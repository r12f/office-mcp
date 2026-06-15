import { randomUUID } from 'node:crypto';
import { EventEmitter } from 'node:events';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import type { AddinToolResult, OfficeMcpCode, RuntimeInfo, SessionDescriptor, ToolFailure } from './types.js';

export type UiHealth = 'up' | 'degraded' | 'down';
export type UiCommandStatus = 'running' | 'success' | 'failure' | 'cancelled' | 'timeout';

export type UiClientRecord = {
  client_id: string;
  transport: 'http' | 'stdio-bridge';
  name: string | null;
  connected_at: string;
  last_activity_at: string;
  in_flight_request_count: number;
};

export type UiCommandRecord = {
  command_id: string;
  mcp_request_id?: string;
  client_id?: string;
  client_name?: string;
  session_id?: string;
  host_app?: string;
  tool: string;
  user_intent?: string;
  status: UiCommandStatus;
  started_at: string;
  deadline_at?: string;
  timeout_ms?: number;
  completed_at?: string;
  elapsed_ms?: number;
  error: UiCommandError | null;
};

export type UiCommandError = {
  office_mcp_code: OfficeMcpCode | 'THROWN';
  message: string;
  tool?: string;
  retriable: boolean;
  partial_effect?: 'none' | 'possible' | 'unknown';
};

export type UiSnapshot = {
  daemon: {
    status: UiHealth;
    version: string;
    uptime_ms: number;
    mcp_endpoint: string;
    addin_endpoint: string;
    config_path: string | null;
    log_path: string | null;
    last_error: string | null;
  };
  clients: UiClientRecord[];
  documents: Record<string, SessionDescriptor[]>;
  current_tasks: UiCommandRecord[];
  recent_commands: UiCommandRecord[];
  document_command_history: Record<string, UiCommandRecord[]>;
};

export type UiStateOptions = {
  version: string;
  mcpEndpoint: string;
  addinEndpoint: string;
  configPath?: string | null;
  logPath?: string | null;
  sessions: () => SessionDescriptor[];
  now?: () => Date;
};

export type UiRuntimeInfo = {
  origin: string;
  stateUrl: string;
  uiUrl: string;
  token: string;
  pid: number;
  createdAt: string;
};

export class UiStateStore {
  private readonly events = new EventEmitter();
  private readonly clients = new Map<string, UiClientRecord>();
  private readonly currentTasks = new Map<string, UiCommandRecord>();
  private readonly recentCommands: UiCommandRecord[] = [];
  private readonly commandsBySession = new Map<string, UiCommandRecord[]>();
  private readonly startedAt: number;
  private health: UiHealth = 'up';
  private lastError: string | null = null;

  constructor(private readonly options: UiStateOptions) {
    this.startedAt = this.options.now?.().getTime() ?? Date.now();
  }

  setHealth(status: UiHealth, lastError: string | null = null): void {
    this.health = status;
    this.lastError = redactText(lastError);
    this.emitSnapshot();
  }

  registerClient(input: { client_id?: string; transport: UiClientRecord['transport']; name?: string | null }): string {
    const now = this.timestamp();
    const clientId = input.client_id ?? randomUUID();
    this.clients.set(clientId, {
      client_id: clientId,
      transport: input.transport,
      name: redactText(input.name),
      connected_at: now,
      last_activity_at: now,
      in_flight_request_count: 0
    });
    this.emitSnapshot();
    return clientId;
  }

  unregisterClient(clientId: string): void {
    this.clients.delete(clientId);
    this.emitSnapshot();
  }

  touchClient(clientId: string): void {
    const client = this.clients.get(clientId);
    if (!client) return;
    client.last_activity_at = this.timestamp();
    this.emitSnapshot();
  }

  startCommand(input: {
    command_id?: string;
    mcp_request_id?: string;
    client_id?: string;
    client_name?: string;
    session_id?: string;
    host_app?: string;
    tool: string;
    user_intent?: string;
    timeout_ms?: number;
  }): string {
    const commandId = input.command_id ?? randomUUID();
    const startedAt = this.timestamp();
    const command: UiCommandRecord = {
      command_id: commandId,
      mcp_request_id: redactText(input.mcp_request_id) ?? undefined,
      client_id: redactText(input.client_id) ?? undefined,
      client_name: redactText(input.client_name) ?? undefined,
      session_id: redactText(input.session_id) ?? undefined,
      host_app: redactText(input.host_app) ?? undefined,
      tool: input.tool,
      user_intent: redactText(input.user_intent) ?? undefined,
      status: 'running',
      started_at: startedAt,
      timeout_ms: input.timeout_ms,
      deadline_at: input.timeout_ms ? new Date(Date.parse(startedAt) + input.timeout_ms).toISOString() : undefined,
      error: null
    };
    this.currentTasks.set(commandId, command);
    if (input.client_id) this.incrementClient(input.client_id, 1);
    this.emitSnapshot();
    return commandId;
  }

  finishCommand(commandId: string, result: AddinToolResult | ToolFailure | Error): void {
    const command = this.currentTasks.get(commandId);
    if (!command) return;
    this.currentTasks.delete(commandId);
    if (command.client_id) this.incrementClient(command.client_id, -1);

    const completedAt = this.timestamp();
    const elapsedMs = Math.max(0, Date.parse(completedAt) - Date.parse(command.started_at));
    const finished: UiCommandRecord = {
      ...command,
      completed_at: completedAt,
      elapsed_ms: elapsedMs,
      ...resultStatus(result)
    };
    this.recentCommands.unshift(finished);
    this.recentCommands.splice(10);
    if (finished.session_id) {
      const sessionCommands = this.commandsBySession.get(finished.session_id) ?? [];
      sessionCommands.unshift(finished);
      sessionCommands.splice(10);
      this.commandsBySession.set(finished.session_id, sessionCommands);
    }
    this.emitSnapshot();
  }

  subscribe(listener: (snapshot: UiSnapshot) => void): () => void {
    this.events.on('snapshot', listener);
    return () => this.events.off('snapshot', listener);
  }

  notifyChanged(): void {
    this.emitSnapshot();
  }

  snapshot(): UiSnapshot {
    const sessions = this.options.sessions().filter((session) => session.status !== 'stale' || session.session_id);
    return {
      daemon: {
        status: this.health,
        version: this.options.version,
        uptime_ms: Math.max(0, (this.options.now?.().getTime() ?? Date.now()) - this.startedAt),
        mcp_endpoint: this.options.mcpEndpoint,
        addin_endpoint: this.options.addinEndpoint,
        config_path: this.options.configPath ?? null,
        log_path: this.options.logPath ?? null,
        last_error: this.lastError
      },
      clients: [...this.clients.values()].map((client) => ({ ...client })),
      documents: groupSessionsByApp(sessions),
      current_tasks: [...this.currentTasks.values()].map(cloneCommand),
      recent_commands: this.recentCommands.map(cloneCommand),
      document_command_history: Object.fromEntries([...this.commandsBySession.entries()].map(([sessionId, commands]) => [sessionId, commands.map(cloneCommand)]))
    };
  }

  private incrementClient(clientId: string, delta: number): void {
    const client = this.clients.get(clientId);
    if (!client) return;
    client.in_flight_request_count = Math.max(0, client.in_flight_request_count + delta);
    client.last_activity_at = this.timestamp();
  }

  private emitSnapshot(): void {
    this.events.emit('snapshot', this.snapshot());
  }

  private timestamp(): string {
    return (this.options.now?.() ?? new Date()).toISOString();
  }
}

export function createUiToken(): string {
  return randomUUID();
}

export function defaultUiRuntimePath(): string {
  if (process.platform === 'win32') {
    return join(process.env.LOCALAPPDATA ?? join(process.env.USERPROFILE ?? 'C:\\Users\\Default', 'AppData', 'Local'), 'office-mcp', 'ui-runtime.json');
  }
  if (process.platform === 'darwin') {
    return join(process.env.HOME ?? '.', 'Library', 'Application Support', 'office-mcp', 'ui-runtime.json');
  }
  return join(process.env.XDG_RUNTIME_DIR ?? join(process.env.HOME ?? '.', '.local', 'state', 'office-mcp'), 'ui-runtime.json');
}

export function writeUiRuntimeFile(path: string, info: UiRuntimeInfo): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, JSON.stringify(info, null, 2), { encoding: 'utf8', mode: 0o600 });
}

export function removeUiRuntimeFile(path: string): void {
  rmSync(path, { force: true });
}

export function authorizeUiRequest(headers: Record<string, string | string[] | undefined>, token: string): boolean {
  const headerToken = firstHeader(headers['x-office-mcp-ui-token']);
  if (headerToken === token) return true;
  const auth = firstHeader(headers.authorization);
  return auth === `Bearer ${token}`;
}

function resultStatus(result: AddinToolResult | ToolFailure | Error): Pick<UiCommandRecord, 'status' | 'error'> {
  if (result instanceof Error) {
    return { status: 'failure', error: { office_mcp_code: 'THROWN', message: redactText(result.message) ?? 'Command failed.', retriable: false } };
  }
  if ('office_mcp_code' in result) return failureStatus(result);
  if (result.ok) return { status: 'success', error: null };
  return failureStatus(result.error);
}

function failureStatus(failure: ToolFailure): Pick<UiCommandRecord, 'status' | 'error'> {
  return {
    status: failure.office_mcp_code === 'TIMEOUT' ? 'timeout' : failure.office_mcp_code === 'CANCELLED' ? 'cancelled' : 'failure',
    error: {
      office_mcp_code: failure.office_mcp_code,
      message: redactText(failure.message) ?? 'Command failed.',
      tool: failure.tool,
      retriable: failure.retriable,
      partial_effect: failure.partial_effect
    }
  };
}

function groupSessionsByApp(sessions: SessionDescriptor[]): Record<string, SessionDescriptor[]> {
  const grouped: Record<string, SessionDescriptor[]> = { word: [], excel: [], powerpoint: [], outlook: [], other: [] };
  for (const session of sessions) {
    const key = ['word', 'excel', 'powerpoint', 'outlook'].includes(session.app) ? session.app : 'other';
    grouped[key].push(session);
  }
  return grouped;
}

function cloneCommand(command: UiCommandRecord): UiCommandRecord {
  return { ...command, error: command.error ? { ...command.error } : null };
}

function firstHeader(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

function redactText(value: string | null | undefined): string | null {
  if (value === undefined || value === null) return null;
  return value
    .replace(/Bearer\s+[A-Za-z0-9._~+\/-]+=*/gi, 'Bearer [redacted]')
    .replace(/(api[_-]?key|shared[_-]?secret|password|passphrase|token)=([^\s&]+)/gi, '$1=[redacted]')
    .replace(/base64,[A-Za-z0-9+/=]+/g, 'base64,[redacted]')
    .slice(0, 500);
}
