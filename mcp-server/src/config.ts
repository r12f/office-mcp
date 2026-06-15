import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { isIP } from 'node:net';

export type DaemonConfig = {
  addin: {
    host: string;
    port: number;
    origin: string;
    pfxPath: string;
    pfxPassphrase: string;
    heartbeatIntervalSec: number;
    heartbeatTimeoutSec: number;
    sessionGraceSec: number;
    maxPendingPerSession: number;
    sharedSecret: string;
  };
  mcp: {
    host: string;
    port: number;
    apiKey: string;
  };
  limits: {
    maxResponseBytes: number;
    maxRequestBytes: number;
    maxWsFrameBytes: number;
    defaultToolTimeoutMs: number;
    requestsPerMinute: number;
  };
  audit: {
    enabled: boolean;
    path: string;
  };
  logging: {
    level: LogLevel;
    file: string;
  };
};

export type LogLevel = 'trace' | 'debug' | 'info' | 'warn' | 'error';

type RawTomlValue = string | number | boolean;
type RawToml = Record<string, Record<string, RawTomlValue>>;

export type LoadConfigOptions = {
  configPath?: string;
};

function intEnvAny(names: string[], fallback: number): number {
  const [name, raw] = readEnvAny(names);
  if (!name || raw === undefined) return fallback;
  const value = Number.parseInt(raw, 10);
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`${name} must be a positive integer`);
  }
  return value;
}

export function loadConfig(options: LoadConfigOptions = {}): DaemonConfig {
  const fileConfig = loadConfigFile(options.configPath ?? process.env.OFFICE_MCP_CONFIG_PATH ?? defaultConfigPath());
  const addinChannel = fileConfig.addin_channel ?? {};
  const mcpHttp = fileConfig.mcp_http ?? {};
  const limits = fileConfig.limits ?? {};
  const audit = fileConfig.audit ?? {};
  const logging = fileConfig.logging ?? {};

  const addinHost = stringEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__BIND', 'OFFICE_MCP_ADDIN_HOST'], stringValue(addinChannel.bind, 'localhost'));
  const addinPort = intEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__PORT', 'OFFICE_MCP_ADDIN_PORT'], intValue(addinChannel.port, 8765));
  const pfxPath = stringEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH', 'OFFICE_MCP_ADDIN_PFX_PATH'], stringValue(addinChannel.certificate_path, join(process.cwd(), '.office-mcp-localhost.pfx')));

  return {
    addin: {
      host: addinHost,
      port: addinPort,
      origin: stringEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__ORIGIN', 'OFFICE_MCP_ADDIN_ORIGIN'], `https://${addinHost}:${addinPort}`),
      pfxPath,
      pfxPassphrase: stringEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE', 'OFFICE_MCP_ADDIN_PFX_PASSPHRASE'], stringValue(addinChannel.certificate_passphrase, 'office-mcp-localhost')),
      heartbeatIntervalSec: intEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_INTERVAL_SEC', 'OFFICE_MCP_ADDIN_HEARTBEAT_INTERVAL_SEC'], intValue(addinChannel.heartbeat_interval_sec, 30)),
      heartbeatTimeoutSec: intEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__HEARTBEAT_TIMEOUT_SEC', 'OFFICE_MCP_ADDIN_HEARTBEAT_TIMEOUT_SEC'], intValue(addinChannel.heartbeat_timeout_sec, 10)),
      sessionGraceSec: intEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__SESSION_GRACE_SEC', 'OFFICE_MCP_ADDIN_SESSION_GRACE_SEC'], intValue(addinChannel.session_grace_sec, 60)),
      maxPendingPerSession: intEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__MAX_PENDING_PER_SESSION', 'OFFICE_MCP_ADDIN_MAX_PENDING_PER_SESSION'], intValue(addinChannel.max_pending_per_session, 4)),
      sharedSecret: stringEnvAny(['OFFICE_MCP_ADDIN_CHANNEL__SHARED_SECRET', 'OFFICE_MCP_ADDIN_SHARED_SECRET'], stringValue(addinChannel.shared_secret, ''))
    },
    mcp: {
      host: stringEnvAny(['OFFICE_MCP_MCP_HTTP__BIND', 'OFFICE_MCP_MCP_HOST'], stringValue(mcpHttp.bind, '127.0.0.1')),
      port: intEnvAny(['OFFICE_MCP_MCP_HTTP__PORT', 'OFFICE_MCP_MCP_PORT'], intValue(mcpHttp.port, 8800)),
      apiKey: stringEnvAny(['OFFICE_MCP_MCP_HTTP__API_KEY', 'OFFICE_MCP_MCP_API_KEY'], stringValue(mcpHttp.api_key, ''))
    },
    limits: {
      maxResponseBytes: intEnvAny(['OFFICE_MCP_LIMITS__MAX_RESPONSE_BYTES', 'OFFICE_MCP_MAX_RESPONSE_BYTES'], intValue(limits.max_response_bytes, 1024 * 1024)),
      maxRequestBytes: intEnvAny(['OFFICE_MCP_LIMITS__MAX_REQUEST_BYTES', 'OFFICE_MCP_MAX_REQUEST_BYTES'], intValue(limits.max_request_bytes, 16 * 1024 * 1024)),
      maxWsFrameBytes: intEnvAny(['OFFICE_MCP_LIMITS__MAX_WS_FRAME_BYTES', 'OFFICE_MCP_MAX_WS_FRAME_BYTES'], intValue(limits.max_ws_frame_bytes, 16 * 1024 * 1024)),
      defaultToolTimeoutMs: intEnvAny(['OFFICE_MCP_LIMITS__DEFAULT_TOOL_TIMEOUT_MS', 'OFFICE_MCP_DEFAULT_TOOL_TIMEOUT_MS'], intValue(limits.default_tool_timeout_ms, 30000)),
      requestsPerMinute: intEnvAny(['OFFICE_MCP_LIMITS__REQUESTS_PER_MINUTE', 'OFFICE_MCP_REQUESTS_PER_MINUTE'], intValue(limits.requests_per_minute, 120))
    },
    audit: {
      enabled: boolEnvAny(['OFFICE_MCP_AUDIT__ENABLED', 'OFFICE_MCP_AUDIT_ENABLED'], boolValue(audit.enabled, false)),
      path: stringEnvAny(['OFFICE_MCP_AUDIT__PATH', 'OFFICE_MCP_AUDIT_PATH'], optionalPathValue(audit.path, defaultAuditPath()))
    },
    logging: {
      level: logLevelValue(stringEnvAny(['OFFICE_MCP_LOGGING__LEVEL', 'OFFICE_MCP_LOG_LEVEL'], stringValue(logging.level, 'info'))),
      file: stringEnvAny(['OFFICE_MCP_LOGGING__FILE', 'OFFICE_MCP_LOG_FILE'], optionalPathValue(logging.file, defaultLogPath()))
    }
  };
}

function stringEnvAny(names: string[], fallback: string): string {
  const [, value] = readEnvAny(names);
  return value ?? fallback;
}

function boolEnvAny(names: string[], fallback: boolean): boolean {
  const [name, raw] = readEnvAny(names);
  if (!name || raw === undefined) return fallback;
  if (raw === 'true') return true;
  if (raw === 'false') return false;
  throw new Error(`${name} must be true or false`);
}

function readEnvAny(names: string[]): [string | undefined, string | undefined] {
  for (const name of names) {
    if (process.env[name] !== undefined) return [name, process.env[name]];
  }
  return [undefined, undefined];
}

function stringValue(value: RawTomlValue | undefined, fallback: string): string {
  if (value === undefined) return fallback;
  if (typeof value !== 'string') throw new Error(`Expected string config value, got ${typeof value}.`);
  return value;
}

function intValue(value: RawTomlValue | undefined, fallback: number): number {
  if (value === undefined) return fallback;
  if (typeof value !== 'number' || !Number.isInteger(value) || value <= 0) {
    throw new Error(`Expected positive integer config value, got ${String(value)}.`);
  }
  return value;
}

function boolValue(value: RawTomlValue | undefined, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  if (typeof value !== 'boolean') throw new Error(`Expected boolean config value, got ${typeof value}.`);
  return value;
}

function optionalPathValue(value: RawTomlValue | undefined, fallback: string): string {
  const path = stringValue(value, fallback);
  return path === '' ? fallback : path;
}

function logLevelValue(value: string): LogLevel {
  if (value === 'trace' || value === 'debug' || value === 'info' || value === 'warn' || value === 'error') return value;
  throw new Error(`logging.level must be one of trace, debug, info, warn, error; got ${value}.`);
}

function loadConfigFile(path: string): RawToml {
  if (!existsSync(path)) return {};
  return parseToml(readFileSync(path, 'utf8'));
}

export function parseToml(input: string): RawToml {
  const result: RawToml = {};
  let currentSection = '';
  for (const [index, rawLine] of input.split(/\r?\n/).entries()) {
    const line = stripTomlComment(rawLine).trim();
    if (!line) continue;
    const sectionMatch = /^\[([A-Za-z0-9_]+)\]$/.exec(line);
    if (sectionMatch) {
      currentSection = sectionMatch[1];
      result[currentSection] ??= {};
      continue;
    }
    const keyValueMatch = /^([A-Za-z0-9_]+)\s*=\s*(.+)$/.exec(line);
    if (!keyValueMatch || !currentSection) {
      throw new Error(`Unsupported TOML syntax at line ${index + 1}: ${rawLine}`);
    }
    result[currentSection]![keyValueMatch[1]] = parseTomlValue(keyValueMatch[2].trim(), index + 1);
  }
  return result;
}

function parseTomlValue(raw: string, lineNumber: number): RawTomlValue {
  if (raw.startsWith('"') && raw.endsWith('"')) {
    return JSON.parse(raw) as string;
  }
  if (raw === 'true') return true;
  if (raw === 'false') return false;
  if (/^[0-9]+$/.test(raw)) return Number.parseInt(raw, 10);
  throw new Error(`Unsupported TOML value at line ${lineNumber}: ${raw}`);
}

function stripTomlComment(line: string): string {
  let inString = false;
  let escaped = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (char === '\\' && inString) {
      escaped = true;
      continue;
    }
    if (char === '"') inString = !inString;
    if (char === '#' && !inString) return line.slice(0, i);
  }
  return line;
}

function defaultConfigPath(): string {
  if (process.platform === 'win32') {
    return join(process.env.APPDATA ?? join(process.env.USERPROFILE ?? 'C:\\Users\\Default', 'AppData', 'Roaming'), 'office-mcp', 'config.toml');
  }
  if (process.platform === 'darwin') {
    return join(process.env.HOME ?? '.', 'Library', 'Application Support', 'office-mcp', 'config.toml');
  }
  return join(process.env.XDG_CONFIG_HOME ?? join(process.env.HOME ?? '.', '.config'), 'office-mcp', 'config.toml');
}

function defaultAuditPath(): string {
  if (process.platform === 'win32') {
    return join(process.env.LOCALAPPDATA ?? join(process.env.USERPROFILE ?? 'C:\\Users\\Default', 'AppData', 'Local'), 'office-mcp', 'audit.jsonl');
  }
  if (process.platform === 'darwin') {
    return join(process.env.HOME ?? '.', 'Library', 'Logs', 'office-mcp', 'audit.jsonl');
  }
  return join(process.env.XDG_STATE_HOME ?? join(process.env.HOME ?? '.', '.local', 'state'), 'office-mcp', 'audit.jsonl');
}

function defaultLogPath(): string {
  if (process.platform === 'win32') {
    return join(process.env.LOCALAPPDATA ?? join(process.env.USERPROFILE ?? 'C:\\Users\\Default', 'AppData', 'Local'), 'office-mcp', 'office-mcp.log');
  }
  if (process.platform === 'darwin') {
    return join(process.env.HOME ?? '.', 'Library', 'Logs', 'office-mcp', 'office-mcp.log');
  }
  return join(process.env.XDG_STATE_HOME ?? join(process.env.HOME ?? '.', '.local', 'state'), 'office-mcp', 'office-mcp.log');
}

export function assertHttpsConfig(config: DaemonConfig): void {
  assertBoundaryAuthConfig(config);
  if (!existsSync(config.addin.pfxPath)) {
    throw new Error(
      `Missing HTTPS certificate PFX at ${config.addin.pfxPath}. Run scripts/export-localhost-dev-cert.ps1 from mcp-server/ or set OFFICE_MCP_ADDIN_PFX_PATH.`
    );
  }
}

export function assertBoundaryAuthConfig(config: DaemonConfig): void {
  if (!isLoopbackHost(config.addin.host) && !config.addin.sharedSecret) {
    throw new Error('Refusing to bind add-in WSS to a non-loopback address without OFFICE_MCP_ADDIN_SHARED_SECRET.');
  }
  if (!isLoopbackHost(config.mcp.host) && !config.mcp.apiKey) {
    throw new Error('Refusing to bind MCP HTTP to a non-loopback address without OFFICE_MCP_MCP_API_KEY.');
  }
}

function isLoopbackHost(host: string): boolean {
  const normalized = host.toLowerCase();
  if (normalized === 'localhost') return true;
  if (normalized === '::1' || normalized === '[::1]') return true;
  if (isIP(normalized) === 4) return normalized.startsWith('127.');
  return false;
}

