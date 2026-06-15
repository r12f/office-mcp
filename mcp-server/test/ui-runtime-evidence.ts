import { execFileSync, spawnSync } from 'node:child_process';
import { createServer } from 'node:http';
import { request as httpsRequest } from 'node:https';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { startDaemon } from '../src/daemon.js';
import type { DaemonConfig } from '../src/config.js';
import type { AddinConnection } from '../src/types.js';
import { EventEmitter } from 'node:events';

type GateStatus = 'passed' | 'failed' | 'skipped';

type EvidenceGate = {
  name: string;
  status: GateStatus;
  started_at: string;
  finished_at: string;
  details: Record<string, unknown>;
};

type EvidenceReport = {
  schema_version: 1;
  generated_at: string;
  kind: 'ui_runtime_evidence';
  ui_url?: string;
  state_url?: string;
  gates: EvidenceGate[];
};

const outputPath = resolve(readOption('--output') ?? '../artifacts/ui-runtime-evidence.json');
const report: EvidenceReport = {
  schema_version: 1,
  generated_at: new Date().toISOString(),
  kind: 'ui_runtime_evidence',
  gates: []
};

const tempDir = mkdtempSync(join(tmpdir(), 'office-mcp-ui-evidence-'));
const runtimePath = join(tempDir, 'ui-runtime.json');
const probeStatePath = join(tempDir, 'ui-state-probe.json');
const previousRuntimePath = process.env.OFFICE_MCP_UI_RUNTIME_PATH;
process.env.OFFICE_MCP_UI_RUNTIME_PATH = runtimePath;

let daemon: Awaited<ReturnType<typeof startDaemon>> | undefined;

try {
  const config = await daemonConfig();
  daemon = await startDaemon(config);
  seedUiState(daemon);
  report.ui_url = `${config.addin.origin}/ui/`;
  report.state_url = `${config.addin.origin}/ui/state`;

  await runRuntimeFileGate();
  await runStateApiGate(config.addin.origin);
  await runEventsStreamGate();
  await runTrayProbeGate();
  await runBrowserSmokeGate();
} finally {
  await daemon?.close().catch(() => undefined);
  if (previousRuntimePath === undefined) delete process.env.OFFICE_MCP_UI_RUNTIME_PATH;
  else process.env.OFFICE_MCP_UI_RUNTIME_PATH = previousRuntimePath;
  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, JSON.stringify(report, null, 2));
  rmSync(tempDir, { recursive: true, force: true });
  console.log(JSON.stringify(report, null, 2));
}

async function runRuntimeFileGate(): Promise<void> {
  await runGate('ui.daemon_runtime_file', async () => {
    const runtime = JSON.parse(readFileSync(runtimePath, 'utf8')) as Record<string, unknown>;
    if (typeof runtime.token !== 'string' || runtime.token.length < 8) throw new Error('Runtime file is missing token.');
    if (typeof runtime.stateUrl !== 'string' || !runtime.stateUrl.endsWith('/ui/state')) throw new Error('Runtime file is missing stateUrl.');
    if (typeof runtime.uiUrl !== 'string' || !runtime.uiUrl.endsWith('/ui/')) throw new Error('Runtime file is missing uiUrl.');
    return { runtime_path: runtimePath, ui_url: runtime.uiUrl, state_url: runtime.stateUrl, pid: runtime.pid };
  });
}

async function runStateApiGate(origin: string): Promise<void> {
  await runGate('ui.state_api_auth_redaction', async () => {
    const runtime = JSON.parse(readFileSync(runtimePath, 'utf8')) as { stateUrl: string; token: string };
    const unauthorized = await httpsText(runtime.stateUrl, { origin, rejectUnauthorized: false });
    if (unauthorized.status !== 401) throw new Error(`Expected 401 without token, got ${unauthorized.status}.`);
    const forbidden = await httpsText(runtime.stateUrl, { origin: 'https://example.invalid', token: runtime.token, rejectUnauthorized: false });
    if (forbidden.status !== 403) throw new Error(`Expected 403 for foreign origin, got ${forbidden.status}.`);
    const authorized = await httpsText(runtime.stateUrl, { origin, token: runtime.token, rejectUnauthorized: false });
    if (authorized.status !== 200) throw new Error(`Expected 200 with token, got ${authorized.status}.`);
    const snapshot = JSON.parse(authorized.body) as Record<string, unknown>;
    writeFileSync(probeStatePath, JSON.stringify(snapshot, null, 2));
    const serialized = JSON.stringify(snapshot);
    if (/secret-value|base64,QUJDREVGRw/.test(serialized)) throw new Error('UI snapshot leaked a seeded secret.');
    if (!/shared_secret=\[redacted\]/.test(serialized)) throw new Error('UI snapshot did not redact shared_secret values.');
    const recent = snapshot.recent_commands as unknown[] | undefined;
    const clients = snapshot.clients as unknown[] | undefined;
    if (!Array.isArray(recent) || recent.length !== 10) throw new Error('Recent command history is not capped at 10.');
    if (!Array.isArray(clients) || clients.length !== 1) throw new Error('Client state was not exposed.');
    return { recent_commands: recent.length, clients: clients.length, bytes: authorized.body.length };
  });
}

async function runEventsStreamGate(): Promise<void> {
  await runGate('ui.events_stream', async () => {
    const runtime = JSON.parse(readFileSync(runtimePath, 'utf8')) as { stateUrl: string; token: string };
    const eventsUrl = runtime.stateUrl.replace(/\/state$/, '/events');
    const body = await readFirstSseEvent(eventsUrl, runtime.token);
    if (!body.includes('event: snapshot')) throw new Error('SSE stream did not emit snapshot event.');
    if (!body.includes('data: ')) throw new Error('SSE stream did not emit data line.');
    return { events_url: eventsUrl, first_event_bytes: body.length };
  });
}

async function runTrayProbeGate(): Promise<void> {
  await runGate('ui.tray_probe', async () => {
    if (process.platform !== 'win32') return { skipped_reason: 'Tray probe is Windows-only.' };
    const script = resolve('..', 'packaging', 'windows', 'office-mcp-tray.ps1');
    const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-InstallRoot', tempDir, '-RuntimePath', runtimePath, '-ProbeStatePath', probeStatePath, '-Probe'], {
      cwd: process.cwd(),
      encoding: 'utf8',
      timeout: 15000,
      env: { ...process.env, OFFICE_MCP_UI_RUNTIME_PATH: runtimePath }
    });
    if (result.status !== 0) throw new Error(`Tray probe failed: ${result.stderr || result.stdout}`);
    const jsonStart = result.stdout.indexOf('{');
    if (jsonStart === -1) throw new Error(`Tray probe did not emit JSON: ${result.stdout}`);
    const evidence = JSON.parse(result.stdout.slice(jsonStart)) as Record<string, unknown>;
    const expected = ['Status: Up', 'Clients: 1', 'Documents: 1', '---', 'Show Office MCP', 'Quit Office MCP'];
    if (JSON.stringify(evidence.menu_items) !== JSON.stringify(expected)) {
      throw new Error(`Tray menu order/counts are wrong: ${JSON.stringify(evidence)}`);
    }
    if (evidence.notify_icon_created !== true || evidence.context_menu_created !== true) throw new Error('Tray probe did not create native objects.');
    if (evidence.state_fetch_ok !== true) throw new Error('Tray probe could not fetch UI state.');
    return evidence;
  });
}

async function runBrowserSmokeGate(): Promise<void> {
  await runGate('ui.browser_smoke', async () => {
    const result = execFileSync(process.execPath, ['./node_modules/tsx/dist/cli.mjs', 'test/ui-browser-smoke.ts'], {
      cwd: process.cwd(),
      encoding: 'utf8',
      maxBuffer: 20 * 1024 * 1024,
      env: process.env
    });
    const jsonStart = result.lastIndexOf('{');
    const parsed = jsonStart === -1 ? {} : JSON.parse(result.slice(jsonStart));
    return { ...parsed, output_bytes: result.length };
  });
}

async function runGate(name: string, run: () => Promise<Record<string, unknown>>): Promise<void> {
  const started = new Date().toISOString();
  try {
    const details = await run();
    report.gates.push({ name, status: 'passed', started_at: started, finished_at: new Date().toISOString(), details });
  } catch (error) {
    report.gates.push({ name, status: 'failed', started_at: started, finished_at: new Date().toISOString(), details: { error: errorMessage(error) } });
  }
}

function seedUiState(running: Awaited<ReturnType<typeof startDaemon>>): void {
  const connection = fakeConnection();
  running.registry.registerRuntime(connection, connection.runtime);
  running.registry.addSession(connection, {
    session_id: '11111111-1111-4111-8111-111111111111',
    instance_id: connection.runtime.instance_id,
    document: {
      title: 'Runtime Evidence.docx',
      filename: 'Runtime Evidence.docx',
      is_dirty: true,
      is_read_only: false,
      protection: { kind: 'irm', rights: null, rights_source: 'unavailable' }
    },
    available_tools: ['word.get_text', 'word.insert_paragraph'],
    is_active: true
  });
  const clientId = running.uiState.registerClient({ client_id: 'client-1', transport: 'http', name: 'copilot-cli/1.0 token=secret-value' });
  for (let index = 0; index < 12; index += 1) {
    const commandId = running.uiState.startCommand({
      client_id: clientId,
      client_name: 'copilot-cli/1.0',
      session_id: '11111111-1111-4111-8111-111111111111',
      host_app: 'word',
      tool: index % 2 === 0 ? 'word.get_text' : 'word.insert_paragraph',
      user_intent: `summarize status token=secret-value base64,QUJDREVGRw== ${index}`,
      timeout_ms: 30000
    });
    running.uiState.finishCommand(commandId, index % 3 === 0 ? {
      ok: false,
      error: { office_mcp_code: 'IRM_DENIED', message: 'shared_secret=secret-value blocked by document policy.', tool: 'word.insert_paragraph', retriable: false, partial_effect: 'none' }
    } : { ok: true, data: { redacted: true } });
  }
}

function fakeConnection(): AddinConnection {
  const socket = new EventEmitter() as AddinConnection['socket'];
  Object.assign(socket, { OPEN: 1, readyState: 1, close() {} });
  return {
    socket,
    runtime: {
      instance_id: '22222222-2222-4222-8222-222222222222',
      host: { app: 'word', version: '16.0', platform: 'pc', build: 'Desktop' },
      add_in: { version: '0.1.5', protocol_version: '1.0' },
      registered_at: new Date().toISOString()
    },
    pending: new Map(),
    queue: Promise.resolve(),
    invokeTool: async () => ({ ok: true, data: {} })
  };
}

async function daemonConfig(): Promise<DaemonConfig> {
  const mcpPort = await freePort();
  const addinPort = await freePort();
  return {
    addin: { host: 'localhost', port: addinPort, origin: `https://localhost:${addinPort}`, pfxPath: '.office-mcp-localhost.pfx', pfxPassphrase: 'office-mcp-localhost', heartbeatIntervalSec: 30, heartbeatTimeoutSec: 10, sessionGraceSec: 60, maxPendingPerSession: 4, sharedSecret: '' },
    mcp: { host: '127.0.0.1', port: mcpPort, apiKey: '' },
    limits: { maxResponseBytes: 1024 * 1024, maxRequestBytes: 16 * 1024 * 1024, maxWsFrameBytes: 16 * 1024 * 1024, defaultToolTimeoutMs: 30000, requestsPerMinute: 120 },
    audit: { enabled: false, path: 'audit.jsonl' },
    logging: { level: 'error', file: '' }
  };
}

function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = createServer();
    server.once('error', reject);
    server.once('listening', () => {
      const address = server.address();
      server.close((error) => {
        if (error) reject(error);
        else if (address && typeof address === 'object') resolve(address.port);
        else reject(new Error('Could not allocate test port.'));
      });
    });
    server.listen(0, '127.0.0.1');
  });
}

function httpsText(url: string, options: { origin?: string; token?: string; rejectUnauthorized?: boolean }): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const headers: Record<string, string> = {};
    if (options.origin) headers.origin = options.origin;
    if (options.token) headers['x-office-mcp-ui-token'] = options.token;
    const req = httpsRequest(url, { headers, rejectUnauthorized: options.rejectUnauthorized }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      res.on('end', () => resolve({ status: res.statusCode ?? 0, body: Buffer.concat(chunks).toString('utf8') }));
    });
    req.on('error', reject);
    req.end();
  });
}

function readFirstSseEvent(url: string, token: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const req = httpsRequest(url, { headers: { 'x-office-mcp-ui-token': token }, rejectUnauthorized: false }, (res) => {
      let body = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => {
        body += chunk;
        if (body.includes('\n\n')) {
          req.destroy();
          resolve(body.slice(0, body.indexOf('\n\n') + 2));
        }
      });
      res.on('end', () => reject(new Error('SSE stream ended before first event.')));
    });
    req.on('error', (error) => {
      if (bodyResolved(error)) return;
      reject(error);
    });
    req.setTimeout(5000, () => {
      req.destroy();
      reject(new Error('Timed out waiting for SSE snapshot.'));
    });
    req.end();
  });
}

function bodyResolved(error: Error): boolean {
  return error.message.includes('socket hang up');
}

function readOption(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
