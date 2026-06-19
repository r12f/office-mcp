import { execFileSync, spawn, spawnSync, type ChildProcess, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { request as httpsRequest } from 'node:https';
import { copyFileSync, existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { setTimeout as delay } from 'node:timers/promises';

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

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const tsxCli = resolve(evidenceRoot, 'node_modules/tsx/dist/cli.mjs');
const cargoCommand = process.env.CARGO || resolve(process.env.USERPROFILE || '', '.cargo/bin/cargo.exe');
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/ui-runtime-evidence.json'));
const report: EvidenceReport = {
  schema_version: 1,
  generated_at: new Date().toISOString(),
  kind: 'ui_runtime_evidence',
  gates: []
};

const tempDir = mkdtempSync(join(tmpdir(), 'office-mcp-ui-evidence-'));
const cargoTargetDir = join(tempDir, 'target');
const runtimePath = join(tempDir, 'ui-runtime.json');
const probeStatePath = join(tempDir, 'ui-state-probe.json');
const productionRuntimePath = join(tempDir, 'production-ui-runtime.json');
const productionLogPath = join(tempDir, 'production-office-mcp.log');
const productionConfigPath = join(tempDir, 'production-config.toml');
const previousRuntimePath = process.env.OFFICE_MCP_UI_RUNTIME_PATH;
process.env.OFFICE_MCP_UI_RUNTIME_PATH = runtimePath;

let daemon: ChildProcess | undefined;

try {
  daemon = startRustUiFixture(runtimePath);
  const runtime = await waitForRuntimeFile(runtimePath);
  report.ui_url = runtime.uiUrl;
  report.state_url = runtime.stateUrl;

  await runRuntimeFileGate();
  await runStateApiGate(runtime.origin);
  await runEventsStreamGate();
  await runTrayProbeGate();
  await runProductionDaemonTrayGate();
  await runBrowserSmokeGate();
} finally {
  await stopProcess(daemon);
  if (previousRuntimePath === undefined) delete process.env.OFFICE_MCP_UI_RUNTIME_PATH;
  else process.env.OFFICE_MCP_UI_RUNTIME_PATH = previousRuntimePath;
  mkdirSync(dirname(outputPath), { recursive: true });
  writeFileSync(outputPath, JSON.stringify(report, null, 2));
  rmSync(tempDir, { recursive: true, force: true });
  console.log(JSON.stringify(report, null, 2));
}

function startRustUiFixture(runtimePath: string): ChildProcess {
  execFileSync(cargoCommand, ['build', '-q', '-p', 'office-mcp-daemon'], {
    cwd: repoRoot,
    env: { ...process.env, CARGO_TARGET_DIR: cargoTargetDir },
    stdio: 'inherit'
  });
  const child = spawn(daemonExecutablePath(), ['evidence', 'ui-fixture'], {
    cwd: repoRoot,
    env: { ...process.env, OFFICE_MCP_UI_RUNTIME_PATH: runtimePath },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', (chunk) => process.stderr.write(chunk));
  child.stdout.setEncoding('utf8');
  child.stdout.on('data', (chunk) => process.stderr.write(chunk));
  child.on('error', (error) => process.stderr.write(`Failed to start Rust UI fixture: ${error.message}\n`));
  return child;
}

async function waitForRuntimeFile(path: string): Promise<{ origin: string; stateUrl: string; uiUrl: string }> {
  for (let index = 0; index < 600; index += 1) {
    if (existsSync(path)) {
      const runtime = JSON.parse(readFileSync(path, 'utf8')) as { origin?: string; stateUrl?: string; uiUrl?: string };
      if (runtime.origin && runtime.stateUrl && runtime.uiUrl) return { origin: runtime.origin, stateUrl: runtime.stateUrl, uiUrl: runtime.uiUrl };
    }
    if (daemon?.exitCode !== null) throw new Error(`Rust UI fixture exited before writing runtime file with code ${daemon?.exitCode}.`);
    await delay(100);
  }
  throw new Error('Timed out waiting for Rust UI fixture runtime file.');
}

async function runRuntimeFileGate(): Promise<void> {
  await runGate('ui.daemon_runtime_file', async () => {
    const runtime = JSON.parse(readFileSync(runtimePath, 'utf8')) as Record<string, unknown>;
    if (typeof runtime.stateUrl !== 'string' || !runtime.stateUrl.endsWith('/ui/state')) throw new Error('Runtime file is missing stateUrl.');
    if (typeof runtime.uiUrl !== 'string' || !runtime.uiUrl.endsWith('/ui/')) throw new Error('Runtime file is missing uiUrl.');
    return { runtime_path: runtimePath, ui_url: runtime.uiUrl, state_url: runtime.stateUrl, pid: runtime.pid };
  });
}

async function runStateApiGate(origin: string): Promise<void> {
  await runGate('ui.state_api_origin_redaction', async () => {
    const runtime = JSON.parse(readFileSync(runtimePath, 'utf8')) as { stateUrl: string };
    const forbidden = await httpsText(runtime.stateUrl, { origin: 'https://example.invalid', rejectUnauthorized: false });
    if (forbidden.status !== 403) throw new Error('Expected 403 for foreign origin, got ' + forbidden.status + '.');
    const authorized = await httpsText(runtime.stateUrl, { origin, rejectUnauthorized: false });
    if (authorized.status !== 200) throw new Error('Expected 200 from trusted origin, got ' + authorized.status + '.');
    const snapshot = JSON.parse(authorized.body) as Record<string, unknown>;
    writeFileSync(probeStatePath, JSON.stringify(snapshot, null, 2));
    const serialized = JSON.stringify(snapshot);
    if (/secret-value|base64,QUJDREVGRw/.test(serialized)) throw new Error('UI snapshot leaked a seeded secret.');
    if (!/certificate_passphrase=\[redacted\]/.test(serialized)) throw new Error('UI snapshot did not redact certificate passphrase values.');
    const recent = snapshot.recent_commands as unknown[] | undefined;
    const clients = snapshot.clients as unknown[] | undefined;
    if (!Array.isArray(recent) || recent.length !== 10) throw new Error('Recent command history is not capped at 10.');
    if (!Array.isArray(clients) || clients.length !== 1) throw new Error('Client state was not exposed.');
    return { recent_commands: recent.length, clients: clients.length, bytes: authorized.body.length };
  });
}

async function runEventsStreamGate(): Promise<void> {
  await runGate('ui.events_stream', async () => {
    const runtime = JSON.parse(readFileSync(runtimePath, 'utf8')) as { stateUrl: string };
    const eventsUrl = runtime.stateUrl.replace(/\/state$/, '/events');
    const body = await readFirstSseEvent(eventsUrl);
    if (!body.includes('event: snapshot')) throw new Error('SSE stream did not emit snapshot event.');
    if (!body.includes('data: ')) throw new Error('SSE stream did not emit data line.');
    return { events_url: eventsUrl, first_event_bytes: body.length };
  });
}

async function runTrayProbeGate(): Promise<void> {
  await runGate('ui.tray_probe', async () => {
    if (process.platform !== 'win32') return { skipped_reason: 'Tray probe is Windows-only.' };
    const installRoot = prepareTrayProbeInstallRoot();
    const script = resolve(repoRoot, 'packaging', 'windows', 'office-mcp-tray.ps1');
    const result = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, '-InstallRoot', installRoot, '-RuntimePath', runtimePath, '-ProbeStatePath', probeStatePath, '-Probe'], {
      cwd: repoRoot,
      encoding: 'utf8',
      timeout: 15000,
      env: { ...process.env, OFFICE_MCP_UI_RUNTIME_PATH: runtimePath }
    });
    if (result.status !== 0) throw new Error(`Tray probe failed: ${result.stderr || result.stdout}`);
    const jsonStart = result.stdout.indexOf('{');
    if (jsonStart === -1) throw new Error(`Tray probe did not emit JSON: ${result.stdout}`);
    const evidence = JSON.parse(result.stdout.slice(jsonStart)) as Record<string, unknown>;
    const snapshot = evidence.snapshot as Record<string, unknown> | undefined;
    const expected = ['Status: Degraded', 'Clients: 1', 'Documents: 1', '---', 'Show Office MCP Control', 'Quit Office MCP Control'];
    if (JSON.stringify(snapshot?.menu_items) !== JSON.stringify(expected)) {
      throw new Error(`Tray menu order/counts are wrong: ${JSON.stringify(evidence)}`);
    }
    if (evidence.native_host !== true) throw new Error('Tray probe did not use the native tray host.');
    if (evidence.state_fetch_ok !== true) throw new Error('Tray probe could not fetch UI state.');
    if (snapshot?.platform !== 'windows-notification-area') throw new Error(`Tray probe reported wrong platform: ${snapshot?.platform}`);
    assertStructuredTraySnapshot(snapshot, 'Tray probe');
    if ((snapshot?.quit_confirmation as Record<string, unknown> | undefined)?.secondary_action !== 'Keep Running') throw new Error('Tray probe did not include quit confirmation details.');
    return evidence;
  });
}

async function runProductionDaemonTrayGate(): Promise<void> {
  await runGate('ui.production_daemon_tray', async () => {
    if (process.platform !== 'win32') return { skipped_reason: 'Native visible tray evidence is Windows-only.' };
    const ports = await reserveProductionPorts();
    writeProductionConfig(ports);
    rmSync(productionRuntimePath, { force: true });
    rmSync(productionLogPath, { force: true });
    const daemonRun = spawn(daemonExecutablePath(), ['daemon', 'run'], {
      cwd: repoRoot,
      env: {
        ...process.env,
        OFFICE_MCP_CONFIG_PATH: productionConfigPath,
        OFFICE_MCP_UI_RUNTIME_PATH: productionRuntimePath,
        OFFICE_MCP_LOGGING__FILE: productionLogPath,
        OFFICE_MCP_LOGGING__LEVEL: 'debug'
      },
      stdio: ['ignore', 'pipe', 'pipe']
    });
    daemonRun.stderr.setEncoding('utf8');
    daemonRun.stderr.on('data', (chunk) => process.stderr.write(chunk));
    daemonRun.stdout.setEncoding('utf8');
    daemonRun.stdout.on('data', (chunk) => process.stderr.write(chunk));
    try {
      const runtime = await waitForRuntimeFileAt(productionRuntimePath, daemonRun);
      const state = await httpsText(runtime.stateUrl, { rejectUnauthorized: false });
      if (state.status !== 200) throw new Error(`Production /ui/state returned ${state.status}.`);
      const snapshot = JSON.parse(state.body) as { daemon?: { status?: string } };
      if (snapshot.daemon?.status !== 'up') throw new Error(`Production daemon status is ${snapshot.daemon?.status}.`);
      const probe = spawnSync(daemonExecutablePath(), ['tray', '--probe', '--runtime-path', productionRuntimePath], {
        cwd: repoRoot,
        encoding: 'utf8',
        timeout: 15000,
        env: { ...process.env, OFFICE_MCP_UI_RUNTIME_PATH: productionRuntimePath }
      });
      if (probe.status !== 0) throw new Error(`Production tray probe failed: ${probe.stderr || probe.stdout}`);
      const probeJson = parseJsonFromStdout(probe.stdout, 'Production tray probe');
      const probeSnapshot = probeJson.snapshot as Record<string, unknown> | undefined;
      if (probeJson.state_fetch_ok !== true) throw new Error(`Production tray probe could not fetch UI state: ${probe.stdout}`);
      if (probeSnapshot?.platform !== 'windows-notification-area') throw new Error(`Wrong tray platform: ${probeSnapshot?.platform}`);
      if (!Array.isArray(probeSnapshot?.menu_items) || probeSnapshot.menu_items[0] !== 'Status: Up') throw new Error(`Tray menu did not report Up: ${probe.stdout}`);
      assertStructuredTraySnapshot(probeSnapshot, 'Production tray probe');
      const logText = await waitForLogLine(productionLogPath, 'created native tray icon', daemonRun);
      if (!logText.includes('windows-notification-area')) throw new Error('Native tray creation log did not include Windows platform evidence.');
      return {
        runtime_path: productionRuntimePath,
        state_url: runtime.stateUrl,
        ui_url: runtime.uiUrl,
        log_path: productionLogPath,
        tray_probe: probeJson,
        tray_log_contains_native_icon: true
      };
    } finally {
      await stopProcess(daemonRun);
    }
  });
}


function assertStructuredTraySnapshot(snapshot: Record<string, unknown> | undefined, label: string): void {
  if (!snapshot) throw new Error(`${label} did not include a tray snapshot.`);
  if (typeof snapshot.tooltip !== 'string' || !/^Office MCP Control - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(snapshot.tooltip)) {
    throw new Error(`${label} did not include a product tooltip: ${JSON.stringify(snapshot)}`);
  }
  const menu = Array.isArray(snapshot.menu) ? snapshot.menu : [];
  const expected = [
    { kind: 'read_only', enabled: false, label: /^Status: (Up|Degraded|Down)$/ },
    { kind: 'read_only', enabled: false, label: /^Clients: \d+$/ },
    { kind: 'read_only', enabled: false, label: /^Documents: \d+$/ },
    { kind: 'separator', enabled: false, label: /^---$/ },
    { kind: 'action', enabled: true, label: /^Show Office MCP Control$/, action: 'show_ui' },
    { kind: 'action', enabled: true, label: /^Quit Office MCP Control$/, action: 'quit' }
  ];
  if (menu.length !== expected.length) throw new Error(`${label} structured menu has ${menu.length} items: ${JSON.stringify(snapshot)}`);
  expected.forEach((rule, index) => {
    const item = menu[index];
    if (!isRecord(item)) throw new Error(`${label} structured menu item ${index} is malformed: ${JSON.stringify(item)}`);
    if (item.kind !== rule.kind || item.enabled !== rule.enabled || typeof item.label !== 'string' || !rule.label.test(item.label)) {
      throw new Error(`${label} structured menu item ${index} is wrong: ${JSON.stringify(item)}`);
    }
    if ('action' in rule && item.action !== rule.action) throw new Error(`${label} structured menu action ${index} is wrong: ${JSON.stringify(item)}`);
  });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
function prepareTrayProbeInstallRoot(): string {
  const daemonExe = daemonExecutablePath();
  if (!existsSync(daemonExe)) {
    execFileSync(cargoCommand, ['build', '-p', 'office-mcp-daemon'], {
      cwd: repoRoot,
      stdio: 'inherit',
      env: { ...process.env, CARGO_TARGET_DIR: cargoTargetDir }
    });
  }
  if (!existsSync(daemonExe)) throw new Error(`Cannot find built daemon executable: ${daemonExe}`);
  copyFileSync(daemonExe, join(tempDir, 'office-mcp-daemon.exe'));
  return tempDir;
}

async function reserveProductionPorts(): Promise<{ addin: number; mcp: number }> {
  const mcp = await reservePort();
  const addin = await reservePort();
  if (mcp === addin) return reserveProductionPorts();
  return { addin, mcp };
}

async function reservePort(): Promise<number> {
  const net = await import('node:net');
  return await new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      if (!address || typeof address === 'string') {
        server.close(() => reject(new Error('Failed to reserve TCP port.')));
        return;
      }
      const port = address.port;
      server.close(() => resolve(port));
    });
  });
}

function writeProductionConfig(ports: { addin: number; mcp: number }): void {
  writeFileSync(productionConfigPath, [
    '[addin_channel]',
    'bind = "localhost"',
    `port = ${ports.addin}`,
    'certificate_path = ".office-mcp-localhost.pfx"',
    'certificate_passphrase = "office-mcp-localhost"',
    '',
    '[mcp_http]',
    'bind = "127.0.0.1"',
    `port = ${ports.mcp}`,
    '',
    '[logging]',
    'level = "debug"',
    `file = ${JSON.stringify(productionLogPath)}`,
    ''
  ].join('\n'));
}

async function waitForRuntimeFileAt(path: string, child: ChildProcess): Promise<{ origin: string; stateUrl: string; uiUrl: string }> {
  for (let index = 0; index < 600; index += 1) {
    if (existsSync(path)) {
      const runtime = JSON.parse(readFileSync(path, 'utf8')) as { origin?: string; stateUrl?: string; uiUrl?: string };
      if (runtime.origin && runtime.stateUrl && runtime.uiUrl) return { origin: runtime.origin, stateUrl: runtime.stateUrl, uiUrl: runtime.uiUrl };
    }
    if (child.exitCode !== null) throw new Error(`Production daemon exited before writing runtime file with code ${child.exitCode}.`);
    await delay(100);
  }
  throw new Error('Timed out waiting for production daemon runtime file.');
}

async function waitForLogLine(path: string, needle: string, child: ChildProcess): Promise<string> {
  for (let index = 0; index < 100; index += 1) {
    if (existsSync(path)) {
      const text = readFileSync(path, 'utf8');
      if (text.includes(needle)) return text;
    }
    if (child.exitCode !== null) throw new Error(`Production daemon exited before logging ${needle}.`);
    await delay(100);
  }
  throw new Error(`Timed out waiting for log line: ${needle}`);
}

function parseJsonFromStdout(stdout: string, label: string): Record<string, unknown> {
  const jsonStart = stdout.indexOf('{');
  if (jsonStart === -1) throw new Error(`${label} did not emit JSON: ${stdout}`);
  return JSON.parse(stdout.slice(jsonStart)) as Record<string, unknown>;
}

function daemonExecutablePath(): string {
  return join(cargoTargetDir, 'debug', process.platform === 'win32' ? 'office-mcp-daemon.exe' : 'office-mcp-daemon');
}

async function runBrowserSmokeGate(): Promise<void> {
  await runGate('ui.browser_smoke', async () => {
    const result = execFileSync(process.execPath, [tsxCli, resolve(evidenceRoot, 'ui-browser-smoke.ts')], {
      cwd: evidenceRoot,
      encoding: 'utf8',
      maxBuffer: 20 * 1024 * 1024,
      env: { ...process.env, OFFICE_MCP_DAEMON_EXE: daemonExecutablePath() }
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

function httpsText(url: string, options: { origin?: string; rejectUnauthorized?: boolean }): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const headers: Record<string, string> = {};
    if (options.origin) headers.origin = options.origin;
    const req = httpsRequest(url, { headers, rejectUnauthorized: options.rejectUnauthorized }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      res.on('end', () => resolve({ status: res.statusCode ?? 0, body: Buffer.concat(chunks).toString('utf8') }));
    });
    req.on('error', reject);
    req.end();
  });
}

function readFirstSseEvent(url: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const req = httpsRequest(url, { rejectUnauthorized: false }, (res) => {
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

async function stopProcess(process: ChildProcess | undefined): Promise<void> {
  if (!process || process.exitCode !== null || process.killed) return;
  process.kill();
  await Promise.race([
    new Promise<void>((resolve) => process.once('exit', () => resolve())),
    delay(3000).then(() => {
      if (process.exitCode === null && !process.killed) process.kill('SIGKILL');
    })
  ]);
}

function readOption(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
