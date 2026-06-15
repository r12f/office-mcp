import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { EventEmitter } from 'node:events';
import { mkdtempSync, rmSync } from 'node:fs';
import { createServer, request as httpRequest } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import WebSocket from 'ws';
import { startDaemon } from '../src/daemon.js';
import type { DaemonConfig } from '../src/config.js';
import type { AddinConnection } from '../src/types.js';

const chromePath = process.env.OFFICE_MCP_CHROME_PATH || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';

type CdpResult<T> = { id: number; result?: T; error?: { message: string } };

async function main(): Promise<void> {
  const watchdog = setTimeout(() => {
    console.error('UI smoke timed out.');
    process.exit(1);
  }, 90000);
  const config = await daemonConfig();
  const daemon = await startDaemon(config);
  let chrome: ChildProcessWithoutNullStreams | undefined;
  let cdp: CdpClient | undefined;
  const userDataDir = mkdtempSync(join(tmpdir(), 'office-mcp-ui-chrome-'));
  try {
    console.error('seed state');
    seedUiState(daemon);
    console.error('launch chrome');
    chrome = await launchChrome(userDataDir, 9339);
    console.error('open page');
    const page = await openPage(9339, `${config.addin.origin}/ui/`);
    console.error('connect cdp');
    cdp = await CdpClient.connect(page.webSocketDebuggerUrl);
    await cdp.send('Security.setIgnoreCertificateErrors', { ignore: true });
    await cdp.send('Page.enable');
    await cdp.send('Runtime.enable');
    console.error('desktop assertions');
    await setViewport(cdp, 1366, 768, false);
    await cdp.send('Page.navigate', { url: `${config.addin.origin}/ui/` });
    await waitFor(cdp, 'document.querySelectorAll("table").length >= 2');
    await assertEval(cdp, 'document.querySelector("#clients table") !== null', 'client table renders');
    await assertEval(cdp, 'document.querySelector("#documents .row.word") !== null', 'word document row renders');
    await assertEval(cdp, 'getComputedStyle(document.querySelector(".layout")).gridTemplateColumns.split(" ").length >= 3', 'desktop layout uses three columns');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#documents .row.word").click()' });
    await waitFor(cdp, 'document.querySelector("#inspector").textContent.includes("Latest 10 Commands")');
    await assertEval(cdp, 'location.search.includes("selected=document")', 'selected row is reflected in URL');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#resultFilter").value = "failure"; document.querySelector("#resultFilter").dispatchEvent(new Event("change"))' });
    await waitFor(cdp, 'location.search.includes("result=failure")');
    await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Tab', windowsVirtualKeyCode: 9, nativeVirtualKeyCode: 9 });
    await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Tab', windowsVirtualKeyCode: 9, nativeVirtualKeyCode: 9 });
    await assertEval(cdp, 'document.activeElement !== document.body', 'keyboard tab reaches a focusable control');
    console.error('narrow assertions');
    await setViewport(cdp, 320, 720, true);
    await assertEval(cdp, 'document.documentElement.scrollWidth <= 320', '320px viewport has no horizontal overflow');
    await assertEval(cdp, 'getComputedStyle(document.querySelector(".layout")).gridTemplateColumns.split(" ").length === 1', 'narrow layout stacks columns');
    console.error('theme assertions');
    await cdp.send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: 'dark' }] });
    await assertEval(cdp, 'getComputedStyle(document.documentElement).colorScheme.includes("dark")', 'dark mode color scheme applies');
    await cdp.send('Emulation.setEmulatedMedia', { features: [{ name: 'forced-colors', value: 'active' }] });
    await assertEval(cdp, 'matchMedia("(forced-colors: active)").matches', 'forced-colors media query is active');
    await cdp.send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-reduced-motion', value: 'reduce' }] });
    await assertEval(cdp, 'matchMedia("(prefers-reduced-motion: reduce)").matches', 'reduced motion media query is active');
    await cdp.send('Emulation.setEmulatedMedia', { features: [] });
    const screenshot = await cdp.send<{ data: string }>('Page.captureScreenshot', { format: 'png' });
    if (screenshot.data.length < 1000) throw new Error('UI smoke failed: daemon screenshot is unexpectedly small');
    console.error('taskpane assertions');
    await setViewport(cdp, 320, 720, true);
    await cdp.send('Page.navigate', { url: `${config.addin.origin}/taskpane.html` });
    await waitFor(cdp, 'document.querySelector(".taskpane-shell") !== null');
    await assertEval(cdp, 'document.documentElement.scrollWidth <= 320', 'taskpane 320px viewport has no horizontal overflow');
    await assertEval(cdp, 'document.querySelector("#settingsToggle").getAttribute("aria-label") === "Open Settings"', 'taskpane settings button is named');
    await assertEval(cdp, 'document.querySelector("#serverVersion") !== null && document.querySelector("#protocolVersion").textContent.trim().length > 0', 'taskpane exposes server and protocol fields');
    await assertEval(cdp, 'document.querySelector("#hostPlatform") !== null && document.querySelector("#documentState") !== null && document.querySelector("#connectionDetail") !== null', 'taskpane exposes host document and connection detail fields');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#settingsToggle").click()' });
    await waitFor(cdp, '!document.querySelector("#settingsPanel").hidden');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#endpointInput").value = "http://localhost:8765/addin"; document.querySelector("#settingsForm").dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }))' });
    await waitFor(cdp, 'document.querySelector("#endpointError").textContent.includes("wss://localhost")');
    await assertEval(cdp, 'document.activeElement === document.querySelector("#endpointInput")', 'taskpane invalid endpoint focuses endpoint field');
    const taskpaneScreenshot = await cdp.send<{ data: string }>('Page.captureScreenshot', { format: 'png' });
    if (taskpaneScreenshot.data.length < 1000) throw new Error('UI smoke failed: taskpane screenshot is unexpectedly small');
    console.log(JSON.stringify({ ok: true, url: `${config.addin.origin}/ui/` }));
  } finally {
    clearTimeout(watchdog);
    await cdp?.close();
    if (chrome) await stopProcess(chrome);
    await daemon.close();
    await removeDirBestEffort(userDataDir);
  }
}

function seedUiState(daemon: Awaited<ReturnType<typeof startDaemon>>): void {
  const connection = fakeConnection();
  daemon.registry.registerRuntime(connection, connection.runtime);
  daemon.registry.addSession(connection, {
    session_id: '11111111-1111-4111-8111-111111111111',
    instance_id: connection.runtime.instance_id,
    document: { title: 'Smoke.docx', filename: 'Smoke.docx', is_dirty: false, is_read_only: false, protection: { kind: null, rights: null, rights_source: 'unavailable' } },
    available_tools: ['word.get_text', 'word.insert_paragraph'],
    is_active: true
  });
  const uiState = daemon.uiState;
  const clientId = uiState.registerClient({ client_id: 'client-1', transport: 'http', name: 'copilot-cli/1.0' });
  for (let index = 0; index < 3; index += 1) {
    const commandId = uiState.startCommand({ client_id: clientId, client_name: 'copilot-cli/1.0', session_id: '11111111-1111-4111-8111-111111111111', host_app: 'word', tool: index === 1 ? 'word.insert_paragraph' : 'word.get_text', user_intent: 'high-level smoke task', timeout_ms: 30000 });
    uiState.finishCommand(commandId, index === 2 ? { ok: false, error: { office_mcp_code: 'IRM_DENIED', message: 'The document blocked the edit.', tool: 'word.insert_paragraph', retriable: false, partial_effect: 'none' } } : { ok: true, data: { redacted: true } });
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

async function launchChrome(userDataDir: string, port: number): Promise<ChildProcessWithoutNullStreams> {
  const chrome = spawn(chromePath, [`--remote-debugging-port=${port}`, '--remote-allow-origins=*', `--user-data-dir=${userDataDir}`, '--headless=new', '--disable-gpu', '--no-first-run', '--ignore-certificate-errors', 'about:blank']);
  chrome.stderr.setEncoding('utf8');
  await waitForHttp(`http://127.0.0.1:${port}/json/version`);
  return chrome;
}

async function openPage(port: number, url: string): Promise<{ webSocketDebuggerUrl: string }> {
  return JSON.parse(await httpRequestText(`http://127.0.0.1:${port}/json/new?${encodeURIComponent(url)}`, 'PUT')) as { webSocketDebuggerUrl: string };
}

async function waitForHttp(url: string): Promise<void> {
  for (let index = 0; index < 50; index += 1) {
    try { await httpRequestText(url); return; } catch { await delay(100); }
  }
  throw new Error(`Timed out waiting for ${url}.`);
}

function httpRequestText(url: string, method = 'GET'): Promise<string> {
  return new Promise((resolve, reject) => {
    const req = httpRequest(url, { method }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      res.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    });
    req.on('error', reject);
    req.end();
  });
}

async function stopProcess(process: ChildProcessWithoutNullStreams): Promise<void> {
  if (process.exitCode !== null || process.killed) return;
  process.kill();
  await Promise.race([
    new Promise<void>((resolve) => process.once('exit', () => resolve())),
    delay(3000)
  ]);
}

async function removeDirBestEffort(path: string): Promise<void> {
  for (let index = 0; index < 5; index += 1) {
    try {
      rmSync(path, { recursive: true, force: true });
      return;
    } catch {
      await delay(250);
    }
  }
}

async function setViewport(cdp: CdpClient, width: number, height: number, mobile: boolean): Promise<void> {
  await cdp.send('Emulation.setDeviceMetricsOverride', { width, height, deviceScaleFactor: 1, mobile });
}

async function waitFor(cdp: CdpClient, expression: string): Promise<void> {
  for (let index = 0; index < 50; index += 1) {
    const value = await cdp.evaluate<boolean>(expression);
    if (value) return;
    await delay(100);
  }
  throw new Error(`Timed out waiting for expression: ${expression}`);
}

async function assertEval(cdp: CdpClient, expression: string, label: string): Promise<void> {
  const value = await cdp.evaluate<boolean>(expression);
  if (!value) throw new Error(`UI smoke failed: ${label}`);
}

class CdpClient {
  private nextId = 1;
  private readonly pending = new Map<number, { resolve: (value: unknown) => void; reject: (error: Error) => void }>();

  private constructor(private readonly socket: WebSocket) {
    socket.on('message', (raw) => {
      const message = JSON.parse(String(raw)) as CdpResult<unknown>;
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  static connect(url: string): Promise<CdpClient> {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Timed out connecting to Chrome DevTools.')), 10000);
      const socket = new WebSocket(url, { rejectUnauthorized: false });
      socket.once('open', () => { clearTimeout(timer); resolve(new CdpClient(socket)); });
      socket.once('error', (error) => { clearTimeout(timer); reject(error); });
    });
  }

  send<T>(method: string, params: Record<string, unknown> = {}): Promise<T> {
    const id = this.nextId++;
    this.socket.send(JSON.stringify({ id, method, params }));
    return new Promise((resolve, reject) => this.pending.set(id, { resolve: resolve as (value: unknown) => void, reject }));
  }

  async evaluate<T>(expression: string): Promise<T> {
    const response = await this.send<{ result: { value: T } }>('Runtime.evaluate', { expression, returnByValue: true });
    return response.result.value;
  }

  close(): Promise<void> {
    return new Promise((resolve) => {
      this.socket.once('close', () => resolve());
      this.socket.close(1000, 'done');
    });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
