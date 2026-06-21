import { execFileSync, spawn, type ChildProcess, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { createServer } from 'node:net';
import { request as httpRequest } from 'node:http';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { setTimeout as delay } from 'node:timers/promises';
import WebSocket from 'ws';

const chromePath = process.env.OFFICE_MCP_CHROME_PATH || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const cargoCommand = process.env.CARGO || resolve(process.env.USERPROFILE || '', '.cargo/bin/cargo.exe');
const daemonExeOverride = process.env.OFFICE_MCP_DAEMON_EXE;

type CdpResult<T> = { id: number; result?: T; error?: { message: string } };

async function main(): Promise<void> {
  const watchdog = setTimeout(() => {
    console.error('UI smoke timed out.');
    process.exit(1);
  }, 180000);
  const fixtureRoot = mkdtempSync(join(tmpdir(), 'office-mcp-ui-runtime-'));
  const cargoTargetDir = join(fixtureRoot, 'target');
  const runtimePath = join(fixtureRoot, 'ui-runtime.json');
  if (!daemonExeOverride) buildRustUiFixture(cargoTargetDir);
  const daemon = startRustUiFixture(runtimePath, cargoTargetDir);
  let chrome: ChildProcessWithoutNullStreams | undefined;
  let cdp: CdpClient | undefined;
  const userDataDir = mkdtempSync(join(tmpdir(), 'office-mcp-ui-chrome-'));
  try {
    const runtime = await waitForRuntimeFile(runtimePath, daemon);
    console.error('launch chrome');
    const chromePort = await reserveTcpPort();
    chrome = await launchChrome(userDataDir, chromePort);
    console.error('open page');
    const page = await openPage(chromePort, runtime.uiUrl);
    console.error('connect cdp');
    cdp = await CdpClient.connect(page.webSocketDebuggerUrl);
    await cdp.send('Security.setIgnoreCertificateErrors', { ignore: true });
    await cdp.send('Page.enable');
    await cdp.send('Runtime.enable');
    const pageDiagnostics = new PageDiagnostics();
    cdp.onMessage((message) => pageDiagnostics.record(message));
    console.error('desktop assertions');
    await setViewport(cdp, 1366, 768, false);
    await cdp.send('Page.navigate', { url: runtime.uiUrl });
    await waitFor(cdp, 'document.querySelectorAll("table").length >= 2');
    await assertEval(cdp, 'document.querySelector("#healthBadge").textContent.includes("Degraded")', 'degraded health badge renders');
    await assertEval(cdp, 'document.querySelector("#lastError").textContent.includes("Certificate reload failed")', 'degraded last error renders');
    await assertEval(cdp, 'document.querySelector("#clients .row") !== null && document.querySelector("#clients").textContent.includes("http")', 'client list renders');
    await assertEval(cdp, 'document.querySelector("#documents .row.word") !== null', 'word document row renders');
    await assertEval(cdp, 'document.querySelector("#documents").textContent.includes("Excel") && document.querySelector("#documents .row.excel") !== null', 'excel document group renders');
    await assertEval(cdp, '(() => { const text = document.querySelector("#documents .row.excel")?.textContent || ""; return text.includes("Dead") && !/stale|reconnecting/i.test(text); })()', 'stale document card renders as dead without reconnect wording');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.classList.contains("document-card")', 'word document card uses compact card class');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.textContent.includes("Runtime Evidence.docx")', 'word document card shows document name');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.textContent.includes("Active")', 'word document card shows active state');
    await assertEval(cdp, '(() => { const text = document.querySelector("#documents .row.word")?.textContent || ""; return text.includes("Session") && text.includes("11111111...") && !text.includes("11111111-1111-4111-8111-111111111111"); })()', 'word document card shows bounded session id');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.textContent.includes("Version 16.0")', 'word document card shows version');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.textContent.includes("2 tools")', 'word document card shows available tool count');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.textContent.includes("Queue 0")', 'word document card shows queue depth');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.textContent.includes("Finished 2")', 'word document card shows finished task count');
    await assertEval(cdp, 'document.querySelector("#documents .row.word")?.textContent.includes("Failed 5")', 'word document card shows failed task count');
    await assertEval(cdp, 'document.querySelector("#documents .doc-history") === null && !document.querySelector("#documents").textContent.includes("Show details") && !document.querySelector("#documents").textContent.includes("Hide details")', 'document list does not duplicate activity history');
    await assertEval(cdp, '(() => { const title = document.querySelector("#documents .row.word .document-card-title strong")?.getBoundingClientRect(); const card = document.querySelector("#documents .row.word")?.getBoundingClientRect(); return title && card && title.width >= Math.min(120, card.width * 0.45); })()', 'document title keeps primary card width');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#appFilter").value = "excel"; document.querySelector("#appFilter").dispatchEvent(new Event("change"))' });
    await waitFor(cdp, 'document.querySelector("#documents .row.excel") !== null && document.querySelector("#documents .row.word") === null');
    await assertEval(cdp, 'document.querySelector("#documents").textContent.includes("Excel") && !document.querySelector("#documents").textContent.includes("Runtime Evidence.docx")', 'document app filter limits visible sessions');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#appFilter").value = "powerpoint"; document.querySelector("#appFilter").dispatchEvent(new Event("change"))' });
    await waitFor(cdp, 'document.querySelector("#documents").textContent.includes("No matching documents")');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#appFilter").value = "all"; document.querySelector("#appFilter").dispatchEvent(new Event("change"))' });
    await waitFor(cdp, 'document.querySelector("#documents .row.word") !== null && document.querySelector("#documents .row.excel") !== null');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#documents .row.word").focus()' });
    await pressKey(cdp, 'Enter', 13);
    await assertEval(cdp, '[...document.querySelectorAll(".metrics dt")].some((node) => node.textContent.trim() === "Active Tasks") && ![...document.querySelectorAll(".metrics dt")].some((node) => node.textContent.trim() === "Running")', 'top metrics use explicit active task label');
    await assertEval(cdp, 'document.activeElement === document.querySelector("#documents .row.word") && document.querySelector("#inspector").textContent.includes("Runtime Evidence.docx")', 'keyboard inspection preserves document row focus');
    await assertEval(cdp, 'document.querySelector("#currentTasks").textContent.includes("Running") && document.querySelector("#taskCount").textContent.trim() === "1"', 'running task state renders');
    await assertEval(cdp, 'document.querySelector("#history").textContent.includes("Success") && document.querySelector("#history").textContent.includes("Failure") && document.querySelector("#history").textContent.includes("TIMEOUT") && document.querySelector("#history").textContent.includes("CANCELLED")', 'history renders success failure timeout and cancelled states');
    await assertEval(cdp, 'document.querySelector("#history tr[data-inspect]").getAttribute("tabindex") === "0" && document.querySelector("#history tr[data-inspect]").getAttribute("role") === "button"', 'history table rows are keyboard focusable buttons');
    await assertEval(cdp, 'document.querySelector("#daemonVersion").textContent.trim().length > 0 && document.querySelector("#daemonUptime").textContent.trim().length > 0', 'daemon details show version and uptime');
    await assertEval(cdp, 'document.querySelector("#configPath").textContent.includes("config.toml") && !document.querySelector("#configPath").textContent.includes("Not configured")', 'daemon details show effective config file path');
    await assertEval(cdp, 'document.querySelector("[data-copy=\\"configPath\\"]").getAttribute("aria-label") === "Copy config path" && document.querySelector("#configPath").closest(".detail-copy") !== null', 'config path has a copy affordance');
    await assertEval(cdp, '(() => { const columns = getComputedStyle(document.querySelector(".details dl")).gridTemplateColumns.trim().split(/\\s+/); return columns.length === 5 && parseFloat(columns[4]) > parseFloat(columns[0]) && parseFloat(columns[4]) >= parseFloat(columns[2]); })()', 'daemon details reserve wider space for last error than compact metadata');
    await assertEval(cdp, 'document.querySelector("#lastError").textContent.trim().length > 0', 'daemon details show last error state');
    await assertEval(cdp, 'document.querySelector(".status-strip > .details") !== null && document.querySelector(".status-strip").nextElementSibling.classList.contains("workspace")', 'daemon details are grouped inside the compact status strip');
    await assertEval(cdp, '(() => { const header = document.querySelector(".status-strip").getBoundingClientRect(); const details = document.querySelector(".details").getBoundingClientRect(); const workspace = document.querySelector(".workspace").getBoundingClientRect(); return details.height <= 56 && workspace.top - header.bottom <= 12; })()', 'daemon header avoids detached oversized detail block');
    await assertEval(cdp, 'getComputedStyle(document.querySelector(".workspace")).gridTemplateColumns.split(" ").length >= 3', 'desktop layout uses three columns');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#documents .row.word").click()' });
    await assertEval(cdp, 'document.querySelector("#documents .doc-history") === null', 'document click does not open inline command history');
    await waitFor(cdp, 'document.querySelector("#inspector").textContent.includes("Runtime Evidence.docx")');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#clearInspector").click()' });
    await waitFor(cdp, 'document.querySelector("#inspector").textContent.trim() === "Select a row."');
    await cdp.send('Runtime.evaluate', { expression: 'window.__copiedText = null; Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async (text) => { window.__copiedText = text; } } }); document.querySelector("[data-copy=\\"mcpEndpoint\\"]").click();' });
    await waitFor(cdp, 'window.__copiedText && window.__copiedText.startsWith("http://127.0.0.1:")');
    await assertEval(cdp, 'document.querySelector("#announcer").textContent.includes("Copied MCP")', 'copy control announces copied MCP endpoint');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#resultFilter").value = "failure"; document.querySelector("#resultFilter").dispatchEvent(new Event("change"))' });
    await waitFor(cdp, 'document.querySelector("#history").textContent.includes("IRM_DENIED")');
    await assertEval(cdp, 'document.querySelector("#history").textContent.includes("IRM_DENIED")', 'failure details expose office_mcp_code');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#history tr[data-inspect]").focus()' });
    await pressKey(cdp, 'Enter', 13);
    await waitFor(cdp, 'document.querySelector("#inspector").textContent.includes("IRM_DENIED")');
    await assertEval(cdp, 'document.activeElement === document.querySelector("#history tr[data-inspect]")', 'keyboard table inspection preserves row focus');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#resultFilter").value = "timeout"; document.querySelector("#resultFilter").dispatchEvent(new Event("change"))' });
    await waitFor(cdp, 'document.querySelector("#history").textContent.includes("TIMEOUT")');
    await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#resultFilter").value = "cancelled"; document.querySelector("#resultFilter").dispatchEvent(new Event("change"))' });
    await waitFor(cdp, 'document.querySelector("#history").textContent.includes("Cancelled")');
    await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Tab', windowsVirtualKeyCode: 9, nativeVirtualKeyCode: 9 });
    await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Tab', windowsVirtualKeyCode: 9, nativeVirtualKeyCode: 9 });
    await assertEval(cdp, 'document.activeElement !== document.body', 'keyboard tab reaches a focusable control');
    console.error('narrow assertions');
    await setViewport(cdp, 320, 720, true);
    await assertEval(cdp, 'document.documentElement.scrollWidth <= 320', '320px viewport has no horizontal overflow');
    await assertEval(cdp, 'getComputedStyle(document.querySelector(".workspace")).gridTemplateColumns.split(" ").length === 1', 'narrow layout stacks columns');
    await assertEval(cdp, 'getComputedStyle(document.querySelector(".details dl")).gridTemplateColumns.split(" ").length === 1', 'daemon details stack on narrow layout');
    await assertEval(cdp, 'document.querySelector(".status-strip > .details") !== null && document.querySelector(".details").getBoundingClientRect().top >= document.querySelector(".identity").getBoundingClientRect().bottom', 'narrow daemon details remain attached below identity metrics');
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
    await assertTaskpane(cdp, pageDiagnostics, `${runtime.origin}/word/taskpane.html`, 'Word', 3);
    await assertTaskpane(cdp, pageDiagnostics, `${runtime.origin}/excel/taskpane.html`, 'Excel', 3);
    await assertTaskpane(cdp, pageDiagnostics, `${runtime.origin}/powerpoint/taskpane.html`, 'PowerPoint', 3);
    console.error('empty-state assertions');
    const emptyRuntimePath = join(mkdtempSync(join(tmpdir(), 'office-mcp-ui-empty-runtime-')), 'ui-runtime.json');
    const emptyDaemon = startRustUiFixture(emptyRuntimePath, cargoTargetDir, 'empty');
    try {
      const emptyRuntime = await waitForRuntimeFile(emptyRuntimePath, emptyDaemon);
      await cdp.send('Page.navigate', { url: emptyRuntime.uiUrl });
      await waitFor(cdp, 'document.querySelector("#healthBadge").textContent.includes("Up")');
      await assertEval(cdp, '[...document.querySelectorAll(".metrics dt")].some((node) => node.textContent.trim() === "Active Tasks") && document.querySelector("#clientCount").textContent.trim() === "0" && document.querySelector("#documentCount").textContent.trim() === "0" && document.querySelector("#taskCount").textContent.trim() === "0"', 'empty state counters render zero with explicit task label');
      await assertEval(cdp, 'document.querySelector("#clients").textContent.includes("No MCP clients connected")', 'empty client state renders');
      await assertEval(cdp, 'document.querySelector("#documents").textContent.includes("No documents connected")', 'empty document state renders');
      await assertEval(cdp, 'document.querySelector("#currentTasks").textContent.includes("No command is running") && document.querySelector("#history").textContent.includes("No command history yet")', 'empty task and history states render');
    } finally {
      await stopProcess(emptyDaemon);
      await removeDirBestEffort(dirname(emptyRuntimePath));
    }
    console.log(JSON.stringify({ ok: true, url: runtime.uiUrl }));
  } finally {
    clearTimeout(watchdog);
    await cdp?.close();
    if (chrome) await stopProcess(chrome);
    await stopProcess(daemon);
    await removeDirBestEffort(userDataDir);
    await removeDirBestEffort(dirname(runtimePath));
  }
}

function buildRustUiFixture(cargoTargetDir: string): void {
  execFileSync(cargoCommand, ['build', '-q', '-p', 'office-mcp-daemon'], {
    cwd: repoRoot,
    env: { ...process.env, CARGO_TARGET_DIR: cargoTargetDir },
    stdio: 'inherit'
  });
}

function startRustUiFixture(runtimePath: string, cargoTargetDir: string, state = 'seeded'): ChildProcess {
  const child = spawn(daemonExecutablePath(cargoTargetDir), ['evidence', 'ui-fixture'], {
    cwd: repoRoot,
    env: { ...process.env, OFFICE_MCP_UI_RUNTIME_PATH: runtimePath, OFFICE_MCP_UI_FIXTURE_STATE: state },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', (chunk) => process.stderr.write(chunk));
  child.stdout.setEncoding('utf8');
  child.stdout.on('data', (chunk) => process.stderr.write(chunk));
  child.on('error', (error) => process.stderr.write(`Failed to start Rust UI fixture: ${error.message}\n`));
  return child;
}

function daemonExecutablePath(cargoTargetDir: string): string {
  return daemonExeOverride || join(cargoTargetDir, 'debug', process.platform === 'win32' ? 'office-mcp-daemon.exe' : 'office-mcp-daemon');
}

async function waitForRuntimeFile(path: string, daemon: ChildProcess): Promise<{ origin: string; stateUrl: string; uiUrl: string }> {
  for (let index = 0; index < 100; index += 1) {
    if (existsSync(path)) {
      const runtime = JSON.parse(readFileSync(path, 'utf8')) as { origin?: string; stateUrl?: string; uiUrl?: string };
      if (runtime.origin && runtime.stateUrl && runtime.uiUrl) return { origin: runtime.origin, stateUrl: runtime.stateUrl, uiUrl: runtime.uiUrl };
    }
    if (daemon.exitCode !== null) throw new Error(`Rust UI fixture exited before writing runtime file with code ${daemon.exitCode}.`);
    await delay(100);
  }
  throw new Error('Timed out waiting for Rust UI fixture runtime file.');
}

async function launchChrome(userDataDir: string, port: number): Promise<ChildProcessWithoutNullStreams> {
  const chrome = spawn(chromePath, [`--remote-debugging-port=${port}`, '--remote-allow-origins=*', `--user-data-dir=${userDataDir}`, '--headless=new', '--disable-gpu', '--no-first-run', '--ignore-certificate-errors', 'about:blank']);
  chrome.stderr.setEncoding('utf8');
  let stderr = '';
  chrome.stderr.on('data', (chunk) => { stderr += String(chunk); });
  try {
    await waitForHttp(`http://127.0.0.1:${port}/json/version`, () => chrome.exitCode !== null ? `Chrome exited with code ${chrome.exitCode}. ${stderr.slice(-2000)}` : stderr.slice(-2000));
    return chrome;
  } catch (error) {
    await stopProcess(chrome);
    throw error;
  }
}

async function openPage(port: number, url: string): Promise<{ webSocketDebuggerUrl: string }> {
  return JSON.parse(await httpRequestText(`http://127.0.0.1:${port}/json/new?${encodeURIComponent(url)}`, 'PUT')) as { webSocketDebuggerUrl: string };
}

async function waitForHttp(url: string, diagnostics = () => ''): Promise<void> {
  for (let index = 0; index < 100; index += 1) {
    try { await httpRequestText(url); return; } catch { await delay(100); }
  }
  const detail = diagnostics();
  throw new Error(`Timed out waiting for ${url}.${detail ? ` ${detail}` : ''}`);
}

function reserveTcpPort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      if (!address || typeof address === 'string') {
        server.close(() => reject(new Error('Could not reserve a TCP port.')));
        return;
      }
      const port = address.port;
      server.close(() => resolve(port));
    });
  });
}

function httpRequestText(url: string, method = 'GET'): Promise<string> {
  return new Promise((resolve, reject) => {
    const req = httpRequest(url, { method }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      res.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    });
    req.on('error', reject);
    req.setTimeout(1000, () => {
      req.destroy(new Error(`Timed out waiting for HTTP response from ${url}.`));
    });
    req.end();
  });
}

async function stopProcess(process: ChildProcess): Promise<void> {
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

async function pressKey(cdp: CdpClient, key: string, code: number): Promise<void> {
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key, windowsVirtualKeyCode: code, nativeVirtualKeyCode: code });
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key, windowsVirtualKeyCode: code, nativeVirtualKeyCode: code });
}

async function assertTaskpane(cdp: CdpClient, pageDiagnostics: PageDiagnostics, url: string, hostName: string, minimumToolGroups: number): Promise<void> {
  await setViewport(cdp, 320, 720, true);
  await cdp.send('Page.navigate', { url });
  await waitFor(cdp, 'document.querySelector(".taskpane-shell") !== null');
  await waitFor(
    cdp,
    'window.__OFFICE_MCP_TASKPANE_READY__ === true',
    () => pageDiagnostics.summary()
  );
  await assertEval(cdp, 'document.documentElement.scrollWidth <= 320', `${hostName} taskpane 320px viewport has no horizontal overflow`);
  await assertEval(cdp, 'document.querySelector("#saveEndpoint").getAttribute("aria-label") === "Reconnect daemon" && document.querySelector("#saveEndpoint").getAttribute("title") === "Reconnect daemon"', `${hostName} taskpane reconnect button has accessible name and tooltip`);
  await assertEval(cdp, '(() => { const buttons = [...document.querySelectorAll(".metadata-copy")]; return buttons.length >= 1 && buttons.every((button) => { const title = button.getAttribute("title"); return Boolean(title && title !== "-"); }); })()', `${hostName} taskpane metadata copy buttons have useful tooltips`);
  await assertEval(cdp, 'document.querySelector("#runtimeVersions") !== null && document.querySelector("#runtimeVersions").textContent.includes("Server") && document.querySelector("#runtimeVersions").textContent.includes("Protocol")', `${hostName} taskpane combines server and protocol metadata`);
  await assertEval(cdp, 'document.querySelectorAll("#serverVersion, #protocolVersion").length === 2 && document.querySelector("#serverVersion").closest("dd") === document.querySelector("#protocolVersion").closest("dd")', `${hostName} taskpane server and protocol share one metadata row`);
  await assertEval(cdp, 'document.querySelector("#hostPlatform") !== null && document.querySelector("#documentState") !== null && document.querySelector("#connectionDetail") !== null', `${hostName} taskpane exposes host document and connection detail fields`);
  await assertEval(cdp, '(() => { const buttons = [...document.querySelectorAll(".metadata-copy")]; return buttons.every((button) => button.title.length >= button.textContent.trim().length); })()', `${hostName} taskpane metadata copy tooltips are at least as complete as visible metadata`);
  await assertEval(cdp, '!/Dirty:\\s*unknown|Read-only:\\s*unknown/i.test(document.querySelector("#documentState").textContent)', `${hostName} taskpane avoids unknown dirty/read-only state`);
  await assertEval(cdp, 'document.querySelector(".tools-panel summary").textContent.includes("Tools") && !document.body.textContent.includes("Available Tools") && !document.body.textContent.includes("Tool Permissions")', `${hostName} taskpane merges available tools and permissions into one surface`);
  await assertEval(cdp, `document.querySelectorAll("#toolList").length === 1 && document.querySelectorAll("#toolPermissionList").length === 0 && document.querySelectorAll(".tool-group").length >= ${minimumToolGroups}`, `${hostName} taskpane renders one grouped tools surface`);
  await assertEval(cdp, '[...document.querySelectorAll(".tool-group")].every((group) => group.tagName === "DETAILS" && /\\d+\\/\\d+/.test(group.querySelector("summary")?.textContent || "") && !/Enabled/.test(group.querySelector("summary")?.textContent || ""))', `${hostName} taskpane tool categories are collapsible with compact counts`);
  await assertEval(cdp, 'document.querySelector(".daemon-endpoint-form") !== null && document.querySelector("#endpointInput").type === "url" && document.querySelector("#endpointInput").name === "daemonEndpoint"', `${hostName} taskpane inline daemon row exposes endpoint URL field`);
  await assertEval(cdp, '(() => { const form = document.querySelector(".daemon-endpoint-form"); const row = form.closest("dd"); const rect = form.getBoundingClientRect(); return row && rect.height > 0 && rect.height < 48 && row.textContent.includes("wss://localhost"); })()', `${hostName} taskpane daemon endpoint row stays compact and inline`);
  await assertEval(cdp, 'document.querySelector("#endpointInput").placeholder.includes("wss://localhost") && document.querySelector("#endpointError") === null && document.querySelector("#connectionDetail") !== null', `${hostName} taskpane endpoint validation uses last error row`);
  await assertEval(cdp, 'document.querySelector("#settingsForm").closest(".summary-panel") !== null', `${hostName} taskpane settings are inline in the summary panel`);
  await assertEval(cdp, '!document.querySelector(".tools-panel").open && document.querySelector(".tools-panel").getBoundingClientRect().height < 48', `${hostName} tools surface stays compact while collapsed`);
  await assertEval(cdp, 'document.querySelector("#currentTaskHeading").getBoundingClientRect().top < 520 && document.querySelector("#historyHeading").getBoundingClientRect().top < 680', `${hostName} taskpane first viewport shows current and recent task regions`);
  await assertEval(cdp, '(() => { const shell = document.querySelector(".taskpane-shell"); const rects = [...shell.children].map((child) => child.getBoundingClientRect()); return rects.every((rect, index) => index === 0 || rect.top - rects[index - 1].bottom <= 12); })()', `${hostName} taskpane avoids large vertical gaps between sections`);
  await cdp.send('Runtime.evaluate', { expression: 'document.querySelector("#endpointInput").focus()' });
  await waitFor(cdp, 'document.activeElement === document.querySelector("#endpointInput")');
  await assertEval(cdp, '(() => { const summary = document.querySelector(".summary-panel"); const form = document.querySelector(".daemon-endpoint-form"); const task = document.querySelector(".current-task-panel"); const summaryRect = summary.getBoundingClientRect(); const formRect = form.getBoundingClientRect(); const taskRect = task.getBoundingClientRect(); return formRect.height > 0 && formRect.top >= summaryRect.top && formRect.bottom <= summaryRect.bottom && taskRect.top - summaryRect.bottom <= 12; })()', `${hostName} inline settings stay inside summary flow without a detached block`);
  await assertEval(cdp, '(() => { const summary = document.querySelector(".summary-panel"); const children = [...summary.children].filter((child) => getComputedStyle(child).display !== "none"); const rects = children.map((child) => child.getBoundingClientRect()).filter((rect) => rect.height > 0); return rects.every((rect, index) => index === 0 || (rect.top >= rects[index - 1].bottom && rect.top - rects[index - 1].bottom <= 12)); })()', `${hostName} summary content has compact non-overlapping vertical spacing`);
  await assertEval(cdp, '(() => { const panels = [...document.querySelectorAll(".panel")].map((panel) => panel.getBoundingClientRect()); return panels.every((rect, index) => index === 0 || rect.top >= panels[index - 1].bottom); })()', `${hostName} taskpane panels do not overlap`);
  await assertEval(cdp, 'document.querySelector("#endpointError") === null && document.querySelector("#connectionDetail") !== null', `${hostName} taskpane keeps endpoint validation on the last error row`);
  const taskpaneScreenshot = await cdp.send<{ data: string }>('Page.captureScreenshot', { format: 'png' });
  if (taskpaneScreenshot.data.length < 1000) throw new Error(`UI smoke failed: ${hostName} taskpane screenshot is unexpectedly small`);
}

async function waitFor(cdp: CdpClient, expression: string, diagnostics = () => ''): Promise<void> {
  for (let index = 0; index < 50; index += 1) {
    const value = await cdp.evaluate<boolean>(expression);
    if (value) return;
    await delay(100);
  }
  const detail = diagnostics();
  throw new Error(`Timed out waiting for expression: ${expression}${detail ? `\n${detail}` : ''}`);
}

async function assertEval(cdp: CdpClient, expression: string, label: string): Promise<void> {
  const value = await cdp.evaluate<boolean>(expression);
  if (!value) throw new Error(`UI smoke failed: ${label}`);
}

class CdpClient {
  private nextId = 1;
  private readonly pending = new Map<number, { resolve: (value: unknown) => void; reject: (error: Error) => void }>();
  private readonly messageHandlers = new Set<(message: Record<string, unknown>) => void>();

  private constructor(private readonly socket: WebSocket) {
    socket.on('message', (raw) => {
      const message = JSON.parse(String(raw)) as CdpResult<unknown> & Record<string, unknown>;
      for (const handler of this.messageHandlers) handler(message);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  onMessage(handler: (message: Record<string, unknown>) => void): void {
    this.messageHandlers.add(handler);
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

class PageDiagnostics {
  private readonly entries: string[] = [];

  record(message: Record<string, unknown>): void {
    if (message.method === 'Runtime.exceptionThrown') {
      const params = message.params as { exceptionDetails?: { text?: string; exception?: { description?: string } } } | undefined;
      const detail = params?.exceptionDetails;
      this.entries.push(`exception: ${detail?.exception?.description || detail?.text || 'unknown'}`);
    }
    if (message.method === 'Page.javascriptDialogOpening') {
      const params = message.params as { message?: string } | undefined;
      this.entries.push(`dialog: ${params?.message || 'unknown'}`);
    }
  }

  summary(): string {
    return this.entries.slice(-10).join('\n');
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
