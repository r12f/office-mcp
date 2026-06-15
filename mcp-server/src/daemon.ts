import { readFileSync } from 'node:fs';
import { createServer as createHttpServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { createServer as createHttpsServer } from 'node:https';
import { basename, dirname, extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { DaemonConfig } from './config.js';
import { assertHttpsConfig } from './config.js';
import { createAddinChannel, handleUpgrade } from './addin-channel.js';
import { McpFrontend } from './mcp-server.js';
import { SessionRegistry } from './session-registry.js';
import { createLogger } from './logger.js';
import { authorizeUiRequest, createUiToken, defaultUiRuntimePath, removeUiRuntimeFile, UiStateStore, writeUiRuntimeFile } from './ui-state.js';
import { SERVER_VERSION } from './types.js';

const PUBLIC_DIR = resolveAddinPublicDir();

export function resolveAddinPublicDir(moduleUrl = import.meta.url, installRoot = process.env.OFFICE_MCP_INSTALL_ROOT): string {
  if (installRoot) return join(installRoot, 'addin', 'public');

  const moduleDir = dirname(fileURLToPath(moduleUrl));
  const parentDir = dirname(moduleDir);
  const serverRoot = basename(parentDir) === 'dist' ? dirname(parentDir) : parentDir;
  return join(dirname(serverRoot), 'addin', 'public');
}

export type RunningDaemon = {
  close: () => Promise<void>;
  registry: SessionRegistry;
  uiToken: string;
  uiState: UiStateStore;
};

export async function startDaemon(config: DaemonConfig): Promise<RunningDaemon> {
  assertHttpsConfig(config);
  const logger = createLogger(config);
  const uiToken = createUiToken();
  let registry: SessionRegistry;
  const uiState = new UiStateStore({
    version: SERVER_VERSION,
    mcpEndpoint: `http://${config.mcp.host}:${config.mcp.port}/mcp`,
    addinEndpoint: `${config.addin.origin}/addin`,
    logPath: config.logging.file || null,
    sessions: () => registry.listSessions()
  });
  registry = new SessionRegistry(config.addin.maxPendingPerSession, uiState);
  const frontend = new McpFrontend(config, registry, uiState);
  const wss = createAddinChannel(config, registry);

  const mcpServer = createHttpServer(async (req, res) => {
    if (!req.url?.startsWith('/mcp')) {
      json(res, 404, { error: 'Not found' });
      return;
    }
    try {
      const body = req.method === 'POST' ? await readJsonWithLimit(req, config.limits.maxRequestBytes) : undefined;
      await frontend.handle(req, res, body);
    } catch (error) {
      if (!res.headersSent) {
        const tooLarge = error instanceof RequestTooLargeError;
        json(res, tooLarge ? 413 : 400, {
          jsonrpc: '2.0',
          id: null,
          error: { code: -32000, message: error instanceof Error ? error.message : 'Bad request' }
        });
      }
    }
  });

  const addinServer = createHttpsServer({
    pfx: readFileSync(config.addin.pfxPath),
    passphrase: config.addin.pfxPassphrase
  }, (req, res) => {
    if (handleUiRequest(req, res, uiState, uiToken, config)) return;
    serveStatic(req, res);
  });

  addinServer.on('upgrade', (req, socket, head) => handleUpgrade(config, wss, req, socket, head));

  await Promise.all([
    listen(mcpServer, config.mcp.port, config.mcp.host),
    listen(addinServer, config.addin.port, config.addin.host)
  ]);

  const pruneTimer = setInterval(() => registry.pruneStaleSessions(config.addin.sessionGraceSec), 5000);
  const uiRuntimePath = process.env.OFFICE_MCP_UI_RUNTIME_PATH || defaultUiRuntimePath();
  writeUiRuntimeFile(uiRuntimePath, {
    origin: config.addin.origin,
    stateUrl: `${config.addin.origin}/ui/state`,
    uiUrl: `${config.addin.origin}/ui/`,
    token: uiToken,
    pid: process.pid,
    createdAt: new Date().toISOString()
  });

  logger.info(`office-mcp MCP endpoint: http://${config.mcp.host}:${config.mcp.port}/mcp`, { component: 'daemon', event: 'mcp_listening', host: config.mcp.host, port: config.mcp.port });
  logger.info(`office-mcp add-in origin: ${config.addin.origin}`, { component: 'daemon', event: 'addin_listening', origin: config.addin.origin });

  return {
    registry,
    uiToken,
    uiState,
    close: async () => {
      clearInterval(pruneTimer);
      await frontend.close();
      wss.close();
      await Promise.all([closeServer(mcpServer), closeServer(addinServer)]);
      removeUiRuntimeFile(uiRuntimePath);
      logger.info('office-mcp daemon stopped', { component: 'daemon', event: 'stopped' });
    }
  };
}

function handleUiRequest(req: IncomingMessage, res: ServerResponse, uiState: UiStateStore, uiToken: string, config: DaemonConfig): boolean {
  const url = new URL(req.url ?? '/', config.addin.origin);
  if (url.pathname === '/ui' || url.pathname === '/ui/') {
    res.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' });
    res.end(renderDaemonUiShell(config, uiToken));
    return true;
  }
  if (url.pathname === '/ui/app.js') {
    res.writeHead(200, { 'content-type': 'text/javascript; charset=utf-8', 'cache-control': 'no-store' });
    res.end(DAEMON_UI_JS);
    return true;
  }
  if (url.pathname === '/ui/app.css') {
    res.writeHead(200, { 'content-type': 'text/css; charset=utf-8', 'cache-control': 'no-store' });
    res.end(DAEMON_UI_CSS);
    return true;
  }
  if (url.pathname !== '/ui/state') return false;
  const origin = req.headers.origin;
  if (origin && origin !== config.addin.origin) {
    res.writeHead(403).end('Forbidden origin');
    return true;
  }
  if (!authorizeUiRequest(req.headers, uiToken)) {
    res.writeHead(401, { 'WWW-Authenticate': 'Bearer' }).end('Unauthorized');
    return true;
  }
  json(res, 200, uiState.snapshot());
  return true;
}

function renderDaemonUiShell(config: DaemonConfig, uiToken: string): string {
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta name="theme-color" content="#F7F8FA" />
    <title>Office MCP</title>
    <link rel="stylesheet" href="/ui/app.css" />
    <script>window.__OFFICE_MCP_UI__ = ${JSON.stringify({ token: uiToken, stateUrl: `${config.addin.origin}/ui/state` })};</script>
    <script defer src="/ui/app.js"></script>
  </head>
  <body>
    <a class="skip-link" href="#main">Skip to Activity</a>
    <main id="main" class="shell">
      <header class="status-strip" aria-label="Daemon Status">
        <div><span id="health" class="badge neutral">Loading…</span><h1>Office MCP</h1></div>
        <div class="strip-metrics"><span id="clientCount">0 Clients</span><span id="documentCount">0 Documents</span><span id="taskCount">0 Running</span></div>
      </header>
      <section class="layout">
        <nav class="panel" aria-labelledby="documentsHeading"><h2 id="documentsHeading">Documents</h2><div id="documents"></div></nav>
        <section class="panel activity" aria-labelledby="activityHeading"><h2 id="activityHeading">Activity</h2><h3>Current Tasks</h3><div id="currentTasks"></div><h3>Recent Command History</h3><div id="history"></div></section>
        <aside class="panel inspector" aria-labelledby="inspectorHeading"><h2 id="inspectorHeading">Inspector</h2><pre id="inspector">Select an item.</pre></aside>
      </section>
      <div id="announcer" class="sr-only" aria-live="polite"></div>
    </main>
  </body>
</html>`;
}

const DAEMON_UI_JS = `(() => {
  const boot = window.__OFFICE_MCP_UI__;
  const health = document.getElementById('health');
  const clientCount = document.getElementById('clientCount');
  const documentCount = document.getElementById('documentCount');
  const taskCount = document.getElementById('taskCount');
  const documents = document.getElementById('documents');
  const currentTasks = document.getElementById('currentTasks');
  const history = document.getElementById('history');
  const inspector = document.getElementById('inspector');
  const announcer = document.getElementById('announcer');

  refresh();
  setInterval(refresh, 2000);

  async function refresh() {
    try {
      const response = await fetch(boot.stateUrl, { headers: { 'x-office-mcp-ui-token': boot.token } });
      if (!response.ok) throw new Error('UI state returned ' + response.status);
      render(await response.json());
    } catch (error) {
      health.textContent = 'Down';
      health.className = 'badge danger';
      announcer.textContent = error.message || 'UI state unavailable.';
    }
  }

  function render(snapshot) {
    const docList = Object.values(snapshot.documents || {}).flat();
    health.textContent = title(snapshot.daemon.status);
    health.className = 'badge ' + tone(snapshot.daemon.status);
    clientCount.textContent = count(snapshot.clients.length, 'Client');
    documentCount.textContent = count(docList.length, 'Document');
    taskCount.textContent = count(snapshot.current_tasks.length, 'Running');
    documents.innerHTML = docList.length ? docList.map(documentRow).join('') : '<p class="empty">No documents connected. Open the add-in in Word.</p>';
    currentTasks.innerHTML = snapshot.current_tasks.length ? snapshot.current_tasks.map(commandRow).join('') : '<p class="empty">No command is running.</p>';
    history.innerHTML = snapshot.recent_commands.length ? snapshot.recent_commands.map(commandRow).join('') : '<p class="empty">No command history yet.</p>';
    for (const button of document.querySelectorAll('[data-inspect]')) {
      button.addEventListener('click', () => { inspector.textContent = JSON.stringify(JSON.parse(button.dataset.inspect), null, 2); });
    }
  }

  function documentRow(doc) {
    return '<button class="row word" type="button" data-inspect=' + attr(JSON.stringify(doc)) + '><span>' + esc(doc.document.title || doc.document.filename || 'Untitled') + '</span><small>' + esc(doc.status) + ' · ' + esc(doc.session_id) + '</small></button>';
  }

  function commandRow(command) {
    const error = command.error ? '<small>' + esc(command.error.office_mcp_code + ': ' + command.error.message) + '</small>' : '';
    return '<button class="row" type="button" data-inspect=' + attr(JSON.stringify(command)) + '><span>' + esc(command.tool) + '</span><small>' + esc(title(command.status)) + ' · ' + formatMs(command.elapsed_ms || 0) + '</small>' + error + '</button>';
  }

  function count(value, noun) { return value + ' ' + noun + (value === 1 ? '' : 's'); }
  function title(value) { return String(value || '').replace(/_/g, ' ').replace(/\\b\\w/g, (char) => char.toUpperCase()); }
  function tone(value) { return value === 'up' || value === 'success' ? 'success' : value === 'degraded' || value === 'running' ? 'warning' : value === 'down' || value === 'failure' ? 'danger' : 'neutral'; }
  function formatMs(ms) { return new Intl.NumberFormat(undefined, { maximumFractionDigits: 1 }).format(ms / 1000) + 's'; }
  function esc(value) { return String(value).replace(/[&<>\"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;' })[char]); }
  function attr(value) { return "'" + esc(value).replace(/'/g, '&#39;') + "'"; }
})();`;

const DAEMON_UI_CSS = `:root{color-scheme:light;--canvas:#f7f8fa;--surface:#fff;--raised:#f2f5f8;--text:#17202a;--muted:#5a6673;--border:#d8dee6;--word:#2b579a;--success:#168a45;--warning:#8a5a00;--danger:#c9352b;--focus:#4c8dff}*{box-sizing:border-box}body{margin:0;background:var(--canvas);color:var(--text);font:13px/1.45 "Segoe UI",system-ui,sans-serif}.skip-link{position:fixed;top:8px;left:8px;transform:translateY(-150%);background:var(--surface);padding:6px 10px;border:1px solid var(--border);border-radius:6px}.skip-link:focus-visible{transform:translateY(0)}button:focus-visible,a:focus-visible{outline:2px solid var(--focus);outline-offset:2px}.shell{display:grid;gap:12px;min-height:100vh;padding:12px}.status-strip{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:12px;border:1px solid var(--border);border-radius:8px;background:var(--surface)}h1,h2,h3{margin:0;letter-spacing:0;text-wrap:balance}h1{font-size:18px}.strip-metrics{display:flex;gap:10px;color:var(--muted);font-variant-numeric:tabular-nums}.layout{display:grid;grid-template-columns:minmax(220px,280px) minmax(360px,1fr) minmax(260px,360px);gap:12px;min-height:0}.panel{min-width:0;padding:12px;border:1px solid var(--border);border-radius:8px;background:var(--surface)}.activity{display:grid;align-content:start;gap:10px}.badge{display:inline-flex;gap:6px;align-items:center;width:fit-content;padding:2px 7px;border:1px solid var(--border);border-radius:999px;background:var(--raised);font-variant-numeric:tabular-nums}.badge:before{width:7px;height:7px;border-radius:50%;background:currentColor;content:""}.success{color:var(--success)}.warning{color:var(--warning)}.danger{color:var(--danger)}.neutral{color:var(--muted)}.row{display:grid;width:100%;min-width:0;gap:2px;margin-top:8px;padding:10px;border:1px solid var(--border);border-left:3px solid var(--border);border-radius:7px;background:var(--raised);color:var(--text);text-align:left;cursor:pointer}.row.word{border-left-color:var(--word)}.row:hover{border-color:var(--focus)}.row span,.row small{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.row small,.empty{color:var(--muted)}pre{overflow:auto;white-space:pre-wrap;overflow-wrap:anywhere;font-family:"Cascadia Mono",Consolas,monospace;font-size:12px}.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap}@media(max-width:900px){.layout{grid-template-columns:1fr}.inspector{order:3}}@media(prefers-reduced-motion:reduce){*,*:before,*:after{transition-duration:.01ms!important;animation-duration:.01ms!important;animation-iteration-count:1!important}}@media(prefers-color-scheme:dark){:root{color-scheme:dark;--canvas:#11161c;--surface:#18212b;--raised:#202b36;--text:#f2f5f8;--muted:#b7c0ca;--border:#344250;--focus:#78a8ff}}@media(forced-colors:active){.status-strip,.panel,.row,.badge{border:1px solid CanvasText}}`;

function listen(server: ReturnType<typeof createHttpServer>, port: number, host: string): Promise<void> {
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, host, () => {
      server.off('error', reject);
      resolve();
    });
  });
}

function closeServer(server: ReturnType<typeof createHttpServer>): Promise<void> {
  return new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
}

class RequestTooLargeError extends Error {}

export async function readJsonWithLimit(req: AsyncIterable<Buffer | string>, maxBytes: number): Promise<unknown> {
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += buffer.length;
    if (total > maxBytes) throw new RequestTooLargeError(`Request body exceeds ${maxBytes} bytes.`);
    chunks.push(buffer);
  }
  if (chunks.length === 0) return undefined;
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  const url = new URL(req.url ?? '/', 'https://localhost');
  if (url.pathname === '/healthz') {
    json(res, 200, { ok: true });
    return;
  }
  if (url.pathname === '/assets/icon-32.png' || url.pathname === '/assets/icon-80.png') {
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/l6lJ6wAAAABJRU5ErkJggg==', 'base64');
    res.writeHead(200, { 'content-type': 'image/png', 'cache-control': 'no-store' });
    res.end(png);
    return;
  }
  const pathname = url.pathname === '/' ? '/taskpane.html' : url.pathname;
  const filePath = normalize(join(PUBLIC_DIR, pathname));
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403).end('Forbidden');
    return;
  }
  try {
    const content = readFileSync(filePath);
    res.writeHead(200, { 'content-type': contentType(filePath), 'cache-control': 'no-store' });
    res.end(content);
  } catch {
    res.writeHead(404).end('Not found');
  }
}

function contentType(path: string): string {
  switch (extname(path)) {
    case '.html': return 'text/html; charset=utf-8';
    case '.js': return 'text/javascript; charset=utf-8';
    case '.css': return 'text/css; charset=utf-8';
    default: return 'application/octet-stream';
  }
}

function json(res: ServerResponse, status: number, body: unknown): void {
  res.writeHead(status, { 'content-type': 'application/json' });
  res.end(JSON.stringify(body));
}

