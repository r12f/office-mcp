import { execFileSync, spawn } from 'node:child_process';
import { existsSync, mkdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import http from 'node:http';

const DRIVER_DIR = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(DRIVER_DIR, '../../../..');
const DEFAULT_TIMEOUT_MS = 120000;

const request = await readRequest();

try {
  const result = await dispatch(request);
  process.stdout.write(JSON.stringify(result ?? {}));
} catch (error) {
  process.stderr.write(error?.message || String(error));
  process.exit(1);
}

async function dispatch({ host, step, context = {} }) {
  switch (step) {
    case 'startDaemon':
      return startDaemon(context);
    case 'createDocument':
      return createDocument(host, context);
    case 'waitForSession':
      return waitForSession(host, context);
    case 'resetContent':
      return resetContent(host, context);
    case 'setupContent':
      return setupContent(host, context);
    case 'callTool':
      return callTool(context);
    case 'verifyResult':
      return verifyResult(context);
    case 'cleanupDocument':
      return cleanupDocument(context);
    case 'stopDaemon':
      return stopDaemon(context);
    default:
      throw new Error(`Unsupported real Office E2E driver step: ${step}`);
  }
}

async function startDaemon(context) {
  const status = daemonStatus();
  if (status.running && status.uiUrl) {
    return daemonFromStatus(status, false);
  }
  const child = spawn('cargo', ['run', '-q', '-p', 'office-mcp-daemon', '--', 'daemon', 'run', '--no-tray'], {
    cwd: REPO_ROOT,
    detached: true,
    stdio: 'ignore',
    windowsHide: true
  });
  child.unref();
  const timeoutMs = Number(context.timeoutMs || DEFAULT_TIMEOUT_MS);
  const started = Date.now();
  while (Date.now() - started <= timeoutMs) {
    await sleep(1000);
    const next = daemonStatus();
    if (next.running && next.uiUrl) return daemonFromStatus(next, true);
  }
  throw new Error(`Timed out waiting ${timeoutMs} ms for office-mcp-daemon to start.`);
}

function daemonFromStatus(status, startedByDriver) {
  const origin = new URL(status.uiUrl).origin;
  return {
    startedByDriver,
    pid: status.pid,
    uiUrl: status.uiUrl,
    stateUrl: status.stateUrl,
    logPath: status.logPath,
    endpoint: process.env.OFFICE_MCP_MCP_ENDPOINT || 'http://127.0.0.1:8800/mcp',
    addinOrigin: origin,
    addinEndpoint: `${origin}/addin`
  };
}

function daemonStatus() {
  const output = execFileSync('cargo', ['run', '-q', '-p', 'office-mcp-daemon', '--', 'daemon', 'status'], {
    cwd: REPO_ROOT,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe']
  });
  return JSON.parse(output);
}

async function createDocument(host, context) {
  const normalizedHost = normalizeHost(host);
  const workDir = resolve(context.workDir || process.env.OFFICE_MCP_E2E_WORK_DIR || tmpdir());
  mkdirSync(workDir, { recursive: true });
  const extension = normalizedHost === 'word' ? 'docx' : normalizedHost === 'excel' ? 'xlsx' : 'pptx';
  const path = resolve(workDir, `office-mcp-e2e-${normalizedHost}-${Date.now()}.${extension}`);
  runPowerShell(documentCreateScript(normalizedHost, path));
  return { host, path };
}

function documentCreateScript(host, path) {
  const literal = psSingle(path);
  if (host === 'word') {
    return `$ErrorActionPreference='Stop'; $app=New-Object -ComObject Word.Application; $app.Visible=$false; $doc=$app.Documents.Add(); $doc.Content.Text='office-mcp e2e baseline'; $doc.SaveAs2('${literal}'); $doc.Close($false); $app.Quit(); Write-Output '${literal}'`;
  }
  if (host === 'excel') {
    return `$ErrorActionPreference='Stop'; $app=New-Object -ComObject Excel.Application; $app.Visible=$false; $wb=$app.Workbooks.Add(); $ws=$wb.Worksheets.Item(1); $ws.Cells.Item(1,1).Value2='office-mcp e2e baseline'; $wb.SaveAs('${literal}'); $wb.Close($false); $app.Quit(); Write-Output '${literal}'`;
  }
  return `$ErrorActionPreference='Stop'; $app=New-Object -ComObject PowerPoint.Application; $pres=$app.Presentations.Add($true); $slide=$pres.Slides.Add(1, 1); $slide.Shapes.Title.TextFrame.TextRange.Text='office-mcp e2e baseline'; $pres.SaveAs('${literal}'); $pres.Close(); $app.Quit(); Write-Output '${literal}'`;
}

async function waitForSession(host, context) {
  const daemon = context.daemon || {};
  const document = context.document || {};
  const timeoutMs = Number(context.timeoutMs || process.env.OFFICE_MCP_E2E_SESSION_TIMEOUT_MS || DEFAULT_TIMEOUT_MS);
  const started = Date.now();
  let latest = [];
  while (Date.now() - started <= timeoutMs) {
    latest = await listSessions(daemon.endpoint);
    const match = latest.find((session) => sessionMatches(session, host, document.path));
    if (match) {
      return {
        sessionId: match.session_id,
        availableTools: Array.isArray(match.available_tools) ? match.available_tools : [],
        descriptor: match
      };
    }
    await sleep(1000);
  }
  throw new Error(`Timed out waiting ${timeoutMs} ms for ${host} add-in session for ${document.path || 'test document'}. Open MCP Control in ${host}, ensure it connects to the daemon, then rerun npm run e2e:tools with OFFICE_MCP_RUN_E2E=1.`);
}

async function resetContent(_host, _context) {
  return { reset: 'external-driver-delegated' };
}

async function setupContent(_host, _context) {
  return { setup: 'external-driver-delegated' };
}

async function callTool(context) {
  const daemon = context.daemon || {};
  const session = context.session || {};
  const toolCase = context.toolCase || {};
  const call = toolCase.call || {};
  const args = { ...(call.arguments || {}), session_id: session.sessionId };
  const result = await mcpToolCall(daemon.endpoint, call.name, args);
  return result;
}

async function verifyResult(context) {
  const toolCase = context.toolCase || {};
  const result = context.result || {};
  if (result.error || result.structuredContent?.error) {
    throw new Error(`${toolCase.tool || 'tool'} returned MCP error: ${JSON.stringify(result.error || result.structuredContent.error)}`);
  }
  return { verified: true, kind: toolCase.verify?.kind || 'readback' };
}

async function cleanupDocument(context) {
  const document = context.document || {};
  if (!document.path) return { deleted: false };
  const resolved = resolve(document.path);
  if (existsSync(resolved)) rmSync(resolved, { force: true });
  return { deleted: !existsSync(resolved), path: resolved };
}

async function stopDaemon(context) {
  const daemon = context.daemon || {};
  if (!daemon.startedByDriver) return { stopped: false };
  execFileSync('cargo', ['run', '-q', '-p', 'office-mcp-daemon', '--', 'daemon', 'stop'], {
    cwd: REPO_ROOT,
    stdio: ['ignore', 'pipe', 'pipe']
  });
  return { stopped: true };
}

async function listSessions(endpoint) {
  const result = await mcpToolCall(endpoint, 'office.list_sessions', {});
  return Array.isArray(result.structuredContent?.sessions) ? result.structuredContent.sessions : [];
}

async function mcpToolCall(endpoint = 'http://127.0.0.1:8800/mcp', name, args) {
  const sessionId = await initializeMcp(endpoint);
  const response = await postJson(endpoint, { jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name, arguments: args } }, sessionId);
  return response.body.result || response.body;
}

async function initializeMcp(endpoint) {
  const response = await postJson(endpoint, { jsonrpc: '2.0', id: 1, method: 'initialize', params: {} });
  const sessionId = response.headers['mcp-session-id'];
  if (!sessionId) throw new Error('MCP initialize response did not include MCP-Session-Id.');
  return sessionId;
}

function postJson(endpoint, body, sessionId) {
  return new Promise((resolvePromise, reject) => {
    const url = new URL(endpoint);
    const text = JSON.stringify(body);
    const request = http.request({
      method: 'POST',
      hostname: url.hostname,
      port: url.port,
      path: url.pathname,
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json',
        'Content-Length': Buffer.byteLength(text),
        ...(sessionId ? { 'MCP-Session-Id': sessionId } : {})
      }
    }, (response) => {
      let responseText = '';
      response.setEncoding('utf8');
      response.on('data', (chunk) => { responseText += chunk; });
      response.on('end', () => {
        try {
          resolvePromise({ headers: response.headers, body: JSON.parse(responseText || '{}') });
        } catch (error) {
          reject(error);
        }
      });
    });
    request.on('error', reject);
    request.end(text);
  });
}

function sessionMatches(session, host, documentPath) {
  const app = String(session.app || session.host?.app || '').toLowerCase();
  if (app !== normalizeHost(host)) return false;
  if (!documentPath) return true;
  const filename = String(session.document?.filename || session.document?.title || '').toLowerCase();
  return filename && documentPath.toLowerCase().includes(filename);
}

function normalizeHost(host) {
  return String(host || '').toLowerCase();
}

function runPowerShell(command) {
  execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe']
  });
}

function psSingle(value) {
  return String(value).replace(/'/g, "''");
}

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

function readRequest() {
  return new Promise((resolvePromise, reject) => {
    let input = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', (chunk) => { input += chunk; });
    process.stdin.on('end', () => {
      try {
        resolvePromise(JSON.parse(input || '{}'));
      } catch (error) {
        reject(error);
      }
    });
  });
}
