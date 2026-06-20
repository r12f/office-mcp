import { execFileSync, spawn } from 'node:child_process';
import { existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import http from 'node:http';

const DRIVER_DIR = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(DRIVER_DIR, '../../../..');
const DAEMON_EXE = resolve(REPO_ROOT, 'target/debug/office-mcp-daemon.exe');
const DEFAULT_WINDOWS_ACTIVATOR = resolve(REPO_ROOT, 'src/office-ctl/common/scripts/activate-office-mcp-addin.ps1');
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
    case 'listTools':
      return listTools(context);
    case 'createDocument':
      return createDocument(host, context);
    case 'activateAddin':
      return activateAddin(host, context);
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
    case 'describeDocumentLifecycle':
      return describeDocumentLifecycle(host, context);
    case 'describeDaemonStatusCommand':
      return describeDaemonStatusCommand();
    default:
      throw new Error(`Unsupported Office E2E driver step: ${step}`);
  }
}

function describeDaemonStatusCommand() {
  const command = daemonStatusCommand();
  return { command: command.command, args: command.args, cwd: command.cwd };
}

function describeDocumentLifecycle(host, context) {
  const normalizedHost = normalizeHost(host);
  const workDir = resolve(context.workDir || process.env.OFFICE_MCP_E2E_WORK_DIR || tmpdir());
  const extension = normalizedHost === 'word' ? 'docx' : normalizedHost === 'excel' ? 'xlsx' : 'pptx';
  const path = resolve(workDir, `office-mcp-e2e-${normalizedHost}-fixture.${extension}`);
  const closePath = `${path}.office-mcp-close`;
  const readyPath = `${path}.office-mcp-ready`;
  const startedPath = `${path}.office-mcp-started`;
  const errorPath = `${path}.office-mcp-error`;
  const stdoutPath = `${path}.office-mcp-stdout.log`;
  const stderrPath = `${path}.office-mcp-stderr.log`;
  const pidPath = `${path}.office-mcp-pid`;
  return {
    host,
    path,
    createdByDriver: true,
    officeWindowMode: officeWindowMode(normalizedHost),
    keeper: { closePath, readyPath, startedPath, errorPath, stdoutPath, stderrPath, pidPath, scriptPath: `${path}.office-mcp-keeper.ps1` },
    script: officeKeeperScript(normalizedHost, path, closePath, readyPath, startedPath, errorPath),
    cleanupScript: officeCleanupScript(normalizedHost, path)
  };
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
  const command = daemonStatusCommand();
  const output = execFileSync(command.command, command.args, {
    cwd: command.cwd,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe']
  });
  return JSON.parse(output);
}

function daemonStatusCommand() {
  if (existsSync(DAEMON_EXE)) {
    return { command: DAEMON_EXE, args: ['daemon', 'status'], cwd: REPO_ROOT };
  }
  return { command: 'cargo', args: ['run', '-q', '-p', 'office-mcp-daemon', '--', 'daemon', 'status'], cwd: REPO_ROOT };
}

async function createDocument(host, context) {
  const normalizedHost = normalizeHost(host);
  const workDir = resolve(context.workDir || process.env.OFFICE_MCP_E2E_WORK_DIR || tmpdir());
  mkdirSync(workDir, { recursive: true });
  const extension = normalizedHost === 'word' ? 'docx' : normalizedHost === 'excel' ? 'xlsx' : 'pptx';
  const path = resolve(workDir, `office-mcp-e2e-${normalizedHost}-${Date.now()}.${extension}`);
  const keeper = startOfficeKeeper(normalizedHost, path);
  await waitForFile(keeper.readyPath, Number(context.keeperTimeoutMs || 30000), keeper);
  return {
    host,
    path,
    createdByDriver: true,
    officeWindowMode: officeWindowMode(normalizedHost),
    keeper
  };
}

async function listTools(context) {
  const daemon = context.daemon || {};
  const result = await mcpToolsList(daemon.endpoint);
  return Array.isArray(result.tools) ? result.tools.map((tool) => tool.name).filter(Boolean) : [];
}

function startOfficeKeeper(host, path) {
  const closePath = `${path}.office-mcp-close`;
  const readyPath = `${path}.office-mcp-ready`;
  const startedPath = `${path}.office-mcp-started`;
  const errorPath = `${path}.office-mcp-error`;
  const scriptPath = `${path}.office-mcp-keeper.ps1`;
  const stdoutPath = `${path}.office-mcp-stdout.log`;
  const stderrPath = `${path}.office-mcp-stderr.log`;
  const pidPath = `${path}.office-mcp-pid`;
  for (const file of [closePath, readyPath, startedPath, errorPath, scriptPath, stdoutPath, stderrPath, pidPath]) {
    rmSync(file, { force: true });
  }
  const script = officeKeeperScript(host, path, closePath, readyPath, startedPath, errorPath);
  writeFileSync(scriptPath, script, 'utf8');
  runOfficePowerShell(scriptPath, stdoutPath, stderrPath);
  return { closePath, readyPath, startedPath, errorPath, stdoutPath, stderrPath, pidPath, scriptPath };
}

function officeKeeperScript(host, path, closePath, readyPath, startedPath, errorPath) {
  const file = psSingle(path);
  const close = psSingle(closePath);
  const ready = psSingle(readyPath);
  const started = psSingle(startedPath);
  const error = psSingle(errorPath);
  const retry = `function Invoke-Retry([scriptblock]$Action) { for ($i=0; $i -lt 30; $i++) { try { return & $Action } catch { if ($_.Exception.Message -notmatch 'RPC_E_CALL_REJECTED|Call was rejected by callee' -or $i -eq 29) { throw }; Start-Sleep -Milliseconds 500 } } }; `;
  const prelude = `${retry}Set-Content -LiteralPath '${started}' -Value 'office-mcp-keeper:start:${host}'; Write-Output 'office-mcp-keeper:start:${host}'; `;
  if (host === 'word') {
    return `$ErrorActionPreference='Stop'; ${prelude}$app=$null; $doc=$null; try { try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { $app=New-Object -ComObject Word.Application }; $app.Visible=$true; $doc=Invoke-Retry { $app.Documents.Add() }; Invoke-Retry { $doc.SaveAs2('${file}') }; New-Item -ItemType File -Path '${ready}' -Force | Out-Null } catch { Set-Content -LiteralPath '${error}' -Value $_.Exception.Message; throw }`;
  }
  if (host === 'excel') {
    return `$ErrorActionPreference='Stop'; ${prelude}$app=$null; $wb=$null; try { try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $app=New-Object -ComObject Excel.Application }; $app.Visible=$true; $app.DisplayAlerts=$false; $wb=Invoke-Retry { $app.Workbooks.Add() }; $ws=$wb.Worksheets.Item(1); $ws.Cells.Item(1,1).Value2='office-mcp e2e baseline'; Invoke-Retry { $wb.SaveAs('${file}') }; New-Item -ItemType File -Path '${ready}' -Force | Out-Null } catch { Set-Content -LiteralPath '${error}' -Value $_.Exception.Message; throw }`;
  }
  return `$ErrorActionPreference='Stop'; ${prelude}$app=$null; $pres=$null; try { try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app=New-Object -ComObject PowerPoint.Application }; $pres=Invoke-Retry { $app.Presentations.Add($true) }; $slide=Invoke-Retry { $pres.Slides.Add(1, 1) }; $slide.Shapes.Title.TextFrame.TextRange.Text='office-mcp e2e baseline'; Invoke-Retry { $pres.SaveAs('${file}') }; New-Item -ItemType File -Path '${ready}' -Force | Out-Null } catch { Set-Content -LiteralPath '${error}' -Value $_.Exception.Message; throw }`;
}

function runOfficePowerShell(scriptPath, stdoutPath, stderrPath) {
  try {
    const stdout = execFileSync('powershell.exe', ['-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      timeout: 45000
    });
    writeFileSync(stdoutPath, stdout || '', 'utf8');
    writeFileSync(stderrPath, '', 'utf8');
  } catch (error) {
    writeFileSync(stdoutPath, error.stdout?.toString() || '', 'utf8');
    writeFileSync(stderrPath, error.stderr?.toString() || error.message || String(error), 'utf8');
    throw error;
  }
}

function officeWindowMode(host) {
  return 'visible';
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

async function activateAddin(host, context) {
  const activation = activationCommand();
  const command = activation.command;
  if (!command) return { activated: false, skipped: 'no-activator-configured' };
  const normalizedHost = normalizeHost(host);
  const document = context.document || {};
  const daemon = context.daemon || {};
  const activatorLogPath = `${document.path || resolve(tmpdir(), `office-mcp-e2e-${normalizedHost}`)}.office-mcp-activator.log`;
  rmSync(activatorLogPath, { force: true });
  const child = spawn(command, [], {
    shell: true,
    windowsHide: true,
    stdio: 'ignore',
    env: {
      ...process.env,
      OFFICE_MCP_E2E_HOST: normalizedHost,
      OFFICE_MCP_E2E_DOCUMENT_PATH: document.path || '',
      OFFICE_MCP_E2E_ADDIN_ORIGIN: daemon.addinOrigin || '',
      OFFICE_MCP_E2E_ADDIN_ENDPOINT: daemon.addinEndpoint || '',
      OFFICE_MCP_E2E_ACTIVATOR_LOG: activatorLogPath
    }
  });
  const exitCode = await waitForChildExit(child, Number(context.timeoutMs || 30000), () => activatorLogDetail(activatorLogPath));
  if (exitCode !== 0) throw new Error(`Office add-in activator exited with code ${exitCode}.${activatorLogDetail(activatorLogPath)}`);
  return { activated: true, activator: command, activator_kind: activation.kind, log_path: activatorLogPath };
}

function activatorLogDetail(path) {
  if (!path || !existsSync(path)) return '';
  const text = readText(path);
  return text ? ` activator log: ${text.slice(-2000)}` : '';
}

function activationCommand() {
  if (process.env.OFFICE_MCP_E2E_ACTIVATOR) {
    return { command: process.env.OFFICE_MCP_E2E_ACTIVATOR, kind: 'custom' };
  }
  if (process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR === '0') {
    return { command: '', kind: 'disabled' };
  }
  if (process.platform !== 'win32' || !existsSync(DEFAULT_WINDOWS_ACTIVATOR)) {
    return { command: '', kind: 'unavailable' };
  }
  return {
    command: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "${DEFAULT_WINDOWS_ACTIVATOR}" -TimeoutSeconds 20`,
    kind: 'default-windows-taskpane'
  };
}

async function resetContent(_host, context) {
  const actions = context.toolCase?.reset?.actions || [];
  const bindings = await runToolActions(context, actions);
  return { reset: actions.length ? 'mcp-actions' : 'external-driver-delegated', actions: actions.length, bindings };
}

async function setupContent(_host, context) {
  const actions = context.toolCase?.setup?.actions || [];
  const bindings = await runToolActions(context, actions);
  return { setup: actions.length ? 'mcp-actions' : 'external-driver-delegated', actions: actions.length, bindings };
}

async function runToolActions(context, actions) {
  const daemon = context.daemon || {};
  const session = context.session || {};
  const bindings = { session_id: session.sessionId, ...(session.bindings || {}) };
  for (const action of actions) {
    if (!action?.tool && !action?.resource) throw new Error('Office E2E setup/reset action must define a tool or resource.');
    const result = action.resource
      ? await mcpResourceRead(daemon.endpoint, resolveBindings(action.resource, bindings))
      : await mcpToolCall(daemon.endpoint, action.tool, { ...resolveBindings(action.arguments || {}, bindings), session_id: session.sessionId });
    if (result.error || result.structuredContent?.error) {
      throw new Error(`Office E2E setup/reset action ${action.tool || action.resource} failed: ${JSON.stringify(result.error || result.structuredContent.error)}`);
    }
    if (action.saveAs) bindings[action.saveAs] = action.resource ? resourceResultData(result) : actionResultData(result);
  }
  return bindings;
}

async function callTool(context) {
  const daemon = context.daemon || {};
  const session = context.session || {};
  const toolCase = context.toolCase || {};
  const call = toolCase.call || {};
  const args = { ...resolveBindings(call.arguments || {}, session.bindings || {}), session_id: session.sessionId };
  const result = await mcpToolCall(daemon.endpoint, call.name, args);
  return result;
}

async function verifyResult(context) {
  const toolCase = context.toolCase || {};
  const result = context.result || {};
  const verifier = toolCase.verify || { kind: 'readback' };
  const error = resultError(result);
  if (error) {
    const code = error.office_mcp_code || error.code;
    if (arrayOf(verifier.allowErrorCodes).includes(code)) {
      return { verified: true, kind: verifier.kind, acceptedErrorCode: code };
    }
    throw new Error(`${toolCase.tool || 'tool'} returned MCP error: ${JSON.stringify(error)}`);
  }
  if (verifier.kind === 'direct-result') {
    assertReadbackExpectations(toolCase.tool || 'tool', result, verifier.expect || {});
    return { verified: true, kind: verifier.kind };
  }
  if (verifier.kind !== 'readback') return { verified: true, kind: verifier.kind };

  const daemon = context.daemon || {};
  const session = context.session || {};
  const bindings = { session_id: session.sessionId, ...(session.bindings || {}), result: actionResultData(result) };
  if (verifier.resource) {
    const resource = resolveBindings(verifier.resource, bindings);
    const readback = await mcpResourceRead(daemon.endpoint, resource);
    assertReadbackExpectations(toolCase.tool || 'tool', resourceResultData(readback), verifier.expect || {});
    return { verified: true, kind: verifier.kind, readbackResource: resource };
  }
  if (!verifier.readbackTool) return { verified: true, kind: verifier.kind, readback: 'not-configured' };

  const readbackArguments = { ...resolveBindings(verifier.readbackArguments || {}, bindings), session_id: session.sessionId };
  const readback = await mcpToolCall(daemon.endpoint, verifier.readbackTool, readbackArguments);
  assertReadbackExpectations(toolCase.tool || 'tool', readback, verifier.expect || {});
  return { verified: true, kind: verifier.kind, readbackTool: verifier.readbackTool };
}

function resultError(result) {
  return result?.error || result?.structuredContent?.error || null;
}

function actionResultData(result) {
  if (result?.structuredContent?.data) return result.structuredContent.data;
  if (result?.structuredContent) return result.structuredContent;
  if (result?.data) return result.data;
  return result || {};
}

function resourceResultData(result) {
  const text = result?.contents?.[0]?.text || result?.result?.contents?.[0]?.text;
  if (typeof text === 'string') {
    try {
      return JSON.parse(text);
    } catch (_error) {
      return { text };
    }
  }
  return actionResultData(result);
}

function resolveBindings(value, bindings) {
  if (typeof value === 'string') return resolveBindingString(value, bindings);
  if (Array.isArray(value)) return value.map((item) => resolveBindings(item, bindings));
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, resolveBindings(item, bindings)]));
  }
  return value;
}

function resolveBindingString(value, bindings) {
  const exact = value.match(/^\$\{([^}]+)\}$/);
  if (exact) return bindingValue(exact[1], bindings);
  return value.replace(/\$\{([^}]+)\}/g, (_match, path) => String(bindingValue(path, bindings)));
}

function bindingValue(path, bindings) {
  const parts = String(path).split('.').filter(Boolean);
  let current = bindings;
  for (const part of parts) {
    if (current && Object.hasOwn(current, part)) current = current[part];
    else throw new Error(`Office E2E binding ${path} was not found.`);
  }
  return current;
}

function assertReadbackExpectations(tool, readback, expect) {
  const text = readbackText(readback);
  for (const marker of arrayOf(expect.contains)) {
    if (!text.includes(marker)) {
      throw new Error(`${tool} readback did not contain expected text ${JSON.stringify(marker)}.`);
    }
  }
  for (const marker of arrayOf(expect.notContains)) {
    if (text.includes(marker)) {
      throw new Error(`${tool} readback still contained forbidden text ${JSON.stringify(marker)}.`);
    }
  }
  let previousIndex = -1;
  for (const marker of arrayOf(expect.orderedContains)) {
    const index = text.indexOf(marker, previousIndex + 1);
    if (index === -1) {
      throw new Error(`${tool} readback did not contain expected ordered text ${JSON.stringify(marker)}.`);
    }
    previousIndex = index;
  }
  for (const assertion of arrayOf(expect.pathEquals)) {
    const actual = readbackPath(readback, assertion.path);
    if (actual !== assertion.value) {
      throw new Error(`${tool} readback path ${assertion.path} expected ${JSON.stringify(assertion.value)} but found ${JSON.stringify(actual)}.`);
    }
  }
  for (const path of arrayOf(expect.pathMissing)) {
    if (readbackPath(readback, path) !== undefined) {
      throw new Error(`${tool} readback path ${path} was expected to be missing.`);
    }
  }
}

function readbackPath(value, path) {
  const root = value?.structuredContent ?? value;
  let current = root;
  for (const part of String(path).split('.').filter(Boolean)) {
    if (current === undefined || current === null) return undefined;
    const key = Array.isArray(current) && /^\d+$/.test(part) ? Number(part) : part;
    current = current[key];
  }
  return current;
}

function readbackText(value) {
  if (typeof value === 'string') return value;
  if (value?.structuredContent) return readbackText(value.structuredContent);
  if (Array.isArray(value?.content)) return value.content.map((item) => readbackText(item)).join('\n');
  if (typeof value?.text === 'string') return value.text;
  if (typeof value?.value === 'string') return value.value;
  return JSON.stringify(value ?? '');
}

function arrayOf(value) {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

async function cleanupDocument(context) {
  const document = context.document || {};
  if (!document.path) return { deleted: false };
  if (document.createdByDriver !== true || !document.keeper?.closePath) {
    return { deleted: false, skipped: 'not-driver-owned' };
  }
  const resolved = resolve(document.path);
  runOfficeCleanup(document.host, resolved, Number(context.timeoutMs || 30000));
  if (existsSync(resolved)) rmSync(resolved, { force: true });
  if (document.keeper.closePath) rmSync(document.keeper.closePath, { force: true });
  if (document.keeper.readyPath) rmSync(document.keeper.readyPath, { force: true });
  if (document.keeper.startedPath) rmSync(document.keeper.startedPath, { force: true });
  if (document.keeper.errorPath) rmSync(document.keeper.errorPath, { force: true });
  if (document.keeper.stdoutPath) rmSync(document.keeper.stdoutPath, { force: true });
  if (document.keeper.stderrPath) rmSync(document.keeper.stderrPath, { force: true });
  if (document.keeper.pidPath) rmSync(document.keeper.pidPath, { force: true });
  if (document.keeper.scriptPath) rmSync(document.keeper.scriptPath, { force: true });
  return { closedByDriver: true, deleted: !existsSync(resolved), path: resolved };
}

function runOfficeCleanup(host, path, timeoutMs) {
  const normalizedHost = normalizeHost(host);
  const script = officeCleanupScript(normalizedHost, path);
  execFileSync('powershell.exe', ['-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: timeoutMs
  });
}

function officeCleanupScript(host, path) {
  const target = psSingle(resolve(path));
  const canonical = `function Canonical($value) { try { return (Get-Item -LiteralPath $value -ErrorAction Stop).FullName.ToLowerInvariant() } catch { return [System.IO.Path]::GetFullPath($value).ToLowerInvariant() } }; $target=Canonical '${target}';`;
  if (host === 'word') {
    return `$ErrorActionPreference='Stop'; ${canonical} $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); foreach ($doc in @($app.Documents)) { if ((Canonical $doc.FullName) -eq $target) { $doc.Close($false); break } }`;
  }
  if (host === 'excel') {
    return `$ErrorActionPreference='Stop'; ${canonical} $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); foreach ($wb in @($app.Workbooks)) { if ((Canonical $wb.FullName) -eq $target) { $wb.Close($false); break } }`;
  }
  return `$ErrorActionPreference='Stop'; ${canonical} $app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); foreach ($pres in @($app.Presentations)) { if ((Canonical $pres.FullName) -eq $target) { $pres.Close(); break } }`;
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

async function mcpResourceRead(endpoint = 'http://127.0.0.1:8800/mcp', uri) {
  const sessionId = await initializeMcp(endpoint);
  const response = await postJson(endpoint, { jsonrpc: '2.0', id: 3, method: 'resources/read', params: { uri } }, sessionId);
  return response.body.result || response.body;
}

async function mcpToolsList(endpoint = 'http://127.0.0.1:8800/mcp') {
  const sessionId = await initializeMcp(endpoint);
  const response = await postJson(endpoint, { jsonrpc: '2.0', id: 4, method: 'tools/list', params: {} }, sessionId);
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

async function waitForFile(path, timeoutMs, keeper) {
  const started = Date.now();
  let missingProcessChecks = 0;
  while (Date.now() - started <= timeoutMs) {
    if (existsSync(path)) return;
    const errorPath = keeper?.errorPath;
    if (errorPath && existsSync(errorPath)) {
      throw new Error(`Office keeper failed before creating ${path}: ${readText(errorPath)}${keeperLogDetail(keeper)}`);
    }
    const pid = keeperPid(keeper);
    if (pid && !processExists(pid)) {
      missingProcessChecks += 1;
    } else {
      missingProcessChecks = 0;
    }
    if (missingProcessChecks >= 10) {
      throw new Error(`Office keeper exited before creating ${path}.${keeperLogDetail(keeper)}`);
    }
    await sleep(100);
  }
  throw new Error(`Timed out waiting ${timeoutMs} ms for ${path}.${keeperLogDetail(keeper)}`);
}

function keeperPid(keeper) {
  if (!keeper) return undefined;
  if (Number.isInteger(keeper.pid) && keeper.pid > 0) return keeper.pid;
  if (keeper.pidPath && existsSync(keeper.pidPath)) {
    const pid = Number(readText(keeper.pidPath));
    if (Number.isInteger(pid) && pid > 0) return pid;
  }
  return undefined;
}

function keeperLogDetail(keeper) {
  if (!keeper) return '';
  const details = [];
  for (const [label, path] of [['stdout', keeper.stdoutPath], ['stderr', keeper.stderrPath]]) {
    if (path && existsSync(path)) {
      const text = readText(path);
      if (text) details.push(`${label}: ${text.slice(-1000)}`);
    }
  }
  return details.length ? ` ${details.join(' ')}` : '';
}

function readText(path) {
  try {
    return execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', `Get-Content -LiteralPath '${psSingle(path)}' -Raw`], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore']
    }).trim();
  } catch {
    return 'unknown keeper error';
  }
}

async function waitForProcessExit(pid, timeoutMs) {
  if (!pid) return;
  const started = Date.now();
  while (Date.now() - started <= timeoutMs) {
    if (!processExists(pid)) return;
    await sleep(100);
  }
  throw new Error(`Timed out waiting ${timeoutMs} ms for Office keeper process ${pid} to exit.`);
}

function waitForChildExit(child, timeoutMs, detail = () => '') {
  return new Promise((resolvePromise, reject) => {
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`Timed out waiting ${timeoutMs} ms for Office add-in activator to exit.${detail()}`));
    }, timeoutMs);
    child.once('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.once('exit', (code) => {
      clearTimeout(timer);
      resolvePromise(code ?? 0);
    });
  });
}

function processExists(pid) {
  try {
    execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', `Get-Process -Id ${Number(pid)} -ErrorAction Stop | Out-Null`], {
      stdio: 'ignore'
    });
    return true;
  } catch {
    return false;
  }
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
