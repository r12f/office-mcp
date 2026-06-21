import { execFileSync, spawn, spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, rmSync, unlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import http from 'node:http';

const DRIVER_DIR = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(DRIVER_DIR, '../../../..');
const DAEMON_EXE = resolve(REPO_ROOT, 'target/debug/office-mcp-daemon.exe');
const DEFAULT_WINDOWS_ACTIVATOR = resolve(REPO_ROOT, 'src/office-ctl/common/scripts/activate-office-mcp-addin.ps1');
const DEFAULT_DAEMON_START_TIMEOUT_MS = 120000;
const DEFAULT_OFFICE_START_TIMEOUT_MS = 90000;
const DEFAULT_ACTIVATOR_TIMEOUT_MS = 120000;
const DEFAULT_ACTIVATION_SESSION_TIMEOUT_MS = 10000;
const DEFAULT_SESSION_TIMEOUT_MS = 10000;
const mcpSessionIds = new Map();

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
  const activationLogPath = activationLogPathForDocument(path, normalizedHost);
  return {
    host,
    path,
    createdByDriver: true,
    officeWindowMode: officeWindowMode(normalizedHost),
    officeProcessIdsBefore: [],
    activationLogPath,
    keeper: { closePath, readyPath, startedPath, errorPath, stdoutPath, stderrPath, pidPath, scriptPath: `${path}.office-mcp-keeper.ps1` },
    script: officeKeeperScript(normalizedHost, path, closePath, readyPath, startedPath, errorPath),
    cleanupScript: officeCleanupScript(normalizedHost, [path], [])
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
  const timeoutMs = Number(context.timeoutMs || process.env.OFFICE_MCP_E2E_DAEMON_START_TIMEOUT_MS || DEFAULT_DAEMON_START_TIMEOUT_MS);
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
  const officeProcessIdsBefore = officeProcessIds(normalizedHost);
  const workDir = resolve(context.workDir || process.env.OFFICE_MCP_E2E_WORK_DIR || tmpdir());
  mkdirSync(workDir, { recursive: true });
  const extension = normalizedHost === 'word' ? 'docx' : normalizedHost === 'excel' ? 'xlsx' : 'pptx';
  const path = resolve(workDir, `office-mcp-e2e-${normalizedHost}-${Date.now()}.${extension}`);
  const keeperTimeoutMs = Number(context.keeperTimeoutMs || process.env.OFFICE_MCP_E2E_OFFICE_START_TIMEOUT_MS || DEFAULT_OFFICE_START_TIMEOUT_MS);
  const powerShellTimeoutMs = Number(context.powerShellTimeoutMs || keeperTimeoutMs + 10000);
  let keeper;
  try {
    keeper = startOfficeKeeper(normalizedHost, path, powerShellTimeoutMs);
    await waitForFile(keeper.readyPath, keeperTimeoutMs, keeper);
  } catch (error) {
    cleanupKeeperArtifacts(keeper, path);
    throw error;
  }
  return {
    host,
    path,
    createdByDriver: true,
    officeWindowMode: officeWindowMode(normalizedHost),
    officeProcessIdsBefore,
    activationLogPath: activationLogPathForDocument(path, normalizedHost),
    keeper
  };
}

function activationLogPathForDocument(path, host) {
  return `${path || resolve(tmpdir(), `office-mcp-e2e-${normalizeHost(host)}`)}.office-mcp-activator.log`;
}

async function listTools(context) {
  const daemon = context.daemon || {};
  const result = await mcpToolsList(daemon.endpoint);
  return Array.isArray(result.tools) ? result.tools.map((tool) => tool.name).filter(Boolean) : [];
}

function startOfficeKeeper(host, path, powerShellTimeoutMs) {
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
  const keeper = { closePath, readyPath, startedPath, errorPath, stdoutPath, stderrPath, pidPath, scriptPath };
  try {
    runOfficePowerShell(scriptPath, stdoutPath, stderrPath, powerShellTimeoutMs);
    return keeper;
  } catch (error) {
    cleanupKeeperArtifacts(keeper, path);
    throw error;
  }
}

function cleanupKeeperArtifacts(keeper, path) {
  for (const file of [
    path,
    keeper?.closePath,
    keeper?.readyPath,
    keeper?.startedPath,
    keeper?.errorPath,
    keeper?.stdoutPath,
    keeper?.stderrPath,
    keeper?.pidPath,
    keeper?.scriptPath
  ]) {
    if (file) rmSync(file, { force: true });
  }
}

function officeKeeperScript(host, path, closePath, readyPath, startedPath, errorPath) {
  const file = psSingle(path);
  const close = psSingle(closePath);
  const ready = psSingle(readyPath);
  const started = psSingle(startedPath);
  const error = psSingle(errorPath);
  const pid = psSingle(`${path}.office-mcp-pid`);
  const retry = `function Invoke-Retry([scriptblock]$Action) { for ($i=0; $i -lt 90; $i++) { try { return & $Action } catch { if ($i -eq 89) { throw }; Start-Sleep -Milliseconds 500 } } }; <# retries RPC_E_CALL_REJECTED and transient Office COM busy states. #> `;
  const processIdHelper = `function Write-OfficeMcpProcessPid([IntPtr]$Handle) { try { $element=[System.Windows.Automation.AutomationElement]::FromHandle($Handle); if ($element) { Set-Content -LiteralPath '${pid}' -Value $element.Current.ProcessId } } catch {} }; `;
  const prelude = `${retry}${processIdHelper}Set-Content -LiteralPath '${started}' -Value 'office-mcp-keeper:start:${host}'; Write-Output 'office-mcp-keeper:start:${host}'; `;
  if (host === 'word') {
    return wordKeeperScript(file, ready, error, prelude);
  }
  if (host === 'excel') {
    return `$ErrorActionPreference='Stop'; Add-Type -AssemblyName UIAutomationClient; ${prelude}$app=$null; $wb=$null; try { $app=Invoke-Retry { New-Object -ComObject Excel.Application }; Invoke-Retry { $app.Visible=$true }; Write-OfficeMcpProcessPid -Handle $app.Hwnd; Invoke-Retry { $app.DisplayAlerts=$false }; $wb=Invoke-Retry { $app.Workbooks.Add() }; $ws=Invoke-Retry { $wb.Worksheets.Item(1) }; Invoke-Retry { $ws.Cells.Item(1,1).Value2='office-mcp e2e baseline' }; Invoke-Retry { $wb.SaveAs('${file}') }; New-Item -ItemType File -Path '${ready}' -Force | Out-Null } catch { Set-Content -LiteralPath '${error}' -Value $_.Exception.Message; throw }`;
  }
  return `$ErrorActionPreference='Stop'; Add-Type -AssemblyName UIAutomationClient; ${prelude}$app=$null; $pres=$null; $createdApp=$false; try { try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $app=New-Object -ComObject PowerPoint.Application; $createdApp=$true }; if ($createdApp -and $app.HWND) { Write-OfficeMcpProcessPid -Handle $app.HWND }; $pres=Invoke-Retry { $app.Presentations.Add($true) }; $slide=Invoke-Retry { $pres.Slides.Add(1, 1) }; $slide.Shapes.Title.TextFrame.TextRange.Text='office-mcp e2e baseline'; Invoke-Retry { $pres.SaveAs('${file}') }; New-Item -ItemType File -Path '${ready}' -Force | Out-Null } catch { Set-Content -LiteralPath '${error}' -Value $_.Exception.Message; throw }`;
}

function wordKeeperScript(file, ready, error, prelude) {
  const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/></Types>';
  const rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>';
  const documentRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/></Relationships>';
  const core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Office MCP E2E</dc:title><dc:creator>office-mcp</dc:creator><cp:lastModifiedBy>office-mcp</cp:lastModifiedBy></cp:coreProperties>';
  const app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office Word</Application></Properties>';
  const document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>office-mcp e2e baseline</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>';
  const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>';
  const settings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
  const fontTable = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Calibri"/></w:fonts>';
  const webSettings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

  return `$ErrorActionPreference='Stop'; function Add-ZipPart([System.IO.Compression.ZipArchive]$Zip,[string]$Name,[string]$Value) { $entry=$Zip.CreateEntry($Name); $writer=New-Object System.IO.StreamWriter($entry.Open()); try { $writer.Write($Value) } finally { $writer.Dispose() } }; function New-OfficeMcpBlankDocx([string]$Path) { Add-Type -AssemblyName System.IO.Compression; Add-Type -AssemblyName System.IO.Compression.FileSystem; if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }; $stream=[System.IO.File]::Open($Path,[System.IO.FileMode]::CreateNew); try { $zip=New-Object System.IO.Compression.ZipArchive($stream,[System.IO.Compression.ZipArchiveMode]::Create); try { Add-ZipPart $zip '[Content_Types].xml' '${psSingle(contentTypes)}'; Add-ZipPart $zip '_rels/.rels' '${psSingle(rootRels)}'; Add-ZipPart $zip 'docProps/core.xml' '${psSingle(core)}'; Add-ZipPart $zip 'docProps/app.xml' '${psSingle(app)}'; Add-ZipPart $zip 'word/_rels/document.xml.rels' '${psSingle(documentRels)}'; Add-ZipPart $zip 'word/document.xml' '${psSingle(document)}'; Add-ZipPart $zip 'word/styles.xml' '${psSingle(styles)}'; Add-ZipPart $zip 'word/settings.xml' '${psSingle(settings)}'; Add-ZipPart $zip 'word/fontTable.xml' '${psSingle(fontTable)}'; Add-ZipPart $zip 'word/webSettings.xml' '${psSingle(webSettings)}' } finally { $zip.Dispose() } } finally { $stream.Dispose() } }; ${prelude}$app=$null; $doc=$null; try { New-OfficeMcpBlankDocx '${file}'; try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { $app=New-Object -ComObject Word.Application }; $app.Visible=$true; $app.DisplayAlerts=0; $confirmConversions=$false; $readOnly=$false; $addToRecentFiles=$false; $doc=Invoke-Retry { $app.Documents.Open('${file}', $confirmConversions, $readOnly, $addToRecentFiles) }; New-Item -ItemType File -Path '${ready}' -Force | Out-Null } catch { Set-Content -LiteralPath '${error}' -Value $_.Exception.Message; throw }`;
}

function runOfficePowerShell(scriptPath, stdoutPath, stderrPath, timeoutMs = 45000) {
  const shell = powerShellCommand(scriptPath);
  const result = spawnSync(shell.command, shell.args, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: timeoutMs,
    windowsHide: true
  });
  writeFileSync(stdoutPath, result.stdout || '', 'utf8');
  writeFileSync(stderrPath, result.stderr || result.error?.message || '', 'utf8');
  if (result.error) throw new Error(`${result.error.message}.${keeperLogDetail({ stdoutPath, stderrPath })}`);
  if (result.status !== 0) {
    throw new Error(`Office keeper PowerShell exited with status ${result.status}.${keeperLogDetail({ stdoutPath, stderrPath })}`);
  }
}

function powerShellCommand(scriptPath) {
  const override = process.env.OFFICE_MCP_E2E_POWERSHELL;
  if (override) {
    const parts = splitCommand(override);
    return { command: parts[0], args: [...parts.slice(1), scriptPath] };
  }
  return { command: 'powershell.exe', args: ['-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath] };
}

function splitCommand(command) {
  const parts = String(command).match(/(?:[^\s"]+|"[^"]*")+/g)?.map((part) => part.replace(/^"|"$/g, '')) || [];
  if (parts.length === 0) throw new Error('OFFICE_MCP_E2E_POWERSHELL must not be empty.');
  return parts;
}

function officeWindowMode(host) {
  return 'visible';
}

async function waitForSession(host, context) {
  const daemon = context.daemon || {};
  const document = context.document || {};
  const timeoutMs = Number(context.timeoutMs || process.env.OFFICE_MCP_E2E_SESSION_TIMEOUT_MS || DEFAULT_SESSION_TIMEOUT_MS);
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
  throw new Error(`Timed out waiting ${timeoutMs} ms for ${host} add-in session for ${document.path || 'test document'}. ${sessionWaitDiagnostic({ host, daemon, document, latest })} Open MCP Control in ${host}, ensure it connects to the daemon, then rerun npm run e2e:tools with OFFICE_MCP_RUN_E2E=1.`);
}

async function activateAddin(host, context) {
  const activation = activationCommand();
  const command = activation.command;
  if (!command) return { activated: false, skipped: 'no-activator-configured' };
  const normalizedHost = normalizeHost(host);
  const document = context.document || {};
  const daemon = context.daemon || {};
  const activatorLogPath = document.activationLogPath || activationLogPathForDocument(document.path, normalizedHost);
  rmSync(activatorLogPath, { force: true });
  const child = spawn(command, [], {
    shell: true,
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
    env: {
      ...process.env,
      OFFICE_MCP_E2E_HOST: normalizedHost,
      OFFICE_MCP_E2E_DOCUMENT_PATH: document.path || '',
      OFFICE_MCP_E2E_ADDIN_ORIGIN: daemon.addinOrigin || '',
      OFFICE_MCP_E2E_ADDIN_ENDPOINT: daemon.addinEndpoint || '',
      OFFICE_MCP_E2E_MANIFEST_PATH: manifestPathForHost(normalizedHost),
      OFFICE_MCP_E2E_ACTIVATOR_LOG: activatorLogPath
    }
  });
  let result;
  try {
    result = await waitForChildExit(child, Number(context.timeoutMs || process.env.OFFICE_MCP_E2E_ACTIVATOR_TIMEOUT_MS || DEFAULT_ACTIVATOR_TIMEOUT_MS), () => activatorLogDetail(activatorLogPath));
  } catch (error) {
    const detected = await registeredSessionAfterActivatorFailure(host, document, daemon, activatorLogPath);
    if (detected) return activationFromDetectedSession(command, activation.kind, activatorLogPath, detected.documentPath, normalizedHost, 'session-detected-after-activator-timeout', detected.document);
    throw error;
  }
  if (result.exitCode !== 0) {
    const detected = await registeredSessionAfterActivatorFailure(host, document, daemon, activatorLogPath);
    if (detected) return activationFromDetectedSession(command, activation.kind, activatorLogPath, detected.documentPath, normalizedHost, 'session-detected-after-activator-failure', detected.document);
    const preservedLogPath = preserveActivatorLogArtifact(activatorLogPath, normalizedHost);
    throw new Error(`Office add-in activator exited with code ${result.exitCode}.${activatorFailureDetail(preservedLogPath || activatorLogPath)}`);
  }
  const parsed = parseActivatorResult(result.stdout);
  if (parsed.document_path && document.path && parsed.document_path !== document.path) {
    parsed.document = { ...document, original_path: document.original_path || document.path, path: parsed.document_path };
  }
  return {
    activated: true,
    ...parsed,
    activator: command,
    activator_kind: activation.kind,
    log_path: activatorLogPath
  };
}

async function registeredSessionAfterActivatorFailure(host, document, daemon, activatorLogPath) {
  if (!daemon.endpoint || !document.path) return null;
  const documentPaths = documentPathsAfterActivationFailure(document, activatorLogPath);
  const timeoutMs = Number(process.env.OFFICE_MCP_E2E_ACTIVATION_SESSION_TIMEOUT_MS || DEFAULT_ACTIVATION_SESSION_TIMEOUT_MS);
  const started = Date.now();
  while (Date.now() - started <= timeoutMs) {
    try {
      const sessions = await listSessions(daemon.endpoint);
      for (const documentPath of documentPaths) {
        const match = sessions.find((session) => sessionMatches(session, host, documentPath));
        if (match) return { session: match, documentPath, document: activatedDocumentContext(document, documentPath) };
      }
    } catch {}
    await sleep(500);
  }
  return null;
}

function documentPathsAfterActivationFailure(document, activatorLogPath) {
  const paths = [document.path].filter(Boolean);
  const log = activatorLogPath && existsSync(activatorLogPath) ? readText(activatorLogPath) : '';
  for (const match of log.matchAll(/^official sideload document=(.+)$/gm)) {
    const path = match[1]?.trim();
    if (path && !paths.includes(path)) paths.push(path);
  }
  return paths;
}

function activatedDocumentContext(document, documentPath) {
  if (!documentPath || documentPath === document.path) return document;
  return { ...document, original_path: document.original_path || document.path, path: documentPath };
}

function activationFromDetectedSession(command, kind, logPath, documentPath, host, activationPath, document) {
  return {
    activated: true,
    host: normalizeHost(host),
    document_path: documentPath,
    document,
    control_opened: false,
    activation_path: activationPath,
    activator: command,
    activator_kind: kind,
    log_path: logPath
  };
}

function parseActivatorResult(stdout) {
  const text = String(stdout || '').trim();
  if (!text) return {};
  const line = text.split(/\r?\n/).reverse().find((entry) => entry.trim().startsWith('{'));
  if (!line) return {};
  try {
    const parsed = JSON.parse(line);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function manifestPathForHost(host) {
  return resolve(REPO_ROOT, `addin-catalog/office-mcp-${host}.xml`);
}

function activatorLogDetail(path) {
  if (!path || !existsSync(path)) return '';
  const text = readText(path);
  return text ? ` activator log: ${text.slice(-8000)}` : '';
}

function sessionWaitDiagnostic({ host, daemon = {}, document = {}, latest = [] }) {
  const fields = [
    'phase=addin-session-register',
    `host=${normalizeHost(host)}`,
    `document=${document.path || 'test document'}`,
    `mcp_endpoint=${daemon.endpoint || ''}`,
    `addin_origin=${daemon.addinOrigin || ''}`,
    `addin_endpoint=${daemon.addinEndpoint || ''}`,
    `latest_sessions=${safeJson(latest)}`
  ];
  if (document.activationLogPath) {
    fields.push(`activation_log_path=${document.activationLogPath}`);
    fields.push(`activation_log_tail=${fileTail(document.activationLogPath)}`);
  }
  if (daemon.logPath) {
    fields.push(`daemon_log_path=${daemon.logPath}`);
    fields.push(`daemon_log_tail=${fileTail(daemon.logPath)}`);
  }
  return fields.join(' ');
}

function fileTail(path, maxLength = 4000) {
  if (!path || !existsSync(path)) return '';
  return compactDiagnosticText(readText(path).slice(-maxLength));
}

function compactDiagnosticText(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function safeJson(value) {
  try {
    return compactDiagnosticText(JSON.stringify(value));
  } catch {
    return 'unserializable';
  }
}

function activatorFailureDetail(path) {
  const tail = activatorLogDetail(path);
  return path ? `${tail} full activator log: ${path}` : tail;
}

function preserveActivatorLogArtifact(path, host) {
  if (!path || !existsSync(path)) return '';
  const artifactsDir = resolve(REPO_ROOT, 'artifacts');
  mkdirSync(artifactsDir, { recursive: true });
  const artifactPath = resolve(artifactsDir, `office-mcp-activator-${normalizeHost(host)}-${Date.now()}.log`);
  copyFileSync(path, artifactPath);
  return artifactPath;
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
    command: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "${DEFAULT_WINDOWS_ACTIVATOR}" -TimeoutSeconds 90`,
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
  const document = context.document || {};
  const bindings = { session_id: session.sessionId, ...(session.bindings || {}) };
  let acceptedErrorCode;
  for (const action of actions) {
    if (!action?.tool && !action?.resource && !action?.driver) throw new Error('Office E2E setup/reset action must define a tool, resource, or driver action.');
    const result = await runSetupAction({ action, bindings, daemon, document, session });
    if (result.error || result.structuredContent?.error) {
      const error = result.error || result.structuredContent.error;
      const code = error.office_mcp_code || error.code;
      if (arrayOf(action.allowErrorCodes).includes(code)) {
        if (action.saveAs) bindings[action.saveAs] = { skipped: true, accepted_error_code: code };
        acceptedErrorCode ??= code;
        continue;
      }
      throw new Error(`Office E2E setup/reset action ${action.tool || action.resource || action.driver} failed: ${JSON.stringify(error)}`);
    }
    if (action.saveAs) bindings[action.saveAs] = action.resource ? resourceResultData(result) : actionResultData(result);
  }
  if (acceptedErrorCode) bindings.__accepted_error_code = acceptedErrorCode;
  return bindings;
}

async function runSetupAction({ action, bindings, daemon, document, session }) {
  if (action.resource) return mcpResourceRead(daemon.endpoint, resolveBindings(action.resource, bindings));
  if (action.tool) return mcpToolCall(daemon.endpoint, action.tool, { ...resolveBindings(action.arguments || {}, bindings), session_id: session.sessionId });
  return runDriverSetupAction(action.driver, { action, bindings, document });
}

function runDriverSetupAction(name, { action, bindings, document }) {
  if (name === 'word.create_tracked_change') {
    const args = resolveBindings(action.arguments || {}, bindings);
    return createWordTrackedChange(document.path, args);
  }
  throw new Error(`Unsupported Office E2E driver setup action: ${name}`);
}

function createWordTrackedChange(documentPath, args = {}) {
  if (process.platform !== 'win32') throw new Error('word.create_tracked_change requires Windows Word COM automation.');
  if (!documentPath) throw new Error('word.create_tracked_change requires a driver-owned document path.');
  const marker = String(args.text || `Tracked change E2E paragraph ${Date.now()}`);
  const script = [
    '$target = ' + psQuoted(resolve(documentPath)),
    '$marker = ' + psQuoted(marker),
    '$word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")',
    '$doc = $null',
    'foreach ($candidate in @($word.Documents)) { if ([System.IO.Path]::GetFullPath($candidate.FullName).ToLowerInvariant() -eq [System.IO.Path]::GetFullPath($target).ToLowerInvariant()) { $doc = $candidate; break } }',
    'if ($null -eq $doc) { throw "Target Word document not open: $target" }',
    '$doc.Activate()',
    '$doc.TrackRevisions = $true',
    '$range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)',
    '$range.InsertAfter([Environment]::NewLine + $marker)',
    '$doc.TrackRevisions = $false',
    '[pscustomobject]@{ document = $doc.FullName; marker = $marker; revisions = $doc.Revisions.Count } | ConvertTo-Json -Depth 4'
  ].join('; ');
  const shell = powerShellInlineCommand(script);
  const output = execFileSync(shell.command, shell.args, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: 30000
  });
  return JSON.parse(output);
}

function powerShellInlineCommand(script) {
  const override = process.env.OFFICE_MCP_E2E_POWERSHELL_INLINE;
  if (override) {
    const parts = splitCommand(override);
    return { command: parts[0], args: [...parts.slice(1), '-Command', script] };
  }
  return { command: 'powershell.exe', args: ['-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script] };
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
      throw new Error(`${tool} readback did not contain expected text ${JSON.stringify(marker)}; actual readback: ${truncateText(text)}.`);
    }
  }
  for (const marker of arrayOf(expect.notContains)) {
    if (text.includes(marker)) {
      throw new Error(`${tool} readback still contained forbidden text ${JSON.stringify(marker)}; actual readback: ${truncateText(text)}.`);
    }
  }
  let previousIndex = -1;
  for (const marker of arrayOf(expect.orderedContains)) {
    const index = text.indexOf(marker, previousIndex + 1);
    if (index === -1) {
      throw new Error(`${tool} readback did not contain expected ordered text ${JSON.stringify(marker)}; actual readback: ${truncateText(text)}.`);
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

function truncateText(text, maxLength = 1200) {
  const value = String(text ?? '');
  return value.length > maxLength ? `${value.slice(0, maxLength)}...` : value;
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
  const initialPaths = driverOwnedCleanupPaths(document);
  const processIds = driverOwnedProcessIds(document);
  runOfficeCleanup(document.host, initialPaths, processIds, Number(context.timeoutMs || 30000));
  const cleanupPaths = [...new Set([...initialPaths, ...driverOwnedCleanupPaths(document)])];
  for (const path of cleanupPaths) await removeFileWithRetry(path, Number(context.deleteTimeoutMs || 10000));
  if (document.keeper.closePath) rmSync(document.keeper.closePath, { force: true });
  if (document.keeper.readyPath) rmSync(document.keeper.readyPath, { force: true });
  if (document.keeper.startedPath) rmSync(document.keeper.startedPath, { force: true });
  if (document.keeper.errorPath) rmSync(document.keeper.errorPath, { force: true });
  if (document.keeper.stdoutPath) rmSync(document.keeper.stdoutPath, { force: true });
  if (document.keeper.stderrPath) rmSync(document.keeper.stderrPath, { force: true });
  if (document.keeper.pidPath) rmSync(document.keeper.pidPath, { force: true });
  if (document.keeper.scriptPath) rmSync(document.keeper.scriptPath, { force: true });
  if (document.activationLogPath) rmSync(document.activationLogPath, { force: true });
  return { closedByDriver: true, deleted: cleanupPaths.every((path) => !existsSync(path)), path: resolve(document.path), deletedPaths: cleanupPaths };
}

function driverOwnedCleanupPaths(document) {
  const paths = new Set();
  for (const path of [document.path, document.original_path]) {
    if (path) paths.add(resolve(path));
  }
  for (const path of officeSideloadCopyCandidates(document.host)) paths.add(resolve(path));
  return [...paths];
}

function driverOwnedProcessIds(document) {
  const ids = new Set();
  const pid = Number(document.keeper?.officeProcessId || (document.keeper?.pidPath && existsSync(document.keeper.pidPath) ? readText(document.keeper.pidPath) : 0));
  if (Number.isInteger(pid) && pid > 0) ids.add(pid);
  for (const id of driverOwnedEmptyOfficeProcessIds(document)) ids.add(id);
  return [...ids];
}

function driverOwnedEmptyOfficeProcessIds(document) {
  const host = normalizeHost(document.host || '');
  if (host !== 'excel') return [];
  const before = new Set(Array.isArray(document.officeProcessIdsBefore) ? document.officeProcessIdsBefore.map(Number) : []);
  return officeProcesses(host)
    .filter(({ id, title }) => !before.has(id) && (title === '' || title === officeAppName(host)))
    .map(({ id }) => id);
}

function officeProcessIds(host) {
  return officeProcesses(host).map(({ id }) => id);
}

function officeProcesses(host) {
  const processName = host === 'word' ? 'WINWORD' : host === 'powerpoint' ? 'POWERPNT' : 'EXCEL';
  const script = `Get-Process -Name '${processName}' -ErrorAction SilentlyContinue | Select-Object Id,MainWindowTitle | ConvertTo-Json -Compress`;
  let output = '';
  try {
    output = execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore']
    }).trim();
  } catch {
    return [];
  }
  if (!output) return [];
  const parsed = JSON.parse(output);
  return (Array.isArray(parsed) ? parsed : [parsed]).map((process) => ({ id: Number(process.Id), title: String(process.MainWindowTitle || '') })).filter(({ id }) => Number.isInteger(id));
}

function officeSideloadCopyCandidates(host) {
  const normalizedHost = normalizeHost(host);
  const specs = {
    word: { extension: 'docx', id: '11111111-aaaa-bbbb-cccc-222222222222' },
    excel: { extension: 'xlsx', id: '33333333-aaaa-bbbb-cccc-444444444444' },
    powerpoint: { extension: 'pptx', id: '44444444-aaaa-bbbb-cccc-555555555555' }
  };
  const spec = specs[normalizedHost];
  if (!spec) return [];
  return [
    ...globTempFiles(`Office add-in ${spec.id}*.${spec.extension}`),
    ...globTempFiles(`${officeAppName(normalizedHost)} add-in ${spec.id}*.${spec.extension}`)
  ];
}

function officeAppName(host) {
  if (host === 'word') return 'Word';
  if (host === 'excel') return 'Excel';
  return 'PowerPoint';
}

function globTempFiles(pattern) {
  const script = `Get-ChildItem -LiteralPath '${psSingle(tmpdir())}' -Filter '${psSingle(pattern)}' -File -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }`;
  const output = execFileSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'ignore']
  }).trim();
  return output ? output.split(/\r?\n/).filter(Boolean) : [];
}

async function removeFileWithRetry(path, timeoutMs) {
  const started = Date.now();
  let lastError;
  while (Date.now() - started <= timeoutMs) {
    if (!existsSync(path)) return;
    try {
      unlinkSync(path);
      return;
    } catch (error) {
      lastError = error;
      if (!['EPERM', 'EBUSY', 'EACCES'].includes(error?.code)) throw error;
      await sleep(200);
    }
  }
  throw lastError || new Error(`Timed out deleting ${path}.`);
}

function runOfficeCleanup(host, paths, processIds, timeoutMs) {
  const normalizedHost = normalizeHost(host);
  const script = officeCleanupScript(normalizedHost, Array.isArray(paths) ? paths : [paths], Array.isArray(processIds) ? processIds : []);
  try {
    execFileSync('powershell.exe', ['-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      timeout: timeoutMs
    });
  } catch (error) {
    const detail = `${error?.stdout || ''}\n${error?.stderr || ''}\n${error?.message || ''}`;
    if (!/MK_E_UNAVAILABLE|Operation unavailable|GetActiveObject/.test(detail)) throw error;
  }
}

function officeCleanupScript(host, paths, processIds = []) {
  const targets = paths.map((path) => `@{ Path='${psSingle(resolve(path))}'; Name='${psSingle(basename(resolve(path)))}' }`).join(',');
  const processIdList = processIds.map((id) => Number(id)).filter((id) => Number.isInteger(id) && id > 0).join(',');
  const canonical = `function Canonical($value) { if ([string]::IsNullOrWhiteSpace($value)) { return '' }; try { return (Get-Item -LiteralPath $value -ErrorAction Stop).FullName.ToLowerInvariant() } catch { try { return [System.IO.Path]::GetFullPath($value).ToLowerInvariant() } catch { return '' } } }; $targets=@(${targets}); foreach ($targetSpec in $targets) { $targetSpec.Path = Canonical $targetSpec.Path };`;
  const helpers = `function Target-Matches($candidate,$caption) { $canonicalCandidate = Canonical $candidate; foreach ($targetSpec in $targets) { if ($canonicalCandidate -and $canonicalCandidate -eq $targetSpec.Path) { return $true }; if (-not [string]::IsNullOrWhiteSpace($caption) -and $caption -like ('*' + $targetSpec.Name + '*')) { return $true } }; return $false }; $processIds=@(${processIdList}); function Close-DriverOwnedDocuments([scriptblock]$CloseDocuments) { & $CloseDocuments }; function Close-DriverOwnedProcessIds { foreach ($processId in $processIds) { try { $process=Get-Process -Id $processId -ErrorAction Stop; $process.CloseMainWindow() | Out-Null; if (-not $process.WaitForExit(5000)) { Stop-Process -Id $processId -Force } } catch {} } }; function Ensure-DriverOwnedProcessIdsExited { foreach ($processId in $processIds) { try { $process=Get-Process -Id $processId -ErrorAction Stop; if (-not $process.WaitForExit(1000)) { Stop-Process -Id $processId -Force } } catch {} } }; function Maybe-QuitEmptyOfficeApplication($app,[int]$count) { if ($count -eq 0) { $app.Quit() } };`;
  if (host === 'word') {
    return `$ErrorActionPreference='Stop'; ${canonical} ${helpers} $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application'); $app.DisplayAlerts=0; Close-DriverOwnedDocuments { foreach ($doc in @($app.Documents)) { if (Target-Matches $doc.FullName $doc.Name) { $doc.Close($false) } }; foreach ($win in @($app.Windows)) { if (Target-Matches $win.Document.FullName $win.Caption) { $win.Document.Close($false) } } }; Maybe-QuitEmptyOfficeApplication $app $app.Documents.Count`;
  }
  if (host === 'excel') {
    return `$ErrorActionPreference='Stop'; ${canonical} ${helpers} function Close-ExcelByWindowTitle { foreach ($process in @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })) { foreach ($targetSpec in $targets) { if (-not [string]::IsNullOrWhiteSpace($process.MainWindowTitle) -and $process.MainWindowTitle -like ('*' + $targetSpec.Name + '*')) { try { $process.CloseMainWindow() | Out-Null; if (-not $process.WaitForExit(5000)) { Stop-Process -Id $process.Id -Force } } catch {} } } } }; try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); $app.DisplayAlerts=$false; Close-DriverOwnedDocuments { foreach ($wb in @($app.Workbooks)) { if (Target-Matches $wb.FullName $wb.Name) { $wb.Close($false) } } }; Maybe-QuitEmptyOfficeApplication $app $app.Workbooks.Count } catch { if ($_.Exception.Message -notmatch 'MK_E_UNAVAILABLE|Operation unavailable|GetActiveObject') { Write-Output $_.Exception.Message } }; Close-ExcelByWindowTitle; Close-DriverOwnedProcessIds; Ensure-DriverOwnedProcessIdsExited`;
  }
  return `$ErrorActionPreference='Stop'; ${canonical} ${helpers} function Close-PowerPointByWindowTitle { foreach ($process in @(Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })) { foreach ($targetSpec in $targets) { if (-not [string]::IsNullOrWhiteSpace($process.MainWindowTitle) -and $process.MainWindowTitle -like ('*' + $targetSpec.Name + '*')) { try { $process.CloseMainWindow() | Out-Null; if (-not $process.WaitForExit(5000)) { Stop-Process -Id $process.Id -Force } } catch {} } } } }; $app=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $app.DisplayAlerts=1; Close-DriverOwnedDocuments { foreach ($pres in @($app.Presentations)) { if (Target-Matches $pres.FullName $pres.Name) { $pres.Saved = $true; $pres.Close() } } }; Maybe-QuitEmptyOfficeApplication $app $app.Presentations.Count; Close-PowerPointByWindowTitle; Close-DriverOwnedProcessIds; Ensure-DriverOwnedProcessIdsExited`;
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
  const sessionId = await mcpSessionId(endpoint);
  const response = await postJson(endpoint, { jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name, arguments: args } }, sessionId);
  return response.body.result || response.body;
}

async function mcpResourceRead(endpoint = 'http://127.0.0.1:8800/mcp', uri) {
  const sessionId = await mcpSessionId(endpoint);
  const response = await postJson(endpoint, { jsonrpc: '2.0', id: 3, method: 'resources/read', params: { uri } }, sessionId);
  return response.body.result || response.body;
}

async function mcpToolsList(endpoint = 'http://127.0.0.1:8800/mcp') {
  const sessionId = await mcpSessionId(endpoint);
  const response = await postJson(endpoint, { jsonrpc: '2.0', id: 4, method: 'tools/list', params: {} }, sessionId);
  return response.body.result || response.body;
}

async function mcpSessionId(endpoint) {
  if (!mcpSessionIds.has(endpoint)) {
    mcpSessionIds.set(endpoint, await initializeMcp(endpoint));
  }
  return mcpSessionIds.get(endpoint);
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
        if (response.statusCode && response.statusCode >= 400) {
          reject(new Error(`MCP HTTP ${response.statusCode}: ${responseText || response.statusMessage || 'request failed'}`));
          return;
        }
        try {
          resolvePromise({ headers: response.headers, body: JSON.parse(responseText || '{}') });
        } catch (error) {
          reject(new Error(`MCP response was not JSON: ${responseText || error.message}`));
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
  throw new Error(`Office keeper did not create the ready sentinel within ${timeoutMs} ms for ${path}.${keeperStateDetail(path, keeper)}`);
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

function keeperStateDetail(readyPath, keeper) {
  if (!keeper) return '';
  const details = [
    `started=${fileState(keeper.startedPath)}`,
    `ready=${existsSync(readyPath) ? 'present' : 'missing'}`,
    `error=${fileState(keeper.errorPath)}`
  ];
  const logs = keeperLogDetail(keeper).trim();
  if (logs) details.push(logs);
  return ` ${details.join(' ')}`;
}

function fileState(path) {
  return path && existsSync(path) ? 'present' : 'missing';
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
    let stdout = '';
    let stderr = '';
    child.stdout?.setEncoding('utf8');
    child.stderr?.setEncoding('utf8');
    child.stdout?.on('data', (chunk) => { stdout += chunk; });
    child.stderr?.on('data', (chunk) => { stderr += chunk; });
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
      resolvePromise({ exitCode: code ?? 0, stdout, stderr });
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

function psQuoted(value) {
  return `'${psSingle(value)}'`;
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
