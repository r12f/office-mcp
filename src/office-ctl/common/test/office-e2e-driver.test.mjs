import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const DRIVER = fileURLToPath(new URL('./office-e2e-driver.mjs', import.meta.url));
const REPO_ROOT = resolve(dirname(DRIVER), '../../../..');
const DEFAULT_ACTIVATOR = resolve(REPO_ROOT, 'src/office-ctl/common/scripts/activate-office-mcp-addin.ps1');
const RUN_OFFICE_COM = process.env.OFFICE_MCP_RUN_E2E === '1';

test('Office E2E driver describes a driver-owned Word lifecycle', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-word-'));
  const create = runDriver({ host: 'Word', step: 'describeDocumentLifecycle', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-word-.*\.docx$/i);
  assert.equal(document.host, 'Word');
  assert.equal(document.createdByDriver, true);
  assert.equal(document.officeWindowMode, 'visible');
  assert.ok(document.keeper?.closePath, 'driver-owned close sentinel is required');
  assert.ok(document.keeper?.startedPath, 'keeper started sentinel is required');
  assert.ok(document.keeper?.pidPath, 'keeper pid file is required');
  assert.ok(document.keeper?.stdoutPath, 'keeper stdout log is required');
  assert.ok(document.keeper?.stderrPath, 'keeper stderr log is required');
  assert.match(document.script, /New-OfficeMcpBlankDocx/);
  assert.match(document.script, /Add-Type -AssemblyName System\.IO\.Compression;/);
  assert.match(document.script, /docProps\/core\.xml/);
  assert.match(document.script, /word\/_rels\/document\.xml\.rels/);
  assert.match(document.script, /word\/styles\.xml/);
  assert.match(document.script, /word\/settings\.xml/);
  assert.match(document.script, /compatibilityMode/);
  assert.match(document.script, /\$app\.DisplayAlerts=0/);
  assert.match(document.script, /Documents\.Open/);
  assert.match(document.script, /\$confirmConversions=\$false/);
  assert.match(document.script, /\$addToRecentFiles=\$false/);
  assert.doesNotMatch(document.script, /SaveAs2/, 'Word E2E creation must not depend on COM SaveAs2');
  assert.match(document.script, /Invoke-Retry/);
  assert.match(document.script, /\$i -lt 90/, 'Word COM creation needs a long retry window for busy desktop hosts');
  assert.match(document.script, /RPC_E_CALL_REJECTED/);
  assert.doesNotMatch(document.script, /;\s*# retries[^\r\n]*Set-Content/, 'retry comment must not comment out the started sentinel');
  assert.match(document.script, /<# retries RPC_E_CALL_REJECTED and transient Office COM busy states\. #>/);
  assert.doesNotMatch(document.script, /\.Content\.Text=/);
  assert.doesNotMatch(document.script, /\.Content\.InsertAfter/);
  assert.doesNotMatch(document.script, /TrackRevisions/);
  assert.match(document.script, /office-mcp-ready/);
  assert.match(document.script, /office-mcp-ready/);
  assert.doesNotMatch(document.script, /\.Quit\(\)/, 'keeper must not quit user Office applications');
});

test('Office E2E driver creates and cleans up Word documents through COM', { skip: !RUN_OFFICE_COM }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-word-com-'));
  const create = runDriver({ host: 'Word', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  const cleanup = runDriver({ host: 'Word', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  assert.equal(JSON.parse(cleanup.stdout).deleted, true);
});

test('Office E2E driver describes a driver-owned Excel lifecycle', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-excel-'));
  const create = runDriver({ host: 'Excel', step: 'describeDocumentLifecycle', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-excel-.*\.xlsx$/i);
  assert.equal(document.host, 'Excel');
  assert.equal(document.createdByDriver, true);
  assert.equal(document.officeWindowMode, 'visible');
  assert.ok(document.keeper?.closePath, 'driver-owned close sentinel is required');
  assert.ok(document.keeper?.startedPath, 'keeper started sentinel is required');
  assert.ok(document.keeper?.pidPath, 'keeper pid file is required');
  assert.ok(document.keeper?.stdoutPath, 'keeper stdout log is required');
  assert.ok(document.keeper?.stderrPath, 'keeper stderr log is required');
  assert.ok(Array.isArray(document.officeProcessIdsBefore), 'driver must record pre-existing Excel processes');
  assert.doesNotMatch(document.script, /GetActiveObject\('Excel\.Application'\)/);
  assert.match(document.script, /New-Object -ComObject Excel\.Application/);
  assert.match(document.script, /function Write-OfficeMcpProcessPid/);
  assert.match(document.script, /Write-OfficeMcpProcessPid -Handle \$app\.Hwnd/);
  assert.match(document.script, /office-mcp-pid/);
  assert.match(document.script, /Invoke-Retry \{ \$app\.Visible=\$true \}/);
  assert.match(document.script, /Invoke-Retry \{ \$app\.DisplayAlerts=\$false \}/);
  assert.match(document.script, /Workbooks\.Add\(\)/);
  assert.match(document.script, /Invoke-Retry \{ \$wb\.Worksheets\.Item\(1\) \}/);
  assert.match(document.script, /Invoke-Retry \{ \$ws\.Cells\.Item\(1,1\)\.Value2='office-mcp e2e baseline' \}/);
  assert.match(document.script, /office-mcp-ready/);
  assert.match(document.script, /office-mcp-ready/);
  assert.doesNotMatch(document.script, /\.Quit\(\)/, 'keeper must not quit user Office applications');
});

test('Office E2E driver creates and cleans up Excel workbooks through COM', { skip: !RUN_OFFICE_COM }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-excel-com-'));
  const create = runDriver({ host: 'Excel', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  const cleanup = runDriver({ host: 'Excel', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  assert.equal(JSON.parse(cleanup.stdout).deleted, true);
});

test('Office E2E driver describes a visible PowerPoint lifecycle', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-powerpoint-'));
  const create = runDriver({ host: 'PowerPoint', step: 'describeDocumentLifecycle', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-powerpoint-.*\.pptx$/i);
  assert.equal(document.host, 'PowerPoint');
  assert.equal(document.createdByDriver, true);
  assert.equal(document.officeWindowMode, 'visible');
  assert.ok(document.keeper?.closePath, 'driver-owned close sentinel is required');
  assert.ok(document.keeper?.startedPath, 'keeper started sentinel is required');
  assert.ok(document.keeper?.pidPath, 'keeper pid file is required');
  assert.ok(document.keeper?.stdoutPath, 'keeper stdout log is required');
  assert.ok(document.keeper?.stderrPath, 'keeper stderr log is required');
  assert.match(document.script, /Presentations\.Add\(\$true\)/);
  assert.match(document.script, /office-mcp-ready/);
  assert.match(document.script, /office-mcp-ready/);
  assert.doesNotMatch(document.script, /\.Quit\(\)/, 'keeper must not quit user Office applications');
});

test('Office E2E driver reuses the built daemon binary for status when available', () => {
  const result = runDriver({ host: 'Word', step: 'describeDaemonStatusCommand' });
  assert.equal(result.status, 0, result.stderr);
  const command = JSON.parse(result.stdout);
  if (existsSync(resolve(REPO_ROOT, 'target/debug/office-mcp-daemon.exe'))) {
    assert.match(command.command, /office-mcp-daemon\.exe$/);
    assert.deepEqual(command.args, ['daemon', 'status']);
  } else {
    assert.equal(command.command, 'cargo');
    assert.deepEqual(command.args, ['run', '-q', '-p', 'office-mcp-daemon', '--', 'daemon', 'status']);
  }
});

test('Office E2E driver cleanup canonicalizes Office document paths', () => {
  const result = runDriver({ host: 'Word', step: 'describeDocumentLifecycle', context: { workDir: mkdtempSync(join(tmpdir(), 'office-mcp-driver-cleanup-script-')) } });
  assert.equal(result.status, 0, result.stderr);
  const document = JSON.parse(result.stdout);
  assert.match(document.cleanupScript, /function Canonical/);
  assert.match(document.cleanupScript, /\[string\]::IsNullOrWhiteSpace\(\$value\)/);
  assert.match(document.cleanupScript, /catch \{ return '' \}/);
  assert.match(document.cleanupScript, /Target-Matches \$doc\.FullName \$doc\.Name/);
  assert.match(document.cleanupScript, /\$targets=@\(/);
  assert.match(document.cleanupScript, /\$targetSpec\.Name/);
  assert.match(document.cleanupScript, /\$win\.Caption/);
  assert.match(document.cleanupScript, /\$app\.DisplayAlerts=0/);
  assert.match(document.cleanupScript, /Close-DriverOwnedDocuments/);
  const driver = readFileSync(DRIVER, 'utf8');
  assert.match(driver, /11111111-aaaa-bbbb-cccc-222222222222/);
  assert.match(document.cleanupScript, /Maybe-QuitEmptyOfficeApplication/);
  assert.match(document.cleanupScript, /\$app\.Quit\(\)/);
});

test('Office E2E driver cleanup covers activated sideload copies and original fixtures', () => {
  const result = runDriver({ host: 'PowerPoint', step: 'describeDocumentLifecycle', context: { workDir: mkdtempSync(join(tmpdir(), 'office-mcp-driver-cleanup-sideload-')) } });
  assert.equal(result.status, 0, result.stderr);
  const document = JSON.parse(result.stdout);
  assert.match(document.cleanupScript, /office-mcp-e2e-powerpoint-fixture\.pptx/);
  assert.match(document.cleanupScript, /Close-DriverOwnedDocuments/);
  assert.match(document.cleanupScript, /Maybe-QuitEmptyOfficeApplication/);
  assert.match(document.cleanupScript, /PowerPoint\.Application[\s\S]*Close-DriverOwnedProcessIds; Ensure-DriverOwnedProcessIdsExited/);
  const driver = readFileSync(DRIVER, 'utf8');
  assert.match(driver, /officeSideloadCopyCandidates/);
  assert.match(driver, /Office add-in \$\{spec\.id\}\*\.\$\{spec\.extension\}/);
  assert.match(driver, /officeAppName\(normalizedHost\)\} add-in \$\{spec\.id\}\*\.\$\{spec\.extension\}/);
  assert.match(driver, /44444444-aaaa-bbbb-cccc-555555555555/);
  assert.match(driver, /const initialPaths = driverOwnedCleanupPaths\(document\)/);
  assert.match(driver, /const cleanupPaths = \[\.\.\.new Set\(\[\.\.\.initialPaths, \.\.\.driverOwnedCleanupPaths\(document\)\]\)\]/);
  assert.match(driver, /MK_E_UNAVAILABLE\|Operation unavailable\|GetActiveObject/);
  assert.match(driver, /function Close-ExcelByWindowTitle/);
  assert.match(driver, /function Close-DriverOwnedProcessIds/);
  assert.match(driver, /function driverOwnedEmptyOfficeProcessIds/);
  assert.match(driver, /title === '' \|\| title === officeAppName\(host\)/);
  assert.match(driver, /document\.officeProcessIdsBefore/);
  assert.match(driver, /const processIds = driverOwnedProcessIds\(document\)/);
  assert.match(driver, /MainWindowTitle -like/);
  assert.match(driver, /Stop-Process -Id \$process\.Id -Force/);
});

test('default Windows add-in activator closes the original file after sideload copy opens', () => {
  const script = readFileSync(resolve(REPO_ROOT, 'src/office-ctl/common/scripts/activate-office-mcp-addin.ps1'), 'utf8');
  assert.match(script, /function Close-DriverDocumentIfDifferent/);
  assert.match(script, /closing original powerpoint presentation after sideload copy/);
  assert.match(script, /Close-DriverDocumentIfDifferent -Application \$app -HostKey \$hostKey -OriginalPath \$DocumentPath -ActivePath \$activeDocumentPath/);
});

test('Office E2E driver cleanup retries deletion while Office releases the document lock', () => {
  const result = runDriver({ host: 'Word', step: 'describeDocumentLifecycle', context: { workDir: mkdtempSync(join(tmpdir(), 'office-mcp-driver-cleanup-retry-')) } });
  assert.equal(result.status, 0, result.stderr);
  const document = JSON.parse(result.stdout);
  assert.match(document.cleanupScript, /Close\(\$false\)/);
  const driver = readFileSync(DRIVER, 'utf8');
  assert.match(driver, /removeFileWithRetry/);
  assert.match(driver, /EPERM|EBUSY|EACCES/);
});

test('Office E2E driver cleanup removes activation logs', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-activation-log-cleanup-'));
  const path = join(dir, 'office-mcp-e2e-word-fixture.docx');
  const activationLogPath = join(dir, 'office-mcp-e2e-word-fixture.docx.office-mcp-activator.log');
  writeFileSync(path, 'fixture');
  writeFileSync(activationLogPath, 'activation log');
  const cleanup = runDriver({
    host: 'Word',
    step: 'cleanupDocument',
    context: {
      document: {
        host: 'Word',
        path,
        createdByDriver: true,
        activationLogPath,
        keeper: { closePath: join(dir, 'close') }
      }
    }
  });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  assert.equal(existsSync(activationLogPath), false);
});

test('Office E2E driver records activation logs before activation can fail', () => {
  const result = runDriver({ host: 'Excel', step: 'describeDocumentLifecycle', context: { workDir: mkdtempSync(join(tmpdir(), 'office-mcp-driver-activation-log-path-')) } });
  assert.equal(result.status, 0, result.stderr);
  const document = JSON.parse(result.stdout);
  assert.equal(document.activationLogPath, `${document.path}.office-mcp-activator.log`);
  assert.match(document.cleanupScript, /Close-DriverOwnedProcessIds/);
  assert.match(document.cleanupScript, /Ensure-DriverOwnedProcessIdsExited/);
  assert.match(document.cleanupScript, /office-mcp-e2e-excel-/);
});

test('Office E2E driver uses a visible PowerPoint window and safe cleanup', { skip: !RUN_OFFICE_COM }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-powerpoint-com-'));
  const create = runDriver({ host: 'PowerPoint', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  const cleanup = runDriver({ host: 'PowerPoint', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  const result = JSON.parse(cleanup.stdout);
  assert.equal(result.closedByDriver, true);
  assert.equal(result.deleted, true);
});

test('Office E2E driver activation step is explicit and configurable', () => {
  const previousDefault = process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR;
  process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR = '0';
  try {
    const skipped = runDriver({ host: 'Word', step: 'activateAddin', context: { document: { path: 'fixture.docx' }, daemon: { addinEndpoint: 'wss://localhost:8765/addin' } } });
    assert.equal(skipped.status, 0, skipped.stderr);
    assert.equal(JSON.parse(skipped.stdout).skipped, 'no-activator-configured');
  } finally {
    restoreEnv('OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR', previousDefault);
  }

  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-activator-'));
  const logPath = join(dir, 'activator-env.json');
  const activatorPath = join(dir, 'activator.mjs');
  writeFileSync(activatorPath, `
import { writeFileSync } from 'node:fs';
writeFileSync(${JSON.stringify(logPath)}, JSON.stringify({
  host: process.env.OFFICE_MCP_E2E_HOST,
  documentPath: process.env.OFFICE_MCP_E2E_DOCUMENT_PATH,
  addinOrigin: process.env.OFFICE_MCP_E2E_ADDIN_ORIGIN,
  addinEndpoint: process.env.OFFICE_MCP_E2E_ADDIN_ENDPOINT
}));
console.log(JSON.stringify({ activated: true, document_path: 'sideload-copy.pptx' }));
`);
  const previous = process.env.OFFICE_MCP_E2E_ACTIVATOR;
  process.env.OFFICE_MCP_E2E_ACTIVATOR = `${process.execPath} ${activatorPath}`;
  try {
    const activated = runDriver({
      host: 'PowerPoint',
      step: 'activateAddin',
      context: {
        document: { path: 'deck.pptx' },
        daemon: { addinOrigin: 'https://localhost:8765', addinEndpoint: 'wss://localhost:8765/addin' },
        timeoutMs: 5000
      }
    });
    assert.equal(activated.status, 0, activated.stderr);
    const activatedResult = JSON.parse(activated.stdout);
    assert.equal(activatedResult.activated, true);
    assert.equal(activatedResult.document_path, 'sideload-copy.pptx');
  } finally {
    restoreEnv('OFFICE_MCP_E2E_ACTIVATOR', previous);
  }

  const env = JSON.parse(readFileSync(logPath, 'utf8'));
  assert.equal(env.host, 'powerpoint');
  assert.equal(env.documentPath, 'deck.pptx');
  assert.equal(env.addinOrigin, 'https://localhost:8765');
  assert.equal(env.addinEndpoint, 'wss://localhost:8765/addin');
});

test('Office E2E driver provides a default Windows add-in activator', () => {
  const previousActivator = process.env.OFFICE_MCP_E2E_ACTIVATOR;
  const previousDryRun = process.env.OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN;
  const previousDefault = process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR;
  delete process.env.OFFICE_MCP_E2E_ACTIVATOR;
  process.env.OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN = '1';
  delete process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR;
  try {
    const activated = runDriver({
      host: 'Excel',
      step: 'activateAddin',
      context: {
        document: { path: 'book.xlsx' },
        daemon: { addinOrigin: 'https://localhost:8765', addinEndpoint: 'wss://localhost:8765/addin' },
        timeoutMs: 5000
      }
    });
    assert.equal(activated.status, 0, activated.stderr);
    const result = JSON.parse(activated.stdout);
    assert.equal(result.activated, true);
  assert.equal(result.activator_kind, 'default-windows-taskpane');
  assert.match(result.activator, /activate-office-mcp-addin\.ps1/);
  const driver = readFileSync(DRIVER, 'utf8');
  assert.match(driver, /OFFICE_MCP_E2E_ACTIVATOR_TIMEOUT_MS \|\| 95000/);
  assert.match(driver, /-TimeoutSeconds 90/);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_ACTIVATOR', previousActivator);
    restoreEnv('OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN', previousDryRun);
    restoreEnv('OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR', previousDefault);
  }
});

test('Office E2E driver accepts activation when the add-in session already registered', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-activation-session-'));
  const activatorPath = join(dir, 'activator.mjs');
  const serverPath = join(dir, 'mcp-server.mjs');
  const documentPath = join(dir, 'fixture.xlsx');
  writeFileSync(documentPath, 'fixture');
  writeFileSync(activatorPath, `process.exit(1);`);
  writeFileSync(serverPath, `
import { createServer } from 'node:http';
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { sessions: [{ session_id: 'excel-session', app: 'excel', document: { filename: 'fixture.xlsx' }, available_tools: ['excel.read_range'] }] } } }));
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const previous = process.env.OFFICE_MCP_E2E_ACTIVATOR;
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    process.env.OFFICE_MCP_E2E_ACTIVATOR = `${process.execPath} ${activatorPath}`;
    const result = runDriver({
      host: 'Excel',
      step: 'activateAddin',
      context: {
        daemon: { endpoint },
        document: { path: documentPath, createdByDriver: true }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    const activation = JSON.parse(result.stdout);
    assert.equal(activation.activated, true);
    assert.equal(activation.activation_path, 'session-detected-after-activator-failure');
  } finally {
    restoreEnv('OFFICE_MCP_E2E_ACTIVATOR', previous);
    server.kill();
  }
});

test('Office E2E driver reuses one MCP HTTP session per driver process', () => {
  const driver = readFileSync(DRIVER, 'utf8');
  assert.match(driver, /const mcpSessionIds = new Map\(\);/);
  assert.match(driver, /async function mcpSessionId\(endpoint\)/);
  assert.match(driver, /mcpSessionIds\.set\(endpoint, await initializeMcp\(endpoint\)\)/);
  assert.match(driver, /const sessionId = await mcpSessionId\(endpoint\);/);
  assert.match(driver, /MCP HTTP \$\{response\.statusCode\}/);
  assert.match(driver, /MCP response was not JSON/);
});

test('default Windows add-in activator can fall back through My Add-ins catalog UI', () => {
  const script = readFileSync(DEFAULT_ACTIVATOR, 'utf8');
  assert.match(script, /office-addin-dev-settings/);
  assert.match(script, /function Invoke-OfficialRegistration/);
  assert.match(script, /office-addin-dev-settings register/);
  assert.match(script, /official registration start app=\$appName manifest=\$ManifestPath/);
  assert.match(script, /registered ribbon path active; attempting to open control panel document=\$DocumentPath/);
  assert.match(script, /Try-OpenControlPanelForDriverDocument -WindowHandle \$driverWindowHandle -Deadline \(Get-Date\)\.AddSeconds\(\[Math\]::Min\(12, \[Math\]::Max\(3, \$TimeoutSeconds \/ 2\)\)\) -AllowCatalogFallback:\$false/);
  assert.match(script, /-ActivationPath "official-registration"/);
  assert.doesNotMatch(script, /catalog fallback skipped for excel/);
  assert.match(script, /sideload/);
  assert.match(script, /cmd\.exe/);
  assert.match(script, /OFFICE_MCP_E2E_MANIFEST_PATH/);
  assert.match(script, /officialDocumentPath/);
  assert.match(script, /Launching .* via/);
  assert.match(script, /function Open-OfficialSideloadDocument/);
  assert.match(script, /Open-OfficialSideloadDocument -Application \$app -HostKey \$hostKey -Path \$activeDocumentPath/);
  assert.match(script, /Start-Process -FilePath \$Path/);
  assert.doesNotMatch(script, /Start-Process -FilePath \$Path -WindowStyle Hidden/, 'sideload workbooks must be opened visibly so the add-in task pane can load');
  assert.match(script, /official sideload document opened via shell path=\$Path/);
  assert.match(script, /official sideload excel document opened via new application path=\$Path/);
  assert.match(script, /return \$excel/);
  assert.match(script, /\$openedApplication = Open-OfficialSideloadDocument -Application \$app -HostKey \$hostKey -Path \$activeDocumentPath/);
  assert.match(script, /if \(\$openedApplication\) \{ \$app = \$openedApplication \}/);
  assert.match(script, /try \{\s*\$driverWindowHandle = Wait-ForDriverDocument -Application \$app -HostKey \$hostKey -Path \$activeDocumentPath -Deadline \$copyDeadline\s*\} catch \{/s);
  assert.match(script, /official sideload copy not active yet; opening document path=\$activeDocumentPath/);
  assert.match(script, /New-Object -ComObject Excel\.Application/);
  assert.match(script, /function Ensure-ExcelSideloadWebExtension/);
  assert.match(script, /function Get-ExcelApplicationForWorkbookOpen/);
  assert.match(script, /excel application handle was stale before workbook reopen/);
  assert.match(script, /excel application reacquired before workbook reopen/);
  assert.match(script, /excel application created before workbook reopen/);
  assert.match(script, /function Open-ExcelWorkbookAfterWebExtensionPatch/);
  assert.match(script, /excel \$Source workbook reopen failed on current application; creating replacement application/);
  assert.match(script, /function Try-OpenExcelPatchedDriverWorkbook/);
  assert.match(script, /function Try-OpenPatchedDriverDocument/);
  assert.match(script, /Try-OpenExcelPatchedDriverWorkbook -Application \$Application -Path \$Path -Deadline \$Deadline/);
  assert.match(script, /Try-OpenPatchedDriverDocument -Application \$app -HostKey \$hostKey -Path \$DocumentPath -Deadline \$deadline/);
  assert.match(script, /word patched driver document unavailable/);
  assert.match(script, /powerpoint patched driver presentation unavailable/);
  assert.match(script, /\$hostKey -in @\("word", "excel", "powerpoint"\)/);
  assert.match(script, /Try-OpenControlPanelForDriverDocument -WindowHandle \$patchedDriverWindowHandle -Deadline \$deadline -AllowCatalogFallback/);
  assert.doesNotMatch(script, /Try-OpenControlPanelForDriverDocument -WindowHandle \$patchedDriverWindowHandle -Deadline \$deadline -AllowCatalogFallback:\$false/);
  assert.match(script, /\$hostKey patched driver document control panel did not open; skipping official sideload fallback to avoid duplicate \$hostKey windows/);
  assert.match(script, /patched driver document did not open Office MCP Control/);
  assert.doesNotMatch(script, /excel patched driver workbook control panel did not open; skipping official sideload fallback to avoid duplicate Excel windows/);
  assert.match(script, /function Reopen-ExcelSideloadDocumentAfterWebExtensionPatch/);
  assert.match(script, /Reopen-ExcelSideloadDocumentAfterWebExtensionPatch -Application \$app -Path \$activeDocumentPath/);
  assert.match(script, /Open-ExcelWorkbookAfterWebExtensionPatch -Application \$Application -Path \$Path -Source "sideload"/);
  assert.match(script, /Open-ExcelWorkbookAfterWebExtensionPatch -Application \$Application -Path \$Path -Source "driver"/);
  assert.match(script, /excel \$Source workbook closed before webextension patch path=\$Path/);
  assert.match(script, /excel \$Source workbook reopened after webextension patch path=\$Path/);
  assert.match(script, /Ensure-ExcelSideloadWebExtension -WorkbookPath \$activeDocumentPath/);
  assert.match(script, /application\/vnd\.ms-office\.webextension\+xml/);
  assert.match(script, /application\/vnd\.ms-office\.webextensiontaskpanes\+xml/);
  assert.match(script, /http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/webextensiontaskpanes/);
  assert.match(script, /Office\.AutoShowTaskpaneWithDocument/);
  assert.match(script, /function Try-OpenControlPanelForDriverDocument/);
  assert.match(script, /Office MCP Control direct click did not show Open Control Panel; trying catalog confirmation/);
  assert.match(script, /\$confirmDeadline = \(Get-Date\)\.AddSeconds\(12\)/);
  assert.match(script, /Try-ConfirmCatalogAddinInstall -WindowHandle \$WindowHandle -Deadline \$confirmDeadline -Source \$Source/);
  assert.match(script, /Wait-ForOpenControlPanel[\s\S]*Office MCP Control direct click did not show Open Control Panel; trying catalog confirmation[\s\S]*Try-ConfirmCatalogAddinInstall/);
  assert.match(script, /\$hostKey -ne "excel"/);
  assert.match(script, /\$tabNames = @\("Add-ins", "Home", "Insert", "My Add-ins"\)/);
  assert.match(script, /if \(\$hostKey -eq "excel" -and -not \$AllowCatalogFallback\) \{\s*\$tabNames = @\("Add-ins"\)\s*\}/s);
  assert.doesNotMatch(script, /\$copyWaitSeconds = if \(\$hostKey -eq "excel"\) \{ 2 \}/, 'Excel sideload copy wait must not be shorter than the Office add-in launch path');
  assert.match(script, /\$copyWaitSeconds = \[Math\]::Min\(8, \[Math\]::Max\(3, \$TimeoutSeconds \/ 3\)\)/);
  assert.match(script, /official sideload copy is active; attempting to open control panel/);
  assert.match(script, /Try-OpenControlPanelForDriverDocument -WindowHandle \$driverWindowHandle -Deadline \$deadline -AllowCatalogFallback/);
  assert.match(script, /control panel best-effort timed out/);
  assert.match(script, /activationPath = if/);
  assert.match(script, /if \(\$panel\.opened\) \{/);
  assert.match(script, /official sideload control panel did not open; continuing activation fallback/);
  assert.doesNotMatch(script, /-ControlOpened \$panel\.opened/);
  assert.match(script, /official-sideload/);
  assert.match(script, /control_opened = \$ControlOpened/);
  assert.match(script, /official sideload timed out/);
  assert.match(script, /Start-Process/);
  assert.match(script, /Wait-ForDriverDocument/);
  assert.match(script, /RootElement/);
  assert.match(script, /Find-GlobalControlByName/);
  assert.match(script, /Find-GlobalControlByAutomationId/);
  assert.match(script, /Wait-ForOpenControlPanel/);
  assert.match(script, /function Try-InvokeNamedControl/);
  assert.match(script, /Find-DescendantByNameLike -Root \$Root -Name \$Name -ProcessId \$processId/);
  assert.match(script, /Find-DescendantByNameLike -Root \$Root -Name \$Name/);
  assert.match(script, /function Invoke-ControlByMouse/);
  assert.match(script, /Invoke-ControlByMouse -Element \$control/);
  assert.doesNotMatch(script, /function Try-InvokeNamedControl[\s\S]*Write-ActivatorLog "invoking control name=\$Name"[\s\S]*Invoke-Control -Element \$control[\s\S]*return \$true[\s\S]*function Try-InvokeAutomationIdControl/);
  assert.match(script, /Try-EnableOfficeMcpAddin/);
  assert.match(script, /Find-DescendantByNameLike -Root \$Root -Name "Office MCP Control"/);
  assert.match(script, /"ControlType\.TabItem"/);
  assert.match(script, /function Try-InvokeOfficeAddinsRibbon/);
  assert.match(script, /Try-InvokeOfficeAddinsRibbon -WindowHandle \$WindowHandle -Window \$window -Deadline \$Deadline -Source "current:\$tabName"/);
  assert.match(script, /OfficeExtensionsShowAddinFlyout/);
  assert.match(script, /Find-GlobalControlByAutomationId -AutomationId \$automationId -ProcessId \$processId/);
  assert.match(script, /\$processId = Get-OfficeProcessIdFromHandle -Handle \$WindowHandle/);
  assert.match(script, /Try-EnableOfficeMcpAddin -Root \$window -WindowHandle \$WindowHandle -Deadline \$Deadline -Source "current:\$tabName"/);
  assert.match(script, /Try-OpenAddinFromCatalog -WindowHandle \$WindowHandle -Window \$nextWindow -Deadline \$Deadline/);
  assert.match(script, /activation_path = "catalog-fallback"/);
  assert.match(script, /control panel visible control sample failed:/);
  assert.match(script, /function Try-ConfirmCatalogAddinInstall/);
  assert.match(script, /\$shortPanelDeadline = \(Get-Date\)\.AddSeconds\(4\)/);
  assert.match(script, /Wait-ForOpenControlPanel -WindowHandle \$WindowHandle -Deadline \$shortPanelDeadline -Source \$Source/);
  assert.match(script, /catalog install confirm invoked name=\$name source=\$Source/);
  assert.match(script, /Find-DescendantByNameLike -Root \$window -Name \$name/);
  assert.match(script, /function Try-OpenControlPanelFromRibbonTabs/);
  assert.match(script, /Try-OpenControlPanelFromRibbonTabs -WindowHandle \$WindowHandle -Deadline \$Deadline -Source "catalog-confirm:\$name"/);
  assert.match(script, /\$tabNames = @\("Add-ins", "Home", "Insert", "My Add-ins"\)/);
  assert.match(script, /\$postCatalogTabNames = @\("Home", "Insert", "Add-ins", "My Add-ins"\)/);
  assert.doesNotMatch(script, /\$tabNames = if \(\$hostKey -eq "excel"\)/);
  assert.match(script, /post-catalog tab scan invoked name=\$tabName source=\$Source/);
  assert.match(script, /function Try-DismissCatalogOverlay/);
  assert.match(script, /function Try-DismissOfficeModalDialog/);
  assert.match(script, /You cannot close Microsoft Excel because a dialog box is open/);
  assert.match(script, /office modal dialog dismissed name=\$name source=\$Source/);
  assert.match(script, /Try-DismissOfficeModalDialog -WindowHandle \$WindowHandle -Deadline \$Deadline -Source "catalog-confirm:\$name"/);
  assert.match(script, /Try-DismissCatalogOverlay -WindowHandle \$WindowHandle -Deadline \$Deadline -Source "catalog-confirm:\$name"/);
  assert.match(script, /function Send-ActivatorKey/);
  assert.match(script, /keybd_event/);
  assert.match(script, /Send-ActivatorKey -WindowHandle \$WindowHandle -Key "\{ESC\}"/);
  assert.match(script, /catalog overlay dismiss sending escape source=\$Source/);
  assert.match(script, /catalog overlay dismiss sent escape source=\$Source/);
  assert.match(script, /catalog overlay dismiss sent second escape source=\$Source/);
  assert.match(script, /foreach \(\$name in @\("Back", "Cancel", "Done"\)\)/);
  assert.match(script, /Find-DescendantByNameLike -Root \$window -Name \$name -ProcessId \$officeProcessId/);
  assert.match(script, /Find-DescendantByNameLike -Root \$window -Name \$name/);
  assert.doesNotMatch(script, /foreach \(\$name in @\("Close", "Back", "Cancel", "Done"\)\)/);
  assert.match(script, /Office MCP Control was clicked but Open Control Panel did not appear yet/);
  assert.match(script, /found Office MCP Control source=\$Source/);
  assert.doesNotMatch(script, /control_name\s*=\s*"Office MCP Control"/, 'clicking the add-in entry alone is not a completed activation');
  assert.doesNotMatch(script, /foreach \(\$name in @\("Insert", "Add-ins", "My Add-ins", "Shared Folder", "Office MCP Control", "Add"\)\)/);
  assert.match(script, /foreach \(\$name in @\("Office MCP Control", "Open Control Panel", "My Add-ins", "Shared Folder"\)\)/);
  assert.match(script, /Get-OfficeStateSnapshot/);
  assert.match(script, /waiting for driver document/);
  assert.match(script, /\$Path -match '\^https\?:\/\/'/);
  assert.match(script, /return \$Path\.Trim\(\)\.ToLowerInvariant\(\)/);
  assert.match(script, /official sideload copy was not visible/);
  assert.match(script, /official sideload failed/);
  assert.match(script, /official sideload error detail/);
  assert.match(script, /Get-PowerPointMainWindowHandle/);
  assert.match(script, /Get-ExcelMainWindowHandle/);
  assert.match(script, /if \(\$HostKey -eq "excel"\) \{ return Get-ExcelMainWindowHandle -Path \$Path \}/);
  assert.match(script, /using excel main window fallback/);
  assert.match(script, /workbookName -eq \$targetName/);
  assert.match(script, /presentationName -eq \$targetName/);
  assert.match(script, /powerpoint presentation window handle unavailable/);
  assert.match(script, /using powerpoint main window fallback/);
  assert.match(script, /function Test-ActivatorDeadline/);
  assert.match(script, /Test-ActivatorDeadline -Deadline \$Deadline/);
  assert.match(script, /try \{/);
  assert.match(script, /My Add-ins/);
  assert.match(script, /Office MCP Control/);
  assert.match(script, /Shared Folder/);
  assert.match(script, /Add/);
});

test('default Windows add-in activator exits before the external driver timeout', () => {
  const driver = readFileSync(DRIVER, 'utf8');
  assert.match(driver, /OFFICE_MCP_E2E_ACTIVATOR_TIMEOUT_MS \|\| 95000/);
  assert.match(driver, /-TimeoutSeconds 90/);
});

test('Office E2E driver callTool posts MCP requests through an injectable endpoint', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-mcp-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: JSON.parse(body) }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 1, result: { protocolVersion: '2024-11-05' } }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 2, result: { structuredContent: { ok: true } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'Word',
      step: 'callTool',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: { call: { name: 'word.get_text', arguments: { limit: 1 } } }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).structuredContent.ok, true);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[0].body.method, 'initialize');
    assert.equal(requests[1].session, 'mcp-session-test');
    assert.equal(requests[1].body.params.name, 'word.get_text');
    assert.equal(requests[1].body.params.arguments.session_id, 'session-1');
  } finally {
    server.kill();
  }
});

test('Office E2E driver listTools reads daemon MCP tools/list names', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-list-tools-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { tools: [{ name: 'office.list_sessions' }, { name: 'word.get_text' }, { name: 'excel.read_range' }] } }));
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({ host: 'Word', step: 'listTools', context: { daemon: { endpoint } } });
    assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(JSON.parse(result.stdout), ['office.list_sessions', 'word.get_text', 'excel.read_range']);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.ok(requests.some((entry) => entry.body.method === 'tools/list'));
  } finally {
    server.kill();
  }
});

test('Office E2E driver verifies readback expectations through an MCP read tool', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-readback-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
    } else {
      const text = ${JSON.stringify('updated marker remains\nfirst marker\nsecond marker')};
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { text, slides: [{ slide_index: 0, layout_name: 'Blank' }] } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'Word',
      step: 'verifyResult',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        result: { structuredContent: { shape_id: 'shape-999' } },
        toolCase: {
          tool: 'powerpoint.add_table',
          verify: {
            kind: 'readback',
            readbackTool: 'powerpoint.read_table',
            readbackArguments: { slide_index: 0, shape_id: '${result.shape_id}' },
            expect: {
              contains: ['updated marker'],
              notContains: ['baseline marker'],
              orderedContains: ['first marker', 'second marker'],
              pathEquals: [{ path: 'slides.0.layout_name', value: 'Blank' }],
              pathMissing: ['slides.1']
            }
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).verified, true);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[1].body.params.name, 'powerpoint.read_table');
    assert.equal(requests[1].body.params.arguments.session_id, 'session-1');
    assert.equal(requests[1].body.params.arguments.shape_id, 'shape-999');
  } finally {
    server.kill();
  }
});

test('Office E2E driver verifies readback expectations through an MCP resource', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-resource-readback-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    if (parsed.method === 'resources/read') {
      const text = JSON.stringify({ comments: [{ comment_id: 'comment-1', content: 'Resolve me E2E', resolved: true }] });
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { contents: [{ uri: parsed.params.uri, mimeType: 'application/json', text }] } }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { comment_id: 'comment-1' } } }));
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'Word',
      step: 'verifyResult',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        result: { structuredContent: { comment_id: 'comment-1' } },
        toolCase: {
          tool: 'word.resolve_comment',
          verify: {
            kind: 'readback',
            resource: 'office://word/${session_id}/comments',
            expect: {
              contains: ['Resolve me E2E'],
              pathEquals: [{ path: 'comments.0.resolved', value: true }]
            }
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).readbackResource, 'office://word/session-1/comments');
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    const resourceRead = requests.find((entry) => entry.body.method === 'resources/read');
    assert.equal(resourceRead.body.params.uri, 'office://word/session-1/comments');
  } finally {
    server.kill();
  }
});

test('Office E2E driver includes readback detail when an expectation fails', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-readback-detail-'));
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { createServer } from 'node:http';
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { actual: 'metadata without marker' } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'Excel',
      step: 'verifyResult',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        result: { structuredContent: { table: 'MissingTable' } },
        toolCase: {
          tool: 'excel.create_table',
          verify: {
            kind: 'readback',
            readbackTool: 'excel.update_table',
            readbackArguments: { table: 'MissingTable', action: 'metadata' },
            expect: { contains: ['ExpectedTable'] }
          }
        }
      }
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /actual readback:/);
    assert.match(result.stderr, /metadata without marker/);
  } finally {
    server.kill();
  }
});

test('Office E2E driver verifies direct-result expectations', async () => {
  const result = runDriver({
    host: 'PowerPoint',
    step: 'verifyResult',
    context: {
      toolCase: {
        tool: 'powerpoint.insert_image',
        verify: {
          kind: 'direct-result',
          expect: {
            pathEquals: [
              { path: 'inserted_image', value: true },
              { path: 'mime_type', value: 'image/png' }
            ]
          }
        }
      },
      result: {
        structuredContent: {
          inserted_image: true,
          mime_type: 'image/png'
        }
      }
    }
  });

  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(result.stdout).verified, true);
});

test('Office E2E driver fails direct-result expectation mismatches', async () => {
  const result = runDriver({
    host: 'PowerPoint',
    step: 'verifyResult',
    context: {
      toolCase: {
        tool: 'powerpoint.insert_image',
        verify: {
          kind: 'direct-result',
          expect: { pathEquals: [{ path: 'inserted_image', value: true }] }
        }
      },
      result: { structuredContent: { inserted_image: false } }
    }
  });

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /inserted_image/);
});

test('Office E2E driver accepts declared host capability errors', async () => {
  const result = runDriver({
    host: 'PowerPoint',
    step: 'verifyResult',
    context: {
      toolCase: {
        tool: 'powerpoint.export_file',
        verify: {
          kind: 'direct-result',
          allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
          expect: { pathEquals: [{ path: 'format', value: 'pdf' }] }
        }
      },
      result: {
        structuredContent: {
          error: {
            office_mcp_code: 'HOST_CAPABILITY_UNAVAILABLE',
            message: 'PowerPoint desktop file export is not available through Office.context.document.getFileAsync in this host.',
            retriable: true,
            partial_effect: 'none'
          }
        }
      }
    }
  });

  assert.equal(result.status, 0, result.stderr);
  const verified = JSON.parse(result.stdout);
  assert.equal(verified.verified, true);
  assert.equal(verified.acceptedErrorCode, 'HOST_CAPABILITY_UNAVAILABLE');
});

test('Office E2E driver setupContent runs declared MCP setup actions', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-setup-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { ok: true } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'Word',
      step: 'setupContent',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: {
          tool: 'word.replace_text',
          setup: {
            actions: [
              { tool: 'word.replace_text', arguments: { find: 'office-mcp e2e baseline', replace: 'baseline marker' } }
            ]
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).actions, 1);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[1].body.params.name, 'word.replace_text');
    assert.equal(requests[1].body.params.arguments.session_id, 'session-1');
    assert.equal(requests[1].body.params.arguments.replace, 'baseline marker');
  } finally {
    server.kill();
  }
});

test('Office E2E driver setupContent accepts declared host capability errors', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-setup-capability-'));
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { createServer } from 'node:http';
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { error: { office_mcp_code: 'HOST_CAPABILITY_UNAVAILABLE', message: 'not available' } } } }));
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'PowerPoint',
      step: 'setupContent',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: {
          setup: {
            actions: [
              { tool: 'powerpoint.add_table', saveAs: 'table', allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'], arguments: { slide_index: 0, values: [['A']] } }
            ]
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    const output = JSON.parse(result.stdout);
    assert.deepEqual(output.bindings.table, { skipped: true, accepted_error_code: 'HOST_CAPABILITY_UNAVAILABLE' });
  } finally {
    server.kill();
  }
});

test('Office E2E driver resolves setup action result references in later calls', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-bindings-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
    } else if (parsed.params?.name === 'powerpoint.add_table') {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { shape_id: 'shape-123' } } }));
    } else if (parsed.params?.name === 'word.insert_content_control') {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { content_control: { content_control_id: 42 } } } }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { text: 'Updated table cell' } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const setup = {
      actions: [
        { tool: 'powerpoint.add_table', saveAs: 'table', arguments: { slide_index: 0, values: [['Old']] } }
      ]
    };
    const setupResult = runDriver({
      host: 'PowerPoint',
      step: 'setupContent',
      context: { daemon: { endpoint }, session: { sessionId: 'session-1' }, toolCase: { setup } }
    });
    assert.equal(setupResult.status, 0, setupResult.stderr);
    const bindings = JSON.parse(setupResult.stdout).bindings;
    assert.equal(bindings.table.shape_id, 'shape-123');

    const callResult = runDriver({
      host: 'PowerPoint',
      step: 'callTool',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1', bindings },
        toolCase: {
          call: {
            name: 'powerpoint.update_table',
            arguments: { slide_index: 0, shape_id: '${table.shape_id}', action: 'set_cell', row_index: 0, column_index: 0, value: 'Updated table cell' }
          }
        }
      }
    });
    assert.equal(callResult.status, 0, callResult.stderr);

    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[3].body.params.name, 'powerpoint.update_table');
    assert.equal(requests[3].body.params.arguments.shape_id, 'shape-123');

    const nestedSetupResult = runDriver({
      host: 'Word',
      step: 'setupContent',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: {
          setup: {
            actions: [
              { tool: 'word.insert_content_control', saveAs: 'controlResult', arguments: { tag: 'e2e' } }
            ]
          }
        }
      }
    });
    assert.equal(nestedSetupResult.status, 0, nestedSetupResult.stderr);
    const nestedBindings = JSON.parse(nestedSetupResult.stdout).bindings;
    assert.equal(nestedBindings.controlResult.content_control.content_control_id, 42);

    const nestedCallResult = runDriver({
      host: 'Word',
      step: 'callTool',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1', bindings: nestedBindings },
        toolCase: {
          call: {
            name: 'word.update_content_control',
            arguments: { content_control_id: '${controlResult.content_control.content_control_id}', text: 'Updated control' }
          }
        }
      }
    });
    assert.equal(nestedCallResult.status, 0, nestedCallResult.stderr);

    const updatedRequests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(updatedRequests[7].body.params.name, 'word.update_content_control');
    assert.equal(updatedRequests[7].body.params.arguments.content_control_id, 42);
  } finally {
    server.kill();
  }
});

test('Office E2E driver binds resource read results for later setup actions', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-resource-bindings-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    if (parsed.method === 'resources/read') {
      const text = JSON.stringify({ changes: [{ index: 0, fingerprint: 'fp-123' }] });
      const contents = [{ uri: parsed.params.uri, mimeType: 'application/json', text }];
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { contents } }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { ok: true } } }));
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'Word',
      step: 'setupContent',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: {
          tool: 'word.update_tracked_change',
          setup: {
            actions: [
              { resource: 'office://word/${session_id}/track_changes', saveAs: 'trackChanges' },
              { tool: 'word.update_tracked_change', arguments: { change_index: 0, action: 'accept', expected_fingerprint: '${trackChanges.changes.0.fingerprint}' } }
            ]
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    const resourceRead = requests.find((entry) => entry.body.method === 'resources/read');
    const toolCall = requests.find((entry) => entry.body.method === 'tools/call' && entry.body.params.name === 'word.update_tracked_change');
    assert.equal(resourceRead.body.params.uri, 'office://word/session-1/track_changes');
    assert.equal(toolCall.body.params.arguments.expected_fingerprint, 'fp-123');
  } finally {
    server.kill();
  }
});

test('Office E2E driver can seed a Word tracked change as a driver setup action', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-tracked-change-'));
  const fakePowerShell = join(dir, 'fake-powershell.mjs');
  const scriptLog = join(dir, 'script.txt');
  writeFileSync(fakePowerShell, `
import { appendFileSync } from 'node:fs';
const script = process.argv.slice(2).join(' ');
appendFileSync(${JSON.stringify(scriptLog)}, script);
console.log(JSON.stringify({ document: 'C:/tmp/seed.docx', marker: 'Tracked change seed', revisions: 1 }));
`);
  const previous = process.env.OFFICE_MCP_E2E_POWERSHELL_INLINE;
  process.env.OFFICE_MCP_E2E_POWERSHELL_INLINE = `${process.execPath} ${fakePowerShell}`;
  try {
    const result = runDriver({
      host: 'Word',
      step: 'setupContent',
      context: {
        document: { path: join(dir, 'seed.docx') },
        session: { sessionId: 'session-1' },
        toolCase: {
          setup: {
            actions: [
              { driver: 'word.create_tracked_change', saveAs: 'trackedChangeSeed', arguments: { text: 'Tracked change seed' } }
            ]
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    const output = JSON.parse(result.stdout);
    assert.equal(output.bindings.trackedChangeSeed.revisions, 1);
    const script = readFileSync(scriptLog, 'utf8');
    assert.match(script, /TrackRevisions = \$true/);
    assert.match(script, /Tracked change seed/);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_POWERSHELL_INLINE', previous);
  }
});

test('Office E2E driver cleanup ignores documents it did not create', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-cleanup-'));
  const path = join(dir, 'user-owned.docx');
  writeFileSync(path, 'do not delete');
  const cleanup = runDriver({ host: 'Word', step: 'cleanupDocument', context: { document: { path } } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  const result = JSON.parse(cleanup.stdout);
  assert.equal(result.deleted, false);
  assert.equal(result.skipped, 'not-driver-owned');
  assert.equal(existsSync(path), true);
});

test('Office E2E driver reports keeper diagnostics when document creation never becomes ready', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-create-fail-'));
  const fakePowerShell = join(dir, 'fake-powershell.mjs');
  writeFileSync(fakePowerShell, `
import { appendFileSync } from 'node:fs';
appendFileSync(1, 'fake keeper stdout\\n');
appendFileSync(2, 'fake keeper stderr\\n');
`);
  const previous = process.env.OFFICE_MCP_E2E_POWERSHELL;
  process.env.OFFICE_MCP_E2E_POWERSHELL = `${process.execPath} ${fakePowerShell}`;
  try {
    const result = runDriver({
      host: 'Word',
      step: 'createDocument',
      context: { workDir: dir, keeperTimeoutMs: 50 }
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /Office keeper did not create the ready sentinel/);
    assert.match(result.stderr, /started=missing/);
    assert.match(result.stderr, /ready=missing/);
    assert.match(result.stderr, /error=missing/);
    assert.match(result.stderr, /stdout: fake keeper stdout/);
    assert.match(result.stderr, /stderr: fake keeper stderr/);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_POWERSHELL', previous);
  }
});

test('Office E2E driver includes keeper logs when document creation exits with an error', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-create-error-'));
  const fakePowerShell = join(dir, 'fake-powershell-error.mjs');
  writeFileSync(fakePowerShell, `
import { appendFileSync } from 'node:fs';
appendFileSync(1, 'keeper started before failure\\n');
appendFileSync(2, 'simulated COM failure\\n');
process.exit(1);
`);
  const previous = process.env.OFFICE_MCP_E2E_POWERSHELL;
  process.env.OFFICE_MCP_E2E_POWERSHELL = `${process.execPath} ${fakePowerShell}`;
  try {
    const result = runDriver({
      host: 'Word',
      step: 'createDocument',
      context: { workDir: dir, keeperTimeoutMs: 50 }
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /Office keeper PowerShell exited with status 1/);
    assert.match(result.stderr, /stdout: keeper started before failure/);
    assert.match(result.stderr, /stderr: simulated COM failure/);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_POWERSHELL', previous);
  }
});

test('Office E2E driver removes keeper sidecars when document creation fails', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-create-cleanup-'));
  const fakePowerShell = join(dir, 'fake-powershell-error.mjs');
  writeFileSync(fakePowerShell, `
import { appendFileSync } from 'node:fs';
appendFileSync(1, 'keeper started before failure\n');
appendFileSync(2, 'simulated COM failure\n');
process.exit(1);
`);
  const previous = process.env.OFFICE_MCP_E2E_POWERSHELL;
  process.env.OFFICE_MCP_E2E_POWERSHELL = `${process.execPath} ${fakePowerShell}`;
  try {
    const result = runDriver({
      host: 'Excel',
      step: 'createDocument',
      context: { workDir: dir, keeperTimeoutMs: 50 }
    });
    assert.notEqual(result.status, 0);
    const residue = readdirSync(dir).filter((name) => name.startsWith('office-mcp-e2e-excel-'));
    assert.deepEqual(residue, []);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_POWERSHELL', previous);
  }
});

test('Office E2E driver includes keeper logs when document creation process times out', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-create-timeout-'));
  const fakePowerShell = join(dir, 'fake-powershell-timeout.mjs');
  writeFileSync(fakePowerShell, `
import { appendFileSync } from 'node:fs';
appendFileSync(1, 'keeper started before timeout\\n');
setTimeout(() => {}, 10000);
`);
  const previous = process.env.OFFICE_MCP_E2E_POWERSHELL;
  process.env.OFFICE_MCP_E2E_POWERSHELL = `${process.execPath} ${fakePowerShell}`;
  try {
    const result = runDriver({
      host: 'Word',
      step: 'createDocument',
      context: { workDir: dir, keeperTimeoutMs: 50, powerShellTimeoutMs: 100 }
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /ETIMEDOUT/);
    assert.match(result.stderr, /stderr: .*ETIMEDOUT/);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_POWERSHELL', previous);
  }
});

test('Office E2E driver rejects add-in activation until a concrete activator is provided', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-empty-session-'));
  const serverPath = join(dir, 'empty-session-server.mjs');
  writeFileSync(serverPath, `
import { createServer } from 'node:http';
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 1, result: {} }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 2, result: { structuredContent: { sessions: [] } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  return firstStdoutLine(server).then((line) => {
    try {
      const { endpoint } = JSON.parse(line);
      const result = runDriver({ host: 'Word', step: 'waitForSession', context: { daemon: { endpoint }, document: { path: 'missing.docx' }, timeoutMs: 1 } });
      assert.notEqual(result.status, 0);
      assert.match(result.stderr, /Open MCP Control in Word/);
    } finally {
      server.kill();
    }
  });
});

function runDriver(payload) {
  return spawnSync(process.execPath, [DRIVER], {
    input: JSON.stringify(payload),
    encoding: 'utf8',
    timeout: 90000
  });
}

function restoreEnv(name, value) {
  if (value === undefined) delete process.env[name];
  else process.env[name] = value;
}

function firstStdoutLine(child) {
  return new Promise((resolve, reject) => {
    let buffer = '';
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => {
      buffer += chunk;
      const index = buffer.indexOf('\n');
      if (index !== -1) resolve(buffer.slice(0, index));
    });
    child.stderr.on('data', (chunk) => {
      if (String(chunk).trim()) reject(new Error(String(chunk)));
    });
    child.on('error', reject);
    child.on('exit', (code) => {
      if (code !== null && code !== 0) reject(new Error(`server exited ${code}`));
    });
  });
}
