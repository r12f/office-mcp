import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const ADDIN_ROOT = process.cwd();

test('Word task pane opts the current document into Office auto-open after connect', () => {
  const source = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(source, /function isWordHost\(info\)/);
  assert.match(source, /Office\.HostType\?\.Word/);
  assert.match(source, /Office\.AutoShowTaskpaneWithDocument/);
  assert.match(source, /Office\.context\.document\.settings\.saveAsync/);
  assert.match(source, /await enableDocumentAutoOpen\(\)/);
});

test('Word add-in manifest and task pane asset versions stay aligned', () => {
  const manifest = readFileSync(join(ADDIN_ROOT, 'manifest.xml'), 'utf8');
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(manifest, /<Version>1\.0\.0\.7<\/Version>/);
  assert.match(manifest, /taskpane\.html\?v=0\.1\.7/);
  assert.match(html, /taskpane\.css\?v=0\.1\.7/);
  assert.match(html, /common\/browser-ui\.js\?v=0\.1\.7/);
  assert.match(html, /common\/addin-channel\.js\?v=0\.1\.7/);
  assert.match(html, /common\/logger\.js\?v=0\.1\.7/);
  assert.match(html, /common\/task-history\.js\?v=0\.1\.7/);
  assert.match(html, /taskpane\.js\?v=0\.1\.7/);
  assert.match(html, /<script async src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/);
  assert.match(js, /ADDIN_VERSION = '0\.1\.7'/);
});

test('Office catalog registration wrapper delegates to shared Word and Excel catalog script', () => {
  const script = readFileSync(join(ADDIN_ROOT, 'scripts', 'register-word-catalog.ps1'), 'utf8');
  const commonScript = readFileSync(join(ADDIN_ROOT, '..', 'common', 'scripts', 'register-office-catalog.ps1'), 'utf8');

  assert.match(script, /register-office-catalog\.ps1/);
  assert.match(commonScript, /src\\office-ctl\\word\\manifest\.xml/);
  assert.match(commonScript, /src\\office-ctl\\excel\\manifest\.xml/);
  assert.match(commonScript, /Word manifest:/);
  assert.match(commonScript, /Excel manifest:/);
  assert.match(commonScript, /TrustedCatalogs\\office-mcp/);
  assert.match(commonScript, /SkipRegistry/);
  assert.match(commonScript, /Remove-DeveloperDebugRegistration/);
  assert.match(commonScript, /WEF\\\\Developer|WEF\\Developer/);
});

test('Office catalog registration script stages Word and Excel manifests without registry mutation', () => {
  const catalogPath = mkdtempSync(join(tmpdir(), 'office-mcp-catalog-'));
  mkdirSync(join(catalogPath, 'word'), { recursive: true });
  mkdirSync(join(catalogPath, 'excel'), { recursive: true });
  writeFileSync(join(catalogPath, 'word', 'manifest.xml'), '<legacy />');
  writeFileSync(join(catalogPath, 'excel', 'manifest.xml'), '<legacy />');

  try {
    const result = spawnSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        join(ADDIN_ROOT, '..', 'common', 'scripts', 'register-office-catalog.ps1'),
        '-CatalogPath',
        catalogPath,
        '-BaseUrl',
        'https://localhost:8766',
        '-SkipRegistry'
      ],
      { cwd: join(ADDIN_ROOT, '..', '..', '..'), encoding: 'utf8' }
    );

    assert.equal(result.status, 0, result.stderr || result.stdout);
    assert.match(result.stdout, /Registered Office trusted catalog:/);
    assert.match(result.stdout, /Catalog URL: \\\\localhost\\[A-Z]\$\\/);
    assert.match(result.stdout, /Manifest origin: https:\/\/localhost:8766/);
    assert.match(result.stdout, /Word manifest:/);
    assert.match(result.stdout, /Excel manifest:/);
    const wordManifest = readFileSync(join(catalogPath, 'office-mcp-word.xml'), 'utf8');
    const excelManifest = readFileSync(join(catalogPath, 'office-mcp-excel.xml'), 'utf8');
    assert.match(wordManifest, /<OfficeApp/);
    assert.match(wordManifest, /https:\/\/localhost:8766\/taskpane\.html\?v=0\.1\.7/);
    assert.match(excelManifest, /<OfficeApp/);
    assert.match(excelManifest, /https:\/\/localhost:8766\/excel\/taskpane\.html\?v=0\.1\.6/);
    assert.doesNotMatch(wordManifest, /https:\/\/localhost:8765/);
    assert.doesNotMatch(excelManifest, /https:\/\/localhost:8765/);
    assert.throws(() => readFileSync(join(catalogPath, 'word', 'manifest.xml'), 'utf8'));
    assert.throws(() => readFileSync(join(catalogPath, 'excel', 'manifest.xml'), 'utf8'));
  } finally {
    rmSync(catalogPath, { force: true, recursive: true });
  }
});

test('Word task pane exposes product UI regions and accessible endpoint settings', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.css'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(html, /id="connectionBadge"/);
  assert.match(html, /id="serverVersion"/);
  assert.match(html, /id="protocolVersion"/);
  assert.match(html, /id="hostPlatform"/);
  assert.match(html, /id="documentState"/);
  assert.match(html, /id="connectionDetail"/);
  assert.match(html, /id="currentTask"/);
  assert.match(html, /id="historyList"/);
  assert.match(html, /class="panel summary-panel"/);
  assert.match(html, /class="tools-panel"/);
  assert.match(html, /id="toolList"/);
  assert.match(html, /id="toolPermissionList"/);
  assert.match(html, /id="enabledToolCount"/);
  assert.match(html, /Enabled 0 of 0/);
  assert.match(html, /aria-live="polite"/);
  assert.match(html, /type="url" inputmode="url" autocomplete="off" spellcheck="false"/);
  assert.match(html, /aria-label="Open Settings"/);
  assert.ok(html.indexOf('id="settingsPanel"') < html.indexOf('id="currentTaskHeading"'));
  assert.ok(html.indexOf('id="toolList"') < html.indexOf('id="settingsPanel"'));
  assert.match(css, /:focus-visible/);
  assert.match(css, /\.summary-panel/);
  assert.match(css, /\.tool-list/);
  assert.match(css, /\.settings-panel/);
  assert.match(css, /\.tool-permission-row/);
  assert.match(css, /\.tool-toggle/);
  assert.match(css, /prefers-reduced-motion/);
  assert.match(css, /forced-colors: active/);
  assert.match(js, /const TOOL_GROUPS = \[/);
  assert.match(js, /const TOOL_METADATA = new Map\(\[/);
  assert.match(js, /TOOL_PERMISSION_STORAGE_KEY/);
  assert.match(js, /const toolListEl = document\.getElementById\('toolList'\)/);
  assert.match(js, /const toolPermissionListEl = document\.getElementById\('toolPermissionList'\)/);
  assert.match(js, /function renderToolSummary\(\)/);
  assert.match(js, /toolCountEl\.textContent = `Enabled \$\{effective\.length\} of \$\{AVAILABLE_TOOLS\.length\}`/);
  assert.match(js, /\$\{tools\.length\}\/\$\{groupTotal\} enabled/);
  assert.match(js, /renderToolSummary\(\)/);
  assert.match(js, /function renderToolPermissions\(\)/);
  assert.match(js, /enabledToolCountEl\.textContent = `Enabled \$\{enabled\.length\} of \$\{AVAILABLE_TOOLS\.length\}`/);
  assert.match(js, /function effectiveTools\(\)/);
  assert.match(js, /available_tools: effectiveTools\(\)/);
  assert.match(js, /sessionUpdatedNotification\(\{/);
  assert.match(js, /patch: \{ available_tools: effectiveTools\(\) \}/);
  assert.match(js, /TOOL_DISABLED_BY_USER/);
  assert.match(js, /new TaskHistoryStore\(\{ redactText \}\)/);
  assert.match(js, /const \{ history, historyLimit \} = taskStore\.snapshot\(\)/);
  assert.match(js, /historyCountEl\.textContent = `\$\{history\.length\} \/ \$\{historyLimit\}`/);
  assert.match(js, /function taskMarkup\(task\)/);
  assert.match(js, /userIntent/);
  assert.match(js, /const requestId = message\.params\?\.request_id \|\| String\(message\.id\)/);
  assert.match(js, /startTask\(requestId, tool, message\.params \|\| \{\}, message\.params\.timeout_ms\)/);
  assert.match(js, /taskStore\.isCancelled\(requestId\)/);
  assert.match(js, /taskStore\.consumeCancellation\(requestId\)/);
  assert.match(js, /finishTask\(requestId, 'success'/);
  assert.match(js, /Cancel requested/);
  assert.match(js, /Deadline/);
  assert.match(js, /office_mcp_code/);
  assert.match(js, /Retriable:/);
  assert.match(js, /Partial effect:/);
  assert.match(js, /storeEndpointOverride\(value\)/);
  assert.match(js, /settingsToggleEl\.addEventListener\('click', handleSettingsClick\)/);
  assert.match(js, /settingsToggleEl\.addEventListener\('keydown', activateSettingsWithKeyboard\)/);
  assert.match(js, /suppressNextSettingsClick/);
  assert.match(js, /function activateSettingsWithKeyboard\(event\)/);
  assert.match(js, /event\.key !== 'Enter' && event\.key !== ' '/);
  assert.match(js, /registerResult\(message, PROTOCOL_VERSION\)/);
  assert.match(js, /protocol_version/);
  assert.match(js, /beforeunload/);
  assert.match(js, /Cancel requested/);
  assert.match(js, /Deadline/);
  assert.match(js, /OfficeCtlCommon/);
  assert.match(js, /OfficeCtlAddinChannel/);
  assert.match(js, /clearEndpointOverride/);
  assert.match(js, /currentOriginEndpoint/);
  assert.match(js, /function tryCurrentOriginEndpointFallback\(failedEndpoint\)/);
  assert.match(js, /OfficeCtlLogger/);
  assert.match(js, /OfficeCtlTaskHistory/);
  assert.match(js, /window\.__OFFICE_MCP_TASKPANE_READY__ = true/);
  assert.match(js, /function whenOfficeReady\(callback\)/);
  assert.match(js, /!window\.Office \|\| typeof Office\.onReady !== 'function'/);
  assert.match(js, /registerRequest\(requestId/);
  assert.match(js, /sessionAddedNotification\(\{/);
  assert.doesNotMatch(js, /function redactText/);
  assert.doesNotMatch(js, /function escapeHtml/);
  assert.doesNotMatch(js, /function configuredEndpoint/);
  assert.doesNotMatch(js, /const taskHistory = \[\]/);
  assert.doesNotMatch(js, /function sanitizeError/);
  assert.doesNotMatch(js, /console\.(log|warn|error)/);
  assert.doesNotMatch(js, /method: 'register'/);
  assert.doesNotMatch(js, /method: 'session\.added'/);
  assert.doesNotMatch(js, /message\.method === 'tool\.invoke'/);
  assert.doesNotMatch(js, /server_version/);
  assert.doesNotMatch(js, /localStorage\.setItem\('office-mcp\.addin-endpoint'/);
  assert.doesNotMatch(js, /sessionStorage\.setItem\('office-mcp\.instance-id'/);
});

test('Word task pane announces session only after successful register response', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const registerBody = functionBody(js, 'register');
  const responseBody = functionBody(js, 'handleRegisterResponse');

  assert.doesNotMatch(registerBody, /announceSession\(/);
  assert.match(responseBody, /serverInfo = registerResult\(message, PROTOCOL_VERSION\)/);
  assert.match(responseBody, /announceSession\(\)\.catch/);
  assert.match(responseBody, /session\.announce\.failed/);
});

function functionBody(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `missing function ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(open + 1, index);
  }
  assert.fail(`unterminated function ${name}`);
}
