import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const ADDIN_ROOT = process.cwd();

test('Excel add-in manifest targets workbook host and versioned task pane', () => {
  const manifest = readFileSync(join(ADDIN_ROOT, 'manifest.xml'), 'utf8');
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(manifest, /<Host Name="Workbook" \/>/);
  assert.match(manifest, /<Set Name="ExcelApi" MinVersion="1\.1" \/>/);
  assert.match(manifest, /excel\/taskpane\.html\?v=0\.1\.7/);
  assert.match(html, /excel\/taskpane\.css\?v=0\.1\.7/);
  assert.match(html, /common\/addin-channel\.js\?v=0\.1\.7/);
  assert.match(html, /excel\/taskpane\.js\?v=0\.1\.7/);
  assert.match(html, /<script async src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/);
  assert.match(js, /ADDIN_VERSION = '0\.1\.7'/);
});

test('Excel task pane uses common channel and registers Excel runtime metadata', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.css'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(html, /id="connectionBadge"/);
  assert.match(html, /id="serverVersion"/);
  assert.match(html, /id="protocolVersion"/);
  assert.match(html, /id="hostPlatform"/);
  assert.match(html, /id="historyList"/);
  assert.match(html, /class="panel summary-panel"/);
  assert.match(html, /class="tools-panel"/);
  assert.match(html, /id="toolList"/);
  assert.match(html, /id="toolPermissionList"/);
  assert.match(html, /id="enabledToolCount"/);
  assert.match(html, /Enabled 0 of 0/);
  assert.match(html, /type="url" inputmode="url" autocomplete="off" spellcheck="false"/);
  assert.match(html, /aria-label="Open Settings"/);
  assert.match(html, /Connecting…/);
  assert.match(html, /wss:\/\/localhost:8765\/addin…/);
  assert.doesNotMatch(html, /Connecting\.\.\./);
  assert.doesNotMatch(html, /addin\.\.\./);
  assert.ok(html.indexOf('id="settingsPanel"') < html.indexOf('id="currentTaskHeading"'));
  assert.ok(html.indexOf('id="toolList"') < html.indexOf('id="settingsPanel"'));
  assert.match(css, /--excel: #217346/);
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
  assert.match(js, /function isExcelHost\(info\)/);
  assert.match(js, /Office\.HostType\?\.Excel/);
  assert.match(js, /Office\.context\?\.diagnostics\?\.host/);
  assert.match(js, /window\.Excel\?\.run/);
  assert.match(js, /isSetSupported\?\.\('ExcelApi', '1\.1'\)/);
  assert.match(js, /app: 'excel'/);
  assert.match(js, /clearEndpointOverride/);
  assert.match(js, /currentOriginEndpoint/);
  assert.match(js, /function tryCurrentOriginEndpointFallback\(failedEndpoint\)/);
  assert.match(js, /new TaskHistoryStore\(\{ redactText \}\)/);
  assert.match(js, /const \{ history, historyLimit \} = taskStore\.snapshot\(\)/);
  assert.match(js, /historyCountEl\.textContent = `\$\{history\.length\} \/ \$\{historyLimit\}`/);
  assert.match(js, /function taskMarkup\(task\)/);
  assert.match(js, /const requestId = message\.params\?\.request_id \|\| String\(message\.id\)/);
  assert.match(js, /taskStore\.start\(requestId, tool, message\.params \|\| \{\}, message\.params\?\.timeout_ms\)/);
  assert.match(js, /taskStore\.isCancelled\(requestId\)/);
  assert.match(js, /taskStore\.consumeCancellation\(requestId\)/);
  assert.match(js, /function cancelledError\(tool\)/);
  assert.match(js, /userIntent/);
  assert.match(js, /Cancel requested/);
  assert.match(js, /Deadline/);
  assert.match(js, /office_mcp_code/);
  assert.match(js, /Retriable:/);
  assert.match(js, /Partial effect:/);
  assert.match(js, /mapped\.office_mcp_code === 'CANCELLED' \? 'cancelled' : 'failure'/);
  assert.match(js, /Discard unsaved endpoint changes\?/);
  assert.match(js, /Connecting…/);
  assert.match(js, /Reconnecting…/);
  assert.match(js, /Registering…/);
  assert.doesNotMatch(js, /Connecting\.\.\./);
  assert.doesNotMatch(js, /Reconnecting\.\.\./);
  assert.doesNotMatch(js, /Registering\.\.\./);
  assert.match(js, /settingsToggleEl\.setAttribute\('aria-label', opening \? 'Close Settings' : 'Open Settings'\)/);
  assert.match(js, /settingsToggleEl\.addEventListener\('click', handleSettingsClick\)/);
  assert.match(js, /settingsToggleEl\.addEventListener\('keydown', activateSettingsWithKeyboard\)/);
  assert.match(js, /suppressNextSettingsClick/);
  assert.match(js, /function activateSettingsWithKeyboard\(event\)/);
  assert.match(js, /event\.key !== 'Enter' && event\.key !== ' '/);
  assert.match(js, /endpointInputEl\.focus\(\)/);
  assert.match(js, /saveEndpointEl\.disabled = true/);
  assert.match(js, /saveEndpointEl\.textContent = 'Saving…'/);
  assert.match(js, /registerRequest\(requestId/);
  assert.match(js, /Office\.AutoShowTaskpaneWithDocument/);
  assert.match(js, /function enableAutoOpen\(\)/);
  assert.match(js, /sessionAddedNotification\(\{/);
  assert.match(js, /excel\.read_range/);
  assert.match(js, /excel\.write_range/);
  assert.match(js, /excel\.add_sheet/);
  assert.match(js, /excel\.set_formula/);
  assert.match(js, /excel\.format_range/);
  assert.match(js, /excel\.create_table/);
  assert.match(js, /excel\.create_chart/);
  assert.match(js, /async function readRange/);
  assert.match(js, /async function writeRange/);
  assert.match(js, /async function addSheet/);
  assert.match(js, /async function setFormula/);
  assert.match(js, /async function formatRange/);
  assert.match(js, /async function createTable/);
  assert.match(js, /async function createChart/);
  assert.match(js, /worksheet\.getRange\(address\)/);
  assert.match(js, /range\.values = args\.values/);
  assert.match(js, /context\.workbook\.worksheets\.add/);
  assert.match(js, /range\.formulas = matrixFromScalar/);
  assert.match(js, /range\.format\.font\.bold/);
  assert.match(js, /range\.numberFormat = matrixFromScalar/);
  assert.match(js, /context\.workbook\.tables\.add/);
  assert.match(js, /worksheet\.charts\.add/);
  assert.match(js, /function chartTypeFrom/);
  assert.match(js, /function requiredString/);
  assert.match(js, /INVALID_ARGUMENT/);
  assert.match(js, /Excel\.run/);
  assert.match(js, /OfficeCtlCommon/);
  assert.match(js, /OfficeCtlAddinChannel/);
  assert.match(js, /OfficeCtlLogger/);
  assert.match(js, /OfficeCtlTaskHistory/);
  assert.match(js, /window\.__OFFICE_MCP_TASKPANE_READY__ = true/);
  assert.match(js, /function whenOfficeReady\(callback\)/);
  assert.match(js, /!window\.Office \|\| typeof Office\.onReady !== 'function'/);
  assert.doesNotMatch(js, /const \{ completed \} = taskStore\.snapshot\(\)/);
  assert.doesNotMatch(js, /console\.(log|warn|error)/);
  assert.doesNotMatch(js, /method: 'register'/);
  assert.doesNotMatch(js, /method: 'session\.added'/);
});


test('Excel task pane keeps settings inline and compact at narrow widths', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.css'), 'utf8');

  assert.match(css, /body \{[\s\S]*min-width: 320px;[\s\S]*overflow-x: hidden;/);
  assert.match(css, /\.taskpane-shell \{[\s\S]*align-content: start;[\s\S]*gap: 10px;[\s\S]*padding: 10px;/);
  assert.match(css, /\.summary-panel \{[\s\S]*display: grid;[\s\S]*gap: 10px;/);
  assert.match(css, /\.empty-state \{[\s\S]*padding: 10px;/);
  assert.match(css, /@media \(min-width: 380px\)/);
  assert.doesNotMatch(css, /\b(min-)?height:\s*(1[2-9]\d|[2-9]\d{2,})px/);
  assert.doesNotMatch(cssRule(css, '.summary-panel'), /\bheight:/);
  assert.doesNotMatch(cssRule(css, '.current-task-panel'), /\bheight:/);
  assert.doesNotMatch(cssRule(css, '.history-panel'), /\bheight:/);
  assert.doesNotMatch(css, /overflow-x:\s*(auto|scroll)/);

  const summaryStart = html.indexOf('class="panel summary-panel"');
  const settingsIndex = html.indexOf('id="settingsPanel"');
  const summaryEnd = html.indexOf('</section>', settingsIndex);
  const currentTaskIndex = html.indexOf('id="currentTaskHeading"');
  assert.ok(summaryStart !== -1 && settingsIndex !== -1, 'summary and settings exist');
  assert.ok(settingsIndex > summaryStart, 'settings panel is inside summary flow');
  assert.ok(settingsIndex < currentTaskIndex, 'settings appears before current task');
  assert.ok(summaryEnd < currentTaskIndex, 'summary closes before current task');
});
test('Excel task pane announces session only after successful register response', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const registerBody = functionBody(js, 'register');
  const responseBody = functionBody(js, 'handleRegisterResponse');

  assert.doesNotMatch(registerBody, /announceSession\(/);
  assert.match(responseBody, /serverInfo = registerResult\(message, PROTOCOL_VERSION\)/);
  assert.match(responseBody, /enableAutoOpen\(\)\.then\(\(\) => announceSession\(\)\)\.catch/);
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

function cssRule(source, selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = source.match(new RegExp(`${escaped}\\s*\\{([\\s\\S]*?)\\}`));
  return match?.[1] || '';
}
