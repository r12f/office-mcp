import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const ADDIN_ROOT = process.cwd();

test('PowerPoint add-in manifest targets presentation host and product identity', () => {
  const manifest = readFileSync(join(ADDIN_ROOT, 'manifest.xml'), 'utf8');
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const packageJson = JSON.parse(readFileSync(join(ADDIN_ROOT, 'package.json'), 'utf8'));

  assert.equal(packageJson.scripts['validate:manifest'], 'office-addin-manifest validate manifest.xml');
  assert.match(packageJson.scripts.check, /^npm run validate:manifest && npm run check:taskpane && npm test$/);
  assert.equal(packageJson.devDependencies['office-addin-manifest'], '2.1.5');

  assert.match(manifest, /<Host Name="Presentation" \/>/);
  assert.match(manifest, /<Set Name="PowerPointApi" MinVersion="1\.1" \/>/);
  assert.match(manifest, /powerpoint\/taskpane\.html\?v=0\.1\.0/);
  assert.match(manifest, /<ProviderName>Office MCP Control<\/ProviderName>/);
  assert.match(manifest, /<DisplayName DefaultValue="Office MCP Control" \/>/);
  assert.match(manifest, /Control live PowerPoint presentations through a local productivity automation control utility\./);
  assert.match(manifest, /<bt:String id="OfficeMcp\.GroupLabel" DefaultValue="Office MCP Control" \/>/);
  assert.match(manifest, /<bt:String id="OfficeMcp\.OpenPane\.Label" DefaultValue="Open Control Panel" \/>/);
  assert.match(manifest, /Office MCP Control for this presentation/);
  assert.match(manifest, /https:\/\/localhost:8765\/assets\/icon-32\.png/);
  assert.match(manifest, /https:\/\/localhost:8765\/assets\/icon-80\.png/);
  assert.doesNotMatch(manifest, /DefaultValue="office-mcp(?: for PowerPoint)?"/);
  assert.doesNotMatch(manifest, /DefaultValue="Open"/);
  assert.match(html, /<title>Office MCP Control<\/title>/);
  assert.match(html, /<img class="product-mark" src="\/assets\/icon-32\.png" width="32" height="32" alt="" aria-hidden="true" \/>/);
  assert.match(html, /<h1>Office MCP Control<\/h1>/);
});

test('PowerPoint task pane uses compact shared product UI shell', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.css'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(html, /powerpoint\/taskpane\.css\?v=0\.1\.0/);
  assert.match(html, /common\/browser-ui\.js\?v=0\.1\.0/);
  assert.match(html, /common\/addin-channel\.js\?v=0\.1\.0/);
  assert.match(html, /common\/logger\.js\?v=0\.1\.0/);
  assert.match(html, /common\/task-history\.js\?v=0\.1\.0/);
  assert.match(html, /powerpoint\/taskpane\.js\?v=0\.1\.0/);
  assert.match(html, /id="runtimeVersions"/);
  assert.match(html, /<dd id="runtimeVersions"><span id="serverVersion">Server Unknown<\/span> \/ <span id="protocolVersion">Protocol 1\.0<\/span><\/dd>/);
  assert.match(html, /Connecting&hellip;/);
  assert.doesNotMatch(html, /Connecting\.\.\./);
  assert.match(html, /<dd id="protection">Not protected<\/dd>/);
  assert.match(html, /<dd id="documentState">Editable<\/dd>/);
  assert.match(html, /class="metadata-copy" data-copy-target="session" aria-label="Copy session ID"/);
  assert.match(html, /class="metadata-copy" data-copy-target="daemon" aria-label="Copy daemon endpoint"/);
  assert.match(html, /class="panel summary-panel"/);
  assert.match(html, /class="tools-panel"/);
  assert.match(html, /<span>Tools<\/span>/);
  assert.match(html, /id="toolList"/);
  assert.match(html, /Enabled 0 of 5/);
  assert.doesNotMatch(html, /Tool Permissions/);
  assert.match(html, /type="url" inputmode="url" autocomplete="off" spellcheck="false"/);
  assert.match(html, /placeholder="wss:\/\/localhost:8765\/addin…"/);
  assert.doesNotMatch(html, /placeholder="wss:\/\/localhost:8765\/addin"/);
  assert.match(html, /aria-label="Open Settings" title="Open Settings"/);
  assert.match(html, /<svg class="control-glyph" viewBox="0 0 24 24" aria-hidden="true" focusable="false">/);
  assert.match(html, /<circle cx="17" cy="12" r="2" \/>/);
  assert.doesNotMatch(html, /⚙|&#9881;/);
  assert.ok(html.indexOf('id="settingsPanel"') < html.indexOf('id="currentTaskHeading"'));
  assert.ok(html.indexOf('id="toolList"') < html.indexOf('id="settingsPanel"'));

  assert.match(css, /--powerpoint: #b7472a/);
  assert.match(css, /body \{[\s\S]*min-width: 320px;[\s\S]*overflow-x: hidden;/);
  assert.match(css, /\.taskpane-shell \{[\s\S]*align-content: start;[\s\S]*gap: 10px;[\s\S]*padding: 10px;/);
  assert.match(css, /\.summary-panel \{[\s\S]*display: grid;[\s\S]*gap: 10px;/);
  assert.match(css, /\.control-glyph \{[\s\S]*width: 18px;[\s\S]*stroke: currentColor;/);
  assert.match(css, /\.tool-permission-row\.is-mutating \{[\s\S]*border-left: 3px solid var\(--powerpoint\);/);
  assert.match(css, /\.metadata-copy \{[\s\S]*display: inline-flex;[\s\S]*min-height: 32px;[\s\S]*text-align: left;/);
  assert.match(css, /\.metadata-copy code \{[\s\S]*text-overflow: ellipsis;[\s\S]*white-space: nowrap;/);
  assert.match(css, /\.task-command-id \{[\s\S]*display: flex;[\s\S]*gap: 4px;/);
  assert.match(css, /\.inline-copy \{[\s\S]*display: inline-flex;[\s\S]*min-height: 32px;[\s\S]*background: transparent;[\s\S]*cursor: pointer;/);
  assert.match(css, /\.metadata-copy:hover,[\s\S]*\.inline-copy:focus-visible \{[\s\S]*background: var\(--surface-raised\);/);
  assert.doesNotMatch(css, /\b(min-)?height:\s*(1[2-9]\d|[2-9]\d{2,})px/);
  assert.doesNotMatch(cssRule(css, '.summary-panel'), /\bheight:/);
  assert.doesNotMatch(css, /overflow-x:\s*(auto|scroll)/);

  assert.match(js, /ADDIN_VERSION = '0\.1\.0'/);
  assert.match(js, /Connecting\\u2026/);
  assert.match(js, /Reconnecting\\u2026/);
  assert.match(js, /Registering\\u2026/);
  assert.match(js, /saveEndpointEl\.textContent = 'Saving\\u2026'/);
  assert.doesNotMatch(js, /Connecting\.\.\./);
  assert.match(js, /const AVAILABLE_TOOLS = \[/);
  assert.match(js, /powerpoint\.add_slide/);
  assert.match(js, /powerpoint\.replace_text/);
  assert.match(js, /powerpoint\.insert_image/);
  assert.match(js, /powerpoint\.apply_layout/);
  assert.match(js, /powerpoint\.export_pdf/);
  assert.match(js, /function isPowerPointHost\(info\)/);
  assert.match(js, /Office\.HostType\?\.PowerPoint/);
  assert.match(js, /Office\.context\?\.requirements\?\.isSetSupported\?\.\('PowerPointApi', '1\.1'\)/);
  assert.match(js, /app: 'powerpoint'/);
  assert.match(js, /supported_features: \['presentation\.session'\]/);
  assert.match(js, /available_tools: effectiveTools\(\)/);
  assert.match(js, /sessionUpdatedNotification\(\{/);
  assert.match(js, /patch: \{ available_tools: effectiveTools\(\) \}/);
  assert.match(js, /TOOL_PERMISSION_STORAGE_KEY/);
  assert.match(js, /TOOL_DISABLED_BY_USER/);
  assert.match(js, /function effectiveTools\(\)/);
  assert.match(js, /function updateToolPermission\(tool, enabled\)/);
  assert.match(js, /<details class="tool-group"><summary class="tool-group-title">/);
  assert.match(js, /<div class="tool-permission-list">\$\{rows\}<\/div>/);
  assert.doesNotMatch(js, /<details class="tool-group" open>/);
  assert.match(js, /rowStateClass = `\$\{enabled \? '' : ' is-disabled'\}\$\{meta\.sideEffect === 'mutating' \? ' is-mutating' : ''\}`/);
  assert.match(js, /const id = `toolPermission-\$\{tool\.replace\(\/\[\^a-z0-9_-\]\/gi, '-'\)\}`/);
  assert.match(js, /for="\$\{id\}"/);
  assert.match(js, /<input id="\$\{id\}" class="tool-toggle"/);
  assert.match(js, /toolListEl\.classList\.toggle\('is-editing-tools', opening\)/);
  assert.match(js, /sessionAddedNotification\(\{/);
  assert.match(js, /new TaskHistoryStore\(\{ redactText \}\)/);
  assert.match(js, /const \{ history, historyLimit \} = taskStore\.snapshot\(\)/);
  assert.match(js, /historyCountEl\.textContent = `\$\{history\.length\} \/ \$\{historyLimit\}`/);
  assert.match(js, /function taskMarkup\(task\)/);
  assert.match(js, /const tone = task\.status === 'success' \? 'status-success' : task\.status === 'running' \? 'status-warning' : task\.status === 'cancelled' \? 'status-neutral' : 'status-danger'/);
  assert.match(js, /Retriable: \$\{valueLabel\(task\.error\.retriable\)\}/);
  assert.match(js, /Partial effect: \$\{escapeHtml\(task\.error\.partial_effect \|\| 'unknown'\)\}/);
  assert.match(js, /const intent = task\.userIntent \? `<div class="task-meta">\$\{escapeHtml\(redactText\(task\.userIntent\)\)\}<\/div>` : ''/);
  assert.match(js, /const deadline = task\.deadlineAt \? `<div class="task-meta">Deadline \$\{escapeHtml\(formatTime\(task\.deadlineAt\)\)\}<\/div>` : ''/);
  assert.match(js, /const cancel = task\.cancelRequested \? '<div class="task-meta">Cancel requested<\/div>' : ''/);
  assert.match(js, /function valueLabel\(value\)/);
  assert.match(js, /if \(value === true\) return 'yes'/);
  assert.match(js, /taskStore\.isCancelled\(requestId\)/);
  assert.match(js, /taskStore\.consumeCancellation\(requestId\)/);
  assert.match(js, /class="task-meta task-command-id"/);
  assert.match(js, /aria-label="Copy command ID"/);
  assert.match(js, /settingsToggleEl\.setAttribute\('title', opening \? 'Close Settings' : 'Open Settings'\)/);
  assert.match(js, /settingsToggleEl\.setAttribute\('title', 'Open Settings'\)/);
  assert.match(js, /middleTruncate\(task\.requestId\)/);
  assert.match(js, /document\.addEventListener\('click', handleMetadataCopy\)/);
  assert.match(js, /async function handleMetadataCopy\(event\)/);
  assert.match(js, /event\.target\.closest\('\[data-copy-target\], \[data-copy-value\]'\)/);
  assert.match(js, /const target = button\.dataset\.copyTarget \? document\.getElementById\(button\.dataset\.copyTarget\) : null/);
  assert.match(js, /const value = button\.dataset\.copyValue \|\| target\?\.textContent\?\.trim\(\)/);
  assert.match(js, /navigator\.clipboard\?\.writeText/);
  assert.match(js, /function setCopyableMetadata\(element, value\)/);
  assert.match(js, /element\.textContent = middleTruncate\(text\)/);
  assert.match(js, /button\.dataset\.copyValue = text/);
  assert.match(js, /function middleTruncate\(value, maxLength = 30\)/);
  assert.match(js, /return `\$\{text\.slice\(0, head\)\}\$\{marker\}\$\{text\.slice\(text\.length - tail\)\}`/);
  assert.match(js, /function fallbackCopy\(value\)/);
  assert.match(js, /clearEndpointOverride/);
  assert.match(js, /currentOriginEndpoint/);
  assert.match(js, /try \{[\s\S]*validateEndpoint\(value\);[\s\S]*storeEndpointOverride\(value\);[\s\S]*\} catch \(error\) \{[\s\S]*endpointErrorEl\.textContent = error\.message \|\| 'Enter a valid wss:\/\/ endpoint\.';[\s\S]*endpointInputEl\.focus\(\);/);
  assert.doesNotMatch(js, /const validation = validateEndpoint\(value\)/);
  assert.doesNotMatch(js, /validation\.ok/);
  assert.match(js, /Office\.AutoShowTaskpaneWithDocument/);
  assert.match(js, /window\.__OFFICE_MCP_TASKPANE_READY__ = true/);
  assert.doesNotMatch(js, /console\.(log|warn|error)/);
});


test('PowerPoint task pane keeps settings inline and compact at narrow widths', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.css'), 'utf8');

  assert.match(css, /body \{[\s\S]*min-width: 320px;[\s\S]*overflow-x: hidden;/);
  assert.match(css, /\.taskpane-shell \{[\s\S]*align-content: start;[\s\S]*gap: 10px;[\s\S]*padding: 10px;/);
  assert.match(css, /\.summary-panel \{[\s\S]*display: grid;[\s\S]*gap: 10px;/);
  assert.match(css, /\.empty-state \{[\s\S]*padding: 10px;/);
  assert.match(css, /#documentTitle \{[\s\S]*display: -webkit-box;[\s\S]*-webkit-line-clamp: 2;/);
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
test('PowerPoint task pane implements advertised tool handlers with host APIs', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(js, /case 'powerpoint\.add_slide':/);
  assert.match(js, /case 'powerpoint\.replace_text':/);
  assert.match(js, /case 'powerpoint\.insert_image':/);
  assert.match(js, /case 'powerpoint\.apply_layout':/);
  assert.match(js, /case 'powerpoint\.export_pdf':/);
  assert.match(js, /async function addSlide\(args\)/);
  assert.match(js, /slides\.add\(slideOptions\(args\)\)/);
  assert.match(js, /added\.shapes\.addTextBox\(title/);
  assert.match(js, /async function replaceText\(args\)/);
  assert.match(js, /shape\.textFrame\?\.textRange/);
  assert.match(js, /range\.text = nextText/);
  assert.match(js, /async function insertImage\(args\)/);
  assert.match(js, /Office\.context\.document\.setSelectedDataAsync\(base64, imageInsertOptions\(args\), callback\)/);
  assert.match(js, /coercionType: Office\.CoercionType\.Image/);
  assert.match(js, /async function applyLayout\(args\)/);
  assert.match(js, /slide\.applyLayout\(layout\)/);
  assert.match(js, /context\.presentation\.slideMasters/);
  assert.match(js, /async function exportPdf\(args\)/);
  assert.match(js, /Office\.context\.document\.getFileAsync\(Office\.FileType\.Pdf/);
  assert.match(js, /file\.getSliceAsync\(index, callback\)/);
  assert.match(js, /file\.closeAsync\(callback\)/);
  assert.match(js, /mime_type: 'application\/pdf'/);
  assert.match(js, /function requiredString\(args, key, message\)/);
  assert.match(js, /INVALID_ARGUMENT/);
  assert.match(js, /NOT_FOUND/);
  assert.match(js, /PowerPoint\.run/);
  assert.doesNotMatch(js, /declared by the daemon contract but is not implemented/);
});

test('PowerPoint task pane announces session only after successful register response', () => {
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
