import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const ADDIN_ROOT = process.cwd();

const POWERPOINT_V1_TOOLS = [
  'powerpoint.get_presentation_info',
  'powerpoint.get_active_view',
  'powerpoint.export_file',
  'powerpoint.update_tags',
  'powerpoint.list_slides',
  'powerpoint.add_slide',
  'powerpoint.update_slide',
  'powerpoint.delete_slide',
  'powerpoint.move_slide',
  'powerpoint.export_slide',
  'powerpoint.list_layouts',
  'powerpoint.apply_layout',
  'powerpoint.get_selection',
  'powerpoint.set_selection',
  'powerpoint.list_shapes',
  'powerpoint.add_text_box',
  'powerpoint.add_shape',
  'powerpoint.insert_image',
  'powerpoint.update_shape',
  'powerpoint.read_text',
  'powerpoint.replace_text',
  'powerpoint.format_text',
  'powerpoint.add_table',
  'powerpoint.read_table',
  'powerpoint.update_table'
];

test('PowerPoint add-in manifest targets presentation host and product identity', () => {
  const manifest = readFileSync(join(ADDIN_ROOT, 'manifest.xml'), 'utf8');
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const packageJson = JSON.parse(readFileSync(join(ADDIN_ROOT, 'package.json'), 'utf8'));

  assert.equal(packageJson.scripts['validate:manifest'], 'office-addin-manifest validate manifest.xml');
  assert.match(packageJson.scripts.check, /^npm run validate:manifest && npm run check:taskpane && npm test$/);
  assert.equal(packageJson.devDependencies['office-addin-manifest'], '2.1.5');

  assert.match(manifest, /<Host Name="Presentation" \/>/);
  assert.match(manifest, /<Set Name="PowerPointApi" MinVersion="1\.1" \/>/);
  assert.match(manifest, /<Version>1\.0\.0\.4<\/Version>/);
  assert.match(manifest, /powerpoint\/taskpane\.html\?v=0\.1\.4/);
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
  assert.match(html, /<title>MCP Control<\/title>/);
  assert.match(html, /<img class="product-mark" src="\/assets\/icon-32\.png" width="32" height="32" alt="" aria-hidden="true" \/>/);
  assert.match(html, /<h1>MCP Control<\/h1>/);
});

test('PowerPoint task pane uses compact shared product UI shell', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.css'), 'utf8');
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  assert.match(html, /powerpoint\/taskpane\.css\?v=0\.1\.4/);
  assert.match(html, /common\/taskpane\.css\?v=0\.1\.4/);
  assert.match(html, /common\/browser-ui\.js\?v=0\.1\.4/);
  assert.match(html, /common\/addin-channel\.js\?v=0\.1\.4/);
  assert.match(html, /common\/logger\.js\?v=0\.1\.4/);
  assert.match(html, /common\/task-history\.js\?v=0\.1\.4/);
  assert.match(html, /common\/main-ui\.js\?v=0\.1\.4/);
  assert.match(html, /powerpoint\/taskpane\.js\?v=0\.1\.4/);
  assert.match(html, /<script src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/);
  assert.doesNotMatch(html, /<script async src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/);
  assert.match(html, /id="runtimeVersions"/);
  assert.match(html, /<dd id="runtimeVersions"><span id="serverVersion">Server Unknown<\/span> \/ <span id="protocolVersion">Protocol 1\.0<\/span><\/dd>/);
  assert.match(html, /Connecting&hellip;/);
  assert.doesNotMatch(html, /Connecting\.\.\./);
  assert.match(html, /<dd id="protection">Not protected<\/dd>/);
  assert.match(html, /<dd id="documentState">Editable<\/dd>/);
  assert.match(html, /class="metadata-copy" data-copy-target="session" aria-label="Copy session ID" title="Copy session ID"/);
  assert.doesNotMatch(html, /class="metadata-copy" data-copy-target="daemon" aria-label="Copy daemon endpoint" title="Copy daemon endpoint"/);
  assert.match(html, /class="panel summary-panel"/);
  assert.match(html, /class="tools-panel"/);
  assert.match(html, /<span>Tools<\/span>/);
  assert.match(html, /id="toolList"/);
  assert.match(html, /0\/25/);
  assert.doesNotMatch(html, /Enabled \d+ of \d+/);
  assert.match(html, /id="toolModeControl" class="tool-mode-control" role="radiogroup" aria-label="Tool capability mode"/);
  assert.match(html, /data-tool-mode="read"/);
  assert.match(html, /data-tool-mode="write"/);
  assert.match(html, /data-tool-mode="all"/);
  assert.doesNotMatch(html, /Tool Permissions/);
  assert.doesNotMatch(html, /id="settingsToggle"/);
  assert.doesNotMatch(html, /id="settingsPanel"/);
  assert.match(html, /type="url" inputmode="url" autocomplete="off" spellcheck="false"/);
  assert.match(html, /placeholder="wss:\/\/localhost:8765\/addin"/);
  assert.match(html, /id="saveEndpoint" class="icon-button reconnect-button" type="submit" aria-label="Reconnect daemon" title="Reconnect daemon"/);
  assert.match(html, /<svg class="control-glyph" viewBox="0 0 24 24" aria-hidden="true" focusable="false">/);
  assert.doesNotMatch(html, /⚙|&#9881;/);
  assert.ok(html.indexOf('class="daemon-endpoint-form"') < html.indexOf('id="runtimeVersions"'));

  assert.match(css, /--powerpoint: #b7472a/);
  const commonCss = readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8');
  assert.match(commonCss, /\.tool-toggle,\r?\n\.group-toggle/);
  assert.match(commonCss, /\.daemon-endpoint-form/);
  assert.match(commonCss, /\.metadata-grid \{[\s\S]*align-items: center;/);
  assert.doesNotMatch(css, /\b(min-)?height:\s*(1[2-9]\d|[2-9]\d{2,})px/);
  assert.doesNotMatch(commonCss, /overflow-x:\s*(auto|scroll)/);
  assert.doesNotMatch(css, /overflow-x:\s*(auto|scroll)/);

  assert.match(js, /ADDIN_VERSION = '0\.1\.4'/);
  assert.match(js, /Connecting\\u2026/);
  assert.match(js, /Reconnecting\\u2026/);
  assert.match(js, /Registering\\u2026/);
  assert.match(js, /saveEndpointEl\.setAttribute\('aria-busy', 'true'\)/);
  assert.match(js, /saveEndpointEl\.removeAttribute\('aria-busy'\)/);
  assert.doesNotMatch(js, /Connecting\.\.\./);
  assert.match(js, /const AVAILABLE_TOOLS = \[/);
  for (const tool of POWERPOINT_V1_TOOLS) assert.match(js, new RegExp(escapeRegExp(tool)));
  assert.doesNotMatch(js, /powerpoint\.export_pdf/);
  assert.match(js, /function isPowerPointHost\(info\)/);
  assert.match(js, /Office\.HostType\?\.PowerPoint/);
  assert.match(js, /Office\.context\?\.requirements\?\.isSetSupported\?\.\('PowerPointApi', '1\.1'\)/);
  assert.match(js, /app: 'powerpoint'/);
  assert.match(js, /supported_features: \['presentation\.session'\]/);
  assert.match(js, /available_tools: effectiveTools\(\)/);
  assert.match(js, /sessionUpdatedNotification\(\{/);
  assert.match(js, /patch: \{ available_tools: effectiveTools\(\) \}/);
  assert.match(js, /TOOL_PERMISSION_STORAGE_KEY/);
  assert.match(js, /TOOL_PERMISSION_MODE_STORAGE_KEY/);
  assert.match(js, /TOOL_DISABLED_BY_USER/);
  assert.match(js, /function effectiveTools\(\)/);
  assert.match(js, /return AVAILABLE_TOOLS\.filter\(\(tool\) => isToolEnabled\(tool\) && isToolAllowedByMode\(tool\)\)/);
  assert.match(js, /function isToolAllowedByMode\(tool\)/);
  assert.match(js, /function handleToolModeChange\(event\)/);
  assert.match(js, /function updateToolPermission\(tool, enabled\)/);
  assert.match(js, /document\.createElement\('details'\)/);
  assert.match(js, /class="group-toggle" type="checkbox" role="switch"/);
  assert.match(js, /toolCountEl\.textContent = `\$\{effective\.length\}\/\$\{AVAILABLE_TOOLS\.length\}`/);
  assert.match(js, /\$\{enabledInGroup\.length\}\/\$\{tools\.length\}/);
  assert.doesNotMatch(js, /Enabled \$\{/);
  assert.match(js, /bindDetailsControl\(input, handleToolPermissionChange\)/);
  assert.match(js, /bindDetailsControl\(input, handleToolGroupPermissionChange\)/);
  assert.match(js, /<div class="tool-permission-list">\$\{rows\}<\/div>/);
  assert.doesNotMatch(js, /<details class="tool-group" open>/);
  assert.match(js, /function handleToolGroupPermissionChange\(event\)/);
  assert.match(js, /const id = `toolPermission-\$\{tool\.replace\(\/\[\^a-z0-9_-\]\/gi, '-'\)\}`/);
  assert.match(js, /for="\$\{id\}"/);
  assert.match(js, /<input id="\$\{id\}" class="tool-toggle"/);
  assert.doesNotMatch(js, /toolListEl\.classList\.toggle\('is-editing-tools'/);
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
  assert.match(js, /function reply\(id, result\) \{\s*return replyJsonRpc\(socket, id, result\);\s*\}/);
  assert.match(js, /taskStore\.isCancelled\(requestId\)/);
  assert.match(js, /taskStore\.consumeCancellation\(requestId\)/);
  assert.match(js, /class="task-meta task-command-id"/);
  assert.match(js, /aria-label="Copy command ID" title="\$\{escapeHtml\(task\.requestId\)\}"/);
  assert.doesNotMatch(js, /settingsToggleEl/);
  assert.match(js, /middleTruncate\(task\.requestId\)/);
  assert.match(js, /document\.addEventListener\('click', handleMetadataCopy\)/);
  assert.match(js, /async function handleMetadataCopy\(event\)/);
  assert.match(js, /event\.target\.closest\('\[data-copy-target\], \[data-copy-value\]'\)/);
  assert.match(js, /button\.title = text === '-' \? button\.getAttribute\('aria-label'\) \|\| '' : text/);
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
  assert.match(js, /ADDIN_VERSION = '0\.1\.4'/);
  assert.doesNotMatch(js, /console\.(log|warn|error)/);
});


test('PowerPoint task pane keeps settings inline and compact at narrow widths', () => {
  const html = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.html'), 'utf8');
  const css = readFileSync(join(ADDIN_ROOT, '..', 'common', 'taskpane.css'), 'utf8');

  assert.match(css, /body \{[\s\S]*min-width: 320px;[\s\S]*overflow-x: hidden;/);
  assert.match(css, /\.taskpane-shell \{[\s\S]*align-content: start;[\s\S]*gap: 10px;[\s\S]*padding: 10px;/);
  assert.match(css, /\.summary-panel \{[\s\S]*display: grid;[\s\S]*gap: 10px;/);
  assert.match(css, /\.empty-state \{[\s\S]*padding: 10px;/);
  assert.match(css, /#documentTitle \{[\s\S]*display: -webkit-box;[\s\S]*-webkit-line-clamp: 2;/);
  assert.doesNotMatch(css, /\b(min-)?height:\s*(1[2-9]\d|[2-9]\d{2,})px/);
  assert.doesNotMatch(cssRule(css, '.summary-panel'), /\bheight:/);
  assert.doesNotMatch(cssRule(css, '.current-task-panel'), /\bheight:/);
  assert.doesNotMatch(cssRule(css, '.history-panel'), /\bheight:/);
  assert.doesNotMatch(css, /overflow-x:\s*(auto|scroll)/);

  const summaryStart = html.indexOf('class="panel summary-panel"');
  const daemonFormIndex = html.indexOf('class="daemon-endpoint-form"');
  const currentTaskIndex = html.indexOf('id="currentTaskHeading"');
  assert.ok(summaryStart !== -1 && daemonFormIndex !== -1, 'summary and daemon form exist');
  assert.ok(daemonFormIndex > summaryStart, 'daemon settings are inline in document metadata');
  assert.ok(daemonFormIndex < currentTaskIndex, 'daemon settings appear before current task');
  assert.doesNotMatch(html, /<section id="settingsPanel"/);
});
test('PowerPoint task pane implements advertised tool handlers with host APIs', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');

  for (const tool of POWERPOINT_V1_TOOLS) {
    assert.match(js, new RegExp(`case '${escapeRegExp(tool)}':`), `missing handler for ${tool}`);
  }
  assert.doesNotMatch(js, /case 'powerpoint\.export_pdf':/);
  assert.match(js, /async function getPresentationInfoTool\(args\)/);
  assert.match(js, /async function getActiveView\(_?args\)/);
  assert.match(js, /async function exportFile\(args\)/);
  assert.match(js, /async function updateTags\(args\)/);
  const updateTagsBody = functionBody(js, 'updateTags');
  assert.match(updateTagsBody, /getItemOrNullObject\(key\)/);
  const tagMetadataBody = functionBody(js, 'tagMetadata');
  assert.match(tagMetadataBody, /normalized_key/);
  assert.match(tagMetadataBody, /host_key/);
  assert.match(js, /async function listSlides\(args\)/);
  assert.match(js, /async function addSlide\(args\)/);
  assert.match(js, /slides\.add\(slideOptions\(args\)\)/);
  assert.match(js, /added\.shapes\.addTextBox\(title/);
  assert.match(js, /async function updateSlide\(args\)/);
  const updateSlideBody = functionBody(js, 'updateSlide');
  assert.match(updateSlideBody, /slide\.load\('id,index,tags\/items\/key,value,layout\/id,layout\/name,shapes\/items\/id'\)/);
  assert.match(updateSlideBody, /slide\.tags\.getItemOrNullObject\(key\)/);
  assert.match(updateSlideBody, /Slide tag updates are not persisted by this PowerPoint host/);
  assert.match(js, /async function deleteSlide\(args\)/);
  assert.match(js, /async function moveSlide\(args\)/);
  assert.match(js, /async function exportSlide\(args\)/);
  assert.match(js, /async function listLayouts\(_?args\)/);
  const listLayoutsBody = functionBody(js, 'listLayouts');
  assert.match(listLayoutsBody, /master\.layouts\.load\('items\/id,name,type'\)/);
  assert.match(js, /async function getSelection\(_?args\)/);
  assert.match(js, /async function setSelection\(args\)/);
  assert.match(js, /async function listShapes\(args\)/);
  const listShapesBody = functionBody(js, 'listShapes');
  assert.match(listShapesBody, /slide\.shapes\.load\('items'\)/);
  assert.match(listShapesBody, /shape\.load\('id,name,type,left,top,width,height,rotation,textFrame\/hasText,textFrame\/textRange\/text'\)/);
  assert.match(js, /async function addTextBox\(args\)/);
  const addTextBoxBody = functionBody(js, 'addTextBox');
  assert.match(addTextBoxBody, /shape\.load\('id,name,type,left,top,width,height,rotation,textFrame\/hasText,textFrame\/textRange\/text'\)/);
  assert.match(js, /async function addShape\(args\)/);
  const addShapeBody = functionBody(js, 'addShape');
  assert.match(addShapeBody, /shape\.load\('id,name,type,left,top,width,height,rotation,textFrame\/hasText,textFrame\/textRange\/text'\)/);
  const shapeMetadataBody = functionBody(js, 'shapeMetadata');
  assert.match(shapeMetadataBody, /safeLoaded\(shape, 'altTextTitle'/);
  assert.match(shapeMetadataBody, /safeLoaded\(shape, 'zOrderPosition'/);
  assert.match(js, /async function updateShape\(args\)/);
  assert.match(js, /async function readText\(args\)/);
  const readTextBody = functionBody(js, 'readText');
  assert.match(readTextBody, /loadSlidesWithShapes\(context, args\)/);
  assert.match(readTextBody, /shape\.load\('id,textFrame\/hasText,textFrame\/textRange\/text'\)/);
  assert.match(js, /async function replaceText\(args\)/);
  const replaceTextBody = functionBody(js, 'replaceText');
  assert.match(replaceTextBody, /loadSlidesWithShapes\(context, args\)/);
  assert.match(replaceTextBody, /shape\.load\('id,textFrame\/hasText,textFrame\/textRange\/text'\)/);
  assert.doesNotMatch(replaceTextBody, /slides\.items/, 'replace_text must not read collection items before sync');
  assert.match(js, /shape\.textFrame\?\.textRange/);
  assert.match(js, /range\.text = nextText/);
  assert.match(replaceTextBody, /isOfficeInvalidArgument\(error\)/);
  assert.match(replaceTextBody, /return await PowerPoint\.run/);
  assert.match(js, /function isOfficeInvalidArgument\(error\)/);
  assert.match(js, /debugInfo\?\.code/);
  assert.match(js, /PowerPoint text replacement is not available in this host\./);
  const loadSlidesWithShapesBody = functionBody(js, 'loadSlidesWithShapes');
  assert.match(loadSlidesWithShapesBody, /slides\.load\('items'\)/);
  assert.match(loadSlidesWithShapesBody, /await context\.sync\(\)/);
  assert.match(loadSlidesWithShapesBody, /slide\.load\('id,index'\)/);
  assert.match(loadSlidesWithShapesBody, /slide\.shapes\.load\('items'\)/);
  assert.match(loadSlidesWithShapesBody, /return slides\.items \|\| \[\]/);
  assert.match(js, /async function formatText\(args\)/);
  assert.match(js, /async function addTable\(args\)/);
  const addTableBody = functionBody(js, 'addTable');
  assert.match(addTableBody, /if \(!shape\.table\) throw hostCapabilityUnavailable/);
  assert.match(js, /async function readTable\(args\)/);
  assert.match(js, /async function updateTable\(args\)/);
  assert.match(js, /async function insertImage\(args\)/);
  assert.match(js, /Office\.context\.document\.setSelectedDataAsync\(base64, imageInsertOptions\(args\), callback\)/);
  assert.match(js, /coercionType: Office\.CoercionType\.Image/);
  assert.match(js, /async function applyLayout\(args\)/);
  assert.match(js, /slide\.applyLayout\(layout\)/);
  assert.match(js, /context\.presentation\.slideMasters/);
  const resolveLayoutBody = functionBody(js, 'resolveLayout');
  assert.match(resolveLayoutBody, /masters\.load\('items\/id,name'\)/);
  assert.match(resolveLayoutBody, /master\.layouts\.load\('items\/id,name,type'\)/);
  assert.doesNotMatch(resolveLayoutBody, /layouts\/items\/id,name,type/);
  assert.doesNotMatch(js, /async function exportPdf\(args\)/);
  assert.match(js, /const fileType = officeFileTypeFrom\(format\)/);
  const exportFileBody = functionBody(js, 'exportFile');
  assert.match(exportFileBody, /isDesktopPowerPointHost\(\)/);
  assert.match(exportFileBody, /PowerPoint desktop file export is not available through Office\.context\.document\.getFileAsync/);
  assert.match(exportFileBody, /POWERPOINT_FILE_EXPORT_TIMEOUT_MS/);
  assert.match(exportFileBody, /Office\.context\.document\.getFileAsync\(fileType/);
  assert.match(exportFileBody, /file\.getSliceAsync\(index, callback\)/);
  assert.match(js, /file\.closeAsync\(callback\)/);
  assert.match(js, /function officeFileTypeFrom\(format\)/);
  assert.match(js, /function isDesktopPowerPointHost\(\)/);
  const isDesktopPowerPointHostBody = functionBody(js, 'isDesktopPowerPointHost');
  assert.match(isDesktopPowerPointHostBody, /Office\.context\?\.platform/);
  assert.match(js, /function officeAsync\(start, options = \{\}\)/);
  assert.match(js, /const timeoutMs = numberOrNull\(options\.timeout_ms\)/);
  assert.match(js, /officeMcpCode: options\.timeout_code \|\| 'HOST_ERROR'/);
  assert.match(js, /function requireRequirementSet\(name, version, feature\)/);
  assert.match(js, /function requiredString\(args, key, message\)/);
  assert.match(js, /INVALID_ARGUMENT/);
  assert.match(js, /NOT_FOUND/);
  assert.match(js, /HOST_CAPABILITY_UNAVAILABLE/);
  assert.match(js, /PowerPoint\.run/);
  assert.doesNotMatch(js, /declared by the daemon contract but is not implemented/);
});

test('PowerPoint presentation info is a non-blocking metadata probe', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const source = functionBody(js, 'getPresentationInfoTool');

  assert.doesNotMatch(source, /PowerPoint\.run/, 'presentation info must not block the task pane on PowerPoint.run');
  assert.doesNotMatch(source, /context\.sync\(/, 'presentation info must not depend on host sync');
  assert.match(source, /getPresentationInfo\(\)/);
  assert.match(source, /slide_count:\s*null/);
  assert.match(source, /include_selection:\s*Boolean\(args\.include_selection\)/);
});

test('PowerPoint owner update tools cover the spec action surface', () => {
  const js = readFileSync(join(ADDIN_ROOT, 'public', 'taskpane.js'), 'utf8');
  const updateShapeBody = functionBody(js, 'updateShape');
  const applyShapePropertiesBody = functionBody(js, 'applyShapeProperties');
  const updateTableBody = functionBody(js, 'updateTable');
  const applyTableValuesBody = functionBody(js, 'applyTableValues');
  const indexedItemsBody = functionBody(js, 'indexedItems');

  for (const action of ['set_properties', 'bring_forward', 'bring_to_front', 'send_backward', 'send_to_back', 'group', 'ungroup', 'delete']) {
    assert.match(updateShapeBody, new RegExp(action), `update_shape must support ${action}`);
  }
  assert.match(updateShapeBody, /slide\.shapes\.addGroup\(args\.shape_ids\)/);
  assert.match(updateShapeBody, /shape\.group\.ungroup\(\)/);
  assert.match(updateShapeBody, /textFrame\/hasText,textFrame\/textRange\/text/);
  for (const property of ['left', 'top', 'width', 'height', 'rotation', 'alt_text_title', 'alt_text_description', 'is_decorative', 'visible']) {
    assert.match(applyShapePropertiesBody, new RegExp(property), `update_shape must own ${property}`);
  }
  assert.match(applyShapePropertiesBody, /shape\.fill\.clear\(\)/);
  assert.match(applyShapePropertiesBody, /shape\.fill\.setSolidColor\(fillColor\)/);
  assert.match(applyShapePropertiesBody, /shape\.fill\.transparency = fillTransparency/);
  assert.match(applyShapePropertiesBody, /shape\.lineFormat\.color = lineColor/);
  assert.match(applyShapePropertiesBody, /shape\.lineFormat\.weight = lineWeight/);
  assert.match(applyShapePropertiesBody, /shape\.lineFormat\.dashStyle = lineDashStyle/);
  assert.match(applyShapePropertiesBody, /shape\.lineFormat\.transparency = lineTransparency/);
  assert.match(applyShapePropertiesBody, /shape\.lineFormat\.visible = Boolean\(args\.line_visible\)/);
  assert.match(applyShapePropertiesBody, /requireRequirementSet\('PowerPointApi', '1\.10', 'shape accessibility and visibility updates'\)/);

  for (const action of ['set_values', 'set_cell', 'add_rows', 'delete_rows', 'add_columns', 'delete_columns', 'merge_cells', 'split_cell', 'clear', 'style', 'delete']) {
    assert.match(updateTableBody, new RegExp(action), `update_table must support ${action}`);
  }
  assert.match(applyTableValuesBody, /tableCell\(table, row, column\)/);
  assert.match(updateTableBody, /requireRequirementSet\('PowerPointApi', '1\.9', 'table structural updates'\)/);
  assert.match(updateTableBody, /table\.rows\.add\(/);
  assert.match(indexedItemsBody, /collection\.getItemAt\(Number\(value\)\)/);
  assert.match(updateTableBody, /table\.columns\.add\(/);
  assert.match(updateTableBody, /table\.columns\.deleteColumns\(indexedItems\(table\.columns/);
  assert.match(updateTableBody, /table\.mergeCells\(/);
  assert.match(updateTableBody, /tableCell\(table, rowIndex, columnIndex\)\.split\(/);
  assert.match(updateTableBody, /table\.clear\(tableClearOptions\(args\)\)/);
  assert.match(updateTableBody, /applyTableStyle\(table, args\)/);
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

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
