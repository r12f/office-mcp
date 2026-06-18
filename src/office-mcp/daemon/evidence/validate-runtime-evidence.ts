import { existsSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

type GateStatus = 'passed' | 'failed' | 'skipped' | 'blocked_by_runtime';

type EvidenceGate = {
  name: string;
  status: GateStatus;
  details?: Record<string, unknown>;
};

type EvidenceReport = {
  schema_version: number;
  generated_at: string;
  kind?: string;
  endpoint: string;
  session_id?: string;
  gates: EvidenceGate[];
};

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const evidencePath = resolve(readOption('--input') ?? join(repoRoot, 'artifacts/runtime-evidence.json'));
const validateUi = hasFlag('--ui');
const requireManualTray = hasFlag('--require-manual-tray');
const manualTrayEvidencePath = readOption('--manual-tray-evidence-path') ?? process.env.OFFICE_MCP_TRAY_MANUAL_EVIDENCE_PATH;
const requireProductVisual = hasFlag('--require-product-visual');
const productVisualEvidencePath = readOption('--product-visual-evidence-path') ?? process.env.OFFICE_MCP_PRODUCT_VISUAL_EVIDENCE_PATH;
const requireIrm = hasFlag('--require-irm');
const requireFullWordSmoke = hasFlag('--require-full-word-smoke');
const requireExcelSmoke = hasFlag('--require-excel-smoke');
const requireComTrackedChanges = hasFlag('--require-com-tracked-changes');
const requireIrmPreflight = hasFlag('--require-irm-preflight');
const requireClaudeDesktopInstallation = hasFlag('--require-claude-desktop-installation');
const requireAgentClientPrompt = hasFlag('--require-agent-client-prompt');
const requireMutation = hasFlag('--require-mutation');
const report = JSON.parse(readFileSync(evidencePath, 'utf8')) as EvidenceReport;
const requiresWordBaseline = !requireExcelSmoke ||
  requireIrm ||
  requireFullWordSmoke ||
  requireComTrackedChanges ||
  requireIrmPreflight ||
  requireClaudeDesktopInstallation ||
  requireAgentClientPrompt ||
  requireMutation;

const requiredPassed = [
  'word.session_discovery',
  'word.runtime_read_smoke',
  'agent_client_stdio_bridge'
];

const fullWordSmokeGates = [
  'word.full_smoke.word-core',
  'word.full_smoke.word-formatting',
  'word.full_smoke.word-review',
  'word.full_smoke.word-resources',
  'word.full_smoke.word-spec-args'
];

const comTrackedChangeGates = [
  'word.tracked_change_com.accept',
  'word.tracked_change_com.reject'
];

const failures: string[] = [];

if (report.schema_version !== 1) failures.push(`Unsupported schema_version: ${report.schema_version}`);

if (validateUi) {
  validateUiEvidence();
  emitSummary();
}

if (requiresWordBaseline && !report.session_id) failures.push('Missing session_id.');

if (requiresWordBaseline) {
  for (const name of requiredPassed) {
    requirePassedGate(name);
  }
} else {
  requirePassedGate('word.session_discovery');
}

if (requireFullWordSmoke) {
  for (const name of fullWordSmokeGates) requirePassedGate(name);
}

if (requireExcelSmoke) requirePassedGate('excel.runtime_smoke');

if (requireMutation) requirePassedGate('word.runtime_mutation_smoke');

if (requireComTrackedChanges) {
  for (const name of comTrackedChangeGates) requirePassedGate(name);
}

if (requireIrmPreflight) requirePassedGate('irm_document_preflight');
if (requireClaudeDesktopInstallation) {
  const gate = requirePassedGate('claude_desktop_installation');
  if (gate && gate.details?.ui_validation_ready !== true) {
    failures.push('Claude Desktop installation gate is not UI-validation ready.');
  }
}

if (requireAgentClientPrompt) requirePassedGate('agent_client_prompt');

const irmGate = gateByName('irm_rights_matrix');
if (requireIrm) {
  if (!irmGate) {
    failures.push('Missing required IRM gate: irm_rights_matrix');
  } else if (irmGate.status !== 'passed') {
    failures.push(`IRM gate is ${irmGate.status}, expected passed.`);
  }
}

const failedGate = report.gates.find((gate) => gate.status === 'failed' || gate.status === 'blocked_by_runtime');
if (failedGate) failures.push(`Gate ${failedGate.name} is ${failedGate.status}.`);

const summary = {
  ok: failures.length === 0,
  require_irm: requireIrm,
  require_full_word_smoke: requireFullWordSmoke,
  require_excel_smoke: requireExcelSmoke,
  require_com_tracked_changes: requireComTrackedChanges,
  require_irm_preflight: requireIrmPreflight,
  require_claude_desktop_installation: requireClaudeDesktopInstallation,
  require_agent_client_prompt: requireAgentClientPrompt,
  require_mutation: requireMutation,
  require_manual_tray: requireManualTray,
  generated_at: report.generated_at,
  endpoint: report.endpoint,
  session_id: report.session_id,
  gates: Object.fromEntries(report.gates.map((gate) => [gate.name, gate.status])),
  failures
};

console.log(JSON.stringify(summary, null, 2));
if (failures.length > 0) process.exit(1);

function validateUiEvidence(): never {
  if (report.kind !== 'ui_runtime_evidence') failures.push(`Unsupported UI evidence kind: ${report.kind ?? 'missing'}`);
  for (const name of [
    'ui.daemon_runtime_file',
    'ui.state_api_origin_redaction',
    'ui.events_stream',
    'ui.tray_probe',
    'ui.production_daemon_tray',
    'ui.browser_smoke'
  ]) {
    requirePassedGate(name);
  }
  if (requireManualTray) validateManualTrayEvidence();
  if (requireProductVisual) validateProductVisualEvidence();
  emitSummary();
}

function validateManualTrayEvidence(): void {
  if (!manualTrayEvidencePath) {
    failures.push('Missing --manual-tray-evidence-path for required manual tray evidence.');
    return;
  }
  const manual = JSON.parse(readFileSync(resolve(manualTrayEvidencePath), 'utf8')) as Record<string, unknown>;
  if (manual.schema_version !== 1) failures.push(`Unsupported manual tray schema_version: ${manual.schema_version}`);
  if (manual.kind !== 'tray_manual_evidence') failures.push(`Unsupported manual tray evidence kind: ${manual.kind ?? 'missing'}`);
  if (manual.platform !== 'win32') failures.push(`Manual tray evidence platform is ${manual.platform}, expected win32.`);
  const observedMenuItems = Array.isArray(manual.observed_menu_items) ? manual.observed_menu_items.filter((item): item is string => typeof item === 'string') : [];
  validateTrayMenuLabels(observedMenuItems, 'Manual tray evidence');
  if (typeof manual.observed_tooltip !== 'string' || !trayTooltipLooksProductReady(manual.observed_tooltip)) {
    failures.push('Manual tray evidence missing product tray tooltip.');
  }
  if (typeof manual.screenshot_path !== 'string' || !screenshotFileLooksLikeImage(resolve(manual.screenshot_path))) {
    failures.push('Manual tray evidence screenshot file does not exist.');
  }
  validateManualTrayDaemonContext(manual.daemon_context);
  for (const [key, label] of [
    ['visible_icon', 'visible tray icon'],
    ['right_click_menu', 'right-click menu'],
    ['show_ui_opened', 'Show Office MCP opened UI'],
    ['passed', 'manual tray evidence passed']
  ] as const) {
    if (manual[key] !== true) failures.push(`Manual tray evidence missing ${label}.`);
  }
}

function validateProductVisualEvidence(): void {
  if (!productVisualEvidencePath) {
    failures.push('Missing --product-visual-evidence-path for required product visual evidence.');
    return;
  }
  const visual = JSON.parse(readFileSync(resolve(productVisualEvidencePath), 'utf8')) as Record<string, unknown>;
  if (visual.schema_version !== 1) failures.push(`Unsupported product visual schema_version: ${visual.schema_version}`);
  if (visual.kind !== 'product_visual_evidence') failures.push(`Unsupported product visual evidence kind: ${visual.kind ?? 'missing'}`);
  if (visual.platform !== 'win32') failures.push(`Product visual evidence platform is ${visual.platform}, expected win32.`);
  if (visual.product_name !== 'Office MCP Control') failures.push('Product visual evidence missing Office MCP Control product name.');
  for (const [key, label] of [
    ['product_text_ready', 'product text on all surfaces'],
    ['catalog_type_ready', 'catalog type metadata'],
    ['catalog_icon_visible', 'catalog icon'],
    ['tray_tooltip_ready', 'tray product tooltip'],
    ['tray_icon_visible', 'visible tray icon'],
    ['tray_menu_native', 'native tray menu'],
    ['quit_confirmation_visible', 'quit confirmation dialog'],
    ['passed', 'product visual evidence passed']
  ] as const) {
    if (visual[key] !== true) failures.push(`Product visual evidence missing ${label}.`);
  }
  if (typeof visual.tray_tooltip !== 'string' || !trayTooltipLooksProductReady(visual.tray_tooltip)) {
    failures.push('Product visual evidence missing product tray tooltip text.');
  }
  if (typeof visual.catalog_type !== 'string' || !/local productivity automation control utility/i.test(visual.catalog_type)) {
    failures.push('Product visual evidence missing local productivity automation/control type metadata.');
  }
  validateProductVisualScreenshots(visual.screenshot_paths);
  validateProductVisualObservations(visual.observations);
  validateExcelTaskpaneVisualEvidence(visual.excel_taskpane);
}

function validateProductVisualScreenshots(paths: unknown): void {
  if (!isRecord(paths)) {
    failures.push('Product visual evidence screenshot paths are malformed.');
    return;
  }
  for (const surface of productVisualSurfaces()) {
    const path = paths[surface];
    if (typeof path !== 'string' || !screenshotFileLooksLikeImage(resolve(path))) {
      failures.push(`Product visual evidence screenshot missing or invalid: ${surface}.`);
    }
  }
}

function validateProductVisualObservations(observations: unknown): void {
  if (!isRecord(observations)) {
    failures.push('Product visual evidence observations are malformed.');
    return;
  }
  for (const surface of productVisualSurfaces().filter((item) => item !== 'tray_tooltip')) {
    const value = observations[surface];
    if (typeof value !== 'string' || !value.includes('Office MCP Control')) {
      failures.push(`Product visual evidence observation missing product name: ${surface}.`);
    }
  }
}

function validateExcelTaskpaneVisualEvidence(taskpane: unknown): void {
  if (!isRecord(taskpane)) {
    failures.push('Product visual evidence missing Excel task pane details.');
    return;
  }
  for (const [key, label] of [
    ['compact_top_block', 'compact top block'],
    ['tools_permissions_merged', 'merged tools and permissions surface'],
    ['inline_settings', 'inline settings']
  ] as const) {
    if (taskpane[key] !== true) failures.push(`Product visual evidence missing Excel ${label}.`);
  }
  if (typeof taskpane.server_protocol_row !== 'string' || !/^Server .+ \/ Protocol .+$/.test(taskpane.server_protocol_row)) {
    failures.push('Product visual evidence missing Excel combined server/protocol row.');
  }
  if (typeof taskpane.document_state !== 'string' || !/^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(taskpane.document_state) || /unknown/i.test(taskpane.document_state)) {
    failures.push('Product visual evidence missing concrete Excel editable/read-only/protected state.');
  }
  if (taskpane.density_ready !== true) failures.push('Product visual evidence missing Excel task pane density pass flag.');
}

function productVisualSurfaces(): string[] {
  return [
    'word_ribbon_command',
    'word_catalog_entry',
    'word_taskpane_title',
    'excel_ribbon_command',
    'excel_catalog_entry',
    'excel_taskpane_title',
    'tray_icon',
    'tray_native_menu',
    'tray_tooltip',
    'tray_quit_confirmation'
  ];
}

function validateManualTrayDaemonContext(context: unknown): void {
  if (context === undefined) return;
  if (!isRecord(context)) {
    failures.push('Manual tray daemon context is malformed.');
    return;
  }
  const status = context.status;
  if (!isRecord(status) || status.ok !== true || status.running !== true || typeof status.uiUrl !== 'string') {
    failures.push('Manual tray daemon context missing running daemon status.');
  }
  const trayProbe = context.tray_probe;
  if (!isRecord(trayProbe) || trayProbe.ok !== true || trayProbe.native_host !== true) {
    failures.push('Manual tray daemon context missing native tray probe success.');
    return;
  }
  const snapshot = trayProbe.snapshot;
  const menuItems = isRecord(snapshot) && Array.isArray(snapshot.menu_items)
    ? snapshot.menu_items.filter((item): item is string => typeof item === 'string')
    : [];
  validateTrayMenuLabels(menuItems, 'Manual tray daemon context live');
  validateStructuredTraySnapshot(snapshot, 'Manual tray daemon context live');
}


function validateTrayMenuLabels(menuItems: string[], label: string): void {
  for (const expected of ['Status:', 'Clients:', 'Documents:', 'Show Office MCP', 'Quit Office MCP']) {
    if (!menuItems.some((item) => item.includes(expected))) {
      failures.push(`${label} missing menu item: ${expected}`);
    }
  }
}

function validateStructuredTraySnapshot(snapshot: unknown, label: string): void {
  if (!isRecord(snapshot)) {
    failures.push(`${label} tray snapshot is malformed.`);
    return;
  }
  if (typeof snapshot.tooltip !== 'string' || !trayTooltipLooksProductReady(snapshot.tooltip)) {
    failures.push(`${label} tray snapshot missing product tooltip.`);
  }
  const menu = Array.isArray(snapshot.menu) ? snapshot.menu : [];
  const expected = [
    { kind: 'read_only', enabled: false, label: /^Status: (Up|Degraded|Down)$/ },
    { kind: 'read_only', enabled: false, label: /^Clients: \d+$/ },
    { kind: 'read_only', enabled: false, label: /^Documents: \d+$/ },
    { kind: 'separator', enabled: false, label: /^---$/ },
    { kind: 'action', enabled: true, label: /^Show Office MCP$/, action: 'show_ui' },
    { kind: 'action', enabled: true, label: /^Quit Office MCP$/, action: 'quit' }
  ];
  if (menu.length !== expected.length) {
    failures.push(`${label} tray menu has ${menu.length} structured items, expected ${expected.length}.`);
    return;
  }
  expected.forEach((rule, index) => {
    const item = menu[index];
    if (!isRecord(item)) {
      failures.push(`${label} tray menu item ${index} is malformed.`);
      return;
    }
    if (item.kind !== rule.kind) failures.push(`${label} tray menu item ${index} has kind ${String(item.kind)}, expected ${rule.kind}.`);
    if (item.enabled !== rule.enabled) failures.push(`${label} tray menu item ${index} has enabled ${String(item.enabled)}, expected ${rule.enabled}.`);
    if (typeof item.label !== 'string' || !rule.label.test(item.label)) failures.push(`${label} tray menu item ${index} has wrong label.`);
    if ('action' in rule && item.action !== rule.action) failures.push(`${label} tray menu item ${index} has action ${String(item.action)}, expected ${rule.action}.`);
  });
}

function trayTooltipLooksProductReady(value: string): boolean {
  return /^Office MCP - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(value);
}
function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function screenshotFileLooksLikeImage(path: string): boolean {
  if (!existsSync(path) || !statSync(path).isFile()) return false;
  const header = readFileSync(path).subarray(0, 12);
  const isPng = header.length >= 8 && header[0] === 0x89 && header[1] === 0x50 && header[2] === 0x4e && header[3] === 0x47 && header[4] === 0x0d && header[5] === 0x0a && header[6] === 0x1a && header[7] === 0x0a;
  const isJpeg = header.length >= 3 && header[0] === 0xff && header[1] === 0xd8 && header[2] === 0xff;
  const isWebp = header.length >= 12 && header.subarray(0, 4).toString('ascii') === 'RIFF' && header.subarray(8, 12).toString('ascii') === 'WEBP';
  const isBmp = header.length >= 2 && header[0] === 0x42 && header[1] === 0x4d;
  return isPng || isJpeg || isWebp || isBmp;
}

function emitSummary(): never {
  console.log(JSON.stringify({
    ok: failures.length === 0,
    ui: validateUi,
    generated_at: report.generated_at,
    kind: report.kind,
    endpoint: report.endpoint,
    session_id: report.session_id,
    gates: Object.fromEntries(report.gates.map((gate) => [gate.name, gate.status])),
    require_manual_tray: requireManualTray,
    manual_tray_evidence_path: manualTrayEvidencePath ? resolve(manualTrayEvidencePath) : undefined,
    require_product_visual: requireProductVisual,
    product_visual_evidence_path: productVisualEvidencePath ? resolve(productVisualEvidencePath) : undefined,
    failures
  }, null, 2));
  process.exit(failures.length > 0 ? 1 : 0);
}

function gateByName(name: string): EvidenceGate | undefined {
  return report.gates.find((gate) => gate.name === name);
}

function requirePassedGate(name: string): EvidenceGate | undefined {
  const gate = gateByName(name);
  if (!gate) {
    failures.push(`Missing required gate: ${name}`);
  } else if (gate.status !== 'passed') {
    failures.push(`Gate ${name} is ${gate.status}, expected passed.`);
  }
  return gate;
}

function readOption(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}

function hasFlag(name: string): boolean {
  return process.argv.includes(name);
}
