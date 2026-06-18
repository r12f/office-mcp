import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { screenshotFileLooksLikeImage } from './image-evidence.js';

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
const requirePowerPointSmoke = hasFlag('--require-powerpoint-smoke');
const requireComTrackedChanges = hasFlag('--require-com-tracked-changes');
const requireIrmPreflight = hasFlag('--require-irm-preflight');
const requireClaudeDesktopInstallation = hasFlag('--require-claude-desktop-installation');
const requireAgentClientPrompt = hasFlag('--require-agent-client-prompt');
const requireMutation = hasFlag('--require-mutation');
const report = JSON.parse(readFileSync(evidencePath, 'utf8')) as EvidenceReport;
const requiresWordBaseline = !(requireExcelSmoke || requirePowerPointSmoke) ||
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
if (requirePowerPointSmoke) {
  const gate = requirePassedGate('powerpoint.runtime_smoke');
  validatePowerPointSmokeGate(gate);
}

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
  require_powerpoint_smoke: requirePowerPointSmoke,
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


function validatePowerPointSmokeGate(gate: EvidenceGate | undefined): void {
  if (!gate || gate.status !== 'passed') return;
  const details = gate.details;
  if (!isRecord(details)) {
    failures.push('PowerPoint smoke gate missing details.');
    return;
  }
  if (typeof details.session_id !== 'string' || details.session_id.length === 0) failures.push('PowerPoint smoke gate missing session_id.');
  if (typeof details.available_tool_count !== 'number' || details.available_tool_count < 5) failures.push('PowerPoint smoke gate missing available tool count.');
  if (details.mutation_proved !== true) failures.push('PowerPoint smoke gate did not prove a mutation path.');
  if (!isRecord(details.add_slide) || typeof details.add_slide.slide_id !== 'string') failures.push('PowerPoint smoke gate missing add_slide proof.');
  if (!isRecord(details.replace_text) || Number(details.replace_text.replacements ?? 0) < 1) failures.push('PowerPoint smoke gate missing replace_text proof.');
  if (!isRecord(details.layout) || typeof details.layout.slide_id !== 'string') failures.push('PowerPoint smoke gate missing apply_layout proof.');
  const pdfSupported = details.pdf_supported === true && details.pdf_mime_type === 'application/pdf' && typeof details.pdf_size === 'number';
  const pdfHostRejection = details.pdf_host_rejection === true;
  if (!pdfSupported && !pdfHostRejection) failures.push('PowerPoint smoke gate missing PDF export success or explicit host-capability rejection.');
}
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
  if (manual.daemon_context_ready !== true) {
    failures.push('Manual tray evidence daemon context is not recorder-ready.');
  }
  validateManualTrayDaemonContext(manual.daemon_context);
  for (const [key, label] of [
    ['visible_icon', 'visible tray icon'],
    ['right_click_menu', 'right-click menu'],
    ['menu_opened_from_tray_icon', 'right-click menu opened from the notification-area tray icon'],
    ['native_menu_appearance_reviewed', 'native tray menu appearance review'],
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
  validateEmbeddedManualTrayEvidence(visual.manual_tray_evidence, visual.manual_tray_evidence_ready);
  if (typeof visual.catalog_type !== 'string' || !/local productivity automation control utility/i.test(visual.catalog_type)) {
    failures.push('Product visual evidence missing local productivity automation/control type metadata.');
  }
  validateProductVisualScreenshots(visual.screenshot_paths);
  validateDistinctProductVisualScreenshots(visual.screenshot_paths);
  validateProductVisualObservations(visual.observations);
  validateProductIdentityReview(visual.product_identity_review);
  validateRenderedLogoReview(visual.rendered_logo_review, visual.rendered_logo_review_ready);
  validateFirstRunIdentity(visual.first_run_identity);
  validateExcelTaskpaneVisualEvidence(visual.excel_taskpane);
  if (visual.daemon_context_ready !== true) {
    failures.push('Product visual evidence daemon context is not recorder-ready.');
  }
  validateProductVisualDaemonContext(visual.daemon_context);
}

function validateEmbeddedManualTrayEvidence(manual: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing embedded manual tray evidence ready flag.');
  if (!isRecord(manual)) {
    failures.push('Product visual evidence missing embedded manual tray evidence.');
    return;
  }
  if (manual.ok !== true) failures.push('Embedded manual tray evidence was not read successfully.');
  if (manual.schema_version !== 1) failures.push(`Unsupported embedded manual tray schema_version: ${manual.schema_version}`);
  if (manual.kind !== 'tray_manual_evidence') failures.push(`Unsupported embedded manual tray evidence kind: ${manual.kind ?? 'missing'}`);
  if (manual.platform !== 'win32') failures.push(`Embedded manual tray evidence platform is ${manual.platform}, expected win32.`);
  const observedMenuItems = Array.isArray(manual.observed_menu_items) ? manual.observed_menu_items.filter((item): item is string => typeof item === 'string') : [];
  validateTrayMenuLabels(observedMenuItems, 'Embedded manual tray evidence');
  if (typeof manual.observed_tooltip !== 'string' || !trayTooltipLooksProductReady(manual.observed_tooltip)) failures.push('Embedded manual tray evidence missing product tray tooltip.');
  if (typeof manual.screenshot_path !== 'string' || !screenshotFileLooksLikeImage(resolve(manual.screenshot_path))) failures.push('Embedded manual tray evidence screenshot file does not exist.');
  validateEmbeddedManualTraySurfaceScreenshots(manual.tray_surface_screenshot_paths, manual.tray_surface_screenshots_exist);
  if (manual.daemon_context_ready !== true) failures.push('Embedded manual tray evidence daemon context is not recorder-ready.');
  validateManualTrayDaemonContext(manual.daemon_context);
  for (const [key, label] of [
    ['visible_icon', 'visible tray icon'],
    ['right_click_menu', 'right-click menu'],
    ['menu_opened_from_tray_icon', 'right-click menu opened from the notification-area tray icon'],
    ['native_menu_appearance_reviewed', 'native tray menu appearance review'],
    ['show_ui_opened', 'Show Office MCP opened UI'],
    ['passed', 'manual tray evidence passed']
  ] as const) {
    if (manual[key] !== true) failures.push(`Embedded manual tray evidence missing ${label}.`);
  }
}

function validateEmbeddedManualTraySurfaceScreenshots(paths: unknown, exists: unknown): void {
  if (!isRecord(paths)) {
    failures.push('Embedded manual tray evidence missing tray surface screenshot paths.');
    return;
  }
  if (!isRecord(exists)) {
    failures.push('Embedded manual tray evidence missing tray surface screenshot existence flags.');
    return;
  }
  for (const surface of trayVisualSurfaces()) {
    const path = paths[surface];
    if (typeof path !== 'string' || !screenshotFileLooksLikeImage(resolve(path))) {
      failures.push(`Embedded manual tray evidence missing or invalid tray surface screenshot: ${surface}.`);
    }
    if (exists[surface] !== true) failures.push(`Embedded manual tray evidence missing ready flag for tray surface screenshot: ${surface}.`);
  }
}

function validateProductIdentityReview(review: unknown): void {
  if (!isRecord(review)) {
    failures.push('Product visual evidence missing product identity review.');
    return;
  }
  for (const [key, label] of [
    ['logo_quality_reviewed', 'logo quality review'],
    ['rendered_size_logo_reviewed', 'rendered-size logo review'],
    ['rendered_logo_review_ready', 'rendered logo review artifact ready flag'],
    ['addin_identity_reviewed', 'add-in first-run identity review'],
    ['word_first_run_identity_reviewed', 'Word first-run identity review'],
    ['excel_first_run_identity_reviewed', 'Excel first-run identity review'],
    ['powerpoint_first_run_identity_reviewed', 'PowerPoint first-run identity review'],
    ['tray_product_polish_reviewed', 'tray product polish review'],
    ['word_first_run_identity_ready', 'Word first-run identity ready flag'],
    ['excel_first_run_identity_ready', 'Excel first-run identity ready flag'],
    ['powerpoint_first_run_identity_ready', 'PowerPoint first-run identity ready flag'],
    ['ready', 'product identity review ready flag']
  ] as const) {
    if (review[key] !== true) failures.push(`Product visual evidence missing ${label}.`);
  }
}

function validateRenderedLogoReview(review: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing rendered logo review ready flag.');
  if (!isRecord(review)) {
    failures.push('Product visual evidence missing rendered logo review artifact.');
    return;
  }
  if (review.ok !== true) failures.push('Product visual rendered logo review was not read successfully.');
  if (review.schema_version !== 1) failures.push(`Unsupported rendered logo review schema_version: ${review.schema_version}`);
  if (review.kind !== 'rendered_logo_review') failures.push(`Unsupported rendered logo review kind: ${review.kind ?? 'missing'}`);
  if (review.product_name !== 'Office MCP Control') failures.push('Rendered logo review missing Office MCP Control product name.');
  if (review.ready !== true) failures.push('Rendered logo review is not ready.');
  if (typeof review.sheet_path !== 'string' || !screenshotFileLooksLikeImage(resolve(review.sheet_path))) {
    failures.push('Rendered logo review contact sheet is missing or invalid.');
  }
  const surfaces = Array.isArray(review.surfaces) ? review.surfaces.filter(isRecord) : [];
  for (const [key, size] of renderedLogoReviewSurfaces()) {
    const surface = surfaces.find((item) => item.key === key);
    if (!surface) {
      failures.push(`Rendered logo review missing surface: ${key}`);
      continue;
    }
    if (surface.rendered_size_px !== size || surface.width !== size || surface.height !== size) {
      failures.push(`Rendered logo review surface ${key} has wrong rendered size.`);
    }
    if (surface.non_empty !== true) failures.push(`Rendered logo review surface ${key} is empty.`);
    if (surface.palette_ready !== true) failures.push(`Rendered logo review surface ${key} is missing product palette.`);
    if (surface.expected_size_ready !== true) failures.push(`Rendered logo review surface ${key} is not expected-size ready.`);
  }
}

function validateFirstRunIdentity(identity: unknown): void {
  if (!isRecord(identity)) {
    failures.push('Product visual evidence missing first-run identity details.');
    return;
  }
  validateHostFirstRunIdentity(identity.word, 'Word');
  validateHostFirstRunIdentity(identity.excel, 'Excel');
  validateHostFirstRunIdentity(identity.powerpoint, 'PowerPoint');
}

function validateHostFirstRunIdentity(identity: unknown, host: string): void {
  if (!isRecord(identity)) {
    failures.push(`Product visual evidence missing ${host} first-run identity details.`);
    return;
  }
  if (identity.manifest_ready !== true) failures.push(`Product visual evidence missing ${host} manifest-derived identity ready flag.`);
  if (typeof identity.display_name !== 'string' || !identity.display_name.includes('Office MCP Control')) {
    failures.push(`Product visual evidence missing ${host} display name product identity.`);
  }
  if (typeof identity.provider !== 'string' || !identity.provider.includes('Office MCP Control')) {
    failures.push(`Product visual evidence missing ${host} provider product name.`);
  }
  if (typeof identity.description !== 'string' || !/local/i.test(identity.description) || !/(productivity|office)/i.test(identity.description) || !/(automation|control)/i.test(identity.description)) {
    failures.push(`Product visual evidence missing ${host} local productivity automation/control description.`);
  }
  if (typeof identity.type !== 'string' || !/local productivity automation control utility/i.test(identity.type)) {
    failures.push(`Product visual evidence missing ${host} local productivity automation/control type metadata.`);
  }
  if (typeof identity.icon_url !== 'string' || !/\/assets\/icon-32\.png/.test(identity.icon_url)) {
    failures.push(`Product visual evidence missing ${host} first-run icon URL.`);
  }
  if (typeof identity.high_resolution_icon_url !== 'string' || !/\/assets\/icon-80\.png/.test(identity.high_resolution_icon_url)) {
    failures.push(`Product visual evidence missing ${host} first-run high-resolution icon URL.`);
  }
  if (identity.ready !== true) failures.push(`Product visual evidence missing ${host} first-run identity ready flag.`);
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

function validateDistinctProductVisualScreenshots(paths: unknown): void {
  if (!isRecord(paths)) return;
  const seenByPath = new Map<string, string>();
  for (const surface of distinctProductVisualSurfaces()) {
    const path = paths[surface];
    if (typeof path !== 'string') continue;
    const normalized = normalizeScreenshotPath(path);
    const previous = seenByPath.get(normalized);
    if (previous) {
      failures.push(`Product visual evidence reuses one live screenshot for distinct surfaces: ${previous} and ${surface}.`);
      continue;
    }
    seenByPath.set(normalized, surface);
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
  validateExcelRuntimeEvidence(taskpane.runtime_evidence, taskpane.runtime_evidence_ready);
  if (taskpane.density_ready !== true) failures.push('Product visual evidence missing Excel task pane density pass flag.');
}

function validateExcelRuntimeEvidence(evidence: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing Excel runtime evidence ready flag.');
  if (!isRecord(evidence)) {
    failures.push('Product visual evidence missing Excel runtime evidence details.');
    return;
  }
  if (evidence.ok !== true) failures.push('Product visual Excel runtime evidence was not read successfully.');
  if (evidence.schema_version !== 1) failures.push(`Unsupported Excel runtime evidence schema_version: ${evidence.schema_version}`);
  if (evidence.smoke_passed !== true) failures.push('Product visual evidence missing passed Excel runtime smoke gate.');
  if (evidence.ready !== true) failures.push('Product visual Excel runtime evidence is not ready.');
  const session = evidence.session;
  const details = evidence.smoke_details;
  if (!isRecord(session) || !isRecord(details)) {
    failures.push('Product visual Excel runtime evidence missing session or smoke details.');
    return;
  }
  const document = isRecord(session.document) ? session.document : undefined;
  const host = isRecord(session.host) ? session.host : undefined;
  if (session.app !== 'excel' || session.status !== 'active') failures.push('Product visual Excel runtime evidence missing active Excel session.');
  if (typeof session.session_id !== 'string' || details.session_id !== session.session_id) failures.push('Product visual Excel runtime evidence session_id mismatch.');
  if (typeof document?.title !== 'string' || document.title.length === 0) failures.push('Product visual Excel runtime evidence missing workbook title.');
  if (host?.app !== 'excel') failures.push('Product visual Excel runtime evidence missing Excel host metadata.');
  if (typeof session.available_tool_count !== 'number' || session.available_tool_count < 7) failures.push('Product visual Excel runtime evidence missing available tool count.');
  if (details.marker_found !== true) failures.push('Product visual Excel runtime evidence missing marker readback.');
  if (!isRecord(details.write) || details.write.wrote_values !== true) failures.push('Product visual Excel runtime evidence missing write_range proof.');
  if (!isRecord(details.formula) || details.formula.wrote_formula !== true) failures.push('Product visual Excel runtime evidence missing set_formula proof.');
  if (!isRecord(details.format) || details.format.formatted !== true) failures.push('Product visual Excel runtime evidence missing format_range proof.');
  if (!isRecord(details.table) || typeof details.table.table !== 'string') failures.push('Product visual Excel runtime evidence missing create_table proof.');
  if (!isRecord(details.chart) || typeof details.chart.chart !== 'string') failures.push('Product visual Excel runtime evidence missing create_chart proof.');
  if (!isRecord(details.sheet) || details.sheet.activated !== true) failures.push('Product visual Excel runtime evidence missing add_sheet activation proof.');
}

function validateProductVisualDaemonContext(context: unknown): void {
  if (!isRecord(context)) {
    failures.push('Product visual evidence missing daemon context for the local build under test.');
    return;
  }
  const status = context.status;
  if (!isRecord(status) || status.ok !== true || status.running !== true || typeof status.uiUrl !== 'string') {
    failures.push('Product visual evidence daemon context missing running daemon status.');
  }
  const trayProbe = context.tray_probe;
  if (!isRecord(trayProbe) || trayProbe.ok !== true || trayProbe.native_host !== true) {
    failures.push('Product visual evidence daemon context missing native tray probe success.');
    return;
  }
  if (trayProbe.state_fetch_ok !== true) {
    failures.push('Product visual evidence daemon context tray probe did not read live UI state.');
  }
  const snapshot = trayProbe.snapshot;
  const menuItems = isRecord(snapshot) && Array.isArray(snapshot.menu_items)
    ? snapshot.menu_items.filter((item): item is string => typeof item === 'string')
    : [];
  validateTrayMenuLabels(menuItems, 'Product visual daemon context live');
  validateStructuredTraySnapshot(snapshot, 'Product visual daemon context live');
}

function trayVisualSurfaces(): string[] {
  return [
    'tray_icon',
    'tray_native_menu',
    'tray_tooltip',
    'tray_quit_confirmation'
  ];
}

function productVisualSurfaces(): string[] {
  return [
    'word_ribbon_command',
    'word_catalog_entry',
    'word_taskpane_title',
    'excel_ribbon_command',
    'excel_catalog_entry',
    'excel_taskpane_title',
    'powerpoint_ribbon_command',
    'powerpoint_catalog_entry',
    'powerpoint_taskpane_title',
    'logo_tray_size',
    'logo_ribbon_size',
    'logo_catalog_thumbnail',
    'logo_daemon_titlebar',
    'logo_installer_metadata',
    'tray_icon',
    'tray_native_menu',
    'tray_tooltip',
    'tray_quit_confirmation'
  ];
}

function distinctProductVisualSurfaces(): string[] {
  return [
    'word_ribbon_command',
    'word_catalog_entry',
    'word_taskpane_title',
    'excel_ribbon_command',
    'excel_catalog_entry',
    'excel_taskpane_title',
    'powerpoint_ribbon_command',
    'powerpoint_catalog_entry',
    'powerpoint_taskpane_title',
    'tray_icon',
    'tray_native_menu',
    'tray_tooltip',
    'tray_quit_confirmation'
  ];
}

function normalizeScreenshotPath(path: string): string {
  const absolute = resolve(path);
  return process.platform === 'win32' ? absolute.toLowerCase() : absolute;
}

function renderedLogoReviewSurfaces(): Array<[string, number]> {
  return [
    ['logo_tray_size', 16],
    ['logo_ribbon_size', 32],
    ['logo_catalog_thumbnail', 80],
    ['logo_daemon_titlebar', 20],
    ['logo_installer_metadata', 256]
  ];
}

function validateManualTrayDaemonContext(context: unknown): void {
  if (!isRecord(context)) {
    failures.push('Manual tray evidence missing daemon context for the local build under test.');
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
  if (trayProbe.state_fetch_ok !== true) {
    failures.push('Manual tray daemon context tray probe did not read live UI state.');
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
