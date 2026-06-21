import { createHash } from 'node:crypto';
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
const brandAssetRoot = join(repoRoot, 'src/office-ctl/common/assets');
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
const requireOfficeToolE2e = hasFlag('--require-office-tool-e2e');
const wordToolE2eReportPath = readOption('--word-tool-e2e-report-path') ?? process.env.OFFICE_MCP_WORD_TOOL_E2E_REPORT_PATH;
const excelToolE2eReportPath = readOption('--excel-tool-e2e-report-path') ?? process.env.OFFICE_MCP_EXCEL_TOOL_E2E_REPORT_PATH;
const powerPointToolE2eReportPath = readOption('--powerpoint-tool-e2e-report-path') ?? process.env.OFFICE_MCP_POWERPOINT_TOOL_E2E_REPORT_PATH;
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

if (requireExcelSmoke) {
  const gate = requirePassedGate('excel.runtime_smoke');
  validateExcelSmokeGate(gate);
}
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

if (requireOfficeToolE2e) validateOfficeToolE2eReports();

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
  require_office_tool_e2e: requireOfficeToolE2e,
  require_manual_tray: requireManualTray,
  office_tool_e2e_report_paths: requireOfficeToolE2e ? {
    word: wordToolE2eReportPath ? resolve(wordToolE2eReportPath) : undefined,
    excel: excelToolE2eReportPath ? resolve(excelToolE2eReportPath) : undefined,
    powerpoint: powerPointToolE2eReportPath ? resolve(powerPointToolE2eReportPath) : undefined
  } : undefined,
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
  validateExactToolCatalog('PowerPoint smoke gate', details.available_tool_count, details.available_tools, powerPointV1Tools());
  if (details.mutation_proved !== true) failures.push('PowerPoint smoke gate did not prove a mutation path.');
  validatePowerPointCategoryProofs(details.tool_category_proofs, 'PowerPoint smoke gate');
  if (!isRecord(details.presentation_info)) failures.push('PowerPoint smoke gate missing presentation info proof.');
  if (!isRecord(details.active_view)) failures.push('PowerPoint smoke gate missing active view proof.');
  if (!isRecord(details.list_slides)) failures.push('PowerPoint smoke gate missing list_slides proof.');
  if (!isRecord(details.add_slide) || typeof details.add_slide.slide_id !== 'string') failures.push('PowerPoint smoke gate missing add_slide proof.');
  if (!isRecord(details.add_text_box) || !isRecord(details.add_text_box.shape)) failures.push('PowerPoint smoke gate missing add_text_box proof.');
  if (!isRecord(details.list_shapes) || !Array.isArray(details.list_shapes.shapes)) failures.push('PowerPoint smoke gate missing list_shapes proof.');
  if (!isRecord(details.read_text) || !Array.isArray(details.read_text.items)) failures.push('PowerPoint smoke gate missing read_text proof.');
  if (!isRecord(details.replace_text) || Number(details.replace_text.replacements ?? 0) < 1) failures.push('PowerPoint smoke gate missing replace_text proof.');
  if (!isRecord(details.format_text) || details.format_text.formatted !== true) failures.push('PowerPoint smoke gate missing format_text proof.');
  if (!isRecord(details.layout) || typeof details.layout.slide_id !== 'string') failures.push('PowerPoint smoke gate missing apply_layout proof.');
  if (!isRecord(details.list_layouts) || !Array.isArray(details.list_layouts.masters)) failures.push('PowerPoint smoke gate missing list_layouts proof.');
  const tableSupported = details.table_supported === true && isRecord(details.add_table) && typeof details.add_table.shape_id === 'string' && isRecord(details.read_table);
  const tableHostRejection = details.table_host_rejection === true;
  if (!tableSupported && !tableHostRejection) failures.push('PowerPoint smoke gate missing table success or explicit host-capability rejection.');
  const exportSupported = details.export_supported === true && details.export_mime_type === 'application/pdf' && typeof details.export_size === 'number';
  const exportHostRejection = details.export_host_rejection === true;
  if (!exportSupported && !exportHostRejection) failures.push('PowerPoint smoke gate missing export_file success or explicit host-capability rejection.');
}

function validateExcelSmokeGate(gate: EvidenceGate | undefined): void {
  if (!gate || gate.status !== 'passed') return;
  const details = gate.details;
  if (!isRecord(details)) {
    failures.push('Excel smoke gate missing details.');
    return;
  }
  if (typeof details.session_id !== 'string' || details.session_id.length === 0) failures.push('Excel smoke gate missing session_id.');
  validateExactToolCatalog('Excel smoke gate', details.available_tool_count, details.available_tools, excelV1Tools());
  if (details.marker_found !== true) failures.push('Excel smoke gate missing marker readback.');
  if (!isRecord(details.workbook_info)) failures.push('Excel smoke gate missing get_workbook_info proof.');
  if (typeof details.sheet_list_count !== 'number') failures.push('Excel smoke gate missing list_sheets proof.');
  if (!isRecord(details.updated_sheet)) failures.push('Excel smoke gate missing update_sheet proof.');
  if (!isRecord(details.deleted_sheet) || details.deleted_sheet.deleted !== true) failures.push('Excel smoke gate missing delete_sheet proof.');
  if (!isRecord(details.used_range)) failures.push('Excel smoke gate missing get_used_range proof.');
  if (!isRecord(details.find_replace)) failures.push('Excel smoke gate missing find_replace_cells proof.');
  if (!isRecord(details.clear)) failures.push('Excel smoke gate missing clear_range proof.');
  if (!isRecord(details.write) || details.write.wrote_values !== true) failures.push('Excel smoke gate missing write_range proof.');
  if (!isRecord(details.formula) || details.formula.wrote_formula !== true) failures.push('Excel smoke gate missing set_formula proof.');
  if (!isRecord(details.format) || details.format.formatted !== true) failures.push('Excel smoke gate missing format_range proof.');
  if (!isRecord(details.table) || typeof details.table.table !== 'string') failures.push('Excel smoke gate missing create_table proof.');
  if (!isRecord(details.table_update)) failures.push('Excel smoke gate missing update_table proof.');
  if (!isRecord(details.sort) || details.sort.sorted !== true) failures.push('Excel smoke gate missing sort_range proof.');
  if (!isRecord(details.filter) || details.filter.filtered !== true) failures.push('Excel smoke gate missing apply_filter proof.');
  if (!isRecord(details.chart) || typeof details.chart.chart !== 'string') failures.push('Excel smoke gate missing create_chart proof.');
  if (!isRecord(details.chart_update) || details.chart_update.updated !== true) failures.push('Excel smoke gate missing update_chart proof.');
  if (!isRecord(details.pivot_table) || typeof details.pivot_table.pivot_table !== 'string') failures.push('Excel smoke gate missing create_pivot_table proof.');
  if (!isRecord(details.pivot_update) || details.pivot_update.refreshed !== true) failures.push('Excel smoke gate missing update_pivot_table proof.');
  if (!isRecord(details.sheet) || details.sheet.activated !== true) failures.push('Excel smoke gate missing add_sheet activation proof.');
}

function excelV1Tools(): string[] {
  return [
    'excel.get_workbook_info',
    'excel.list_sheets',
    'excel.add_sheet',
    'excel.update_sheet',
    'excel.delete_sheet',
    'excel.get_used_range',
    'excel.read_range',
    'excel.write_range',
    'excel.clear_range',
    'excel.find_replace_cells',
    'excel.set_formula',
    'excel.format_range',
    'excel.sort_range',
    'excel.apply_filter',
    'excel.create_table',
    'excel.update_table',
    'excel.create_chart',
    'excel.update_chart',
    'excel.create_pivot_table',
    'excel.update_pivot_table'
  ];
}

function powerPointV1Tools(): string[] {
  return [
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
}

function wordV1Tools(): string[] {
  return [
    'word.get_text',
    'word.get_outline',
    'word.get_paragraph',
    'word.find_text',
    'word.get_selection',
    'word.insert_paragraph',
    'word.insert_table',
    'word.insert_image',
    'word.insert_page_break',
    'word.insert_list',
    'word.replace_text',
    'word.update_paragraph',
    'word.delete_range',
    'word.apply_formatting',
    'word.apply_style',
    'word.read_table',
    'word.update_table',
    'word.list_content_controls',
    'word.insert_content_control',
    'word.update_content_control',
    'word.delete_content_control',
    'word.add_comment',
    'word.resolve_comment',
    'word.update_tracked_change',
    'word.save'
  ];
}

function validateOfficeToolE2eReports(): void {
  validateOfficeToolE2eReport('Word', wordToolE2eReportPath);
  validateOfficeToolE2eReport('Excel', excelToolE2eReportPath);
  validateOfficeToolE2eReport('PowerPoint', powerPointToolE2eReportPath);
}

function validateOfficeToolE2eReport(host: 'Word' | 'Excel' | 'PowerPoint', path: string | undefined): void {
  if (!path) {
    failures.push(`Missing --${hostToolE2ePathOption(host)} for required Office tool E2E report.`);
    return;
  }
  let e2e: Record<string, unknown>;
  try {
    e2e = JSON.parse(readFileSync(resolve(path), 'utf8')) as Record<string, unknown>;
  } catch (error) {
    failures.push(`${host} Office tool E2E report could not be read: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }

  if (e2e.schema_version !== 1) failures.push(`${host} Office tool E2E report unsupported schema_version: ${e2e.schema_version}`);
  if (e2e.kind !== 'office_tool_e2e_report') failures.push(`${host} Office tool E2E report has unsupported kind: ${e2e.kind ?? 'missing'}`);
  if (e2e.host !== host) failures.push(`${host} Office tool E2E report host is ${String(e2e.host)}, expected ${host}.`);
  if (e2e.passed !== true) failures.push(`${host} Office tool E2E report did not pass.`);
  validateOfficeToolE2eLifecycle(`${host} Office tool E2E report`, e2e.lifecycle_counts);
  validateOfficeToolE2eActivation(`${host} Office tool E2E report`, e2e.addin_activation);

  const advertisedTools = stringArray(e2e.advertised_tools);
  const sessionTools = stringArray(e2e.session_available_tools);
  const executedTools = stringArray(e2e.executed_tools);
  if (advertisedTools.length === 0) failures.push(`${host} Office tool E2E report missing advertised tools.`);
  if (!sameStrings(advertisedTools, sessionTools)) failures.push(`${host} Office tool E2E report session tools do not match advertised tools.`);
  if (!sameStrings(advertisedTools, executedTools)) failures.push(`${host} Office tool E2E report executed tools do not match advertised tools.`);
  validateOfficeToolRuns(`${host} Office tool E2E report`, advertisedTools, e2e.tool_runs);

  if (!isRecord(e2e.daemon) || typeof e2e.daemon.endpoint !== 'string') failures.push(`${host} Office tool E2E report missing daemon endpoint.`);
  if (!isRecord(e2e.document) || typeof e2e.document.path !== 'string') failures.push(`${host} Office tool E2E report missing driver-owned document path.`);
  if (!isRecord(e2e.session) || typeof e2e.session.session_id !== 'string') failures.push(`${host} Office tool E2E report missing session ID.`);
}

function validateOfficeToolE2eActivation(label: string, activation: unknown): void {
  if (!isRecord(activation)) {
    failures.push(`${label} missing add-in activation proof.`);
    return;
  }
  if (activation.activated !== true) failures.push(`${label} add-in activation did not run.`);
  if (typeof activation.skipped === 'string' && activation.skipped.length > 0) failures.push(`${label} add-in activation was skipped: ${activation.skipped}.`);
}

function validateOfficeToolE2eLifecycle(label: string, lifecycle: unknown): void {
  if (!isRecord(lifecycle)) {
    failures.push(`${label} missing lifecycle counts.`);
    return;
  }
  for (const key of ['start_daemon', 'list_tools', 'create_document', 'activate_addin', 'wait_for_session', 'cleanup_document', 'stop_daemon']) {
    if (lifecycle[key] !== 1) failures.push(`${label} lifecycle ${key} is ${String(lifecycle[key])}, expected 1.`);
  }
}

function validateOfficeToolRuns(label: string, advertisedTools: string[], runs: unknown): void {
  if (!Array.isArray(runs)) {
    failures.push(`${label} missing tool runs.`);
    return;
  }
  const runTools = runs.filter(isRecord).map((run) => run.tool).filter((tool): tool is string => typeof tool === 'string');
  if (!sameStrings(advertisedTools, runTools)) failures.push(`${label} tool runs do not match advertised tools.`);
  for (const run of runs.filter(isRecord)) {
    const tool = typeof run.tool === 'string' ? run.tool : '<missing>';
    if (run.passed !== true) failures.push(`${label} tool ${tool} did not pass.`);
    if (typeof run.id !== 'string' || run.id.length === 0) failures.push(`${label} tool ${tool} missing run ID.`);
    if (typeof run.setup_action_count !== 'number' || run.setup_action_count < 1) failures.push(`${label} tool ${tool} missing setup action proof.`);
    const verifier = run.verifier;
    if (!isRecord(verifier) || (verifier.kind !== 'direct-result' && verifier.kind !== 'readback')) {
      failures.push(`${label} tool ${tool} missing verifier kind.`);
    }
    const expectationKeys = isRecord(verifier) && Array.isArray(verifier.expectation_keys) ? verifier.expectation_keys : [];
    if (expectationKeys.length === 0) failures.push(`${label} tool ${tool} missing verifier expectations.`);
  }
}

function hostToolE2ePathOption(host: 'Word' | 'Excel' | 'PowerPoint'): string {
  if (host === 'Word') return 'word-tool-e2e-report-path';
  if (host === 'Excel') return 'excel-tool-e2e-report-path';
  return 'powerpoint-tool-e2e-report-path';
}

function stringArray(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === 'string') : [];
}

function validateExactToolCatalog(label: string, count: unknown, tools: unknown, expected: string[]): void {
  if (count !== expected.length) {
    failures.push(`${label} missing ${expected.length}-tool available tool count.`);
  }
  if (!sameStrings(stringArray(tools), expected)) {
    failures.push(`${label} available tools are not aligned with v1 catalog.`);
  }
}

function sameStrings(left: string[], right: string[]): boolean {
  return JSON.stringify([...left].sort()) === JSON.stringify([...right].sort());
}

function validatePowerPointCategoryProofs(value: unknown, label: string): void {
  const proofs = isRecord(value) ? value : undefined;
  for (const category of ['presentation', 'slides', 'layout', 'shapes', 'text', 'tables']) {
    if (proofs?.[category] !== true) failures.push(`${label} missing ${category} category proof.`);
  }
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
  if (manual.tray_menu_surface_kind !== 'native' || manual.tray_menu_surface_native !== true) failures.push('Manual tray evidence surface is not native.');
  const observedMenuItems = Array.isArray(manual.observed_menu_items) ? manual.observed_menu_items.filter((item): item is string => typeof item === 'string') : [];
  validateTrayMenuLabels(observedMenuItems, 'Manual tray evidence');
  if (typeof manual.observed_tooltip !== 'string' || !trayTooltipLooksProductReady(manual.observed_tooltip)) {
    failures.push('Manual tray evidence missing product tray tooltip.');
  }
  if (typeof manual.screenshot_path !== 'string' || !screenshotFileLooksLikeImage(resolve(manual.screenshot_path))) {
    failures.push('Manual tray evidence screenshot file does not exist.');
  }
  validateManualTraySurfaceScreenshots(manual.tray_surface_screenshot_paths, manual.tray_surface_screenshots_exist, 'Manual tray evidence');
  if (manual.tray_surface_screenshots_ready !== true) failures.push('Manual tray evidence missing tray surface screenshots ready flag.');
  if (manual.tray_surface_screenshots_distinct !== true) failures.push('Manual tray evidence reuses one screenshot for multiple tray surfaces.');
  if (manual.daemon_context_ready !== true) {
    failures.push('Manual tray evidence daemon context is not recorder-ready.');
  }
  validateManualTrayDaemonContext(manual.daemon_context);
  validateManualTrayObservedSnapshotBinding(manual, 'Manual tray evidence');
  for (const [key, label] of [
    ['visible_icon', 'visible tray icon'],
    ['right_click_menu', 'right-click menu'],
    ['menu_opened_from_tray_icon', 'right-click menu opened from the notification-area tray icon'],
    ['native_menu_appearance_reviewed', 'native tray menu appearance review'],
    ['menu_anchored_to_tray_icon', 'native menu anchored to the notification-area tray icon'],
    ['os_native_menu_behavior_reviewed', 'OS-native tray menu spacing, hover, and theme behavior review'],
    ['keyboard_menu_access_reviewed', 'keyboard access for native tray menu actions'],
    ['native_quit_confirmation_reviewed', 'native quit confirmation review'],
    ['native_tray_interaction_ready', 'native tray interaction ready flag'],
    ['tray_menu_surface_native', 'native tray menu surface'],
    ['show_ui_opened', 'Show Office MCP Control opened UI'],
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
  if (!productCatalogTypeLooksReady(visual.catalog_type)) {
    failures.push('Product visual evidence missing mature local productivity automation/control type metadata.');
  }
  if (visual.tray_menu_surface_kind !== 'native' || visual.tray_menu_surface_native !== true) {
    failures.push('Product visual evidence tray menu surface is not native.');
  }
  validateProductVisualTrayBinding(visual);
  validateProductVisualScreenshots(visual.screenshot_paths, visual.screenshots_exist);
  validateProductVisualScreenshotFreshness(visual.screenshot_paths, visual.screenshot_metadata, visual.screenshots_fresh, visual.screenshots_fresh_ready, visual.recorded_at);
  validateDistinctProductVisualScreenshots(visual.screenshot_paths);
  validateProductVisualObservations(visual.observations);
  validateProductIdentityReview(visual.product_identity_review);
  validateCatalogIdentityReview(visual.catalog_identity_review, visual.catalog_identity_review_ready);
  validateRenderedLogoReview(visual.rendered_logo_review, visual.rendered_logo_review_ready);
  validateFirstRunIdentity(visual.first_run_identity);
  validateFirstRunCatalogBinding(visual.first_run_identity, visual.catalog_identity_review);
  validateDaemonMainWindowVisualEvidence(visual.daemon_main_window);
  validateWordRuntimeEvidence(visual.word_runtime_evidence, visual.word_runtime_evidence_ready);
  validateExcelRuntimeEvidence(visual.excel_runtime_evidence, visual.excel_runtime_evidence_ready);
  validatePowerPointRuntimeEvidence(visual.powerpoint_runtime_evidence, visual.powerpoint_runtime_evidence_ready);
  validateProductVisualOfficeToolE2e(visual.office_tool_e2e, visual.office_tool_e2e_ready);
  validateWordTaskpaneVisualEvidence(visual.word_taskpane);
  validateExcelTaskpaneVisualEvidence(visual.excel_taskpane);
  validatePowerPointTaskpaneVisualEvidence(visual.powerpoint_taskpane);
  validateTaskpaneRuntimeBinding('Word', visual.word_runtime_evidence, visual.word_taskpane);
  validateTaskpaneRuntimeBinding('Excel', visual.excel_runtime_evidence, visual.excel_taskpane);
  validateTaskpaneRuntimeBinding('PowerPoint', visual.powerpoint_runtime_evidence, visual.powerpoint_taskpane);
  validateOfficeToolE2eRuntimeBinding('Word', visual.word_runtime_evidence, visual.office_tool_e2e);
  validateOfficeToolE2eRuntimeBinding('Excel', visual.excel_runtime_evidence, visual.office_tool_e2e);
  validateOfficeToolE2eRuntimeBinding('PowerPoint', visual.powerpoint_runtime_evidence, visual.office_tool_e2e);
  if (visual.daemon_context_ready !== true) {
    failures.push('Product visual evidence daemon context is not recorder-ready.');
  }
  validateProductVisualDaemonContext(visual.daemon_context);
}

function validateCatalogIdentityReview(review: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing catalog identity review ready flag.');
  if (!isRecord(review)) {
    failures.push('Product visual evidence missing catalog identity review.');
    return;
  }
  if (review.ok !== true) failures.push('Product visual catalog identity review was not read successfully.');
  if (review.schema_version !== 1) failures.push(`Unsupported catalog identity review schema_version: ${review.schema_version}`);
  if (review.kind !== 'catalog_identity_review') failures.push(`Unsupported catalog identity review kind: ${review.kind ?? 'missing'}`);
  if (review.product_name !== 'Office MCP Control') failures.push('Catalog identity review missing Office MCP Control product name.');
  if (review.ready !== true) failures.push('Catalog identity review is not ready.');
  if (!productCatalogTypeLooksReady(review.catalog_type)) failures.push('Catalog identity review missing mature local productivity automation/control type metadata.');
  if (typeof review.shared_origin !== 'string' || !/^https:\/\/localhost:\d+$/.test(review.shared_origin)) failures.push('Catalog identity review missing shared local daemon origin.');
  if (!isRecord(review.hosts)) {
    failures.push('Catalog identity review missing host details.');
    return;
  }
  validateCatalogIdentityHost(review.hosts.word, 'Word', '/word/taskpane.html');
  validateCatalogIdentityHost(review.hosts.excel, 'Excel', '/excel/taskpane.html');
  validateCatalogIdentityHost(review.hosts.powerpoint, 'PowerPoint', '/powerpoint/taskpane.html');
}

function validateCatalogIdentityHost(host: unknown, label: string, taskpanePath: string): void {
  if (!isRecord(host)) {
    failures.push(`Catalog identity review missing ${label} details.`);
    return;
  }
  if (host.ready !== true) failures.push(`Catalog identity review missing ${label} ready flag.`);
  if (host.display_name !== 'Office MCP Control') failures.push(`Catalog identity review missing ${label} product display name.`);
  if (host.provider !== 'Office MCP Control') failures.push(`Catalog identity review missing ${label} product provider.`);
  if (typeof host.description !== 'string' || !/local productivity automation control utility/i.test(host.description)) failures.push(`Catalog identity review missing ${label} product description.`);
  if (host.group_label !== 'Office MCP Control') failures.push(`Catalog identity review missing ${label} product ribbon group label.`);
  if (host.command_label !== 'Open Control Panel') failures.push(`Catalog identity review missing ${label} Open Control Panel command label.`);
  if (typeof host.taskpane_url !== 'string' || !host.taskpane_url.includes(taskpanePath)) failures.push(`Catalog identity review missing ${label} task pane URL.`);
  if (typeof host.icon_url !== 'string' || !/\/assets\/icon-32\.png$/.test(host.icon_url)) failures.push(`Catalog identity review missing ${label} generated icon URL.`);
  if (typeof host.high_resolution_icon_url !== 'string' || !/\/assets\/icon-80\.png$/.test(host.high_resolution_icon_url)) failures.push(`Catalog identity review missing ${label} generated high-resolution icon URL.`);
}

function validateEmbeddedManualTrayEvidence(manual: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing embedded manual tray evidence ready flag.');
  if (!isRecord(manual)) {
    failures.push('Product visual evidence missing embedded manual tray evidence.');
    return;
  }
  if (manual.ok !== true) failures.push('Embedded manual tray evidence was not read successfully.');
  if (manual.tray_menu_surface_kind !== 'native' || manual.tray_menu_surface_native !== true) failures.push('Embedded manual tray evidence surface is not native.');
  if (manual.schema_version !== 1) failures.push(`Unsupported embedded manual tray schema_version: ${manual.schema_version}`);
  if (manual.kind !== 'tray_manual_evidence') failures.push(`Unsupported embedded manual tray evidence kind: ${manual.kind ?? 'missing'}`);
  if (manual.platform !== 'win32') failures.push(`Embedded manual tray evidence platform is ${manual.platform}, expected win32.`);
  const observedMenuItems = Array.isArray(manual.observed_menu_items) ? manual.observed_menu_items.filter((item): item is string => typeof item === 'string') : [];
  validateTrayMenuLabels(observedMenuItems, 'Embedded manual tray evidence');
  if (typeof manual.observed_tooltip !== 'string' || !trayTooltipLooksProductReady(manual.observed_tooltip)) failures.push('Embedded manual tray evidence missing product tray tooltip.');
  if (typeof manual.screenshot_path !== 'string' || !screenshotFileLooksLikeImage(resolve(manual.screenshot_path))) failures.push('Embedded manual tray evidence screenshot file does not exist.');
  validateManualTraySurfaceScreenshots(manual.tray_surface_screenshot_paths, manual.tray_surface_screenshots_exist, 'Embedded manual tray evidence');
  if (manual.tray_surface_screenshots_ready !== true) failures.push('Embedded manual tray evidence missing tray surface screenshots ready flag.');
  if (manual.tray_surface_screenshots_distinct !== true) failures.push('Embedded manual tray evidence reuses one screenshot for multiple tray surfaces.');
  if (manual.daemon_context_ready !== true) failures.push('Embedded manual tray evidence daemon context is not recorder-ready.');
  validateManualTrayDaemonContext(manual.daemon_context);
  validateManualTrayObservedSnapshotBinding(manual, 'Embedded manual tray evidence');
  for (const [key, label] of [
    ['visible_icon', 'visible tray icon'],
    ['right_click_menu', 'right-click menu'],
    ['menu_opened_from_tray_icon', 'right-click menu opened from the notification-area tray icon'],
    ['native_menu_appearance_reviewed', 'native tray menu appearance review'],
    ['menu_anchored_to_tray_icon', 'native menu anchored to the notification-area tray icon'],
    ['os_native_menu_behavior_reviewed', 'OS-native tray menu spacing, hover, and theme behavior review'],
    ['keyboard_menu_access_reviewed', 'keyboard access for native tray menu actions'],
    ['native_quit_confirmation_reviewed', 'native quit confirmation review'],
    ['native_tray_interaction_ready', 'native tray interaction ready flag'],
    ['tray_menu_surface_native', 'native tray menu surface'],
    ['show_ui_opened', 'Show Office MCP Control opened UI'],
    ['passed', 'manual tray evidence passed']
  ] as const) {
    if (manual[key] !== true) failures.push(`Embedded manual tray evidence missing ${label}.`);
  }
}

function validateProductVisualTrayBinding(visual: Record<string, unknown>): void {
  const manual = isRecord(visual.manual_tray_evidence) ? visual.manual_tray_evidence : undefined;
  if (manual) {
    if (typeof visual.tray_tooltip === 'string' && typeof manual.observed_tooltip === 'string' && visual.tray_tooltip !== manual.observed_tooltip) {
      failures.push('Product visual evidence tray tooltip does not match embedded manual tray evidence.');
    }
    if (visual.tray_menu_surface_kind !== manual.tray_menu_surface_kind || visual.tray_menu_surface_native !== manual.tray_menu_surface_native) {
      failures.push('Product visual evidence tray menu surface metadata does not match embedded manual tray evidence.');
    }
  }

  const context = isRecord(visual.daemon_context) ? visual.daemon_context : undefined;
  const trayProbe = isRecord(context?.tray_probe) ? context.tray_probe : undefined;
  const snapshot = isRecord(trayProbe?.snapshot) ? trayProbe.snapshot : undefined;
  if (snapshot && typeof visual.tray_tooltip === 'string' && typeof snapshot.tooltip === 'string' && visual.tray_tooltip !== snapshot.tooltip) {
    failures.push('Product visual evidence tray tooltip does not match daemon tray snapshot.');
  }
}

function validateManualTraySurfaceScreenshots(paths: unknown, exists: unknown, label: string): void {
  if (!isRecord(paths)) {
    failures.push(`${label} missing tray surface screenshot paths.`);
    return;
  }
  if (!isRecord(exists)) {
    failures.push(`${label} missing tray surface screenshot existence flags.`);
    return;
  }
  const seenByPath = new Map<string, string>();
  for (const surface of trayVisualSurfaces()) {
    const path = paths[surface];
    if (typeof path !== 'string' || !screenshotFileLooksLikeImage(resolve(path))) {
      failures.push(`${label} missing or invalid tray surface screenshot: ${surface}.`);
    } else {
      const normalized = normalizeScreenshotPath(path);
      const previous = seenByPath.get(normalized);
      if (previous) {
        failures.push(`${label} reuses one tray surface screenshot for distinct surfaces: ${previous} and ${surface}.`);
      } else {
        seenByPath.set(normalized, surface);
      }
    }
    if (exists[surface] !== true) failures.push(`${label} missing ready flag for tray surface screenshot: ${surface}.`);
  }
}

function validateProductIdentityReview(review: unknown): void {
  if (!isRecord(review)) {
    failures.push('Product visual evidence missing product identity review.');
    return;
  }
  for (const [key, label] of [
    ['logo_quality_reviewed', 'logo quality review'],
    ['logo_future_office_control_reviewed', 'future office-control logo review'],
    ['final_logo_user_surface_reviewed', 'final logo user-surface review'],
    ['current_logo_screenshot_feedback_reviewed', 'current screenshot logo feedback review'],
    ['rendered_size_logo_reviewed', 'rendered-size logo review'],
    ['rendered_logo_review_ready', 'rendered logo review artifact ready flag'],
    ['catalog_identity_review_ready', 'catalog identity review ready flag'],
    ['addin_identity_reviewed', 'add-in first-run identity review'],
    ['addin_title_icon_type_reviewed', 'add-in title, icon, and type/category review'],
    ['addin_installable_surface_reviewed', 'add-in installable-software surface review'],
    ['current_addin_screenshot_feedback_reviewed', 'current screenshot add-in feedback review'],
    ['word_first_run_identity_reviewed', 'Word first-run identity review'],
    ['excel_first_run_identity_reviewed', 'Excel first-run identity review'],
    ['powerpoint_first_run_identity_reviewed', 'PowerPoint first-run identity review'],
    ['tray_product_polish_reviewed', 'tray product polish review'],
    ['tray_native_first_impression_reviewed', 'tray native first-impression review'],
    ['tray_normal_windows_launch_reviewed', 'tray normal Windows launch review'],
    ['current_tray_screenshot_feedback_reviewed', 'current screenshot tray feedback review'],
    ['current_screenshot_feedback_ready', 'current screenshot feedback ready flag'],
    ['word_first_run_identity_ready', 'Word first-run identity ready flag'],
    ['excel_first_run_identity_ready', 'Excel first-run identity ready flag'],
    ['powerpoint_first_run_identity_ready', 'PowerPoint first-run identity ready flag'],
    ['word_runtime_evidence_ready', 'Word runtime evidence ready flag'],
    ['excel_runtime_evidence_ready', 'Excel runtime evidence ready flag'],
    ['powerpoint_runtime_evidence_ready', 'PowerPoint runtime evidence ready flag'],
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
  validateRenderedLogoDesignReview(review.design_review);
  if (typeof review.sheet_path !== 'string' || !screenshotFileLooksLikeImage(resolve(review.sheet_path))) {
    failures.push('Rendered logo review contact sheet is missing or invalid.');
  }
  validateRenderedLogoAssetFingerprints(review);
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

function validateRenderedLogoAssetFingerprints(review: Record<string, unknown>): void {
  if (typeof review.source_asset_sha256 !== 'string') {
    failures.push('Rendered logo review missing source asset SHA-256 fingerprint.');
  } else if (review.source_asset_sha256 !== sha256File(join(brandAssetRoot, 'brand-mark.svg'))) {
    failures.push('Rendered logo review source asset fingerprint does not match current brand-mark.svg.');
  }
  const surfaces = Array.isArray(review.surfaces) ? review.surfaces.filter(isRecord) : [];
  for (const [key] of renderedLogoReviewSurfaces()) {
    const surface = surfaces.find((item) => item.key === key);
    if (!surface) continue;
    if (typeof surface.asset_sha256 !== 'string') {
      failures.push(`Rendered logo review missing asset SHA-256 fingerprint: ${key}.`);
      continue;
    }
    const assetName = typeof surface.asset_path === 'string' ? surface.asset_path.split(/[\\/]/).pop() : undefined;
    if (!assetName || !/^icon-\d+\.png$/.test(assetName)) {
      failures.push(`Rendered logo review missing generated icon asset path: ${key}.`);
      continue;
    }
    if (surface.asset_sha256 !== sha256File(join(brandAssetRoot, assetName))) {
      failures.push(`Rendered logo review asset fingerprint does not match current generated icon: ${key}.`);
    }
  }
}

function sha256File(path: string): string | undefined {
  try {
    return createHash('sha256').update(readFileSync(path)).digest('hex');
  } catch (_error) {
    return undefined;
  }
}

function validateRenderedLogoDesignReview(review: unknown): void {
  if (!isRecord(review)) {
    failures.push('Rendered logo review missing design review.');
    return;
  }
  if (review.ready !== true) failures.push('Rendered logo design review is not ready.');
  validateRenderedLogoConceptPass(review.concept_pass);
  if (typeof review.future_office_control_brief !== 'string' || !/future office control/i.test(review.future_office_control_brief) || !/routing|operator|control/i.test(review.future_office_control_brief) || !/without .*Office-owned app marks/i.test(review.future_office_control_brief)) {
    failures.push('Rendered logo design review missing future office control brief.');
  }
  if (typeof review.office_productivity_metaphor !== 'string' || !/document|pane|office/i.test(review.office_productivity_metaphor)) failures.push('Rendered logo design review missing office productivity metaphor.');
  if (typeof review.user_control_metaphor !== 'string' || !/control|command|operator/i.test(review.user_control_metaphor)) failures.push('Rendered logo design review missing user control metaphor.');
  if (typeof review.futuristic_maturity !== 'string' || !/mature|futuristic|desktop utility/i.test(review.futuristic_maturity)) failures.push('Rendered logo design review missing mature futuristic utility rationale.');
  if (typeof review.non_microsoft_distinction !== 'string' || !/Office logos/i.test(review.non_microsoft_distinction) || !/Microsoft 365 gradients/i.test(review.non_microsoft_distinction) || !/PowerPoint slide silhouettes/i.test(review.non_microsoft_distinction) || !/Outlook envelope marks/i.test(review.non_microsoft_distinction) || !/gear-only/i.test(review.non_microsoft_distinction)) failures.push('Rendered logo design review missing non-Microsoft distinction.');
  const rejected = Array.isArray(review.rejects_generic_readings) ? review.rejects_generic_readings : [];
  for (const item of ['settings', 'file', 'debug console', 'ai-only', 'microsoft office clone']) {
    if (!rejected.includes(item)) failures.push(`Rendered logo design review does not reject ${item}.`);
  }
}

function validateRenderedLogoConceptPass(conceptPass: unknown): void {
  if (!isRecord(conceptPass)) {
    failures.push('Rendered logo design review missing concept pass.');
    return;
  }
  if (conceptPass.ready !== true) failures.push('Rendered logo concept pass is not ready.');
  if (conceptPass.selected_direction !== 'Command Console Panes') failures.push('Rendered logo concept pass missing selected Command Console Panes direction.');
  if (typeof conceptPass.minimum_concepts_reviewed !== 'number' || conceptPass.minimum_concepts_reviewed < 3) failures.push('Rendered logo concept pass must review at least three concepts.');
  const concepts = Array.isArray(conceptPass.concepts) ? conceptPass.concepts.filter(isRecord) : [];
  if (concepts.length < 3) failures.push('Rendered logo concept pass missing reviewed concepts.');
  if (!concepts.some((concept) => concept.name === 'Command Console Panes' && concept.decision === 'selected' && typeof concept.rationale === 'string' && /office productivity, local routing, and deliberate user control/i.test(concept.rationale))) failures.push('Rendered logo concept pass missing selected office-control rationale.');
  if (!concepts.some((concept) => concept.name === 'Orbiting Document Hub' && concept.decision === 'rejected' && typeof concept.rationale === 'string' && /generic sync or cloud connector/i.test(concept.rationale))) failures.push('Rendered logo concept pass missing rejected hub rationale.');
  if (!concepts.some((concept) => concept.name === 'Shielded Automation Badge' && concept.decision === 'rejected' && typeof concept.rationale === 'string' && /endpoint protection software/i.test(concept.rationale))) failures.push('Rendered logo concept pass missing rejected shield rationale.');
  const rejectedPatterns = Array.isArray(conceptPass.rejected_patterns) ? conceptPass.rejected_patterns : [];
  for (const item of ['gear-only settings mark', 'Office-like app tile', 'host-app color block', 'generic document thumbnail', 'terminal/debug glyph', 'AI sparkle motif']) {
    if (!rejectedPatterns.includes(item)) failures.push(`Rendered logo concept pass does not reject ${item}.`);
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
  if (!productCatalogTypeLooksReady(identity.type)) {
    failures.push(`Product visual evidence missing ${host} mature local productivity automation/control type metadata.`);
  }
  if (typeof identity.icon_url !== 'string' || !/\/assets\/icon-32\.png/.test(identity.icon_url)) {
    failures.push(`Product visual evidence missing ${host} first-run icon URL.`);
  }
  if (typeof identity.high_resolution_icon_url !== 'string' || !/\/assets\/icon-80\.png/.test(identity.high_resolution_icon_url)) {
    failures.push(`Product visual evidence missing ${host} first-run high-resolution icon URL.`);
  }
  if (identity.ready !== true) failures.push(`Product visual evidence missing ${host} first-run identity ready flag.`);
}

function validateFirstRunCatalogBinding(firstRun: unknown, catalogReview: unknown): void {
  if (!isRecord(firstRun) || !isRecord(catalogReview) || !isRecord(catalogReview.hosts)) return;
  validateHostFirstRunCatalogBinding('Word', firstRun.word, catalogReview.hosts.word);
  validateHostFirstRunCatalogBinding('Excel', firstRun.excel, catalogReview.hosts.excel);
  validateHostFirstRunCatalogBinding('PowerPoint', firstRun.powerpoint, catalogReview.hosts.powerpoint);
}

function validateHostFirstRunCatalogBinding(label: 'Word' | 'Excel' | 'PowerPoint', firstRun: unknown, catalogHost: unknown): void {
  if (!isRecord(firstRun) || !isRecord(catalogHost)) return;
  for (const [firstRunField, catalogField, fieldLabel] of [
    ['display_name', 'display_name', 'display name'],
    ['provider', 'provider', 'provider'],
    ['description', 'description', 'description'],
    ['icon_url', 'icon_url', 'icon URL'],
    ['high_resolution_icon_url', 'high_resolution_icon_url', 'high-resolution icon URL']
  ] as const) {
    if (typeof firstRun[firstRunField] !== 'string' || typeof catalogHost[catalogField] !== 'string' || firstRun[firstRunField] !== catalogHost[catalogField]) {
      failures.push(`Product visual evidence ${label} first-run identity ${fieldLabel} does not match catalog identity review.`);
    }
  }
  if (typeof firstRun.taskpane_url === 'string' && typeof catalogHost.taskpane_url === 'string' && firstRun.taskpane_url !== catalogHost.taskpane_url) {
    failures.push(`Product visual evidence ${label} first-run identity task pane URL does not match catalog identity review.`);
  }
}

function validateDaemonMainWindowVisualEvidence(mainWindow: unknown): void {
  if (!isRecord(mainWindow)) {
    failures.push('Product visual evidence missing daemon main window review.');
    return;
  }
  for (const [key, label] of [
    ['reviewed', 'daemon main window product review'],
    ['compact_status_details_reviewed', 'daemon main window compact status/details review'],
    ['three_column_layout_reviewed', 'daemon main window three-column layout review'],
    ['screenshot_exists', 'daemon main window screenshot existence flag'],
    ['screenshot_fresh', 'daemon main window fresh screenshot flag'],
    ['screenshot_ready', 'daemon main window screenshot ready flag'],
    ['ready', 'daemon main window ready flag']
  ] as const) {
    if (mainWindow[key] !== true) failures.push(`Product visual evidence missing ${label}.`);
  }
  if (typeof mainWindow.screenshot_path !== 'string' || !screenshotFileLooksLikeImage(resolve(mainWindow.screenshot_path))) {
    failures.push('Product visual evidence missing or invalid daemon main window screenshot.');
  }
  if (typeof mainWindow.observation !== 'string' || !mainWindow.observation.includes('Office MCP Control')) {
    failures.push('Product visual evidence missing daemon main window product observation.');
  }
}

function productCatalogTypeLooksReady(value: unknown): boolean {
  if (typeof value !== 'string') return false;
  return /local productivity automation control utility/i.test(value) && !/(add-in|task pane|developer tool|mcp server|protocol bridge|sample|debug|experimental|office-mcp-(word|excel|powerpoint))/i.test(value);
}

function validateProductVisualScreenshots(paths: unknown, exists: unknown): void {
  if (!isRecord(paths)) {
    failures.push('Product visual evidence screenshot paths are malformed.');
    return;
  }
  if (!isRecord(exists)) {
    failures.push('Product visual evidence screenshot ready flags are malformed.');
    return;
  }
  for (const surface of productVisualSurfaces()) {
    const path = paths[surface];
    if (typeof path !== 'string' || !screenshotFileLooksLikeImage(resolve(path))) {
      failures.push(`Product visual evidence screenshot missing or invalid: ${surface}.`);
    }
    if (exists[surface] !== true) failures.push(`Product visual evidence screenshot ready flag is false: ${surface}.`);
  }
}

function validateProductVisualScreenshotFreshness(paths: unknown, metadata: unknown, freshFlags: unknown, ready: unknown, recordedAt: unknown): void {
  if (ready !== true) failures.push('Product visual evidence screenshots are not fresh for the recorded run.');
  if (!isRecord(paths) || !isRecord(metadata) || !isRecord(freshFlags)) {
    failures.push('Product visual evidence screenshot freshness metadata is malformed.');
    return;
  }
  const recordedAtMs = typeof recordedAt === 'string' ? Date.parse(recordedAt) : Number.NaN;
  if (!Number.isFinite(recordedAtMs)) failures.push('Product visual evidence missing valid recorded_at timestamp for screenshot freshness.');
  for (const surface of productVisualSurfaces()) {
    const path = paths[surface];
    const item = metadata[surface];
    if (!isRecord(item)) {
      failures.push(`Product visual evidence missing screenshot freshness metadata: ${surface}.`);
      continue;
    }
    const resolvedPath = typeof path === 'string' ? resolve(path) : undefined;
    if (typeof item.path !== 'string' || !resolvedPath || resolve(item.path) !== resolvedPath) failures.push(`Product visual evidence screenshot freshness path mismatch: ${surface}.`);
    if (item.ready !== true || item.fresh !== true || freshFlags[surface] !== true) failures.push(`Product visual evidence screenshot is stale: ${surface}.`);
    if (typeof item.mtime !== 'string' || !Number.isFinite(Date.parse(item.mtime))) failures.push(`Product visual evidence screenshot missing valid mtime: ${surface}.`);
    if (typeof item.recorded_at !== 'string' || !Number.isFinite(Date.parse(item.recorded_at))) failures.push(`Product visual evidence screenshot missing valid metadata recorded_at: ${surface}.`);
    if (typeof item.age_ms !== 'number' || item.age_ms < 0) failures.push(`Product visual evidence screenshot missing non-negative age: ${surface}.`);
    if (typeof item.freshness_window_ms !== 'number' || item.freshness_window_ms < 0) failures.push(`Product visual evidence screenshot missing freshness window: ${surface}.`);
    if (typeof item.size_bytes !== 'number' || item.size_bytes <= 0) failures.push(`Product visual evidence screenshot missing file size: ${surface}.`);
    if (Number.isFinite(recordedAtMs) && typeof item.mtime === 'string') {
      const mtimeMs = Date.parse(item.mtime);
      const windowMs = typeof item.freshness_window_ms === 'number' ? item.freshness_window_ms : 0;
      if (!Number.isFinite(mtimeMs) || recordedAtMs - mtimeMs > windowMs) failures.push(`Product visual evidence screenshot is older than freshness window: ${surface}.`);
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

function validatePowerPointRuntimeEvidence(evidence: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing PowerPoint runtime evidence ready flag.');
  if (!isRecord(evidence)) {
    failures.push('Product visual evidence missing PowerPoint runtime evidence details.');
    return;
  }
  if (evidence.ok !== true) failures.push('Product visual PowerPoint runtime evidence was not read successfully.');
  if (evidence.schema_version !== 1) failures.push(`Unsupported PowerPoint runtime evidence schema_version: ${evidence.schema_version}`);
  if (evidence.smoke_passed !== true) failures.push('Product visual evidence missing passed PowerPoint runtime smoke gate.');
  if (evidence.ready !== true) failures.push('Product visual PowerPoint runtime evidence is not ready.');
  const session = evidence.session;
  const details = evidence.smoke_details;
  if (!isRecord(session) || !isRecord(details)) {
    failures.push('Product visual PowerPoint runtime evidence missing session or smoke details.');
    return;
  }
  const document = isRecord(session.document) ? session.document : undefined;
  const host = isRecord(session.host) ? session.host : undefined;
  if (session.app !== 'powerpoint' || session.status !== 'active') failures.push('Product visual PowerPoint runtime evidence missing active PowerPoint session.');
  if (typeof session.session_id !== 'string' || details.session_id !== session.session_id) failures.push('Product visual PowerPoint runtime evidence session_id mismatch.');
  if (typeof document?.title !== 'string' || document.title.length === 0) failures.push('Product visual PowerPoint runtime evidence missing presentation title.');
  if (host?.app !== 'powerpoint') failures.push('Product visual PowerPoint runtime evidence missing PowerPoint host metadata.');
  validateExactToolCatalog('Product visual PowerPoint runtime evidence', session.available_tool_count, details.available_tools, powerPointV1Tools());
  if (details.mutation_proved !== true) failures.push('Product visual PowerPoint runtime evidence did not prove mutation path.');
  validatePowerPointCategoryProofs(details.tool_category_proofs, 'Product visual PowerPoint runtime evidence');
  if (!isRecord(details.presentation_info)) failures.push('Product visual PowerPoint runtime evidence missing presentation info proof.');
  if (!isRecord(details.active_view)) failures.push('Product visual PowerPoint runtime evidence missing active view proof.');
  if (!isRecord(details.list_slides)) failures.push('Product visual PowerPoint runtime evidence missing list_slides proof.');
  if (!isRecord(details.add_slide) || typeof details.add_slide.slide_id !== 'string') failures.push('Product visual PowerPoint runtime evidence missing add_slide proof.');
  if (!isRecord(details.add_text_box) || !isRecord(details.add_text_box.shape)) failures.push('Product visual PowerPoint runtime evidence missing add_text_box proof.');
  if (!isRecord(details.list_shapes) || !Array.isArray(details.list_shapes.shapes)) failures.push('Product visual PowerPoint runtime evidence missing list_shapes proof.');
  if (!isRecord(details.read_text) || !Array.isArray(details.read_text.items)) failures.push('Product visual PowerPoint runtime evidence missing read_text proof.');
  if (!isRecord(details.replace_text) || Number(details.replace_text.replacements ?? 0) < 1) failures.push('Product visual PowerPoint runtime evidence missing replace_text proof.');
  if (!isRecord(details.format_text) || details.format_text.formatted !== true) failures.push('Product visual PowerPoint runtime evidence missing format_text proof.');
  if (!isRecord(details.layout) || typeof details.layout.slide_id !== 'string') failures.push('Product visual PowerPoint runtime evidence missing apply_layout proof.');
  if (!isRecord(details.list_layouts) || !Array.isArray(details.list_layouts.masters)) failures.push('Product visual PowerPoint runtime evidence missing list_layouts proof.');
  const tableSupported = details.table_supported === true && isRecord(details.add_table) && typeof details.add_table.shape_id === 'string' && isRecord(details.read_table);
  const tableHostRejection = details.table_host_rejection === true;
  if (!tableSupported && !tableHostRejection) failures.push('Product visual PowerPoint runtime evidence missing table success or explicit host-capability rejection.');
  const exportSupported = details.export_supported === true && details.export_mime_type === 'application/pdf' && typeof details.export_size === 'number';
  const exportHostRejection = details.export_host_rejection === true;
  if (!exportSupported && !exportHostRejection) failures.push('Product visual PowerPoint runtime evidence missing export_file success or explicit host-capability rejection.');
}

function validateWordRuntimeEvidence(evidence: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing Word runtime evidence ready flag.');
  if (!isRecord(evidence)) {
    failures.push('Product visual evidence missing Word runtime evidence details.');
    return;
  }
  if (evidence.ok !== true) failures.push('Product visual Word runtime evidence was not read successfully.');
  if (evidence.schema_version !== 1) failures.push(`Unsupported Word runtime evidence schema_version: ${evidence.schema_version}`);
  if (evidence.smoke_passed !== true) failures.push('Product visual evidence missing passed Word runtime smoke gates.');
  if (evidence.ready !== true) failures.push('Product visual Word runtime evidence is not ready.');
  const session = evidence.session;
  const details = evidence.smoke_details;
  if (!isRecord(session) || !isRecord(details)) {
    failures.push('Product visual Word runtime evidence missing session or smoke details.');
    return;
  }
  const document = isRecord(session.document) ? session.document : undefined;
  const host = isRecord(session.host) ? session.host : undefined;
  if (session.app !== 'word' || session.status !== 'active') failures.push('Product visual Word runtime evidence missing active Word session.');
  if (typeof session.session_id !== 'string') failures.push('Product visual Word runtime evidence missing session_id.');
  if (typeof session.session_id !== 'string' || details.session_id !== session.session_id) failures.push('Product visual Word runtime evidence session_id mismatch.');
  if (typeof document?.title !== 'string' || document.title.length === 0) failures.push('Product visual Word runtime evidence missing document title.');
  if (host?.app !== 'word') failures.push('Product visual Word runtime evidence missing Word host metadata.');
  validateExactToolCatalog('Product visual Word runtime evidence', session.available_tool_count, details.available_tools, wordV1Tools());
  if (details.available_tool_count !== wordV1Tools().length) failures.push('Product visual Word runtime evidence missing read smoke available tool count.');
  if (Number(details.paragraph_0_text_length ?? 0) <= 0) failures.push('Product visual Word runtime evidence missing paragraph read proof.');
  if (Number(details.document_text_length ?? 0) <= 0) failures.push('Product visual Word runtime evidence missing document read proof.');
  if (Number(details.find_count ?? 0) < 1) failures.push('Product visual Word runtime evidence missing mutation readback.');
  if (details.full_smoke_passed !== true) failures.push('Product visual Word runtime evidence missing full smoke proof.');
  if (details.com_tracked_change_passed !== true) failures.push('Product visual Word runtime evidence missing COM tracked-change proof.');
}

function validateProductVisualOfficeToolE2e(evidence: unknown, ready: unknown): void {
  if (ready !== true) failures.push('Product visual evidence missing Office tool E2E ready flag.');
  if (!isRecord(evidence)) {
    failures.push('Product visual evidence missing Office tool E2E reports.');
    return;
  }
  validateEmbeddedOfficeToolE2eReport('Word', evidence.word);
  validateEmbeddedOfficeToolE2eReport('Excel', evidence.excel);
  validateEmbeddedOfficeToolE2eReport('PowerPoint', evidence.powerpoint);
}

function validateEmbeddedOfficeToolE2eReport(host: 'Word' | 'Excel' | 'PowerPoint', evidence: unknown): void {
  if (!isRecord(evidence)) {
    failures.push(`Product visual evidence missing Office tool E2E ${host} report.`);
    return;
  }
  if (evidence.ok !== true) failures.push(`Product visual Office tool E2E ${host} report was not read successfully.`);
  if (evidence.ready !== true) failures.push(`Product visual Office tool E2E ${host} report is not ready.`);
  if (evidence.schema_version !== 1) failures.push(`Product visual Office tool E2E ${host} report unsupported schema_version: ${evidence.schema_version}`);
  if (evidence.kind !== 'office_tool_e2e_report') failures.push(`Product visual Office tool E2E ${host} report has unsupported kind: ${evidence.kind ?? 'missing'}`);
  if (evidence.report_host !== host && evidence.host !== host) failures.push(`Product visual Office tool E2E ${host} report host mismatch.`);
  if (evidence.passed !== true) failures.push(`Product visual Office tool E2E ${host} report did not pass.`);
  validateOfficeToolE2eLifecycle(`Product visual Office tool E2E ${host} report`, evidence.lifecycle_counts);
  validateOfficeToolE2eActivation(`Product visual Office tool E2E ${host} report`, evidence.addin_activation);
  const advertisedTools = stringArray(evidence.advertised_tools);
  const sessionTools = stringArray(evidence.session_available_tools);
  const executedTools = stringArray(evidence.executed_tools);
  if (advertisedTools.length === 0) failures.push(`Product visual Office tool E2E ${host} report missing advertised tools.`);
  if (!sameStrings(advertisedTools, sessionTools)) failures.push(`Product visual Office tool E2E ${host} report session tools do not match advertised tools.`);
  if (!sameStrings(advertisedTools, executedTools)) failures.push(`Product visual Office tool E2E ${host} report executed tools do not match advertised tools.`);
  validateOfficeToolRuns(`Product visual Office tool E2E ${host} report`, advertisedTools, evidence.tool_runs);
}

function validateWordTaskpaneVisualEvidence(taskpane: unknown): void {
  validateTaskpaneDensityEvidence(taskpane, 'Word');
  if (isRecord(taskpane)) validateWordRuntimeEvidence(taskpane.runtime_evidence, taskpane.runtime_evidence_ready);
}

function validateExcelTaskpaneVisualEvidence(taskpane: unknown): void {
  validateTaskpaneDensityEvidence(taskpane, 'Excel');
  if (isRecord(taskpane)) validateExcelRuntimeEvidence(taskpane.runtime_evidence, taskpane.runtime_evidence_ready);
}

function validatePowerPointTaskpaneVisualEvidence(taskpane: unknown): void {
  validateTaskpaneDensityEvidence(taskpane, 'PowerPoint');
  if (isRecord(taskpane)) validatePowerPointRuntimeEvidence(taskpane.runtime_evidence, taskpane.runtime_evidence_ready);
}

function validateTaskpaneDensityEvidence(taskpane: unknown, label: 'Word' | 'Excel' | 'PowerPoint'): void {
  if (!isRecord(taskpane)) {
    failures.push(`Product visual evidence missing ${label} task pane details.`);
    return;
  }
  for (const [key, fieldLabel] of [
    ['compact_top_block', 'compact top block'],
    ['tools_permissions_merged', 'merged tools and permissions surface'],
    ['inline_settings', 'inline settings']
  ] as const) {
    if (taskpane[key] !== true) failures.push(`Product visual evidence missing ${label} ${fieldLabel}.`);
  }
  if (typeof taskpane.server_protocol_row !== 'string' || !/^Server .+ \/ Protocol .+$/.test(taskpane.server_protocol_row)) {
    failures.push(`Product visual evidence missing ${label} combined server/protocol row.`);
  }
  if (typeof taskpane.document_state !== 'string' || !/^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(taskpane.document_state) || /unknown/i.test(taskpane.document_state)) {
    failures.push(`Product visual evidence missing concrete ${label} editable/read-only/protected state.`);
  }
  if (taskpane.density_ready !== true) failures.push(`Product visual evidence missing ${label} task pane density pass flag.`);
}

function validateTaskpaneRuntimeBinding(label: 'Word' | 'Excel' | 'PowerPoint', runtime: unknown, taskpane: unknown): void {
  const topSessionId = sessionIdFromRuntimeEvidence(runtime);
  const taskpaneRuntime = isRecord(taskpane) ? taskpane.runtime_evidence : undefined;
  const taskpaneSessionId = sessionIdFromRuntimeEvidence(taskpaneRuntime);
  if (!topSessionId || !taskpaneSessionId || topSessionId !== taskpaneSessionId) {
    failures.push(`Product visual evidence ${label} task pane runtime evidence does not match top-level runtime evidence session.`);
  }
}

function sessionIdFromRuntimeEvidence(runtime: unknown): string | undefined {
  if (!isRecord(runtime)) return undefined;
  const session = isRecord(runtime.session) ? runtime.session : undefined;
  return typeof session?.session_id === 'string' ? session.session_id : undefined;
}

function validateOfficeToolE2eRuntimeBinding(label: 'Word' | 'Excel' | 'PowerPoint', runtime: unknown, officeToolE2e: unknown): void {
  const runtimeSessionId = sessionIdFromRuntimeEvidence(runtime);
  const reports = isRecord(officeToolE2e) ? officeToolE2e : undefined;
  const key = label.toLowerCase();
  const report = reports && isRecord(reports[key]) ? reports[key] : undefined;
  const session = report && isRecord(report.session) ? report.session : undefined;
  const e2eSessionId = typeof session?.session_id === 'string' ? session.session_id : undefined;
  if (!runtimeSessionId || !e2eSessionId || runtimeSessionId !== e2eSessionId) {
    failures.push(`Product visual evidence ${label} Office tool E2E report session does not match runtime evidence session.`);
  }
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
  validateExactToolCatalog('Product visual Excel runtime evidence', session.available_tool_count, details.available_tools, excelV1Tools());
  if (details.marker_found !== true) failures.push('Product visual Excel runtime evidence missing marker readback.');
  if (!isRecord(details.workbook_info)) failures.push('Product visual Excel runtime evidence missing get_workbook_info proof.');
  if (typeof details.sheet_list_count !== 'number') failures.push('Product visual Excel runtime evidence missing list_sheets proof.');
  if (!isRecord(details.updated_sheet)) failures.push('Product visual Excel runtime evidence missing update_sheet proof.');
  if (!isRecord(details.deleted_sheet) || details.deleted_sheet.deleted !== true) failures.push('Product visual Excel runtime evidence missing delete_sheet proof.');
  if (!isRecord(details.used_range)) failures.push('Product visual Excel runtime evidence missing get_used_range proof.');
  if (!isRecord(details.find_replace)) failures.push('Product visual Excel runtime evidence missing find_replace_cells proof.');
  if (!isRecord(details.clear)) failures.push('Product visual Excel runtime evidence missing clear_range proof.');
  if (!isRecord(details.write) || details.write.wrote_values !== true) failures.push('Product visual Excel runtime evidence missing write_range proof.');
  if (!isRecord(details.formula) || details.formula.wrote_formula !== true) failures.push('Product visual Excel runtime evidence missing set_formula proof.');
  if (!isRecord(details.format) || details.format.formatted !== true) failures.push('Product visual Excel runtime evidence missing format_range proof.');
  if (!isRecord(details.table) || typeof details.table.table !== 'string') failures.push('Product visual Excel runtime evidence missing create_table proof.');
  if (!isRecord(details.table_update)) failures.push('Product visual Excel runtime evidence missing update_table proof.');
  if (!isRecord(details.sort) || details.sort.sorted !== true) failures.push('Product visual Excel runtime evidence missing sort_range proof.');
  if (!isRecord(details.filter) || details.filter.filtered !== true) failures.push('Product visual Excel runtime evidence missing apply_filter proof.');
  if (!isRecord(details.chart) || typeof details.chart.chart !== 'string') failures.push('Product visual Excel runtime evidence missing create_chart proof.');
  if (!isRecord(details.chart_update) || details.chart_update.updated !== true) failures.push('Product visual Excel runtime evidence missing update_chart proof.');
  if (!isRecord(details.pivot_table) || typeof details.pivot_table.pivot_table !== 'string') failures.push('Product visual Excel runtime evidence missing create_pivot_table proof.');
  if (!isRecord(details.pivot_update) || details.pivot_update.refreshed !== true) failures.push('Product visual Excel runtime evidence missing update_pivot_table proof.');
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
    'daemon_main_window',
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
    'daemon_main_window',
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

function validateManualTrayObservedSnapshotBinding(manual: Record<string, unknown>, label: string): void {
  const context = isRecord(manual.daemon_context) ? manual.daemon_context : undefined;
  const trayProbe = isRecord(context?.tray_probe) ? context.tray_probe : undefined;
  const snapshot = isRecord(trayProbe?.snapshot) ? trayProbe.snapshot : undefined;
  if (!snapshot) return;

  if (typeof manual.observed_tooltip === 'string' && typeof snapshot.tooltip === 'string' && manual.observed_tooltip !== snapshot.tooltip) {
    failures.push(`${label} observed tooltip does not match daemon tray snapshot.`);
  }

  const observedMenu = Array.isArray(manual.observed_menu_items) ? manual.observed_menu_items.filter((item): item is string => typeof item === 'string') : [];
  const snapshotMenu = Array.isArray(snapshot.menu_items) ? snapshot.menu_items.filter((item): item is string => typeof item === 'string') : [];
  if (observedMenu.length > 0 && !sameMenuItems(observedMenu, snapshotMenu)) {
    failures.push(`${label} observed menu items do not match daemon tray snapshot.`);
  }
}

function sameMenuItems(left: string[], right: string[]): boolean {
  return normalizeObservedMenuItems(left).join('\n') === normalizeObservedMenuItems(right).join('\n');
}

function normalizeObservedMenuItems(items: string[]): string[] {
  return items.filter((item) => item !== '---');
}


function validateTrayMenuLabels(menuItems: string[], label: string): void {
  for (const expected of ['Status:', 'Clients:', 'Documents:', 'Show Office MCP Control', 'Quit Office MCP Control']) {
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
    { kind: 'action', enabled: true, label: /^Show Office MCP Control$/, action: 'show_ui' },
    { kind: 'action', enabled: true, label: /^Quit Office MCP Control$/, action: 'quit' }
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
  return /^Office MCP Control - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(value);
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
