import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { screenshotFileLooksLikeImage } from './image-evidence.js';

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/product-visual-evidence.json'));
const tester = readOption('--tester') ?? process.env.USERNAME ?? process.env.USER ?? 'unknown';
const notes = readOption('--notes');
const daemonBin = readOption('--daemon-bin');
const renderedLogoReviewPath = readOption('--rendered-logo-review-path');
const excelRuntimeEvidencePath = readOption('--excel-runtime-evidence-path');
const manualTrayEvidencePath = readOption('--manual-tray-evidence-path');
const wordManifestPath = resolve(readOption('--word-manifest-path') ?? join(repoRoot, 'src/office-ctl/word/manifest.xml'));
const excelManifestPath = resolve(readOption('--excel-manifest-path') ?? join(repoRoot, 'src/office-ctl/excel/manifest.xml'));
const powerPointManifestPath = resolve(readOption('--powerpoint-manifest-path') ?? join(repoRoot, 'src/office-ctl/powerpoint/manifest.xml'));

const requiredSurfaces = [
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

const observations = Object.fromEntries(requiredSurfaces.map((surface) => [surface, readOption(`--${surface.replaceAll('_', '-')}`)]));
const screenshotPaths = Object.fromEntries(requiredSurfaces.map((surface) => [surface, screenshotPathFor(surface)]));
const screenshotsExist = Object.fromEntries(
  requiredSurfaces.map((surface) => [surface, typeof screenshotPaths[surface] === 'string' && screenshotFileLooksLikeImage(resolve(screenshotPaths[surface] as string))])
);
const productName = readOption('--product-name') ?? 'Office MCP Control';
const trayTooltip = readOption('--tray-tooltip');
const catalogType = readOption('--catalog-type');
const catalogIconVisible = booleanFlag('--catalog-icon-visible');
const trayMenuNative = booleanFlag('--tray-menu-native');
const trayIconVisible = booleanFlag('--tray-icon-visible');
const quitConfirmationVisible = booleanFlag('--quit-confirmation-visible');
const logoQualityReviewed = booleanFlag('--logo-quality-reviewed');
const addinIdentityReviewed = booleanFlag('--addin-identity-reviewed');
const trayProductPolishReviewed = booleanFlag('--tray-product-polish-reviewed');
const renderedSizeLogoReviewed = booleanFlag('--rendered-size-logo-reviewed');
const wordFirstRunIdentityReviewed = booleanFlag('--word-first-run-identity-reviewed');
const excelFirstRunIdentityReviewed = booleanFlag('--excel-first-run-identity-reviewed');
const powerPointFirstRunIdentityReviewed = booleanFlag('--powerpoint-first-run-identity-reviewed');
const wordCatalogProvider = readOption('--word-catalog-provider');
const wordCatalogDescription = readOption('--word-catalog-description');
const wordCatalogType = readOption('--word-catalog-type');
const excelCatalogProvider = readOption('--excel-catalog-provider');
const excelCatalogDescription = readOption('--excel-catalog-description');
const excelCatalogType = readOption('--excel-catalog-type');
const powerPointCatalogProvider = readOption('--powerpoint-catalog-provider');
const powerPointCatalogDescription = readOption('--powerpoint-catalog-description');
const powerPointCatalogType = readOption('--powerpoint-catalog-type');
const excelCompactTopBlock = booleanFlag('--excel-compact-top-block');
const excelToolsPermissionsMerged = booleanFlag('--excel-tools-permissions-merged');
const excelInlineSettings = booleanFlag('--excel-inline-settings');
const excelServerProtocolRow = readOption('--excel-server-protocol-row');
const excelDocumentState = readOption('--excel-document-state');
const daemonContext = daemonBin ? readDaemonContext(resolve(daemonBin)) : undefined;
const daemonContextReady = daemonContextLooksReady(daemonContext);
const renderedLogoReview = renderedLogoReviewPath ? readRenderedLogoReview(resolve(renderedLogoReviewPath)) : undefined;
const renderedLogoReviewReady = renderedLogoReviewLooksReady(renderedLogoReview);
const excelRuntimeEvidence = excelRuntimeEvidencePath ? readExcelRuntimeEvidence(resolve(excelRuntimeEvidencePath)) : undefined;
const excelRuntimeEvidenceReady = excelRuntimeEvidenceLooksReady(excelRuntimeEvidence);
const manualTrayEvidence = manualTrayEvidencePath ? readManualTrayEvidence(resolve(manualTrayEvidencePath)) : undefined;
const manualTrayEvidenceReady = manualTrayEvidenceLooksReady(manualTrayEvidence);
const wordManifestIdentity = readManifestIdentity(wordManifestPath);
const excelManifestIdentity = readManifestIdentity(excelManifestPath);
const powerPointManifestIdentity = readManifestIdentity(powerPointManifestPath);

const productTextReady = requiredSurfaces.filter((surface) => surface !== 'tray_tooltip').every((surface) => typeof observations[surface] === 'string' && (observations[surface] as string).includes(productName));
const allScreenshotsExist = Object.values(screenshotsExist).every(Boolean);
const trayTooltipReady = typeof trayTooltip === 'string' && /^Office MCP - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(trayTooltip);
const catalogTypeReady = typeof catalogType === 'string' && /local productivity automation control utility/i.test(catalogType);
const wordFirstRunIdentity = firstRunIdentity(wordManifestIdentity, wordCatalogProvider, wordCatalogDescription, wordCatalogType, catalogType);
const excelFirstRunIdentity = firstRunIdentity(excelManifestIdentity, excelCatalogProvider, excelCatalogDescription, excelCatalogType, catalogType);
const powerPointFirstRunIdentity = firstRunIdentity(powerPointManifestIdentity, powerPointCatalogProvider, powerPointCatalogDescription, powerPointCatalogType, catalogType);
const wordFirstRunIdentityReady = wordFirstRunIdentityReviewed && catalogIdentityLooksReady(wordFirstRunIdentity);
const excelFirstRunIdentityReady = excelFirstRunIdentityReviewed && catalogIdentityLooksReady(excelFirstRunIdentity);
const powerPointFirstRunIdentityReady = powerPointFirstRunIdentityReviewed && catalogIdentityLooksReady(powerPointFirstRunIdentity);
const excelServerProtocolReady = typeof excelServerProtocolRow === 'string' && /^Server .+ \/ Protocol .+$/.test(excelServerProtocolRow);
const excelDocumentStateReady = typeof excelDocumentState === 'string' && /^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(excelDocumentState) && !/unknown/i.test(excelDocumentState);
const excelTaskpaneDensityReady = excelCompactTopBlock && excelToolsPermissionsMerged && excelInlineSettings && excelServerProtocolReady && excelDocumentStateReady && excelRuntimeEvidenceReady;
const productIdentityReviewReady = logoQualityReviewed && renderedSizeLogoReviewed && renderedLogoReviewReady && addinIdentityReviewed && wordFirstRunIdentityReady && excelFirstRunIdentityReady && powerPointFirstRunIdentityReady && trayProductPolishReviewed;
const passed = productTextReady && allScreenshotsExist && trayTooltipReady && catalogTypeReady && catalogIconVisible && trayMenuNative && trayIconVisible && quitConfirmationVisible && manualTrayEvidenceReady && excelTaskpaneDensityReady && productIdentityReviewReady && renderedLogoReviewReady && daemonContextReady;

const evidence = {
  schema_version: 1,
  kind: 'product_visual_evidence',
  recorded_at: new Date().toISOString(),
  tester,
  platform: process.platform,
  product_name: productName,
  required_surfaces: requiredSurfaces,
  observations,
  screenshot_paths: screenshotPaths,
  screenshots_exist: screenshotsExist,
  product_text_ready: productTextReady,
  catalog_type: catalogType,
  catalog_type_ready: catalogTypeReady,
  catalog_icon_visible: catalogIconVisible,
  tray_tooltip: trayTooltip,
  tray_tooltip_ready: trayTooltipReady,
  tray_icon_visible: trayIconVisible,
  tray_menu_native: trayMenuNative,
  quit_confirmation_visible: quitConfirmationVisible,
  manual_tray_evidence: manualTrayEvidence,
  manual_tray_evidence_ready: manualTrayEvidenceReady,
  product_identity_review: {
    logo_quality_reviewed: logoQualityReviewed,
    rendered_size_logo_reviewed: renderedSizeLogoReviewed,
    rendered_logo_review_ready: renderedLogoReviewReady,
    addin_identity_reviewed: addinIdentityReviewed,
    word_first_run_identity_reviewed: wordFirstRunIdentityReviewed,
    excel_first_run_identity_reviewed: excelFirstRunIdentityReviewed,
    powerpoint_first_run_identity_reviewed: powerPointFirstRunIdentityReviewed,
    tray_product_polish_reviewed: trayProductPolishReviewed,
    word_first_run_identity_ready: wordFirstRunIdentityReady,
    excel_first_run_identity_ready: excelFirstRunIdentityReady,
    powerpoint_first_run_identity_ready: powerPointFirstRunIdentityReady,
    ready: productIdentityReviewReady
  },
  first_run_identity: {
    word: {
      manifest_path: wordManifestPath,
      ...wordFirstRunIdentity,
      ready: wordFirstRunIdentityReady
    },
    excel: {
      manifest_path: excelManifestPath,
      ...excelFirstRunIdentity,
      ready: excelFirstRunIdentityReady
    },
    powerpoint: {
      manifest_path: powerPointManifestPath,
      ...powerPointFirstRunIdentity,
      ready: powerPointFirstRunIdentityReady
    }
  },
  rendered_logo_review: renderedLogoReview,
  rendered_logo_review_ready: renderedLogoReviewReady,
  excel_taskpane: {
    compact_top_block: excelCompactTopBlock,
    tools_permissions_merged: excelToolsPermissionsMerged,
    inline_settings: excelInlineSettings,
    server_protocol_row: excelServerProtocolRow,
    server_protocol_row_ready: excelServerProtocolReady,
    document_state: excelDocumentState,
    document_state_ready: excelDocumentStateReady,
    runtime_evidence: excelRuntimeEvidence,
    runtime_evidence_ready: excelRuntimeEvidenceReady,
    density_ready: excelTaskpaneDensityReady
  },
  daemon_context: daemonContext,
  daemon_context_ready: daemonContextReady,
  notes,
  passed
};

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, JSON.stringify(evidence, null, 2));
console.log(JSON.stringify(evidence, null, 2));
if (!passed) process.exit(1);

function screenshotPathFor(surface: string): string | undefined {
  return readOption(`--${surface.replaceAll('_', '-')}-screenshot`);
}

function booleanFlag(name: string): boolean {
  const value = readOption(name);
  if (value === undefined) return false;
  return ['1', 'true', 'yes', 'passed'].includes(value.toLowerCase());
}

function readOption(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  if (index === -1) return undefined;
  return process.argv[index + 1];
}

function firstRunIdentity(manifest: Record<string, unknown>, providerOverride: string | undefined, descriptionOverride: string | undefined, typeOverride: string | undefined, catalogTypeFallback: string | undefined): Record<string, unknown> {
  return {
    display_name: manifest.display_name,
    provider: providerOverride ?? manifest.provider,
    description: descriptionOverride ?? manifest.description,
    type: typeOverride ?? catalogTypeFallback,
    icon_url: manifest.icon_url,
    high_resolution_icon_url: manifest.high_resolution_icon_url,
    manifest_ready: manifest.ready === true
  };
}

function catalogIdentityLooksReady(identity: Record<string, unknown>): boolean {
  return typeof identity.display_name === 'string' && identity.display_name.includes(productName)
    && typeof identity.provider === 'string' && identity.provider.includes(productName)
    && typeof identity.description === 'string' && /local/i.test(identity.description) && /(productivity|office)/i.test(identity.description) && /(automation|control)/i.test(identity.description)
    && typeof identity.type === 'string' && /local productivity automation control utility/i.test(identity.type)
    && typeof identity.icon_url === 'string' && /\/assets\/icon-32\.png/.test(identity.icon_url)
    && typeof identity.high_resolution_icon_url === 'string' && /\/assets\/icon-80\.png/.test(identity.high_resolution_icon_url)
    && identity.manifest_ready === true;
}

function readManifestIdentity(path: string): Record<string, unknown> {
  try {
    const xml = readFileSync(path, 'utf8');
    return {
      path,
      ready: true,
      display_name: extractDefaultValue(xml, 'DisplayName'),
      provider: extractElementText(xml, 'ProviderName'),
      description: extractDefaultValue(xml, 'Description'),
      icon_url: extractDefaultValue(xml, 'IconUrl'),
      high_resolution_icon_url: extractDefaultValue(xml, 'HighResolutionIconUrl')
    };
  } catch (error) {
    return { path, ready: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function extractElementText(xml: string, element: string): string | undefined {
  return new RegExp(`<${element}>([^<]+)<\\/${element}>`).exec(xml)?.[1];
}

function extractDefaultValue(xml: string, element: string): string | undefined {
  return new RegExp(`<${element}[^>]*DefaultValue="([^"]+)"`).exec(xml)?.[1];
}

function readRenderedLogoReview(path: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    return { path, ok: true, ...parsed };
  } catch (error) {
    return { path, ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function renderedLogoReviewLooksReady(review: Record<string, unknown> | undefined): boolean {
  if (!review) return false;
  if (review.ok !== true || review.schema_version !== 1 || review.kind !== 'rendered_logo_review' || review.product_name !== productName || review.ready !== true) return false;
  if (typeof review.sheet_path !== 'string' || !screenshotFileLooksLikeImage(resolve(review.sheet_path))) return false;
  const surfaces = Array.isArray(review.surfaces) ? review.surfaces.filter(isRecord) : [];
  return [
    ['logo_tray_size', 16],
    ['logo_ribbon_size', 32],
    ['logo_catalog_thumbnail', 80],
    ['logo_daemon_titlebar', 20],
    ['logo_installer_metadata', 256]
  ].every(([key, size]) => surfaces.some((surface) => surface.key === key && surface.rendered_size_px === size && surface.width === size && surface.height === size && surface.non_empty === true && surface.palette_ready === true && surface.expected_size_ready === true));
}

function readExcelRuntimeEvidence(path: string): Record<string, unknown> {
  try {
    const report = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    const gates = Array.isArray(report.gates) ? report.gates.filter(isRecord) : [];
    const discovery = gates.find((gate) => gate.name === 'word.session_discovery' && gate.status === 'passed');
    const smoke = gates.find((gate) => gate.name === 'excel.runtime_smoke' && gate.status === 'passed');
    const session = excelSessionFromDiscovery(discovery, smoke);
    const details = isRecord(smoke?.details) ? smoke.details : undefined;
    return {
      path,
      ok: true,
      schema_version: report.schema_version,
      endpoint: report.endpoint,
      generated_at: report.generated_at,
      session,
      smoke_details: details,
      smoke_passed: Boolean(smoke),
      ready: excelRuntimeDetailsLookReady(session, details)
    };
  } catch (error) {
    return { path, ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function excelSessionFromDiscovery(discovery: Record<string, unknown> | undefined, smoke: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  const smokeDetails = isRecord(smoke?.details) ? smoke.details : undefined;
  const smokeSessionId = typeof smokeDetails?.session_id === 'string' ? smokeDetails.session_id : undefined;
  const discoveryDetails = isRecord(discovery?.details) ? discovery.details : undefined;
  const sessions = Array.isArray(discoveryDetails?.sessions) ? discoveryDetails.sessions.filter(isRecord) : [];
  return sessions.find((session) => session.app === 'excel' && (!smokeSessionId || session.session_id === smokeSessionId));
}

function excelRuntimeEvidenceLooksReady(evidence: Record<string, unknown> | undefined): boolean {
  if (!evidence) return false;
  return evidence.ok === true && evidence.schema_version === 1 && evidence.smoke_passed === true && evidence.ready === true;
}

function excelRuntimeDetailsLookReady(session: Record<string, unknown> | undefined, details: Record<string, unknown> | undefined): boolean {
  if (!session || !details) return false;
  const document = isRecord(session.document) ? session.document : undefined;
  const host = isRecord(session.host) ? session.host : undefined;
  return session.app === 'excel'
    && session.status === 'active'
    && typeof session.session_id === 'string'
    && typeof details.session_id === 'string'
    && session.session_id === details.session_id
    && typeof document?.title === 'string'
    && document.title.length > 0
    && host?.app === 'excel'
    && typeof session.available_tool_count === 'number'
    && session.available_tool_count >= 7
    && details.marker_found === true
    && isRecord(details.write) && details.write.wrote_values === true
    && isRecord(details.formula) && details.formula.wrote_formula === true
    && isRecord(details.format) && details.format.formatted === true
    && isRecord(details.table) && typeof details.table.table === 'string'
    && isRecord(details.chart) && typeof details.chart.chart === 'string'
    && isRecord(details.sheet) && details.sheet.activated === true;
}

function readManualTrayEvidence(path: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    return { path, ok: true, ...parsed };
  } catch (error) {
    return { path, ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function manualTrayEvidenceLooksReady(evidence: Record<string, unknown> | undefined): boolean {
  if (!evidence) return false;
  if (evidence.ok !== true || evidence.schema_version !== 1 || evidence.kind !== 'tray_manual_evidence' || evidence.platform !== 'win32' || evidence.passed !== true) return false;
  if (evidence.visible_icon !== true || evidence.right_click_menu !== true || evidence.menu_opened_from_tray_icon !== true || evidence.native_menu_appearance_reviewed !== true || evidence.show_ui_opened !== true) return false;
  if (typeof evidence.observed_tooltip !== 'string' || !/^Office MCP - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(evidence.observed_tooltip)) return false;
  if (typeof evidence.screenshot_path !== 'string' || !screenshotFileLooksLikeImage(resolve(evidence.screenshot_path))) return false;
  const items = Array.isArray(evidence.observed_menu_items) ? evidence.observed_menu_items.filter((item): item is string => typeof item === 'string') : [];
  return ['Status:', 'Clients:', 'Documents:', 'Show Office MCP', 'Quit Office MCP'].every((expected) => items.some((item) => item.includes(expected)))
    && evidence.daemon_context_ready === true
    && daemonContextLooksReady(isRecord(evidence.daemon_context) ? evidence.daemon_context : undefined);
}

function readDaemonContext(binaryPath: string): Record<string, unknown> {
  return {
    binary_path: binaryPath,
    status: runJson(binaryPath, ['daemon', 'status']),
    tray_probe: runJson(binaryPath, ['tray', '--probe'])
  };
}

function runJson(binaryPath: string, args: string[]): Record<string, unknown> {
  const result = spawnSync(binaryPath, args, { encoding: 'utf8', shell: process.platform === 'win32' && binaryPath.toLowerCase().endsWith('.cmd') });
  if (result.status !== 0) {
    return {
      ok: false,
      error: result.error instanceof Error ? result.error.message : undefined,
      exit_code: result.status,
      stderr: result.stderr?.trim() ?? '',
      stdout: result.stdout?.trim() ?? ''
    };
  }
  try {
    const parsed = JSON.parse(result.stdout) as Record<string, unknown>;
    return { ok: true, ...parsed };
  } catch (error) {
    return {
      ok: false,
      parse_error: error instanceof Error ? error.message : String(error),
      stdout: result.stdout.trim()
    };
  }
}

function daemonContextLooksReady(context: Record<string, unknown> | undefined): boolean {
  if (!context) return false;
  const status = context.status;
  const trayProbe = context.tray_probe;
  return isRecord(status) && status.ok === true && status.running === true && typeof status.uiUrl === 'string'
    && isRecord(trayProbe) && trayProbe.ok === true && trayProbe.native_host === true && trayProbe.state_fetch_ok === true
    && traySnapshotLooksReady(trayProbe.snapshot);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function traySnapshotLooksReady(snapshot: unknown): boolean {
  if (!isRecord(snapshot)) return false;
  const tooltipReady = typeof snapshot.tooltip === 'string' && /^Office MCP - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(snapshot.tooltip);
  const menuItems = Array.isArray(snapshot.menu_items) ? snapshot.menu_items.filter((item): item is string => typeof item === 'string') : [];
  const menuReady = ['Status:', 'Clients:', 'Documents:', 'Show Office MCP', 'Quit Office MCP']
    .every((expected) => menuItems.some((item) => item.includes(expected)));
  return tooltipReady && menuReady;
}
