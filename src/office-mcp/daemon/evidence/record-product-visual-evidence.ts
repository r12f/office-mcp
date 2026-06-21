import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { screenshotFileLooksLikeImage } from './image-evidence.js';

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const brandAssetRoot = join(repoRoot, 'src/office-ctl/common/assets');
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/product-visual-evidence.json'));
const tester = readOption('--tester') ?? process.env.USERNAME ?? process.env.USER ?? 'unknown';
const notes = readOption('--notes');
const daemonBin = readOption('--daemon-bin');
const renderedLogoReviewPath = readOption('--rendered-logo-review-path') ?? process.env.OFFICE_MCP_RENDERED_LOGO_REVIEW_PATH;
const wordRuntimeEvidencePath = readOption('--word-runtime-evidence-path') ?? process.env.OFFICE_MCP_WORD_RUNTIME_EVIDENCE_PATH;
const excelRuntimeEvidencePath = readOption('--excel-runtime-evidence-path') ?? process.env.OFFICE_MCP_EXCEL_RUNTIME_EVIDENCE_PATH;
const powerPointRuntimeEvidencePath = readOption('--powerpoint-runtime-evidence-path') ?? process.env.OFFICE_MCP_POWERPOINT_RUNTIME_EVIDENCE_PATH;
const wordToolE2eReportPath = readOption('--word-tool-e2e-report-path') ?? process.env.OFFICE_MCP_WORD_TOOL_E2E_REPORT_PATH;
const excelToolE2eReportPath = readOption('--excel-tool-e2e-report-path') ?? process.env.OFFICE_MCP_EXCEL_TOOL_E2E_REPORT_PATH;
const powerPointToolE2eReportPath = readOption('--powerpoint-tool-e2e-report-path') ?? process.env.OFFICE_MCP_POWERPOINT_TOOL_E2E_REPORT_PATH;
const manualTrayEvidencePath = readOption('--manual-tray-evidence-path') ?? process.env.OFFICE_MCP_TRAY_MANUAL_EVIDENCE_PATH;
const catalogIdentityReviewPath = readOption('--catalog-identity-review-path') ?? process.env.OFFICE_MCP_CATALOG_IDENTITY_REVIEW_PATH;
const wordManifestPath = resolve(readOption('--word-manifest-path') ?? join(repoRoot, 'src/office-ctl/word/manifest.xml'));
const excelManifestPath = resolve(readOption('--excel-manifest-path') ?? join(repoRoot, 'src/office-ctl/excel/manifest.xml'));
const powerPointManifestPath = resolve(readOption('--powerpoint-manifest-path') ?? join(repoRoot, 'src/office-ctl/powerpoint/manifest.xml'));
const productName = readOption('--product-name') ?? 'Office MCP Control';
const renderedLogoReview = renderedLogoReviewPath ? readRenderedLogoReview(resolve(renderedLogoReviewPath)) : undefined;
const renderedLogoReviewReady = renderedLogoReviewLooksReady(renderedLogoReview);
const manualTrayEvidence = manualTrayEvidencePath ? readManualTrayEvidence(resolve(manualTrayEvidencePath)) : undefined;
const manualTrayEvidenceReady = manualTrayEvidenceLooksReady(manualTrayEvidence);
const catalogIdentityReview = catalogIdentityReviewPath ? readCatalogIdentityReview(resolve(catalogIdentityReviewPath)) : undefined;
const catalogIdentityReviewReady = catalogIdentityReviewLooksReady(catalogIdentityReview);

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

const excelV1Tools = [
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

const powerPointV1Tools = [
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

const wordV1Tools = [
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

const observations = Object.fromEntries(requiredSurfaces.map((surface) => [surface, observationFor(surface)]));
const screenshotPaths = Object.fromEntries(requiredSurfaces.map((surface) => [surface, screenshotPathFor(surface)]));
const screenshotsExist = Object.fromEntries(
  requiredSurfaces.map((surface) => [surface, typeof screenshotPaths[surface] === 'string' && screenshotFileLooksLikeImage(resolve(screenshotPaths[surface] as string))])
);
const screenshotMetadata = Object.fromEntries(requiredSurfaces.map((surface) => [surface, screenshotMetadataFor(screenshotPaths[surface])]));
const screenshotsFresh = Object.fromEntries(requiredSurfaces.map((surface) => [surface, screenshotMetadata[surface]?.fresh === true]));
const trayTooltip = readOption('--tray-tooltip');
const catalogType = readOption('--catalog-type') ?? (typeof catalogIdentityReview?.catalog_type === 'string' ? catalogIdentityReview.catalog_type : undefined);
const catalogIconVisible = booleanFlag('--catalog-icon-visible');
const trayMenuNative = booleanFlag('--tray-menu-native');
const trayMenuSurfaceKind = readOption('--tray-menu-surface-kind') ?? (typeof manualTrayEvidence?.tray_menu_surface_kind === 'string' ? manualTrayEvidence.tray_menu_surface_kind : undefined);
const trayMenuSurfaceNative = trayMenuSurfaceKind === 'native';
const trayIconVisible = booleanFlag('--tray-icon-visible');
const quitConfirmationVisible = booleanFlag('--quit-confirmation-visible');
const logoQualityReviewed = booleanFlag('--logo-quality-reviewed');
const logoFutureOfficeControlReviewed = booleanFlag('--logo-future-office-control-reviewed');
const finalLogoUserSurfaceReviewed = booleanFlag('--final-logo-user-surface-reviewed');
const currentLogoScreenshotFeedbackReviewed = booleanFlag('--current-logo-screenshot-feedback-reviewed');
const addinIdentityReviewed = booleanFlag('--addin-identity-reviewed');
const addinTitleIconTypeReviewed = booleanFlag('--addin-title-icon-type-reviewed');
const addinInstallableSurfaceReviewed = booleanFlag('--addin-installable-surface-reviewed');
const currentAddinScreenshotFeedbackReviewed = booleanFlag('--current-addin-screenshot-feedback-reviewed');
const trayProductPolishReviewed = booleanFlag('--tray-product-polish-reviewed');
const trayNativeFirstImpressionReviewed = booleanFlag('--tray-native-first-impression-reviewed');
const trayNormalWindowsLaunchReviewed = booleanFlag('--tray-normal-windows-launch-reviewed');
const currentTrayScreenshotFeedbackReviewed = booleanFlag('--current-tray-screenshot-feedback-reviewed');
const daemonMainWindowReviewed = booleanFlag('--daemon-main-window-reviewed');
const daemonMainWindowCompactReviewed = booleanFlag('--daemon-main-window-compact-reviewed');
const daemonMainWindowThreeColumnReviewed = booleanFlag('--daemon-main-window-three-column-reviewed');
const renderedSizeLogoReviewed = booleanFlag('--rendered-size-logo-reviewed');
const wordFirstRunIdentityReviewed = booleanFlag('--word-first-run-identity-reviewed');
const excelFirstRunIdentityReviewed = booleanFlag('--excel-first-run-identity-reviewed');
const powerPointFirstRunIdentityReviewed = booleanFlag('--powerpoint-first-run-identity-reviewed');
const wordCatalogProvider = readOption('--word-catalog-provider') ?? catalogIdentityHostValue('word', 'provider');
const wordCatalogDescription = readOption('--word-catalog-description') ?? catalogIdentityHostValue('word', 'description');
const wordCatalogType = readOption('--word-catalog-type') ?? (catalogIdentityReviewReady ? catalogType : undefined);
const excelCatalogProvider = readOption('--excel-catalog-provider') ?? catalogIdentityHostValue('excel', 'provider');
const excelCatalogDescription = readOption('--excel-catalog-description') ?? catalogIdentityHostValue('excel', 'description');
const excelCatalogType = readOption('--excel-catalog-type') ?? (catalogIdentityReviewReady ? catalogType : undefined);
const powerPointCatalogProvider = readOption('--powerpoint-catalog-provider') ?? catalogIdentityHostValue('powerpoint', 'provider');
const powerPointCatalogDescription = readOption('--powerpoint-catalog-description') ?? catalogIdentityHostValue('powerpoint', 'description');
const powerPointCatalogType = readOption('--powerpoint-catalog-type') ?? (catalogIdentityReviewReady ? catalogType : undefined);
const wordCompactTopBlock = booleanFlag('--word-compact-top-block');
const wordToolsPermissionsMerged = booleanFlag('--word-tools-permissions-merged');
const wordInlineSettings = booleanFlag('--word-inline-settings');
const wordServerProtocolRow = readOption('--word-server-protocol-row');
const wordDocumentState = readOption('--word-document-state');
const excelCompactTopBlock = booleanFlag('--excel-compact-top-block');
const excelToolsPermissionsMerged = booleanFlag('--excel-tools-permissions-merged');
const excelInlineSettings = booleanFlag('--excel-inline-settings');
const excelServerProtocolRow = readOption('--excel-server-protocol-row');
const excelDocumentState = readOption('--excel-document-state');
const powerPointCompactTopBlock = booleanFlag('--powerpoint-compact-top-block');
const powerPointToolsPermissionsMerged = booleanFlag('--powerpoint-tools-permissions-merged');
const powerPointInlineSettings = booleanFlag('--powerpoint-inline-settings');
const powerPointServerProtocolRow = readOption('--powerpoint-server-protocol-row');
const powerPointDocumentState = readOption('--powerpoint-document-state');
const daemonContext = daemonBin ? readDaemonContext(resolve(daemonBin)) : undefined;
const daemonContextReady = daemonContextLooksReady(daemonContext);
const wordRuntimeEvidence = wordRuntimeEvidencePath ? readWordRuntimeEvidence(resolve(wordRuntimeEvidencePath)) : undefined;
const wordRuntimeEvidenceReady = wordRuntimeEvidenceLooksReady(wordRuntimeEvidence);
const excelRuntimeEvidence = excelRuntimeEvidencePath ? readExcelRuntimeEvidence(resolve(excelRuntimeEvidencePath)) : undefined;
const excelRuntimeEvidenceReady = excelRuntimeEvidenceLooksReady(excelRuntimeEvidence);
const powerPointRuntimeEvidence = powerPointRuntimeEvidencePath ? readPowerPointRuntimeEvidence(resolve(powerPointRuntimeEvidencePath)) : undefined;
const powerPointRuntimeEvidenceReady = powerPointRuntimeEvidenceLooksReady(powerPointRuntimeEvidence);
const officeToolE2e = {
  word: wordToolE2eReportPath ? readOfficeToolE2eReport('Word', resolve(wordToolE2eReportPath)) : undefined,
  excel: excelToolE2eReportPath ? readOfficeToolE2eReport('Excel', resolve(excelToolE2eReportPath)) : undefined,
  powerpoint: powerPointToolE2eReportPath ? readOfficeToolE2eReport('PowerPoint', resolve(powerPointToolE2eReportPath)) : undefined
};
const officeToolE2eReady = officeToolE2eReportLooksReady(officeToolE2e.word) && officeToolE2eReportLooksReady(officeToolE2e.excel) && officeToolE2eReportLooksReady(officeToolE2e.powerpoint);
const wordManifestIdentity = readManifestIdentity(wordManifestPath);
const excelManifestIdentity = readManifestIdentity(excelManifestPath);
const powerPointManifestIdentity = readManifestIdentity(powerPointManifestPath);

const productTextReady = requiredSurfaces.filter((surface) => surface !== 'tray_tooltip').every((surface) => typeof observations[surface] === 'string' && (observations[surface] as string).includes(productName));
const allScreenshotsExist = Object.values(screenshotsExist).every(Boolean);
const allScreenshotsFresh = Object.values(screenshotsFresh).every(Boolean);
const trayTooltipReady = typeof trayTooltip === 'string' && /^Office MCP Control - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(trayTooltip);
const catalogTypeReady = productCatalogTypeLooksReady(catalogType);
const wordFirstRunIdentity = firstRunIdentity(wordManifestIdentity, wordCatalogProvider, wordCatalogDescription, wordCatalogType, catalogType);
const excelFirstRunIdentity = firstRunIdentity(excelManifestIdentity, excelCatalogProvider, excelCatalogDescription, excelCatalogType, catalogType);
const powerPointFirstRunIdentity = firstRunIdentity(powerPointManifestIdentity, powerPointCatalogProvider, powerPointCatalogDescription, powerPointCatalogType, catalogType);
const wordFirstRunIdentityReady = wordFirstRunIdentityReviewed && catalogIdentityLooksReady(wordFirstRunIdentity);
const excelFirstRunIdentityReady = excelFirstRunIdentityReviewed && catalogIdentityLooksReady(excelFirstRunIdentity);
const powerPointFirstRunIdentityReady = powerPointFirstRunIdentityReviewed && catalogIdentityLooksReady(powerPointFirstRunIdentity);
const wordServerProtocolReady = typeof wordServerProtocolRow === 'string' && /^Server .+ \/ Protocol .+$/.test(wordServerProtocolRow);
const wordDocumentStateReady = typeof wordDocumentState === 'string' && /^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(wordDocumentState) && !/unknown/i.test(wordDocumentState);
const wordTaskpaneDensityReady = wordCompactTopBlock && wordToolsPermissionsMerged && wordInlineSettings && wordServerProtocolReady && wordDocumentStateReady && wordRuntimeEvidenceReady;
const excelServerProtocolReady = typeof excelServerProtocolRow === 'string' && /^Server .+ \/ Protocol .+$/.test(excelServerProtocolRow);
const excelDocumentStateReady = typeof excelDocumentState === 'string' && /^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(excelDocumentState) && !/unknown/i.test(excelDocumentState);
const excelTaskpaneDensityReady = excelCompactTopBlock && excelToolsPermissionsMerged && excelInlineSettings && excelServerProtocolReady && excelDocumentStateReady && excelRuntimeEvidenceReady;
const powerPointServerProtocolReady = typeof powerPointServerProtocolRow === 'string' && /^Server .+ \/ Protocol .+$/.test(powerPointServerProtocolRow);
const powerPointDocumentStateReady = typeof powerPointDocumentState === 'string' && /^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(powerPointDocumentState) && !/unknown/i.test(powerPointDocumentState);
const powerPointTaskpaneDensityReady = powerPointCompactTopBlock && powerPointToolsPermissionsMerged && powerPointInlineSettings && powerPointServerProtocolReady && powerPointDocumentStateReady && powerPointRuntimeEvidenceReady;
const currentScreenshotFeedbackReady = currentLogoScreenshotFeedbackReviewed && currentAddinScreenshotFeedbackReviewed && currentTrayScreenshotFeedbackReviewed;
const daemonMainWindowReady = daemonMainWindowReviewed && daemonMainWindowCompactReviewed && daemonMainWindowThreeColumnReviewed;
const productIdentityReviewReady = logoQualityReviewed && logoFutureOfficeControlReviewed && finalLogoUserSurfaceReviewed && currentLogoScreenshotFeedbackReviewed && renderedSizeLogoReviewed && renderedLogoReviewReady && catalogIdentityReviewReady && addinIdentityReviewed && addinTitleIconTypeReviewed && addinInstallableSurfaceReviewed && currentAddinScreenshotFeedbackReviewed && wordFirstRunIdentityReady && excelFirstRunIdentityReady && powerPointFirstRunIdentityReady && wordRuntimeEvidenceReady && excelRuntimeEvidenceReady && powerPointRuntimeEvidenceReady && trayProductPolishReviewed && trayNativeFirstImpressionReviewed && trayNormalWindowsLaunchReviewed && currentTrayScreenshotFeedbackReviewed;
const passed = productTextReady && allScreenshotsExist && allScreenshotsFresh && trayTooltipReady && catalogTypeReady && catalogIdentityReviewReady && catalogIconVisible && trayMenuNative && trayMenuSurfaceNative && trayIconVisible && quitConfirmationVisible && manualTrayEvidenceReady && officeToolE2eReady && wordTaskpaneDensityReady && excelTaskpaneDensityReady && powerPointTaskpaneDensityReady && daemonMainWindowReady && productIdentityReviewReady && renderedLogoReviewReady && wordRuntimeEvidenceReady && excelRuntimeEvidenceReady && powerPointRuntimeEvidenceReady && daemonContextReady;

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
  screenshot_metadata: screenshotMetadata,
  screenshots_fresh: screenshotsFresh,
  screenshots_fresh_ready: allScreenshotsFresh,
  product_text_ready: productTextReady,
  catalog_type: catalogType,
  catalog_type_ready: catalogTypeReady,
  catalog_identity_review: catalogIdentityReview,
  catalog_identity_review_ready: catalogIdentityReviewReady,
  catalog_icon_visible: catalogIconVisible,
  tray_tooltip: trayTooltip,
  tray_tooltip_ready: trayTooltipReady,
  tray_icon_visible: trayIconVisible,
  tray_menu_native: trayMenuNative,
  tray_menu_surface_kind: trayMenuSurfaceKind,
  tray_menu_surface_native: trayMenuSurfaceNative,
  quit_confirmation_visible: quitConfirmationVisible,
  manual_tray_evidence: manualTrayEvidence,
  manual_tray_evidence_ready: manualTrayEvidenceReady,
  product_identity_review: {
    logo_quality_reviewed: logoQualityReviewed,
    logo_future_office_control_reviewed: logoFutureOfficeControlReviewed,
    final_logo_user_surface_reviewed: finalLogoUserSurfaceReviewed,
    current_logo_screenshot_feedback_reviewed: currentLogoScreenshotFeedbackReviewed,
    rendered_size_logo_reviewed: renderedSizeLogoReviewed,
    rendered_logo_review_ready: renderedLogoReviewReady,
    catalog_identity_review_ready: catalogIdentityReviewReady,
    addin_identity_reviewed: addinIdentityReviewed,
    addin_title_icon_type_reviewed: addinTitleIconTypeReviewed,
    addin_installable_surface_reviewed: addinInstallableSurfaceReviewed,
    current_addin_screenshot_feedback_reviewed: currentAddinScreenshotFeedbackReviewed,
    word_first_run_identity_reviewed: wordFirstRunIdentityReviewed,
    excel_first_run_identity_reviewed: excelFirstRunIdentityReviewed,
    powerpoint_first_run_identity_reviewed: powerPointFirstRunIdentityReviewed,
    tray_product_polish_reviewed: trayProductPolishReviewed,
    tray_native_first_impression_reviewed: trayNativeFirstImpressionReviewed,
    tray_normal_windows_launch_reviewed: trayNormalWindowsLaunchReviewed,
    current_tray_screenshot_feedback_reviewed: currentTrayScreenshotFeedbackReviewed,
    current_screenshot_feedback_ready: currentScreenshotFeedbackReady,
    word_first_run_identity_ready: wordFirstRunIdentityReady,
    excel_first_run_identity_ready: excelFirstRunIdentityReady,
    powerpoint_first_run_identity_ready: powerPointFirstRunIdentityReady,
    word_runtime_evidence_ready: wordRuntimeEvidenceReady,
    excel_runtime_evidence_ready: excelRuntimeEvidenceReady,
    powerpoint_runtime_evidence_ready: powerPointRuntimeEvidenceReady,
    ready: productIdentityReviewReady
  },
  daemon_main_window: {
    reviewed: daemonMainWindowReviewed,
    compact_status_details_reviewed: daemonMainWindowCompactReviewed,
    three_column_layout_reviewed: daemonMainWindowThreeColumnReviewed,
    ready: daemonMainWindowReady
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
  word_runtime_evidence: wordRuntimeEvidence,
  word_runtime_evidence_ready: wordRuntimeEvidenceReady,
  excel_runtime_evidence: excelRuntimeEvidence,
  excel_runtime_evidence_ready: excelRuntimeEvidenceReady,
  powerpoint_runtime_evidence: powerPointRuntimeEvidence,
  powerpoint_runtime_evidence_ready: powerPointRuntimeEvidenceReady,
  office_tool_e2e: officeToolE2e,
  office_tool_e2e_ready: officeToolE2eReady,
  word_taskpane: {
    compact_top_block: wordCompactTopBlock,
    tools_permissions_merged: wordToolsPermissionsMerged,
    inline_settings: wordInlineSettings,
    server_protocol_row: wordServerProtocolRow,
    server_protocol_row_ready: wordServerProtocolReady,
    document_state: wordDocumentState,
    document_state_ready: wordDocumentStateReady,
    runtime_evidence: wordRuntimeEvidence,
    runtime_evidence_ready: wordRuntimeEvidenceReady,
    density_ready: wordTaskpaneDensityReady
  },
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
  powerpoint_taskpane: {
    compact_top_block: powerPointCompactTopBlock,
    tools_permissions_merged: powerPointToolsPermissionsMerged,
    inline_settings: powerPointInlineSettings,
    server_protocol_row: powerPointServerProtocolRow,
    server_protocol_row_ready: powerPointServerProtocolReady,
    document_state: powerPointDocumentState,
    document_state_ready: powerPointDocumentStateReady,
    runtime_evidence: powerPointRuntimeEvidence,
    runtime_evidence_ready: powerPointRuntimeEvidenceReady,
    density_ready: powerPointTaskpaneDensityReady
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

function observationFor(surface: string): string | undefined {
  const explicit = readOption(`--${surface.replaceAll('_', '-')}`);
  if (explicit) return explicit;
  const logoSurface = renderedLogoReviewSurface(surface);
  if (logoSurface) {
    const label = typeof logoSurface.label === 'string' ? logoSurface.label : surface.replaceAll('_', ' ');
    return `${productName} rendered logo review ${label}`;
  }
  if (manualTrayEvidenceReady && isTraySurface(surface)) return `${productName} manual tray evidence ${surface.replaceAll('_', ' ')}`;
  return undefined;
}

function screenshotPathFor(surface: string): string | undefined {
  return readOption(`--${surface.replaceAll('_', '-')}-screenshot`) ?? renderedLogoReviewScreenshotPath(surface) ?? manualTrayScreenshotPath(surface);
}

function screenshotMetadataFor(path: string | undefined): Record<string, unknown> {
  if (!path) return { ready: false, fresh: false, error: 'missing path' };
  const resolvedPath = resolve(path);
  try {
    const stats = statSync(resolvedPath);
    const recordedAtMs = Date.now();
    const ageMs = Math.max(0, recordedAtMs - stats.mtimeMs);
    const fresh = ageMs <= screenshotFreshnessWindowMs();
    return {
      path: resolvedPath,
      size_bytes: stats.size,
      mtime: stats.mtime.toISOString(),
      recorded_at: new Date(recordedAtMs).toISOString(),
      age_ms: ageMs,
      freshness_window_ms: screenshotFreshnessWindowMs(),
      fresh,
      ready: fresh
    };
  } catch (error) {
    return { path: resolvedPath, ready: false, fresh: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function screenshotFreshnessWindowMs(): number {
  const value = readOption('--screenshot-freshness-window-ms') ?? process.env.OFFICE_MCP_SCREENSHOT_FRESHNESS_WINDOW_MS;
  const parsed = value ? Number(value) : 30 * 60 * 1000;
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : 30 * 60 * 1000;
}

function manualTrayScreenshotPath(surface: string): string | undefined {
  if (!manualTrayEvidenceReady || !isTraySurface(surface)) return undefined;
  const paths = isRecord(manualTrayEvidence?.tray_surface_screenshot_paths) ? manualTrayEvidence.tray_surface_screenshot_paths : undefined;
  const path = paths?.[surface];
  return typeof path === 'string' ? path : undefined;
}

function isTraySurface(surface: string): boolean {
  return ['tray_icon', 'tray_native_menu', 'tray_tooltip', 'tray_quit_confirmation'].includes(surface);
}

function renderedLogoReviewScreenshotPath(surface: string): string | undefined {
  const logoSurface = renderedLogoReviewSurface(surface);
  if (typeof logoSurface?.screenshot_path === 'string') return logoSurface.screenshot_path;
  if (typeof renderedLogoReview?.sheet_path === 'string' && renderedLogoSurfaceSpecs().some(([key]) => key === surface)) return renderedLogoReview.sheet_path;
  return undefined;
}

function renderedLogoReviewSurface(surface: string): Record<string, unknown> | undefined {
  if (!renderedLogoReviewReady || !Array.isArray(renderedLogoReview?.surfaces)) return undefined;
  return renderedLogoReview.surfaces.filter(isRecord).find((item) => item.key === surface);
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
    && productCatalogTypeLooksReady(identity.type)
    && typeof identity.icon_url === 'string' && /\/assets\/icon-32\.png/.test(identity.icon_url)
    && typeof identity.high_resolution_icon_url === 'string' && /\/assets\/icon-80\.png/.test(identity.high_resolution_icon_url)
    && identity.manifest_ready === true;
}

function productCatalogTypeLooksReady(value: unknown): boolean {
  if (typeof value !== 'string') return false;
  return /local productivity automation control utility/i.test(value) && !/(add-in|task pane|developer tool|mcp server|protocol bridge|sample|debug|experimental|office-mcp-(word|excel|powerpoint))/i.test(value);
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
  if (!renderedLogoDesignReviewLooksReady(review.design_review)) return false;
  if (!renderedLogoAssetFingerprintsLookCurrent(review)) return false;
  if (typeof review.sheet_path !== 'string' || !screenshotFileLooksLikeImage(resolve(review.sheet_path))) return false;
  const surfaces = Array.isArray(review.surfaces) ? review.surfaces.filter(isRecord) : [];
  return renderedLogoSurfaceSpecs().every(([key, size]) => surfaces.some((surface) => surface.key === key && surface.rendered_size_px === size && surface.width === size && surface.height === size && surface.non_empty === true && surface.palette_ready === true && surface.expected_size_ready === true));
}

function renderedLogoAssetFingerprintsLookCurrent(review: Record<string, unknown>): boolean {
  if (typeof review.source_asset_sha256 !== 'string' || review.source_asset_sha256 !== sha256File(join(brandAssetRoot, 'brand-mark.svg'))) return false;
  const surfaces = Array.isArray(review.surfaces) ? review.surfaces.filter(isRecord) : [];
  return renderedLogoSurfaceSpecs().every(([key]) => {
    const surface = surfaces.find((item) => item.key === key);
    if (!surface || typeof surface.asset_sha256 !== 'string') return false;
    const assetName = typeof surface.asset_path === 'string' ? surface.asset_path.split(/[\\/]/).pop() : undefined;
    if (!assetName || !/^icon-\d+\.png$/.test(assetName)) return false;
    return surface.asset_sha256 === sha256File(join(brandAssetRoot, assetName));
  });
}

function sha256File(path: string): string | undefined {
  try {
    return createHash('sha256').update(readFileSync(path)).digest('hex');
  } catch (_error) {
    return undefined;
  }
}

function renderedLogoDesignReviewLooksReady(review: unknown): boolean {
  if (!isRecord(review) || review.ready !== true) return false;
  const rejectedReadings = Array.isArray(review.rejects_generic_readings) ? review.rejects_generic_readings : [];
  return typeof review.future_office_control_brief === 'string'
    && renderedLogoConceptPassLooksReady(review.concept_pass)
    && /future office control/i.test(review.future_office_control_brief)
    && /routing|operator|control/i.test(review.future_office_control_brief)
    && /without .*Office-owned app marks/i.test(review.future_office_control_brief)
    && typeof review.office_productivity_metaphor === 'string' && /document|pane|office/i.test(review.office_productivity_metaphor)
    && typeof review.user_control_metaphor === 'string' && /control|command|operator/i.test(review.user_control_metaphor)
    && typeof review.futuristic_maturity === 'string' && /mature|futuristic|desktop utility/i.test(review.futuristic_maturity)
    && typeof review.non_microsoft_distinction === 'string' && /Office logos/i.test(review.non_microsoft_distinction) && /Microsoft 365 gradients/i.test(review.non_microsoft_distinction) && /PowerPoint slide silhouettes/i.test(review.non_microsoft_distinction) && /Outlook envelope marks/i.test(review.non_microsoft_distinction) && /gear-only/i.test(review.non_microsoft_distinction)
    && ['settings', 'file', 'debug console', 'ai-only', 'microsoft office clone'].every((item) => rejectedReadings.includes(item));
}

function renderedLogoConceptPassLooksReady(conceptPass: unknown): boolean {
  if (!isRecord(conceptPass) || conceptPass.ready !== true || conceptPass.selected_direction !== 'Command Console Panes') return false;
  const minimum = typeof conceptPass.minimum_concepts_reviewed === 'number' ? conceptPass.minimum_concepts_reviewed : 3;
  const concepts = Array.isArray(conceptPass.concepts) ? conceptPass.concepts.filter(isRecord) : [];
  const rejectedPatterns = Array.isArray(conceptPass.rejected_patterns) ? conceptPass.rejected_patterns : [];
  return minimum >= 3
    && concepts.length >= minimum
    && concepts.some((concept) => concept.name === 'Command Console Panes' && concept.decision === 'selected' && typeof concept.rationale === 'string' && /office productivity, local routing, and deliberate user control/i.test(concept.rationale))
    && concepts.some((concept) => concept.name === 'Orbiting Document Hub' && concept.decision === 'rejected' && typeof concept.rationale === 'string' && /generic sync or cloud connector/i.test(concept.rationale))
    && concepts.some((concept) => concept.name === 'Shielded Automation Badge' && concept.decision === 'rejected' && typeof concept.rationale === 'string' && /endpoint protection software/i.test(concept.rationale))
    && ['gear-only settings mark', 'Office-like app tile', 'host-app color block', 'generic document thumbnail', 'terminal/debug glyph', 'AI sparkle motif'].every((item) => rejectedPatterns.includes(item));
}
function renderedLogoSurfaceSpecs(): Array<[string, number]> {
  return [
    ['logo_tray_size', 16],
    ['logo_ribbon_size', 32],
    ['logo_catalog_thumbnail', 80],
    ['logo_daemon_titlebar', 20],
    ['logo_installer_metadata', 256]
  ];
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

function readWordRuntimeEvidence(path: string): Record<string, unknown> {
  try {
    const report = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    const gates = Array.isArray(report.gates) ? report.gates.filter(isRecord) : [];
    const discovery = gates.find((gate) => gate.name === 'word.session_discovery' && gate.status === 'passed');
    const readSmoke = gates.find((gate) => gate.name === 'word.runtime_read_smoke' && gate.status === 'passed');
    const mutationSmoke = gates.find((gate) => gate.name === 'word.runtime_mutation_smoke' && gate.status === 'passed');
    const fullSmokePassed = ['word-core', 'word-formatting', 'word-review', 'word-resources', 'word-spec-args']
      .every((mode) => gates.some((gate) => gate.name === `word.full_smoke.${mode}` && gate.status === 'passed'));
    const comTrackedChangePassed = ['accept', 'reject']
      .every((action) => gates.some((gate) => gate.name === `word.tracked_change_com.${action}` && gate.status === 'passed'));
    const session = wordSessionFromDiscovery(discovery, report.session_id);
    const readDetails = isRecord(readSmoke?.details) ? readSmoke.details : undefined;
    const mutationDetails = isRecord(mutationSmoke?.details) ? mutationSmoke.details : undefined;
    const details = {
      session_id: report.session_id,
      available_tool_count: readDetails?.available_tool_count,
      available_tools: readDetails?.available_tools,
      paragraph_0_text_length: readDetails?.paragraph_0_text_length,
      document_text_length: readDetails?.document_text_length,
      find_count: mutationDetails?.find_count,
      full_smoke_passed: fullSmokePassed,
      com_tracked_change_passed: comTrackedChangePassed
    };
    return {
      path,
      ok: true,
      schema_version: report.schema_version,
      endpoint: report.endpoint,
      generated_at: report.generated_at,
      session,
      smoke_details: details,
      smoke_passed: Boolean(readSmoke && mutationSmoke && fullSmokePassed && comTrackedChangePassed),
      ready: wordRuntimeDetailsLookReady(session, details)
    };
  } catch (error) {
    return { path, ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function wordSessionFromDiscovery(discovery: Record<string, unknown> | undefined, reportSessionId: unknown): Record<string, unknown> | undefined {
  const discoveryDetails = isRecord(discovery?.details) ? discovery.details : undefined;
  const sessions = Array.isArray(discoveryDetails?.sessions) ? discoveryDetails.sessions.filter(isRecord) : [];
  const sessionId = typeof reportSessionId === 'string' ? reportSessionId : undefined;
  return sessions.find((session) => session.app === 'word' && (!sessionId || session.session_id === sessionId));
}

function wordRuntimeEvidenceLooksReady(evidence: Record<string, unknown> | undefined): boolean {
  if (!evidence) return false;
  return evidence.ok === true && evidence.schema_version === 1 && evidence.smoke_passed === true && evidence.ready === true;
}

function readPowerPointRuntimeEvidence(path: string): Record<string, unknown> {
  try {
    const report = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    const gates = Array.isArray(report.gates) ? report.gates.filter(isRecord) : [];
    const discovery = gates.find((gate) => gate.name === 'word.session_discovery' && gate.status === 'passed');
    const smoke = gates.find((gate) => gate.name === 'powerpoint.runtime_smoke' && gate.status === 'passed');
    const session = powerPointSessionFromDiscovery(discovery, smoke);
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
      ready: powerPointRuntimeDetailsLookReady(session, details)
    };
  } catch (error) {
    return { path, ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function powerPointSessionFromDiscovery(discovery: Record<string, unknown> | undefined, smoke: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  const smokeDetails = isRecord(smoke?.details) ? smoke.details : undefined;
  const smokeSessionId = typeof smokeDetails?.session_id === 'string' ? smokeDetails.session_id : undefined;
  const discoveryDetails = isRecord(discovery?.details) ? discovery.details : undefined;
  const sessions = Array.isArray(discoveryDetails?.sessions) ? discoveryDetails.sessions.filter(isRecord) : [];
  return sessions.find((session) => session.app === 'powerpoint' && (!smokeSessionId || session.session_id === smokeSessionId));
}

function powerPointRuntimeEvidenceLooksReady(evidence: Record<string, unknown> | undefined): boolean {
  if (!evidence) return false;
  return evidence.ok === true && evidence.schema_version === 1 && evidence.smoke_passed === true && evidence.ready === true;
}

function readOfficeToolE2eReport(host: 'Word' | 'Excel' | 'PowerPoint', path: string): Record<string, unknown> {
  try {
    const report = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    return {
      path,
      ok: true,
      host,
      schema_version: report.schema_version,
      kind: report.kind,
      report_host: report.host,
      passed: report.passed,
      addin_activation: report.addin_activation,
      lifecycle_counts: report.lifecycle_counts,
      advertised_tools: report.advertised_tools,
      session_available_tools: report.session_available_tools,
      executed_tools: report.executed_tools,
      tool_runs: report.tool_runs,
      ready: officeToolE2eDetailsLookReady(host, report)
    };
  } catch (error) {
    return { path, ok: false, host, error: error instanceof Error ? error.message : String(error), ready: false };
  }
}

function officeToolE2eReportLooksReady(report: Record<string, unknown> | undefined): boolean {
  return Boolean(report && report.ok === true && report.ready === true);
}

function officeToolE2eDetailsLookReady(host: 'Word' | 'Excel' | 'PowerPoint', report: Record<string, unknown>): boolean {
  if (report.schema_version !== 1 || report.kind !== 'office_tool_e2e_report' || report.host !== host || report.passed !== true) return false;
  if (!officeToolE2eLifecycleLooksReady(report.lifecycle_counts)) return false;
  if (!officeToolE2eActivationLooksReady(report.addin_activation)) return false;
  const advertisedTools = stringArray(report.advertised_tools);
  const sessionTools = stringArray(report.session_available_tools);
  const executedTools = stringArray(report.executed_tools);
  if (advertisedTools.length === 0 || !sameStrings(advertisedTools, sessionTools) || !sameStrings(advertisedTools, executedTools)) return false;
  if (!isRecord(report.daemon) || typeof report.daemon.endpoint !== 'string') return false;
  if (!isRecord(report.document) || typeof report.document.path !== 'string') return false;
  if (!isRecord(report.session) || typeof report.session.session_id !== 'string') return false;
  return officeToolRunsLookReady(advertisedTools, report.tool_runs);
}

function officeToolE2eActivationLooksReady(activation: unknown): boolean {
  return isRecord(activation)
    && activation.activated === true
    && typeof activation.skipped !== 'string';
}

function officeToolE2eLifecycleLooksReady(lifecycle: unknown): boolean {
  if (!isRecord(lifecycle)) return false;
  return ['start_daemon', 'list_tools', 'create_document', 'activate_addin', 'wait_for_session', 'cleanup_document', 'stop_daemon']
    .every((key) => lifecycle[key] === 1);
}

function officeToolRunsLookReady(advertisedTools: string[], runs: unknown): boolean {
  if (!Array.isArray(runs)) return false;
  const runRecords = runs.filter(isRecord);
  const runTools = runRecords.map((run) => run.tool).filter((tool): tool is string => typeof tool === 'string');
  return sameStrings(advertisedTools, runTools) && runRecords.every((run) => {
    const verifier = run.verifier;
    const expectationKeys = isRecord(verifier) && Array.isArray(verifier.expectation_keys) ? verifier.expectation_keys : [];
    return run.passed === true
      && typeof run.id === 'string'
      && typeof run.setup_action_count === 'number'
      && run.setup_action_count >= 1
      && isRecord(verifier)
      && (verifier.kind === 'direct-result' || verifier.kind === 'readback')
      && expectationKeys.length > 0;
  });
}

function stringArray(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === 'string') : [];
}

function exactToolCatalogReady(count: unknown, tools: unknown, expected: string[]): boolean {
  return count === expected.length && sameStrings(stringArray(tools), expected);
}

function sameStrings(left: string[], right: string[]): boolean {
  return JSON.stringify([...left].sort()) === JSON.stringify([...right].sort());
}

function powerPointRuntimeDetailsLookReady(session: Record<string, unknown> | undefined, details: Record<string, unknown> | undefined): boolean {
  if (!session || !details) return false;
  const document = isRecord(session.document) ? session.document : undefined;
  const host = isRecord(session.host) ? session.host : undefined;
  const exportSupported = details.export_supported === true && details.export_mime_type === 'application/pdf' && typeof details.export_size === 'number';
  const exportHostRejection = details.export_host_rejection === true;
  const tableSupported = details.table_supported === true && isRecord(details.add_table) && typeof details.add_table.shape_id === 'string' && isRecord(details.read_table);
  const tableHostRejection = details.table_host_rejection === true;
  const categoryProofs = isRecord(details.tool_category_proofs) ? details.tool_category_proofs : undefined;
  return session.app === 'powerpoint'
    && session.status === 'active'
    && typeof session.session_id === 'string'
    && typeof details.session_id === 'string'
    && session.session_id === details.session_id
    && typeof document?.title === 'string'
    && document.title.length > 0
    && host?.app === 'powerpoint'
    && exactToolCatalogReady(session.available_tool_count, details.available_tools, powerPointV1Tools)
    && details.mutation_proved === true
    && categoryProofs?.presentation === true
    && categoryProofs?.slides === true
    && categoryProofs?.layout === true
    && categoryProofs?.shapes === true
    && categoryProofs?.text === true
    && categoryProofs?.tables === true
    && isRecord(details.presentation_info)
    && isRecord(details.active_view)
    && isRecord(details.list_slides)
    && isRecord(details.add_slide)
    && typeof details.add_slide.slide_id === 'string'
    && isRecord(details.add_text_box)
    && isRecord(details.add_text_box.shape)
    && isRecord(details.list_shapes)
    && Array.isArray(details.list_shapes.shapes)
    && isRecord(details.read_text)
    && Array.isArray(details.read_text.items)
    && isRecord(details.replace_text)
    && Number(details.replace_text.replacements ?? 0) >= 1
    && isRecord(details.format_text)
    && details.format_text.formatted === true
    && isRecord(details.layout)
    && typeof details.layout.slide_id === 'string'
    && isRecord(details.list_layouts)
    && Array.isArray(details.list_layouts.masters)
    && (tableSupported || tableHostRejection)
    && (exportSupported || exportHostRejection);
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
    && exactToolCatalogReady(session.available_tool_count, details.available_tools, excelV1Tools)
    && details.marker_found === true
    && isRecord(details.workbook_info)
    && typeof details.sheet_list_count === 'number'
    && isRecord(details.updated_sheet)
    && isRecord(details.deleted_sheet) && details.deleted_sheet.deleted === true
    && isRecord(details.used_range)
    && isRecord(details.find_replace)
    && isRecord(details.clear)
    && isRecord(details.write) && details.write.wrote_values === true
    && isRecord(details.formula) && details.formula.wrote_formula === true
    && isRecord(details.format) && details.format.formatted === true
    && isRecord(details.table) && typeof details.table.table === 'string'
    && isRecord(details.table_update)
    && isRecord(details.sort) && details.sort.sorted === true
    && isRecord(details.filter) && details.filter.filtered === true
    && isRecord(details.chart) && typeof details.chart.chart === 'string'
    && isRecord(details.chart_update) && details.chart_update.updated === true
    && isRecord(details.pivot_table) && typeof details.pivot_table.pivot_table === 'string'
    && isRecord(details.pivot_update) && details.pivot_update.refreshed === true
    && isRecord(details.sheet) && details.sheet.activated === true;
}

function wordRuntimeDetailsLookReady(session: Record<string, unknown> | undefined, details: Record<string, unknown> | undefined): boolean {
  if (!session || !details) return false;
  const document = isRecord(session.document) ? session.document : undefined;
  const host = isRecord(session.host) ? session.host : undefined;
  return session.app === 'word'
    && session.status === 'active'
    && typeof session.session_id === 'string'
    && typeof document?.title === 'string'
    && document.title.length > 0
    && host?.app === 'word'
    && exactToolCatalogReady(session.available_tool_count, details.available_tools, wordV1Tools)
    && details.available_tool_count === wordV1Tools.length
    && Number(details.paragraph_0_text_length ?? 0) > 0
    && Number(details.document_text_length ?? 0) > 0
    && Number(details.find_count ?? 0) >= 1
    && details.full_smoke_passed === true
    && details.com_tracked_change_passed === true;
}

function readManualTrayEvidence(path: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    return { path, ok: true, ...parsed };
  } catch (error) {
    return { path, ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function readCatalogIdentityReview(path: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;
    return { path, ok: true, ...parsed };
  } catch (error) {
    return { path, ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

function catalogIdentityReviewLooksReady(review: Record<string, unknown> | undefined): boolean {
  if (!review) return false;
  if (review.ok !== true || review.schema_version !== 1 || review.kind !== 'catalog_identity_review' || review.product_name !== productName || review.ready !== true) return false;
  if (!productCatalogTypeLooksReady(review.catalog_type)) return false;
  if (typeof review.shared_origin !== 'string' || !/^https:\/\/localhost:\d+$/.test(review.shared_origin)) return false;
  const hosts = isRecord(review.hosts) ? review.hosts : {};
  return ['word', 'excel', 'powerpoint'].every((host) => catalogIdentityHostLooksReady(hosts[host]));
}

function catalogIdentityHostLooksReady(host: unknown): boolean {
  if (!isRecord(host)) return false;
  return host.ready === true
    && host.display_name === productName
    && host.provider === productName
    && typeof host.description === 'string' && /local productivity automation control utility/i.test(host.description)
    && host.group_label === productName
    && host.command_label === 'Open Control Panel'
    && typeof host.icon_url === 'string' && /\/assets\/icon-32\.png$/.test(host.icon_url)
    && typeof host.high_resolution_icon_url === 'string' && /\/assets\/icon-80\.png$/.test(host.high_resolution_icon_url);
}

function catalogIdentityHostValue(host: string, field: string): string | undefined {
  if (!catalogIdentityReviewReady || !isRecord(catalogIdentityReview?.hosts)) return undefined;
  const hostReview = catalogIdentityReview.hosts[host];
  if (!isRecord(hostReview)) return undefined;
  const value = hostReview[field];
  return typeof value === 'string' ? value : undefined;
}

function manualTrayEvidenceLooksReady(evidence: Record<string, unknown> | undefined): boolean {
  if (!evidence) return false;
  if (evidence.ok !== true || evidence.schema_version !== 1 || evidence.kind !== 'tray_manual_evidence' || evidence.platform !== 'win32' || evidence.passed !== true) return false;
  if (!manualTrayInteractionLooksReady(evidence) || evidence.tray_menu_surface_native !== true || evidence.tray_menu_surface_kind !== 'native' || evidence.show_ui_opened !== true) return false;
  if (typeof evidence.observed_tooltip !== 'string' || !/^Office MCP Control - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(evidence.observed_tooltip)) return false;
  if (typeof evidence.screenshot_path !== 'string' || !screenshotFileLooksLikeImage(resolve(evidence.screenshot_path))) return false;
  const items = Array.isArray(evidence.observed_menu_items) ? evidence.observed_menu_items.filter((item): item is string => typeof item === 'string') : [];
  return ['Status:', 'Clients:', 'Documents:', 'Show Office MCP Control', 'Quit Office MCP Control'].every((expected) => items.some((item) => item.includes(expected)))
    && evidence.daemon_context_ready === true
    && daemonContextLooksReady(isRecord(evidence.daemon_context) ? evidence.daemon_context : undefined);
}

function manualTrayInteractionLooksReady(evidence: Record<string, unknown>): boolean {
  return evidence.visible_icon === true
    && evidence.right_click_menu === true
    && evidence.menu_opened_from_tray_icon === true
    && evidence.native_menu_appearance_reviewed === true
    && evidence.menu_anchored_to_tray_icon === true
    && evidence.os_native_menu_behavior_reviewed === true
    && evidence.keyboard_menu_access_reviewed === true
    && evidence.native_quit_confirmation_reviewed === true
    && evidence.native_tray_interaction_ready === true;
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
  const tooltipReady = typeof snapshot.tooltip === 'string' && /^Office MCP Control - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(snapshot.tooltip);
  const menuItems = Array.isArray(snapshot.menu_items) ? snapshot.menu_items.filter((item): item is string => typeof item === 'string') : [];
  const menuReady = ['Status:', 'Clients:', 'Documents:', 'Show Office MCP Control', 'Quit Office MCP Control']
    .every((expected) => menuItems.some((item) => item.includes(expected)));
  return tooltipReady && menuReady && structuredTrayMenuLooksReady(snapshot.menu);
}

function structuredTrayMenuLooksReady(menu: unknown): boolean {
  if (!Array.isArray(menu)) return false;
  const expected = [
    { kind: 'read_only', enabled: false, label: /^Status: (Up|Degraded|Down)$/ },
    { kind: 'read_only', enabled: false, label: /^Clients: \d+$/ },
    { kind: 'read_only', enabled: false, label: /^Documents: \d+$/ },
    { kind: 'separator', enabled: false, label: /^---$/ },
    { kind: 'action', enabled: true, label: /^Show Office MCP Control$/, action: 'show_ui' },
    { kind: 'action', enabled: true, label: /^Quit Office MCP Control$/, action: 'quit' }
  ];
  if (menu.length !== expected.length) return false;
  return expected.every((rule, index) => {
    const item = menu[index];
    if (!isRecord(item)) return false;
    if (item.kind !== rule.kind || item.enabled !== rule.enabled) return false;
    if (typeof item.label !== 'string' || !rule.label.test(item.label)) return false;
    return !('action' in rule) || item.action === rule.action;
  });
}
