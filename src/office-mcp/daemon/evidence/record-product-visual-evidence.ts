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

const requiredSurfaces = [
  'word_ribbon_command',
  'word_catalog_entry',
  'word_taskpane_title',
  'excel_ribbon_command',
  'excel_catalog_entry',
  'excel_taskpane_title',
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
const wordCatalogProvider = readOption('--word-catalog-provider');
const wordCatalogDescription = readOption('--word-catalog-description');
const wordCatalogType = readOption('--word-catalog-type');
const excelCatalogProvider = readOption('--excel-catalog-provider');
const excelCatalogDescription = readOption('--excel-catalog-description');
const excelCatalogType = readOption('--excel-catalog-type');
const excelCompactTopBlock = booleanFlag('--excel-compact-top-block');
const excelToolsPermissionsMerged = booleanFlag('--excel-tools-permissions-merged');
const excelInlineSettings = booleanFlag('--excel-inline-settings');
const excelServerProtocolRow = readOption('--excel-server-protocol-row');
const excelDocumentState = readOption('--excel-document-state');
const daemonContext = daemonBin ? readDaemonContext(resolve(daemonBin)) : undefined;
const daemonContextReady = daemonContextLooksReady(daemonContext);
const renderedLogoReview = renderedLogoReviewPath ? readRenderedLogoReview(resolve(renderedLogoReviewPath)) : undefined;
const renderedLogoReviewReady = renderedLogoReviewLooksReady(renderedLogoReview);

const productTextReady = requiredSurfaces.filter((surface) => surface !== 'tray_tooltip').every((surface) => typeof observations[surface] === 'string' && (observations[surface] as string).includes(productName));
const allScreenshotsExist = Object.values(screenshotsExist).every(Boolean);
const trayTooltipReady = typeof trayTooltip === 'string' && /^Office MCP - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(trayTooltip);
const catalogTypeReady = typeof catalogType === 'string' && /local productivity automation control utility/i.test(catalogType);
const wordFirstRunIdentityReady = wordFirstRunIdentityReviewed && catalogIdentityLooksReady(wordCatalogProvider, wordCatalogDescription, wordCatalogType);
const excelFirstRunIdentityReady = excelFirstRunIdentityReviewed && catalogIdentityLooksReady(excelCatalogProvider, excelCatalogDescription, excelCatalogType);
const excelServerProtocolReady = typeof excelServerProtocolRow === 'string' && /^Server .+ \/ Protocol .+$/.test(excelServerProtocolRow);
const excelDocumentStateReady = typeof excelDocumentState === 'string' && /^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(excelDocumentState) && !/unknown/i.test(excelDocumentState);
const excelTaskpaneDensityReady = excelCompactTopBlock && excelToolsPermissionsMerged && excelInlineSettings && excelServerProtocolReady && excelDocumentStateReady;
const productIdentityReviewReady = logoQualityReviewed && renderedSizeLogoReviewed && renderedLogoReviewReady && addinIdentityReviewed && wordFirstRunIdentityReady && excelFirstRunIdentityReady && trayProductPolishReviewed;
const passed = productTextReady && allScreenshotsExist && trayTooltipReady && catalogTypeReady && catalogIconVisible && trayMenuNative && trayIconVisible && quitConfirmationVisible && excelTaskpaneDensityReady && productIdentityReviewReady && renderedLogoReviewReady && daemonContextReady;

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
  product_identity_review: {
    logo_quality_reviewed: logoQualityReviewed,
    rendered_size_logo_reviewed: renderedSizeLogoReviewed,
    rendered_logo_review_ready: renderedLogoReviewReady,
    addin_identity_reviewed: addinIdentityReviewed,
    word_first_run_identity_reviewed: wordFirstRunIdentityReviewed,
    excel_first_run_identity_reviewed: excelFirstRunIdentityReviewed,
    tray_product_polish_reviewed: trayProductPolishReviewed,
    word_first_run_identity_ready: wordFirstRunIdentityReady,
    excel_first_run_identity_ready: excelFirstRunIdentityReady,
    ready: productIdentityReviewReady
  },
  first_run_identity: {
    word: {
      provider: wordCatalogProvider,
      description: wordCatalogDescription,
      type: wordCatalogType,
      ready: wordFirstRunIdentityReady
    },
    excel: {
      provider: excelCatalogProvider,
      description: excelCatalogDescription,
      type: excelCatalogType,
      ready: excelFirstRunIdentityReady
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

function catalogIdentityLooksReady(provider: string | undefined, description: string | undefined, type: string | undefined): boolean {
  return typeof provider === 'string' && provider.includes(productName)
    && typeof description === 'string' && /local/i.test(description) && /(productivity|office)/i.test(description) && /(automation|control)/i.test(description)
    && typeof type === 'string' && /local productivity automation control utility/i.test(type);
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
