import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/product-visual-evidence.json'));
const tester = readOption('--tester') ?? process.env.USERNAME ?? process.env.USER ?? 'unknown';
const notes = readOption('--notes');
const daemonBin = readOption('--daemon-bin');

const requiredSurfaces = [
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

const observations = Object.fromEntries(requiredSurfaces.map((surface) => [surface, readOption(`--${surface.replaceAll('_', '-')}`)]));
const screenshotPaths = Object.fromEntries(requiredSurfaces.map((surface) => [surface, screenshotPathFor(surface)]));
const screenshotsExist = Object.fromEntries(
  requiredSurfaces.map((surface) => [surface, typeof screenshotPaths[surface] === 'string' && existsSync(resolve(screenshotPaths[surface] as string))])
);
const productName = readOption('--product-name') ?? 'Office MCP Control';
const trayTooltip = readOption('--tray-tooltip');
const catalogType = readOption('--catalog-type');
const catalogIconVisible = booleanFlag('--catalog-icon-visible');
const trayMenuNative = booleanFlag('--tray-menu-native');
const trayIconVisible = booleanFlag('--tray-icon-visible');
const quitConfirmationVisible = booleanFlag('--quit-confirmation-visible');
const excelCompactTopBlock = booleanFlag('--excel-compact-top-block');
const excelToolsPermissionsMerged = booleanFlag('--excel-tools-permissions-merged');
const excelInlineSettings = booleanFlag('--excel-inline-settings');
const excelServerProtocolRow = readOption('--excel-server-protocol-row');
const excelDocumentState = readOption('--excel-document-state');
const daemonContext = daemonBin ? readDaemonContext(resolve(daemonBin)) : undefined;
const daemonContextReady = daemonContextLooksReady(daemonContext);

const productTextReady = requiredSurfaces.filter((surface) => surface !== 'tray_tooltip').every((surface) => typeof observations[surface] === 'string' && (observations[surface] as string).includes(productName));
const allScreenshotsExist = Object.values(screenshotsExist).every(Boolean);
const trayTooltipReady = typeof trayTooltip === 'string' && /^Office MCP - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(trayTooltip);
const catalogTypeReady = typeof catalogType === 'string' && /local productivity automation control utility/i.test(catalogType);
const excelServerProtocolReady = typeof excelServerProtocolRow === 'string' && /^Server .+ \/ Protocol .+$/.test(excelServerProtocolRow);
const excelDocumentStateReady = typeof excelDocumentState === 'string' && /^(Editable|Editable, unsaved changes|Read-only|Protected.*)$/i.test(excelDocumentState) && !/unknown/i.test(excelDocumentState);
const excelTaskpaneDensityReady = excelCompactTopBlock && excelToolsPermissionsMerged && excelInlineSettings && excelServerProtocolReady && excelDocumentStateReady;
const passed = productTextReady && allScreenshotsExist && trayTooltipReady && catalogTypeReady && catalogIconVisible && trayMenuNative && trayIconVisible && quitConfirmationVisible && excelTaskpaneDensityReady && daemonContextReady;

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
    && isRecord(trayProbe) && trayProbe.ok === true && trayProbe.native_host === true && trayProbe.state_fetch_ok === true;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
