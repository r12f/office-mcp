import { mkdirSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { screenshotFileLooksLikeImage } from './image-evidence.js';

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/tray-manual-evidence.json'));
const tester = readOption('--tester') ?? process.env.USERNAME ?? process.env.USER ?? 'unknown';
const screenshotPath = readOption('--screenshot-path');
const traySurfaceScreenshots = traySurfaceScreenshotPaths();
const notes = readOption('--notes');
const observedTooltip = readOption('--tooltip');
const daemonBin = readOption('--daemon-bin');

const visibleIcon = booleanFlag('--visible-icon');
const rightClickMenu = booleanFlag('--right-click-menu');
const menuOpenedFromTrayIcon = booleanFlag('--menu-opened-from-tray-icon');
const nativeMenuAppearanceReviewed = booleanFlag('--native-menu-appearance-reviewed');
const showUiOpened = booleanFlag('--show-ui-opened');
const expectedItems = ['Status:', 'Clients:', 'Documents:', 'Show Office MCP', 'Quit Office MCP'];
const observedMenuItems = readRepeatedOption('--menu-item');
const menuContainsRequiredItems = expectedItems.every((expected) =>
  observedMenuItems.some((item) => item.includes(expected))
);
const screenshotExists = screenshotPath ? screenshotFileLooksLikeImage(resolve(screenshotPath)) : false;
const traySurfaceScreenshotsExist = Object.fromEntries(
  Object.entries(traySurfaceScreenshots).map(([surface, path]) => [surface, typeof path === 'string' && screenshotFileLooksLikeImage(resolve(path))])
);
const tooltipLooksProductReady = typeof observedTooltip === 'string' && /^Office MCP - (Up|Degraded|Down) - \d+ clients - \d+ documents$/.test(observedTooltip);
const daemonContext = daemonBin ? readDaemonContext(resolve(daemonBin)) : undefined;
const daemonContextReady = daemonContextLooksReady(daemonContext);
const passed = visibleIcon && rightClickMenu && menuOpenedFromTrayIcon && nativeMenuAppearanceReviewed && showUiOpened && menuContainsRequiredItems && tooltipLooksProductReady && screenshotExists && daemonContextReady;

const evidence = {
  schema_version: 1,
  kind: 'tray_manual_evidence',
  recorded_at: new Date().toISOString(),
  tester,
  platform: process.platform,
  visible_icon: visibleIcon,
  right_click_menu: rightClickMenu,
  menu_opened_from_tray_icon: menuOpenedFromTrayIcon,
  native_menu_appearance_reviewed: nativeMenuAppearanceReviewed,
  show_ui_opened: showUiOpened,
  observed_menu_items: observedMenuItems,
  observed_tooltip: observedTooltip,
  expected_menu_items: expectedItems,
  menu_contains_required_items: menuContainsRequiredItems,
  tooltip_product_ready: tooltipLooksProductReady,
  screenshot_path: screenshotPath ? resolve(screenshotPath) : undefined,
  screenshot_exists: screenshotExists,
  tray_surface_screenshot_paths: traySurfaceScreenshots,
  tray_surface_screenshots_exist: traySurfaceScreenshotsExist,
  daemon_context: daemonContext,
  daemon_context_ready: daemonContextReady,
  notes,
  passed
};

mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, JSON.stringify(evidence, null, 2));
console.log(JSON.stringify(evidence, null, 2));
if (!passed) process.exit(1);

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

function readRepeatedOption(name: string): string[] {
  const values: string[] = [];
  for (let index = 0; index < process.argv.length; index += 1) {
    if (process.argv[index] === name && process.argv[index + 1]) values.push(process.argv[index + 1]);
  }
  return values;
}

function traySurfaceScreenshotPaths(): Record<string, string | undefined> {
  return {
    tray_icon: normalizedOptionPath('--tray-icon-screenshot'),
    tray_native_menu: normalizedOptionPath('--tray-native-menu-screenshot'),
    tray_tooltip: normalizedOptionPath('--tray-tooltip-screenshot'),
    tray_quit_confirmation: normalizedOptionPath('--tray-quit-confirmation-screenshot')
  };
}

function normalizedOptionPath(name: string): string | undefined {
  const value = readOption(name);
  return value ? resolve(value) : undefined;
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
  const menuReady = expectedItems.every((expected) => menuItems.some((item) => item.includes(expected)));
  return tooltipReady && menuReady;
}
