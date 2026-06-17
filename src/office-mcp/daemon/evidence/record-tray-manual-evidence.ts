import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const evidenceRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(evidenceRoot, '../../../..');
const outputPath = resolve(readOption('--output') ?? join(repoRoot, 'artifacts/tray-manual-evidence.json'));
const tester = readOption('--tester') ?? process.env.USERNAME ?? process.env.USER ?? 'unknown';
const screenshotPath = readOption('--screenshot-path');
const notes = readOption('--notes');

const visibleIcon = booleanFlag('--visible-icon');
const rightClickMenu = booleanFlag('--right-click-menu');
const showUiOpened = booleanFlag('--show-ui-opened');
const expectedItems = ['Status:', 'Clients:', 'Documents:', 'Show Office MCP', 'Quit Office MCP'];
const observedMenuItems = readRepeatedOption('--menu-item');
const menuContainsRequiredItems = expectedItems.every((expected) =>
  observedMenuItems.some((item) => item.includes(expected))
);
const screenshotExists = screenshotPath ? existsSync(resolve(screenshotPath)) : false;
const passed = visibleIcon && rightClickMenu && showUiOpened && menuContainsRequiredItems && screenshotExists;

const evidence = {
  schema_version: 1,
  kind: 'tray_manual_evidence',
  recorded_at: new Date().toISOString(),
  tester,
  platform: process.platform,
  visible_icon: visibleIcon,
  right_click_menu: rightClickMenu,
  show_ui_opened: showUiOpened,
  observed_menu_items: observedMenuItems,
  expected_menu_items: expectedItems,
  menu_contains_required_items: menuContainsRequiredItems,
  screenshot_path: screenshotPath ? resolve(screenshotPath) : undefined,
  screenshot_exists: screenshotExists,
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
