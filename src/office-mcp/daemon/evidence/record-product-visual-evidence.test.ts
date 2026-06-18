import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';

const TSX = './node_modules/tsx/dist/cli.mjs';
const RECORDER = resolve(process.cwd(), 'record-product-visual-evidence.ts');
const SURFACES = [
  'word-ribbon-command',
  'word-catalog-entry',
  'word-taskpane-title',
  'excel-ribbon-command',
  'excel-catalog-entry',
  'excel-taskpane-title',
  'tray-icon',
  'tray-native-menu',
  'tray-tooltip',
  'tray-quit-confirmation'
];

test('product visual evidence recorder requires all product surfaces', () => {
  withScreenshots((dir, screenshots) => {
    const output = join(dir, 'product-visual-evidence.json');
    const passing = runRecorder(output, screenshots);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.kind, 'product_visual_evidence');
    assert.equal(evidence.product_text_ready, true);
    assert.equal(evidence.catalog_type_ready, true);
    assert.equal(evidence.tray_tooltip_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal(evidence.passed, true);

    const missingScreenshots = { ...screenshots };
    missingScreenshots['tray-native-menu'] = join(dir, 'missing-tray-menu.png');
    const missingTray = runRecorder(join(dir, 'missing-tray.json'), missingScreenshots);
    assert.notEqual(missingTray.status, 0);
    const failed = JSON.parse(outputText(missingTray.stdout)) as Record<string, unknown>;
    assert.equal(failed.passed, false);
    assert.equal((failed.screenshots_exist as Record<string, unknown>).tray_native_menu, false);
  });
});

function runRecorder(output: string, screenshots: Record<string, string>, ...extra: string[]): ReturnType<typeof spawnSync> {
  const args = [
    TSX,
    RECORDER,
    '--output', output,
    '--catalog-type', 'Local productivity automation control utility',
    '--catalog-icon-visible', 'true',
    '--tray-tooltip', 'Office MCP - Up - 0 clients - 0 documents',
    '--tray-icon-visible', 'true',
    '--tray-menu-native', 'true',
    '--quit-confirmation-visible', 'true',
    '--excel-compact-top-block', 'true',
    '--excel-tools-permissions-merged', 'true',
    '--excel-inline-settings', 'true',
    '--excel-server-protocol-row', 'Server 0.1.0 / Protocol 1.0',
    '--excel-document-state', 'Editable'
  ];
  for (const surface of SURFACES) {
    args.push(`--${surface}`, `Office MCP Control ${surface}`);
    args.push(`--${surface}-screenshot`, screenshots[surface]);
  }
  args.push(...extra);
  return spawnSync(process.execPath, args, { cwd: process.cwd(), encoding: 'utf8' });
}

function withScreenshots(callback: (dir: string, screenshots: Record<string, string>) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-product-visual-test-'));
  try {
    const screenshots: Record<string, string> = {};
    for (const surface of SURFACES) {
      const path = join(dir, `${surface}.png`);
      writeFileSync(path, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));
      screenshots[surface] = path;
    }
    callback(dir, screenshots);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function outputText(value: string | Buffer): string {
  return typeof value === 'string' ? value : value.toString('utf8');
}