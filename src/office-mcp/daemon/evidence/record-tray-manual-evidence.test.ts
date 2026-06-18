import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';

const TSX = './node_modules/tsx/dist/cli.mjs';
const RECORDER = resolve(process.cwd(), 'record-tray-manual-evidence.ts');

test('manual tray evidence recorder requires product tooltip', () => {
  withTrayScreenshot((dir, screenshotPath) => {
    const output = join(dir, 'tray-evidence.json');
    const passing = runRecorder(output, screenshotPath, '--tooltip', 'Office MCP - Up - 0 clients - 0 documents');
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.tooltip_product_ready, true);
    assert.equal(evidence.passed, true);

    const missingTooltip = runRecorder(join(dir, 'missing-tooltip.json'), screenshotPath);
    assert.notEqual(missingTooltip.status, 0);
    const failed = JSON.parse(outputText(missingTooltip.stdout)) as Record<string, unknown>;
    assert.equal(failed.tooltip_product_ready, false);
    assert.equal(failed.passed, false);
  });
});

function runRecorder(output: string, screenshotPath: string, ...extra: string[]): ReturnType<typeof spawnSync> {
  return spawnSync(process.execPath, [
    TSX,
    RECORDER,
    '--output', output,
    '--visible-icon', 'true',
    '--right-click-menu', 'true',
    '--show-ui-opened', 'true',
    '--screenshot-path', screenshotPath,
    '--menu-item', 'Status: Up',
    '--menu-item', 'Clients: 0',
    '--menu-item', 'Documents: 0',
    '--menu-item', 'Show Office MCP',
    '--menu-item', 'Quit Office MCP',
    ...extra
  ], { cwd: process.cwd(), encoding: 'utf8' });
}

function outputText(value: string | Buffer): string {
  return typeof value === 'string' ? value : value.toString('utf8');
}

function withTrayScreenshot(callback: (dir: string, screenshotPath: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-tray-recorder-test-'));
  try {
    const screenshotPath = join(dir, 'tray.png');
    writeFileSync(screenshotPath, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));
    callback(dir, screenshotPath);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
