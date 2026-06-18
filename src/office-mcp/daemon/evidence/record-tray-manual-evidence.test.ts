import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import test from 'node:test';
import { tinyPng } from './image-evidence.js';

const TSX = './node_modules/tsx/dist/cli.mjs';
const RECORDER = resolve(process.cwd(), 'record-tray-manual-evidence.ts');

test('manual tray evidence recorder requires product tooltip', () => {
  withTrayScreenshot((dir, screenshotPath) => {
    const daemonBin = writeFakeDaemon(dir);
    const output = join(dir, 'tray-evidence.json');
    const passing = runRecorder(output, screenshotPath, '--tooltip', 'Office MCP - Up - 0 clients - 0 documents', '--daemon-bin', daemonBin);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.tooltip_product_ready, true);
    assert.equal(evidence.daemon_context_ready, true);
    assert.equal((evidence.tray_surface_screenshots_exist as Record<string, boolean>).tray_icon, true);
    assert.equal((evidence.tray_surface_screenshots_exist as Record<string, boolean>).tray_native_menu, true);
    assert.equal((evidence.tray_surface_screenshots_exist as Record<string, boolean>).tray_tooltip, true);
    assert.equal((evidence.tray_surface_screenshots_exist as Record<string, boolean>).tray_quit_confirmation, true);
    assert.equal(evidence.passed, true);

    const missingTooltip = runRecorder(join(dir, 'missing-tooltip.json'), screenshotPath, '--daemon-bin', daemonBin);
    assert.notEqual(missingTooltip.status, 0);
    const failed = JSON.parse(outputText(missingTooltip.stdout)) as Record<string, unknown>;
    assert.equal(failed.tooltip_product_ready, false);
    assert.equal(failed.passed, false);
  });
});

test('manual tray evidence recorder requires daemon context before passing', () => {
  withTrayScreenshot((dir, screenshotPath) => {
    const output = join(dir, 'missing-daemon-context.json');
    const result = runRecorder(output, screenshotPath, '--tooltip', 'Office MCP - Up - 0 clients - 0 documents');
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('manual tray evidence recorder requires native tray menu review', () => {
  withTrayScreenshot((dir, screenshotPath) => {
    const daemonBin = writeFakeDaemon(dir);
    const output = join(dir, 'missing-native-menu-review.json');
    const result = runRecorder(
      output,
      screenshotPath,
      '--tooltip', 'Office MCP - Up - 0 clients - 0 documents',
      '--daemon-bin', daemonBin,
      '--skip-native-menu-review-flags'
    );
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.menu_opened_from_tray_icon, false);
    assert.equal(evidence.native_menu_appearance_reviewed, false);
    assert.equal(evidence.passed, false);
  });
});

test('manual tray evidence recorder requires tray probe live state', () => {
  withTrayScreenshot((dir, screenshotPath) => {
    const daemonBin = writeFakeDaemon(dir, false);
    const output = join(dir, 'missing-live-state.json');
    const result = runRecorder(output, screenshotPath, '--tooltip', 'Office MCP - Up - 0 clients - 0 documents', '--daemon-bin', daemonBin);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('manual tray evidence recorder requires live tray menu snapshot', () => {
  withTrayScreenshot((dir, screenshotPath) => {
    const daemonBin = writeFakeDaemon(dir, true, ['Status: Up', 'Clients: 0']);
    const output = join(dir, 'missing-menu-items.json');
    const result = runRecorder(output, screenshotPath, '--tooltip', 'Office MCP - Up - 0 clients - 0 documents', '--daemon-bin', daemonBin);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('manual tray evidence recorder rejects truncated screenshots', () => {
  withTrayScreenshot((dir, screenshotPath) => {
    writeFileSync(screenshotPath, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));
    const daemonBin = writeFakeDaemon(dir);
    const output = join(dir, 'truncated-screenshot.json');
    const result = runRecorder(output, screenshotPath, '--tooltip', 'Office MCP - Up - 0 clients - 0 documents', '--daemon-bin', daemonBin);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.screenshot_exists, false);
    assert.equal(evidence.passed, false);
  });
});

function runRecorder(output: string, screenshotPath: string, ...extra: string[]): ReturnType<typeof spawnSync> {
  const dir = dirname(output);
  const trayIconScreenshot = writeSurfaceScreenshot(dir, 'tray-icon.png');
  const trayNativeMenuScreenshot = writeSurfaceScreenshot(dir, 'tray-native-menu.png');
  const trayTooltipScreenshot = writeSurfaceScreenshot(dir, 'tray-tooltip.png');
  const trayQuitConfirmationScreenshot = writeSurfaceScreenshot(dir, 'tray-quit-confirmation.png');
  const skipNativeMenuReviewFlags = extra.includes('--skip-native-menu-review-flags');
  const filteredExtra = extra.filter((item) => item !== '--skip-native-menu-review-flags');
  const reviewArgs = skipNativeMenuReviewFlags ? [] : [
    '--menu-opened-from-tray-icon', 'true',
    '--native-menu-appearance-reviewed', 'true'
  ];
  return spawnSync(process.execPath, [
    TSX,
    RECORDER,
    '--output', output,
    '--visible-icon', 'true',
    '--right-click-menu', 'true',
    ...reviewArgs,
    '--show-ui-opened', 'true',
    '--screenshot-path', screenshotPath,
    '--tray-icon-screenshot', trayIconScreenshot,
    '--tray-native-menu-screenshot', trayNativeMenuScreenshot,
    '--tray-tooltip-screenshot', trayTooltipScreenshot,
    '--tray-quit-confirmation-screenshot', trayQuitConfirmationScreenshot,
    '--menu-item', 'Status: Up',
    '--menu-item', 'Clients: 0',
    '--menu-item', 'Documents: 0',
    '--menu-item', 'Show Office MCP',
    '--menu-item', 'Quit Office MCP',
    ...filteredExtra
  ], { cwd: process.cwd(), encoding: 'utf8' });
}

function writeSurfaceScreenshot(dir: string, name: string): string {
  const path = join(dir, name);
  writeFileSync(path, tinyPng());
  return path;
}

function outputText(value: string | Buffer): string {
  return typeof value === 'string' ? value : value.toString('utf8');
}

function withTrayScreenshot(callback: (dir: string, screenshotPath: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-tray-recorder-test-'));
  try {
    const screenshotPath = join(dir, 'tray.png');
    writeFileSync(screenshotPath, tinyPng());
    callback(dir, screenshotPath);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function writeFakeDaemon(dir: string, stateFetchOk = true, menuItems = ['Status: Up', 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP', 'Quit Office MCP']): string {
  const daemonBin = join(dir, process.platform === 'win32' ? 'daemon.cmd' : 'daemon.sh');
  writeFileSync(daemonBin, fakeDaemonScript(stateFetchOk, menuItems));
  chmodSync(daemonBin, 0o755);
  return daemonBin;
}

function fakeDaemonScript(stateFetchOk: boolean, menuItems: string[]): string {
  const status = JSON.stringify({ running: true, uiUrl: 'https://localhost:8765/ui/' });
  const trayProbe = JSON.stringify({ native_host: true, state_fetch_ok: stateFetchOk, snapshot: { tooltip: 'Office MCP - Up - 0 clients - 0 documents', menu_items: menuItems } });
  if (process.platform === 'win32') {
    return `@echo off\r\nif "%1"=="daemon" echo ${status}\r\nif "%1"=="tray" echo ${trayProbe}\r\n`;
  }
  return `#!/bin/sh\nif [ "$1" = "daemon" ]; then printf '%s\\n' '${status}'; fi\nif [ "$1" = "tray" ]; then printf '%s\\n' '${trayProbe}'; fi\n`;
}
