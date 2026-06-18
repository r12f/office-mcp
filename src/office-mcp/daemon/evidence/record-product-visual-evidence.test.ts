import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
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
    const daemonBin = writeFakeDaemon(dir);
    const output = join(dir, 'product-visual-evidence.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.kind, 'product_visual_evidence');
    assert.equal(evidence.product_text_ready, true);
    assert.equal(evidence.catalog_type_ready, true);
    assert.equal(evidence.tray_tooltip_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal(evidence.daemon_context_ready, true);
    assert.equal(evidence.passed, true);

    const missingScreenshots = { ...screenshots };
    missingScreenshots['tray-native-menu'] = join(dir, 'missing-tray-menu.png');
    const missingTray = runRecorder(join(dir, 'missing-tray.json'), missingScreenshots, '--daemon-bin', daemonBin);
    assert.notEqual(missingTray.status, 0);
    const failed = JSON.parse(outputText(missingTray.stdout)) as Record<string, unknown>;
    assert.equal(failed.passed, false);
    assert.equal((failed.screenshots_exist as Record<string, unknown>).tray_native_menu, false);
  });
});

test('product visual evidence recorder requires daemon context before passing', () => {
  withScreenshots((dir, screenshots) => {
    const output = join(dir, 'missing-daemon-context.json');
    const result = runRecorder(output, screenshots);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires tray probe live state', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir, false);
    const output = join(dir, 'missing-live-state.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires live tray menu snapshot', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir, true, ['Status: Up', 'Clients: 0']);
    const output = join(dir, 'missing-menu-items.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder can bind evidence to daemon context', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);

    const output = join(dir, 'product-visual-evidence.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const context = evidence.daemon_context as Record<string, unknown>;
    assert.equal(context.binary_path, resolve(daemonBin));
    assert.equal((context.status as Record<string, unknown>).running, true);
    assert.equal((context.tray_probe as Record<string, unknown>).native_host, true);
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

function writeFakeDaemon(dir: string, stateFetchOk = true, menuItems = ['Status: Up', 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP', 'Quit Office MCP']): string {
  const daemonBin = join(dir, process.platform === 'win32' ? 'daemon.cmd' : 'daemon.sh');
  writeFileSync(daemonBin, fakeDaemonScript(stateFetchOk, menuItems));
  chmodSync(daemonBin, 0o755);
  return daemonBin;
}

function fakeDaemonScript(stateFetchOk: boolean, menuItems: string[]): string {
  const status = JSON.stringify({ running: true, uiUrl: 'https://localhost:8765/ui/' });
  const trayProbe = JSON.stringify({
    native_host: true,
    state_fetch_ok: stateFetchOk,
    snapshot: {
      tooltip: 'Office MCP - Up - 0 clients - 0 documents',
      menu_items: menuItems,
      menu: [
        { kind: 'read_only', enabled: false, label: 'Status: Up' },
        { kind: 'read_only', enabled: false, label: 'Clients: 0' },
        { kind: 'read_only', enabled: false, label: 'Documents: 0' },
        { kind: 'separator', enabled: false, label: '---' },
        { kind: 'action', enabled: true, label: 'Show Office MCP', action: 'show_ui' },
        { kind: 'action', enabled: true, label: 'Quit Office MCP', action: 'quit' }
      ]
    }
  });
  if (process.platform === 'win32') {
    return `@echo off\r\nif "%1"=="daemon" echo ${status}\r\nif "%1"=="tray" echo ${trayProbe}\r\n`;
  }
  return `#!/bin/sh\nif [ "$1" = "daemon" ]; then printf '%s\\n' '${status}'; fi\nif [ "$1" = "tray" ]; then printf '%s\\n' '${trayProbe}'; fi\n`;
}
