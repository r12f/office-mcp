import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import test from 'node:test';
import { tinyPng } from './image-evidence.js';

const TSX = './node_modules/tsx/dist/cli.mjs';
const RECORDER = resolve(process.cwd(), 'record-product-visual-evidence.ts');
const SURFACES = [
  'word-ribbon-command',
  'word-catalog-entry',
  'word-taskpane-title',
  'excel-ribbon-command',
  'excel-catalog-entry',
  'excel-taskpane-title',
  'logo-tray-size',
  'logo-ribbon-size',
  'logo-catalog-thumbnail',
  'logo-daemon-titlebar',
  'logo-installer-metadata',
  'tray-icon',
  'tray-native-menu',
  'tray-tooltip',
  'tray-quit-confirmation'
];

test('product visual evidence recorder requires all product surfaces', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'product-visual-evidence.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.kind, 'product_visual_evidence');
    assert.equal(evidence.product_text_ready, true);
    assert.equal(evidence.catalog_type_ready, true);
    assert.equal(evidence.tray_tooltip_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).excel.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.display_name, 'Office MCP Control');
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.icon_url, 'https://localhost:8765/assets/icon-32.png');
    assert.equal(evidence.rendered_logo_review_ready, true);
    assert.equal(evidence.daemon_context_ready, true);
    assert.equal(evidence.passed, true);

    const missingScreenshots = { ...screenshots };
    missingScreenshots['tray-native-menu'] = join(dir, 'missing-tray-menu.png');
    const missingTray = runRecorder(join(dir, 'missing-tray.json'), missingScreenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.notEqual(missingTray.status, 0);
    const failed = JSON.parse(outputText(missingTray.stdout)) as Record<string, unknown>;
    assert.equal(failed.passed, false);
    assert.equal((failed.screenshots_exist as Record<string, unknown>).tray_native_menu, false);
  });
});

test('product visual evidence recorder requires product identity review flags', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-product-review.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--skip-product-review-flags');
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal((evidence.product_identity_review as Record<string, unknown>).ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires rendered-size and first-run identity review', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-first-run-review.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--skip-rendered-logo-and-first-run-flags');
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const review = evidence.product_identity_review as Record<string, unknown>;
    assert.equal(review.rendered_size_logo_reviewed, false);
    assert.equal(review.word_first_run_identity_ready, false);
    assert.equal(review.excel_first_run_identity_ready, false);
    assert.equal(review.ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires rendered logo review artifact', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const output = join(dir, 'missing-rendered-logo-review.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin);
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.rendered_logo_review_ready, false);
    assert.equal(evidence.passed, false);

    const brokenReviewPath = writeRenderedLogoReview(dir, false);
    const broken = runRecorder(join(dir, 'broken-rendered-logo-review.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', brokenReviewPath);
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(outputText(broken.stdout)) as Record<string, unknown>;
    assert.equal(evidence.rendered_logo_review_ready, false);
  });
});

test('product visual evidence recorder requires daemon context before passing', () => {
  withScreenshots((dir, screenshots) => {
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-daemon-context.json');
    const result = runRecorder(output, screenshots, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires tray probe live state', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir, false);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-live-state.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires live tray menu snapshot', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir, true, ['Status: Up', 'Clients: 0']);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-menu-items.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder rejects truncated screenshots', () => {
  withScreenshots((dir, screenshots) => {
    writeFileSync(screenshots['tray-icon'], Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'truncated-screenshot.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal((evidence.screenshots_exist as Record<string, unknown>).tray_icon, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder can bind evidence to daemon context', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);

    const output = join(dir, 'product-visual-evidence.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const context = evidence.daemon_context as Record<string, unknown>;
    assert.equal(context.binary_path, resolve(daemonBin));
    assert.equal((context.status as Record<string, unknown>).running, true);
    assert.equal((context.tray_probe as Record<string, unknown>).native_host, true);
  });
});

function runRecorder(output: string, screenshots: Record<string, string>, ...extra: string[]): ReturnType<typeof spawnSync> {
  const skipProductReviewFlags = extra.includes('--skip-product-review-flags');
  const skipRenderedLogoAndFirstRunFlags = extra.includes('--skip-rendered-logo-and-first-run-flags');
  const filteredExtra = extra.filter((item) => item !== '--skip-product-review-flags' && item !== '--skip-rendered-logo-and-first-run-flags');
  const hasWordManifest = filteredExtra.includes('--word-manifest-path');
  const hasExcelManifest = filteredExtra.includes('--excel-manifest-path');
  const outputDir = dirname(output);
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
    '--word-catalog-provider', 'Office MCP Control',
    '--word-catalog-description', 'Local office productivity automation and control utility',
    '--word-catalog-type', 'Local productivity automation control utility',
    '--excel-catalog-provider', 'Office MCP Control',
    '--excel-catalog-description', 'Local office productivity automation and control utility',
    '--excel-catalog-type', 'Local productivity automation control utility',
    '--excel-compact-top-block', 'true',
    '--excel-tools-permissions-merged', 'true',
    '--excel-inline-settings', 'true',
    '--excel-server-protocol-row', 'Server 0.1.0 / Protocol 1.0',
    '--excel-document-state', 'Editable'
  ];
  if (!hasWordManifest) args.push('--word-manifest-path', writeManifest(outputDir, 'word'));
  if (!hasExcelManifest) args.push('--excel-manifest-path', writeManifest(outputDir, 'excel'));
  if (!skipProductReviewFlags) {
    args.push(
      '--logo-quality-reviewed', 'true',
      '--addin-identity-reviewed', 'true',
      '--tray-product-polish-reviewed', 'true'
    );
  }
  if (!skipRenderedLogoAndFirstRunFlags) {
    args.push(
      '--rendered-size-logo-reviewed', 'true',
      '--word-first-run-identity-reviewed', 'true',
      '--excel-first-run-identity-reviewed', 'true'
    );
  }
  for (const surface of SURFACES) {
    args.push(`--${surface}`, `Office MCP Control ${surface}`);
    args.push(`--${surface}-screenshot`, screenshots[surface]);
  }
  args.push(...filteredExtra);
  return spawnSync(process.execPath, args, { cwd: process.cwd(), encoding: 'utf8' });
}

function withScreenshots(callback: (dir: string, screenshots: Record<string, string>) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-product-visual-test-'));
  try {
    const screenshots: Record<string, string> = {};
    for (const surface of SURFACES) {
      const path = join(dir, `${surface}.png`);
      writeFileSync(path, tinyPng());
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

function writeRenderedLogoReview(dir: string, ready = true): string {
  const sheetPath = join(dir, 'rendered-logo-review.png');
  writeFileSync(sheetPath, tinyPng());
  const path = join(dir, 'rendered-logo-review.json');
  const surfaces = [
    ['logo_tray_size', 16],
    ['logo_ribbon_size', 32],
    ['logo_catalog_thumbnail', 80],
    ['logo_daemon_titlebar', 20],
    ['logo_installer_metadata', 256]
  ].map(([key, size]) => ({
    key,
    rendered_size_px: ready ? size : 1,
    width: ready ? size : 1,
    height: ready ? size : 1,
    non_empty: ready,
    palette_ready: ready,
    expected_size_ready: ready,
    screenshot_path: sheetPath
  }));
  writeFileSync(path, JSON.stringify({
    schema_version: 1,
    kind: 'rendered_logo_review',
    product_name: 'Office MCP Control',
    sheet_path: sheetPath,
    surfaces,
    ready
  }, null, 2));
  return path;
}

function writeManifest(dir: string, host: 'word' | 'excel'): string {
  const path = join(dir, `${host}-manifest.xml`);
  const context = host === 'word' ? 'Word documents' : 'Excel workbooks';
  writeFileSync(path, `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp>
  <ProviderName>Office MCP Control</ProviderName>
  <DisplayName DefaultValue="Office MCP Control" />
  <Description DefaultValue="Control live ${context} through a local productivity automation control utility." />
  <IconUrl DefaultValue="https://localhost:8765/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:8765/assets/icon-80.png" />
</OfficeApp>
`);
  return path;
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
