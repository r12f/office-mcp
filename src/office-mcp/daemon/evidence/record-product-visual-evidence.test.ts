import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { chmodSync, mkdtempSync, readFileSync, rmSync, statSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import test from 'node:test';
import { tinyPng } from './image-evidence.js';

const TSX = './node_modules/tsx/dist/cli.mjs';
const RECORDER = resolve(process.cwd(), 'record-product-visual-evidence.ts');
const REPO_ROOT = resolve(process.cwd(), '../../../..');
const ASSET_ROOT = resolve(REPO_ROOT, 'src/office-ctl/common/assets');
const LOGO_SURFACES = [
  'logo-tray-size',
  'logo-ribbon-size',
  'logo-catalog-thumbnail',
  'logo-daemon-titlebar',
  'logo-installer-metadata'
];

const TRAY_SURFACES = [
  'tray-icon',
  'tray-native-menu',
  'tray-tooltip',
  'tray-quit-confirmation'
];

const SURFACES = [
  'word-ribbon-command',
  'word-catalog-entry',
  'word-taskpane-title',
  'excel-ribbon-command',
  'excel-catalog-entry',
  'excel-taskpane-title',
  'powerpoint-ribbon-command',
  'powerpoint-catalog-entry',
  'powerpoint-taskpane-title',
  'daemon-main-window',
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
    const wordRuntimeEvidencePath = writeWordRuntimeEvidence(dir);
    const excelRuntimeEvidencePath = writeExcelRuntimeEvidence(dir);
    const powerPointRuntimeEvidencePath = writePowerPointRuntimeEvidence(dir);
    const output = join(dir, 'product-visual-evidence.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--word-runtime-evidence-path', wordRuntimeEvidencePath, '--excel-runtime-evidence-path', excelRuntimeEvidencePath, '--powerpoint-runtime-evidence-path', powerPointRuntimeEvidencePath);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.kind, 'product_visual_evidence');
    assert.equal(evidence.product_text_ready, true);
    assert.equal(evidence.screenshots_fresh_ready, true);
    assert.equal((evidence.screenshots_fresh as Record<string, unknown>).word_ribbon_command, true);
    assert.equal(typeof ((evidence.screenshot_metadata as Record<string, Record<string, unknown>>).word_ribbon_command.age_ms), 'number');
    assert.equal(evidence.catalog_type_ready, true);
    assert.equal(evidence.tray_tooltip_ready, true);
    assert.equal(evidence.tray_menu_surface_kind, 'native');
    assert.equal(evidence.tray_menu_surface_native, true);
    assert.equal((evidence.word_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal((evidence.word_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal((evidence.daemon_main_window as Record<string, unknown>).ready, true);
    assert.equal((evidence.daemon_main_window as Record<string, unknown>).compact_status_details_reviewed, true);
    assert.equal((evidence.screenshots_exist as Record<string, unknown>).daemon_main_window, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).logo_future_office_control_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).final_logo_user_surface_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).current_logo_screenshot_feedback_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).addin_title_icon_type_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).addin_installable_surface_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).current_addin_screenshot_feedback_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).tray_native_first_impression_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).tray_normal_windows_launch_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).current_tray_screenshot_feedback_reviewed, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).current_screenshot_feedback_ready, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).excel.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).powerpoint.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.display_name, 'Office MCP Control');
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.icon_url, 'https://localhost:8765/assets/icon-32.png');
    assert.equal(evidence.rendered_logo_review_ready, true);
    assert.equal(evidence.word_runtime_evidence_ready, true);
    assert.equal(evidence.excel_runtime_evidence_ready, true);
    assert.equal(evidence.powerpoint_runtime_evidence_ready, true);
    assert.equal(evidence.daemon_context_ready, true);
    assert.equal(evidence.passed, true);

    const missingScreenshots = { ...screenshots };
    missingScreenshots['tray-native-menu'] = join(dir, 'missing-tray-menu.png');
    const missingTray = runRecorder(join(dir, 'missing-tray.json'), missingScreenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--word-runtime-evidence-path', wordRuntimeEvidencePath, '--excel-runtime-evidence-path', excelRuntimeEvidencePath, '--powerpoint-runtime-evidence-path', powerPointRuntimeEvidencePath);
    assert.notEqual(missingTray.status, 0);
    const failed = JSON.parse(outputText(missingTray.stdout)) as Record<string, unknown>;
    assert.equal(failed.passed, false);
    assert.equal((failed.screenshots_exist as Record<string, unknown>).tray_native_menu, false);
  });
});

test('product visual evidence recorder rejects stale screenshots', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir, true, ['Status: Degraded', 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP Control', 'Quit Office MCP Control']);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const staleDate = new Date(Date.now() - 60 * 60 * 1000);
    utimesSync(screenshots['word-ribbon-command'], staleDate, staleDate);
    const output = join(dir, 'stale-screenshot.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--screenshot-freshness-window-ms', '1000');

    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.screenshots_fresh_ready, false);
    assert.equal((evidence.screenshots_fresh as Record<string, unknown>).word_ribbon_command, false);
    assert.equal((evidence.screenshot_metadata as Record<string, Record<string, unknown>>).word_ribbon_command.fresh, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder ties daemon main window review to its screenshot', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const missingScreenshots = { ...screenshots, 'daemon-main-window': join(dir, 'missing-daemon-main-window.png') };
    const output = join(dir, 'missing-daemon-main-window.json');
    const result = runRecorder(output, missingScreenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);

    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const daemonMainWindow = evidence.daemon_main_window as Record<string, unknown>;
    assert.equal((evidence.screenshots_exist as Record<string, unknown>).daemon_main_window, false);
    assert.equal(daemonMainWindow.reviewed, true);
    assert.equal(daemonMainWindow.screenshot_ready, false);
    assert.equal(daemonMainWindow.ready, false);
    assert.equal(evidence.passed, false);
  });
});


test('deployment spec manual tray evidence command matches current native tray gates', () => {
  const deployment = readFileSync(resolve(process.cwd(), '../../../..', 'doc/spec/07-deployment.md'), 'utf8');
  const commandLine = deployment.split('\n').find((line) => line.includes('npm run evidence:record-tray-manual')) ?? '';

  for (const required of [
    '--daemon-bin',
    '--visible-icon',
    '--right-click-menu',
    '--menu-opened-from-tray-icon',
    '--native-menu-appearance-reviewed',
    '--menu-anchored-to-tray-icon',
    '--os-native-menu-behavior-reviewed',
    '--keyboard-menu-access-reviewed',
    '--native-quit-confirmation-reviewed',
    '--menu-surface-kind native',
    '--show-ui-opened',
    '--tooltip',
    '--screenshot-path',
    '--tray-icon-screenshot',
    '--tray-native-menu-screenshot',
    '--tray-tooltip-screenshot',
    '--tray-quit-confirmation-screenshot',
    '--screenshot-freshness-window-ms'
  ]) {
    assert.match(commandLine, new RegExp(required));
  }
  assert.match(deployment, /freshness metadata/i);
  assert.match(deployment, /stale screenshots/i);
  assert.equal(optionValue(commandLine.split(/\s+/), '--screenshot-path'), optionValue(commandLine.split(/\s+/), '--tray-icon-screenshot'));
});

test('roadmap documents manual tray screenshot freshness gate', () => {
  const roadmap = readFileSync(resolve(process.cwd(), '../../../..', 'doc/spec/08-roadmap.md'), 'utf8');
  assert.match(roadmap, /manual tray evidence[\s\S]*freshness metadata/i);
  assert.match(roadmap, /manual tray evidence[\s\S]*stale screenshots/i);
});

test('deployment spec product visual evidence command matches current product visual gates', () => {
  const deployment = readFileSync(resolve(process.cwd(), '../../../..', 'doc/spec/07-deployment.md'), 'utf8');
  const commandLine = deployment.split('\n').find((line) => line.includes('npm run evidence:record-product-visual')) ?? '';
  const catalogIdentityLine = deployment.split('\n').find((line) => line.includes('record-catalog-identity-review.mjs')) ?? '';

  for (const required of [
    '--catalog-identity-review-path',
    '--word-tool-e2e-report-path',
    '--excel-tool-e2e-report-path',
    '--powerpoint-tool-e2e-report-path',
    '--word-runtime-evidence-path',
    '--excel-runtime-evidence-path',
    '--powerpoint-runtime-evidence-path',
    '--powerpoint-ribbon-command',
    '--powerpoint-ribbon-command-screenshot',
    '--powerpoint-catalog-entry',
    '--powerpoint-catalog-entry-screenshot',
    '--powerpoint-taskpane-title',
    '--powerpoint-taskpane-title-screenshot',
    '--daemon-main-window',
    '--daemon-main-window-screenshot',
    '--daemon-main-window-reviewed',
    '--daemon-main-window-compact-reviewed',
    '--daemon-main-window-three-column-reviewed',
    '--logo-quality-reviewed',
    '--logo-future-office-control-reviewed',
    '--final-logo-user-surface-reviewed',
    '--current-logo-screenshot-feedback-reviewed',
    '--rendered-size-logo-reviewed',
    '--addin-identity-reviewed',
    '--addin-title-icon-type-reviewed',
    '--addin-installable-surface-reviewed',
    '--current-addin-screenshot-feedback-reviewed',
    '--word-first-run-identity-reviewed',
    '--excel-first-run-identity-reviewed',
    '--powerpoint-first-run-identity-reviewed',
    '--tray-product-polish-reviewed',
    '--tray-native-first-impression-reviewed',
    '--tray-normal-windows-launch-reviewed',
    '--current-tray-screenshot-feedback-reviewed',
    '--word-compact-top-block',
    '--word-tools-permissions-merged',
    '--word-inline-settings',
    '--word-server-protocol-row',
    '--word-document-state',
    '--excel-compact-top-block',
    '--excel-tools-permissions-merged',
    '--excel-inline-settings',
    '--excel-server-protocol-row',
    '--excel-document-state',
    '--powerpoint-compact-top-block',
    '--powerpoint-tools-permissions-merged',
    '--powerpoint-inline-settings',
    '--powerpoint-server-protocol-row',
    '--powerpoint-document-state',
    '--screenshot-freshness-window-ms'
  ]) {
    assert.match(commandLine, new RegExp(required));
  }
  assert.doesNotMatch(commandLine, /--powerpoint-catalog-provider/);
  assert.doesNotMatch(commandLine, /--word-catalog-type/);
  assert.match(catalogIdentityLine, /record-catalog-identity-review\.mjs/);
  assert.match(catalogIdentityLine, /--catalog-path/);
  assert.match(catalogIdentityLine, /artifacts\\portable-stage\\addin-catalog/);
  assert.doesNotMatch(catalogIdentityLine, /\.\.\\\.\.\\\.\.\\\.\.\\addin-catalog/);
  assert.match(catalogIdentityLine, /--output .*catalog-identity-review\.json/);
});

test('deployment spec Word runtime evidence command matches self-contained gate', () => {
  const deployment = readFileSync(resolve(process.cwd(), '../../../..', 'doc/spec/07-deployment.md'), 'utf8');

  assert.match(deployment, /npm run evidence:word/);
  assert.match(deployment, /runtime-evidence-word\.json/);
  assert.doesNotMatch(deployment, /--require-mutation|--require-full-word-smoke|--require-com-tracked-changes/);
  assert.doesNotMatch(deployment, /runtime-evidence-full\.json/);
  assert.doesNotMatch(deployment, /npm run evidence:runtime -- --endpoint http:\/\/127\.0\.0\.1:8800\/mcp --output .*runtime-evidence-full\.json/);
});

test('deployment spec describes current Word Excel and PowerPoint product surface', () => {
  const deployment = readFileSync(resolve(process.cwd(), '../../../..', 'doc/spec/07-deployment.md'), 'utf8');
  const mcpSurface = readFileSync(resolve(process.cwd(), '../../../..', 'doc/spec/03-mcp-tool-surface.md'), 'utf8');

  assert.match(deployment, /Current product scope is Word, Excel, and PowerPoint/);
  assert.match(deployment, /`office-ctl\/word\/`/);
  assert.match(deployment, /`office-ctl\/excel\/`/);
  assert.match(deployment, /`src\/office-ctl\/powerpoint\/`/);
  assert.match(mcpSurface, /"available_tool_count": 25/);
  assert.doesNotMatch(mcpSurface, /"available_tool_count": 27/);
  assert.match(deployment, /Word task pane\s+`https:\/\/localhost:8765\/word\/taskpane\.html`/);
  assert.match(deployment, /Excel task pane\s+`https:\/\/localhost:8765\/excel\/taskpane\.html`/);
  assert.match(deployment, /PowerPoint task pane\s+`https:\/\/localhost:8765\/powerpoint\/taskpane\.html`/);
  assert.doesNotMatch(deployment, /PowerPoint add-in scaffold/);
  assert.doesNotMatch(deployment, /PowerPoint and Outlook planned/);
  assert.doesNotMatch(deployment, /Add-in task pane: `https:\/\/localhost:8765\/taskpane\.html`/);
});


test('product visual evidence recorder rejects experimental catalog type and non-native tray surface', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'unfinished-product-surface.json');
    const result = runRecorder(
      output,
      screenshots,
      '--daemon-bin', daemonBin,
      '--rendered-logo-review-path', renderedLogoReviewPath,
      '--catalog-type', 'Local productivity automation control utility protocol bridge debug panel',
      '--word-catalog-type', 'Local productivity automation control utility protocol bridge debug panel',
      '--tray-menu-surface-kind', 'webview'
    );
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.catalog_type_ready, false);
    assert.equal(evidence.tray_menu_surface_kind, 'webview');
    assert.equal(evidence.tray_menu_surface_native, false);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.ready, false);
    assert.equal(evidence.passed, false);
  });
});
test('product visual evidence recorder derives logo surfaces from rendered review artifact', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'derived-logo-surfaces.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--skip-logo-surface-args');
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const screenshotPaths = evidence.screenshot_paths as Record<string, string>;
    const screenshotsExist = evidence.screenshots_exist as Record<string, boolean>;
    const observations = evidence.observations as Record<string, string>;
    const renderedLogoReview = evidence.rendered_logo_review as Record<string, unknown>;
    const sheetPath = renderedLogoReview.sheet_path;

    for (const surface of LOGO_SURFACES) {
      const key = surface.replaceAll('-', '_');
      assert.equal(screenshotPaths[key], sheetPath);
      assert.equal(screenshotsExist[key], true);
      assert.match(observations[key], /Office MCP Control rendered logo review/);
    }
    assert.equal(evidence.product_text_ready, true);
    assert.equal(evidence.passed, true);
  });
});

test('product visual evidence recorder derives tray surfaces from manual tray artifact', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const manualTrayEvidencePath = writeManualTrayEvidence(dir, true, 'Degraded');
    const output = join(dir, 'derived-tray-surfaces.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--manual-tray-evidence-path', manualTrayEvidencePath, '--skip-tray-surface-args', '--skip-tray-tooltip-arg');
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const screenshotPaths = evidence.screenshot_paths as Record<string, string>;
    const screenshotsExist = evidence.screenshots_exist as Record<string, boolean>;
    const observations = evidence.observations as Record<string, string>;
    const manualTray = evidence.manual_tray_evidence as Record<string, unknown>;
    const manualPaths = manualTray.tray_surface_screenshot_paths as Record<string, string>;

    for (const surface of TRAY_SURFACES) {
      const key = surface.replaceAll('-', '_');
      assert.equal(screenshotPaths[key], manualPaths[key]);
      assert.equal(screenshotsExist[key], true);
      if (key === 'tray_tooltip') assert.equal(observations[key], manualTray.observed_tooltip);
      else assert.match(observations[key], /Office MCP Control manual tray evidence/);
    }
    assert.equal(evidence.tray_tooltip, manualTray.observed_tooltip);
    assert.equal(evidence.tray_tooltip_ready, true);
    assert.equal(evidence.product_text_ready, true);
    assert.equal(evidence.passed, true);
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


test('product visual evidence recorder requires final user-surface polish reviews', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-final-user-surface-reviews.json');
    const result = runRecorder(
      output,
      screenshots,
      '--daemon-bin', daemonBin,
      '--rendered-logo-review-path', renderedLogoReviewPath,
      '--skip-product-review-flags',
      '--logo-quality-reviewed', 'true',
      '--addin-identity-reviewed', 'true',
      '--tray-product-polish-reviewed', 'true'
    );
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const review = evidence.product_identity_review as Record<string, unknown>;
    assert.equal(review.logo_quality_reviewed, true);
    assert.equal(review.logo_future_office_control_reviewed, false);
    assert.equal(review.final_logo_user_surface_reviewed, false);
    assert.equal(review.addin_identity_reviewed, true);
    assert.equal(review.addin_title_icon_type_reviewed, false);
    assert.equal(review.addin_installable_surface_reviewed, false);
    assert.equal(review.tray_product_polish_reviewed, true);
    assert.equal(review.tray_native_first_impression_reviewed, false);
    assert.equal(review.tray_normal_windows_launch_reviewed, false);
    assert.equal(review.ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires current screenshot feedback reviews', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-current-screenshot-feedback-reviews.json');
    const result = runRecorder(
      output,
      screenshots,
      '--daemon-bin', daemonBin,
      '--rendered-logo-review-path', renderedLogoReviewPath,
      '--skip-product-review-flags',
      '--logo-quality-reviewed', 'true',
      '--logo-future-office-control-reviewed', 'true',
      '--final-logo-user-surface-reviewed', 'true',
      '--addin-identity-reviewed', 'true',
      '--addin-title-icon-type-reviewed', 'true',
      '--addin-installable-surface-reviewed', 'true',
      '--tray-product-polish-reviewed', 'true',
      '--tray-native-first-impression-reviewed', 'true',
      '--tray-normal-windows-launch-reviewed', 'true'
    );
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const review = evidence.product_identity_review as Record<string, unknown>;
    assert.equal(review.current_logo_screenshot_feedback_reviewed, false);
    assert.equal(review.current_addin_screenshot_feedback_reviewed, false);
    assert.equal(review.current_tray_screenshot_feedback_reviewed, false);
    assert.equal(review.current_screenshot_feedback_ready, false);
    assert.equal(review.ready, false);
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
    assert.equal(review.powerpoint_first_run_identity_ready, false);
    assert.equal(review.ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires rendered logo design review', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const renderedLogoReview = JSON.parse(readFileSync(renderedLogoReviewPath, 'utf8')) as Record<string, unknown>;
    renderedLogoReview.design_review = { ready: false, rejects_generic_readings: [] };
    writeFileSync(renderedLogoReviewPath, JSON.stringify(renderedLogoReview, null, 2));
    const output = join(dir, 'missing-rendered-logo-design-review.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.rendered_logo_review_ready, false);
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

test('product visual evidence recorder requires rendered logo concept pass', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReview = JSON.parse(readFileSync(writeRenderedLogoReview(dir), 'utf8')) as Record<string, unknown>;
    const designReview = renderedLogoReview.design_review as Record<string, unknown>;
    designReview.concept_pass = { ready: false, concepts: [] };
    const renderedLogoReviewPath = join(dir, 'missing-rendered-logo-concept-pass.json');
    writeFileSync(renderedLogoReviewPath, JSON.stringify(renderedLogoReview, null, 2));
    const output = join(dir, 'missing-rendered-logo-concept-pass-evidence.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);

    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8'));
    assert.equal(evidence.rendered_logo_review_ready, false);
  });
});

test('product visual evidence recorder requires rendered logo asset fingerprints to match current assets', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const review = JSON.parse(readFileSync(renderedLogoReviewPath, 'utf8')) as Record<string, unknown>;
    review.source_asset_sha256 = '0'.repeat(64);
    const surfaces = review.surfaces as Array<Record<string, unknown>>;
    surfaces[0].asset_sha256 = '0'.repeat(64);
    writeFileSync(renderedLogoReviewPath, JSON.stringify(review, null, 2));
    const output = join(dir, 'stale-rendered-logo-review.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.rendered_logo_review_ready, false);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).rendered_logo_review_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires Excel runtime evidence', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-excel-runtime-evidence.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--excel-runtime-evidence-path', join(dir, 'missing-excel.json'));
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.excel_runtime_evidence_ready, false);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).excel_runtime_evidence_ready, false);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).density_ready, false);

    const broken = runRecorder(join(dir, 'broken-excel-runtime-evidence.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--excel-runtime-evidence-path', writeExcelRuntimeEvidence(dir, false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(outputText(broken.stdout)) as Record<string, unknown>;
    assert.equal(evidence.excel_runtime_evidence_ready, false);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
  });
});

test('product visual evidence recorder requires Word runtime evidence', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-word-runtime-evidence.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--word-runtime-evidence-path', join(dir, 'missing-word.json'));
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.word_runtime_evidence_ready, false);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).word_runtime_evidence_ready, false);
    assert.equal((evidence.word_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
    assert.equal((evidence.word_taskpane as Record<string, unknown>).density_ready, false);

    const broken = runRecorder(join(dir, 'broken-word-runtime-evidence.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--word-runtime-evidence-path', writeWordRuntimeEvidence(dir, false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(outputText(broken.stdout)) as Record<string, unknown>;
    assert.equal(evidence.word_runtime_evidence_ready, false);
    assert.equal((evidence.word_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
    assert.equal((evidence.word_taskpane as Record<string, unknown>).density_ready, false);
  });
});

test('product visual evidence recorder requires embedded manual tray evidence', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-manual-tray-evidence.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--manual-tray-evidence-path', join(dir, 'missing-tray.json'));
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.manual_tray_evidence_ready, false);
    assert.equal(evidence.passed, false);

    const broken = runRecorder(join(dir, 'broken-manual-tray-evidence.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--manual-tray-evidence-path', writeManualTrayEvidence(dir, false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(outputText(broken.stdout)) as Record<string, unknown>;
    assert.equal(evidence.manual_tray_evidence_ready, false);
  });
});

test('product visual evidence recorder rejects manual tray artifacts without ready surface screenshots', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const manualTrayEvidencePath = writeManualTrayEvidence(dir, true);
    const manualTray = JSON.parse(readFileSync(manualTrayEvidencePath, 'utf8')) as Record<string, unknown>;
    manualTray.tray_surface_screenshots_ready = false;
    manualTray.tray_surface_screenshots_distinct = false;
    manualTray.tray_surface_screenshots_exist = { tray_icon: true };
    manualTray.tray_surface_screenshot_paths = { tray_icon: screenshots.tray_icon };
    writeFileSync(manualTrayEvidencePath, JSON.stringify(manualTray, null, 2));

    const output = join(dir, 'manual-tray-missing-surfaces.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--manual-tray-evidence-path', manualTrayEvidencePath, '--skip-tray-surface-args');
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.manual_tray_evidence_ready, false);
    assert.equal((evidence.screenshots_exist as Record<string, unknown>).tray_native_menu, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder rejects manual tray artifacts without primary screenshot binding', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const manualTrayEvidencePath = writeManualTrayEvidence(dir, true);
    const manualTray = JSON.parse(readFileSync(manualTrayEvidencePath, 'utf8')) as Record<string, unknown>;
    manualTray.primary_screenshot_matches_tray_icon = false;
    writeFileSync(manualTrayEvidencePath, JSON.stringify(manualTray, null, 2));

    const output = join(dir, 'manual-tray-missing-primary-binding.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--manual-tray-evidence-path', manualTrayEvidencePath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.manual_tray_evidence_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder rejects manual tray artifacts without observed snapshot binding', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const manualTrayEvidencePath = writeManualTrayEvidence(dir, true);
    const manualTray = JSON.parse(readFileSync(manualTrayEvidencePath, 'utf8')) as Record<string, unknown>;
    delete manualTray.observed_snapshot_binding_ready;
    writeFileSync(manualTrayEvidencePath, JSON.stringify(manualTray, null, 2));

    const output = join(dir, 'manual-tray-missing-observed-snapshot-binding.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--manual-tray-evidence-path', manualTrayEvidencePath);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.manual_tray_evidence_ready, false);
    assert.equal(evidence.passed, false);
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
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--skip-manual-tray-evidence');
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
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--skip-manual-tray-evidence');
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
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--skip-manual-tray-evidence');
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

test('product visual evidence recorder requires structured tray menu snapshot', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir, true, undefined, false);
    const output = join(dir, 'missing-structured-tray-menu.json');
    const result = runRecorder(output, screenshots, '--daemon-bin', daemonBin);
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.daemon_context_ready, false);
    assert.equal(evidence.passed, false);
  });
});


test('product visual evidence recorder requires PowerPoint task pane density review', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-powerpoint-taskpane-density.json');
    const result = runRecorder(
      output,
      screenshots,
      '--daemon-bin', daemonBin,
      '--rendered-logo-review-path', renderedLogoReviewPath,
      '--powerpoint-document-state', 'unknown'
    );
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const taskpane = evidence.powerpoint_taskpane as Record<string, unknown>;
    assert.equal(taskpane.document_state, 'unknown');
    assert.equal(taskpane.document_state_ready, false);
    assert.equal(taskpane.density_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires PowerPoint runtime evidence', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-powerpoint-runtime-evidence.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--powerpoint-runtime-evidence-path', join(dir, 'missing-powerpoint.json'));
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.powerpoint_runtime_evidence_ready, false);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).powerpoint_runtime_evidence_ready, false);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).density_ready, false);

    const broken = runRecorder(join(dir, 'broken-powerpoint-runtime-evidence.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--powerpoint-runtime-evidence-path', writePowerPointRuntimeEvidence(dir, false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(outputText(broken.stdout)) as Record<string, unknown>;
    assert.equal(evidence.powerpoint_runtime_evidence_ready, false);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).density_ready, false);
  });
});

test('product visual evidence recorder requires Office tool E2E reports', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-office-tool-e2e.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--word-tool-e2e-report-path', join(dir, 'missing-word-e2e.json'));
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.office_tool_e2e_ready, false);
    assert.equal((evidence.office_tool_e2e as Record<string, Record<string, unknown>>).word.ready, false);

    const broken = runRecorder(join(dir, 'broken-office-tool-e2e.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--excel-tool-e2e-report-path', writeOfficeToolE2eReport(dir, 'Excel', false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(readFileSync(join(dir, 'broken-office-tool-e2e.json'), 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.office_tool_e2e_ready, false);
    assert.equal((evidence.office_tool_e2e as Record<string, Record<string, unknown>>).excel.ready, false);

    const skippedActivationPath = writeOfficeToolE2eReport(dir, 'PowerPoint', true);
    const skippedActivationReport = JSON.parse(readFileSync(skippedActivationPath, 'utf8')) as Record<string, unknown>;
    skippedActivationReport.addin_activation = { activated: false, skipped: 'no-activator-configured' };
    writeFileSync(skippedActivationPath, JSON.stringify(skippedActivationReport, null, 2));
    const skippedActivation = runRecorder(join(dir, 'skipped-office-tool-e2e-activation.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--powerpoint-tool-e2e-report-path', skippedActivationPath);
    assert.notEqual(skippedActivation.status, 0);
    evidence = JSON.parse(readFileSync(join(dir, 'skipped-office-tool-e2e-activation.json'), 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.office_tool_e2e_ready, false);
    assert.equal((evidence.office_tool_e2e as Record<string, Record<string, unknown>>).powerpoint.ready, false);

    const weakActivationPath = writeOfficeToolE2eReport(dir, 'PowerPoint', true);
    const weakActivationReport = JSON.parse(readFileSync(weakActivationPath, 'utf8')) as Record<string, unknown>;
    weakActivationReport.addin_activation = { activated: true };
    writeFileSync(weakActivationPath, JSON.stringify(weakActivationReport, null, 2));
    const weakActivation = runRecorder(join(dir, 'weak-office-tool-e2e-activation.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--powerpoint-tool-e2e-report-path', weakActivationPath);
    assert.notEqual(weakActivation.status, 0);
    evidence = JSON.parse(readFileSync(join(dir, 'weak-office-tool-e2e-activation.json'), 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.office_tool_e2e_ready, false);
    assert.equal((evidence.office_tool_e2e as Record<string, Record<string, unknown>>).powerpoint.ready, false);

    const skippedCleanupPath = writeOfficeToolE2eReport(dir, 'PowerPoint', true);
    const skippedCleanupReport = JSON.parse(readFileSync(skippedCleanupPath, 'utf8')) as Record<string, unknown>;
    skippedCleanupReport.cleanup = { closed_by_driver: false, deleted: false, deleted_path_count: 0, skipped: 'manual-debug' };
    writeFileSync(skippedCleanupPath, JSON.stringify(skippedCleanupReport, null, 2));
    const skippedCleanup = runRecorder(join(dir, 'skipped-office-tool-e2e-cleanup.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--powerpoint-tool-e2e-report-path', skippedCleanupPath);
    assert.notEqual(skippedCleanup.status, 0);
    evidence = JSON.parse(readFileSync(join(dir, 'skipped-office-tool-e2e-cleanup.json'), 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.office_tool_e2e_ready, false);
    assert.equal((evidence.office_tool_e2e as Record<string, Record<string, unknown>>).powerpoint.ready, false);

    const passing = runRecorder(join(dir, 'office-tool-e2e-ready.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath);
    assert.equal(passing.status, 0, outputText(passing.stdout) + outputText(passing.stderr));
    evidence = JSON.parse(readFileSync(join(dir, 'office-tool-e2e-ready.json'), 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.office_tool_e2e_ready, true);
    const officeToolE2e = evidence.office_tool_e2e as Record<string, Record<string, unknown>>;
    assert.equal(officeToolE2e.word.ready, true);
    assert.deepEqual(officeToolE2e.word.daemon, { endpoint: 'http://127.0.0.1:8765/mcp' });
    assert.deepEqual(officeToolE2e.word.document, { path: 'word-fixture' });
    assert.deepEqual(officeToolE2e.word.session, { session_id: 'word-session', available_tool_count: 2 });
    assert.equal(officeToolE2e.excel.ready, true);
    assert.deepEqual(officeToolE2e.excel.daemon, { endpoint: 'http://127.0.0.1:8765/mcp' });
    assert.deepEqual(officeToolE2e.excel.document, { path: 'excel-fixture' });
    assert.deepEqual(officeToolE2e.excel.session, { session_id: 'excel-session', available_tool_count: 2 });
    assert.equal(officeToolE2e.powerpoint.ready, true);
    assert.deepEqual(officeToolE2e.powerpoint.daemon, { endpoint: 'http://127.0.0.1:8765/mcp' });
    assert.deepEqual(officeToolE2e.powerpoint.document, { path: 'powerpoint-fixture' });
    assert.deepEqual(officeToolE2e.powerpoint.session, { session_id: 'powerpoint-session', available_tool_count: 2 });
  });
});

test('product visual evidence recorder requires Office tool E2E cleanup deleted paths', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const catalogIdentityReviewPath = writeCatalogIdentityReview(dir);
    const wordToolE2eReportPath = writeOfficeToolE2eReport(dir, 'Word', true);
    const wordToolE2e = JSON.parse(readFileSync(wordToolE2eReportPath, 'utf8')) as Record<string, unknown>;
    wordToolE2e.cleanup = { closed_by_driver: true, deleted: true, deleted_path_count: 1 };
    writeFileSync(wordToolE2eReportPath, JSON.stringify(wordToolE2e, null, 2));

    const output = join(dir, 'office-tool-e2e-missing-deleted-paths.json');
    const result = runRecorder(
      output,
      screenshots,
      '--daemon-bin', daemonBin,
      '--rendered-logo-review-path', renderedLogoReviewPath,
      '--catalog-identity-review-path', catalogIdentityReviewPath,
      '--word-tool-e2e-report-path', wordToolE2eReportPath,
      '--excel-tool-e2e-report-path', writeOfficeToolE2eReport(dir, 'Excel', true),
      '--powerpoint-tool-e2e-report-path', writeOfficeToolE2eReport(dir, 'PowerPoint', true),
      '--word-runtime-evidence-path', writeWordRuntimeEvidence(dir),
      '--excel-runtime-evidence-path', writeExcelRuntimeEvidence(dir),
      '--powerpoint-runtime-evidence-path', writePowerPointRuntimeEvidence(dir)
    );

    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.office_tool_e2e_ready, false);
    const officeToolE2e = evidence.office_tool_e2e as Record<string, Record<string, unknown>>;
    assert.equal(officeToolE2e.word.ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder requires Word task pane density review', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-word-taskpane-density.json');
    const result = runRecorder(
      output,
      screenshots,
      '--daemon-bin', daemonBin,
      '--rendered-logo-review-path', renderedLogoReviewPath,
      '--word-document-state', 'unknown'
    );
    assert.notEqual(result.status, 0);
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    const taskpane = evidence.word_taskpane as Record<string, unknown>;
    assert.equal(taskpane.document_state, 'unknown');
    assert.equal(taskpane.document_state_ready, false);
    assert.equal(taskpane.density_ready, false);
    assert.equal(evidence.passed, false);
  });
});

test('product visual evidence recorder reads evidence artifact paths from environment', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const wordRuntimeEvidencePath = writeWordRuntimeEvidence(dir);
    const excelRuntimeEvidencePath = writeExcelRuntimeEvidence(dir);
    const powerPointRuntimeEvidencePath = writePowerPointRuntimeEvidence(dir);
    const manualTrayEvidencePath = writeManualTrayEvidence(dir);
    const output = join(dir, 'product-visual-env-evidence.json');
    const result = runRecorder(
      output,
      screenshots,
      '--daemon-bin', daemonBin,
      '--skip-logo-surface-args',
      '--skip-tray-surface-args',
      '--env-rendered-logo-review-path', renderedLogoReviewPath,
      '--env-word-runtime-evidence-path', wordRuntimeEvidencePath,
      '--env-excel-runtime-evidence-path', excelRuntimeEvidencePath,
      '--env-powerpoint-runtime-evidence-path', powerPointRuntimeEvidencePath,
      '--env-manual-tray-evidence-path', manualTrayEvidencePath
    );

    assert.equal(result.status, 0, outputText(result.stderr) || outputText(result.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.rendered_logo_review_ready, true);
    assert.equal(evidence.catalog_identity_review_ready, true);
    assert.equal((evidence.word_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal((evidence.powerpoint_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal(evidence.powerpoint_runtime_evidence_ready, true);
    assert.equal(evidence.word_runtime_evidence_ready, true);
    assert.equal(evidence.manual_tray_evidence_ready, true);
    assert.equal(evidence.passed, true);
  });
});

test('product visual evidence recorder requires catalog identity review artifact', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const output = join(dir, 'missing-catalog-identity-review.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--skip-catalog-identity-review');
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.catalog_identity_review_ready, false);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).catalog_identity_review_ready, false);
    assert.equal(evidence.passed, false);

    const broken = runRecorder(join(dir, 'broken-catalog-identity-review.json'), screenshots, '--daemon-bin', daemonBin, '--catalog-identity-review-path', writeCatalogIdentityReview(dir, false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(readFileSync(join(dir, 'broken-catalog-identity-review.json'), 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.catalog_identity_review_ready, false);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).catalog_identity_review_ready, false);
    assert.equal(evidence.passed, false);
  });
});

function runRecorder(output: string, screenshots: Record<string, string>, ...extra: string[]): ReturnType<typeof spawnSync> {
  const skipProductReviewFlags = extra.includes('--skip-product-review-flags');
  const skipRenderedLogoAndFirstRunFlags = extra.includes('--skip-rendered-logo-and-first-run-flags');
  const skipLogoSurfaceArgs = extra.includes('--skip-logo-surface-args');
  const skipTraySurfaceArgs = extra.includes('--skip-tray-surface-args');
  const skipTrayTooltipArg = extra.includes('--skip-tray-tooltip-arg');
  const skipManualTrayEvidence = extra.includes('--skip-manual-tray-evidence');
  const skipCatalogIdentityReview = extra.includes('--skip-catalog-identity-review');
  const skipOfficeToolE2eReports = extra.includes('--skip-office-tool-e2e-reports');
  const envRenderedLogoReviewPath = optionValue(extra, '--env-rendered-logo-review-path');
  const envWordRuntimeEvidencePath = optionValue(extra, '--env-word-runtime-evidence-path');
  const envExcelRuntimeEvidencePath = optionValue(extra, '--env-excel-runtime-evidence-path');
  const envPowerPointRuntimeEvidencePath = optionValue(extra, '--env-powerpoint-runtime-evidence-path');
  const envManualTrayEvidencePath = optionValue(extra, '--env-manual-tray-evidence-path');
  const envCatalogIdentityReviewPath = optionValue(extra, '--env-catalog-identity-review-path');
  const explicitTrayMenuSurfaceKind = extra.includes('--tray-menu-surface-kind');
  const explicitCatalogType = extra.includes('--catalog-type');
  const explicitWordCatalogType = extra.includes('--word-catalog-type');
  const explicitExcelCatalogType = extra.includes('--excel-catalog-type');
  const explicitPowerPointCatalogType = extra.includes('--powerpoint-catalog-type');
  const explicitWordDocumentState = extra.includes('--word-document-state');
  const explicitPowerPointDocumentState = extra.includes('--powerpoint-document-state');
  const filteredExtra = extra.filter((item, index) => {
    const previous = extra[index - 1];
    if (previous === '--env-rendered-logo-review-path' || previous === '--env-word-runtime-evidence-path' || previous === '--env-excel-runtime-evidence-path' || previous === '--env-powerpoint-runtime-evidence-path' || previous === '--env-manual-tray-evidence-path' || previous === '--env-catalog-identity-review-path') return false;
    return item !== '--skip-product-review-flags'
      && item !== '--skip-rendered-logo-and-first-run-flags'
      && item !== '--skip-logo-surface-args'
      && item !== '--skip-tray-surface-args'
      && item !== '--skip-tray-tooltip-arg'
      && item !== '--skip-manual-tray-evidence'
      && item !== '--skip-catalog-identity-review'
      && item !== '--skip-office-tool-e2e-reports'
      && item !== '--env-rendered-logo-review-path'
      && item !== '--env-word-runtime-evidence-path'
      && item !== '--env-excel-runtime-evidence-path'
      && item !== '--env-powerpoint-runtime-evidence-path'
      && item !== '--env-manual-tray-evidence-path'
      && item !== '--env-catalog-identity-review-path';
  });
  const hasWordManifest = filteredExtra.includes('--word-manifest-path');
  const hasExcelManifest = filteredExtra.includes('--excel-manifest-path');
  const hasPowerPointManifest = filteredExtra.includes('--powerpoint-manifest-path');
  const hasWordRuntimeEvidence = filteredExtra.includes('--word-runtime-evidence-path');
  const hasExcelRuntimeEvidence = filteredExtra.includes('--excel-runtime-evidence-path');
  const hasPowerPointRuntimeEvidence = filteredExtra.includes('--powerpoint-runtime-evidence-path');
  const hasManualTrayEvidence = filteredExtra.includes('--manual-tray-evidence-path');
  const hasCatalogIdentityReview = filteredExtra.includes('--catalog-identity-review-path');
  const hasWordToolE2eReport = filteredExtra.includes('--word-tool-e2e-report-path');
  const hasExcelToolE2eReport = filteredExtra.includes('--excel-tool-e2e-report-path');
  const hasPowerPointToolE2eReport = filteredExtra.includes('--powerpoint-tool-e2e-report-path');
  const outputDir = dirname(output);
  const args = [
    TSX,
    RECORDER,
    '--output', output,
    ...(explicitCatalogType ? [] : ['--catalog-type', 'Local productivity automation control utility']),
    '--catalog-icon-visible', 'true',
    ...(skipTrayTooltipArg ? [] : ['--tray-tooltip', 'Office MCP Control - Up - 0 clients - 0 documents']),
    '--tray-icon-visible', 'true',
    '--tray-menu-native', 'true',
    ...(explicitTrayMenuSurfaceKind ? [] : ['--tray-menu-surface-kind', 'native']),
    '--quit-confirmation-visible', 'true',
    '--word-catalog-provider', 'Office MCP Control',
    '--word-catalog-description', 'Local office productivity automation and control utility',
    ...(explicitWordCatalogType ? [] : ['--word-catalog-type', 'Local productivity automation control utility']),
    '--excel-catalog-provider', 'Office MCP Control',
    '--excel-catalog-description', 'Local office productivity automation and control utility',
    ...(explicitExcelCatalogType ? [] : ['--excel-catalog-type', 'Local productivity automation control utility']),
    '--powerpoint-catalog-provider', 'Office MCP Control',
    '--powerpoint-catalog-description', 'Local office productivity automation and control utility',
    ...(explicitPowerPointCatalogType ? [] : ['--powerpoint-catalog-type', 'Local productivity automation control utility']),
    '--word-compact-top-block', 'true',
    '--word-tools-permissions-merged', 'true',
    '--word-inline-settings', 'true',
    '--word-server-protocol-row', 'Server 0.1.0 / Protocol 1.0',
    ...(explicitWordDocumentState ? [] : ['--word-document-state', 'Editable']),
    '--excel-compact-top-block', 'true',
    '--excel-tools-permissions-merged', 'true',
    '--excel-inline-settings', 'true',
    '--excel-server-protocol-row', 'Server 0.1.0 / Protocol 1.0',
    '--excel-document-state', 'Editable',
    '--powerpoint-compact-top-block', 'true',
    '--powerpoint-tools-permissions-merged', 'true',
    '--powerpoint-inline-settings', 'true',
    '--powerpoint-server-protocol-row', 'Server 0.1.0 / Protocol 1.0',
    ...(explicitPowerPointDocumentState ? [] : ['--powerpoint-document-state', 'Editable']),
    '--daemon-main-window-reviewed', 'true',
    '--daemon-main-window-compact-reviewed', 'true',
    '--daemon-main-window-three-column-reviewed', 'true'
  ];
  if (!hasWordManifest) args.push('--word-manifest-path', writeManifest(outputDir, 'word'));
  if (!hasExcelManifest) args.push('--excel-manifest-path', writeManifest(outputDir, 'excel'));
  if (!hasPowerPointManifest) args.push('--powerpoint-manifest-path', writeManifest(outputDir, 'powerpoint'));
  if (!hasWordRuntimeEvidence) args.push('--word-runtime-evidence-path', writeWordRuntimeEvidence(outputDir));
  if (!hasExcelRuntimeEvidence) args.push('--excel-runtime-evidence-path', writeExcelRuntimeEvidence(outputDir));
  if (!hasPowerPointRuntimeEvidence) args.push('--powerpoint-runtime-evidence-path', writePowerPointRuntimeEvidence(outputDir));
  if (!hasManualTrayEvidence && !skipManualTrayEvidence) args.push('--manual-tray-evidence-path', writeManualTrayEvidence(outputDir));
  if (!hasCatalogIdentityReview && !skipCatalogIdentityReview) args.push('--catalog-identity-review-path', writeCatalogIdentityReview(outputDir));
  if (!skipOfficeToolE2eReports) {
    if (!hasWordToolE2eReport) args.push('--word-tool-e2e-report-path', writeOfficeToolE2eReport(outputDir, 'Word'));
    if (!hasExcelToolE2eReport) args.push('--excel-tool-e2e-report-path', writeOfficeToolE2eReport(outputDir, 'Excel'));
    if (!hasPowerPointToolE2eReport) args.push('--powerpoint-tool-e2e-report-path', writeOfficeToolE2eReport(outputDir, 'PowerPoint'));
  }
  if (!skipProductReviewFlags) {
    args.push(
      '--logo-quality-reviewed', 'true',
      '--logo-future-office-control-reviewed', 'true',
      '--final-logo-user-surface-reviewed', 'true',
      '--current-logo-screenshot-feedback-reviewed', 'true',
      '--addin-identity-reviewed', 'true',
      '--addin-title-icon-type-reviewed', 'true',
      '--addin-installable-surface-reviewed', 'true',
      '--current-addin-screenshot-feedback-reviewed', 'true',
      '--tray-product-polish-reviewed', 'true',
      '--tray-native-first-impression-reviewed', 'true',
      '--tray-normal-windows-launch-reviewed', 'true',
      '--current-tray-screenshot-feedback-reviewed', 'true'
    );
  }
  if (!skipRenderedLogoAndFirstRunFlags) {
    args.push(
      '--rendered-size-logo-reviewed', 'true',
      '--word-first-run-identity-reviewed', 'true',
      '--excel-first-run-identity-reviewed', 'true',
      '--powerpoint-first-run-identity-reviewed', 'true'
    );
  }
  for (const surface of SURFACES) {
    if (skipLogoSurfaceArgs && LOGO_SURFACES.includes(surface)) continue;
    if (skipTraySurfaceArgs && TRAY_SURFACES.includes(surface)) continue;
    args.push(`--${surface}`, `Office MCP Control ${surface}`);
    args.push(`--${surface}-screenshot`, screenshots[surface]);
  }
  args.push(...filteredExtra);
  const env = {
    ...process.env,
    ...(envRenderedLogoReviewPath ? { OFFICE_MCP_RENDERED_LOGO_REVIEW_PATH: envRenderedLogoReviewPath } : {}),
    ...(envWordRuntimeEvidencePath ? { OFFICE_MCP_WORD_RUNTIME_EVIDENCE_PATH: envWordRuntimeEvidencePath } : {}),
    ...(envExcelRuntimeEvidencePath ? { OFFICE_MCP_EXCEL_RUNTIME_EVIDENCE_PATH: envExcelRuntimeEvidencePath } : {}),
    ...(envPowerPointRuntimeEvidencePath ? { OFFICE_MCP_POWERPOINT_RUNTIME_EVIDENCE_PATH: envPowerPointRuntimeEvidencePath } : {}),
    ...(envManualTrayEvidencePath ? { OFFICE_MCP_TRAY_MANUAL_EVIDENCE_PATH: envManualTrayEvidencePath } : {}),
    ...(envCatalogIdentityReviewPath ? { OFFICE_MCP_CATALOG_IDENTITY_REVIEW_PATH: envCatalogIdentityReviewPath } : {})
  };
  return spawnSync(process.execPath, args, { cwd: process.cwd(), encoding: 'utf8', env });
}

function optionValue(args: string[], name: string): string | undefined {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : undefined;
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

function writeFakeDaemon(
  dir: string,
  stateFetchOk = true,
  menuItems = ['Status: Up', 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP Control', 'Quit Office MCP Control'],
  includeStructuredMenu = true
): string {
  const daemonBin = join(dir, process.platform === 'win32' ? 'daemon.cmd' : 'daemon.sh');
  writeFileSync(daemonBin, fakeDaemonScript(stateFetchOk, menuItems, includeStructuredMenu));
  chmodSync(daemonBin, 0o755);
  return daemonBin;
}

function writeRenderedLogoReview(dir: string, ready = true): string {
  const sheetPath = join(dir, 'rendered-logo-review.png');
  writeFileSync(sheetPath, tinyPng());
  const path = join(dir, 'rendered-logo-review.json');
  const surfaces = [
    ['logo_tray_size', 16, 'icon-16.png'],
    ['logo_ribbon_size', 32, 'icon-32.png'],
    ['logo_catalog_thumbnail', 80, 'icon-80.png'],
    ['logo_daemon_titlebar', 20, 'icon-20.png'],
    ['logo_installer_metadata', 256, 'icon-256.png']
  ].map(([key, size, asset]) => ({
    key,
    asset_path: resolve(ASSET_ROOT, String(asset)),
    asset_sha256: ready ? sha256File(resolve(ASSET_ROOT, String(asset))) : '0'.repeat(64),
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
    source_asset_path: resolve(ASSET_ROOT, 'brand-mark.svg'),
    source_asset_sha256: ready ? sha256File(resolve(ASSET_ROOT, 'brand-mark.svg')) : '0'.repeat(64),
    sheet_path: sheetPath,
    design_review: renderedLogoDesignReview(ready),
    surfaces,
    ready
  }, null, 2));
  return path;
}

function renderedLogoDesignReview(ready: boolean) {
  return {
    future_office_control_brief: ready ? 'Future office control: routing geometry and operator control without Office-owned app marks.' : '',
    concept_pass: renderedLogoConceptPass(ready),
    office_productivity_metaphor: ready ? 'Abstract document panes communicate office productivity.' : '',
    user_control_metaphor: ready ? 'Command routing and operator nodes communicate local user control.' : '',
    futuristic_maturity: ready ? 'Mature slightly futuristic desktop utility geometry.' : '',
    non_microsoft_distinction: ready ? 'Avoids Office logos, Microsoft 365 gradients, Word silhouettes, Excel grid marks, PowerPoint slide silhouettes, Outlook envelope marks, and gear-only artwork.' : '',
    rejects_generic_readings: ready ? ['settings', 'file', 'debug console', 'ai-only', 'microsoft office clone'] : [],
    ready
  };
}

function sha256File(path: string): string {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

function renderedLogoConceptPass(ready: boolean) {
  return {
    ready,
    selected_direction: ready ? 'Command Console Panes' : '',
    minimum_concepts_reviewed: ready ? 3 : 0,
    concepts: ready ? [
      {
        name: 'Command Console Panes',
        decision: 'selected',
        rationale: 'Layered panes communicate office productivity, local routing, and deliberate user control at release sizes.'
      },
      {
        name: 'Orbiting Document Hub',
        decision: 'rejected',
        rationale: 'The hub read as a generic sync or cloud connector and lost the operator-control affordance.'
      },
      {
        name: 'Shielded Automation Badge',
        decision: 'rejected',
        rationale: 'The badge looked closer to endpoint protection software than an office control utility.'
      }
    ] : [],
    rejected_patterns: ready ? ['gear-only settings mark', 'Office-like app tile', 'host-app color block', 'generic document thumbnail', 'terminal/debug glyph', 'AI sparkle motif'] : []
  };
}
function writeManifest(dir: string, host: 'word' | 'excel' | 'powerpoint'): string {
  const path = join(dir, `${host}-manifest.xml`);
  const context = host === 'word' ? 'Word documents' : host === 'excel' ? 'Excel workbooks' : 'PowerPoint presentations';
  writeFileSync(path, `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp>
  <ProviderName>Office MCP Control</ProviderName>
  <DisplayName DefaultValue="Office MCP Control" />
  <Description DefaultValue="Control live ${context} through a local productivity automation control utility." />
  <IconUrl DefaultValue="https://localhost:8765/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:8765/assets/icon-80.png" />
  <bt:String id="OfficeMcp.GroupLabel" DefaultValue="Office MCP Control" />
</OfficeApp>
`);
  return path;
}

function writeExcelRuntimeEvidence(dir: string, ready = true): string {
  return writeHostRuntimeEvidence(dir, 'excel', 'Excel Workbook', ready);
}

function writeWordRuntimeEvidence(dir: string, ready = true): string {
  return writeHostRuntimeEvidence(dir, 'word', 'Word Document', ready);
}

function writePowerPointRuntimeEvidence(dir: string, ready = true): string {
  return writeHostRuntimeEvidence(dir, 'powerpoint', 'PowerPoint Presentation', ready);
}

function writeHostRuntimeEvidence(dir: string, app: 'word' | 'excel' | 'powerpoint', title: string, ready = true): string {
  const path = join(dir, `${app}-runtime-${ready ? 'ready' : 'broken'}.json`);
  const sessionId = `${app}-runtime-session`;
  const sessions = ready ? [{
    app,
    status: 'active',
    session_id: sessionId,
    document: { title },
    host: { app, platform: 'pc', version: '16.0' }
  }] : [];
  writeFileSync(path, JSON.stringify({
    schema_version: 1,
    endpoint: 'http://127.0.0.1:8800/mcp',
    generated_at: new Date().toISOString(),
    session_id: sessionId,
    gates: [
      { name: 'word.session_discovery', status: 'passed', details: { sessions } },
      { name: 'agent_client_stdio_bridge', status: 'passed', details: { session_count: sessions.length } },
      { name: 'irm_rights_matrix', status: 'skipped', details: { reason: 'fixture' } }
    ]
  }, null, 2));
  return path;
}
function writeOfficeToolE2eReport(dir: string, host: 'Word' | 'Excel' | 'PowerPoint', ready = true): string {
  const key = host.toLowerCase();
  const path = join(dir, `office-tool-e2e-${key}-${ready ? 'ready' : 'broken'}.json`);
  const tools = host === 'Word'
    ? ['word.get_text', 'word.insert_paragraph']
    : host === 'Excel'
      ? ['excel.get_workbook_info', 'excel.write_range']
      : ['powerpoint.get_presentation_info', 'powerpoint.add_slide'];
  const now = new Date().toISOString();
  writeFileSync(path, JSON.stringify({
    schema_version: 1,
    kind: 'office_tool_e2e_report',
    host,
    started_at: now,
    finished_at: now,
    passed: ready,
    daemon: { endpoint: 'http://127.0.0.1:8765/mcp' },
    document: { path: `${key}-fixture` },
    addin_activation: { activated: ready, activator: 'office-ui-activator', activation_path: ready ? 'official-sideload' : undefined },
    session: { session_id: `${key}-session`, available_tool_count: tools.length },
    lifecycle_counts: {
      start_daemon: 1,
      list_tools: 1,
      create_document: 1,
      activate_addin: 1,
      wait_for_session: 1,
      cleanup_document: 1,
      stop_daemon: 1
    },
    cleanup: {
      closed_by_driver: ready,
      deleted: ready,
      deleted_path_count: ready ? 1 : 0,
      deleted_paths: ready ? [`${key}-fixture`] : []
    },
    advertised_tools: tools,
    session_available_tools: tools,
    executed_tools: ready ? tools : tools.slice(0, -1),
    tool_runs: tools.map((tool, index) => ({
      id: `e2e-${tool.replace(/[^a-z0-9]+/gi, '-')}`,
      tool,
      setup_action_count: 1,
      verifier: { kind: index === 0 ? 'direct-result' : 'readback', expectation_keys: ['contains'] },
      passed: ready
    }))
  }, null, 2));
  return path;
}

function writeSurfaceScreenshot(dir: string, name: string): string {
  const path = join(dir, name);
  writeFileSync(path, tinyPng());
  return path;
}

function writeManualTrayEvidence(dir: string, ready = true, status: 'Up' | 'Degraded' = 'Up'): string {
  const traySurfaceScreenshotPaths = Object.fromEntries(TRAY_SURFACES.map((surface) => [surface.replaceAll('-', '_'), writeSurfaceScreenshot(dir, `${surface}.png`)]));
  const primaryScreenshotPath = traySurfaceScreenshotPaths.tray_icon;
  const traySurfaceScreenshotsExist = Object.fromEntries(TRAY_SURFACES.map((surface) => [surface.replaceAll('-', '_'), ready]));
  const traySurfaceScreenshotsFresh = Object.fromEntries(TRAY_SURFACES.map((surface) => [surface.replaceAll('-', '_'), ready]));
  const path = join(dir, `manual-tray-${ready ? 'ready' : 'broken'}.json`);
  writeFileSync(path, JSON.stringify({
    schema_version: 1,
    kind: 'tray_manual_evidence',
    platform: 'win32',
    visible_icon: ready,
    right_click_menu: ready,
    menu_opened_from_tray_icon: ready,
    native_menu_appearance_reviewed: ready,
    menu_anchored_to_tray_icon: ready,
    os_native_menu_behavior_reviewed: ready,
    keyboard_menu_access_reviewed: ready,
    native_quit_confirmation_reviewed: ready,
    native_tray_interaction_ready: ready,
    tray_menu_surface_kind: ready ? 'native' : 'webview',
    tray_menu_surface_native: ready,
    show_ui_opened: ready,
    observed_menu_items: ready ? [`Status: ${status}`, 'Clients: 0', 'Documents: 0', 'Show Office MCP Control', 'Quit Office MCP Control'] : ['Status: Up'],
    observed_tooltip: `Office MCP Control - ${status} - 0 clients - 0 documents`,
    screenshot_path: primaryScreenshotPath,
    primary_screenshot_matches_tray_icon: ready,
    tray_surface_screenshot_paths: traySurfaceScreenshotPaths,
    tray_surface_screenshots_exist: traySurfaceScreenshotsExist,
    tray_surface_screenshots_ready: ready,
    tray_surface_screenshot_metadata: manualTrayScreenshotMetadataFor(traySurfaceScreenshotPaths, ready),
    tray_surface_screenshots_fresh: traySurfaceScreenshotsFresh,
    tray_surface_screenshots_fresh_ready: ready,
    tray_surface_screenshots_distinct: ready,
    daemon_context: manualTrayDaemonContext(ready, status),
    daemon_context_ready: ready,
    observed_snapshot_binding_ready: ready,
    passed: ready
  }, null, 2));
  return path;
}

function manualTrayScreenshotMetadataFor(screenshots: Record<string, string>, ready: boolean): Record<string, Record<string, unknown>> {
  return Object.fromEntries(Object.entries(screenshots).map(([surface, path]) => {
    if (!ready) return [surface, { path, ready: false, fresh: false }];
    const stats = statSync(path);
    return [surface, {
      path,
      size_bytes: stats.size,
      mtime: stats.mtime.toISOString(),
      recorded_at: new Date().toISOString(),
      age_ms: 0,
      freshness_window_ms: 30 * 60 * 1000,
      fresh: true,
      ready: true
    }];
  }));
}

function writeCatalogIdentityReview(dir: string, ready = true): string {
  const path = join(dir, `catalog-identity-${ready ? 'ready' : 'broken'}.json`);
  const hosts = Object.fromEntries(['word', 'excel', 'powerpoint'].map((host) => [host, catalogIdentityHost(host, ready)]));
  writeFileSync(path, JSON.stringify({
    schema_version: 1,
    kind: 'catalog_identity_review',
    product_name: 'Office MCP Control',
    catalog_path: dir,
    catalog_type: ready ? 'Local productivity automation control utility' : 'Task Pane Add-in protocol bridge',
    shared_origin: ready ? 'https://localhost:8765' : null,
    hosts,
    ready,
    failures: ready ? [] : ['Catalog type is not product-ready.']
  }, null, 2));
  return path;
}

function catalogIdentityHost(host: string, ready = true): Record<string, unknown> {
  const taskpanePath = host === 'powerpoint' ? '/powerpoint/taskpane.html' : `/${host}/taskpane.html`;
  return {
    key: host,
    label: host === 'word' ? 'Word' : host === 'excel' ? 'Excel' : 'PowerPoint',
    display_name: ready ? 'Office MCP Control' : `office-mcp-${host}`,
    provider: ready ? 'Office MCP Control' : 'office-mcp',
    description: ready ? 'Control live documents through a local productivity automation control utility.' : 'Experimental protocol bridge debug panel.',
    icon_url: ready ? 'https://localhost:8765/assets/icon-32.png' : 'https://localhost:8765/assets/blank.png',
    high_resolution_icon_url: ready ? 'https://localhost:8765/assets/icon-80.png' : 'https://localhost:8765/assets/blank.png',
    group_label: ready ? 'Office MCP Control' : 'Office MCP',
    command_label: ready ? 'Open Control Panel' : 'Open',
    tooltip: ready ? 'Open Office MCP Control.' : 'Open task pane.',
    taskpane_url: `https://localhost:8765${taskpanePath}?v=0.1.0`,
    origin: ready ? 'https://localhost:8765' : null,
    ready,
    failures: ready ? [] : ['Prototype metadata.']
  };
}

function manualTrayDaemonContext(ready = true, status: 'Up' | 'Degraded' = 'Up') {
  return {
    status: { ok: ready, running: ready, uiUrl: 'https://localhost:8765/ui/' },
    tray_probe: {
      ok: ready,
      native_host: ready,
      state_fetch_ok: ready,
      snapshot: {
        tooltip: `Office MCP Control - ${status} - 0 clients - 0 documents`,
        menu_items: [`Status: ${status}`, 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP Control', 'Quit Office MCP Control'],
        menu: structuredTrayMenu(status)
      }
    }
  };
}

function fakeDaemonScript(stateFetchOk: boolean, menuItems: string[], includeStructuredMenu: boolean): string {
  const status = JSON.stringify({ running: true, uiUrl: 'https://localhost:8765/ui/' });
  const stateLabel = menuItems.find((item) => item.startsWith('Status: '))?.slice('Status: '.length) ?? 'Up';
  const snapshot: Record<string, unknown> = {
    tooltip: `Office MCP Control - ${stateLabel} - 0 clients - 0 documents`,
    menu_items: menuItems
  };
  if (includeStructuredMenu) snapshot.menu = structuredTrayMenu(stateLabel as 'Up' | 'Degraded');
  const trayProbe = JSON.stringify({
    native_host: true,
    state_fetch_ok: stateFetchOk,
    snapshot
  });
  if (process.platform === 'win32') {
    return `@echo off\r\nif "%1"=="daemon" echo ${status}\r\nif "%1"=="tray" echo ${trayProbe}\r\n`;
  }
  return `#!/bin/sh\nif [ "$1" = "daemon" ]; then printf '%s\\n' '${status}'; fi\nif [ "$1" = "tray" ]; then printf '%s\\n' '${trayProbe}'; fi\n`;
}

function structuredTrayMenu(status: 'Up' | 'Degraded' = 'Up'): Array<Record<string, unknown>> {
  return [
    { kind: 'read_only', enabled: false, label: `Status: ${status}` },
    { kind: 'read_only', enabled: false, label: 'Clients: 0' },
    { kind: 'read_only', enabled: false, label: 'Documents: 0' },
    { kind: 'separator', enabled: false, label: '---' },
    { kind: 'action', enabled: true, label: 'Show Office MCP Control', action: 'show_ui' },
    { kind: 'action', enabled: true, label: 'Quit Office MCP Control', action: 'quit' }
  ];
}
