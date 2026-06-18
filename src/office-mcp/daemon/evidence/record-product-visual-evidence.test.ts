import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import test from 'node:test';
import { tinyPng } from './image-evidence.js';

const TSX = './node_modules/tsx/dist/cli.mjs';
const RECORDER = resolve(process.cwd(), 'record-product-visual-evidence.ts');
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
    const excelRuntimeEvidencePath = writeExcelRuntimeEvidence(dir);
    const powerPointRuntimeEvidencePath = writePowerPointRuntimeEvidence(dir);
    const output = join(dir, 'product-visual-evidence.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--excel-runtime-evidence-path', excelRuntimeEvidencePath, '--powerpoint-runtime-evidence-path', powerPointRuntimeEvidencePath);
    assert.equal(passing.status, 0, outputText(passing.stderr) || outputText(passing.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.kind, 'product_visual_evidence');
    assert.equal(evidence.product_text_ready, true);
    assert.equal(evidence.catalog_type_ready, true);
    assert.equal(evidence.tray_tooltip_ready, true);
    assert.equal(evidence.tray_menu_surface_kind, 'native');
    assert.equal(evidence.tray_menu_surface_native, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).density_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal((evidence.product_identity_review as Record<string, unknown>).ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).excel.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).powerpoint.ready, true);
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.display_name, 'Office MCP Control');
    assert.equal((evidence.first_run_identity as Record<string, Record<string, unknown>>).word.icon_url, 'https://localhost:8765/assets/icon-32.png');
    assert.equal(evidence.rendered_logo_review_ready, true);
    assert.equal(evidence.powerpoint_runtime_evidence_ready, true);
    assert.equal(evidence.daemon_context_ready, true);
    assert.equal(evidence.passed, true);

    const missingScreenshots = { ...screenshots };
    missingScreenshots['tray-native-menu'] = join(dir, 'missing-tray-menu.png');
    const missingTray = runRecorder(join(dir, 'missing-tray.json'), missingScreenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--excel-runtime-evidence-path', excelRuntimeEvidencePath, '--powerpoint-runtime-evidence-path', powerPointRuntimeEvidencePath);
    assert.notEqual(missingTray.status, 0);
    const failed = JSON.parse(outputText(missingTray.stdout)) as Record<string, unknown>;
    assert.equal(failed.passed, false);
    assert.equal((failed.screenshots_exist as Record<string, unknown>).tray_native_menu, false);
  });
});

test('README product visual evidence command matches current PowerPoint gates', () => {
  const readme = readFileSync(resolve(process.cwd(), '../../../..', 'README.md'), 'utf8');
  const commandLine = readme.split('\n').find((line) => line.includes('npm run evidence:record-product-visual')) ?? '';

  for (const required of [
    '--powerpoint-runtime-evidence-path',
    '--powerpoint-ribbon-command',
    '--powerpoint-ribbon-command-screenshot',
    '--powerpoint-catalog-entry',
    '--powerpoint-catalog-entry-screenshot',
    '--powerpoint-taskpane-title',
    '--powerpoint-taskpane-title-screenshot',
    '--powerpoint-catalog-provider',
    '--powerpoint-catalog-description',
    '--powerpoint-catalog-type',
    '--powerpoint-first-run-identity-reviewed'
  ]) {
    assert.match(commandLine, new RegExp(required));
  }
});

test('README describes current Word Excel and PowerPoint product surface', () => {
  const readme = readFileSync(resolve(process.cwd(), '../../../..', 'README.md'), 'utf8');

  assert.match(readme, /exposes Word, Excel, and PowerPoint \(with Outlook planned\)/);
  assert.match(readme, /implementation is in place for Word, Excel, and PowerPoint/);
  assert.match(readme, /Word, Excel, and PowerPoint task pane add-ins/);
  assert.match(readme, /doc\/spec\/04-excel-capabilities\.md/);
  assert.match(readme, /PowerPoint v1 presentation tools/);
  assert.match(readme, /cd \.\.\\excel\s+npm run check\s+cd \.\.\\powerpoint\s+npm run check/);
  assert.match(readme, /src\/office-ctl\/powerpoint\/` \| PowerPoint add-in package/);
  assert.match(readme, /Word task pane: `https:\/\/localhost:8765\/word\/taskpane\.html`/);
  assert.match(readme, /Excel task pane: `https:\/\/localhost:8765\/excel\/taskpane\.html`/);
  assert.match(readme, /PowerPoint task pane: `https:\/\/localhost:8765\/powerpoint\/taskpane\.html`/);
  assert.match(readme, /`office-ctl\/word\/`, `office-ctl\/excel\/`, and\s+`office-ctl\/powerpoint\/`/);
  assert.doesNotMatch(readme, /PowerPoint add-in scaffold/);
  assert.doesNotMatch(readme, /PowerPoint and Outlook planned/);
  assert.doesNotMatch(readme, /Add-in task pane: `https:\/\/localhost:8765\/taskpane\.html`/);
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
    const manualTrayEvidencePath = writeManualTrayEvidence(dir);
    const output = join(dir, 'derived-tray-surfaces.json');
    const passing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--manual-tray-evidence-path', manualTrayEvidencePath, '--skip-tray-surface-args');
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
      if (key === 'tray_tooltip') assert.equal(observations[key], 'Office MCP - Up - 0 clients - 0 documents');
      else assert.match(observations[key], /Office MCP Control manual tray evidence/);
    }
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

test('product visual evidence recorder requires Excel runtime evidence', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
    const output = join(dir, 'missing-excel-runtime-evidence.json');
    const missing = runRecorder(output, screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--excel-runtime-evidence-path', join(dir, 'missing-excel.json'));
    assert.notEqual(missing.status, 0);
    let evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).density_ready, false);

    const broken = runRecorder(join(dir, 'broken-excel-runtime-evidence.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--excel-runtime-evidence-path', writeExcelRuntimeEvidence(dir, false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(outputText(broken.stdout)) as Record<string, unknown>;
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, false);
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

    const broken = runRecorder(join(dir, 'broken-powerpoint-runtime-evidence.json'), screenshots, '--daemon-bin', daemonBin, '--rendered-logo-review-path', renderedLogoReviewPath, '--powerpoint-runtime-evidence-path', writePowerPointRuntimeEvidence(dir, false));
    assert.notEqual(broken.status, 0);
    evidence = JSON.parse(outputText(broken.stdout)) as Record<string, unknown>;
    assert.equal(evidence.powerpoint_runtime_evidence_ready, false);
  });
});

test('product visual evidence recorder reads evidence artifact paths from environment', () => {
  withScreenshots((dir, screenshots) => {
    const daemonBin = writeFakeDaemon(dir);
    const renderedLogoReviewPath = writeRenderedLogoReview(dir);
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
      '--env-excel-runtime-evidence-path', excelRuntimeEvidencePath,
      '--env-powerpoint-runtime-evidence-path', powerPointRuntimeEvidencePath,
      '--env-manual-tray-evidence-path', manualTrayEvidencePath
    );

    assert.equal(result.status, 0, outputText(result.stderr) || outputText(result.stdout));
    const evidence = JSON.parse(readFileSync(output, 'utf8')) as Record<string, unknown>;
    assert.equal(evidence.rendered_logo_review_ready, true);
    assert.equal((evidence.excel_taskpane as Record<string, unknown>).runtime_evidence_ready, true);
    assert.equal(evidence.powerpoint_runtime_evidence_ready, true);
    assert.equal(evidence.manual_tray_evidence_ready, true);
    assert.equal(evidence.passed, true);
  });
});

function runRecorder(output: string, screenshots: Record<string, string>, ...extra: string[]): ReturnType<typeof spawnSync> {
  const skipProductReviewFlags = extra.includes('--skip-product-review-flags');
  const skipRenderedLogoAndFirstRunFlags = extra.includes('--skip-rendered-logo-and-first-run-flags');
  const skipLogoSurfaceArgs = extra.includes('--skip-logo-surface-args');
  const skipTraySurfaceArgs = extra.includes('--skip-tray-surface-args');
  const skipManualTrayEvidence = extra.includes('--skip-manual-tray-evidence');
  const envRenderedLogoReviewPath = optionValue(extra, '--env-rendered-logo-review-path');
  const envExcelRuntimeEvidencePath = optionValue(extra, '--env-excel-runtime-evidence-path');
  const envPowerPointRuntimeEvidencePath = optionValue(extra, '--env-powerpoint-runtime-evidence-path');
  const envManualTrayEvidencePath = optionValue(extra, '--env-manual-tray-evidence-path');
  const explicitTrayMenuSurfaceKind = extra.includes('--tray-menu-surface-kind');
  const explicitCatalogType = extra.includes('--catalog-type');
  const explicitWordCatalogType = extra.includes('--word-catalog-type');
  const explicitExcelCatalogType = extra.includes('--excel-catalog-type');
  const explicitPowerPointCatalogType = extra.includes('--powerpoint-catalog-type');
  const filteredExtra = extra.filter((item, index) => {
    const previous = extra[index - 1];
    if (previous === '--env-rendered-logo-review-path' || previous === '--env-excel-runtime-evidence-path' || previous === '--env-powerpoint-runtime-evidence-path' || previous === '--env-manual-tray-evidence-path') return false;
    return item !== '--skip-product-review-flags'
      && item !== '--skip-rendered-logo-and-first-run-flags'
      && item !== '--skip-logo-surface-args'
      && item !== '--skip-tray-surface-args'
      && item !== '--skip-manual-tray-evidence'
      && item !== '--env-rendered-logo-review-path'
      && item !== '--env-excel-runtime-evidence-path'
      && item !== '--env-powerpoint-runtime-evidence-path'
      && item !== '--env-manual-tray-evidence-path';
  });
  const hasWordManifest = filteredExtra.includes('--word-manifest-path');
  const hasExcelManifest = filteredExtra.includes('--excel-manifest-path');
  const hasPowerPointManifest = filteredExtra.includes('--powerpoint-manifest-path');
  const hasExcelRuntimeEvidence = filteredExtra.includes('--excel-runtime-evidence-path');
  const hasPowerPointRuntimeEvidence = filteredExtra.includes('--powerpoint-runtime-evidence-path');
  const hasManualTrayEvidence = filteredExtra.includes('--manual-tray-evidence-path');
  const outputDir = dirname(output);
  const args = [
    TSX,
    RECORDER,
    '--output', output,
    ...(explicitCatalogType ? [] : ['--catalog-type', 'Local productivity automation control utility']),
    '--catalog-icon-visible', 'true',
    '--tray-tooltip', 'Office MCP - Up - 0 clients - 0 documents',
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
    '--excel-compact-top-block', 'true',
    '--excel-tools-permissions-merged', 'true',
    '--excel-inline-settings', 'true',
    '--excel-server-protocol-row', 'Server 0.1.0 / Protocol 1.0',
    '--excel-document-state', 'Editable'
  ];
  if (!hasWordManifest) args.push('--word-manifest-path', writeManifest(outputDir, 'word'));
  if (!hasExcelManifest) args.push('--excel-manifest-path', writeManifest(outputDir, 'excel'));
  if (!hasPowerPointManifest) args.push('--powerpoint-manifest-path', writeManifest(outputDir, 'powerpoint'));
  if (!hasExcelRuntimeEvidence) args.push('--excel-runtime-evidence-path', writeExcelRuntimeEvidence(outputDir));
  if (!hasPowerPointRuntimeEvidence) args.push('--powerpoint-runtime-evidence-path', writePowerPointRuntimeEvidence(outputDir));
  if (!hasManualTrayEvidence && !skipManualTrayEvidence) args.push('--manual-tray-evidence-path', writeManualTrayEvidence(outputDir));
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
    ...(envExcelRuntimeEvidencePath ? { OFFICE_MCP_EXCEL_RUNTIME_EVIDENCE_PATH: envExcelRuntimeEvidencePath } : {}),
    ...(envPowerPointRuntimeEvidencePath ? { OFFICE_MCP_POWERPOINT_RUNTIME_EVIDENCE_PATH: envPowerPointRuntimeEvidencePath } : {}),
    ...(envManualTrayEvidencePath ? { OFFICE_MCP_TRAY_MANUAL_EVIDENCE_PATH: envManualTrayEvidencePath } : {})
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
</OfficeApp>
`);
  return path;
}

function writeExcelRuntimeEvidence(dir: string, ready = true): string {
  const path = join(dir, `excel-runtime-${ready ? 'ready' : 'broken'}.json`);
  const sessionId = '11111111-2222-3333-4444-555555555555';
  const sessions = ready ? [{
    app: 'excel',
    available_tool_count: 7,
    document: { title: 'Excel Workbook' },
    host: { app: 'excel', platform: 'pc', version: '16.0' },
    session_id: sessionId,
    status: 'active'
  }] : [];
  const smokeDetails = {
    session_id: sessionId,
    marker_found: ready,
    write: { wrote_values: ready },
    formula: { wrote_formula: ready },
    format: { formatted: ready },
    table: ready ? { table: 'OfficeMcpTable' } : {},
    chart: ready ? { chart: 'Chart 1' } : {},
    sheet: { activated: ready }
  };
  writeFileSync(path, JSON.stringify({
    schema_version: 1,
    generated_at: new Date().toISOString(),
    endpoint: 'http://127.0.0.1:8800/mcp',
    gates: [
      { name: 'word.session_discovery', status: 'passed', details: { sessions } },
      { name: 'excel.runtime_smoke', status: ready ? 'passed' : 'failed', details: smokeDetails }
    ]
  }, null, 2));
  return path;
}

function writePowerPointRuntimeEvidence(dir: string, ready = true): string {
  const path = join(dir, `powerpoint-runtime-${ready ? 'ready' : 'broken'}.json`);
  const sessionId = '22222222-3333-4444-5555-666666666666';
  const sessions = ready ? [{
    app: 'powerpoint',
    available_tool_count: 5,
    document: { title: 'PowerPoint Presentation' },
    host: { app: 'powerpoint', platform: 'pc', version: '16.0' },
    session_id: sessionId,
    status: 'active'
  }] : [];
  const smokeDetails = {
    session_id: sessionId,
    available_tool_count: ready ? 5 : 0,
    add_slide: ready ? { slide_id: 'slide-1', slide_index: 0 } : {},
    replace_text: { replacements: ready ? 1 : 0 },
    layout: ready ? { slide_id: 'slide-1', slide_index: 0, layout_name: 'Title Only' } : {},
    mutation_proved: ready,
    pdf_supported: false,
    pdf_host_rejection: ready
  };
  writeFileSync(path, JSON.stringify({
    schema_version: 1,
    generated_at: new Date().toISOString(),
    endpoint: 'http://127.0.0.1:8800/mcp',
    gates: [
      { name: 'word.session_discovery', status: 'passed', details: { sessions } },
      { name: 'powerpoint.runtime_smoke', status: ready ? 'passed' : 'failed', details: smokeDetails }
    ]
  }, null, 2));
  return path;
}

function writeSurfaceScreenshot(dir: string, name: string): string {
  const path = join(dir, name);
  writeFileSync(path, tinyPng());
  return path;
}

function writeManualTrayEvidence(dir: string, ready = true): string {
  const screenshotPath = join(dir, `manual-tray-${ready ? 'ready' : 'broken'}.png`);
  writeFileSync(screenshotPath, tinyPng());
  const traySurfaceScreenshotPaths = Object.fromEntries(TRAY_SURFACES.map((surface) => [surface.replaceAll('-', '_'), writeSurfaceScreenshot(dir, `${surface}.png`)]));
  const traySurfaceScreenshotsExist = Object.fromEntries(TRAY_SURFACES.map((surface) => [surface.replaceAll('-', '_'), ready]));
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
    observed_menu_items: ready ? ['Status: Up', 'Clients: 0', 'Documents: 0', 'Show Office MCP', 'Quit Office MCP'] : ['Status: Up'],
    observed_tooltip: 'Office MCP - Up - 0 clients - 0 documents',
    screenshot_path: screenshotPath,
    tray_surface_screenshot_paths: traySurfaceScreenshotPaths,
    tray_surface_screenshots_exist: traySurfaceScreenshotsExist,
    tray_surface_screenshots_ready: ready,
    daemon_context: manualTrayDaemonContext(ready),
    daemon_context_ready: ready,
    passed: ready
  }, null, 2));
  return path;
}

function manualTrayDaemonContext(ready = true) {
  return {
    status: { ok: ready, running: ready, uiUrl: 'https://localhost:8765/ui/' },
    tray_probe: {
      ok: ready,
      native_host: ready,
      state_fetch_ok: ready,
      snapshot: {
        tooltip: 'Office MCP - Up - 0 clients - 0 documents',
        menu_items: ['Status: Up', 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP', 'Quit Office MCP']
      }
    }
  };
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
