import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';
import { tinyPng } from './image-evidence.js';

test('runtime evidence validator accepts required runtime gates and optional IRM skip', () => {
  withEvidenceFile(report('skipped'), (path) => {
    const result = runValidator(path);
    assert.equal(result.status, 0, outputText(result.stderr));
    assert.equal(JSON.parse(outputText(result.stdout)).ok, true);
  });
});

test('runtime evidence validator can require IRM evidence', () => {
  withEvidenceFile(report('skipped'), (path) => {
    const result = runValidator(path, '--require-irm');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /IRM gate is skipped/);
  });

  withEvidenceFile(report('passed'), (path) => {
    const result = runValidator(path, '--require-irm');
    assert.equal(result.status, 0, outputText(result.stderr));
  });
});

test('runtime evidence validator can require full Word smoke evidence', () => {
  withEvidenceFile(report('skipped'), (path) => {
    const result = runValidator(path, '--require-full-word-smoke');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: word\.full_smoke\.word-core/);
  });

  const full = report('skipped');
  full.gates.push(
    gate('word.full_smoke.word-core', 'passed'),
    gate('word.full_smoke.word-formatting', 'passed'),
    gate('word.full_smoke.word-review', 'passed'),
    gate('word.full_smoke.word-resources', 'passed'),
    gate('word.full_smoke.word-spec-args', 'passed')
  );
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-full-word-smoke');
    assert.equal(result.status, 0, outputText(result.stderr));
  });
});

test('runtime evidence validator can require Excel smoke evidence', () => {
  withEvidenceFile(report('skipped'), (path) => {
    const result = runValidator(path, '--require-excel-smoke');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: excel\.runtime_smoke/);
  });

  const full = report('skipped');
  full.gates.push(gate('excel.runtime_smoke', 'passed'));
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-excel-smoke');
    assert.equal(result.status, 0, outputText(result.stderr));
  });
});

test('runtime evidence validator accepts Excel-only smoke without Word runtime gates', () => {
  const excelOnly = excelOnlyReport();
  withEvidenceFile(excelOnly, (path) => {
    const result = runValidator(path, '--require-excel-smoke');
    assert.equal(result.status, 0, outputText(result.stdout) + outputText(result.stderr));
    const summary = JSON.parse(outputText(result.stdout)) as { ok: boolean; session_id?: string };
    assert.equal(summary.ok, true);
    assert.equal(summary.session_id, undefined);
  });

  excelOnly.gates = excelOnly.gates.filter((item) => item.name !== 'excel.runtime_smoke');
  withEvidenceFile(excelOnly, (path) => {
    const result = runValidator(path, '--require-excel-smoke');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: excel\.runtime_smoke/);
  });
});

test('runtime evidence validator can require COM tracked-change evidence', () => {
  withEvidenceFile(report('skipped'), (path) => {
    const result = runValidator(path, '--require-com-tracked-changes');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: word\.tracked_change_com\.accept/);
  });

  const full = report('skipped');
  full.gates.push(
    gate('word.tracked_change_com.accept', 'passed'),
    gate('word.tracked_change_com.reject', 'passed')
  );
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-com-tracked-changes');
    assert.equal(result.status, 0, outputText(result.stderr));
  });
});

test('runtime evidence validator can require external preflight evidence', () => {
  withEvidenceFile(report('skipped'), (path) => {
    const result = runValidator(path, '--require-irm-preflight', '--require-claude-desktop-installation');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: irm_document_preflight/);
    assert.match(outputText(result.stdout), /Missing required gate: claude_desktop_installation/);
  });

  const full = report('skipped');
  full.gates.push(
    gate('irm_document_preflight', 'passed'),
    gate('claude_desktop_installation', 'passed', { ui_validation_ready: false })
  );
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-irm-preflight', '--require-claude-desktop-installation');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /not UI-validation ready/);
  });

  full.gates = full.gates.filter((item) => item.name !== 'claude_desktop_installation');
  full.gates.push(gate('claude_desktop_installation', 'passed', { config_has_office_mcp: true, app_detected: false, ui_validation_ready: false }));
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-irm-preflight', '--require-claude-desktop-installation');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /not UI-validation ready/);
  });

  full.gates = full.gates.filter((item) => item.name !== 'claude_desktop_installation');
  full.gates.push(gate('claude_desktop_installation', 'passed', { config_has_office_mcp: true, app_detected: true, ui_validation_ready: true }));
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-irm-preflight', '--require-claude-desktop-installation');
    assert.equal(result.status, 0, outputText(result.stderr));
  });
});

test('runtime evidence validator can require agent client evidence', () => {
  withEvidenceFile(report('passed'), (path) => {
    const result = runValidator(path, '--require-agent-client-prompt');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: agent_client_prompt/);
  });

  const full = report('passed');
  full.gates.push(gate('agent_client_prompt', 'passed', {
    kind: 'agent_client_prompt',
    prompt: 'what does paragraph 1 of my open Word doc say?',
    observed_answer: 'Document 2026-05-31 (6)',
    passed: true
  }));
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-agent-client-prompt');
    assert.equal(result.status, 0, outputText(result.stderr));
  });
});

test('runtime evidence validator can require mutation evidence', () => {
  const broken = report('passed');
  broken.gates = broken.gates.filter((item) => item.name !== 'word.runtime_mutation_smoke');

  withEvidenceFile(broken, (path) => {
    const result = runValidator(path);
    assert.equal(result.status, 0, outputText(result.stderr));
  });

  withEvidenceFile(broken, (path) => {
    const result = runValidator(path, '--require-mutation');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: word\.runtime_mutation_smoke/);
  });
});

test('runtime evidence validator accepts UI runtime evidence gates', () => {
  const ui = uiReport();
  withEvidenceFile(ui, (path) => {
    const result = runValidator(path, '--ui');
    assert.equal(result.status, 0, outputText(result.stderr));
    assert.equal(JSON.parse(outputText(result.stdout)).ok, true);
  });

  ui.gates = ui.gates.filter((item) => item.name !== 'ui.tray_probe');
  withEvidenceFile(ui, (path) => {
    const result = runValidator(path, '--ui');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: ui\.tray_probe/);
  });

  const missingProductionTray = uiReport();
  missingProductionTray.gates = missingProductionTray.gates.filter((item) => item.name !== 'ui.production_daemon_tray');
  withEvidenceFile(missingProductionTray, (path) => {
    const result = runValidator(path, '--ui');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: ui\.production_daemon_tray/);
  });
});

test('runtime evidence validator can require product visual evidence', () => {
  const ui = uiReport();
  withEvidenceFile(ui, (path) => {
    const result = runValidator(path, '--ui', '--require-product-visual');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing --product-visual-evidence-path/);
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const passing = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.equal(passing.status, 0, outputText(passing.stdout) + outputText(passing.stderr));
      assert.equal(JSON.parse(outputText(passing.stdout)).ok, true);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.observations.word_ribbon_command = 'Raw task pane';
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /observation missing product name: word_ribbon_command/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      writeFileSync(broken.screenshot_paths.tray_icon, 'not an image');
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /screenshot missing or invalid: tray_icon/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      writeFileSync(broken.screenshot_paths.tray_icon, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /screenshot missing or invalid: tray_icon/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.excel_taskpane.document_state = 'unknown';
      broken.excel_taskpane.document_state_ready = false;
      broken.excel_taskpane.density_ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /concrete Excel editable\/read-only\/protected state/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.excel_taskpane.runtime_evidence_ready = false;
      broken.excel_taskpane.runtime_evidence.smoke_details.marker_found = false;
      (broken.excel_taskpane.runtime_evidence.smoke_details as Record<string, unknown>).chart = {};
      broken.excel_taskpane.density_ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Excel runtime evidence ready flag/);
      assert.match(outputText(result.stdout), /Excel runtime evidence missing marker readback/);
      assert.match(outputText(result.stdout), /Excel runtime evidence missing create_chart proof/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.daemon_context = undefined as unknown as ReturnType<typeof manualTrayDaemonContext>;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /missing daemon context/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.daemon_context = manualTrayDaemonContext(['Status: Up', 'Clients: 0']);
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Product visual daemon context live missing menu item: Documents:/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.daemon_context = manualTrayDaemonContext(undefined, false);
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Product visual evidence daemon context tray probe did not read live UI state/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.daemon_context_ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Product visual evidence daemon context is not recorder-ready/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.product_identity_review.logo_quality_reviewed = false;
      broken.product_identity_review.ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Product visual evidence missing logo quality review/);
      assert.match(outputText(result.stdout), /Product visual evidence missing product identity review ready flag/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.product_identity_review.rendered_size_logo_reviewed = false;
      broken.product_identity_review.rendered_logo_review_ready = false;
      broken.rendered_logo_review_ready = false;
      broken.product_identity_review.word_first_run_identity_ready = false;
      broken.first_run_identity.word.type = 'Experimental protocol bridge';
      broken.first_run_identity.word.ready = false;
      broken.product_identity_review.ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /rendered-size logo review/);
      assert.match(outputText(result.stdout), /rendered logo review ready flag/);
      assert.match(outputText(result.stdout), /Word first-run identity ready flag/);
      assert.match(outputText(result.stdout), /Word local productivity automation\/control type metadata/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.rendered_logo_review.surfaces = broken.rendered_logo_review.surfaces.filter((item) => item.key !== 'logo_tray_size');
      broken.rendered_logo_review.ready = false;
      broken.rendered_logo_review_ready = false;
      broken.product_identity_review.rendered_logo_review_ready = false;
      broken.product_identity_review.ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /missing rendered logo review ready flag/);
      assert.match(outputText(result.stdout), /Rendered logo review missing surface: logo_tray_size/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.first_run_identity.excel.display_name = 'office-mcp-excel';
      broken.first_run_identity.excel.icon_url = 'https://localhost:8765/assets/blank.png';
      broken.first_run_identity.excel.manifest_ready = false;
      broken.first_run_identity.excel.ready = false;
      broken.product_identity_review.excel_first_run_identity_ready = false;
      broken.product_identity_review.ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Excel manifest-derived identity ready flag/);
      assert.match(outputText(result.stdout), /Excel display name product identity/);
      assert.match(outputText(result.stdout), /Excel first-run icon URL/);
    });
  });
});
test('runtime evidence validator can require manual Windows tray evidence', () => {
  const ui = uiReport();
  withEvidenceFile(ui, (path) => {
    const result = runValidator(path, '--ui', '--require-manual-tray');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing --manual-tray-evidence-path/);
  });

  withEvidenceFile(ui, (uiPath) => {
    withEvidenceFile(manualTrayReport(false), (manualPath) => {
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray evidence missing visible tray icon/);
      assert.match(outputText(result.stdout), /Manual tray evidence screenshot file does not exist/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const passing = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.equal(passing.status, 0, outputText(passing.stdout) + outputText(passing.stderr));
      assert.equal(JSON.parse(outputText(passing.stdout)).ok, true);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      broken.daemon_context = undefined;
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray evidence missing daemon context/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const manual = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      manual.daemon_context = manualTrayDaemonContext();
      writeFileSync(manualPath, JSON.stringify(manual, null, 2));
      const passing = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.equal(passing.status, 0, outputText(passing.stdout) + outputText(passing.stderr));
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      broken.observed_menu_items = ['Status: Up', 'Clients: 0'];
      broken.menu_contains_required_items = true;
      broken.screenshot_exists = true;
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray evidence missing menu item: Documents:/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      writeFileSync(broken.screenshot_path, 'not an image');
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray evidence screenshot file does not exist/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      writeFileSync(broken.screenshot_path, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray evidence screenshot file does not exist/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      broken.daemon_context = manualTrayDaemonContext(['Status: Up', 'Clients: 0']);
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray daemon context live missing menu item: Documents:/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      broken.daemon_context = manualTrayDaemonContext(undefined, false);
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray daemon context tray probe did not read live UI state/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      broken.daemon_context_ready = false;
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray evidence daemon context is not recorder-ready/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      broken.menu_opened_from_tray_icon = false;
      broken.native_menu_appearance_reviewed = false;
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /right-click menu opened from the notification-area tray icon/);
      assert.match(outputText(result.stdout), /native tray menu appearance review/);
    });
  });
});

test('runtime evidence validator rejects missing or failed required gates', () => {
  const broken = report('passed');
  broken.gates = broken.gates.filter((item) => item.name !== 'word.runtime_read_smoke');

  withEvidenceFile(broken, (path) => {
    const result = runValidator(path);
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: word\.runtime_read_smoke/);
  });
});

function runValidator(path: string, ...args: string[]): ReturnType<typeof spawnSync> {
  return spawnSync(process.execPath, ['./node_modules/tsx/dist/cli.mjs', evidencePath('validate-runtime-evidence.ts'), '--input', path, ...args], {
    cwd: process.cwd(),
    encoding: 'utf8'
  });
}

function evidencePath(file: string): string {
  return resolve(process.cwd(), file);
}

function outputText(value: string | Buffer): string {
  return typeof value === 'string' ? value : value.toString('utf8');
}

function withEvidenceFile(data: unknown, callback: (path: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-evidence-validator-'));
  try {
    const path = join(dir, 'runtime-evidence.json');
    writeFileSync(path, JSON.stringify(data, null, 2));
    callback(path);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function withManualTrayEvidence(passed: boolean, callback: (path: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-manual-tray-evidence-'));
  try {
    const screenshotPath = join(dir, 'tray-visible.png');
    if (passed) writeFileSync(screenshotPath, tinyPng());
    const path = join(dir, 'tray-manual-evidence.json');
    writeFileSync(path, JSON.stringify(manualTrayReport(passed, screenshotPath), null, 2));
    callback(path);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function report(irmStatus: 'passed' | 'skipped') {
  const now = new Date().toISOString();
  return {
    schema_version: 1,
    generated_at: now,
    endpoint: 'http://127.0.0.1:8800/mcp',
    session_id: '24248c3e-11a9-48b8-a922-dc4f58dbb2de',
    gates: [
      gate('word.session_discovery', 'passed'),
      gate('word.runtime_read_smoke', 'passed'),
      gate('word.runtime_mutation_smoke', 'passed'),
      gate('agent_client_stdio_bridge', 'passed'),
      gate('irm_rights_matrix', irmStatus)
    ]
  };
}

function gate(name: string, status: 'passed' | 'skipped', details: Record<string, unknown> = {}) {
  const now = new Date().toISOString();
  return { name, status, started_at: now, finished_at: now, details };
}

function uiReport() {
  const now = new Date().toISOString();
  return {
    schema_version: 1,
    generated_at: now,
    kind: 'ui_runtime_evidence',
    gates: [
      gate('ui.daemon_runtime_file', 'passed'),
      gate('ui.state_api_origin_redaction', 'passed'),
      gate('ui.events_stream', 'passed'),
      gate('ui.tray_probe', 'passed'),
      gate('ui.production_daemon_tray', 'passed'),
      gate('ui.browser_smoke', 'passed')
    ]
  };
}

function withProductVisualEvidence(passed: boolean, callback: (path: string) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-product-visual-evidence-'));
  try {
    const screenshots: Record<string, string> = {};
    for (const surface of productVisualSurfaces()) {
      const screenshotPath = join(dir, `${surface}.png`);
      if (passed) writeFileSync(screenshotPath, tinyPng());
      screenshots[surface] = screenshotPath;
    }
    const path = join(dir, 'product-visual-evidence.json');
    writeFileSync(path, JSON.stringify(productVisualReport(passed, screenshots), null, 2));
    callback(path);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function productVisualReport(passed: boolean, screenshots: Record<string, string>) {
  const observations = Object.fromEntries(productVisualSurfaces().map((surface) => [surface, `Office MCP Control ${surface}`]));
  return {
    schema_version: 1,
    kind: 'product_visual_evidence',
    recorded_at: new Date().toISOString(),
    tester: 'test',
    platform: 'win32',
    product_name: 'Office MCP Control',
    required_surfaces: productVisualSurfaces(),
    observations,
    screenshot_paths: screenshots,
    screenshots_exist: Object.fromEntries(productVisualSurfaces().map((surface) => [surface, passed])),
    product_text_ready: passed,
    catalog_type: 'Local productivity automation control utility',
    catalog_type_ready: passed,
    catalog_icon_visible: passed,
    tray_tooltip: 'Office MCP - Up - 0 clients - 0 documents',
    tray_tooltip_ready: passed,
    tray_icon_visible: passed,
    tray_menu_native: passed,
    quit_confirmation_visible: passed,
    product_identity_review: {
      logo_quality_reviewed: passed,
      rendered_size_logo_reviewed: passed,
      rendered_logo_review_ready: passed,
      addin_identity_reviewed: passed,
      word_first_run_identity_reviewed: passed,
      excel_first_run_identity_reviewed: passed,
      tray_product_polish_reviewed: passed,
      word_first_run_identity_ready: passed,
      excel_first_run_identity_ready: passed,
      ready: passed
    },
    first_run_identity: {
      word: {
        manifest_ready: passed,
        display_name: 'Office MCP Control',
        provider: 'Office MCP Control',
        description: 'Local office productivity automation and control utility',
        type: 'Local productivity automation control utility',
        icon_url: 'https://localhost:8765/assets/icon-32.png',
        high_resolution_icon_url: 'https://localhost:8765/assets/icon-80.png',
        ready: passed
      },
      excel: {
        manifest_ready: passed,
        display_name: 'Office MCP Control',
        provider: 'Office MCP Control',
        description: 'Local office productivity automation and control utility',
        type: 'Local productivity automation control utility',
        icon_url: 'https://localhost:8765/assets/icon-32.png',
        high_resolution_icon_url: 'https://localhost:8765/assets/icon-80.png',
        ready: passed
      }
    },
    rendered_logo_review: renderedLogoReview(passed, screenshots.logo_tray_size),
    rendered_logo_review_ready: passed,
    excel_taskpane: {
      compact_top_block: passed,
      tools_permissions_merged: passed,
      inline_settings: passed,
      server_protocol_row: 'Server 0.1.0 / Protocol 1.0',
      server_protocol_row_ready: passed,
      document_state: 'Editable',
      document_state_ready: passed,
      runtime_evidence: excelRuntimeEvidence(passed),
      runtime_evidence_ready: passed,
      density_ready: passed
    },
    daemon_context: manualTrayDaemonContext(),
    daemon_context_ready: passed,
    passed
  };
}

function excelRuntimeEvidence(passed: boolean) {
  const sessionId = '11111111-2222-3333-4444-555555555555';
  return {
    ok: passed,
    schema_version: 1,
    endpoint: 'http://127.0.0.1:8800/mcp',
    generated_at: new Date().toISOString(),
    smoke_passed: passed,
    ready: passed,
    session: {
      app: 'excel',
      status: 'active',
      session_id: sessionId,
      available_tool_count: 7,
      document: { title: 'Excel Workbook' },
      host: { app: 'excel', platform: 'pc', version: '16.0' }
    },
    smoke_details: {
      session_id: sessionId,
      marker_found: passed,
      write: { wrote_values: passed },
      formula: { wrote_formula: passed },
      format: { formatted: passed },
      table: { table: 'OfficeMcpTable' },
      chart: { chart: 'Chart 1' },
      sheet: { activated: passed }
    }
  };
}

function renderedLogoReview(passed: boolean, sheetPath: string) {
  return {
    ok: passed,
    schema_version: 1,
    kind: 'rendered_logo_review',
    product_name: 'Office MCP Control',
    sheet_path: sheetPath,
    ready: passed,
    surfaces: [
      ['logo_tray_size', 16],
      ['logo_ribbon_size', 32],
      ['logo_catalog_thumbnail', 80],
      ['logo_daemon_titlebar', 20],
      ['logo_installer_metadata', 256]
    ].map(([key, size]) => ({
      key,
      rendered_size_px: size,
      width: size,
      height: size,
      non_empty: passed,
      palette_ready: passed,
      expected_size_ready: passed
    }))
  };
}

function productVisualSurfaces(): string[] {
  return [
    'word_ribbon_command',
    'word_catalog_entry',
    'word_taskpane_title',
    'excel_ribbon_command',
    'excel_catalog_entry',
    'excel_taskpane_title',
    'logo_tray_size',
    'logo_ribbon_size',
    'logo_catalog_thumbnail',
    'logo_daemon_titlebar',
    'logo_installer_metadata',
    'tray_icon',
    'tray_native_menu',
    'tray_tooltip',
    'tray_quit_confirmation'
  ];
}
function manualTrayReport(passed: boolean, screenshotPath = 'C:\\temp\\tray.png') {
  return {
    schema_version: 1,
    kind: 'tray_manual_evidence',
    recorded_at: new Date().toISOString(),
    tester: 'test',
    platform: 'win32',
    visible_icon: passed,
    right_click_menu: passed,
    menu_opened_from_tray_icon: passed,
    native_menu_appearance_reviewed: passed,
    show_ui_opened: passed,
    observed_menu_items: ['Status: Up', 'Clients: 0', 'Documents: 0', 'Show Office MCP', 'Quit Office MCP'],
    observed_tooltip: 'Office MCP - Up - 0 clients - 0 documents',
    expected_menu_items: ['Status:', 'Clients:', 'Documents:', 'Show Office MCP', 'Quit Office MCP'],
    menu_contains_required_items: passed,
    screenshot_path: screenshotPath,
    screenshot_exists: passed,
    daemon_context: manualTrayDaemonContext() as ReturnType<typeof manualTrayDaemonContext> | undefined,
    daemon_context_ready: passed,
    passed
  };
}

function manualTrayDaemonContext(menuItems = ['Status: Up', 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP', 'Quit Office MCP'], stateFetchOk = true) {
  return {
    binary_path: 'C:\\Code\\office-mcp\\target\\debug\\office-mcp-daemon.exe',
    status: {
      ok: true,
      running: true,
      pid: 1234,
      uiUrl: 'https://localhost:8765/ui/'
    },
    tray_probe: {
      ok: true,
      native_host: true,
      state_fetch_ok: stateFetchOk,
      snapshot: {
        menu_items: menuItems,
        menu: structuredMenu(menuItems),
        tooltip: 'Office MCP - Up - 0 clients - 0 documents',
        platform: 'windows-notification-area'
      }
    }
  };
}

function excelOnlyReport() {
  const now = new Date().toISOString();
  return {
    schema_version: 1,
    generated_at: now,
    endpoint: 'http://127.0.0.1:8800/mcp',
    gates: [
      gate('word.session_discovery', 'passed'),
      gate('excel.runtime_smoke', 'passed')
    ]
  };
}

function structuredMenu(menuItems: string[]) {
  return menuItems.map((label, index) => {
    if (label === '---') return { kind: 'separator', label, enabled: false };
    if (label === 'Show Office MCP') return { kind: 'action', label, action: 'show_ui', enabled: true };
    if (label === 'Quit Office MCP') return { kind: 'action', label, action: 'quit', enabled: true };
    return { kind: 'read_only', label, enabled: false };
  });
}
