import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';
import { tinyPng } from './image-evidence.js';

const REPO_ROOT = resolve(process.cwd(), '../../../..');
const ASSET_ROOT = resolve(REPO_ROOT, 'src/office-ctl/common/assets');

const POWERPOINT_V1_TOOLS = [
  'powerpoint.get_presentation_info',
  'powerpoint.get_active_view',
  'powerpoint.export_file',
  'powerpoint.update_tags',
  'powerpoint.list_slides',
  'powerpoint.add_slide',
  'powerpoint.update_slide',
  'powerpoint.delete_slide',
  'powerpoint.move_slide',
  'powerpoint.export_slide',
  'powerpoint.list_layouts',
  'powerpoint.apply_layout',
  'powerpoint.get_selection',
  'powerpoint.set_selection',
  'powerpoint.list_shapes',
  'powerpoint.add_text_box',
  'powerpoint.add_shape',
  'powerpoint.insert_image',
  'powerpoint.update_shape',
  'powerpoint.read_text',
  'powerpoint.replace_text',
  'powerpoint.format_text',
  'powerpoint.add_table',
  'powerpoint.read_table',
  'powerpoint.update_table'
];

const EXCEL_V1_TOOLS = [
  'excel.get_workbook_info',
  'excel.list_sheets',
  'excel.add_sheet',
  'excel.update_sheet',
  'excel.delete_sheet',
  'excel.get_used_range',
  'excel.read_range',
  'excel.write_range',
  'excel.clear_range',
  'excel.find_replace_cells',
  'excel.set_formula',
  'excel.format_range',
  'excel.sort_range',
  'excel.apply_filter',
  'excel.create_table',
  'excel.update_table',
  'excel.create_chart',
  'excel.update_chart',
  'excel.create_pivot_table',
  'excel.update_pivot_table'
];

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

  const full = excelOnlyReport();
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-excel-smoke');
    assert.equal(result.status, 0, outputText(result.stderr));
  });

  const weak = excelOnlyReport();
  const weakDiscovery = weak.gates[0].details as { sessions: Array<{ available_tool_count: number }> };
  const weakSmoke = weak.gates[1].details as Record<string, unknown>;
  weakDiscovery.sessions[0].available_tool_count = 7;
  weakSmoke.available_tool_count = 7;
  weakSmoke.available_tools = EXCEL_V1_TOOLS.slice(0, 7);
  delete weakSmoke.sort;
  delete weakSmoke.pivot_table;
  withEvidenceFile(weak, (path) => {
    const result = runValidator(path, '--require-excel-smoke');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Excel smoke gate missing 20-tool available tool count/);
    assert.match(outputText(result.stdout), /Excel smoke gate available tools are not aligned with v1 catalog/);
    assert.match(outputText(result.stdout), /Excel smoke gate missing sort_range proof/);
    assert.match(outputText(result.stdout), /Excel smoke gate missing create_pivot_table proof/);
  });
});

test('runtime evidence validator rejects passed Excel smoke gates without v1 proof details', () => {
  const weak = report('skipped');
  weak.gates.push(gate('excel.runtime_smoke', 'passed'));
  withEvidenceFile(weak, (path) => {
    const result = runValidator(path, '--require-excel-smoke');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Excel smoke gate missing session_id/);
    assert.match(outputText(result.stdout), /Excel smoke gate missing 20-tool available tool count/);
    assert.match(outputText(result.stdout), /Excel smoke gate missing marker readback/);
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

test('runtime evidence validator can require PowerPoint smoke evidence', () => {
  withEvidenceFile(report('skipped'), (path) => {
    const result = runValidator(path, '--require-powerpoint-smoke');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing required gate: powerpoint\.runtime_smoke/);
  });

  const full = powerpointOnlyReport();
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-powerpoint-smoke');
    assert.equal(result.status, 0, outputText(result.stdout) + outputText(result.stderr));
    const summary = JSON.parse(outputText(result.stdout)) as { ok: boolean; session_id?: string; require_powerpoint_smoke?: boolean };
    assert.equal(summary.ok, true);
    assert.equal(summary.session_id, undefined);
    assert.equal(summary.require_powerpoint_smoke, true);
  });

  full.gates[1].details.export_supported = false;
  full.gates[1].details.export_host_rejection = false;
  withEvidenceFile(full, (path) => {
    const result = runValidator(path, '--require-powerpoint-smoke');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /export_file success or explicit host-capability rejection/);
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

test('runtime evidence validator can require Office tool E2E reports', () => {
  withEvidenceFile(report('passed'), (path) => {
    const result = runValidator(path, '--require-office-tool-e2e');
    assert.notEqual(result.status, 0);
    assert.match(outputText(result.stdout), /Missing --word-tool-e2e-report-path/);
    assert.match(outputText(result.stdout), /Missing --excel-tool-e2e-report-path/);
    assert.match(outputText(result.stdout), /Missing --powerpoint-tool-e2e-report-path/);
  });

  withEvidenceFile(report('passed'), (path) => {
    withOfficeToolE2eReports(true, (reports) => {
      const result = runValidator(
        path,
        '--require-office-tool-e2e',
        '--word-tool-e2e-report-path', reports.word,
        '--excel-tool-e2e-report-path', reports.excel,
        '--powerpoint-tool-e2e-report-path', reports.powerpoint
      );
      assert.equal(result.status, 0, outputText(result.stdout) + outputText(result.stderr));
      const summary = JSON.parse(outputText(result.stdout)) as { ok: boolean; require_office_tool_e2e?: boolean };
      assert.equal(summary.ok, true);
      assert.equal(summary.require_office_tool_e2e, true);
    });
  });

  withEvidenceFile(report('passed'), (path) => {
    withOfficeToolE2eReports(true, (reports) => {
      const broken = JSON.parse(readFileSync(reports.word, 'utf8')) as ReturnType<typeof officeToolE2eReport>;
      broken.lifecycle_counts.create_document = 2;
      writeFileSync(reports.word, JSON.stringify(broken, null, 2));
      const result = runValidator(path, '--require-office-tool-e2e', '--word-tool-e2e-report-path', reports.word, '--excel-tool-e2e-report-path', reports.excel, '--powerpoint-tool-e2e-report-path', reports.powerpoint);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Word Office tool E2E report lifecycle create_document is 2, expected 1/);
    });
  });

  withEvidenceFile(report('passed'), (path) => {
    withOfficeToolE2eReports(true, (reports) => {
      const broken = JSON.parse(readFileSync(reports.word, 'utf8')) as Record<string, unknown>;
      broken.addin_activation = { activated: false, skipped: 'no-activator-configured' };
      writeFileSync(reports.word, JSON.stringify(broken, null, 2));
      const result = runValidator(path, '--require-office-tool-e2e', '--word-tool-e2e-report-path', reports.word, '--excel-tool-e2e-report-path', reports.excel, '--powerpoint-tool-e2e-report-path', reports.powerpoint);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Word Office tool E2E report add-in activation did not run/);
    });
  });

  withEvidenceFile(report('passed'), (path) => {
    withOfficeToolE2eReports(true, (reports) => {
      const broken = JSON.parse(readFileSync(reports.excel, 'utf8')) as ReturnType<typeof officeToolE2eReport>;
      broken.executed_tools = broken.executed_tools.slice(0, -1);
      writeFileSync(reports.excel, JSON.stringify(broken, null, 2));
      const result = runValidator(path, '--require-office-tool-e2e', '--word-tool-e2e-report-path', reports.word, '--excel-tool-e2e-report-path', reports.excel, '--powerpoint-tool-e2e-report-path', reports.powerpoint);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Excel Office tool E2E report executed tools do not match advertised tools/);
    });
  });

  withEvidenceFile(report('passed'), (path) => {
    withOfficeToolE2eReports(true, (reports) => {
      const broken = JSON.parse(readFileSync(reports.powerpoint, 'utf8')) as ReturnType<typeof officeToolE2eReport>;
      broken.tool_runs[1].passed = false;
      writeFileSync(reports.powerpoint, JSON.stringify(broken, null, 2));
      const result = runValidator(path, '--require-office-tool-e2e', '--word-tool-e2e-report-path', reports.word, '--excel-tool-e2e-report-path', reports.excel, '--powerpoint-tool-e2e-report-path', reports.powerpoint);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /PowerPoint Office tool E2E report tool powerpoint\.add_slide did not pass/);
    });
  });
});

test('README documents Office tool E2E report validation command', () => {
  const readme = readFileSync(resolve(process.cwd(), '../../../..', 'README.md'), 'utf8');
  const commandLine = readme.split('\n').find((line) => line.includes('--require-office-tool-e2e') && line.includes('--word-tool-e2e-report-path')) ?? '';

  assert.match(readme, /OFFICE_MCP_RUN_E2E = '1'/);
  assert.match(readme, /OFFICE_MCP_E2E_ACTIVATOR/);
  assert.match(readme, /activate-office-mcp-addin\.ps1/);
  assert.match(readme, /OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR=0/);
  assert.match(readme, /no-activator-configured/);
  assert.match(readme, /must not restart Office, recreate the document, or reconnect/);
  assert.match(readme, /one table-driven\s+loop across the host's complete tool catalog/);
  assert.match(readme, /office-tool-e2e-<host>\.json/);
  assert.match(commandLine, /npm run evidence:validate/);
  assert.match(commandLine, /--require-office-tool-e2e/);
  assert.match(commandLine, /--word-tool-e2e-report-path .*office-tool-e2e-word\.json/);
  assert.match(commandLine, /--excel-tool-e2e-report-path .*office-tool-e2e-excel\.json/);
  assert.match(commandLine, /--powerpoint-tool-e2e-report-path .*office-tool-e2e-powerpoint\.json/);
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
      const passing = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      passing.screenshot_paths.logo_ribbon_size = passing.screenshot_paths.logo_tray_size;
      passing.screenshot_paths.logo_catalog_thumbnail = passing.screenshot_paths.logo_tray_size;
      passing.screenshot_paths.logo_daemon_titlebar = passing.screenshot_paths.logo_tray_size;
      passing.screenshot_paths.logo_installer_metadata = passing.screenshot_paths.logo_tray_size;
      passing.screenshot_metadata.logo_ribbon_size = passing.screenshot_metadata.logo_tray_size;
      passing.screenshot_metadata.logo_catalog_thumbnail = passing.screenshot_metadata.logo_tray_size;
      passing.screenshot_metadata.logo_daemon_titlebar = passing.screenshot_metadata.logo_tray_size;
      passing.screenshot_metadata.logo_installer_metadata = passing.screenshot_metadata.logo_tray_size;
      writeFileSync(visualPath, JSON.stringify(passing, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.equal(result.status, 0, outputText(result.stdout) + outputText(result.stderr));
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.screenshot_paths.excel_ribbon_command = broken.screenshot_paths.word_ribbon_command;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /reuses one live screenshot for distinct surfaces: word_ribbon_command and excel_ribbon_command/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.screenshot_paths.tray_native_menu = broken.screenshot_paths.tray_icon;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /reuses one live screenshot for distinct surfaces: tray_icon and tray_native_menu/);
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
      broken.screenshots_fresh_ready = false;
      broken.screenshots_fresh.word_ribbon_command = false;
      broken.screenshot_metadata.word_ribbon_command.fresh = false;
      broken.screenshot_metadata.word_ribbon_command.ready = false;
      broken.screenshot_metadata.word_ribbon_command.mtime = new Date(Date.now() - 60 * 60 * 1000).toISOString();
      broken.screenshot_metadata.word_ribbon_command.age_ms = 60 * 60 * 1000;
      broken.screenshot_metadata.word_ribbon_command.freshness_window_ms = 1000;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /screenshots are not fresh for the recorded run/);
      assert.match(outputText(result.stdout), /screenshot is stale: word_ribbon_command/);
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
      broken.word_taskpane.document_state = 'unknown';
      broken.word_taskpane.document_state_ready = false;
      broken.word_taskpane.density_ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /concrete Word editable\/read-only\/protected state/);
      assert.match(outputText(result.stdout), /Word task pane density pass flag/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.word_runtime_evidence_ready = false;
      broken.product_identity_review.word_runtime_evidence_ready = false;
      broken.word_taskpane.runtime_evidence_ready = false;
      broken.word_taskpane.runtime_evidence.smoke_details.find_count = 0;
      (broken.word_taskpane.runtime_evidence.smoke_details as Record<string, unknown>).com_tracked_change_passed = false;
      broken.word_taskpane.density_ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Word runtime evidence ready flag/);
      assert.match(outputText(result.stdout), /Word runtime evidence missing mutation readback/);
      assert.match(outputText(result.stdout), /Word runtime evidence missing COM tracked-change proof/);
      assert.match(outputText(result.stdout), /Word task pane density pass flag/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.daemon_main_window.compact_status_details_reviewed = false;
      broken.daemon_main_window.three_column_layout_reviewed = false;
      broken.daemon_main_window.ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /daemon main window compact status\/details review/);
      assert.match(outputText(result.stdout), /daemon main window three-column layout review/);
      assert.match(outputText(result.stdout), /daemon main window ready flag/);
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
      broken.powerpoint_taskpane.document_state = 'unknown';
      broken.powerpoint_taskpane.document_state_ready = false;
      broken.powerpoint_taskpane.density_ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /concrete PowerPoint editable\/read-only\/protected state/);
      assert.match(outputText(result.stdout), /PowerPoint task pane density pass flag/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.powerpoint_runtime_evidence_ready = false;
      broken.product_identity_review.powerpoint_runtime_evidence_ready = false;
      broken.powerpoint_runtime_evidence.smoke_details.mutation_proved = false;
      broken.powerpoint_runtime_evidence.smoke_details.replace_text = { replacements: 0 };
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /PowerPoint runtime evidence ready flag/);
      assert.match(outputText(result.stdout), /PowerPoint runtime evidence did not prove mutation path/);
      assert.match(outputText(result.stdout), /PowerPoint runtime evidence missing replace_text proof/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.powerpoint_taskpane.runtime_evidence_ready = false;
      broken.powerpoint_taskpane.runtime_evidence.smoke_details.mutation_proved = false;
      broken.powerpoint_taskpane.density_ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /PowerPoint runtime evidence ready flag/);
      assert.match(outputText(result.stdout), /PowerPoint runtime evidence did not prove mutation path/);
      assert.match(outputText(result.stdout), /PowerPoint task pane density pass flag/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.catalog_type = 'Local productivity automation control utility protocol bridge debug panel';
      broken.first_run_identity.word.type = 'Local productivity automation control utility sample debug add-in';
      broken.tray_menu_surface_kind = 'webview';
      broken.tray_menu_surface_native = false;
      broken.manual_tray_evidence.tray_menu_surface_kind = 'html';
      broken.manual_tray_evidence.tray_menu_surface_native = false;
      broken.manual_tray_evidence.tray_surface_screenshots_distinct = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /mature local productivity automation\/control type metadata/);
      assert.match(outputText(result.stdout), /tray menu surface is not native/);
      assert.match(outputText(result.stdout), /Embedded manual tray evidence surface is not native/);
      assert.match(outputText(result.stdout), /Embedded manual tray evidence reuses one screenshot for multiple tray surfaces/);
      assert.match(outputText(result.stdout), /Word mature local productivity automation\/control type metadata/);
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
      broken.manual_tray_evidence.menu_opened_from_tray_icon = false;
      broken.manual_tray_evidence.native_menu_appearance_reviewed = false;
      broken.manual_tray_evidence.menu_anchored_to_tray_icon = false;
      broken.manual_tray_evidence.os_native_menu_behavior_reviewed = false;
      broken.manual_tray_evidence.keyboard_menu_access_reviewed = false;
      broken.manual_tray_evidence.native_quit_confirmation_reviewed = false;
      broken.manual_tray_evidence.native_tray_interaction_ready = false;
      broken.manual_tray_evidence_ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /embedded manual tray evidence ready flag/);
      assert.match(outputText(result.stdout), /right-click menu opened from the notification-area tray icon/);
      assert.match(outputText(result.stdout), /native tray menu appearance review/);
      assert.match(outputText(result.stdout), /native menu anchored to the notification-area tray icon/);
      assert.match(outputText(result.stdout), /OS-native tray menu spacing, hover, and theme behavior review/);
      assert.match(outputText(result.stdout), /keyboard access for native tray menu actions/);
      assert.match(outputText(result.stdout), /native quit confirmation review/);
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
      broken.office_tool_e2e_ready = false;
      broken.office_tool_e2e.word.ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Product visual evidence missing Office tool E2E ready flag/);
      assert.match(outputText(result.stdout), /Product visual Office tool E2E Word report is not ready/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.office_tool_e2e.excel.executed_tools = broken.office_tool_e2e.excel.executed_tools.slice(0, -1);
      broken.office_tool_e2e.excel.ready = false;
      broken.office_tool_e2e_ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Product visual Office tool E2E Excel report executed tools do not match advertised tools/);
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
      broken.product_identity_review.final_logo_user_surface_reviewed = false;
      broken.product_identity_review.addin_installable_surface_reviewed = false;
      broken.product_identity_review.tray_normal_windows_launch_reviewed = false;
      broken.product_identity_review.ready = false;
      broken.passed = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /final logo user-surface review/);
      assert.match(outputText(result.stdout), /add-in installable-software surface review/);
      assert.match(outputText(result.stdout), /tray normal Windows launch review/);
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
      broken.product_identity_review.logo_future_office_control_reviewed = false;
      broken.product_identity_review.addin_title_icon_type_reviewed = false;
      broken.product_identity_review.tray_native_first_impression_reviewed = false;
      broken.product_identity_review.ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /future office-control logo review/);
      assert.match(outputText(result.stdout), /add-in title, icon, and type\/category review/);
      assert.match(outputText(result.stdout), /tray native first-impression review/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.product_identity_review.current_logo_screenshot_feedback_reviewed = false;
      broken.product_identity_review.current_addin_screenshot_feedback_reviewed = false;
      broken.product_identity_review.current_tray_screenshot_feedback_reviewed = false;
      broken.product_identity_review.current_screenshot_feedback_ready = false;
      broken.product_identity_review.ready = false;
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /current screenshot logo feedback review/);
      assert.match(outputText(result.stdout), /current screenshot add-in feedback review/);
      assert.match(outputText(result.stdout), /current screenshot tray feedback review/);
      assert.match(outputText(result.stdout), /current screenshot feedback ready flag/);
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
      assert.match(outputText(result.stdout), /Word mature local productivity automation\/control type metadata/);
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
      broken.catalog_identity_review_ready = false;
      broken.catalog_identity_review.ready = false;
      broken.catalog_identity_review.hosts.word.display_name = 'office-mcp-word';
      broken.catalog_identity_review.hosts.excel.group_label = 'Office MCP';
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /catalog identity review ready flag/);
      assert.match(outputText(result.stdout), /Catalog identity review missing Word product display name/);
      assert.match(outputText(result.stdout), /Catalog identity review missing Excel product ribbon group label/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.rendered_logo_review.design_review.concept_pass = {
        ready: false,
        selected_direction: '',
        minimum_concepts_reviewed: 0,
        concepts: [],
        rejected_patterns: []
      };
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Rendered logo concept pass is not ready/);
      assert.match(outputText(result.stdout), /Rendered logo concept pass must review at least three concepts/);
    });
  });

  withEvidenceFile(ui, (uiPath) => {
    withProductVisualEvidence(true, (visualPath) => {
      const broken = JSON.parse(readFileSync(visualPath, 'utf8')) as ReturnType<typeof productVisualReport>;
      broken.rendered_logo_review.source_asset_sha256 = '0'.repeat(64);
      broken.rendered_logo_review.surfaces[0].asset_sha256 = '0'.repeat(64);
      writeFileSync(visualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-product-visual', '--product-visual-evidence-path', visualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /source asset fingerprint does not match current brand-mark\.svg/);
      assert.match(outputText(result.stdout), /asset fingerprint does not match current generated icon: logo_tray_size/);
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
      broken.tray_menu_surface_kind = 'webview';
      broken.tray_menu_surface_native = false;
      broken.tray_surface_screenshots_distinct = false;
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray evidence surface is not native/);
      assert.match(outputText(result.stdout), /Manual tray evidence reuses one screenshot for multiple tray surfaces/);
      assert.match(outputText(result.stdout), /native tray menu surface/);
    });
  });
  withEvidenceFile(ui, (uiPath) => {
    withManualTrayEvidence(true, (manualPath) => {
      const broken = JSON.parse(readFileSync(manualPath, 'utf8')) as ReturnType<typeof manualTrayReport>;
      broken.menu_opened_from_tray_icon = false;
      broken.native_menu_appearance_reviewed = false;
      broken.menu_anchored_to_tray_icon = false;
      broken.os_native_menu_behavior_reviewed = false;
      broken.keyboard_menu_access_reviewed = false;
      broken.native_quit_confirmation_reviewed = false;
      broken.native_tray_interaction_ready = false;
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /right-click menu opened from the notification-area tray icon/);
      assert.match(outputText(result.stdout), /native tray menu appearance review/);
      assert.match(outputText(result.stdout), /native menu anchored to the notification-area tray icon/);
      assert.match(outputText(result.stdout), /OS-native tray menu spacing, hover, and theme behavior review/);
      assert.match(outputText(result.stdout), /keyboard access for native tray menu actions/);
      assert.match(outputText(result.stdout), /native quit confirmation review/);
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
    if (passed) for (const path of Object.values(traySurfaceScreenshotPaths(screenshotPath))) writeFileSync(path, tinyPng());
    const path = join(dir, 'tray-manual-evidence.json');
    writeFileSync(path, JSON.stringify(manualTrayReport(passed, screenshotPath), null, 2));
    callback(path);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function withOfficeToolE2eReports(passed: boolean, callback: (reports: Record<'word' | 'excel' | 'powerpoint', string>) => void): void {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-tool-e2e-reports-'));
  try {
    const reports = {
      word: join(dir, 'office-tool-e2e-word.json'),
      excel: join(dir, 'office-tool-e2e-excel.json'),
      powerpoint: join(dir, 'office-tool-e2e-powerpoint.json')
    };
    writeFileSync(reports.word, JSON.stringify(officeToolE2eReport('Word', ['word.get_text', 'word.insert_paragraph'], passed), null, 2));
    writeFileSync(reports.excel, JSON.stringify(officeToolE2eReport('Excel', ['excel.get_workbook_info', 'excel.write_range'], passed), null, 2));
    writeFileSync(reports.powerpoint, JSON.stringify(officeToolE2eReport('PowerPoint', ['powerpoint.get_presentation_info', 'powerpoint.add_slide'], passed), null, 2));
    callback(reports);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function officeToolE2eReport(host: 'Word' | 'Excel' | 'PowerPoint', tools: string[], passed: boolean) {
  const now = new Date().toISOString();
  return {
    schema_version: 1,
    kind: 'office_tool_e2e_report',
    host,
    started_at: now,
    finished_at: now,
    passed,
    daemon: { endpoint: 'http://127.0.0.1:8765/mcp' },
    document: { path: `${host.toLowerCase()}-fixture` },
    addin_activation: { activated: true, activator: 'office-ui-activator' },
    session: { session_id: `${host.toLowerCase()}-session`, available_tool_count: tools.length },
    lifecycle_counts: {
      start_daemon: 1,
      list_tools: 1,
      create_document: 1,
      activate_addin: 1,
      wait_for_session: 1,
      cleanup_document: 1,
      stop_daemon: 1
    },
    advertised_tools: tools,
    session_available_tools: tools,
    executed_tools: tools,
    tool_runs: tools.map((tool, index) => ({
      id: `e2e-${tool.replace(/[^a-z0-9]+/gi, '-')}`,
      tool,
      started_at: now,
      finished_at: now,
      setup_action_count: 1,
      verifier: { kind: index === 0 ? 'direct-result' : 'readback', expectation_keys: ['contains'] },
      passed
    }))
  };
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
    if (passed) for (const path of Object.values(traySurfaceScreenshotPaths(screenshots.tray_icon))) writeFileSync(path, tinyPng());
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
    screenshot_metadata: screenshotMetadataFor(screenshots, passed),
    screenshots_fresh: Object.fromEntries(productVisualSurfaces().map((surface) => [surface, passed])),
    screenshots_fresh_ready: passed,
    product_text_ready: passed,
    catalog_type: 'Local productivity automation control utility',
    catalog_type_ready: passed,
    catalog_identity_review: catalogIdentityReview(passed),
    catalog_identity_review_ready: passed,
    catalog_icon_visible: passed,
    tray_tooltip: 'Office MCP Control - Up - 0 clients - 0 documents',
    tray_tooltip_ready: passed,
    tray_icon_visible: passed,
    tray_menu_native: passed,
    tray_menu_surface_kind: passed ? 'native' : 'webview',
    tray_menu_surface_native: passed,
    quit_confirmation_visible: passed,
    manual_tray_evidence: manualTrayReport(passed, screenshots.tray_icon),
    manual_tray_evidence_ready: passed,
    product_identity_review: {
      logo_quality_reviewed: passed,
      logo_future_office_control_reviewed: passed,
      final_logo_user_surface_reviewed: passed,
      current_logo_screenshot_feedback_reviewed: passed,
      rendered_size_logo_reviewed: passed,
      rendered_logo_review_ready: passed,
      addin_identity_reviewed: passed,
      addin_title_icon_type_reviewed: passed,
      addin_installable_surface_reviewed: passed,
      current_addin_screenshot_feedback_reviewed: passed,
      word_first_run_identity_reviewed: passed,
      excel_first_run_identity_reviewed: passed,
      powerpoint_first_run_identity_reviewed: passed,
      tray_product_polish_reviewed: passed,
      tray_native_first_impression_reviewed: passed,
      tray_normal_windows_launch_reviewed: passed,
      current_tray_screenshot_feedback_reviewed: passed,
      current_screenshot_feedback_ready: passed,
      word_first_run_identity_ready: passed,
      excel_first_run_identity_ready: passed,
      powerpoint_first_run_identity_ready: passed,
      word_runtime_evidence_ready: passed,
      powerpoint_runtime_evidence_ready: passed,
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
      },
      powerpoint: {
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
    word_runtime_evidence: wordRuntimeEvidence(passed),
    word_runtime_evidence_ready: passed,
    daemon_main_window: {
      reviewed: passed,
      compact_status_details_reviewed: passed,
      three_column_layout_reviewed: passed,
      ready: passed
    },
    powerpoint_runtime_evidence: powerPointRuntimeEvidence(passed),
    powerpoint_runtime_evidence_ready: passed,
    office_tool_e2e: {
      word: embeddedOfficeToolE2eReport('Word', ['word.get_text', 'word.insert_paragraph'], passed),
      excel: embeddedOfficeToolE2eReport('Excel', ['excel.get_workbook_info', 'excel.write_range'], passed),
      powerpoint: embeddedOfficeToolE2eReport('PowerPoint', ['powerpoint.get_presentation_info', 'powerpoint.add_slide'], passed)
    },
    office_tool_e2e_ready: passed,
    word_taskpane: {
      compact_top_block: passed,
      tools_permissions_merged: passed,
      inline_settings: passed,
      server_protocol_row: 'Server 0.1.0 / Protocol 1.0',
      server_protocol_row_ready: passed,
      document_state: 'Editable',
      document_state_ready: passed,
      runtime_evidence: wordRuntimeEvidence(passed),
      runtime_evidence_ready: passed,
      density_ready: passed
    },
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
    powerpoint_taskpane: {
      compact_top_block: passed,
      tools_permissions_merged: passed,
      inline_settings: passed,
      server_protocol_row: 'Server 0.1.0 / Protocol 1.0',
      server_protocol_row_ready: passed,
      document_state: 'Editable',
      document_state_ready: passed,
      runtime_evidence: powerPointRuntimeEvidence(passed),
      runtime_evidence_ready: passed,
      density_ready: passed
    },
    daemon_context: manualTrayDaemonContext(),
    daemon_context_ready: passed,
    passed
  };
}

function screenshotMetadataFor(screenshots: Record<string, string>, passed: boolean): Record<string, Record<string, unknown>> {
  return Object.fromEntries(productVisualSurfaces().map((surface) => {
    const path = screenshots[surface];
    if (!passed) return [surface, { path, ready: false, fresh: false }];
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

function catalogIdentityReview(passed: boolean) {
  return {
    ok: passed,
    schema_version: 1,
    kind: 'catalog_identity_review',
    product_name: 'Office MCP Control',
    catalog_path: 'C:\\catalog',
    catalog_type: passed ? 'Local productivity automation control utility' : 'Task Pane Add-in protocol bridge',
    shared_origin: passed ? 'https://localhost:8765' : null,
    hosts: {
      word: catalogIdentityHost('word', passed),
      excel: catalogIdentityHost('excel', passed),
      powerpoint: catalogIdentityHost('powerpoint', passed)
    },
    ready: passed,
    failures: passed ? [] : ['Catalog type is not product-ready.']
  };
}

function catalogIdentityHost(host: string, passed: boolean) {
  const taskpanePath = host === 'powerpoint' ? '/powerpoint/taskpane.html' : `/${host}/taskpane.html`;
  return {
    key: host,
    label: host === 'word' ? 'Word' : host === 'excel' ? 'Excel' : 'PowerPoint',
    display_name: passed ? 'Office MCP Control' : `office-mcp-${host}`,
    provider: passed ? 'Office MCP Control' : 'office-mcp',
    description: passed ? 'Control live documents through a local productivity automation control utility.' : 'Experimental protocol bridge debug panel.',
    group_label: passed ? 'Office MCP Control' : 'Office MCP',
    command_label: passed ? 'Open Control Panel' : 'Open',
    taskpane_url: `https://localhost:8765${taskpanePath}?v=0.1.0`,
    icon_url: passed ? 'https://localhost:8765/assets/icon-32.png' : 'https://localhost:8765/assets/blank.png',
    high_resolution_icon_url: passed ? 'https://localhost:8765/assets/icon-80.png' : 'https://localhost:8765/assets/blank.png',
    ready: passed,
    failures: passed ? [] : ['Prototype metadata.']
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
      available_tool_count: 20,
      document: { title: 'Excel Workbook' },
      host: { app: 'excel', platform: 'pc', version: '16.0' }
    },
    smoke_details: {
      session_id: sessionId,
      marker_found: passed,
      workbook_info: { sheet_count: 2, table_count: 1 },
      sheet_list_count: 1,
      updated_sheet: { updated: passed },
      deleted_sheet: { deleted: passed },
      used_range: { address: 'Sheet1!A1:C5' },
      find_replace: { replaced: passed, replaced_count: 1 },
      clear: { cleared: passed },
      write: { wrote_values: passed },
      formula: { wrote_formula: passed },
      format: { formatted: passed },
      table: { table: 'OfficeMcpTable' },
      table_update: { updated: passed },
      sort: { sorted: passed },
      filter: { filtered: passed },
      chart: { chart: 'Chart 1' },
      chart_update: { updated: passed },
      pivot_table: { pivot_table: 'OfficeMcpPivot' },
      pivot_update: { refreshed: passed },
      sheet: { activated: passed }
    }
  };
}

function wordRuntimeEvidence(passed: boolean) {
  const sessionId = '00000000-1111-2222-3333-444444444444';
  return {
    ok: passed,
    schema_version: 1,
    endpoint: 'http://127.0.0.1:8800/mcp',
    generated_at: new Date().toISOString(),
    session: {
      app: 'word',
      status: 'active',
      session_id: sessionId,
      available_tool_count: 25,
      document: { title: 'Word Document' },
      host: { app: 'word', platform: 'pc', version: '16.0' }
    },
    smoke_details: {
      session_id: sessionId,
      available_tool_count: 25,
      paragraph_0_text_length: passed ? 23 : 0,
      document_text_length: passed ? 23 : 0,
      find_count: passed ? 1 : 0,
      full_smoke_passed: passed,
      com_tracked_change_passed: passed
    },
    smoke_passed: passed,
    ready: passed
  };
}

function powerPointRuntimeEvidence(passed: boolean) {
  const sessionId = '22222222-3333-4444-5555-666666666666';
  return {
    ok: passed,
    schema_version: 1,
    endpoint: 'http://127.0.0.1:8800/mcp',
    generated_at: new Date().toISOString(),
    smoke_passed: passed,
    ready: passed,
    session: {
      app: 'powerpoint',
      status: 'active',
      session_id: sessionId,
      available_tool_count: 25,
      document: { title: 'PowerPoint Presentation' },
      host: { app: 'powerpoint', platform: 'pc', version: '16.0' }
    },
    smoke_details: {
      session_id: sessionId,
      available_tool_count: 25,
      available_tools: POWERPOINT_V1_TOOLS,
      presentation_info: { slide_count: 1 },
      active_view: { active_view: 'edit' },
      list_slides: { slides: [{ slide_id: 'slide-1', slide_index: 0 }] },
      add_slide: { slide_id: 'slide-1', slide_index: 0 },
      add_text_box: { shape: { shape_id: 'shape-1' } },
      list_shapes: { shapes: [{ shape_id: 'shape-1' }] },
      read_text: { items: [{ shape_id: 'shape-1', text: 'Office MCP' }] },
      replace_text: { replacements: passed ? 1 : 0 },
      format_text: { shape_id: 'shape-1', formatted: passed },
      list_layouts: { masters: [{ id: 'master-1', layouts: [{ id: 'layout-1' }] }] },
      layout: { slide_id: 'slide-1', slide_index: 0, layout_name: 'Title Only' },
      add_table: { shape_id: 'table-1' },
      read_table: { shape_id: 'table-1', values: [['Office']] },
      mutation_proved: passed,
      tool_category_proofs: {
        presentation: passed,
        slides: passed,
        layout: passed,
        shapes: passed,
        text: passed,
        tables: passed
      },
      export_supported: false,
      export_host_rejection: passed,
      table_supported: passed,
      table_host_rejection: false
    }
  };
}

function embeddedOfficeToolE2eReport(host: 'Word' | 'Excel' | 'PowerPoint', tools: string[], passed: boolean) {
  return {
    path: `C:\\Code\\office-mcp\\artifacts\\office-tool-e2e-${host.toLowerCase()}.json`,
    ok: passed,
    host,
    schema_version: 1,
    kind: 'office_tool_e2e_report',
    report_host: host,
    passed,
    addin_activation: { activated: passed, activator: 'office-ui-activator' },
    lifecycle_counts: {
      start_daemon: 1,
      list_tools: 1,
      create_document: 1,
      activate_addin: 1,
      wait_for_session: 1,
      cleanup_document: 1,
      stop_daemon: 1
    },
    advertised_tools: tools,
    session_available_tools: tools,
    executed_tools: tools,
    tool_runs: tools.map((tool, index) => ({
      id: `e2e-${tool.replace(/[^a-z0-9]+/gi, '-')}`,
      tool,
      setup_action_count: 1,
      verifier: { kind: index === 0 ? 'direct-result' : 'readback', expectation_keys: ['contains'] },
      passed
    })),
    ready: passed
  };
}

function renderedLogoReview(passed: boolean, sheetPath: string) {
  return {
    ok: passed,
    schema_version: 1,
    kind: 'rendered_logo_review',
    product_name: 'Office MCP Control',
    source_asset_path: resolve(ASSET_ROOT, 'brand-mark.svg'),
    source_asset_sha256: passed ? sha256File(resolve(ASSET_ROOT, 'brand-mark.svg')) : '0'.repeat(64),
    sheet_path: sheetPath,
    ready: passed,
    design_review: renderedLogoDesignReview(passed),
    surfaces: [
      ['logo_tray_size', 16, 'icon-16.png'],
      ['logo_ribbon_size', 32, 'icon-32.png'],
      ['logo_catalog_thumbnail', 80, 'icon-80.png'],
      ['logo_daemon_titlebar', 20, 'icon-20.png'],
      ['logo_installer_metadata', 256, 'icon-256.png']
    ].map(([key, size, asset]) => ({
      key,
      asset_path: resolve(ASSET_ROOT, String(asset)),
      asset_sha256: passed ? sha256File(resolve(ASSET_ROOT, String(asset))) : '0'.repeat(64),
      rendered_size_px: size,
      width: size,
      height: size,
      non_empty: passed,
      palette_ready: passed,
      expected_size_ready: passed
    }))
  };
}

function sha256File(path: string): string {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

function renderedLogoDesignReview(passed: boolean) {
  return {
    future_office_control_brief: passed ? 'Future office control: routing geometry and operator control without Office-owned app marks.' : '',
    concept_pass: renderedLogoConceptPass(passed),
    office_productivity_metaphor: passed ? 'Abstract document panes communicate office productivity.' : '',
    user_control_metaphor: passed ? 'Command routing and operator nodes communicate local user control.' : '',
    futuristic_maturity: passed ? 'Mature slightly futuristic desktop utility geometry.' : '',
    non_microsoft_distinction: passed ? 'Avoids Office logos, Microsoft 365 gradients, Word silhouettes, Excel grid marks, PowerPoint slide silhouettes, Outlook envelope marks, and gear-only artwork.' : '',
    rejects_generic_readings: passed ? ['settings', 'file', 'debug console', 'ai-only', 'microsoft office clone'] : [],
    ready: passed
  };
}

function renderedLogoConceptPass(passed: boolean) {
  return {
    ready: passed,
    selected_direction: passed ? 'Command Console Panes' : '',
    minimum_concepts_reviewed: passed ? 3 : 0,
    concepts: passed ? [
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
    rejected_patterns: passed ? ['gear-only settings mark', 'Office-like app tile', 'host-app color block', 'generic document thumbnail', 'terminal/debug glyph', 'AI sparkle motif'] : []
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
    'powerpoint_ribbon_command',
    'powerpoint_catalog_entry',
    'powerpoint_taskpane_title',
    'daemon_main_window',
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
    ok: passed,
    schema_version: 1,
    kind: 'tray_manual_evidence',
    recorded_at: new Date().toISOString(),
    tester: 'test',
    platform: 'win32',
    visible_icon: passed,
    right_click_menu: passed,
    menu_opened_from_tray_icon: passed,
    native_menu_appearance_reviewed: passed,
    menu_anchored_to_tray_icon: passed,
    os_native_menu_behavior_reviewed: passed,
    keyboard_menu_access_reviewed: passed,
    native_quit_confirmation_reviewed: passed,
    native_tray_interaction_ready: passed,
    tray_menu_surface_kind: passed ? 'native' : 'webview',
    tray_menu_surface_native: passed,
    show_ui_opened: passed,
    observed_menu_items: ['Status: Up', 'Clients: 0', 'Documents: 0', 'Show Office MCP Control', 'Quit Office MCP Control'],
    observed_tooltip: 'Office MCP Control - Up - 0 clients - 0 documents',
    expected_menu_items: ['Status:', 'Clients:', 'Documents:', 'Show Office MCP Control', 'Quit Office MCP Control'],
    menu_contains_required_items: passed,
    screenshot_path: screenshotPath,
    tray_surface_screenshot_paths: traySurfaceScreenshotPaths(screenshotPath),
    tray_surface_screenshots_exist: Object.fromEntries(trayVisualSurfaces().map((surface) => [surface, passed])),
    tray_surface_screenshots_ready: passed,
    tray_surface_screenshots_distinct: passed,
    screenshot_exists: passed,
    daemon_context: manualTrayDaemonContext() as ReturnType<typeof manualTrayDaemonContext> | undefined,
    daemon_context_ready: passed,
    passed
  };
}

function traySurfaceScreenshotPaths(basePath: string): Record<string, string> {
  const extensionIndex = basePath.toLowerCase().lastIndexOf('.png');
  const prefix = extensionIndex === -1 ? basePath : basePath.slice(0, extensionIndex);
  return Object.fromEntries(trayVisualSurfaces().map((surface) => [surface, `${prefix}-${surface}.png`]));
}

function trayVisualSurfaces(): string[] {
  return [
    'tray_icon',
    'tray_native_menu',
    'tray_tooltip',
    'tray_quit_confirmation'
  ];
}

function manualTrayDaemonContext(menuItems = ['Status: Up', 'Clients: 0', 'Documents: 0', '---', 'Show Office MCP Control', 'Quit Office MCP Control'], stateFetchOk = true) {
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
        tooltip: 'Office MCP Control - Up - 0 clients - 0 documents',
        platform: 'windows-notification-area'
      }
    }
  };
}

function powerpointOnlyReport() {
  const now = new Date().toISOString();
  return {
    schema_version: 1,
    generated_at: now,
    endpoint: 'http://127.0.0.1:8800/mcp',
    gates: [
      gate('word.session_discovery', 'passed', { sessions: [{ app: 'powerpoint', session_id: 'ppt-session' }] }),
      gate('powerpoint.runtime_smoke', 'passed', {
        session_id: 'ppt-session',
        available_tool_count: 25,
        available_tools: POWERPOINT_V1_TOOLS,
        presentation_info: { slide_count: 1 },
        active_view: { active_view: 'edit' },
        list_slides: { slides: [{ slide_id: 'slide-1', slide_index: 0 }] },
        add_slide: { slide_id: 'slide-1', slide_index: 0, added: true },
        add_text_box: { shape: { shape_id: 'shape-1' } },
        list_shapes: { shapes: [{ shape_id: 'shape-1' }] },
        read_text: { items: [{ shape_id: 'shape-1', text: 'Office MCP' }] },
        replace_text: { replacements: 1, touched_shapes: [{ slide_id: 'slide-1', shape_id: 'shape-1' }] },
        format_text: { shape_id: 'shape-1', formatted: true },
        list_layouts: { masters: [{ id: 'master-1', layouts: [{ id: 'layout-1' }] }] },
        layout: { slide_id: 'slide-1', slide_index: 0, layout_name: 'Title Only' },
        add_table: { shape_id: 'table-1' },
        read_table: { shape_id: 'table-1', values: [['Office']] },
        mutation_proved: true,
        tool_category_proofs: {
          presentation: true,
          slides: true,
          layout: true,
          shapes: true,
          text: true,
          tables: true
        },
        export_supported: false,
        export_host_rejection: true,
        table_supported: true,
        table_host_rejection: false
      })
    ]
  };
}
function excelOnlyReport() {
  const now = new Date().toISOString();
  const sessionId = 'excel-session-1';
  return {
    schema_version: 1,
    generated_at: now,
    endpoint: 'http://127.0.0.1:8800/mcp',
    gates: [
      gate('word.session_discovery', 'passed', {
        sessions: [{
          app: 'excel',
          status: 'active',
          session_id: sessionId,
          available_tool_count: 20,
          document: { title: 'E2E workbook.xlsx' },
          host: { app: 'excel', platform: 'pc', version: '16.0' }
        }]
      }),
      gate('excel.runtime_smoke', 'passed', excelSmokeDetails(sessionId))
    ]
  };
}

function excelSmokeDetails(sessionId: string) {
  return {
    session_id: sessionId,
    document_title: 'E2E workbook.xlsx',
    available_tool_count: 20,
    available_tools: EXCEL_V1_TOOLS,
    marker_found: true,
    workbook_info: { active_sheet: 'Sheet1', sheet_count: 1, table_count: 0 },
    sheet_list_count: 1,
    read_before_address: 'Sheet1!A1:B2',
    updated_sheet: { name: 'Renamed Sheet' },
    deleted_sheet: { deleted: true },
    used_range: { address: 'Renamed Sheet!A1:C5' },
    find_replace: { replaced: 1 },
    clear: { cleared: true },
    write: { wrote_values: true },
    formula: { wrote_formula: true },
    format: { formatted: true },
    table: { table: 'E2ETable' },
    table_update: { updated: true },
    sort: { sorted: true },
    filter: { filtered: true },
    chart: { chart: 'E2E Chart' },
    chart_update: { updated: true },
    pivot_table: { pivot_table: 'E2EPivot' },
    pivot_update: { refreshed: true },
    sheet: { activated: true }
  };
}

function structuredMenu(menuItems: string[]) {
  return menuItems.map((label, index) => {
    if (label === '---') return { kind: 'separator', label, enabled: false };
    if (label === 'Show Office MCP Control') return { kind: 'action', label, action: 'show_ui', enabled: true };
    if (label === 'Quit Office MCP Control') return { kind: 'action', label, action: 'quit', enabled: true };
    return { kind: 'read_only', label, enabled: false };
  });
}
