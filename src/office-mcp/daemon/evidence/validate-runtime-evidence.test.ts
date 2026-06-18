import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';

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
      broken.daemon_context = manualTrayDaemonContext(['Status: Up', 'Clients: 0']);
      writeFileSync(manualPath, JSON.stringify(broken, null, 2));
      const result = runValidator(uiPath, '--ui', '--require-manual-tray', '--manual-tray-evidence-path', manualPath);
      assert.notEqual(result.status, 0);
      assert.match(outputText(result.stdout), /Manual tray daemon context missing live menu item: Documents:/);
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
    if (passed) writeFileSync(screenshotPath, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]));
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

function manualTrayReport(passed: boolean, screenshotPath = 'C:\\temp\\tray.png') {
  return {
    schema_version: 1,
    kind: 'tray_manual_evidence',
    recorded_at: new Date().toISOString(),
    tester: 'test',
    platform: 'win32',
    visible_icon: passed,
    right_click_menu: passed,
    show_ui_opened: passed,
    observed_menu_items: ['Status: Up', 'Clients: 0', 'Documents: 0', 'Show Office MCP', 'Quit Office MCP'],
    expected_menu_items: ['Status:', 'Clients:', 'Documents:', 'Show Office MCP', 'Quit Office MCP'],
    menu_contains_required_items: passed,
    screenshot_path: screenshotPath,
    screenshot_exists: passed,
    daemon_context: undefined as ReturnType<typeof manualTrayDaemonContext> | undefined,
    passed
  };
}

function manualTrayDaemonContext(menuItems = ['Status: Up', 'Clients: 0', 'Documents: 0', 'Show Office MCP', 'Quit Office MCP']) {
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
      snapshot: {
        menu_items: menuItems,
        platform: 'windows-notification-area'
      },
      state_fetch_ok: true
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
