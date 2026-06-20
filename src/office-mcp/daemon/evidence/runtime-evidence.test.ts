import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';

test('runtime evidence harness covers Word, agent client bridge, and IRM gates', () => {
  const source = readFileSync(evidencePath('runtime-evidence.ts'), 'utf8');
  const packageJson = readFileSync(evidencePath('package.json'), 'utf8');

  assert.match(source, /OFFICE_MCP_MCP_ENDPOINT/);
  assert.match(source, /word\.runtime_read_smoke/);
  assert.match(source, /agent_client_stdio_bridge/);
  assert.match(source, /irm_rights_matrix/);
  assert.match(source, /--include-mutation/);
  assert.match(source, /--include-full-word-smoke/);
  assert.match(source, /--include-com-tracked-changes/);
  assert.match(source, /--irm-mode/);
  assert.match(source, /--irm-document-path/);
  assert.match(source, /--wait-for-session-ms/);
  assert.match(source, /process\.argv\.lastIndexOf\(name\)/);
  assert.match(source, /--agent-client-evidence-path/);
  assert.match(source, /word\.wait_for_requested_session/);
  assert.match(source, /claude_desktop_installation/);
  assert.match(source, /agent_client_prompt/);
  assert.match(source, /irm_document_preflight/);
  assert.match(source, /const wantsWordBaseline = !\(includeExcelSmoke \|\| includePowerPointSmoke\) \|\| wantsWordRuntime/);
  assert.match(source, /if \(!sessionId && wantsWordBaseline\)/);
  assert.match(source, /else if \(sessionId\)/);
  assert.match(source, /wantsWordBaseline \? requestedSessionId \?\? selectWordSessionId/);
  assert.match(source, /function selectWordSessionId/);
  assert.match(source, /host\?\.app \?\? ''\)\.toLowerCase\(\) === 'word'/);
  assert.match(source, /open Office MCP Control, then rerun this script/);
  assert.doesNotMatch(source, /load the office-mcp task pane|load the add-in task pane/);
  assert.doesNotMatch(packageJson, /evidence:(irm|excel)[^\n]+--endpoint http:\/\/127\.0\.0\.1:8800\/mcp/);
});

test('runtime evidence harness writes structured failure evidence when daemon is unavailable', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-runtime-evidence-'));
  try {
    const output = join(dir, 'evidence.json');
    const result = spawnSync(process.execPath, [
      './node_modules/tsx/dist/cli.mjs',
      evidencePath('runtime-evidence.ts'),
      '--endpoint',
      'http://127.0.0.1:1/mcp',
      '--output',
      output
    ], { cwd: process.cwd(), encoding: 'utf8' });

    assert.equal(result.status, 0, result.stderr);
    const report = JSON.parse(readFileSync(output, 'utf8')) as { schema_version: number; gates: Array<{ name: string; status: string; details: { error?: string } }> };
    assert.equal(report.schema_version, 1);
    assert.equal(report.gates.length, 1);
    assert.equal(report.gates[0].name, 'runtime_evidence_harness');
    assert.equal(report.gates[0].status, 'failed');
    assert.match(String(report.gates[0].details.error), /ECONNREFUSED|fetch failed|connect/i);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('Excel capability spec exists and tracks the v1 tool catalog', () => {
  const spec = readFileSync(resolve(process.cwd(), '../../../../doc/spec/04-excel-capabilities.md'), 'utf8');
  const source = readFileSync(evidencePath('runtime-evidence.ts'), 'utf8');

  for (const tool of [
    'excel.get_workbook_info',
    'excel.list_sheets',
    'excel.read_range',
    'excel.write_range',
    'excel.add_sheet',
    'excel.update_sheet',
    'excel.delete_sheet',
    'excel.get_used_range',
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
  ]) {
    assert.match(spec, new RegExp(tool.replace('.', '\\.')));
    assert.match(source, new RegExp(tool.replace('.', '\\.')));
  }
  assert.match(spec, /ExcelApi 1\.1/);
  assert.match(spec, /npm run evidence:excel/);
  assert.match(source, /workbook_info/);
  assert.match(source, /sheet_list_count/);
  assert.match(source, /used_range/);
  assert.match(source, /find_replace/);
  assert.match(source, /sort/);
  assert.match(source, /filter/);
  assert.match(source, /table_update/);
  assert.match(source, /chart_update/);
  assert.match(source, /pivot_table/);
});

test('Excel runtime evidence can own the live Office E2E lifecycle', () => {
  const source = readFileSync(evidencePath('runtime-evidence.ts'), 'utf8');
  const packageJson = readFileSync(evidencePath('package.json'), 'utf8');

  assert.match(source, /--excel-e2e-session/);
  assert.match(source, /OFFICE_MCP_E2E_DRIVER/);
  assert.match(source, /office-e2e-driver\.mjs/);
  assert.match(source, /runExcelE2eSessionSmoke/);
  assert.match(source, /runOfficeE2eDriverStep\('Excel', 'startDaemon'/);
  assert.match(source, /runOfficeE2eDriverStep\('Excel', 'createDocument'/);
  assert.match(source, /runOfficeE2eDriverStep\('Excel', 'activateAddin'/);
  assert.match(source, /runOfficeE2eDriverStep\('Excel', 'waitForSession'/);
  assert.match(source, /runOfficeE2eDriverStep\('Excel', 'cleanupDocument'/);
  assert.match(source, /runOfficeE2eDriverStep\('Excel', 'stopDaemon'/);
  assert.match(source, /finally \{/);
  assert.match(source, /await runExcelSmokeGate\(sessionId\)/);
  assert.match(source, /excel\.e2e_session/);
  assert.match(source, /available_tools: availableTools/);
  assert.match(packageJson, /"evidence:excel"[^\n]+--excel-e2e-session/);
});

test('Excel runtime smoke uses worksheet names within Excel limits', () => {
  const source = readFileSync(evidencePath('runtime-evidence.ts'), 'utf8');

  assert.match(source, /const sheetName = `Mcp\$\{runId\}`/);
  assert.match(source, /const renamedSheetName = `McpR\$\{runId\}`/);
  assert.match(source, /const cleanupSheetName = `McpC\$\{runId\}`/);
  assert.doesNotMatch(source, /OfficeMcpSmoke\$\{Date\.now\(\)\}Renamed/);
});

test('PowerPoint runtime evidence harness proves session, tools, mutation, and host rejections', () => {
  const spec = readFileSync(resolve(process.cwd(), '../../../../doc/spec/08-roadmap.md'), 'utf8');
  const source = readFileSync(evidencePath('runtime-evidence.ts'), 'utf8');

  assert.match(spec, /Add live PowerPoint runtime smoke evidence against a real presentation/);
  assert.match(source, /--include-powerpoint-smoke/);
  assert.match(source, /runWaitForHostSessionGate\('powerpoint', waitForSessionMs\)/);
  assert.match(source, /runGate\(`\$\{hostApp\}\.wait_for_session`/);
  assert.match(source, /powerpoint\.runtime_smoke/);
  assert.match(source, /selectHostSessionId\(sessions, 'powerpoint'\)/);
  assert.match(source, /No connected PowerPoint add-in session/);

  for (const tool of [
    'office.get_session_info',
    'powerpoint.get_presentation_info',
    'powerpoint.get_active_view',
    'powerpoint.list_slides',
    'powerpoint.list_layouts',
    'powerpoint.add_slide',
    'powerpoint.add_text_box',
    'powerpoint.list_shapes',
    'powerpoint.read_text',
    'powerpoint.replace_text',
    'powerpoint.format_text',
    'powerpoint.apply_layout',
    'powerpoint.add_table',
    'powerpoint.read_table',
    'powerpoint.export_file'
  ]) {
    assert.match(source, new RegExp(tool.replace('.', '\\.')));
  }

  assert.match(source, /availableTools\.length/);
  assert.match(source, /available_tools: availableTools/);
  assert.match(source, /availableTools\.includes\('powerpoint\.export_file'\)/);
  assert.match(source, /availableTools\.includes\('powerpoint\.export_pdf'\)/);
  assert.match(source, /mutation_proved: typeof addSlide\.slide_id === 'string' && Number\(replaceText\.replacements \?\? 0\) >= 1/);
  assert.match(source, /exportSupported/);
  assert.match(source, /exportHostRejection/);
  assert.match(source, /tableSupported/);
  assert.match(source, /tableHostRejection/);
  assert.match(source, /PowerPoint file export failed without explicit host-capability rejection/);
  assert.match(source, /PowerPoint table creation failed without explicit host-capability rejection/);

  for (const category of ['presentation', 'slides', 'layout', 'shapes', 'text', 'tables']) {
    assert.match(source, new RegExp(`${category}:`));
  }
});

test('add-in communication specs stay metadata-only at the local boundary', () => {
  const specPaths = [
    '../../../../doc/spec/02-registration-protocol.md',
    '../../../../doc/spec/05-security.md',
    '../../../../doc/spec/09-ui.md'
  ];
  const specs = specPaths
    .map((path) => readFileSync(resolve(process.cwd(), path), 'utf8'))
    .join('\n');

  for (const forbidden of forbiddenLocalAdmissionTerms()) {
    assert.doesNotMatch(specs, new RegExp(forbidden, 'i'));
  }

  assert.match(specs, /loopback binding/);
  assert.match(specs, /exact `Origin` validation/);
  assert.match(specs, /MUST remain metadata-only/);
});

function evidencePath(file: string): string {
  return resolve(process.cwd(), file);
}

function forbiddenLocalAdmissionTerms(): string[] {
  const join = (...parts: string[]) => parts.join('');
  const spacer = '[-_ ]?';

  return [
    join('api', spacer, 'key'),
    join('shared', spacer, 'secret'),
    'pairing',
    join('credential', spacer, 'exchange'),
    join('per-launch ui ', 'token'),
    join('application-level ', 'credential'),
    join('local ', 'credential')
  ];
}
