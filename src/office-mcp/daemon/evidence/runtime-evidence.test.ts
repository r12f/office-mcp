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
  assert.match(source, /agent_client_stdio_bridge/);
  assert.match(source, /irm_rights_matrix/);
  assert.match(source, /--irm-mode/);
  assert.match(source, /--irm-document-path/);
  assert.match(source, /--wait-for-session-ms/);
  assert.match(source, /process\.argv\.lastIndexOf\(name\)/);
  assert.match(source, /--agent-client-evidence-path/);
  assert.match(source, /word\.wait_for_requested_session/);
  assert.match(source, /claude_desktop_installation/);
  assert.match(source, /agent_client_prompt/);
  assert.match(source, /irm_document_preflight/);
  assert.doesNotMatch(source, /runtime_smoke|full_smoke|e2e_session/);
  assert.doesNotMatch(source, /includeExcelSmoke|includePowerPointSmoke|runExcelSmokeGate|runPowerPointSmokeGate|runWordMutationGate/);
  assert.doesNotMatch(source, /office-e2e-driver\.mjs|runOfficeE2eDriverStep/);
  assert.match(source, /function selectWordSessionId/);
  assert.match(source, /host\?\.app \?\? ''\)\.toLowerCase\(\) === 'word'/);
  assert.match(source, /open Office MCP Control, then rerun this script/);
  assert.doesNotMatch(source, /load the office-mcp task pane|load the add-in task pane/);
  assert.doesNotMatch(packageJson, /include-.*smoke|e2e-session|require-.*smoke/);
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

test('Excel capability spec points tool coverage at Office tool E2E', () => {
  const spec = readFileSync(resolve(process.cwd(), '../../../../doc/spec/04-excel-capabilities.md'), 'utf8');
  const e2e = readFileSync(resolve(process.cwd(), '../../../../src/office-ctl/excel/test/tool-e2e.test.mjs'), 'utf8');

  for (const tool of [
    'excel.get_workbook_info',
    'excel.list_sheets',
    'excel.read_range',
    'excel.write_range',
    'excel.add_sheet',
    'excel.update_sheet',
    'excel.delete_sheet',
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
    assert.match(e2e, new RegExp(tool.replace('.', '\\.')));
  }
  assert.match(spec, /ExcelApi 1\.1/);
  assert.match(spec, /npm run e2e:tools/);
  assert.doesNotMatch(spec, /excel\.runtime_smoke|require-excel-smoke/);
});

test('Office tool E2E owns live Office lifecycle and per-tool coverage', () => {
  const packageJson = readFileSync(evidencePath('package.json'), 'utf8');
  const contract = readFileSync(resolve(process.cwd(), '../../../../src/office-ctl/common/test/tool-e2e-contract.test.mjs'), 'utf8');

  assert.match(packageJson, /"evidence:excel"[^\n]+office-ctl\/excel[^\n]+e2e:tools/);
  assert.match(packageJson, /"evidence:powerpoint"[^\n]+office-ctl\/powerpoint[^\n]+e2e:tools/);
  assert.match(contract, /shared Office tool E2E loop drives daemon, document, setup, calls, verification, and cleanup/);
  assert.match(contract, /office_tool_e2e_report/);
});

test('UI browser smoke follows the current inline task pane settings contract', () => {
  const source = readFileSync(evidencePath('ui-browser-smoke.ts'), 'utf8');

  assert.match(source, /\.daemon-endpoint-form/);
  assert.match(source, /#saveEndpoint/);
  assert.match(source, /#endpointInput/);
  assert.match(source, /#connectionDetail/);
  assert.match(source, /taskpane endpoint validation uses last error row/);
  assert.match(source, /taskpane keeps endpoint validation on the last error row/);
  assert.doesNotMatch(source, /#endpointError"\)\.getAttribute/);
  assert.doesNotMatch(source, /requestSubmit\(\)/);
  assert.doesNotMatch(source, /#settingsToggle/);
  assert.doesNotMatch(source, /#settingsPanel/);
  assert.doesNotMatch(source, /is-editing-tools/);
  assert.doesNotMatch(source, /textContent\.includes\("Enabled"\)/);
});

test('daemon UI spec keeps diagnostic detail values selectable and copyable', () => {
  const uiSpec = readFileSync(resolve(process.cwd(), '../../../../doc/spec/09-ui.md'), 'utf8');
  const roadmap = readFileSync(resolve(process.cwd(), '../../../../doc/spec/08-roadmap.md'), 'utf8');

  assert.match(uiSpec, /top daemon details rail[\s\S]*selectable text/i);
  assert.match(uiSpec, /Config[\s\S]*Log[\s\S]*Last Error[\s\S]*copy/i);
  assert.match(uiSpec, /Log[\s\S]*wrap/i);
  assert.match(uiSpec, /Last Error[\s\S]*must not be ellipsis-only/i);
  assert.match(roadmap, /Config\/Log paths and `Last error` values are selectable/i);
  assert.doesNotMatch(uiSpec, /`Log` can use a medium flexible\s+column with truncation/i);
  assert.doesNotMatch(roadmap, /keep `Log` truncatable/i);
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
