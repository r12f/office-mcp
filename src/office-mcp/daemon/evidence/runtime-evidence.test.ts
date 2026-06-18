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

  for (const tool of [
    'excel.read_range',
    'excel.write_range',
    'excel.add_sheet',
    'excel.set_formula',
    'excel.format_range',
    'excel.create_table',
    'excel.create_chart'
  ]) {
    assert.match(spec, new RegExp(tool.replace('.', '\\.')));
  }
  assert.match(spec, /ExcelApi 1\.1/);
  assert.match(spec, /npm run evidence:excel/);
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
