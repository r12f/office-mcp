import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

test('runtime evidence harness covers Word, agent client bridge, and IRM gates', () => {
  const source = readFileSync(join(process.cwd(), 'test', 'runtime-evidence.ts'), 'utf8');

  assert.match(source, /word\.runtime_read_smoke/);
  assert.match(source, /agent_client_stdio_bridge/);
  assert.match(source, /irm_rights_matrix/);
  assert.match(source, /--include-mutation/);
  assert.match(source, /--include-full-word-smoke/);
  assert.match(source, /--include-com-tracked-changes/);
  assert.match(source, /--irm-mode/);
  assert.match(source, /--irm-document-path/);
  assert.match(source, /--wait-for-session-ms/);
  assert.match(source, /--agent-client-evidence-path/);
  assert.match(source, /word\.wait_for_requested_session/);
  assert.match(source, /claude_desktop_installation/);
  assert.match(source, /agent_client_prompt/);
  assert.match(source, /irm_document_preflight/);
});

test('runtime evidence harness writes structured failure evidence when daemon is unavailable', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-runtime-evidence-'));
  try {
    const output = join(dir, 'evidence.json');
    const result = spawnSync(process.execPath, [
      './node_modules/tsx/dist/cli.mjs',
      'test/runtime-evidence.ts',
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



