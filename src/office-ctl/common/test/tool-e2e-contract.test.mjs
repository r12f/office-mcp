import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import {
  assertConcreteE2eCases,
  assertE2eCaseCoverage,
  daemonCatalogTools,
  directResult,
  e2eCase,
  excelReadback,
  powerpointReadback,
  readbackByResource,
  readbackByTool,
  requireOfficeE2eDriver,
  runOfficeToolE2e,
  wordReadback
} from './tool-e2e-contract.mjs';

test('concrete E2E case gate accepts setup actions, calls, and verifiers', () => {
  assertConcreteE2eCases({
    host: 'Word',
    cases: {
      'word.replace_text': e2eCase('word.replace_text', {
        setup: {
          actions: [
            { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'baseline' } }
          ]
        },
        args: { find: 'baseline', replace: 'updated' },
        verify: readbackByTool('word.get_text', {
          arguments: { limit: 20 },
          expect: { contains: ['updated'], notContains: ['baseline'] }
        })
      }),
      'word.add_comment': e2eCase('word.add_comment', {
        setup: {
          actions: [
            { tool: 'word.insert_paragraph', arguments: { anchor: { kind: 'end_of_document' }, text: 'comment target' } }
          ]
        },
        args: { anchor: { kind: 'after_text', text: 'comment target' }, text: 'E2E comment' },
        verify: readbackByResource('office://word/${session_id}/comments', {
          expect: { contains: ['E2E comment'] }
        })
      })
    }
  });
});

test('concrete E2E case gate rejects weak placeholder cases', () => {
  assert.throws(
    () => assertConcreteE2eCases({
      host: 'Word',
      cases: { 'word.read': e2eCase('word.read') }
    }),
    /word\.read case must define setup\.actions/
  );
  assert.throws(
    () => assertConcreteE2eCases({
      host: 'Word',
      cases: {
        'word.read': e2eCase('word.read', {
          setup: { actions: [{ tool: 'word.insert_paragraph', arguments: {} }] },
          verify: { kind: 'direct-result', expect: {} }
        })
      }
    }),
    /word\.read verifier must define at least one expectation/
  );
  assert.throws(
    () => assertConcreteE2eCases({
      host: 'Word',
      cases: {
        'word.write': e2eCase('word.write', {
          setup: { actions: [{ tool: 'word.insert_paragraph', arguments: {} }] },
          verify: { kind: 'readback', expect: { contains: ['updated'] } }
        })
      }
    }),
    /word\.write readback verifier must define readbackTool or resource/
  );
});

test('shared Office tool E2E loop drives daemon, document, setup, calls, verification, and cleanup', async () => {
  const events = [];
  const driver = {
    async startDaemon() {
      events.push('startDaemon');
      return { endpoint: 'http://127.0.0.1:0/mcp' };
    },
    async listTools() {
      events.push('listTools');
      return ['office.list_sessions', 'word.read', 'word.write', 'excel.read'];
    },
    async createDocument() {
      events.push('createDocument');
      return { path: 'fixture.docx' };
    },
    async activateAddin(document) {
      events.push(`activateAddin:${document.path}`);
      return { activated: true };
    },
    async waitForSession(document) {
      events.push(`waitForSession:${document.path}`);
      return { sessionId: 'session-1', availableTools: ['word.read', 'word.write'] };
    },
    async resetContent(toolCase, session) {
      events.push(`reset:${toolCase.tool}:${session.sessionId}`);
    },
    async setupContent(toolCase) {
      events.push(`setup:${toolCase.tool}:${toolCase.setup}`);
    },
    async callTool(toolCase, session) {
      events.push(`call:${toolCase.call.name}:${session.sessionId}`);
      return { ok: true, tool: toolCase.tool };
    },
    async verifyResult(toolCase, result) {
      events.push(`verify:${toolCase.tool}:${toolCase.verify.kind}:${result.ok}`);
    },
    async cleanupDocument(document) {
      events.push(`cleanupDocument:${document.path}`);
    },
    async stopDaemon() {
      events.push('stopDaemon');
    }
  };
  const cases = {
    'word.read': e2eCase('word.read', { setup: 'read baseline', verify: 'direct-result' }),
    'word.write': e2eCase('word.write', { setup: 'write baseline', args: { text: 'updated' } })
  };

  await runOfficeToolE2e({ host: 'Word', cases, driver });

  assert.deepEqual(events, [
    'startDaemon',
    'listTools',
    'createDocument',
    'activateAddin:fixture.docx',
    'waitForSession:fixture.docx',
    'reset:word.read:session-1',
    'setup:word.read:read baseline',
    'call:word.read:session-1',
    'verify:word.read:direct-result:true',
    'reset:word.write:session-1',
    'setup:word.write:write baseline',
    'call:word.write:session-1',
    'verify:word.write:readback:true',
    'cleanupDocument:fixture.docx',
    'stopDaemon'
  ]);
  assert.equal(events.filter((event) => event === 'createDocument').length, 1);
  assert.equal(events.filter((event) => event.startsWith('activateAddin:')).length, 1);
  assert.equal(events.filter((event) => event.startsWith('waitForSession:')).length, 1);
  assert.equal(events.filter((event) => event.startsWith('call:')).length, 2);
});

test('shared Office tool E2E loop writes lifecycle and per-tool report evidence', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-report-'));
  const reportPath = join(dir, 'word-tools-report.json');
  const driver = {
    async startDaemon() {
      return { endpoint: 'http://127.0.0.1:8765/mcp' };
    },
    async listTools() {
      return ['word.read', 'word.write'];
    },
    async createDocument() {
      return { path: 'fixture.docx' };
    },
    async waitForSession() {
      return { sessionId: 'session-1', availableTools: ['word.read', 'word.write'] };
    },
    async resetContent() {},
    async setupContent() {},
    async callTool(toolCase) {
      return { ok: true, tool: toolCase.tool };
    },
    async verifyResult() {}
  };
  const cases = {
    'word.read': e2eCase('word.read', { setup: { actions: [{ tool: 'word.insert_paragraph', arguments: {} }] }, verify: directResult({ pathEquals: [{ path: 'ok', value: true }] }) }),
    'word.write': e2eCase('word.write', { setup: { actions: [{ tool: 'word.insert_paragraph', arguments: {} }] }, verify: readbackByTool('word.get_text', { arguments: {}, expect: { contains: ['updated'] } }) })
  };

  await runOfficeToolE2e({ host: 'Word', cases, driver, reportPath });

  const report = JSON.parse(readFileSync(reportPath, 'utf8'));
  assert.equal(report.schema_version, 1);
  assert.equal(report.kind, 'office_tool_e2e_report');
  assert.equal(report.host, 'Word');
  assert.equal(report.passed, true);
  assert.deepEqual(report.lifecycle_counts, {
    start_daemon: 1,
    list_tools: 1,
    create_document: 1,
    activate_addin: 0,
    wait_for_session: 1,
    cleanup_document: 0,
    stop_daemon: 0
  });
  assert.deepEqual(report.advertised_tools, ['word.read', 'word.write']);
  assert.deepEqual(report.session_available_tools, ['word.read', 'word.write']);
  assert.deepEqual(report.executed_tools, ['word.read', 'word.write']);
  assert.equal(report.tool_runs.length, 2);
  assert.equal(report.tool_runs[0].tool, 'word.read');
  assert.equal(report.tool_runs[0].verifier.kind, 'direct-result');
  assert.equal(report.tool_runs[0].passed, true);
  assert.equal(report.tool_runs[1].verifier.kind, 'readback');
  assert.equal(report.tool_runs[1].passed, true);
});

test('shared Office tool E2E loop writes failed report evidence before rethrowing', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-report-fail-'));
  const reportPath = join(dir, 'word-tools-report.json');
  const driver = {
    async startDaemon() {},
    async listTools() {
      return ['word.read'];
    },
    async createDocument() {
      return { path: 'fixture.docx' };
    },
    async waitForSession() {
      return { sessionId: 'session-1', availableTools: ['word.read'] };
    },
    async resetContent() {},
    async setupContent() {},
    async callTool() {
      return { ok: true };
    },
    async verifyResult() {
      throw new Error('verification failed');
    }
  };

  await assert.rejects(
    () => runOfficeToolE2e({
      host: 'Word',
      cases: { 'word.read': e2eCase('word.read', { setup: { actions: [{ tool: 'word.insert_paragraph', arguments: {} }] }, verify: directResult({ contains: ['read'] }) }) },
      driver,
      reportPath
    }),
    /verification failed/
  );

  const report = JSON.parse(readFileSync(reportPath, 'utf8'));
  assert.equal(report.passed, false);
  assert.equal(report.error.message, 'verification failed');
  assert.equal(report.tool_runs.length, 1);
  assert.equal(report.tool_runs[0].tool, 'word.read');
  assert.equal(report.tool_runs[0].passed, false);
  assert.equal(report.tool_runs[0].error.message, 'verification failed');
});

test('shared Office tool E2E loop records per-tool run metadata without body text', async () => {
  const records = [];
  const driver = {
    async startDaemon() {},
    async listTools() {
      return ['word.read'];
    },
    async createDocument() {
      return {};
    },
    async waitForSession() {
      return { sessionId: 'session-1', availableTools: ['word.read'] };
    },
    async resetContent(toolCase, _session, context) {
      records.push({ phase: 'reset', tool: toolCase.tool, run: context.run.id, hasSetupBody: Object.hasOwn(context.run, 'setup') });
    },
    async setupContent(toolCase, _session, context) {
      records.push({ phase: 'setup', tool: toolCase.tool, run: context.run.id, hasSetupBody: Object.hasOwn(context.run, 'setup') });
    },
    async callTool(toolCase, _session, context) {
      records.push({ phase: 'call', tool: toolCase.tool, run: context.run.id, requestId: context.run.requestId });
      return { ok: true };
    },
    async verifyResult(toolCase, _result, _session, context) {
      records.push({ phase: 'verify', tool: toolCase.tool, run: context.run.id, requestId: context.run.requestId });
    }
  };

  await runOfficeToolE2e({
    host: 'Word',
    cases: { 'word.read': e2eCase('word.read', { setup: 'secret body text', verify: 'direct-result' }) },
    driver
  });

  assert.deepEqual(records.map((record) => record.phase), ['reset', 'setup', 'call', 'verify']);
  assert.ok(records.every((record) => record.tool === 'word.read'));
  assert.ok(records.every((record) => record.run === 'word.read'));
  assert.ok(records.every((record) => record.hasSetupBody !== true));
  assert.match(records.find((record) => record.phase === 'call').requestId, /^e2e-word-read-/);
});

test('shared Office tool E2E loop preserves concrete readback verifier metadata', async () => {
  const verifiers = [];
  const driver = {
    async startDaemon() {},
    async listTools() {
      return ['word.write'];
    },
    async createDocument() {
      return {};
    },
    async waitForSession() {
      return { sessionId: 'session-1', availableTools: ['word.write'] };
    },
    async resetContent() {},
    async setupContent() {},
    async callTool() {
      return { ok: true };
    },
    async verifyResult(toolCase) {
      verifiers.push(toolCase.verify);
    }
  };
  const cases = {
    'word.write': e2eCase('word.write', {
      setup: { kind: 'document-text', text: 'baseline marker' },
      args: { text: 'updated marker' },
      verify: {
        kind: 'readback',
        readbackTool: 'word.get_text',
        readbackArguments: { limit: 20 },
        expect: { contains: ['updated marker'], notContains: ['baseline marker'] }
      }
    })
  };

  await runOfficeToolE2e({ host: 'Word', cases, driver });

  assert.deepEqual(verifiers, [{
    kind: 'readback',
    readbackTool: 'word.get_text',
    readbackArguments: { limit: 20 },
    expect: { contains: ['updated marker'], notContains: ['baseline marker'] }
  }]);
});

test('shared Office tool E2E loop carries setup bindings into call and verify steps', async () => {
  const observed = [];
  const driver = {
    async startDaemon() {},
    async listTools() {
      return ['powerpoint.update_table'];
    },
    async createDocument() {
      return {};
    },
    async waitForSession() {
      return { sessionId: 'session-1', availableTools: ['powerpoint.update_table'] };
    },
    async resetContent() {},
    async setupContent() {
      return { bindings: { table: { shape_id: 'shape-123' } } };
    },
    async callTool(_toolCase, session) {
      observed.push(['call', session.bindings.table.shape_id]);
      return { structuredContent: { ok: true } };
    },
    async verifyResult(_toolCase, _result, session) {
      observed.push(['verify', session.bindings.table.shape_id]);
    }
  };

  await runOfficeToolE2e({
    host: 'PowerPoint',
    cases: { 'powerpoint.update_table': e2eCase('powerpoint.update_table') },
    driver
  });

  assert.deepEqual(observed, [['call', 'shape-123'], ['verify', 'shape-123']]);
});

test('E2E case coverage accepts concrete readback verifiers with explicit expectations', () => {
  const toolCase = e2eCase('word.replace_text', {
    verify: {
      kind: 'readback',
      readbackTool: 'word.get_text',
      readbackArguments: { limit: 20 },
      expect: { contains: ['updated'], notContains: ['baseline'] }
    }
  });

  assert.equal(toolCase.verify.kind, 'readback');
  assert.equal(toolCase.verify.readbackTool, 'word.get_text');
  assert.deepEqual(toolCase.verify.expect, { contains: ['updated'], notContains: ['baseline'] });
});

test('E2E case coverage accepts resource readback verifiers with explicit expectations', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-resource-verifier-'));
  const addinRoot = join(dir, 'word');
  const publicRoot = join(addinRoot, 'public');
  mkdirSync(publicRoot, { recursive: true });
  writeFileSync(join(publicRoot, 'taskpane.js'), "const AVAILABLE_TOOLS = ['word.add_comment'];\n");
  const catalogPath = join(dir, 'catalog.rs');
  writeFileSync(catalogPath, 'pub const WORD_V1_TOOLS: &[&str] = &["word.add_comment"];\n');
  const toolCase = e2eCase('word.add_comment', {
    verify: {
      kind: 'readback',
      resource: 'office://word/${session_id}/comments',
      expect: { contains: ['E2E comment'] }
    }
  });

  assertE2eCaseCoverage({
    addinRoot,
    host: 'Word',
    catalogPath,
    cases: { 'word.add_comment': toolCase }
  });
  assert.equal(toolCase.verify.resource, 'office://word/${session_id}/comments');
});

test('E2E readback helpers create standard direct, tool, and resource verifiers', () => {
  assert.deepEqual(directResult({ pathEquals: [{ path: 'ok', value: true }] }), {
    kind: 'direct-result',
    expect: { pathEquals: [{ path: 'ok', value: true }] }
  });
  assert.deepEqual(readbackByTool('word.get_text', { arguments: { limit: 20 }, expect: { contains: ['updated'] } }), {
    kind: 'readback',
    readbackTool: 'word.get_text',
    readbackArguments: { limit: 20 },
    expect: { contains: ['updated'] }
  });
  assert.deepEqual(readbackByResource('office://word/${session_id}/comments', { expect: { contains: ['E2E comment'] } }), {
    kind: 'readback',
    resource: 'office://word/${session_id}/comments',
    expect: { contains: ['E2E comment'] }
  });
});

test('E2E readback helpers cover current Word, Excel, and PowerPoint object owners', () => {
  assert.deepEqual(wordReadback.documentText({ contains: ['paragraph'] }), readbackByTool('word.get_text', { arguments: { limit: 20 }, expect: { contains: ['paragraph'] } }));
  assert.deepEqual(wordReadback.paragraph(0, { contains: ['paragraph'] }), readbackByTool('word.get_paragraph', { arguments: { index: 0 }, expect: { contains: ['paragraph'] } }));
  assert.deepEqual(wordReadback.table(0, { contains: ['cell'] }), readbackByTool('word.read_table', { arguments: { table_index: 0 }, expect: { contains: ['cell'] } }));
  assert.deepEqual(wordReadback.contentControls('tag', { contains: ['title'] }), readbackByTool('word.list_content_controls', { arguments: { tag: 'tag' }, expect: { contains: ['title'] } }));
  assert.deepEqual(wordReadback.comments({ contains: ['comment'] }), readbackByResource('office://word/${session_id}/comments', { expect: { contains: ['comment'] } }));
  assert.deepEqual(wordReadback.trackChanges({ notContains: ['deleted'] }), readbackByResource('office://word/${session_id}/track_changes', { expect: { notContains: ['deleted'] } }));

  assert.deepEqual(excelReadback.workbook({ contains: ['Sheet1'] }), readbackByTool('excel.get_workbook_info', { arguments: {}, expect: { contains: ['Sheet1'] } }));
  assert.deepEqual(excelReadback.sheets({ contains: ['Sheet1'] }), readbackByTool('excel.list_sheets', { arguments: {}, expect: { contains: ['Sheet1'] } }));
  assert.deepEqual(excelReadback.range('Sheet1', 'A1:B2', { contains: ['value'] }), readbackByTool('excel.read_range', { arguments: { sheet: 'Sheet1', address: 'A1:B2' }, expect: { contains: ['value'] } }));
  assert.deepEqual(excelReadback.table('Table1', { contains: ['Table1'] }), readbackByTool('excel.update_table', { arguments: { table: 'Table1', action: 'metadata' }, expect: { contains: ['Table1'] } }));
  assert.deepEqual(excelReadback.chart('Sheet1', 'Chart 1', { contains: ['Chart 1'] }), readbackByTool('excel.update_chart', { arguments: { sheet: 'Sheet1', chart: 'Chart 1', action: 'metadata' }, expect: { contains: ['Chart 1'] } }));
  assert.deepEqual(excelReadback.pivotTable('Pivot1', { contains: ['Pivot1'] }), readbackByTool('excel.update_pivot_table', { arguments: { pivot_table: 'Pivot1', action: 'metadata' }, expect: { contains: ['Pivot1'] } }));

  assert.deepEqual(powerpointReadback.presentation({ contains: ['slides'] }), readbackByTool('powerpoint.get_presentation_info', { arguments: {}, expect: { contains: ['slides'] } }));
  assert.deepEqual(powerpointReadback.slides({ contains: ['Title'] }), readbackByTool('powerpoint.list_slides', { arguments: {}, expect: { contains: ['Title'] } }));
  assert.deepEqual(powerpointReadback.shapes(0, { contains: ['rectangle'] }), readbackByTool('powerpoint.list_shapes', { arguments: { slide_index: 0 }, expect: { contains: ['rectangle'] } }));
  assert.deepEqual(powerpointReadback.text(0, { contains: ['text'] }), readbackByTool('powerpoint.read_text', { arguments: { slide_index: 0 }, expect: { contains: ['text'] } }));
  assert.deepEqual(powerpointReadback.table(0, '${table.shape_id}', { contains: ['cell'] }), readbackByTool('powerpoint.read_table', { arguments: { slide_index: 0, shape_id: '${table.shape_id}' }, expect: { contains: ['cell'] } }));
});

test('E2E case coverage checks daemon catalog tools as well as task pane tools', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-catalog-'));
  const addinRoot = join(dir, 'word');
  const publicRoot = join(addinRoot, 'public');
  mkdirSync(publicRoot, { recursive: true });
  writeFileSync(join(publicRoot, 'taskpane.js'), "const AVAILABLE_TOOLS = ['word.read'];\n");
  const catalogPath = join(dir, 'catalog.rs');
  writeFileSync(catalogPath, 'pub const WORD_V1_TOOLS: &[&str] = &["word.read"];\n');

  assertE2eCaseCoverage({
    addinRoot,
    host: 'Word',
    catalogPath,
    cases: { 'word.read': e2eCase('word.read', { verify: 'direct-result' }) }
  });
});

test('E2E case coverage fails when daemon catalog exposes an uncovered tool', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-catalog-missing-'));
  const addinRoot = join(dir, 'word');
  const publicRoot = join(addinRoot, 'public');
  mkdirSync(publicRoot, { recursive: true });
  writeFileSync(join(publicRoot, 'taskpane.js'), "const AVAILABLE_TOOLS = ['word.read'];\n");
  const catalogPath = join(dir, 'catalog.rs');
  writeFileSync(catalogPath, 'pub const WORD_V1_TOOLS: &[&str] = &["word.read", "word.catalog_only"];\n');

  assert.throws(
    () => assertE2eCaseCoverage({
      addinRoot,
      host: 'Word',
      catalogPath,
      cases: { 'word.read': e2eCase('word.read', { verify: 'direct-result' }) }
    }),
    /Word E2E cases must cover every daemon catalog tool exactly/
  );
});

test('daemonCatalogTools reads the Rust runtime catalog for each Office host', () => {
  assert.ok(daemonCatalogTools('Word').includes('word.get_text'));
  assert.ok(daemonCatalogTools('Excel').includes('excel.read_range'));
  assert.ok(daemonCatalogTools('PowerPoint').includes('powerpoint.list_slides'));
});

test('shared Office tool E2E loop fails when session tools and case table differ', async () => {
  const driver = {
    async startDaemon() {},
    async listTools() {
      return ['word.read'];
    },
    async createDocument() {
      return {};
    },
    async waitForSession() {
      return { sessionId: 'session-1', availableTools: ['word.read', 'word.missing'] };
    },
    async resetContent() {},
    async setupContent() {},
    async callTool() {},
    async verifyResult() {},
    async cleanupDocument() {},
    async stopDaemon() {}
  };
  const cases = {
    'word.read': e2eCase('word.read')
  };

  await assert.rejects(
    () => runOfficeToolE2e({ host: 'Word', cases, driver }),
    /Word E2E session tools must match the case table exactly/
  );
});

test('shared Office tool E2E loop fails when daemon tools/list exposes an uncovered host tool', async () => {
  const driver = {
    async startDaemon() {},
    async listTools() {
      return ['office.list_sessions', 'word.read', 'word.catalog_only', 'excel.read_range'];
    },
    async createDocument() {
      return {};
    },
    async waitForSession() {
      return { sessionId: 'session-1', availableTools: ['word.read'] };
    },
    async resetContent() {},
    async setupContent() {},
    async callTool() {},
    async verifyResult() {},
    async cleanupDocument() {},
    async stopDaemon() {}
  };

  await assert.rejects(
    () => runOfficeToolE2e({ host: 'Word', cases: { 'word.read': e2eCase('word.read') }, driver }),
    /Word E2E daemon tools\/list must match the case table exactly/
  );
});

test('shared Office tool E2E loop still cleans up when a verifier fails', async () => {
  const events = [];
  const driver = {
    async startDaemon() {
      events.push('startDaemon');
    },
    async listTools() {
      events.push('listTools');
      return ['word.read'];
    },
    async createDocument() {
      events.push('createDocument');
      return {};
    },
    async waitForSession() {
      events.push('waitForSession');
      return { sessionId: 'session-1', availableTools: ['word.read'] };
    },
    async resetContent() {
      events.push('reset');
    },
    async setupContent() {
      events.push('setup');
    },
    async callTool() {
      events.push('call');
      return {};
    },
    async verifyResult() {
      events.push('verify');
      throw new Error('verification failed');
    },
    async cleanupDocument() {
      events.push('cleanupDocument');
    },
    async stopDaemon() {
      events.push('stopDaemon');
    }
  };

  await assert.rejects(
    () => runOfficeToolE2e({ host: 'Word', cases: { 'word.read': e2eCase('word.read') }, driver }),
    /verification failed/
  );
  assert.deepEqual(events, ['startDaemon', 'listTools', 'createDocument', 'waitForSession', 'reset', 'setup', 'call', 'verify', 'cleanupDocument', 'stopDaemon']);
});

test('external Office E2E driver adapter exchanges one JSON request per step', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-driver-'));
  const logPath = join(dir, 'calls.jsonl');
  const driverPath = join(dir, 'mock-driver.mjs');
  const reportPath = join(dir, 'report.json');
  writeFileSync(driverPath, `
import { appendFileSync } from 'node:fs';
const input = await new Promise((resolve) => {
  let body = '';
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', (chunk) => { body += chunk; });
  process.stdin.on('end', () => resolve(body));
});
const request = JSON.parse(input);
appendFileSync(${JSON.stringify(logPath)}, JSON.stringify({ host: request.host, step: request.step, context: request.context }) + '\\n');
const responses = {
  startDaemon: { endpoint: 'http://127.0.0.1:8765/mcp' },
  listTools: ['word.read'],
  createDocument: { path: 'external.docx' },
  activateAddin: { activated: true },
  waitForSession: { sessionId: 'session-1', availableTools: ['word.read'] },
  resetContent: { reset: true },
  setupContent: { setup: true },
  callTool: { ok: true },
  verifyResult: { verified: true },
  cleanupDocument: { cleaned: true },
  stopDaemon: { stopped: true }
};
process.stdout.write(JSON.stringify(responses[request.step] || {}));
`);
  const previousRun = process.env.OFFICE_MCP_RUN_E2E;
  const previousDriver = process.env.OFFICE_MCP_E2E_DRIVER;
  process.env.OFFICE_MCP_RUN_E2E = '1';
  process.env.OFFICE_MCP_E2E_DRIVER = driverPath;
  try {
    await runOfficeToolE2e({
      host: 'Word',
      cases: { 'word.read': e2eCase('word.read', { verify: 'direct-result' }) },
      driver: requireOfficeE2eDriver('Word'),
      reportPath
    });
  } finally {
    restoreEnv('OFFICE_MCP_RUN_E2E', previousRun);
    restoreEnv('OFFICE_MCP_E2E_DRIVER', previousDriver);
  }

  const calls = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
  assert.deepEqual(calls.map((call) => call.step), [
    'startDaemon',
    'listTools',
    'createDocument',
    'activateAddin',
    'waitForSession',
    'resetContent',
    'setupContent',
    'callTool',
    'verifyResult',
    'cleanupDocument',
    'stopDaemon'
  ]);
  assert.equal(calls[0].host, 'Word');
  assert.equal(calls[5].context.toolCase.tool, 'word.read');
  assert.equal(calls[7].context.session.sessionId, 'session-1');
  const report = JSON.parse(readFileSync(reportPath, 'utf8'));
  assert.equal(report.passed, true);
  assert.deepEqual(report.lifecycle_counts, {
    start_daemon: 1,
    list_tools: 1,
    create_document: 1,
    activate_addin: 1,
    wait_for_session: 1,
    cleanup_document: 1,
    stop_daemon: 1
  });
  assert.deepEqual(report.executed_tools, ['word.read']);
});

test('external Office E2E driver adapter reports failed step stderr', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-driver-fail-'));
  const driverPath = join(dir, 'failing-driver.mjs');
  writeFileSync(driverPath, `
console.error('driver exploded');
process.exit(7);
`);
  const previousRun = process.env.OFFICE_MCP_RUN_E2E;
  const previousDriver = process.env.OFFICE_MCP_E2E_DRIVER;
  process.env.OFFICE_MCP_RUN_E2E = '1';
  process.env.OFFICE_MCP_E2E_DRIVER = driverPath;
  try {
    await assert.rejects(
      () => requireOfficeE2eDriver('Word').startDaemon({ host: 'Word' }),
      /Word E2E driver step startDaemon failed with exit code 7: driver exploded/
    );
  } finally {
    restoreEnv('OFFICE_MCP_RUN_E2E', previousRun);
    restoreEnv('OFFICE_MCP_E2E_DRIVER', previousDriver);
  }
});

function restoreEnv(name, value) {
  if (value === undefined) {
    delete process.env[name];
  } else {
    process.env[name] = value;
  }
}
