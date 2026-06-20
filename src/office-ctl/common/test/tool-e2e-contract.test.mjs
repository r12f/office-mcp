import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { e2eCase, requireRealOfficeE2eDriver, runOfficeToolE2e } from './tool-e2e-contract.mjs';

test('shared Office tool E2E loop drives daemon, document, setup, calls, verification, and cleanup', async () => {
  const events = [];
  const driver = {
    async startDaemon() {
      events.push('startDaemon');
      return { endpoint: 'http://127.0.0.1:0/mcp' };
    },
    async createDocument() {
      events.push('createDocument');
      return { path: 'fixture.docx' };
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
    'createDocument',
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
});

test('shared Office tool E2E loop records per-tool run metadata without body text', async () => {
  const records = [];
  const driver = {
    async startDaemon() {},
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

test('shared Office tool E2E loop fails when session tools and case table differ', async () => {
  const driver = {
    async startDaemon() {},
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

test('shared Office tool E2E loop still cleans up when a verifier fails', async () => {
  const events = [];
  const driver = {
    async startDaemon() {
      events.push('startDaemon');
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
  assert.deepEqual(events, ['startDaemon', 'createDocument', 'waitForSession', 'reset', 'setup', 'call', 'verify', 'cleanupDocument', 'stopDaemon']);
});

test('external Office E2E driver adapter exchanges one JSON request per step', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-e2e-driver-'));
  const logPath = join(dir, 'calls.jsonl');
  const driverPath = join(dir, 'mock-driver.mjs');
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
  createDocument: { path: 'external.docx' },
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
      driver: requireRealOfficeE2eDriver('Word')
    });
  } finally {
    restoreEnv('OFFICE_MCP_RUN_E2E', previousRun);
    restoreEnv('OFFICE_MCP_E2E_DRIVER', previousDriver);
  }

  const calls = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
  assert.deepEqual(calls.map((call) => call.step), [
    'startDaemon',
    'createDocument',
    'waitForSession',
    'resetContent',
    'setupContent',
    'callTool',
    'verifyResult',
    'cleanupDocument',
    'stopDaemon'
  ]);
  assert.equal(calls[0].host, 'Word');
  assert.equal(calls[3].context.toolCase.tool, 'word.read');
  assert.equal(calls[5].context.session.sessionId, 'session-1');
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
      () => requireRealOfficeE2eDriver('Word').startDaemon({ host: 'Word' }),
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
