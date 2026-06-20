import assert from 'node:assert/strict';
import test from 'node:test';
import { e2eCase, runOfficeToolE2e } from './tool-e2e-contract.mjs';

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
