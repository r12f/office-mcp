import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const REPO_ROOT = join(process.cwd(), '..', '..', '..');

function loadLogger() {
  const source = readFileSync(join(REPO_ROOT, 'src', 'office-ctl', 'common', 'logger.js'), 'utf8');
  const entries = [];
  const console = {
    info: (...args) => entries.push(['info', ...args]),
    warn: (...args) => entries.push(['warn', ...args]),
    error: (...args) => entries.push(['error', ...args])
  };
  const context = vm.createContext({
    Array,
    Object,
    String,
    console,
    globalThis: {}
  });
  context.globalThis = context;
  vm.runInContext(source, context, { filename: 'logger.js' });
  return { loggerModule: context.OfficeCtlLogger, entries };
}

test('common logger writes scoped redacted records', () => {
  const { loggerModule, entries } = loadLogger();
  const logger = new loggerModule.AddinLogger({
    redactText: (value) => String(value).replace(/secret/g, '[redacted]')
  });

  logger.info('session.added', {
    sessionId: 'session-1',
    document: { title: 'Plan', text: 'secret body', nested: { content: 'secret paragraph' } },
    userIntent: 'insert secret summary'
  });

  assert.equal(entries.length, 1);
  assert.equal(entries[0][0], 'info');
  assert.equal(entries[0][1], 'office-mcp session.added');
  assert.equal(entries[0][2].document.text, '[redacted]');
  assert.equal(entries[0][2].document.nested.content, '[redacted]');
  assert.equal(entries[0][2].userIntent, 'insert [redacted] summary');
});
