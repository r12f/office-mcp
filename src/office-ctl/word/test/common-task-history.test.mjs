import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const REPO_ROOT = join(process.cwd(), '..', '..', '..');

function loadTaskHistory() {
  const source = readFileSync(join(REPO_ROOT, 'src', 'office-ctl', 'common', 'task-history.js'), 'utf8');
  const context = vm.createContext({
    Date,
    Object,
    Set,
    String,
    globalThis: {}
  });
  context.globalThis = context;
  vm.runInContext(source, context, { filename: 'task-history.js' });
  return context.OfficeCtlTaskHistory;
}

test('common task history tracks current task and bounded completed history', () => {
  const { TaskHistoryStore } = loadTaskHistory();
  let now = 1000;
  const store = new TaskHistoryStore({
    historyLimit: 2,
    now: () => now,
    toIso: () => `iso-${now}`,
    redactText: (value) => String(value).replace('secret', '[redacted]')
  });

  store.start('request-1', 'word.insert_paragraph', { client_meta: { user_intent: 'insert secret text' } }, 5000);
  let snapshot = store.snapshot();
  assert.equal(snapshot.currentTask.requestId, 'request-1');
  assert.equal(snapshot.currentTask.tool, 'word.insert_paragraph');
  assert.equal(snapshot.currentTask.deadlineAt, 6000);
  assert.equal(snapshot.currentTask.userIntent, 'insert secret text');

  now = 1100;
  const finished = store.finish('request-1', 'failure', 100, { message: 'secret failed', partialEffect: 'none' });
  assert.equal(finished.requestId, 'request-1');
  assert.equal(finished.error.message, '[redacted] failed');
  assert.equal(finished.error.partial_effect, 'none');
  store.start('request-2', 'word.get_text', {}, null);
  store.finish('request-2', 'success', 10);
  store.start('request-3', 'word.save', {}, null);
  store.finish('request-3', 'success', 20);

  snapshot = store.snapshot();
  assert.equal(snapshot.currentTask, null);
  assert.equal(snapshot.history.length, 2);
  assert.equal(snapshot.history[0].tool, 'word.save');
  assert.equal(snapshot.history[0].requestId, 'request-3');
  assert.equal(snapshot.history[1].tool, 'word.get_text');
  assert.equal(snapshot.history[1].requestId, 'request-2');
});

test('common task history records cancellation state', () => {
  const { TaskHistoryStore } = loadTaskHistory();
  const store = new TaskHistoryStore({ now: () => 2000 });

  store.cancel('request-1');
  assert.equal(store.isCancelled('request-1'), true);
  store.start('request-1', 'word.get_text', {}, null);
  assert.equal(store.snapshot().currentTask.cancelRequested, true);
  assert.equal(store.consumeCancellation('request-1'), true);
  assert.equal(store.isCancelled('request-1'), false);
});
