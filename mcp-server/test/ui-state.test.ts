import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { authorizeUiRequest, removeUiRuntimeFile, UiStateStore, writeUiRuntimeFile } from '../src/ui-state.js';

test('UI state snapshots redact secrets and cap recent command history', () => {
  let tick = 0;
  const store = new UiStateStore({
    version: '0.1.0',
    mcpEndpoint: 'http://127.0.0.1:8800/mcp',
    addinEndpoint: 'https://localhost:8765/addin',
    sessions: () => [],
    now: () => new Date(Date.UTC(2026, 0, 1, 0, 0, tick++))
  });

  for (let index = 0; index < 12; index += 1) {
    const id = store.startCommand({
      tool: 'word.insert_paragraph',
      session_id: '11111111-1111-4111-8111-111111111111',
      user_intent: `token=secret-${index} insert private body text`
    });
    store.finishCommand(id, {
      ok: false,
      error: {
        office_mcp_code: 'IRM_DENIED',
        message: `shared_secret=secret-${index} Word denied the edit.`,
        tool: 'word.insert_paragraph',
        retriable: false,
        partial_effect: 'none'
      }
    });
  }

  const snapshot = store.snapshot();
  assert.equal(snapshot.recent_commands.length, 10);
  assert.equal(snapshot.current_tasks.length, 0);
  const serialized = JSON.stringify(snapshot);
  assert.doesNotMatch(serialized, /secret-11/);
  assert.doesNotMatch(serialized, /shared_secret=secret/);
  assert.match(serialized, /shared_secret=\[redacted\]/);
  assert.match(serialized, /token=\[redacted\]/);
});

test('UI token accepts bearer or explicit UI token header only', () => {
  assert.equal(authorizeUiRequest({ authorization: 'Bearer ui-token' }, 'ui-token'), true);
  assert.equal(authorizeUiRequest({ 'x-office-mcp-ui-token': 'ui-token' }, 'ui-token'), true);
  assert.equal(authorizeUiRequest({ authorization: 'Bearer wrong' }, 'ui-token'), false);
  assert.equal(authorizeUiRequest({}, 'ui-token'), false);
});

test('UI state tracks clients and in-flight request counts', () => {
  const store = new UiStateStore({
    version: '0.1.0',
    mcpEndpoint: 'http://127.0.0.1:8800/mcp',
    addinEndpoint: 'https://localhost:8765/addin',
    sessions: () => []
  });

  const clientId = store.registerClient({ transport: 'http', name: 'test-client' });
  const commandId = store.startCommand({ client_id: clientId, client_name: 'test-client', tool: 'word.get_text' });

  let snapshot = store.snapshot();
  assert.equal(snapshot.clients.length, 1);
  assert.equal(snapshot.clients[0].name, 'test-client');
  assert.equal(snapshot.clients[0].in_flight_request_count, 1);

  store.finishCommand(commandId, { ok: true, data: {} });
  snapshot = store.snapshot();
  assert.equal(snapshot.clients[0].in_flight_request_count, 0);

  store.unregisterClient(clientId);
  assert.equal(store.snapshot().clients.length, 0);
});

test('UI runtime file is written for the native tray and removed on shutdown', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-ui-runtime-'));
  const path = join(dir, 'ui-runtime.json');
  try {
    writeUiRuntimeFile(path, {
      origin: 'https://localhost:8765',
      stateUrl: 'https://localhost:8765/ui/state',
      uiUrl: 'https://localhost:8765/ui/',
      token: 'ui-token',
      pid: 123,
      createdAt: '2026-01-01T00:00:00.000Z'
    });

    const parsed = JSON.parse(readFileSync(path, 'utf8')) as { token?: string; stateUrl?: string };
    assert.equal(parsed.token, 'ui-token');
    assert.equal(parsed.stateUrl, 'https://localhost:8765/ui/state');

    removeUiRuntimeFile(path);
    assert.equal(existsSync(path), false);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
