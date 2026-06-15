import assert from 'node:assert/strict';
import test from 'node:test';
import { EventEmitter } from 'node:events';
import { SessionRegistry, ToolInvocationError } from '../src/session-registry.js';
import type { AddinConnection } from '../src/types.js';

test('lists sessions and returns full session info', () => {
  const registry = new SessionRegistry();
  const connection = fakeConnection(async () => ({ ok: true, data: {} }));
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: '11111111-1111-4111-8111-111111111111',
    instance_id: connection.runtime.instance_id,
    document: { title: 'Example.docx', protection: { kind: null, rights: null, rights_source: 'unavailable' } },
    available_tools: ['word.get_text'],
    is_active: null
  });

  assert.equal(registry.listSessions().length, 1);
  assert.deepEqual(registry.getSessionInfo('11111111-1111-4111-8111-111111111111')?.available_tools, ['word.get_text']);
});

test('rejects unsupported session tool before dispatch', async () => {
  const registry = new SessionRegistry();
  const connection = fakeConnection(async () => {
    throw new Error('should not dispatch');
  });
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: '22222222-2222-4222-8222-222222222222',
    instance_id: connection.runtime.instance_id,
    document: {},
    available_tools: ['word.get_text'],
    is_active: null
  });

  await assert.rejects(
    registry.invoke('22222222-2222-4222-8222-222222222222', 'word.insert_paragraph', { session_id: '22222222-2222-4222-8222-222222222222' }, 1000),
    (error) => error instanceof ToolInvocationError && error.failure.office_mcp_code === 'HOST_CAPABILITY_UNAVAILABLE'
  );
});

test('returns NO_SESSIONS when a tool targets a session before any add-in connects', async () => {
  const registry = new SessionRegistry();

  await assert.rejects(
    registry.invoke('44444444-4444-4444-8444-444444444444', 'word.get_text', { session_id: '44444444-4444-4444-8444-444444444444' }, 1000),
    (error) => error instanceof ToolInvocationError && error.failure.office_mcp_code === 'NO_SESSIONS' && error.failure.retriable === true
  );
});

test('returns SESSION_NOT_FOUND when other sessions exist but the requested id is unknown', async () => {
  const registry = new SessionRegistry();
  const connection = fakeConnection(async () => ({ ok: true, data: {} }));
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: '55555555-5555-4555-8555-555555555555',
    instance_id: connection.runtime.instance_id,
    document: {},
    available_tools: ['word.get_text'],
    is_active: null
  });

  await assert.rejects(
    registry.invoke('66666666-6666-4666-8666-666666666666', 'word.get_text', { session_id: '66666666-6666-4666-8666-666666666666' }, 1000),
    (error) => error instanceof ToolInvocationError && error.failure.office_mcp_code === 'SESSION_NOT_FOUND'
  );
});

test('forwards supported tool calls to the owning add-in connection', async () => {
  const registry = new SessionRegistry();
  const seen: unknown[] = [];
  const connection = fakeConnection(async (_sessionId, tool, args) => {
    seen.push({ tool, args });
    return { ok: true, data: { text: 'hello' } };
  });
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: '33333333-3333-4333-8333-333333333333',
    instance_id: connection.runtime.instance_id,
    document: {},
    available_tools: ['word.get_text'],
    is_active: null
  });

  const result = await registry.invoke('33333333-3333-4333-8333-333333333333', 'word.get_text', { session_id: '33333333-3333-4333-8333-333333333333' }, 1000);
  assert.deepEqual(result, { ok: true, data: { text: 'hello' } });
  assert.equal(seen.length, 1);
});

function fakeConnection(invokeTool: AddinConnection['invokeTool']): AddinConnection {
  const socket = new EventEmitter() as AddinConnection['socket'];
  Object.assign(socket, {
    OPEN: 1,
    readyState: 1,
    close() {}
  });
  return {
    socket,
    runtime: {
      instance_id: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa',
      host: { app: 'word' },
      add_in: { version: '0.1.0', protocol_version: '1.0' },
      registered_at: new Date().toISOString()
    },
    pending: new Map(),
    queue: Promise.resolve(),
    invokeTool
  };
}
