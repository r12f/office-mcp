import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const REPO_ROOT = join(process.cwd(), '..', '..', '..');

function storage(initial = {}) {
  const values = new Map(Object.entries(initial));
  return {
    getItem(key) {
      return values.has(key) ? values.get(key) : null;
    },
    removeItem(key) {
      values.delete(key);
    },
    setItem(key, value) {
      values.set(key, String(value));
    }
  };
}

function loadChannel(options = {}) {
  const source = readFileSync(join(REPO_ROOT, 'src', 'office-ctl', 'common', 'addin-channel.js'), 'utf8');
  const cryptoValue = options.crypto === undefined
    ? { randomUUID: options.randomUUID || (() => 'generated-id') }
    : options.crypto;
  const context = vm.createContext({
    Error,
    JSON,
    Math,
    Object,
    String,
    URL,
    crypto: cryptoValue,
    globalThis: {},
    localStorage: options.localStorage || storage(),
    location: options.location || { origin: 'https://localhost:8765' },
    sessionStorage: options.sessionStorage || storage()
  });
  context.globalThis = context;
  vm.runInContext(source, context, { filename: 'addin-channel.js' });
  return context.OfficeCtlAddinChannel;
}

test('common add-in channel manages endpoint configuration', () => {
  const localStorage = storage();
  const channel = loadChannel({ localStorage });

  assert.equal(channel.configuredEndpoint(), 'wss://localhost:8765/addin');
  assert.equal(channel.saveEndpointOverride('wss://localhost:8765/custom-addin'), 'wss://localhost:8765/custom-addin');
  assert.equal(channel.configuredEndpoint(), 'wss://localhost:8765/custom-addin');
  assert.equal(channel.currentOriginEndpoint(), 'wss://localhost:8765/addin');
  assert.equal(channel.clearEndpointOverride(), 'wss://localhost:8765/addin');
  assert.equal(channel.configuredEndpoint(), 'wss://localhost:8765/addin');
  assert.throws(() => channel.validateEndpoint('https://localhost:8765/addin'), /wss:\/\/localhost/);
  assert.throws(() => channel.validateEndpoint('wss://example.invalid/addin'), /wss:\/\/localhost/);
});

test('common add-in channel clears stale invalid endpoint overrides', () => {
  const localStorage = storage({ 'office-mcp.addin-endpoint': 'https://localhost:8765/addin' });
  const channel = loadChannel({ localStorage });

  assert.equal(channel.configuredEndpoint(), 'wss://localhost:8765/addin');
  assert.equal(localStorage.getItem('office-mcp.addin-endpoint'), null);
});

test('common add-in channel clears endpoint overrides from a stale manifest origin', () => {
  const localStorage = storage({ 'office-mcp.addin-endpoint': 'wss://localhost:8766/addin' });
  const channel = loadChannel({ localStorage, location: { origin: 'https://localhost:8765' } });

  assert.equal(channel.configuredEndpoint(), 'wss://localhost:8765/addin');
  assert.equal(localStorage.getItem('office-mcp.addin-endpoint'), null);
});

test('common add-in channel derives fallback endpoint from the manifest origin', () => {
  const channel = loadChannel({ location: { origin: 'https://localhost:8766' } });

  assert.equal(channel.currentOriginEndpoint(), 'wss://localhost:8766/addin');
  assert.equal(channel.configuredEndpoint(), 'wss://localhost:8766/addin');
});

test('common add-in channel persists runtime IDs and register request IDs', () => {
  const sessionStorage = storage();
  let counter = 0;
  const channel = loadChannel({ sessionStorage, randomUUID: () => `id-${++counter}` });

  const firstIds = channel.runtimeIds();
  assert.equal(firstIds.instanceId, 'id-1');
  assert.equal(firstIds.sessionId, 'id-2');
  const secondIds = channel.runtimeIds();
  assert.equal(secondIds.instanceId, 'id-1');
  assert.equal(secondIds.sessionId, 'id-2');
  channel.rememberRegisterRequest('register-1');
  assert.equal(channel.isRegisterResponse({ id: 'register-1' }), true);
  assert.equal(channel.isRegisterResponse({ id: 'other' }), false);
  channel.clearRegisterRequest();
  assert.equal(channel.isRegisterResponse({ id: 'register-1' }), false);
});

test('common add-in channel falls back when Office WebView lacks crypto.randomUUID', () => {
  const sessionStorage = storage();
  const channel = loadChannel({
    crypto: {},
    sessionStorage,
    random: () => 0.5
  });

  const ids = channel.runtimeIds();
  assert.match(ids.instanceId, /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
  assert.match(ids.sessionId, /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
  assert.match(channel.createRequestId({ random: () => 0.25 }), /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
});

test('common add-in channel parses and sends JSON-RPC messages', () => {
  const channel = loadChannel();
  const sent = [];
  const socket = { readyState: 1, send: (value) => sent.push(value) };

  assert.deepEqual(channel.parseJsonRpc('{"jsonrpc":"2.0","id":"1"}'), { jsonrpc: '2.0', id: '1' });
  assert.equal(channel.parseJsonRpc('not json'), null);
  assert.equal(channel.sendJsonRpc(socket, { jsonrpc: '2.0', method: 'ping' }), true);
  assert.equal(channel.reply(socket, '1', { ok: true }), true);
  assert.equal(channel.sendJsonRpc({ readyState: 3, send: () => assert.fail('must not send') }, {}), false);
  assert.deepEqual(sent.map((value) => JSON.parse(value)), [
    { jsonrpc: '2.0', method: 'ping' },
    { jsonrpc: '2.0', id: '1', result: { ok: true } }
  ]);
  assert.equal(channel.reconnectDelay(2, () => 0.25), 2125);
});

test('common add-in channel builds and classifies protocol messages', () => {
  const channel = loadChannel();

  assert.equal(JSON.stringify(channel.registerRequest('register-1', { instance_id: 'instance-1' })), JSON.stringify({
    jsonrpc: '2.0',
    id: 'register-1',
    method: 'register',
    params: { instance_id: 'instance-1' }
  }));
  assert.equal(JSON.stringify(channel.sessionAddedNotification({ session_id: 'session-1' })), JSON.stringify({
    jsonrpc: '2.0',
    method: 'session.added',
    params: { session_id: 'session-1' }
  }));
  assert.equal(JSON.stringify(channel.sessionUpdatedNotification({ session_id: 'session-1', patch: { available_tools: [] } })), JSON.stringify({
    jsonrpc: '2.0',
    method: 'session.updated',
    params: { session_id: 'session-1', patch: { available_tools: [] } }
  }));
  const registerResult = channel.registerResult({ result: { server_version: '0.2.0', protocol_version: '1.1' } }, '1.0');
  assert.equal(registerResult.serverVersion, '0.2.0');
  assert.equal(registerResult.protocolVersion, '1.1');
  const fallbackRegisterResult = channel.registerResult({}, '1.0');
  assert.equal(fallbackRegisterResult.serverVersion, 'Unknown');
  assert.equal(fallbackRegisterResult.protocolVersion, '1.0');
  assert.equal(channel.isToolInvoke({ method: 'tool.invoke' }), true);
  assert.equal(channel.isPing({ method: 'ping' }), true);
  assert.equal(channel.isToolCancel({ method: 'tool.cancel', params: { request_id: 'request-1' } }), true);
  assert.equal(channel.isToolCancel({ method: 'tool.cancel', params: {} }), false);
});
