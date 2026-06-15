import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { createServer } from 'node:http';
import test from 'node:test';
import type { IncomingMessage } from 'node:http';
import type { Duplex } from 'node:stream';
import { once } from 'node:events';
import { WebSocket, WebSocketServer } from 'ws';
import { createAddinChannel, handleUpgrade } from '../src/addin-channel.js';
import type { DaemonConfig } from '../src/config.js';
import { SessionRegistry } from '../src/session-registry.js';
import { ToolInvocationError } from '../src/session-registry.js';

test('rejects add-in websocket upgrades from foreign origins', () => {
  const wss = new WebSocketServer({ noServer: true });
  let destroyed = false;
  let written = '';
  const socket = new EventEmitter() as Duplex;
  Object.assign(socket, {
    write(chunk: string) { written += chunk; return true; },
    destroy() { destroyed = true; return socket; }
  });
  const req = { url: '/addin', headers: { origin: 'https://evil.example' } } as IncomingMessage;

  handleUpgrade(config(), wss, req, socket, Buffer.alloc(0));

  assert.equal(destroyed, true);
  assert.match(written, /403 Forbidden/);
  wss.close();
});

test('closes add-in websocket after missed heartbeat threshold', async () => {
  const testConfig = config();
  testConfig.addin.heartbeatIntervalSec = 0.1;
  testConfig.addin.heartbeatTimeoutSec = 0.1;
  const registry = new SessionRegistry();
  const wss = createAddinChannel(testConfig, registry);
  const server = createServer();
  server.on('upgrade', (req, socket, head) => handleUpgrade(testConfig, wss, req, socket, head));
  await listen(server);
  const address = server.address();
  assert(address && typeof address === 'object');

  const client = new WebSocket(`ws://127.0.0.1:${address.port}/addin`, { origin: testConfig.addin.origin });
  try {
    await once(client, 'open');
    client.send(JSON.stringify({
      jsonrpc: '2.0',
      id: 'register-1',
      method: 'register',
      params: {
        instance_id: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa',
        host: { app: 'word' },
        add_in: { version: '0.1.0', protocol_version: '1.0' }
      }
    }));
    client.send(JSON.stringify({
      jsonrpc: '2.0',
      method: 'session.added',
      params: {
        session_id: 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb',
        instance_id: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa',
        document: { title: 'Heartbeat.docx' },
        available_tools: ['word.get_text'],
        is_active: null
      }
    }));

    await waitFor(() => registry.listSessions().length === 1);
    const [code] = await once(client, 'close') as [number, Buffer];
    assert.equal(code, 4002);
    assert.equal(registry.getSessionInfo('bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb')?.status, 'stale');
  } finally {
    client.close();
    wss.close();
    await close(server);
  }
});

test('sends tool.cancel and returns TIMEOUT when an add-in invocation exceeds its deadline', async () => {
  const testConfig = config();
  testConfig.addin.heartbeatIntervalSec = 60;
  const registry = new SessionRegistry();
  const wss = createAddinChannel(testConfig, registry);
  const server = createServer();
  server.on('upgrade', (req, socket, head) => handleUpgrade(testConfig, wss, req, socket, head));
  await listen(server);
  const address = server.address();
  assert(address && typeof address === 'object');

  const client = new WebSocket(`ws://127.0.0.1:${address.port}/addin`, { origin: testConfig.addin.origin });
  const messages: unknown[] = [];
  client.on('message', (raw) => messages.push(JSON.parse(raw.toString())));
  try {
    await once(client, 'open');
    client.send(JSON.stringify({
      jsonrpc: '2.0',
      id: 'register-1',
      method: 'register',
      params: {
        instance_id: 'cccccccc-cccc-4ccc-8ccc-cccccccccccc',
        host: { app: 'word' },
        add_in: { version: '0.1.0', protocol_version: '1.0' }
      }
    }));
    client.send(JSON.stringify({
      jsonrpc: '2.0',
      method: 'session.added',
      params: {
        session_id: 'dddddddd-dddd-4ddd-8ddd-dddddddddddd',
        instance_id: 'cccccccc-cccc-4ccc-8ccc-cccccccccccc',
        document: { title: 'Timeout.docx' },
        available_tools: ['word.get_text'],
        is_active: null
      }
    }));
    await waitFor(() => registry.getSessionInfo('dddddddd-dddd-4ddd-8ddd-dddddddddddd')?.status === 'active');

    await assert.rejects(
      registry.invoke('dddddddd-dddd-4ddd-8ddd-dddddddddddd', 'word.get_text', { session_id: 'dddddddd-dddd-4ddd-8ddd-dddddddddddd' }, 25),
      (error) => error instanceof ToolInvocationError && error.failure.office_mcp_code === 'TIMEOUT' && error.failure.partial_effect === 'unknown'
    );
    await waitFor(() => messages.some((message) => isMethod(message, 'tool.cancel')));
    const cancel = messages.find((message) => isMethod(message, 'tool.cancel')) as { params: { reason: string; request_id: string } };
    assert.equal(cancel.params.reason, 'deadline_expired');
    assert.equal(typeof cancel.params.request_id, 'string');
  } finally {
    client.close();
    wss.close();
    await close(server);
  }
});

function config(): DaemonConfig {
  return {
    addin: {
      host: 'localhost',
      port: 8765,
      origin: 'https://localhost:8765',
      pfxPath: '.office-mcp-localhost.pfx',
      pfxPassphrase: 'office-mcp-localhost',
      heartbeatIntervalSec: 30,
      heartbeatTimeoutSec: 10,
      sessionGraceSec: 60,
      maxPendingPerSession: 4,
      sharedSecret: ''
    },
    mcp: { host: '127.0.0.1', port: 8800, apiKey: '' },
    limits: { maxResponseBytes: 1024 * 1024, maxRequestBytes: 16 * 1024 * 1024, maxWsFrameBytes: 16 * 1024 * 1024, defaultToolTimeoutMs: 30000, requestsPerMinute: 120 },
    audit: { enabled: false, path: 'audit.jsonl' },
    logging: { level: 'info', file: '' }
  };
}

function listen(server: ReturnType<typeof createServer>): Promise<void> {
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      server.off('error', reject);
      resolve();
    });
  });
}

function close(server: ReturnType<typeof createServer>): Promise<void> {
  return new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
}

async function waitFor(predicate: () => boolean): Promise<void> {
  const deadline = Date.now() + 1000;
  while (Date.now() < deadline) {
    if (predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  assert.equal(predicate(), true);
}

function isMethod(message: unknown, method: string): message is { method: string; params?: unknown } {
  return typeof message === 'object' && message !== null && 'method' in message && (message as { method?: unknown }).method === method;
}
