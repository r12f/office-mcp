import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import test from 'node:test';
import type { DaemonConfig } from '../src/config.js';
import { createMcpServer } from '../src/mcp-server.js';
import { SessionRegistry } from '../src/session-registry.js';
import type { AddinConnection } from '../src/types.js';

const SESSION_ID = '77777777-7777-4777-8777-777777777777';

test('writes opt-in audit records for successful tool calls without document payload', async () => {
  await withAuditClient(async ({ client, auditPath }) => {
    await client.callTool({ name: 'word.get_text', arguments: { session_id: SESSION_ID, offset: 0, limit: 10 } });

    const records = readAuditRecords(auditPath);
    assert.equal(records.length, 1);
    assert.equal(records[0].tool, 'word.get_text');
    assert.equal(records[0].session_id, SESSION_ID);
    assert.equal(records[0].ok, true);
    assert.equal(typeof records[0].duration_ms, 'number');
    assert.equal('text' in records[0], false);
  });
});

test('writes opt-in audit records for failed tool calls', async () => {
  await withAuditClient(async ({ client, auditPath }) => {
    await client.callTool({ name: 'word.insert_paragraph', arguments: { session_id: SESSION_ID, text: 'secret body', anchor: { kind: 'end_of_document' } } });

    const records = readAuditRecords(auditPath);
    assert.equal(records.length, 1);
    assert.equal(records[0].tool, 'word.insert_paragraph');
    assert.equal(records[0].session_id, SESSION_ID);
    assert.equal(records[0].ok, false);
    assert.equal(records[0].error_code, 'HOST_CAPABILITY_UNAVAILABLE');
    assert.equal(JSON.stringify(records).includes('secret body'), false);
  });
});

async function withAuditClient(callback: (context: { client: Client; auditPath: string }) => Promise<void>): Promise<void> {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-audit-'));
  const auditPath = join(dir, 'audit.jsonl');
  const registry = new SessionRegistry();
  const connection = fakeConnection(async (_sessionId, tool) => ({ ok: true, data: { tool, text: 'document body' } }));
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: SESSION_ID,
    instance_id: connection.runtime.instance_id,
    document: { title: 'Audit.docx' },
    available_tools: ['word.get_text'],
    is_active: null
  });

  const server = createMcpServer(config(auditPath), registry);
  const client = new Client({ name: 'audit-log-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  try {
    await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
    await callback({ client, auditPath });
  } finally {
    await client.close();
    await server.close();
    rmSync(dir, { recursive: true, force: true });
  }
}

function readAuditRecords(path: string): Array<Record<string, unknown>> {
  return readFileSync(path, 'utf8').trim().split('\n').map((line) => JSON.parse(line) as Record<string, unknown>);
}

function fakeConnection(invokeTool: AddinConnection['invokeTool']): AddinConnection {
  const socket = new EventEmitter() as AddinConnection['socket'];
  Object.assign(socket, { OPEN: 1, readyState: 1, close() {} });
  return {
    socket,
    runtime: {
      instance_id: 'dddddddd-dddd-4ddd-8ddd-dddddddddddd',
      host: { app: 'word' },
      add_in: { version: '0.1.0', protocol_version: '1.0' },
      registered_at: new Date().toISOString()
    },
    pending: new Map(),
    queue: Promise.resolve(),
    invokeTool
  };
}

function config(auditPath: string): DaemonConfig {
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
    audit: { enabled: true, path: auditPath },
    logging: { level: 'info', file: '' }
  };
}
