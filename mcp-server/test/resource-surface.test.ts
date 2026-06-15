import assert from 'node:assert/strict';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import test from 'node:test';
import { createMcpServer } from '../src/mcp-server.js';
import { SessionRegistry } from '../src/session-registry.js';
import type { DaemonConfig } from '../src/config.js';
import type { AddinConnection } from '../src/types.js';
import { EventEmitter } from 'node:events';

const SESSION_ID = '44444444-4444-4444-8444-444444444444';

test('MCP server exposes and routes the Word v1 resource surface', async () => {
  const invoked: Array<{ tool: string; args: Record<string, unknown> }> = [];
  const registry = new SessionRegistry();
  const connection = fakeConnection(async (_sessionId, tool, args) => {
    invoked.push({ tool, args });
    return { ok: true, data: dataFor(tool) };
  });
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: SESSION_ID,
    instance_id: connection.runtime.instance_id,
    document: { title: 'Resources.docx' },
    available_tools: ['word.get_text', 'word.get_outline', 'word.get_paragraph', 'word.get_selection'],
    is_active: null
  });

  const server = createMcpServer(config(), registry);
  const client = new Client({ name: 'resource-surface-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  const resources = await client.listResources();
  assert.deepEqual(resources.resources.map((resource) => resource.uri).sort(), [
    'office://sessions',
    `office://word/${SESSION_ID}/comments`,
    `office://word/${SESSION_ID}/selection`,
    `office://word/${SESSION_ID}/structure`,
    `office://word/${SESSION_ID}/track_changes`
  ].sort());

  const templates = await client.listResourceTemplates();
  assert.deepEqual(templates.resourceTemplates.map((template) => template.uriTemplate).sort(), [
    'office://word/{session_id}/comments',
    'office://word/{session_id}/document{?offset,limit}',
    'office://word/{session_id}/paragraph/{index}',
    'office://word/{session_id}/selection',
    'office://word/{session_id}/structure',
    'office://word/{session_id}/track_changes'
  ]);

  await readJson(client, `office://word/${SESSION_ID}/document?offset=2&limit=3`);
  await readJson(client, `office://word/${SESSION_ID}/structure`);
  await readJson(client, `office://word/${SESSION_ID}/paragraph/5`);
  await readJson(client, `office://word/${SESSION_ID}/comments`);
  await readJson(client, `office://word/${SESSION_ID}/track_changes`);
  await readJson(client, `office://word/${SESSION_ID}/selection`);

  assert.deepEqual(invoked, [
    { tool: 'word.get_text', args: { session_id: SESSION_ID, offset: 2, limit: 3 } },
    { tool: 'word._get_structure', args: { session_id: SESSION_ID } },
    { tool: 'word.get_paragraph', args: { session_id: SESSION_ID, index: 5 } },
    { tool: 'word._get_comments', args: { session_id: SESSION_ID } },
    { tool: 'word._get_tracked_changes', args: { session_id: SESSION_ID } },
    { tool: 'word.get_selection', args: { session_id: SESSION_ID } }
  ]);

  await client.close();
  await server.close();
});

async function readJson(client: Client, uri: string): Promise<unknown> {
  const result = await client.readResource({ uri });
  assert.equal(result.contents.length, 1);
  const content = result.contents[0];
  assert.equal(content.mimeType, 'application/json');
  assert.equal(content.uri, uri);
  assert.ok('text' in content);
  return JSON.parse(content.text);
}


test('tracked-change mutation preserves add-in stale-index failures', async () => {
  const registry = new SessionRegistry();
  const connection = fakeConnection(async () => ({
    ok: false,
    error: {
      office_mcp_code: 'STALE_INDEX',
      message: 'Tracked change fingerprint mismatch; re-read track_changes before mutating.',
      retriable: false,
      partial_effect: 'none'
    }
  }));
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: SESSION_ID,
    instance_id: connection.runtime.instance_id,
    document: { title: 'Tracked.docx' },
    available_tools: ['word.accept_change'],
    is_active: null
  });

  const server = createMcpServer(config(), registry);
  const client = new Client({ name: 'tracked-change-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  const result = await client.callTool({
    name: 'word.accept_change',
    arguments: { session_id: SESSION_ID, change_index: 0, expected_fingerprint: 'old' }
  });

  assert.equal(result.isError, true);
  assert.deepEqual(result.structuredContent, {
    ok: false,
    error: {
      office_mcp_code: 'STALE_INDEX',
      message: 'Tracked change fingerprint mismatch; re-read track_changes before mutating.',
      retriable: false,
      partial_effect: 'none'
    }
  });

  await client.close();
  await server.close();
});

test('tool responses exceeding configured size return MAX_RESPONSE_SIZE', async () => {
  const testConfig = config();
  testConfig.limits.maxResponseBytes = 20;
  const registry = new SessionRegistry();
  const connection = fakeConnection(async () => ({ ok: true, data: { text: 'this response is intentionally too large' } }));
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: SESSION_ID,
    instance_id: connection.runtime.instance_id,
    document: { title: 'Large.docx' },
    available_tools: ['word.get_text'],
    is_active: null
  });

  const server = createMcpServer(testConfig, registry);
  const client = new Client({ name: 'max-response-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  const result = await client.callTool({
    name: 'word.get_text',
    arguments: { session_id: SESSION_ID, offset: 0, limit: 1 }
  });

  assert.equal(result.isError, true);
  assert.deepEqual((result.structuredContent as { error: { office_mcp_code: string; max_response_bytes: number } }).error, {
    office_mcp_code: 'MAX_RESPONSE_SIZE',
    message: 'Tool response exceeds 20 bytes.',
    retriable: false,
    max_response_bytes: 20
  });

  await client.close();
  await server.close();
});

test('resource responses exceeding configured size throw MAX_RESPONSE_SIZE', async () => {
  const testConfig = config();
  testConfig.limits.maxResponseBytes = 20;
  const registry = new SessionRegistry();
  const connection = fakeConnection(async () => ({ ok: true, data: { text: 'this resource is intentionally too large' } }));
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: SESSION_ID,
    instance_id: connection.runtime.instance_id,
    document: { title: 'LargeResource.docx' },
    available_tools: ['word.get_text'],
    is_active: null
  });

  const server = createMcpServer(testConfig, registry);
  const client = new Client({ name: 'max-resource-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  await assert.rejects(
    client.readResource({ uri: `office://word/${SESSION_ID}/document?offset=0&limit=1` }),
    /MAX_RESPONSE_SIZE|exceeds/
  );

  await client.close();
  await server.close();
});
function dataFor(tool: string): unknown {
  switch (tool) {
    case 'word.get_text':
      return { text: 'body' };
    case 'word._get_structure':
      return { outline: [], headings: [], lists: [], tables: [] };
    case 'word.get_paragraph':
      return { index: 5, text: 'paragraph' };
    case 'word._get_comments':
      return { comments: [] };
    case 'word._get_tracked_changes':
      return { changes: [] };
    case 'word.get_selection':
      return { text: '', is_empty: true };
    default:
      throw new Error(`Unexpected tool ${tool}`);
  }
}

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
      instance_id: 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb',
      host: { app: 'word' },
      add_in: { version: '0.1.0', protocol_version: '1.0' },
      registered_at: new Date().toISOString()
    },
    pending: new Map(),
    queue: Promise.resolve(),
    invokeTool
  };
}

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



