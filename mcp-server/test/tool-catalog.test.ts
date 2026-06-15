import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import test from 'node:test';
import { createMcpServer } from '../src/mcp-server.js';
import { SessionRegistry } from '../src/session-registry.js';
import type { DaemonConfig } from '../src/config.js';
import type { AddinConnection } from '../src/types.js';

const WORD_V1_TOOLS = [
  'office.get_session_info',
  'office.list_sessions',
  'word.accept_change',
  'word.add_column',
  'word.add_comment',
  'word.add_row',
  'word.apply_formatting',
  'word.apply_style',
  'word.delete_range',
  'word.find_text',
  'word.format_cell',
  'word.get_outline',
  'word.get_paragraph',
  'word.get_selection',
  'word.get_text',
  'word.insert_heading',
  'word.insert_image',
  'word.insert_list',
  'word.insert_page_break',
  'word.insert_paragraph',
  'word.insert_table',
  'word.read_table',
  'word.reject_change',
  'word.replace_text',
  'word.resolve_comment',
  'word.save',
  'word.set_heading_level',
  'word.update_cell',
  'word.update_paragraph'
].sort();

test('MCP server exposes the complete Word v1 tool catalog', async () => {
  const server = createMcpServer(config(), new SessionRegistry());
  const client = new Client({ name: 'tool-catalog-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  const tools = await client.listTools();
  assert.deepEqual(tools.tools.map((tool) => tool.name).sort(), WORD_V1_TOOLS);

  await client.close();
  await server.close();
});

test('MCP server exposes the Word v1 prompt surface', async () => {
  const server = createMcpServer(config(), new SessionRegistry());
  const client = new Client({ name: 'prompt-catalog-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  const prompts = await client.listPrompts();
  assert.deepEqual(prompts.prompts.map((prompt) => prompt.name).sort(), [
    'extract_action_items',
    'polish_section',
    'summarize_document'
  ]);

  const summary = await client.getPrompt({
    name: 'summarize_document',
    arguments: { session_id: '66666666-6666-4666-8666-666666666666' }
  });
  assert.match(summary.messages[0].content.type === 'text' ? summary.messages[0].content.text : '', /office:\/\/word\/66666666-6666-4666-8666-666666666666\/document/);
  assert.match(summary.messages[0].content.type === 'text' ? summary.messages[0].content.text : '', /word\.add_comment/);

  const polish = await client.getPrompt({
    name: 'polish_section',
    arguments: { session_id: '66666666-6666-4666-8666-666666666666', heading: 'Findings' }
  });
  assert.match(polish.messages[0].content.type === 'text' ? polish.messages[0].content.text : '', /Findings/);
  assert.match(polish.messages[0].content.type === 'text' ? polish.messages[0].content.text : '', /explicit approval/);

  await client.close();
  await server.close();
});


test('MCP schemas accept Word v1 anchor and edit argument variants', async () => {
  const seen: Array<{ tool: string; args: Record<string, unknown> }> = [];
  const registry = new SessionRegistry();
  const connection = fakeConnection(async (_sessionId, tool, args) => {
    seen.push({ tool, args });
    return { ok: true, data: { ok: true } };
  });
  registry.registerRuntime(connection, connection.runtime);
  registry.addSession(connection, {
    session_id: '55555555-5555-4555-8555-555555555555',
    instance_id: connection.runtime.instance_id,
    document: { title: 'Schema.docx' },
    available_tools: ['word.insert_paragraph', 'word.replace_text', 'word.add_comment'],
    is_active: null
  });

  const server = createMcpServer(config(), registry);
  const client = new Client({ name: 'schema-variant-test', version: '0.1.0' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);

  await client.callTool({
    name: 'word.insert_paragraph',
    arguments: {
      session_id: '55555555-5555-4555-8555-555555555555',
      text: 'Formatted',
      anchor: { kind: 'heading', text: 'Intro', level: 1 },
      formatting: { bold: true, color: '#112233' }
    }
  });
  await client.callTool({
    name: 'word.add_comment',
    arguments: {
      session_id: '55555555-5555-4555-8555-555555555555',
      anchor: { kind: 'bookmark', name: 'TargetBookmark' },
      text: 'Bookmark comment'
    }
  });
  await client.callTool({
    name: 'word.replace_text',
    arguments: {
      session_id: '55555555-5555-4555-8555-555555555555',
      find: 'old',
      replace: 'new',
      scope: { paragraph_range: [1, 3], selection_only: false },
      partial_ok: true,
      dry_run: true
    }
  });

  assert.equal(seen.length, 3);
  assert.deepEqual(seen[0].args.anchor, { kind: 'heading', text: 'Intro', level: 1 });
  assert.deepEqual(seen[0].args.formatting, { bold: true, color: '#112233' });
  assert.deepEqual(seen[1].args.anchor, { kind: 'bookmark', name: 'TargetBookmark' });
  assert.deepEqual(seen[2].args.scope, { paragraph_range: [1, 3], selection_only: false });
  assert.equal(seen[2].args.partial_ok, true);

  await client.close();
  await server.close();
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
      instance_id: 'cccccccc-cccc-4ccc-8ccc-cccccccccccc',
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

