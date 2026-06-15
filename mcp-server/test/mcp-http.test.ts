import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { McpFrontend } from '../src/mcp-server.js';
import { SessionRegistry } from '../src/session-registry.js';
import type { DaemonConfig } from '../src/config.js';

test('rejects MCP HTTP requests from foreign browser origins', async () => {
  const response = fakeResponse();
  await new McpFrontend(config(), new SessionRegistry()).handle(fakeRequest({ origin: 'https://evil.example' }), response);

  assert.equal(response.statusCode, 403);
  assert.equal(response.body, 'Forbidden origin');
});

test('requires bearer auth when MCP API key is configured', async () => {
  const testConfig = config();
  testConfig.mcp.apiKey = 'secret';
  const response = fakeResponse();
  await new McpFrontend(testConfig, new SessionRegistry()).handle(fakeRequest(), response);

  assert.equal(response.statusCode, 401);
  assert.equal(response.headers['WWW-Authenticate'], 'Bearer');
});

test('rejects non-bearer MCP auth when API key is configured', async () => {
  const testConfig = config();
  testConfig.mcp.apiKey = 'secret';
  const response = fakeResponse();
  await new McpFrontend(testConfig, new SessionRegistry()).handle(fakeRequest({ authorization: 'Basic secret' }), response);

  assert.equal(response.statusCode, 401);
});

test('rejects wrong-length MCP bearer auth without throwing', async () => {
  const testConfig = config();
  testConfig.mcp.apiKey = 'secret';
  const response = fakeResponse();
  await new McpFrontend(testConfig, new SessionRegistry()).handle(fakeRequest({ authorization: 'Bearer s' }), response);

  assert.equal(response.statusCode, 401);
});

test('accepts configured MCP bearer auth before dispatch', async () => {
  const testConfig = config();
  testConfig.mcp.apiKey = 'secret';
  const response = fakeResponse();
  await new McpFrontend(testConfig, new SessionRegistry()).handle(fakeRequest({ authorization: 'Bearer secret' }), response);

  assert.notEqual(response.statusCode, 401);
});

test('rate limits MCP HTTP requests per source', async () => {
  const testConfig = config();
  testConfig.limits.requestsPerMinute = 1;
  const frontend = new McpFrontend(testConfig, new SessionRegistry());

  const first = fakeResponse();
  await frontend.handle(fakeRequest(), first);
  const second = fakeResponse();
  await frontend.handle(fakeRequest(), second);

  assert.equal(second.statusCode, 429);
  assert.equal(second.headers['Retry-After'], '60');
});

test('MCP frontend wires HTTP session lifecycle into UI client tracking', () => {
  const source = readFileSync('src/mcp-server.ts', 'utf8');

  assert.match(source, /uiState\?\.registerClient/);
  assert.match(source, /uiState\?\.unregisterClient/);
  assert.match(source, /uiState\?\.touchClient/);
});

function fakeRequest(headers: Record<string, string> = {}): IncomingMessage {
  return Object.assign(new EventEmitter(), {
    method: 'GET',
    url: '/mcp',
    headers,
    socket: { remoteAddress: '127.0.0.1' }
  }) as IncomingMessage;
}

type CapturedResponse = ServerResponse & { statusCode?: number; body: string; headers: Record<string, string> };

function fakeResponse(): CapturedResponse {
  const response = new EventEmitter() as unknown as CapturedResponse;
  response.body = '';
  response.headers = {};
  response.writeHead = ((statusCode: number, headers?: Record<string, string>) => {
    response.statusCode = statusCode;
    response.headers = headers ?? {};
    return response;
  }) as unknown as CapturedResponse['writeHead'];
  response.end = ((chunk?: string) => {
    if (chunk) response.body += chunk;
    return response;
  }) as unknown as CapturedResponse['end'];
  return response;
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
