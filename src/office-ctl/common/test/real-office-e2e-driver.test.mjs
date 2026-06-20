import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const DRIVER = fileURLToPath(new URL('./real-office-e2e-driver.mjs', import.meta.url));

test('real Office E2E driver creates and cleans up Word documents through COM', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-real-driver-word-'));
  const create = runDriver({ host: 'Word', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-word-.*\.docx$/i);
  assert.equal(document.host, 'Word');

  const cleanup = runDriver({ host: 'Word', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  assert.equal(JSON.parse(cleanup.stdout).deleted, true);
});

test('real Office E2E driver creates and cleans up Excel workbooks through COM', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-real-driver-excel-'));
  const create = runDriver({ host: 'Excel', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-excel-.*\.xlsx$/i);
  assert.equal(document.host, 'Excel');

  const cleanup = runDriver({ host: 'Excel', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  assert.equal(JSON.parse(cleanup.stdout).deleted, true);
});

test('real Office E2E driver callTool posts MCP requests through an injectable endpoint', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-real-driver-mcp-'));
  const logPath = join(dir, 'mcp-requests.jsonl');
  const serverPath = join(dir, 'mcp-server.mjs');
  writeFileSync(serverPath, `
import { appendFileSync } from 'node:fs';
import { createServer } from 'node:http';
const logPath = ${JSON.stringify(logPath)};
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: JSON.parse(body) }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 1, result: { protocolVersion: '2024-11-05' } }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 2, result: { structuredContent: { ok: true } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  try {
    const { endpoint } = JSON.parse(await firstStdoutLine(server));
    const result = runDriver({
      host: 'Word',
      step: 'callTool',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: { call: { name: 'word.get_text', arguments: { limit: 1 } } }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).structuredContent.ok, true);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[0].body.method, 'initialize');
    assert.equal(requests[1].session, 'mcp-session-test');
    assert.equal(requests[1].body.params.name, 'word.get_text');
    assert.equal(requests[1].body.params.arguments.session_id, 'session-1');
  } finally {
    server.kill();
  }
});

test('real Office E2E driver rejects add-in activation until a concrete activator is provided', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-real-driver-empty-session-'));
  const serverPath = join(dir, 'empty-session-server.mjs');
  writeFileSync(serverPath, `
import { createServer } from 'node:http';
const server = createServer((request, response) => {
  let body = '';
  request.setEncoding('utf8');
  request.on('data', (chunk) => { body += chunk; });
  request.on('end', () => {
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 1, result: {} }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: 2, result: { structuredContent: { sessions: [] } } }));
    }
  });
});
server.listen(0, '127.0.0.1', () => {
  const address = server.address();
  console.log(JSON.stringify({ endpoint: 'http://127.0.0.1:' + address.port + '/mcp' }));
});
`);
  const server = spawn(process.execPath, [serverPath], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  return firstStdoutLine(server).then((line) => {
    try {
      const { endpoint } = JSON.parse(line);
      const result = runDriver({ host: 'Word', step: 'waitForSession', context: { daemon: { endpoint }, document: { path: 'missing.docx' }, timeoutMs: 1 } });
      assert.notEqual(result.status, 0);
      assert.match(result.stderr, /Open MCP Control in Word/);
    } finally {
      server.kill();
    }
  });
});

function runDriver(payload) {
  return spawnSync(process.execPath, [DRIVER], {
    input: JSON.stringify(payload),
    encoding: 'utf8',
    timeout: 30000
  });
}

function firstStdoutLine(child) {
  return new Promise((resolve, reject) => {
    let buffer = '';
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => {
      buffer += chunk;
      const index = buffer.indexOf('\n');
      if (index !== -1) resolve(buffer.slice(0, index));
    });
    child.stderr.on('data', (chunk) => {
      if (String(chunk).trim()) reject(new Error(String(chunk)));
    });
    child.on('error', reject);
    child.on('exit', (code) => {
      if (code !== null && code !== 0) reject(new Error(`server exited ${code}`));
    });
  });
}
