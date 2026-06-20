import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const DRIVER = fileURLToPath(new URL('./office-e2e-driver.mjs', import.meta.url));
const REPO_ROOT = resolve(dirname(DRIVER), '../../../..');
const DEFAULT_ACTIVATOR = resolve(REPO_ROOT, 'src/office-ctl/common/scripts/activate-office-mcp-addin.ps1');
const RUN_OFFICE_COM = process.env.OFFICE_MCP_RUN_E2E === '1';

test('Office E2E driver describes a driver-owned Word lifecycle', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-word-'));
  const create = runDriver({ host: 'Word', step: 'describeDocumentLifecycle', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-word-.*\.docx$/i);
  assert.equal(document.host, 'Word');
  assert.equal(document.createdByDriver, true);
  assert.equal(document.officeWindowMode, 'visible');
  assert.ok(document.keeper?.closePath, 'driver-owned close sentinel is required');
  assert.ok(document.keeper?.startedPath, 'keeper started sentinel is required');
  assert.ok(document.keeper?.pidPath, 'keeper pid file is required');
  assert.ok(document.keeper?.stdoutPath, 'keeper stdout log is required');
  assert.ok(document.keeper?.stderrPath, 'keeper stderr log is required');
  assert.match(document.script, /Documents\.Add\(\)/);
  assert.match(document.script, /Invoke-Retry/);
  assert.match(document.script, /RPC_E_CALL_REJECTED/);
  assert.doesNotMatch(document.script, /\.Content\.Text=/);
  assert.doesNotMatch(document.script, /\.Content\.InsertAfter/);
  assert.doesNotMatch(document.script, /TrackRevisions/);
  assert.match(document.script, /office-mcp-ready/);
  assert.match(document.script, /office-mcp-ready/);
  assert.doesNotMatch(document.script, /\.Quit\(\)/, 'keeper must not quit user Office applications');
});

test('Office E2E driver creates and cleans up Word documents through COM', { skip: !RUN_OFFICE_COM }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-word-com-'));
  const create = runDriver({ host: 'Word', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  const cleanup = runDriver({ host: 'Word', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  assert.equal(JSON.parse(cleanup.stdout).deleted, true);
});

test('Office E2E driver describes a driver-owned Excel lifecycle', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-excel-'));
  const create = runDriver({ host: 'Excel', step: 'describeDocumentLifecycle', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-excel-.*\.xlsx$/i);
  assert.equal(document.host, 'Excel');
  assert.equal(document.createdByDriver, true);
  assert.equal(document.officeWindowMode, 'visible');
  assert.ok(document.keeper?.closePath, 'driver-owned close sentinel is required');
  assert.ok(document.keeper?.startedPath, 'keeper started sentinel is required');
  assert.ok(document.keeper?.pidPath, 'keeper pid file is required');
  assert.ok(document.keeper?.stdoutPath, 'keeper stdout log is required');
  assert.ok(document.keeper?.stderrPath, 'keeper stderr log is required');
  assert.match(document.script, /Workbooks\.Add\(\)/);
  assert.match(document.script, /office-mcp-ready/);
  assert.match(document.script, /office-mcp-ready/);
  assert.doesNotMatch(document.script, /\.Quit\(\)/, 'keeper must not quit user Office applications');
});

test('Office E2E driver creates and cleans up Excel workbooks through COM', { skip: !RUN_OFFICE_COM }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-excel-com-'));
  const create = runDriver({ host: 'Excel', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  const cleanup = runDriver({ host: 'Excel', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  assert.equal(JSON.parse(cleanup.stdout).deleted, true);
});

test('Office E2E driver describes a visible PowerPoint lifecycle', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-powerpoint-'));
  const create = runDriver({ host: 'PowerPoint', step: 'describeDocumentLifecycle', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  assert.match(document.path, /office-mcp-e2e-powerpoint-.*\.pptx$/i);
  assert.equal(document.host, 'PowerPoint');
  assert.equal(document.createdByDriver, true);
  assert.equal(document.officeWindowMode, 'visible');
  assert.ok(document.keeper?.closePath, 'driver-owned close sentinel is required');
  assert.ok(document.keeper?.startedPath, 'keeper started sentinel is required');
  assert.ok(document.keeper?.pidPath, 'keeper pid file is required');
  assert.ok(document.keeper?.stdoutPath, 'keeper stdout log is required');
  assert.ok(document.keeper?.stderrPath, 'keeper stderr log is required');
  assert.match(document.script, /Presentations\.Add\(\$true\)/);
  assert.match(document.script, /office-mcp-ready/);
  assert.match(document.script, /office-mcp-ready/);
  assert.doesNotMatch(document.script, /\.Quit\(\)/, 'keeper must not quit user Office applications');
});

test('Office E2E driver reuses the built daemon binary for status when available', () => {
  const result = runDriver({ host: 'Word', step: 'describeDaemonStatusCommand' });
  assert.equal(result.status, 0, result.stderr);
  const command = JSON.parse(result.stdout);
  if (existsSync(resolve(REPO_ROOT, 'target/debug/office-mcp-daemon.exe'))) {
    assert.match(command.command, /office-mcp-daemon\.exe$/);
    assert.deepEqual(command.args, ['daemon', 'status']);
  } else {
    assert.equal(command.command, 'cargo');
    assert.deepEqual(command.args, ['run', '-q', '-p', 'office-mcp-daemon', '--', 'daemon', 'status']);
  }
});

test('Office E2E driver cleanup canonicalizes Office document paths', () => {
  const result = runDriver({ host: 'Word', step: 'describeDocumentLifecycle', context: { workDir: mkdtempSync(join(tmpdir(), 'office-mcp-driver-cleanup-script-')) } });
  assert.equal(result.status, 0, result.stderr);
  const document = JSON.parse(result.stdout);
  assert.match(document.cleanupScript, /function Canonical/);
  assert.match(document.cleanupScript, /Canonical \$doc\.FullName/);
});

test('Office E2E driver uses a visible PowerPoint window and safe cleanup', { skip: !RUN_OFFICE_COM }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-powerpoint-com-'));
  const create = runDriver({ host: 'PowerPoint', step: 'createDocument', context: { workDir: dir } });
  assert.equal(create.status, 0, create.stderr);
  const document = JSON.parse(create.stdout);
  const cleanup = runDriver({ host: 'PowerPoint', step: 'cleanupDocument', context: { document } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  const result = JSON.parse(cleanup.stdout);
  assert.equal(result.closedByDriver, true);
  assert.equal(result.deleted, true);
});

test('Office E2E driver activation step is explicit and configurable', () => {
  const previousDefault = process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR;
  process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR = '0';
  try {
    const skipped = runDriver({ host: 'Word', step: 'activateAddin', context: { document: { path: 'fixture.docx' }, daemon: { addinEndpoint: 'wss://localhost:8765/addin' } } });
    assert.equal(skipped.status, 0, skipped.stderr);
    assert.equal(JSON.parse(skipped.stdout).skipped, 'no-activator-configured');
  } finally {
    restoreEnv('OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR', previousDefault);
  }

  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-activator-'));
  const logPath = join(dir, 'activator-env.json');
  const activatorPath = join(dir, 'activator.mjs');
  writeFileSync(activatorPath, `
import { writeFileSync } from 'node:fs';
writeFileSync(${JSON.stringify(logPath)}, JSON.stringify({
  host: process.env.OFFICE_MCP_E2E_HOST,
  documentPath: process.env.OFFICE_MCP_E2E_DOCUMENT_PATH,
  addinOrigin: process.env.OFFICE_MCP_E2E_ADDIN_ORIGIN,
  addinEndpoint: process.env.OFFICE_MCP_E2E_ADDIN_ENDPOINT
}));
`);
  const previous = process.env.OFFICE_MCP_E2E_ACTIVATOR;
  process.env.OFFICE_MCP_E2E_ACTIVATOR = `${process.execPath} ${activatorPath}`;
  try {
    const activated = runDriver({
      host: 'PowerPoint',
      step: 'activateAddin',
      context: {
        document: { path: 'deck.pptx' },
        daemon: { addinOrigin: 'https://localhost:8765', addinEndpoint: 'wss://localhost:8765/addin' },
        timeoutMs: 5000
      }
    });
    assert.equal(activated.status, 0, activated.stderr);
    assert.equal(JSON.parse(activated.stdout).activated, true);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_ACTIVATOR', previous);
  }

  const env = JSON.parse(readFileSync(logPath, 'utf8'));
  assert.equal(env.host, 'powerpoint');
  assert.equal(env.documentPath, 'deck.pptx');
  assert.equal(env.addinOrigin, 'https://localhost:8765');
  assert.equal(env.addinEndpoint, 'wss://localhost:8765/addin');
});

test('Office E2E driver provides a default Windows add-in activator', () => {
  const previousActivator = process.env.OFFICE_MCP_E2E_ACTIVATOR;
  const previousDryRun = process.env.OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN;
  const previousDefault = process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR;
  delete process.env.OFFICE_MCP_E2E_ACTIVATOR;
  process.env.OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN = '1';
  delete process.env.OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR;
  try {
    const activated = runDriver({
      host: 'Excel',
      step: 'activateAddin',
      context: {
        document: { path: 'book.xlsx' },
        daemon: { addinOrigin: 'https://localhost:8765', addinEndpoint: 'wss://localhost:8765/addin' },
        timeoutMs: 5000
      }
    });
    assert.equal(activated.status, 0, activated.stderr);
    const result = JSON.parse(activated.stdout);
    assert.equal(result.activated, true);
    assert.equal(result.activator_kind, 'default-windows-taskpane');
    assert.match(result.activator, /activate-office-mcp-addin\.ps1/);
  } finally {
    restoreEnv('OFFICE_MCP_E2E_ACTIVATOR', previousActivator);
    restoreEnv('OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN', previousDryRun);
    restoreEnv('OFFICE_MCP_E2E_USE_DEFAULT_ACTIVATOR', previousDefault);
  }
});

test('default Windows add-in activator can fall back through My Add-ins catalog UI', () => {
  const script = readFileSync(DEFAULT_ACTIVATOR, 'utf8');
  assert.match(script, /My Add-ins/);
  assert.match(script, /Office MCP Control/);
  assert.match(script, /Shared Folder/);
  assert.match(script, /Add/);
});

test('Office E2E driver callTool posts MCP requests through an injectable endpoint', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-mcp-'));
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

test('Office E2E driver listTools reads daemon MCP tools/list names', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-list-tools-'));
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
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { tools: [{ name: 'office.list_sessions' }, { name: 'word.get_text' }, { name: 'excel.read_range' }] } }));
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
    const result = runDriver({ host: 'Word', step: 'listTools', context: { daemon: { endpoint } } });
    assert.equal(result.status, 0, result.stderr);
    assert.deepEqual(JSON.parse(result.stdout), ['office.list_sessions', 'word.get_text', 'excel.read_range']);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.ok(requests.some((entry) => entry.body.method === 'tools/list'));
  } finally {
    server.kill();
  }
});

test('Office E2E driver verifies readback expectations through an MCP read tool', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-readback-'));
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
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
    } else {
      const text = ${JSON.stringify('updated marker remains\nfirst marker\nsecond marker')};
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { text, slides: [{ slide_index: 0, layout_name: 'Blank' }] } } }));
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
      step: 'verifyResult',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        result: { structuredContent: { shape_id: 'shape-999' } },
        toolCase: {
          tool: 'powerpoint.add_table',
          verify: {
            kind: 'readback',
            readbackTool: 'powerpoint.read_table',
            readbackArguments: { slide_index: 0, shape_id: '${result.shape_id}' },
            expect: {
              contains: ['updated marker'],
              notContains: ['baseline marker'],
              orderedContains: ['first marker', 'second marker'],
              pathEquals: [{ path: 'slides.0.layout_name', value: 'Blank' }],
              pathMissing: ['slides.1']
            }
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).verified, true);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[1].body.params.name, 'powerpoint.read_table');
    assert.equal(requests[1].body.params.arguments.session_id, 'session-1');
    assert.equal(requests[1].body.params.arguments.shape_id, 'shape-999');
  } finally {
    server.kill();
  }
});

test('Office E2E driver verifies readback expectations through an MCP resource', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-resource-readback-'));
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
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    if (parsed.method === 'resources/read') {
      const text = JSON.stringify({ comments: [{ comment_id: 'comment-1', content: 'Resolve me E2E', resolved: true }] });
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { contents: [{ uri: parsed.params.uri, mimeType: 'application/json', text }] } }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { comment_id: 'comment-1' } } }));
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
      step: 'verifyResult',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        result: { structuredContent: { comment_id: 'comment-1' } },
        toolCase: {
          tool: 'word.resolve_comment',
          verify: {
            kind: 'readback',
            resource: 'office://word/${session_id}/comments',
            expect: {
              contains: ['Resolve me E2E'],
              pathEquals: [{ path: 'comments.0.resolved', value: true }]
            }
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).readbackResource, 'office://word/session-1/comments');
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    const resourceRead = requests.find((entry) => entry.body.method === 'resources/read');
    assert.equal(resourceRead.body.params.uri, 'office://word/session-1/comments');
  } finally {
    server.kill();
  }
});

test('Office E2E driver verifies direct-result expectations', async () => {
  const result = runDriver({
    host: 'PowerPoint',
    step: 'verifyResult',
    context: {
      toolCase: {
        tool: 'powerpoint.insert_image',
        verify: {
          kind: 'direct-result',
          expect: {
            pathEquals: [
              { path: 'inserted_image', value: true },
              { path: 'mime_type', value: 'image/png' }
            ]
          }
        }
      },
      result: {
        structuredContent: {
          inserted_image: true,
          mime_type: 'image/png'
        }
      }
    }
  });

  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(result.stdout).verified, true);
});

test('Office E2E driver fails direct-result expectation mismatches', async () => {
  const result = runDriver({
    host: 'PowerPoint',
    step: 'verifyResult',
    context: {
      toolCase: {
        tool: 'powerpoint.insert_image',
        verify: {
          kind: 'direct-result',
          expect: { pathEquals: [{ path: 'inserted_image', value: true }] }
        }
      },
      result: { structuredContent: { inserted_image: false } }
    }
  });

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /inserted_image/);
});

test('Office E2E driver accepts declared host capability errors', async () => {
  const result = runDriver({
    host: 'PowerPoint',
    step: 'verifyResult',
    context: {
      toolCase: {
        tool: 'powerpoint.export_file',
        verify: {
          kind: 'direct-result',
          allowErrorCodes: ['HOST_CAPABILITY_UNAVAILABLE'],
          expect: { pathEquals: [{ path: 'format', value: 'pdf' }] }
        }
      },
      result: {
        structuredContent: {
          error: {
            office_mcp_code: 'HOST_CAPABILITY_UNAVAILABLE',
            message: 'PowerPoint desktop file export is not available through Office.context.document.getFileAsync in this host.',
            retriable: true,
            partial_effect: 'none'
          }
        }
      }
    }
  });

  assert.equal(result.status, 0, result.stderr);
  const verified = JSON.parse(result.stdout);
  assert.equal(verified.verified, true);
  assert.equal(verified.acceptedErrorCode, 'HOST_CAPABILITY_UNAVAILABLE');
});

test('Office E2E driver setupContent runs declared MCP setup actions', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-setup-'));
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
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { ok: true } } }));
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
      step: 'setupContent',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: {
          tool: 'word.replace_text',
          setup: {
            actions: [
              { tool: 'word.replace_text', arguments: { find: 'office-mcp e2e baseline', replace: 'baseline marker' } }
            ]
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).actions, 1);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[1].body.params.name, 'word.replace_text');
    assert.equal(requests[1].body.params.arguments.session_id, 'session-1');
    assert.equal(requests[1].body.params.arguments.replace, 'baseline marker');
  } finally {
    server.kill();
  }
});

test('Office E2E driver resolves setup action result references in later calls', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-bindings-'));
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
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
    } else if (parsed.params?.name === 'powerpoint.add_table') {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { shape_id: 'shape-123' } } }));
    } else if (parsed.params?.name === 'word.insert_content_control') {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { content_control: { content_control_id: 42 } } } }));
    } else {
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { text: 'Updated table cell' } } }));
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
    const setup = {
      actions: [
        { tool: 'powerpoint.add_table', saveAs: 'table', arguments: { slide_index: 0, values: [['Old']] } }
      ]
    };
    const setupResult = runDriver({
      host: 'PowerPoint',
      step: 'setupContent',
      context: { daemon: { endpoint }, session: { sessionId: 'session-1' }, toolCase: { setup } }
    });
    assert.equal(setupResult.status, 0, setupResult.stderr);
    const bindings = JSON.parse(setupResult.stdout).bindings;
    assert.equal(bindings.table.shape_id, 'shape-123');

    const callResult = runDriver({
      host: 'PowerPoint',
      step: 'callTool',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1', bindings },
        toolCase: {
          call: {
            name: 'powerpoint.update_table',
            arguments: { slide_index: 0, shape_id: '${table.shape_id}', action: 'set_cell', row_index: 0, column_index: 0, value: 'Updated table cell' }
          }
        }
      }
    });
    assert.equal(callResult.status, 0, callResult.stderr);

    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(requests[3].body.params.name, 'powerpoint.update_table');
    assert.equal(requests[3].body.params.arguments.shape_id, 'shape-123');

    const nestedSetupResult = runDriver({
      host: 'Word',
      step: 'setupContent',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: {
          setup: {
            actions: [
              { tool: 'word.insert_content_control', saveAs: 'controlResult', arguments: { tag: 'e2e' } }
            ]
          }
        }
      }
    });
    assert.equal(nestedSetupResult.status, 0, nestedSetupResult.stderr);
    const nestedBindings = JSON.parse(nestedSetupResult.stdout).bindings;
    assert.equal(nestedBindings.controlResult.content_control.content_control_id, 42);

    const nestedCallResult = runDriver({
      host: 'Word',
      step: 'callTool',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1', bindings: nestedBindings },
        toolCase: {
          call: {
            name: 'word.update_content_control',
            arguments: { content_control_id: '${controlResult.content_control.content_control_id}', text: 'Updated control' }
          }
        }
      }
    });
    assert.equal(nestedCallResult.status, 0, nestedCallResult.stderr);

    const updatedRequests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    assert.equal(updatedRequests[7].body.params.name, 'word.update_content_control');
    assert.equal(updatedRequests[7].body.params.arguments.content_control_id, 42);
  } finally {
    server.kill();
  }
});

test('Office E2E driver binds resource read results for later setup actions', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-resource-bindings-'));
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
    const parsed = JSON.parse(body);
    appendFileSync(logPath, JSON.stringify({ session: request.headers['mcp-session-id'] || null, body: parsed }) + '\\n');
    response.setHeader('Content-Type', 'application/json');
    if (!request.headers['mcp-session-id']) {
      response.setHeader('MCP-Session-Id', 'mcp-session-test');
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: {} }));
      return;
    }
    if (parsed.method === 'resources/read') {
      const text = JSON.stringify({ changes: [{ index: 0, fingerprint: 'fp-123' }] });
      const contents = [{ uri: parsed.params.uri, mimeType: 'application/json', text }];
      response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { contents } }));
      return;
    }
    response.end(JSON.stringify({ jsonrpc: '2.0', id: parsed.id, result: { structuredContent: { ok: true } } }));
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
      step: 'setupContent',
      context: {
        daemon: { endpoint },
        session: { sessionId: 'session-1' },
        toolCase: {
          tool: 'word.update_tracked_change',
          setup: {
            actions: [
              { resource: 'office://word/${session_id}/track_changes', saveAs: 'trackChanges' },
              { tool: 'word.update_tracked_change', arguments: { change_index: 0, action: 'accept', expected_fingerprint: '${trackChanges.changes.0.fingerprint}' } }
            ]
          }
        }
      }
    });
    assert.equal(result.status, 0, result.stderr);
    const requests = readFileSync(logPath, 'utf8').trim().split('\n').map((line) => JSON.parse(line));
    const resourceRead = requests.find((entry) => entry.body.method === 'resources/read');
    const toolCall = requests.find((entry) => entry.body.method === 'tools/call' && entry.body.params.name === 'word.update_tracked_change');
    assert.equal(resourceRead.body.params.uri, 'office://word/session-1/track_changes');
    assert.equal(toolCall.body.params.arguments.expected_fingerprint, 'fp-123');
  } finally {
    server.kill();
  }
});

test('Office E2E driver cleanup ignores documents it did not create', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-cleanup-'));
  const path = join(dir, 'user-owned.docx');
  writeFileSync(path, 'do not delete');
  const cleanup = runDriver({ host: 'Word', step: 'cleanupDocument', context: { document: { path } } });
  assert.equal(cleanup.status, 0, cleanup.stderr);
  const result = JSON.parse(cleanup.stdout);
  assert.equal(result.deleted, false);
  assert.equal(result.skipped, 'not-driver-owned');
  assert.equal(existsSync(path), true);
});

test('Office E2E driver rejects add-in activation until a concrete activator is provided', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-driver-empty-session-'));
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
    timeout: 90000
  });
}

function restoreEnv(name, value) {
  if (value === undefined) delete process.env[name];
  else process.env[name] = value;
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
