import assert from 'node:assert/strict';
import test from 'node:test';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { createServer } from 'node:http';
import { request as httpsRequest } from 'node:https';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { readJsonWithLimit, resolveAddinPublicDir, startDaemon } from '../src/daemon.js';
import type { DaemonConfig } from '../src/config.js';

test('parses JSON request bodies within the configured size limit', async () => {
  const body = await readJsonWithLimit(chunks(['{"ok":true}']), 20);

  assert.deepEqual(body, { ok: true });
});

test('rejects JSON request bodies exceeding the configured size limit', async () => {
  await assert.rejects(
    readJsonWithLimit(chunks(['{"text":"too large"}']), 10),
    /exceeds 10 bytes/
  );
});

test('resolves add-in public assets beside the server package in source checkout', () => {
  const moduleUrl = pathToFileURL(join('C:\\Code\\office-mcp', 'mcp-server', 'src', 'daemon.ts')).href;

  assert.equal(
    resolveAddinPublicDir(moduleUrl, ''),
    join('C:\\Code\\office-mcp', 'addin', 'public')
  );
});

test('resolves add-in public assets from the MSI install root at runtime', () => {
  const moduleUrl = pathToFileURL(join('C:\\Users\\User\\AppData\\Local\\office-mcp', 'mcp-server', 'dist', 'src', 'daemon.js')).href;

  assert.equal(
    resolveAddinPublicDir(moduleUrl, 'C:\\Users\\User\\AppData\\Local\\office-mcp'),
    join('C:\\Users\\User\\AppData\\Local\\office-mcp', 'addin', 'public')
  );
});

test('daemon UI state endpoint requires UI token and rejects foreign origins', async () => {
  const config = await daemonConfig();
  const runtimeDir = mkdtempSync(join(tmpdir(), 'office-mcp-daemon-ui-'));
  const runtimePath = join(runtimeDir, 'ui-runtime.json');
  const previousRuntimePath = process.env.OFFICE_MCP_UI_RUNTIME_PATH;
  process.env.OFFICE_MCP_UI_RUNTIME_PATH = runtimePath;
  const daemon = await startDaemon(config);
  try {
    const runtime = JSON.parse(readFileSync(runtimePath, 'utf8')) as { token?: string; stateUrl?: string; uiUrl?: string };
    assert.equal(runtime.token, daemon.uiToken);
    assert.equal(runtime.stateUrl, `${config.addin.origin}/ui/state`);
    assert.equal(runtime.uiUrl, `${config.addin.origin}/ui/`);

    const endpoint = `${config.addin.origin}/ui/state`;
    const shell = await httpsJson(`${config.addin.origin}/ui/`);
    assert.equal(shell.status, 200);
    assert.match(shell.body, /Office MCP/);
    assert.match(shell.body, /__OFFICE_MCP_UI__/);

    const script = await httpsJson(`${config.addin.origin}/ui/app.js`);
    assert.equal(script.status, 200);
    assert.match(script.body, /fetch\(boot\.stateUrl/);

    const css = await httpsJson(`${config.addin.origin}/ui/app.css`);
    assert.equal(css.status, 200);
    assert.match(css.body, /prefers-reduced-motion/);

    const unauthorized = await httpsJson(endpoint);
    assert.equal(unauthorized.status, 401);

    const foreign = await httpsJson(endpoint, { origin: 'https://evil.example', 'x-office-mcp-ui-token': daemon.uiToken });
    assert.equal(foreign.status, 403);

    const ok = await httpsJson(endpoint, { origin: config.addin.origin, 'x-office-mcp-ui-token': daemon.uiToken });
    assert.equal(ok.status, 200);
    const body = JSON.parse(ok.body) as { daemon?: { status?: string }; documents?: Record<string, unknown[]> };
    assert.equal(body.daemon?.status, 'up');
    assert.deepEqual(body.documents?.word, []);
  } finally {
    await daemon.close();
    assert.equal(existsSync(runtimePath), false);
    if (previousRuntimePath === undefined) delete process.env.OFFICE_MCP_UI_RUNTIME_PATH;
    else process.env.OFFICE_MCP_UI_RUNTIME_PATH = previousRuntimePath;
    rmSync(runtimeDir, { recursive: true, force: true });
  }
});

async function* chunks(values: string[]): AsyncIterable<Buffer> {
  for (const value of values) yield Buffer.from(value);
}

async function daemonConfig(): Promise<DaemonConfig> {
  const mcpPort = await freePort();
  const addinPort = await freePort();
  return {
    addin: {
      host: 'localhost',
      port: addinPort,
      origin: `https://localhost:${addinPort}`,
      pfxPath: '.office-mcp-localhost.pfx',
      pfxPassphrase: 'office-mcp-localhost',
      heartbeatIntervalSec: 30,
      heartbeatTimeoutSec: 10,
      sessionGraceSec: 60,
      maxPendingPerSession: 4,
      sharedSecret: ''
    },
    mcp: { host: '127.0.0.1', port: mcpPort, apiKey: '' },
    limits: { maxResponseBytes: 1024 * 1024, maxRequestBytes: 16 * 1024 * 1024, maxWsFrameBytes: 16 * 1024 * 1024, defaultToolTimeoutMs: 30000, requestsPerMinute: 120 },
    audit: { enabled: false, path: 'audit.jsonl' },
    logging: { level: 'error', file: '' }
  };
}

function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = createServer();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      server.close((error) => {
        if (error) reject(error);
        else if (address && typeof address === 'object') resolve(address.port);
        else reject(new Error('Could not allocate test port.'));
      });
    });
  });
}

function httpsJson(url: string, headers: Record<string, string> = {}): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const req = httpsRequest(url, { method: 'GET', headers, rejectUnauthorized: false }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
      res.on('end', () => resolve({ status: res.statusCode ?? 0, body: Buffer.concat(chunks).toString('utf8') }));
    });
    req.on('error', reject);
    req.end();
  });
}
