import { readFileSync } from 'node:fs';
import { createServer as createHttpServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { createServer as createHttpsServer } from 'node:https';
import { basename, dirname, extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { DaemonConfig } from './config.js';
import { assertHttpsConfig } from './config.js';
import { createAddinChannel, handleUpgrade } from './addin-channel.js';
import { McpFrontend } from './mcp-server.js';
import { SessionRegistry } from './session-registry.js';
import { createLogger } from './logger.js';

const PUBLIC_DIR = resolveAddinPublicDir();

export function resolveAddinPublicDir(moduleUrl = import.meta.url, installRoot = process.env.OFFICE_MCP_INSTALL_ROOT): string {
  if (installRoot) return join(installRoot, 'addin', 'public');

  const moduleDir = dirname(fileURLToPath(moduleUrl));
  const parentDir = dirname(moduleDir);
  const serverRoot = basename(parentDir) === 'dist' ? dirname(parentDir) : parentDir;
  return join(dirname(serverRoot), 'addin', 'public');
}

export type RunningDaemon = {
  close: () => Promise<void>;
  registry: SessionRegistry;
};

export async function startDaemon(config: DaemonConfig): Promise<RunningDaemon> {
  assertHttpsConfig(config);
  const logger = createLogger(config);
  const registry = new SessionRegistry(config.addin.maxPendingPerSession);
  const frontend = new McpFrontend(config, registry);
  const wss = createAddinChannel(config, registry);

  const mcpServer = createHttpServer(async (req, res) => {
    if (!req.url?.startsWith('/mcp')) {
      json(res, 404, { error: 'Not found' });
      return;
    }
    try {
      const body = req.method === 'POST' ? await readJsonWithLimit(req, config.limits.maxRequestBytes) : undefined;
      await frontend.handle(req, res, body);
    } catch (error) {
      if (!res.headersSent) {
        const tooLarge = error instanceof RequestTooLargeError;
        json(res, tooLarge ? 413 : 400, {
          jsonrpc: '2.0',
          id: null,
          error: { code: -32000, message: error instanceof Error ? error.message : 'Bad request' }
        });
      }
    }
  });

  const addinServer = createHttpsServer({
    pfx: readFileSync(config.addin.pfxPath),
    passphrase: config.addin.pfxPassphrase
  }, (req, res) => serveStatic(req, res));

  addinServer.on('upgrade', (req, socket, head) => handleUpgrade(config, wss, req, socket, head));

  await Promise.all([
    listen(mcpServer, config.mcp.port, config.mcp.host),
    listen(addinServer, config.addin.port, config.addin.host)
  ]);

  const pruneTimer = setInterval(() => registry.pruneStaleSessions(config.addin.sessionGraceSec), 5000);

  logger.info(`office-mcp MCP endpoint: http://${config.mcp.host}:${config.mcp.port}/mcp`, { component: 'daemon', event: 'mcp_listening', host: config.mcp.host, port: config.mcp.port });
  logger.info(`office-mcp add-in origin: ${config.addin.origin}`, { component: 'daemon', event: 'addin_listening', origin: config.addin.origin });

  return {
    registry,
    close: async () => {
      clearInterval(pruneTimer);
      await frontend.close();
      wss.close();
      await Promise.all([closeServer(mcpServer), closeServer(addinServer)]);
      logger.info('office-mcp daemon stopped', { component: 'daemon', event: 'stopped' });
    }
  };
}

function listen(server: ReturnType<typeof createHttpServer>, port: number, host: string): Promise<void> {
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, host, () => {
      server.off('error', reject);
      resolve();
    });
  });
}

function closeServer(server: ReturnType<typeof createHttpServer>): Promise<void> {
  return new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
}

class RequestTooLargeError extends Error {}

export async function readJsonWithLimit(req: AsyncIterable<Buffer | string>, maxBytes: number): Promise<unknown> {
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of req) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    total += buffer.length;
    if (total > maxBytes) throw new RequestTooLargeError(`Request body exceeds ${maxBytes} bytes.`);
    chunks.push(buffer);
  }
  if (chunks.length === 0) return undefined;
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

function serveStatic(req: IncomingMessage, res: ServerResponse): void {
  const url = new URL(req.url ?? '/', 'https://localhost');
  if (url.pathname === '/healthz') {
    json(res, 200, { ok: true });
    return;
  }
  if (url.pathname === '/assets/icon-32.png' || url.pathname === '/assets/icon-80.png') {
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/l6lJ6wAAAABJRU5ErkJggg==', 'base64');
    res.writeHead(200, { 'content-type': 'image/png', 'cache-control': 'no-store' });
    res.end(png);
    return;
  }
  const pathname = url.pathname === '/' ? '/taskpane.html' : url.pathname;
  const filePath = normalize(join(PUBLIC_DIR, pathname));
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403).end('Forbidden');
    return;
  }
  try {
    const content = readFileSync(filePath);
    res.writeHead(200, { 'content-type': contentType(filePath), 'cache-control': 'no-store' });
    res.end(content);
  } catch {
    res.writeHead(404).end('Not found');
  }
}

function contentType(path: string): string {
  switch (extname(path)) {
    case '.html': return 'text/html; charset=utf-8';
    case '.js': return 'text/javascript; charset=utf-8';
    case '.css': return 'text/css; charset=utf-8';
    default: return 'application/octet-stream';
  }
}

function json(res: ServerResponse, status: number, body: unknown): void {
  res.writeHead(status, { 'content-type': 'application/json' });
  res.end(JSON.stringify(body));
}

