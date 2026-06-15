import assert from 'node:assert/strict';
import test from 'node:test';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { readJsonWithLimit, resolveAddinPublicDir } from '../src/daemon.js';

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

async function* chunks(values: string[]): AsyncIterable<Buffer> {
  for (const value of values) yield Buffer.from(value);
}
