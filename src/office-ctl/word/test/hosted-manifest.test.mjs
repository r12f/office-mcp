import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';

const ADDIN_ROOT = process.cwd();
const SCRIPT = join(ADDIN_ROOT, 'scripts', 'render-hosted-manifest.ps1');
const POWERSHELL = process.platform === 'win32' ? 'powershell' : 'pwsh';

test('hosted manifest renderer emits public office-mcp.dev URLs', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-hosted-manifest-'));
  try {
    const output = join(dir, 'manifest.xml');
    const result = spawnSync(POWERSHELL, [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      SCRIPT,
      '-BaseUrl',
      'https://office-mcp.dev',
      '-AddinId',
      '22222222-aaaa-bbbb-cccc-333333333333',
      '-AddinVersion',
      '1.2.3.4',
      '-AssetVersion',
      '1.2.3',
      '-OutputPath',
      output
    ], { encoding: 'utf8' });

    assert.equal(result.status, 0, result.stderr);
    const xml = readFileSync(output, 'utf8');
    assert.match(xml, /<Id>22222222-aaaa-bbbb-cccc-333333333333<\/Id>/);
    assert.match(xml, /<Version>1\.2\.3\.4<\/Version>/);
    assert.match(xml, /<DisplayName DefaultValue="Office MCP Control" \/>/);
    assert.match(xml, /DefaultValue="Open Control Panel"/);
    assert.match(xml, /https:\/\/office-mcp\.dev\/taskpane\.html\?v=1\.2\.3/);
    assert.match(xml, /https:\/\/office-mcp\.dev\/assets\/icon-32\.png/);
    assert.match(xml, /https:\/\/office-mcp\.dev\/assets\/icon-80\.png/);
    assert.doesNotMatch(xml, /localhost|127\.0\.0\.1/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('hosted manifest renderer rejects loopback and non-HTTPS origins', () => {
  const loopback = spawnSync(POWERSHELL, [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    SCRIPT,
    '-BaseUrl',
    'https://localhost:8765'
  ], { encoding: 'utf8' });
  assert.notEqual(loopback.status, 0);
  assert.match(loopback.stderr, /loopback/i);

  const http = spawnSync(POWERSHELL, [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    SCRIPT,
    '-BaseUrl',
    'http://office-mcp.dev'
  ], { encoding: 'utf8' });
  assert.notEqual(http.status, 0);
  assert.match(http.stderr, /https:\/\//);
});
