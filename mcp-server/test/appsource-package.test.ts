import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const REPO_ROOT = join(process.cwd(), '..');
const SCRIPT = join(REPO_ROOT, 'addin', 'scripts', 'build-appsource-package.ps1');
const POWERSHELL = process.platform === 'win32' ? 'powershell' : 'pwsh';

test('AppSource package builder emits submission artifacts without loopback URLs', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-appsource-'));
  try {
    const result = spawnSync(POWERSHELL, [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      SCRIPT,
      '-Version',
      '1.2.3',
      '-BaseUrl',
      'https://office-mcp.dev',
      '-AddinId',
      '22222222-aaaa-bbbb-cccc-333333333333',
      '-AddinVersion',
      '1.2.3.4',
      '-OutputDir',
      dir
    ], { encoding: 'utf8' });

    assert.equal(result.status, 0, result.stderr);

    const manifest = readFileSync(join(dir, 'manifest-1.2.3.xml'), 'utf8');
    assert.match(manifest, /https:\/\/office-mcp\.dev\/taskpane\.html\?v=1\.2\.3/);
    assert.doesNotMatch(manifest, /localhost|127\.0\.0\.1/);

    const metadata = JSON.parse(readFileSync(join(dir, 'appsource-metadata-1.2.3.json'), 'utf8')) as Record<string, unknown>;
    assert.equal(metadata.manifest, 'manifest-1.2.3.xml');
    assert.equal(metadata.addin_bundle, 'office-mcp-addin-1.2.3.zip');
    assert.match(String(metadata.manifest_sha256), /^[0-9a-f]{64}$/);
    assert.match(String(metadata.addin_bundle_sha256), /^[0-9a-f]{64}$/);

    const checklist = readFileSync(join(dir, 'appsource-checklist-1.2.3.md'), 'utf8');
    assert.match(checklist, /External gates before Partner Center submission/);
    assert.match(checklist, /Microsoft AppSource validation review/);

    assert.ok(readFileSync(join(dir, 'office-mcp-addin-1.2.3.zip')).byteLength > 1000);
    assert.ok(readFileSync(join(dir, 'office-mcp-appsource-1.2.3.zip')).byteLength > 1000);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('AppSource package builder rejects loopback package origins', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-appsource-reject-'));
  try {
    const result = spawnSync(POWERSHELL, [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      SCRIPT,
      '-BaseUrl',
      'https://localhost:8765',
      '-OutputDir',
      dir
    ], { encoding: 'utf8' });

    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /loopback/i);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
