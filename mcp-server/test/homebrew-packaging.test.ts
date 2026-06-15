import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const REPO_ROOT = join(process.cwd(), '..');
const SCRIPT = join(REPO_ROOT, 'packaging', 'homebrew', 'render-formula.ps1');
const TEMPLATE = join(REPO_ROOT, 'packaging', 'homebrew', 'Formula', 'office-mcp.rb.in');
const POWERSHELL = process.platform === 'win32' ? 'powershell' : 'pwsh';

test('Homebrew formula template describes the daemon service layout', () => {
  const template = readFileSync(TEMPLATE, 'utf8');

  assert.match(template, /class OfficeMcp < Formula/);
  assert.match(template, /depends_on "node@22"/);
  assert.match(template, /OFFICE_MCP_INSTALL_ROOT/);
  assert.match(template, /OFFICE_MCP_CONFIG_PATH/);
  assert.match(template, /service do/);
  assert.match(template, /"daemon", "run"/);
  assert.match(template, /config endpoints/);
});

test('Homebrew formula renderer creates a release formula', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-homebrew-'));
  try {
    const output = join(dir, 'office-mcp.rb');
    const result = spawnSync(POWERSHELL, [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      SCRIPT,
      '-TarballUrl',
      'https://github.com/office-mcp/office-mcp/releases/download/v0.1.0/office-mcp-0.1.0-aarch64-darwin.tar.gz',
      '-TarballSha256',
      '0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef',
      '-OutputPath',
      output
    ], { encoding: 'utf8' });

    assert.equal(result.status, 0, result.stderr);
    const formula = readFileSync(output, 'utf8');
    assert.match(formula, /url "https:\/\/github\.com\/office-mcp\/office-mcp\/releases\/download\/v0\.1\.0\/office-mcp-0\.1\.0-aarch64-darwin\.tar\.gz"/);
    assert.match(formula, /sha256 "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"/);
    assert.doesNotMatch(formula, /{{TARBALL_/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('Homebrew formula renderer rejects unsafe release inputs', () => {
  const badUrl = spawnSync(POWERSHELL, [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    SCRIPT,
    '-TarballUrl',
    'http://example.test/office-mcp.tar.gz',
    '-TarballSha256',
    '0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef'
  ], { encoding: 'utf8' });
  assert.notEqual(badUrl.status, 0);
  assert.match(badUrl.stderr, /https:\/\//);

  const badSha = spawnSync(POWERSHELL, [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    SCRIPT,
    '-TarballUrl',
    'https://example.test/office-mcp.tar.gz',
    '-TarballSha256',
    'ABC'
  ], { encoding: 'utf8' });
  assert.notEqual(badSha.status, 0);
  assert.match(badSha.stderr, /SHA-256/);
});
