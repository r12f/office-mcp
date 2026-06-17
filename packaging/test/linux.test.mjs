import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';

const PACKAGING_ROOT = process.cwd();
const REPO_ROOT = resolve(PACKAGING_ROOT, '..');
const SCRIPT = join(PACKAGING_ROOT, 'linux', 'render-systemd-unit.ps1');
const TEMPLATE = join(PACKAGING_ROOT, 'linux', 'office-mcp.service.in');
const POWERSHELL = process.platform === 'win32' ? 'powershell' : 'pwsh';

test('Linux systemd user unit template runs the Rust daemon service', () => {
  const template = readFileSync(TEMPLATE, 'utf8');

  assert.match(template, /Description=Office MCP daemon/);
  assert.match(template, /OFFICE_MCP_INSTALL_ROOT=\{\{INSTALL_ROOT\}\}/);
  assert.match(template, /OFFICE_MCP_CONFIG_PATH=\{\{CONFIG_PATH\}\}/);
  assert.match(template, /ExecStart=\{\{INSTALL_ROOT\}\}\/office-mcp-daemon daemon run/);
  assert.match(template, /OFFICE_MCP_INSTALL_ROOT=\{\{INSTALL_ROOT\}\}/);
  assert.match(template, /Restart=on-failure/);
  assert.match(template, /WantedBy=default\.target/);
  assert.doesNotMatch(template, /node|reference-node|mcp-server|dist\/src\/cli\.js/);
});

test('Linux systemd unit renderer creates a user service file', () => {
  const dir = mkdtempSync(join(tmpdir(), 'office-mcp-systemd-'));
  try {
    const output = join(dir, 'office-mcp.service');
    const result = spawnSync(POWERSHELL, [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      SCRIPT,
      '-InstallRoot',
      '/opt/office-mcp/libexec',
      '-ConfigPath',
      '/home/alice/.config/office-mcp/config.toml',
      '-OutputPath',
      output
    ], { cwd: REPO_ROOT, encoding: 'utf8' });

    assert.equal(result.status, 0, result.stderr);
    const unit = readFileSync(output, 'utf8');
    assert.match(unit, /Environment=OFFICE_MCP_INSTALL_ROOT=\/opt\/office-mcp\/libexec/);
    assert.match(unit, /Environment=OFFICE_MCP_CONFIG_PATH=\/home\/alice\/\.config\/office-mcp\/config\.toml/);
    assert.match(unit, /ExecStart=\/opt\/office-mcp\/libexec\/office-mcp-daemon daemon run/);
    assert.doesNotMatch(unit, /node|reference-node|mcp-server|dist\/src\/cli\.js/);
    assert.doesNotMatch(unit, /\{\{/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('Linux systemd unit renderer rejects non-absolute paths', () => {
  const result = spawnSync(POWERSHELL, [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    SCRIPT,
    '-InstallRoot',
    'relative/path',
    '-ConfigPath',
    '/home/alice/.config/office-mcp/config.toml'
  ], { cwd: REPO_ROOT, encoding: 'utf8' });

  assert.notEqual(result.status, 0);
  assert.match(result.stderr, /absolute Unix path/);
});
