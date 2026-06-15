import assert from 'node:assert/strict';
import test from 'node:test';
import { join } from 'node:path';
import { buildClaudeDesktopConfig } from '../src/client-config.js';
import { readFileSync } from 'node:fs';

test('builds Claude Desktop config for a source checkout', () => {
  const config = buildClaudeDesktopConfig({ mode: 'dev', cwd: 'C:\\Code\\office-mcp\\mcp-server' });

  assert.deepEqual(config, {
    mcpServers: {
      'office-mcp': {
        command: 'node',
        args: ['dist/src/cli.js', 'stdio'],
        cwd: 'C:\\Code\\office-mcp\\mcp-server'
      }
    }
  });
});

test('builds Claude Desktop config for the Windows MSI install layout', () => {
  const installRoot = 'C:\\Users\\Riff\\AppData\\Local\\office-mcp';
  const config = buildClaudeDesktopConfig({ mode: 'installed', installRoot });

  assert.deepEqual(config, {
    mcpServers: {
      'office-mcp': {
        command: 'powershell.exe',
        args: ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(installRoot, 'office-mcp.ps1'), 'stdio']
      }
    }
  });
  assert.equal('cwd' in config.mcpServers['office-mcp'], false);
});

test('uses OFFICE_MCP_INSTALL_ROOT for installed Claude Desktop config by default', () => {
  const previousInstallRoot = process.env.OFFICE_MCP_INSTALL_ROOT;
  process.env.OFFICE_MCP_INSTALL_ROOT = 'D:\\Apps\\office-mcp';
  try {
    const config = buildClaudeDesktopConfig({ mode: 'installed' });

    assert.deepEqual(config.mcpServers['office-mcp'].args, [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      join('D:\\Apps\\office-mcp', 'office-mcp.ps1'),
      'stdio'
    ]);
  } finally {
    if (previousInstallRoot === undefined) {
      delete process.env.OFFICE_MCP_INSTALL_ROOT;
    } else {
      process.env.OFFICE_MCP_INSTALL_ROOT = previousInstallRoot;
    }
  }
});

test('CLI exposes daemon UI command in usage text', () => {
  const source = readFileSync(join(process.cwd(), 'src', 'cli.ts'), 'utf8');

  assert.match(source, /command\[0\] === 'ui'/);
  assert.match(source, /office-mcp ui/);
});
