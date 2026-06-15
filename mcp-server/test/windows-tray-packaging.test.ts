import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const REPO_ROOT = join(process.cwd(), '..');

test('Windows tray controller exposes required notification-area menu', () => {
  const script = readFileSync(join(REPO_ROOT, 'packaging', 'windows', 'office-mcp-tray.ps1'), 'utf8');

  assert.match(script, /System\.Windows\.Forms\.NotifyIcon/);
  assert.match(script, /Status:/);
  assert.match(script, /Clients:/);
  assert.match(script, /Documents:/);
  assert.match(script, /Show Office MCP/);
  assert.match(script, /Quit Office MCP/);
  assert.match(script, /ui-runtime\.json/);
  assert.match(script, /Invoke-RestMethod/);
  assert.match(script, /x-office-mcp-ui-token/);
  assert.match(script, /ContextMenuStrip/);
  assert.match(script, /Start-Process/);
  assert.match(script, /"daemon", "start"/);
  assert.doesNotMatch(script, /Start-ScheduledTask/);
});

test('Windows packaging includes the tray controller in installer payload', () => {
  const buildScript = readFileSync(join(REPO_ROOT, 'packaging', 'windows', 'build-windows-msi.ps1'), 'utf8');
  const installScript = readFileSync(join(REPO_ROOT, 'packaging', 'windows', 'install-windows.ps1'), 'utf8');
  const productWxs = readFileSync(join(REPO_ROOT, 'packaging', 'wix', 'Product.wxs'), 'utf8');

  assert.match(buildScript, /office-mcp-tray\.ps1/);
  assert.match(buildScript, /Tray launcher must expose the required notification-area menu/);
  assert.match(installScript, /office-mcp-tray\.ps1/);
  assert.match(installScript, /Tray launcher:/);
  assert.match(installScript, /New-ScheduledTaskAction[\s\S]*\$trayLauncherPath/);
  assert.match(productWxs, /Windows\\CurrentVersion\\Run/);
  assert.match(productWxs, /office-mcp-tray\.ps1/);
  assert.doesNotMatch(productWxs, /office-mcp-daemon\.ps1/);
});
