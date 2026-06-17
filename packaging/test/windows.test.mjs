import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import test from 'node:test';

const PACKAGING_ROOT = process.cwd();

test('Windows tray controller exposes required notification-area menu', () => {
  const script = readFileSync(join(PACKAGING_ROOT, 'windows', 'office-mcp-tray.ps1'), 'utf8');
  const repoRoot = join(PACKAGING_ROOT, '..');
  const rustTrayHost = readFileSync(join(repoRoot, 'src', 'office-mcp', 'daemon', 'src', 'tray_host.rs'), 'utf8');
  const rustTrayController = readFileSync(join(repoRoot, 'src', 'office-mcp', 'daemon', 'src', 'tray_controller.rs'), 'utf8');

  assert.match(script, /office-mcp-daemon\.exe/);
  assert.match(script, /"tray"/);
  assert.match(script, /--probe/);
  assert.match(rustTrayHost, /tray_icon::/);
  assert.match(rustTrayHost, /tao::event_loop/);
  assert.match(rustTrayHost, /MenuItem::with_id/);
  assert.match(rustTrayHost, /office-mcp-show/);
  assert.match(rustTrayHost, /office-mcp-quit/);
  assert.match(rustTrayController, /Status:/);
  assert.match(rustTrayController, /Clients:/);
  assert.match(rustTrayController, /Documents:/);
  assert.match(rustTrayController, /current_tasks/);
  assert.match(rustTrayController, /Show Office MCP/);
  assert.match(rustTrayController, /Quit Office MCP/);
  assert.match(rustTrayController, /Keep Running/);
  assert.match(rustTrayController, /running tasks/);
  assert.doesNotMatch(script, /node\.exe/);
  assert.doesNotMatch(script, /reference-node/);
  assert.doesNotMatch(script, /dist\src\cli\.js/);
  assert.doesNotMatch(script, /System\.Windows\.Forms\.NotifyIcon/);
  assert.doesNotMatch(script, /ContextMenuStrip/);
  assert.doesNotMatch(script, new RegExp(['x-office-mcp-ui', 'token'].join('-'), 'i'));
  assert.doesNotMatch(script, new RegExp(['ui', 'Token'].join('')));
  assert.doesNotMatch(script, /Start-ScheduledTask/);
});

test('Windows packaging includes the tray controller in installer payload', () => {
  const buildScript = readFileSync(join(PACKAGING_ROOT, 'windows', 'build-windows-msi.ps1'), 'utf8');
  const installScript = readFileSync(join(PACKAGING_ROOT, 'windows', 'install-windows.ps1'), 'utf8');
  const productWxs = readFileSync(join(PACKAGING_ROOT, 'wix', 'Product.wxs'), 'utf8');

  assert.match(buildScript, /office-mcp-tray\.ps1/);
  assert.match(buildScript, /cargo build --release -p office-mcp-daemon/);
  assert.match(buildScript, /office-mcp-daemon\.exe/);
  assert.match(buildScript, /office-mcp\\ui\\index\.html/);
  assert.match(buildScript, /office-ctl\\common\\addin-channel\.js/);
  assert.match(buildScript, /office-ctl\\common\\browser-ui\.js/);
  assert.match(buildScript, /office-ctl\\common\\logger\.js/);
  assert.match(buildScript, /office-ctl\\common\\task-history\.js/);
  assert.match(buildScript, /office-ctl\\excel\\manifest\.xml/);
  assert.match(buildScript, /office-ctl\\excel\\public\\taskpane\.js/);
  assert.match(buildScript, /addin-catalog\\office-mcp-word\.xml/);
  assert.match(buildScript, /addin-catalog\\office-mcp-excel\.xml/);
  assert.match(buildScript, /scripts\\export-localhost-dev-cert\.ps1/);
  assert.match(buildScript, /Tray launcher must delegate to the native Rust tray host/);
  assert.doesNotMatch(buildScript, /node\\node\.exe/);
  assert.doesNotMatch(buildScript, /dist\\src\\cli\.js/);
  assert.doesNotMatch(buildScript, /reference-node\\scripts/);
  assert.match(installScript, /office-mcp-tray\.ps1/);
  assert.match(installScript, /cargo build --release -p office-mcp-daemon/);
  assert.match(installScript, /office-mcp-daemon\.exe/);
  assert.match(installScript, /office-mcp\\ui/);
  assert.match(installScript, /office-ctl\\common/);
  assert.match(installScript, /office-ctl\\word/);
  assert.match(installScript, /office-ctl\\excel/);
  assert.match(installScript, /installedWordRoot/);
  assert.match(installScript, /office-mcp-word\.xml/);
  assert.match(installScript, /office-mcp-excel\.xml/);
  assert.match(installScript, /commonRoot/);
  assert.match(installScript, /function ConvertTo-OfficeCatalogUrl/);
  assert.match(installScript, /\\\\localhost/);
  assert.match(installScript, /-Name Url -Value \$catalogUrl/);
  assert.match(installScript, /Catalog URL: \$catalogUrl/);
  assert.doesNotMatch(installScript, new RegExp(['npm run', 'daemon'].join(' ')));
  assert.match(installScript, /Tray launcher:/);
  assert.match(installScript, /New-ScheduledTaskAction[\s\S]*office-mcp-daemon\.exe[\s\S]*tray/);
  assert.match(productWxs, /Windows\\CurrentVersion\\Run/);
  assert.match(productWxs, /office-mcp-daemon\.exe&quot; tray/);
  assert.doesNotMatch(productWxs, /office-mcp-daemon\.ps1/);
  assert.doesNotMatch(productWxs, /office-mcp-tray\.ps1/);
});

test('Windows localhost certificate helper lives under packaging', () => {
  const helper = readFileSync(join(PACKAGING_ROOT, 'windows', 'export-localhost-dev-cert.ps1'), 'utf8');

  assert.match(helper, /Cert:\\CurrentUser\\My/);
  assert.match(helper, /Cert:\\CurrentUser\\Root/);
  assert.match(helper, /Export-PfxCertificate/);
  assert.doesNotMatch(helper, /Import-PfxCertificate/);
});
