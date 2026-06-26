param(
  [string]$Version = "0.1.4",
  [string]$OutputDir = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) "artifacts"),
  [switch]$SkipNpmInstall
)

$ErrorActionPreference = "Stop"

function Assert-LastExitCode([string]$CommandName) {
  if ($LASTEXITCODE -ne 0) {
    throw "$CommandName failed with exit code $LASTEXITCODE"
  }
}

function Reset-DirectoryInside([string]$Path, [string]$Parent) {
  New-Item -ItemType Directory -Force -Path $Parent | Out-Null
  $resolvedParent = (Resolve-Path -LiteralPath $Parent).Path
  if (Test-Path -LiteralPath $Path) {
    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    if (-not $resolvedPath.StartsWith($resolvedParent, [System.StringComparison]::OrdinalIgnoreCase)) {
      throw "Refusing to remove staging directory outside output directory: $resolvedPath"
    }
    Remove-Item -LiteralPath $resolvedPath -Recurse -Force
  }
  New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

function Assert-PortableStagePayload([string]$StageRoot) {
  $requiredPaths = @(
    "office-mcp-daemon.exe",
    "office-mcp\ui\index.html",
    "office-mcp\ui\app.css",
    "office-mcp\ui\app.js",
    "office-ctl\common\addin-channel.js",
    "office-ctl\common\browser-ui.js",
    "office-ctl\common\logger.js",
    "office-ctl\common\task-history.js",
    "office-ctl\common\assets\brand-mark.svg",
    "office-ctl\common\assets\icon-16.png",
    "office-ctl\common\assets\icon-32.png",
    "office-ctl\common\assets\icon-80.png",
    "office-ctl\common\assets\icon-256.png",
    "scripts\export-localhost-dev-cert.ps1",
    "office-ctl\word\manifest.xml",
    "office-ctl\word\public\taskpane.html",
    "office-ctl\word\public\taskpane.css",
    "office-ctl\word\public\taskpane.js",
    "office-ctl\excel\manifest.xml",
    "office-ctl\excel\public\taskpane.html",
    "office-ctl\excel\public\taskpane.css",
    "office-ctl\excel\public\taskpane.js",
    "office-ctl\powerpoint\manifest.xml",
    "office-ctl\powerpoint\public\taskpane.html",
    "office-ctl\powerpoint\public\taskpane.css",
    "office-ctl\powerpoint\public\taskpane.js",
    "addin-catalog\office-mcp-word.xml",
    "addin-catalog\office-mcp-excel.xml",
    "addin-catalog\office-mcp-powerpoint.xml",
    "install.ps1",
    "uninstall.ps1",
    "README-install.txt",
    "config.toml"
  )

  foreach ($relativePath in $requiredPaths) {
    $path = Join-Path $StageRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
      throw "Portable staging payload is missing required path: $relativePath"
    }
  }

  foreach ($forbiddenRootScript in @("office-mcp.ps1", "office-mcp-daemon.ps1", "office-mcp-tray.ps1", "office-mcp-env.ps1")) {
    if (Test-Path -LiteralPath (Join-Path $StageRoot $forbiddenRootScript)) {
      throw "Portable staging payload must not expose duplicate root launcher: $forbiddenRootScript"
    }
  }

  $installScript = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "install.ps1")
  if ($installScript -notmatch "ConvertTo-OfficeCatalogUrl" -or $installScript -notmatch "\\\\localhost" -or $installScript -notmatch "export-localhost-dev-cert\.ps1" -or $installScript -notmatch "-CreateIfMissing" -or $installScript -notmatch "6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57") {
    throw "Portable install script must register a GUID-based UNC Office catalog and create the localhost certificate."
  }
  foreach ($requiredCatalogInstallBehavior in @("Assert-OfficeHostsClosed", "Remove-OfficeAddinCache", "Remove-CustomUiValidationCache")) {
    if ($installScript -notmatch $requiredCatalogInstallBehavior) {
      throw "Portable install script must $requiredCatalogInstallBehavior before completing Office catalog registration."
    }
  }
  foreach ($envName in @("OFFICE_MCP_INSTALL_ROOT", "OFFICE_MCP_CONFIG_PATH", "OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH", "OFFICE_MCP_ADDIN_CHANNEL__PORT", "OFFICE_MCP_MCP_HTTP__PORT")) {
    if ($installScript -notmatch $envName) {
      throw "Portable install script must set $envName."
    }
  }
  if ($installScript -notmatch "office-mcp-daemon\.exe" -or $installScript -notmatch "Start-Process" -or $installScript -notmatch "'daemon', 'run'" -or $installScript -notmatch "WindowStyle Hidden") {
    throw "Portable install script must start the native daemon runtime without a visible console."
  }
  if ($installScript -match "office-mcp-tray\.ps1" -or $installScript -match "office-mcp-env\.ps1") {
    throw "Portable install script must not depend on duplicate launcher scripts."
  }
}

function New-PortableZip([string]$StageRoot, [string]$OutputPath) {
  if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
  }
  Compress-Archive -Path (Join-Path $StageRoot '*') -DestinationPath $OutputPath -Force
  if (-not (Test-Path -LiteralPath $OutputPath)) {
    throw "Portable ZIP was not created: $OutputPath"
  }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$rustDaemonRoot = Join-Path $repoRoot "src\office-mcp\daemon"
$evidenceRoot = Join-Path $repoRoot "src\office-mcp\daemon\evidence"
$uiRoot = Join-Path $rustDaemonRoot "src\ui\assets"
$commonRoot = Join-Path $repoRoot "src\office-ctl\common"
$addinRoot = Join-Path $repoRoot "src\office-ctl\word"
$excelAddinRoot = Join-Path $repoRoot "src\office-ctl\excel"
$powerPointAddinRoot = Join-Path $repoRoot "src\office-ctl\powerpoint"
$zipOutputPath = Join-Path $OutputDir "office-mcp-windows-portable-$Version-x64.zip"
$stageRoot = Join-Path $OutputDir "portable-stage"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

Push-Location $evidenceRoot
try {
  if (-not $SkipNpmInstall) {
    if (Test-Path (Join-Path $evidenceRoot "node_modules")) {
      npm install
      Assert-LastExitCode "npm install"
    } else {
      npm ci
      Assert-LastExitCode "npm ci"
    }
  }
  npm run check
  Assert-LastExitCode "npm run check"
} finally {
  Pop-Location
}

Push-Location $repoRoot
try {
  cargo build --release -p office-mcp-daemon
  Assert-LastExitCode "cargo build --release -p office-mcp-daemon"
} finally {
  Pop-Location
}

Push-Location $addinRoot
try {
  if (-not $SkipNpmInstall) {
    npm ci
    Assert-LastExitCode "npm ci"
  }
  npm run check
  Assert-LastExitCode "npm run check"
} finally {
  Pop-Location
}

Push-Location $excelAddinRoot
try {
  if (-not $SkipNpmInstall) {
    npm ci
    Assert-LastExitCode "npm ci"
  }
  npm run check
  Assert-LastExitCode "npm run check"
} finally {
  Pop-Location
}

Push-Location $powerPointAddinRoot
try {
  if (-not $SkipNpmInstall) {
    npm ci
    Assert-LastExitCode "npm ci"
  }
  npm run check
  Assert-LastExitCode "npm run check"
} finally {
  Pop-Location
}

Reset-DirectoryInside -Path $stageRoot -Parent $OutputDir

$stageUiRoot = Join-Path $stageRoot "office-mcp\ui"
$stageCommonRoot = Join-Path $stageRoot "office-ctl\common"
$stageAddinRoot = Join-Path $stageRoot "office-ctl\word"
$stageExcelAddinRoot = Join-Path $stageRoot "office-ctl\excel"
$stagePowerPointAddinRoot = Join-Path $stageRoot "office-ctl\powerpoint"
$stageCatalogRoot = Join-Path $stageRoot "addin-catalog"
$stageScriptsRoot = Join-Path $stageRoot "scripts"
New-Item -ItemType Directory -Force -Path $stageScriptsRoot, $stageUiRoot, $stageCommonRoot, $stageAddinRoot, $stageExcelAddinRoot, $stagePowerPointAddinRoot, $stageCatalogRoot | Out-Null

Copy-Item -Force -Path (Join-Path $repoRoot "target\release\office-mcp-daemon.exe") -Destination (Join-Path $stageRoot "office-mcp-daemon.exe")
Copy-Item -Recurse -Force -Path (Join-Path $uiRoot "*") -Destination $stageUiRoot
Copy-Item -Recurse -Force -Path (Join-Path $commonRoot "*") -Destination $stageCommonRoot
Copy-Item -Force -Path (Join-Path $repoRoot "packaging\windows\export-localhost-dev-cert.ps1") -Destination (Join-Path $stageScriptsRoot "export-localhost-dev-cert.ps1")

Copy-Item -Force -Path (Join-Path $addinRoot "manifest.xml") -Destination $stageAddinRoot
Copy-Item -Recurse -Force -Path (Join-Path $addinRoot "public") -Destination $stageAddinRoot
Copy-Item -Force -Path (Join-Path $excelAddinRoot "manifest.xml") -Destination $stageExcelAddinRoot
Copy-Item -Recurse -Force -Path (Join-Path $excelAddinRoot "public") -Destination $stageExcelAddinRoot
Copy-Item -Force -Path (Join-Path $powerPointAddinRoot "manifest.xml") -Destination $stagePowerPointAddinRoot
Copy-Item -Recurse -Force -Path (Join-Path $powerPointAddinRoot "public") -Destination $stagePowerPointAddinRoot
& (Join-Path $commonRoot "scripts\register-office-catalog.ps1") -RepoRoot $repoRoot -CatalogPath $stageCatalogRoot -BaseUrl "https://localhost:8765" -SkipRegistry

@'
[addin_channel]
bind = "localhost"
port = 8765
heartbeat_interval_sec = 30
heartbeat_timeout_sec = 10
session_grace_sec = 60
max_pending_per_session = 4
certificate_path = ""

[mcp_http]
bind = "127.0.0.1"
port = 8800

[limits]
max_response_bytes = 1048576
max_request_bytes = 16777216
max_ws_frame_bytes = 16777216
default_tool_timeout_ms = 30000
requests_per_minute = 120

[audit]
enabled = false
path = ""

[logging]
level = "info"
file = ""
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "config.toml")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Set-OfficeMcpPortableEnvironment {
  param(
    [Parameter(Mandatory = $true)][string]$InstallRoot
  )

  $pfxPath = Join-Path $InstallRoot '.office-mcp-localhost.pfx'
  $env:OFFICE_MCP_INSTALL_ROOT = $InstallRoot
  $env:OFFICE_MCP_CONFIG_PATH = Join-Path $InstallRoot 'config.toml'
  $env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH = $pfxPath
  $env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE = 'office-mcp-localhost'
  $env:OFFICE_MCP_ADDIN_CHANNEL__BIND = 'localhost'
  $env:OFFICE_MCP_ADDIN_CHANNEL__PORT = '8765'
  $env:OFFICE_MCP_MCP_HTTP__BIND = '127.0.0.1'
  $env:OFFICE_MCP_MCP_HTTP__PORT = '8800'
}

Set-OfficeMcpPortableEnvironment -InstallRoot $installRoot

function ConvertTo-OfficeCatalogUrl {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $resolvedPath = [System.IO.Path]::GetFullPath($Path)
  if ($resolvedPath.StartsWith('\', [System.StringComparison]::Ordinal)) {
    return $resolvedPath
  }

  $root = [System.IO.Path]::GetPathRoot($resolvedPath)
  if ([string]::IsNullOrWhiteSpace($root) -or $root.Length -lt 2 -or $root[1] -ne ':') {
    throw "CatalogPath must be an absolute drive path or UNC path: $Path"
  }

  $drive = $root.Substring(0, 1).ToUpperInvariant()
  $relativePath = $resolvedPath.Substring($root.Length).TrimStart('\', '/')
  return "\\localhost\$drive`$\$relativePath"
}

function Assert-OfficeHostsClosed {
  $running = @()
  foreach ($processName in @('WINWORD', 'EXCEL', 'POWERPNT')) {
    if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
      $running += $processName
    }
  }
  if ($running.Count -gt 0) {
    throw "Close Word, Excel, and PowerPoint before installing Office MCP Control so Office can reload the trusted add-in catalog. Running processes: $($running -join ', ')."
  }
}

function Remove-DeveloperDebugRegistration {
  param(
    [Parameter(Mandatory = $true)][string]$AddinId
  )

  $developerRoot = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
  $developerValue = Get-ItemProperty -LiteralPath $developerRoot -Name $AddinId -ErrorAction SilentlyContinue
  if ($developerValue) {
    Remove-ItemProperty -LiteralPath $developerRoot -Name $AddinId -Force
  }

  $developerKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\$AddinId"
  if (Test-Path -LiteralPath $developerKey) {
    Remove-Item -LiteralPath $developerKey -Recurse -Force
  }
}

function Remove-OfficeAddinCache {
  $wefRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
  if (-not (Test-Path -LiteralPath $wefRoot)) { return }

  $hosts = @(
    @{ Name = 'Word'; AddinId = '11111111-aaaa-bbbb-cccc-222222222222' },
    @{ Name = 'Excel'; AddinId = '33333333-aaaa-bbbb-cccc-444444444444' },
    @{ Name = 'PowerPoint'; AddinId = '44444444-aaaa-bbbb-cccc-555555555555' }
  )

  foreach ($officeHost in $hosts) {
    $addinId = $officeHost.AddinId
    $hostName = $officeHost.Name

    Remove-DeveloperDebugRegistration -AddinId $addinId

    $addinInfoRoot = Join-Path $wefRoot "AddinInfo\1\filesystem\$hostName\1"
    if (Test-Path -LiteralPath $addinInfoRoot) {
      Get-ChildItem -LiteralPath $addinInfoRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name.StartsWith($addinId, [System.StringComparison]::OrdinalIgnoreCase) } |
        Remove-Item -Recurse -Force
    }

    Get-ChildItem -LiteralPath $wefRoot -Recurse -File -ErrorAction SilentlyContinue |
      Where-Object { $_.FullName -match '\\Manifests\\' -and $_.Name.StartsWith($addinId, [System.StringComparison]::OrdinalIgnoreCase) } |
      Remove-Item -Force

    $appCommandsRoot = Join-Path $wefRoot 'AppCommands'
    if (Test-Path -LiteralPath $appCommandsRoot) {
      Get-ChildItem -LiteralPath $appCommandsRoot -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object {
          ($_.FullName -match '\\TrustedCatalog\\' -and $_.Name.StartsWith($addinId, [System.StringComparison]::OrdinalIgnoreCase)) -or
          ($_.Name -like "$hostName.RibbonCache.*")
        } |
        Remove-Item -Force
    }
  }
}

function Remove-CustomUiValidationCache {
  $cacheKey = 'HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache'
  if (-not (Test-Path -LiteralPath $cacheKey)) { return }

  $addinIds = @(
    '11111111-aaaa-bbbb-cccc-222222222222',
    '33333333-aaaa-bbbb-cccc-444444444444',
    '44444444-aaaa-bbbb-cccc-555555555555'
  )
  $cache = Get-ItemProperty -LiteralPath $cacheKey
  foreach ($addinId in $addinIds) {
    $cache.PSObject.Properties |
      Where-Object { $_.Name.StartsWith($addinId, [System.StringComparison]::OrdinalIgnoreCase) } |
      ForEach-Object { Remove-ItemProperty -LiteralPath $cacheKey -Name $_.Name -Force }
  }
}

$catalogPath = Join-Path $installRoot 'addin-catalog'
$catalogUrl = ConvertTo-OfficeCatalogUrl -Path $catalogPath
$catalogId = '{6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}'
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$legacyCatalogKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp'
Assert-OfficeHostsClosed
if (Test-Path -LiteralPath $legacyCatalogKey) {
  Remove-Item -LiteralPath $legacyCatalogKey -Recurse -Force
}
New-Item -Path $catalogKey -Force | Out-Null
Set-ItemProperty -Path $catalogKey -Name Id -Value $catalogId
Set-ItemProperty -Path $catalogKey -Name Url -Value $catalogUrl
Set-ItemProperty -Path $catalogKey -Name Flags -Value 1 -Type DWord
Remove-OfficeAddinCache
Remove-CustomUiValidationCache

$pfxPath = Join-Path $installRoot '.office-mcp-localhost.pfx'
if (-not (Test-Path -LiteralPath $pfxPath)) {
  & (Join-Path $installRoot 'scripts\export-localhost-dev-cert.ps1') -OutputPath $pfxPath -CreateIfMissing
}

$daemonExe = Join-Path $installRoot 'office-mcp-daemon.exe'
$escapedRoot = [System.Text.RegularExpressions.Regex]::Escape($installRoot)
$existingDaemon = Get-CimInstance Win32_Process -Filter "Name = 'office-mcp-daemon.exe'" -ErrorAction SilentlyContinue |
  Where-Object { $_.CommandLine -match $escapedRoot -and $_.CommandLine -match 'daemon' -and $_.CommandLine -match 'run' } |
  Select-Object -First 1

if ($existingDaemon) {
  Write-Output "Office MCP Control daemon is already running. PID: $($existingDaemon.ProcessId)"
} else {
  Start-Process -FilePath $daemonExe -ArgumentList 'daemon', 'run' -WorkingDirectory $installRoot -WindowStyle Hidden
  Write-Output 'Office MCP Control daemon started.'
}

Write-Output 'Office MCP Control install completed.'
Write-Output "Install root: $installRoot"
Write-Output "Catalog folder: $catalogPath"
Write-Output "Catalog URL: $catalogUrl"
Write-Output 'Office add-in catalog cache cleared. Reopen Word, Excel, or PowerPoint and add Office MCP Control from the Shared Folder catalog if it does not appear automatically.'
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "install.ps1")

@'
$ErrorActionPreference = 'Stop'
$catalogId = '{6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}'
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$legacyCatalogKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp'
foreach ($key in @($catalogKey, $legacyCatalogKey)) {
  if (Test-Path -LiteralPath $key) {
    Remove-Item -LiteralPath $key -Recurse -Force
  }
}
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
Remove-ItemProperty -LiteralPath $runKey -Name 'office-mcp' -Force -ErrorAction SilentlyContinue
Write-Output 'Office MCP Control user registry entries removed.'
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "uninstall.ps1")

@"
Office MCP Control portable Windows package

This folder is the install directory. No hidden copy step is required.

Contents:
- office-mcp-daemon.exe: native daemon, tray host, and MCP server.
- office-mcp\ui\: daemon control panel assets.
- office-ctl\word, office-ctl\excel, office-ctl\powerpoint: Office add-in bundles.
- addin-catalog\: Word, Excel, and PowerPoint shared-folder catalog manifests.
- scripts\export-localhost-dev-cert.ps1: creates/exports the localhost HTTPS certificate.
- install.ps1: registers the current user's Office trusted catalog, creates the localhost certificate if needed, and starts the daemon runtime with tray support.
- uninstall.ps1: removes Office MCP Control user registry entries.

What install.ps1 changes:
- Writes HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}
  - Id = {6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}
  - Url = the UNC path for this folder's addin-catalog directory
  - Flags = 1, which means Show in Menu
- Removes the old invalid HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp key if present.
- Creates .office-mcp-localhost.pfx in this folder if it is missing.
- Uses CurrentUser certificate stores only.

Install:
1. Extract the zip to the folder where you want Office MCP Control to live.
2. Close Word, Excel, and PowerPoint.
3. Run PowerShell from this folder:
   powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
4. Reopen Office and use Home > Add-ins > Advanced > Shared Folder to add Office MCP Control if Office does not show it automatically.

install.ps1 refuses to continue while Word, Excel, or PowerPoint is still running. Office only reloads trusted add-in catalogs on startup, so leaving an Office host open would make the Trust Center look empty even after registry registration succeeds.

Default endpoints:
- MCP: http://127.0.0.1:8800/mcp
- Add-in UI/WSS origin: https://localhost:8765
"@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "README-install.txt")

Assert-PortableStagePayload -StageRoot $stageRoot
New-PortableZip -StageRoot $stageRoot -OutputPath $zipOutputPath
Write-Output "Built $zipOutputPath"
