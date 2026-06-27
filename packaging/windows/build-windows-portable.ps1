param(
  [string]$Version = "0.1.6",
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
  foreach ($requiredInstallBehavior in @("param\(", "OFFICE_MCP_INSTALL_ROOT", "Copy-PortablePayload", "Stop-OfficeMcpDaemons", "Remove-LegacyVersionedInstallRoots", "Remove-StaleOfficeMcpTrustedCatalogs", "Read-Host", "-InstallRoot")) {
    if ($installScript -notmatch $requiredInstallBehavior) {
      throw "Portable install script must include fixed-root upgrade behavior: $requiredInstallBehavior."
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
param(
  [string]$InstallRoot = $(if (-not [string]::IsNullOrWhiteSpace($env:OFFICE_MCP_INSTALL_ROOT)) { $env:OFFICE_MCP_INSTALL_ROOT } else { Join-Path $env:LOCALAPPDATA 'office-mcp' }),
  [switch]$CloseOfficeHosts,
  [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'
$packageRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $MyInvocation.MyCommand.Path))
$installRoot = [System.IO.Path]::GetFullPath($InstallRoot)

function Test-SamePath {
  param(
    [Parameter(Mandatory = $true)][string]$Left,
    [Parameter(Mandatory = $true)][string]$Right
  )

  return [System.String]::Equals(
    [System.IO.Path]::GetFullPath($Left).TrimEnd('\'),
    [System.IO.Path]::GetFullPath($Right).TrimEnd('\'),
    [System.StringComparison]::OrdinalIgnoreCase)
}

function Copy-PortablePayload {
  param(
    [Parameter(Mandatory = $true)][string]$PackageRoot,
    [Parameter(Mandatory = $true)][string]$InstallRoot
  )

  if (Test-SamePath -Left $PackageRoot -Right $InstallRoot) { return }

  New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null

  foreach ($relativePath in @('office-mcp-daemon.exe', 'office-mcp', 'office-ctl', 'addin-catalog', 'scripts', 'install.ps1', 'uninstall.ps1', 'README-install.txt')) {
    $targetPath = Join-Path $InstallRoot $relativePath
    if (Test-Path -LiteralPath $targetPath) {
      Remove-Item -LiteralPath $targetPath -Recurse -Force
    }
  }

  foreach ($relativePath in @('office-mcp-daemon.exe', 'office-mcp', 'office-ctl', 'addin-catalog', 'scripts', 'install.ps1', 'uninstall.ps1', 'README-install.txt')) {
    $sourcePath = Join-Path $PackageRoot $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath)) {
      throw "Package payload is missing required path: $relativePath"
    }
    Copy-Item -LiteralPath $sourcePath -Destination $InstallRoot -Recurse -Force
  }

  $targetConfigPath = Join-Path $InstallRoot 'config.toml'
  if (-not (Test-Path -LiteralPath $targetConfigPath)) {
    Copy-Item -LiteralPath (Join-Path $PackageRoot 'config.toml') -Destination $targetConfigPath -Force
  }
}

function Stop-OfficeMcpDaemons {
  $processes = @(Get-CimInstance Win32_Process -Filter "Name = 'office-mcp-daemon.exe'" -ErrorAction SilentlyContinue)
  foreach ($processInfo in $processes) {
    try {
      Stop-Process -Id $processInfo.ProcessId -Force -ErrorAction Stop
      Write-Output "Stopped existing Office MCP Control daemon. PID: $($processInfo.ProcessId)"
    } catch {
      throw "Failed to stop existing Office MCP Control daemon PID $($processInfo.ProcessId): $($_.Exception.Message)"
    }
  }
}

function Remove-LegacyVersionedInstallRoots {
  param(
    [Parameter(Mandatory = $true)][string]$InstallRoot
  )

  if (-not (Test-Path -LiteralPath $InstallRoot)) { return }
  Get-ChildItem -LiteralPath $InstallRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^v\d+\.\d+\.\d+' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'office-mcp-daemon.exe')) } |
    ForEach-Object {
      Remove-Item -LiteralPath $_.FullName -Recurse -Force
      Write-Output "Removed legacy versioned install root: $($_.FullName)"
    }
}

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

function Get-RunningOfficeHosts {
  $running = @()
  foreach ($processName in @('WINWORD', 'EXCEL', 'POWERPNT')) {
    $processes = @(Get-Process -Name $processName -ErrorAction SilentlyContinue)
    foreach ($process in $processes) {
      $running += [pscustomobject]@{ Name = $processName; Id = $process.Id }
    }
  }
  return $running
}

function Assert-OfficeHostsClosed {
  $running = @(Get-RunningOfficeHosts)
  if ($running.Count -gt 0) {
    $display = ($running | ForEach-Object { "$($_.Name)($($_.Id))" }) -join ', '
    if (-not $CloseOfficeHosts) {
      if ($NonInteractive -or -not [Environment]::UserInteractive) {
        throw "Word, Excel, or PowerPoint is running: $display. Re-run after closing Office, or pass -CloseOfficeHosts to let the installer close them."
      }
      $answer = Read-Host "Office MCP Control must close these Office apps so the trusted catalog reloads: $display. Close them now? [y/N]"
      if ($answer -notin @('y', 'Y', 'yes', 'YES')) {
        throw "Install cancelled because Office hosts are still running: $display"
      }
    }
    foreach ($hostProcess in $running) {
      Stop-Process -Id $hostProcess.Id -Force
      Write-Output "Closed Office host $($hostProcess.Name). PID: $($hostProcess.Id)"
    }
  }
}

function Remove-StaleOfficeMcpTrustedCatalogs {
  param(
    [Parameter(Mandatory = $true)][string]$CurrentCatalogKey
  )

  $trustedRoot = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
  if (-not (Test-Path -LiteralPath $trustedRoot)) { return }
  Get-ChildItem -LiteralPath $trustedRoot -ErrorAction SilentlyContinue |
    Where-Object { $_.PSPath -ne $CurrentCatalogKey } |
    ForEach-Object {
      $props = Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue
      $url = [string]$props.Url
      if ($url -match '\\office-mcp\\v\d+\.\d+\.\d+.*\\addin-catalog') {
        Remove-Item -LiteralPath $_.PSPath -Recurse -Force
        Write-Output "Removed stale Office MCP trusted catalog: $($_.PSChildName)"
      }
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
Stop-OfficeMcpDaemons
Copy-PortablePayload -PackageRoot $packageRoot -InstallRoot $installRoot
Remove-LegacyVersionedInstallRoots -InstallRoot $installRoot
Set-OfficeMcpPortableEnvironment -InstallRoot $installRoot
if (Test-Path -LiteralPath $legacyCatalogKey) {
  Remove-Item -LiteralPath $legacyCatalogKey -Recurse -Force
}
Remove-StaleOfficeMcpTrustedCatalogs -CurrentCatalogKey $catalogKey
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

This folder is installation media. The installer copies Office MCP Control into
a stable install root, %LOCALAPPDATA%\office-mcp by default, so upgrades do not
create a new versioned install folder each time.

Contents:
- office-mcp-daemon.exe: native daemon, tray host, and MCP server.
- office-mcp\ui\: daemon control panel assets.
- office-ctl\word, office-ctl\excel, office-ctl\powerpoint: Office add-in bundles.
- addin-catalog\: Word, Excel, and PowerPoint shared-folder catalog manifests.
- scripts\export-localhost-dev-cert.ps1: creates/exports the localhost HTTPS certificate.
- install.ps1: installs or upgrades the fixed install root, registers the current user's Office trusted catalog, creates the localhost certificate if needed, and starts the daemon runtime with tray support.
- uninstall.ps1: removes Office MCP Control user registry entries.

What install.ps1 changes:
- Stops existing Office MCP Control daemon processes before replacing runtime files.
- Writes HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}
  - Id = {6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}
  - Url = the UNC path for this folder's addin-catalog directory
  - Flags = 1, which means Show in Menu
- Removes stale Office MCP trusted catalog paths and the old invalid HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp key if present.
- Removes safe-to-identify legacy versioned install roots such as %LOCALAPPDATA%\office-mcp\v0.1.5.
- Creates .office-mcp-localhost.pfx in the install root if it is missing.
- Uses CurrentUser certificate stores only.

Install:
1. Extract the zip to a temporary folder where you can inspect it.
2. Run PowerShell from this folder:
   powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1

Custom install root:
   powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -InstallRoot D:\Apps\OfficeMcp

Office host handling:
If Word, Excel, or PowerPoint is running, install.ps1 shows the running hosts
and asks before closing them. In non-interactive runs, close Office first or use
-CloseOfficeHosts.

After install:
3. The installer prints the install root. From that folder, run:
   .\office-mcp-daemon.exe daemon status
4. Reopen Office and use Home > Add-ins > Advanced > Shared Folder to add Office MCP Control if Office does not show it automatically.

Default endpoints:
- MCP: http://127.0.0.1:8800/mcp
- Add-in UI/WSS origin: https://localhost:8765
"@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "README-install.txt")

Assert-PortableStagePayload -StageRoot $stageRoot
New-PortableZip -StageRoot $stageRoot -OutputPath $zipOutputPath
Write-Output "Built $zipOutputPath"
