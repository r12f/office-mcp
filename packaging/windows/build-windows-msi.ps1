param(
  [string]$Version = "0.1.0",
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

function New-WixPayloadFragment([string]$StageRoot, [string]$OutputPath) {
  $resolvedStageRoot = (Resolve-Path -LiteralPath $StageRoot).Path
  $directories = New-Object System.Collections.Generic.List[string]
  $refs = New-Object System.Collections.Generic.List[string]

  Add-WixDirectory -BasePath $resolvedStageRoot -Path $resolvedStageRoot -Directories $directories -Refs $refs

  $content = @()
  $content += '<?xml version="1.0" encoding="UTF-8"?>'
  $content += '<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">'
  $content += '  <Fragment>'
  $content += '    <DirectoryRef Id="INSTALLFOLDER">'
  $content += $directories
  $content += '    </DirectoryRef>'
  $content += '  </Fragment>'
  $content += '  <Fragment>'
  $content += '    <ComponentGroup Id="OfficeMcpPayloadComponents">'
  $content += $refs
  $content += '    </ComponentGroup>'
  $content += '  </Fragment>'
  $content += '</Wix>'
  $content | Set-Content -Encoding UTF8 -Path $OutputPath
}

function Add-WixDirectory(
  [string]$BasePath,
  [string]$Path,
  [System.Collections.Generic.List[string]]$Directories,
  [System.Collections.Generic.List[string]]$Refs
) {
  foreach ($childDir in Get-ChildItem -LiteralPath $Path -Directory | Sort-Object FullName) {
    $dirId = "Dir_" + (Get-StableId $BasePath $childDir.FullName)
    $name = ConvertTo-WixXml $childDir.Name
    $Directories.Add("      <Directory Id=`"$dirId`" Name=`"$name`">")
    Add-WixDirectory -BasePath $BasePath -Path $childDir.FullName -Directories $Directories -Refs $Refs
    $Directories.Add("      </Directory>")
  }

  foreach ($file in Get-ChildItem -LiteralPath $Path -File | Sort-Object FullName) {
    $componentId = "Cmp_" + (Get-StableId $BasePath $file.FullName)
    $fileId = "File_" + (Get-StableId $BasePath ($file.FullName + ":file"))
    $source = ConvertTo-WixXml $file.FullName
    $name = ConvertTo-WixXml $file.Name
    $Directories.Add("      <Component Id=`"$componentId`" Guid=`"*`">")
    $Directories.Add("        <File Id=`"$fileId`" Source=`"$source`" Name=`"$name`" KeyPath=`"yes`" />")
    $Directories.Add("      </Component>")
    $Refs.Add("      <ComponentRef Id=`"$componentId`" />")
  }
}

function Get-StableId([string]$BasePath, [string]$Path) {
  $baseUri = [System.Uri]((Join-Path $BasePath "") -replace '\\', '/')
  $pathUri = [System.Uri]($Path -replace '\\', '/')
  $relative = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($pathUri).ToString()).ToLowerInvariant()
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($relative)
  $sha = [System.Security.Cryptography.SHA1]::Create()
  try {
    return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '')
  } finally {
    $sha.Dispose()
  }
}

function ConvertTo-WixXml([string]$Value) {
  return [System.Security.SecurityElement]::Escape($Value)
}

function Assert-MsiStagePayload([string]$StageRoot, [string]$GeneratedWxsPath) {
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
    "office-mcp-env.ps1",
    "office-mcp-install-user.ps1",
    "install-user.ps1",
    "uninstall-user.ps1",
    "start-daemon.ps1",
    "README-install.txt",
    "config.toml",
    "office-mcp-daemon.ps1",
    "office-mcp-tray.ps1",
    "office-mcp.ps1"
  )

  foreach ($relativePath in $requiredPaths) {
    $path = Join-Path $StageRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
      throw "MSI staging payload is missing required path: $relativePath"
    }
  }

  $daemonLauncher = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp-daemon.ps1")
  if ($daemonLauncher -notmatch "office-mcp-env\.ps1" -or $daemonLauncher -notmatch "office-mcp-daemon\.exe" -or $daemonLauncher -notmatch "daemon run") {
    throw "Daemon launcher must use the packaged Rust daemon and run it."
  }

  $cliLauncher = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp.ps1")
  if ($cliLauncher -notmatch "office-mcp-env\.ps1" -or $cliLauncher -notmatch "office-mcp-daemon\.exe" -or $cliLauncher -notmatch "@args") {
    throw "CLI launcher must use the packaged Rust daemon and forward arguments."
  }

  $trayLauncher = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp-tray.ps1")
  if ($trayLauncher -notmatch "office-mcp-daemon\.exe" -or $trayLauncher -notmatch "tray" -or $trayLauncher -notmatch "--probe") {
    throw "Tray launcher must delegate to the native Rust tray host."
  }

  $generatedWxs = Get-Content -Raw -LiteralPath $GeneratedWxsPath
  foreach ($needle in @("office-mcp-daemon.exe", "index.html", "app.js", "taskpane.js", "brand-mark.svg", "icon-32.png", "icon-80.png", "addin-catalog", "office-mcp-env.ps1", "office-mcp-daemon.ps1")) {
    if (-not $generatedWxs.Contains($needle)) {
      throw "Generated WiX payload is missing expected entry: $needle"
    }
  }

  $envScript = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp-env.ps1")
  if ($envScript -notmatch "OFFICE_MCP_INSTALL_ROOT") {
    throw "MSI environment script must set OFFICE_MCP_INSTALL_ROOT."
  }
  if ($envScript -notmatch "OFFICE_MCP_CONFIG_PATH") {
    throw "MSI environment script must set OFFICE_MCP_CONFIG_PATH."
  }
  foreach ($envName in @("OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH", "OFFICE_MCP_ADDIN_CHANNEL__PORT", "OFFICE_MCP_MCP_HTTP__PORT")) {
    if ($envScript -notmatch $envName) {
      throw "MSI environment script must set $envName."
    }
  }

  $installUserScript = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp-install-user.ps1")
  if ($installUserScript -notmatch "ConvertTo-OfficeCatalogUrl" -or $installUserScript -notmatch "\\\\localhost" -or $installUserScript -notmatch "export-localhost-dev-cert\.ps1" -or $installUserScript -notmatch "-CreateIfMissing" -or $installUserScript -notmatch "6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57") {
    throw "MSI user configuration script must register a GUID-based UNC Office catalog and create the localhost certificate."
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
$wxsPath = Join-Path $repoRoot "packaging\wix\Product.wxs"
$outputPath = Join-Path $OutputDir "office-mcp-setup-$Version-x64.msi"
$zipOutputPath = Join-Path $OutputDir "office-mcp-windows-portable-$Version-x64.zip"
$stageRoot = Join-Path $OutputDir "msi-stage"
$generatedWxsPath = Join-Path $OutputDir "msi-payload.wxs"

if (-not (Test-Path $wxsPath)) {
  throw "Missing WiX source file: $wxsPath"
}

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
Copy-Item -Force -Path (Join-Path $repoRoot "packaging\windows\office-mcp-tray.ps1") -Destination (Join-Path $stageRoot "office-mcp-tray.ps1")

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
$pfxPath = Join-Path $installRoot '.office-mcp-localhost.pfx'

$env:OFFICE_MCP_INSTALL_ROOT = $installRoot
$env:OFFICE_MCP_CONFIG_PATH = Join-Path $installRoot 'config.toml'
$env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH = $pfxPath
$env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE = 'office-mcp-localhost'
$env:OFFICE_MCP_ADDIN_CHANNEL__BIND = 'localhost'
$env:OFFICE_MCP_ADDIN_CHANNEL__PORT = '8765'
$env:OFFICE_MCP_MCP_HTTP__BIND = '127.0.0.1'
$env:OFFICE_MCP_MCP_HTTP__PORT = '8800'
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "office-mcp-env.ps1")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $installRoot 'office-mcp-env.ps1')
$daemonExe = Join-Path $installRoot 'office-mcp-daemon.exe'
$pfxPath = Join-Path $installRoot '.office-mcp-localhost.pfx'

if (-not (Test-Path -LiteralPath $pfxPath)) {
  & (Join-Path $installRoot 'scripts\export-localhost-dev-cert.ps1') -OutputPath $pfxPath -CreateIfMissing
}

& $daemonExe daemon run
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "office-mcp-daemon.ps1")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $installRoot 'office-mcp-env.ps1')

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

$catalogPath = Join-Path $installRoot 'addin-catalog'
$catalogUrl = ConvertTo-OfficeCatalogUrl -Path $catalogPath
$catalogId = '{6D178D62-0D2E-4BD6-9F03-5F7FCA34EC57}'
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$legacyCatalogKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp'
if (Test-Path -LiteralPath $legacyCatalogKey) {
  Remove-Item -LiteralPath $legacyCatalogKey -Recurse -Force
}
New-Item -Path $catalogKey -Force | Out-Null
Set-ItemProperty -Path $catalogKey -Name Id -Value $catalogId
Set-ItemProperty -Path $catalogKey -Name Url -Value $catalogUrl
Set-ItemProperty -Path $catalogKey -Name Flags -Value 1 -Type DWord

$pfxPath = Join-Path $installRoot '.office-mcp-localhost.pfx'
if (-not (Test-Path -LiteralPath $pfxPath)) {
  & (Join-Path $installRoot 'scripts\export-localhost-dev-cert.ps1') -OutputPath $pfxPath -CreateIfMissing
}
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "office-mcp-install-user.ps1")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
& (Join-Path $installRoot 'office-mcp-install-user.ps1')
Write-Output "Office MCP Control user install completed."
Write-Output "Install root: $installRoot"
Write-Output "Catalog folder: $(Join-Path $installRoot 'addin-catalog')"
Write-Output "Daemon launcher: $(Join-Path $installRoot 'start-daemon.ps1')"
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "install-user.ps1")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $installRoot 'office-mcp-env.ps1')
$daemonExe = Join-Path $installRoot 'office-mcp-daemon.exe'
& $daemonExe daemon run
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "start-daemon.ps1")

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
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "uninstall-user.ps1")

@"
Office MCP Control portable Windows package

This folder is the install directory. No hidden copy step is required.

Contents:
- office-mcp-daemon.exe: native daemon, tray host, and MCP server.
- office-mcp\ui\: daemon control panel assets.
- office-ctl\word, office-ctl\excel, office-ctl\powerpoint: Office add-in bundles.
- addin-catalog\: Word, Excel, and PowerPoint shared-folder catalog manifests.
- scripts\export-localhost-dev-cert.ps1: creates/exports the localhost HTTPS certificate.
- install-user.ps1: registers the current user's Office trusted catalog and creates the localhost certificate if needed.
- start-daemon.ps1: starts the daemon from this folder.
- uninstall-user.ps1: removes Office MCP Control user registry entries.

What install-user.ps1 changes:
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
   powershell -NoProfile -ExecutionPolicy Bypass -File .\install-user.ps1
4. Start the daemon:
   powershell -NoProfile -ExecutionPolicy Bypass -File .\start-daemon.ps1
5. Reopen Office and use Home > Add-ins > Advanced > Shared Folder to add Office MCP Control if Office does not show it automatically.

Default endpoints:
- MCP: http://127.0.0.1:8800/mcp
- Add-in UI/WSS origin: https://localhost:8765
"@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "README-install.txt")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $installRoot 'office-mcp-env.ps1')
$daemonExe = Join-Path $installRoot 'office-mcp-daemon.exe'
& $daemonExe @args
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "office-mcp.ps1")

New-WixPayloadFragment -StageRoot $stageRoot -OutputPath $generatedWxsPath
Assert-MsiStagePayload -StageRoot $stageRoot -GeneratedWxsPath $generatedWxsPath
New-PortableZip -StageRoot $stageRoot -OutputPath $zipOutputPath

Push-Location $repoRoot
try {
  dotnet tool restore
  Assert-LastExitCode "dotnet tool restore"
  dotnet wix build $wxsPath $generatedWxsPath -arch x64 -d "RepoRoot=$repoRoot" -d "PackageVersion=$Version" -o $outputPath
  Assert-LastExitCode "dotnet wix build"
} finally {
  Pop-Location
}

if (-not (Test-Path $outputPath)) {
  throw "MSI was not created: $outputPath"
}

Write-Output "Built $outputPath"
Write-Output "Built $zipOutputPath"
