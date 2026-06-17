param(
  [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "office-mcp"),
  [string]$ConfigRoot = (Join-Path $env:APPDATA "office-mcp"),
  [string]$TaskName = "office-mcp",
  [switch]$SkipScheduledTask,
  [switch]$SkipCertificateExport
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$rustDaemonRoot = Join-Path $repoRoot "src\office-mcp\daemon"
$evidenceRoot = Join-Path $repoRoot "src\office-mcp\daemon\evidence"
$uiRoot = Join-Path $repoRoot "src\office-mcp\ui"
$commonRoot = Join-Path $repoRoot "src\office-ctl\common"
$addinRoot = Join-Path $repoRoot "src\office-ctl\word"
$excelAddinRoot = Join-Path $repoRoot "src\office-ctl\excel"
$catalogPath = Join-Path $InstallRoot "addin-catalog"
$pfxPath = Join-Path $InstallRoot ".office-mcp-localhost.pfx"

function ConvertTo-OfficeCatalogUrl {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $resolvedPath = [System.IO.Path]::GetFullPath($Path)
  if ($resolvedPath.StartsWith("\\", [System.StringComparison]::Ordinal)) {
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

if (-not (Test-Path (Join-Path $addinRoot "manifest.xml"))) {
  throw "Cannot find src/office-ctl/word/manifest.xml under $repoRoot."
}
if (-not (Test-Path (Join-Path $excelAddinRoot "manifest.xml"))) {
  throw "Cannot find src/office-ctl/excel/manifest.xml under $repoRoot."
}
if (-not (Test-Path (Join-Path $rustDaemonRoot "Cargo.toml"))) {
  throw "Cannot find src/office-mcp/daemon/Cargo.toml under $repoRoot."
}
if (-not (Test-Path (Join-Path $evidenceRoot "package.json"))) {
  throw "Cannot find src/office-mcp/daemon/evidence/package.json under $repoRoot."
}

New-Item -ItemType Directory -Force -Path $InstallRoot, $ConfigRoot, $catalogPath | Out-Null

Push-Location $repoRoot
try {
  cargo build --release -p office-mcp-daemon
} finally {
  Pop-Location
}

Push-Location $evidenceRoot
try {
  npm ci
  npm run check
} finally {
  Pop-Location
}

Push-Location $addinRoot
try {
  npm ci
  npm run check
} finally {
  Pop-Location
}

Push-Location $excelAddinRoot
try {
  npm ci
  npm run check
} finally {
  Pop-Location
}

Copy-Item -Force -Path (Join-Path $addinRoot "manifest.xml") -Destination (Join-Path $catalogPath "office-mcp-word.xml")
Copy-Item -Force -Path (Join-Path $excelAddinRoot "manifest.xml") -Destination (Join-Path $catalogPath "office-mcp-excel.xml")
Copy-Item -Force -Path (Join-Path $repoRoot "target\release\office-mcp-daemon.exe") -Destination (Join-Path $InstallRoot "office-mcp-daemon.exe")
$installedUiRoot = Join-Path $InstallRoot "office-mcp\ui"
$installedCommonRoot = Join-Path $InstallRoot "office-ctl\common"
$installedWordRoot = Join-Path $InstallRoot "office-ctl\word"
$installedExcelRoot = Join-Path $InstallRoot "office-ctl\excel"
New-Item -ItemType Directory -Force -Path $installedUiRoot, $installedCommonRoot, $installedWordRoot, $installedExcelRoot | Out-Null
Copy-Item -Recurse -Force -Path (Join-Path $uiRoot "*") -Destination $installedUiRoot
Copy-Item -Recurse -Force -Path (Join-Path $commonRoot "*") -Destination $installedCommonRoot
Copy-Item -Force -Path (Join-Path $addinRoot "manifest.xml") -Destination $installedWordRoot
Copy-Item -Recurse -Force -Path (Join-Path $addinRoot "public") -Destination $installedWordRoot
Copy-Item -Force -Path (Join-Path $excelAddinRoot "manifest.xml") -Destination $installedExcelRoot
Copy-Item -Recurse -Force -Path (Join-Path $excelAddinRoot "public") -Destination $installedExcelRoot
$trayLauncherPath = Join-Path $InstallRoot "office-mcp-tray.ps1"
Copy-Item -Force -Path (Join-Path $repoRoot "packaging\windows\office-mcp-tray.ps1") -Destination $trayLauncherPath

$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp"
$catalogUrl = ConvertTo-OfficeCatalogUrl -Path $catalogPath
New-Item -Path $catalogKey -Force | Out-Null
Set-ItemProperty -Path $catalogKey -Name Id -Value "office-mcp"
Set-ItemProperty -Path $catalogKey -Name Url -Value $catalogUrl
Set-ItemProperty -Path $catalogKey -Name Flags -Value 1 -Type DWord

if (-not $SkipCertificateExport) {
  & (Join-Path $repoRoot "packaging\windows\export-localhost-dev-cert.ps1") -OutputPath $pfxPath
}

$configPath = Join-Path $ConfigRoot "config.env.ps1"
@"
`$env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH = '$pfxPath'
`$env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE = 'office-mcp-localhost'
`$env:OFFICE_MCP_ADDIN_CHANNEL__BIND = 'localhost'
`$env:OFFICE_MCP_ADDIN_CHANNEL__PORT = '8765'
`$env:OFFICE_MCP_MCP_HTTP__BIND = '127.0.0.1'
`$env:OFFICE_MCP_MCP_HTTP__PORT = '8800'
"@ | Set-Content -Encoding ASCII -Path $configPath

$launcherPath = Join-Path $InstallRoot "office-mcp-daemon.ps1"
@"
`$ErrorActionPreference = 'Stop'
. '$configPath'
& (Join-Path '$InstallRoot' 'office-mcp-daemon.exe') daemon run
"@ | Set-Content -Encoding ASCII -Path $launcherPath

if (-not $SkipScheduledTask) {
  $action = New-ScheduledTaskAction -Execute (Join-Path $InstallRoot "office-mcp-daemon.exe") -Argument "tray"
  $trigger = New-ScheduledTaskTrigger -AtLogOn
  $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel LeastPrivilege
  $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
}

Write-Output "Installed office-mcp developer bootstrap."
Write-Output "Install root: $InstallRoot"
Write-Output "Config script: $configPath"
Write-Output "Catalog URL: $catalogUrl"
Write-Output "Daemon launcher: $launcherPath"
Write-Output "Tray launcher: $(Join-Path $InstallRoot "office-mcp-daemon.exe") tray"
if (-not $SkipScheduledTask) { Write-Output "Scheduled task: $TaskName" }
