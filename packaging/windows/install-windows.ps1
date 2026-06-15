param(
  [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "office-mcp"),
  [string]$ConfigRoot = (Join-Path $env:APPDATA "office-mcp"),
  [string]$TaskName = "office-mcp",
  [switch]$SkipScheduledTask,
  [switch]$SkipCertificateExport
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$serverRoot = Join-Path $repoRoot "mcp-server"
$addinRoot = Join-Path $repoRoot "addin"
$catalogPath = Join-Path $InstallRoot "addin-catalog"
$pfxPath = Join-Path $InstallRoot ".office-mcp-localhost.pfx"

if (-not (Test-Path (Join-Path $serverRoot "package.json"))) {
  throw "Cannot find mcp-server/package.json under $repoRoot. Run this script from the checked-out office-mcp repo."
}
if (-not (Test-Path (Join-Path $addinRoot "manifest.xml"))) {
  throw "Cannot find addin/manifest.xml under $repoRoot."
}

New-Item -ItemType Directory -Force -Path $InstallRoot, $ConfigRoot, $catalogPath | Out-Null

Push-Location $serverRoot
try {
  npm ci
  npm run build
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

Copy-Item -Force -Path (Join-Path $addinRoot "manifest.xml") -Destination (Join-Path $catalogPath "manifest.xml")

$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp"
New-Item -Path $catalogKey -Force | Out-Null
Set-ItemProperty -Path $catalogKey -Name Id -Value "office-mcp"
Set-ItemProperty -Path $catalogKey -Name Url -Value $catalogPath
Set-ItemProperty -Path $catalogKey -Name Flags -Value 1 -Type DWord

if (-not $SkipCertificateExport) {
  & (Join-Path $serverRoot "scripts\export-localhost-dev-cert.ps1") -OutputPath $pfxPath
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
Set-Location '$serverRoot'
npm run daemon
"@ | Set-Content -Encoding ASCII -Path $launcherPath

if (-not $SkipScheduledTask) {
  $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$launcherPath`""
  $trigger = New-ScheduledTaskTrigger -AtLogOn
  $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel LeastPrivilege
  $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
}

Write-Output "Installed office-mcp developer bootstrap."
Write-Output "Install root: $InstallRoot"
Write-Output "Config script: $configPath"
Write-Output "Catalog URL: $catalogPath"
Write-Output "Daemon launcher: $launcherPath"
if (-not $SkipScheduledTask) { Write-Output "Scheduled task: $TaskName" }
