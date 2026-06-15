param(
  [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "office-mcp"),
  [string]$ConfigRoot = (Join-Path $env:APPDATA "office-mcp"),
  [string]$TaskName = "office-mcp",
  [switch]$PurgeConfig
)

$ErrorActionPreference = "Stop"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp"
if (Test-Path $catalogKey) {
  Remove-Item -Path $catalogKey -Recurse -Force
}

if (Test-Path $InstallRoot) {
  $resolved = (Resolve-Path -LiteralPath $InstallRoot).Path
  $localAppData = (Resolve-Path -LiteralPath $env:LOCALAPPDATA).Path
  if (-not $resolved.StartsWith($localAppData, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to remove install root outside LOCALAPPDATA: $resolved"
  }
  Remove-Item -LiteralPath $resolved -Recurse -Force
}

if ($PurgeConfig -and (Test-Path $ConfigRoot)) {
  $resolvedConfig = (Resolve-Path -LiteralPath $ConfigRoot).Path
  $appData = (Resolve-Path -LiteralPath $env:APPDATA).Path
  if (-not $resolvedConfig.StartsWith($appData, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to remove config root outside APPDATA: $resolvedConfig"
  }
  Remove-Item -LiteralPath $resolvedConfig -Recurse -Force
}

Write-Output "Uninstalled office-mcp developer bootstrap."
Write-Output "Removed scheduled task: $TaskName"
Write-Output "Removed trusted catalog registry key: $catalogKey"
Write-Output "Removed install root: $InstallRoot"
if ($PurgeConfig) { Write-Output "Removed config root: $ConfigRoot" }
