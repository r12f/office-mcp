param(
  [Parameter(Mandatory = $true)][string]$InstallRoot,
  [Parameter(Mandatory = $true)][string]$ConfigPath,
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

function Assert-AbsoluteUnixPath([string]$Name, [string]$Value) {
  if (-not $Value.StartsWith('/')) {
    throw "$Name must be an absolute Unix path."
  }
  if ($Value -match '[\r\n\0]') {
    throw "$Name contains an invalid control character."
  }
}

Assert-AbsoluteUnixPath "InstallRoot" $InstallRoot
Assert-AbsoluteUnixPath "ConfigPath" $ConfigPath

if (-not $OutputPath) {
  $OutputPath = Join-Path $PSScriptRoot "office-mcp.service"
}

$templatePath = Join-Path $PSScriptRoot "office-mcp.service.in"
$content = Get-Content -Raw -LiteralPath $templatePath
$content = $content.Replace('{{INSTALL_ROOT}}', $InstallRoot.TrimEnd('/'))
$content = $content.Replace('{{CONFIG_PATH}}', $ConfigPath)

$outputDir = Split-Path -Parent $OutputPath
if ($outputDir) {
  New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}
$content | Set-Content -Encoding UTF8 -Path $OutputPath
Write-Host "Rendered $OutputPath"
