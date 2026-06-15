param(
  [Parameter(Mandatory = $true)]
  [string]$TarballUrl,

  [Parameter(Mandatory = $true)]
  [string]$TarballSha256,

  [string]$TemplatePath = "",
  [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

if (-not $TemplatePath) {
  $TemplatePath = Join-Path $PSScriptRoot "Formula\office-mcp.rb.in"
}
if (-not $OutputPath) {
  $OutputPath = Join-Path $PSScriptRoot "Formula\office-mcp.rb"
}

if ($TarballUrl -notmatch '^https://') {
  throw "TarballUrl must start with https://"
}
if ($TarballSha256 -notmatch '^[0-9a-f]{64}$') {
  throw "TarballSha256 must be a lowercase 64-character SHA-256 hex digest."
}
if (-not (Test-Path -LiteralPath $TemplatePath)) {
  throw "Missing formula template: $TemplatePath"
}

$content = Get-Content -Raw -LiteralPath $TemplatePath
$content = $content.Replace('{{TARBALL_URL}}', $TarballUrl).Replace('{{TARBALL_SHA256}}', $TarballSha256)

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null
$content | Set-Content -Encoding ASCII -Path $OutputPath
Write-Output "Rendered $OutputPath"
