param(
  [string]$SourceManifest = (Join-Path (Split-Path -Parent $PSScriptRoot) "manifest.xml"),
  [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "dist\manifest-hosted.xml"),
  [string]$BaseUrl = "https://office-mcp.dev",
  [string]$AddinId = "44444444-aaaa-bbbb-cccc-555555555555",
  [string]$AddinVersion = "1.0.0.4",
  [string]$AssetVersion = "0.1.4",
  [string]$SupportUrl = "https://github.com/office-mcp/office-mcp"
)

$ErrorActionPreference = "Stop"

& (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) "common\scripts\render-hosted-manifest.ps1") `
  -SourceManifest $SourceManifest `
  -OutputPath $OutputPath `
  -BaseUrl $BaseUrl `
  -TaskpanePath "/powerpoint/taskpane.html" `
  -AddinId $AddinId `
  -AddinVersion $AddinVersion `
  -AssetVersion $AssetVersion `
  -SupportUrl $SupportUrl
