$ErrorActionPreference = "Stop"

$repo = "r12f/office-mcp"
$assetPattern = "office-mcp-windows-portable-*-x64.zip"

$release = Invoke-RestMethod "https://api.github.com/repos/$repo/releases" |
  Where-Object { $_.assets.name -like $assetPattern } |
  Select-Object -First 1

if (-not $release) {
  throw "No Windows portable release asset found for $repo."
}

$asset = $release.assets |
  Where-Object { $_.name -like $assetPattern } |
  Select-Object -First 1

if (-not $asset) {
  throw "Release $($release.tag_name) does not contain $assetPattern."
}

$installRoot = Join-Path $env:LOCALAPPDATA ("office-mcp\" + $release.tag_name)
$zipPath = Join-Path $env:TEMP $asset.name

Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath
New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
Expand-Archive -LiteralPath $zipPath -DestinationPath $installRoot -Force
& (Join-Path $installRoot "install.ps1")
