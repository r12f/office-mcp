$ErrorActionPreference = "Stop"

$repo = "r12f/office-mcp"
$assetPattern = "office-mcp-windows-portable-*-x64.zip"
$headers = @{ "User-Agent" = "office-mcp-installer" }

$releaseJson = (Invoke-WebRequest -Headers $headers -Uri "https://api.github.com/repos/$repo/releases?per_page=100").Content
$releases = @($releaseJson | ConvertFrom-Json)
$release = $releases |
  Where-Object {
    -not $_.draft -and
    @(($_.assets | Where-Object { $_.name -like $assetPattern })).Count -gt 0
  } |
  Sort-Object { [datetime]$_.published_at } -Descending |
  Select-Object -First 1

if (-not $release) {
  throw "No published Windows portable release asset found for $repo. Publish a non-draft release containing $assetPattern."
}

$asset = $release.assets |
  Where-Object { $_.name -like $assetPattern } |
  Select-Object -First 1

if (-not $asset) {
  throw "Release $($release.tag_name) does not contain $assetPattern."
}

$installRoot = Join-Path $env:LOCALAPPDATA ("office-mcp\" + $release.tag_name)
$zipPath = Join-Path $env:TEMP $asset.name

Write-Host "Installing Office MCP Control $($release.tag_name) to $installRoot"
Invoke-WebRequest -Headers $headers -Uri $asset.browser_download_url -OutFile $zipPath
New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
Expand-Archive -LiteralPath $zipPath -DestinationPath $installRoot -Force
& (Join-Path $installRoot "install.ps1")
