param(
  [string]$Version = "0.1.0",
  [string]$BaseUrl = "https://office-mcp.dev",
  [string]$AddinId = "11111111-aaaa-bbbb-cccc-222222222222",
  [string]$AddinVersion = "1.0.0.4",
  [string]$SupportUrl = "https://github.com/office-mcp/office-mcp",
  [string]$OutputDir = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))) "artifacts\appsource")
)

$ErrorActionPreference = "Stop"

function Assert-File([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw "Missing required file: $Path"
  }
}

function Get-Sha256([string]$Path) {
  $stream = [System.IO.File]::OpenRead((Resolve-Path -LiteralPath $Path).Path)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    return ([System.BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
  } finally {
    $sha.Dispose()
    $stream.Dispose()
  }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$addinRoot = Split-Path -Parent $PSScriptRoot
$publicRoot = Join-Path $addinRoot "public"
$assetRoot = Join-Path (Split-Path -Parent $addinRoot) "common\assets"
$manifestPath = Join-Path $OutputDir "manifest-$Version.xml"
$bundlePath = Join-Path $OutputDir "office-mcp-addin-$Version.zip"
$bundleStagePath = Join-Path $OutputDir "office-mcp-addin-bundle-$Version"
$checklistPath = Join-Path $OutputDir "appsource-checklist-$Version.md"
$metadataPath = Join-Path $OutputDir "appsource-metadata-$Version.json"
$packagePath = Join-Path $OutputDir "office-mcp-appsource-$Version.zip"

Assert-File (Join-Path $addinRoot "manifest.xml")
Assert-File (Join-Path $publicRoot "taskpane.html")
Assert-File (Join-Path $publicRoot "taskpane.css")
Assert-File (Join-Path $publicRoot "taskpane.js")
Assert-File (Join-Path $assetRoot "brand-mark.svg")
Assert-File (Join-Path $assetRoot "icon-32.png")
Assert-File (Join-Path $assetRoot "icon-80.png")

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

& (Join-Path $PSScriptRoot "render-hosted-manifest.ps1") `
  -SourceManifest (Join-Path $addinRoot "manifest.xml") `
  -OutputPath $manifestPath `
  -BaseUrl $BaseUrl `
  -AddinId $AddinId `
  -AddinVersion $AddinVersion `
  -AssetVersion $Version `
  -SupportUrl $SupportUrl | Out-Null

$manifestXml = Get-Content -Raw -LiteralPath $manifestPath
if ($manifestXml -match 'localhost|127\.0\.0\.1') {
  throw "AppSource manifest contains loopback URLs."
}
if ($manifestXml -notmatch [regex]::Escape($BaseUrl.TrimEnd('/'))) {
  throw "AppSource manifest does not reference $BaseUrl."
}

if (Test-Path -LiteralPath $bundlePath) { Remove-Item -LiteralPath $bundlePath -Force }
if (Test-Path -LiteralPath $bundleStagePath) { Remove-Item -LiteralPath $bundleStagePath -Recurse -Force }
New-Item -ItemType Directory -Force -Path $bundleStagePath | Out-Null
Copy-Item -Recurse -Force -Path (Join-Path $publicRoot "*") -Destination $bundleStagePath
Copy-Item -Recurse -Force -Path $assetRoot -Destination (Join-Path $bundleStagePath "assets")
Compress-Archive -Path (Join-Path $bundleStagePath "*") -DestinationPath $bundlePath -CompressionLevel Optimal
Remove-Item -LiteralPath $bundleStagePath -Recurse -Force

$metadata = [ordered]@{
  name = "Office MCP Control"
  slug = "office-mcp"
  category = "Productivity"
  type = "Local productivity automation control utility"
  version = $Version
  manifest = Split-Path -Leaf $manifestPath
  manifest_sha256 = Get-Sha256 $manifestPath
  addin_bundle = Split-Path -Leaf $bundlePath
  addin_bundle_sha256 = Get-Sha256 $bundlePath
  public_base_url = $BaseUrl.TrimEnd('/')
  support_url = $SupportUrl
  notes = @(
    "Partner Center account and AppSource validation review are external gates.",
    "Office webview access from the hosted HTTPS task pane to the local daemon must be validated before submission."
  )
}
$metadata | ConvertTo-Json -Depth 4 | Set-Content -Encoding ASCII -Path $metadataPath

@"
# office-mcp AppSource submission checklist

Generated for version `$Version`.

## Included artifacts

- `$([System.IO.Path]::GetFileName($manifestPath))`
- `$([System.IO.Path]::GetFileName($bundlePath))`
- `$([System.IO.Path]::GetFileName($metadataPath))`

## Automated gates completed

- Hosted manifest rendered from `src/office-ctl/word/manifest.xml`.
- Manifest uses `$($BaseUrl.TrimEnd('/'))` and contains no loopback URLs.
- Add-in static bundle includes `taskpane.html`, `taskpane.css`, `taskpane.js`, and generated `assets/*` product icons.
- SHA-256 digests are recorded in `$([System.IO.Path]::GetFileName($metadataPath))`.

## External gates before Partner Center submission

- Host `$($BaseUrl.TrimEnd('/'))/taskpane.html` and `$($BaseUrl.TrimEnd('/'))/assets/*` over public HTTPS.
- Host the manifest at `$($BaseUrl.TrimEnd('/'))/manifest.xml` or the chosen Partner Center URL.
- Validate Office webview behavior from the hosted origin to the local daemon endpoint.
- Provide Partner Center listing metadata with category `Productivity` and type `Local productivity automation control utility`.
- Provide screenshots, privacy/support URLs, and reviewer notes.
- Complete Microsoft AppSource validation review.
"@ | Set-Content -Encoding ASCII -Path $checklistPath

if (Test-Path -LiteralPath $packagePath) { Remove-Item -LiteralPath $packagePath -Force }
Compress-Archive -Path $manifestPath, $bundlePath, $metadataPath, $checklistPath -DestinationPath $packagePath -CompressionLevel Optimal

Write-Output "Built AppSource package: $packagePath"
