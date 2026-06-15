param(
  [string]$CatalogPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "addin-catalog")
)

$ErrorActionPreference = "Stop"

New-Item -ItemType Directory -Force -Path $CatalogPath | Out-Null
$manifestPath = Join-Path (Split-Path -Parent $PSScriptRoot) "manifest.xml"
Copy-Item -Force -Path $manifestPath -Destination (Join-Path $CatalogPath "manifest.xml")

$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp"
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name Id -Value "office-mcp"
Set-ItemProperty -Path $key -Name Url -Value $CatalogPath
Set-ItemProperty -Path $key -Name Flags -Value 1 -Type DWord

Write-Output "Registered Word trusted catalog: $CatalogPath"
Write-Output "Catalog URL: $CatalogPath"
