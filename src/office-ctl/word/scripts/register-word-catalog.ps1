param(
  [string]$CatalogPath = "",
  [string]$BaseUrl = "https://localhost:8765"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$arguments = @{
  RepoRoot = $repoRoot
}

if (-not [string]::IsNullOrWhiteSpace($CatalogPath)) {
  $arguments.CatalogPath = $CatalogPath
}
$arguments.BaseUrl = $BaseUrl

& (Join-Path $repoRoot "src\office-ctl\common\scripts\register-office-catalog.ps1") @arguments
