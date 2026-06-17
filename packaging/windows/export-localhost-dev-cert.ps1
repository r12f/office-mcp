param(
  [string]$OutputPath = (Join-Path (Get-Location) ".office-mcp-localhost.pfx"),
  [string]$Password = "office-mcp-localhost"
)

$ErrorActionPreference = "Stop"

$cert = Get-ChildItem Cert:\CurrentUser\My |
  Where-Object { $_.Subject -eq "CN=localhost" -and $_.HasPrivateKey } |
  Sort-Object NotAfter -Descending |
  Select-Object -First 1

if (-not $cert) {
  throw "No CurrentUser\My localhost certificate with a private key was found. Run 'dotnet dev-certs https' first."
}

$trusted = Get-ChildItem Cert:\CurrentUser\Root |
  Where-Object { $_.Thumbprint -eq $cert.Thumbprint } |
  Select-Object -First 1

if (-not $trusted) {
  throw "The localhost certificate is not trusted in CurrentUser\Root. Import it only after explicit approval."
}

$secure = ConvertTo-SecureString -String $Password -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath $OutputPath -Password $secure | Out-Null

Write-Output "Exported $OutputPath"
Write-Output "Set OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH=$OutputPath if you start the daemon outside this repo."
