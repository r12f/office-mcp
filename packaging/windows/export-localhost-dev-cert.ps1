param(
  [string]$OutputPath = (Join-Path (Get-Location) ".office-mcp-localhost.pfx"),
  [string]$Password = "office-mcp-localhost",
  [switch]$CreateIfMissing
)

$ErrorActionPreference = "Stop"

$cert = Get-ChildItem Cert:\CurrentUser\My |
  Where-Object { $_.Subject -eq "CN=localhost" -and $_.HasPrivateKey } |
  Sort-Object NotAfter -Descending |
  Select-Object -First 1

if (-not $cert) {
  if (-not $CreateIfMissing) {
    throw "No CurrentUser\My localhost certificate with a private key was found. Run this helper with -CreateIfMissing to create one."
  }

  $cert = New-SelfSignedCertificate `
    -DnsName "localhost" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears(2)
}

$trusted = Get-ChildItem Cert:\CurrentUser\Root |
  Where-Object { $_.Thumbprint -eq $cert.Thumbprint } |
  Select-Object -First 1

if (-not $trusted) {
  if (-not $CreateIfMissing) {
    throw "The localhost certificate is not trusted in CurrentUser\Root. Run this helper with -CreateIfMissing to trust it for the current user."
  }

  $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
  try {
    $rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $rootStore.Add($cert)
  } finally {
    $rootStore.Close()
  }
}

$secure = ConvertTo-SecureString -String $Password -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath $OutputPath -Password $secure | Out-Null

Write-Output "Exported $OutputPath"
Write-Output "Set OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH=$OutputPath if you start the daemon outside this repo."
