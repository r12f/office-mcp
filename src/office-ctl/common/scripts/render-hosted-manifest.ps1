param(
  [Parameter(Mandatory = $true)][string]$SourceManifest,
  [Parameter(Mandatory = $true)][string]$OutputPath,
  [string]$BaseUrl = "https://office-mcp.dev",
  [Parameter(Mandatory = $true)][string]$TaskpanePath,
  [Parameter(Mandatory = $true)][string]$AddinId,
  [Parameter(Mandatory = $true)][string]$AddinVersion,
  [Parameter(Mandatory = $true)][string]$AssetVersion,
  [string]$SupportUrl = "https://github.com/office-mcp/office-mcp"
)

$ErrorActionPreference = "Stop"

if ($BaseUrl -notmatch '^https://') {
  throw "BaseUrl must start with https://"
}
if ($BaseUrl -match 'localhost|127\.0\.0\.1') {
  throw "Hosted manifests must not use a loopback BaseUrl."
}
if ($TaskpanePath -notmatch '^/') {
  throw "TaskpanePath must start with /."
}
if ($AddinId -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
  throw "AddinId must be a GUID."
}
if ($AddinVersion -notmatch '^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$') {
  throw "AddinVersion must be a four-part Office manifest version."
}
if ($SupportUrl -notmatch '^https://') {
  throw "SupportUrl must start with https://"
}
if (-not (Test-Path -LiteralPath $SourceManifest)) {
  throw "Missing source manifest: $SourceManifest"
}

$base = $BaseUrl.TrimEnd('/')
$taskpaneUrl = "$base$TaskpanePath`?v=$AssetVersion"
$icon16Url = "$base/assets/icon-16.png"
$icon32Url = "$base/assets/icon-32.png"
$icon80Url = "$base/assets/icon-80.png"

[xml]$manifest = Get-Content -Raw -LiteralPath $SourceManifest
$ns = New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
$ns.AddNamespace("o", "http://schemas.microsoft.com/office/appforoffice/1.1")
$ns.AddNamespace("bt", "http://schemas.microsoft.com/office/officeappbasictypes/1.0")

function Set-NodeText([xml]$Document, [System.Xml.XmlNamespaceManager]$NamespaceManager, [string]$XPath, [string]$Value) {
  $node = $Document.SelectSingleNode($XPath, $NamespaceManager)
  if (-not $node) { throw "Manifest node not found: $XPath" }
  $node.InnerText = $Value
}

function Set-DefaultValue([xml]$Document, [System.Xml.XmlNamespaceManager]$NamespaceManager, [string]$XPath, [string]$Value) {
  $node = $Document.SelectSingleNode($XPath, $NamespaceManager)
  if (-not $node) { throw "Manifest node not found: $XPath" }
  $node.SetAttribute("DefaultValue", $Value)
}

Set-NodeText $manifest $ns "/o:OfficeApp/o:Id" $AddinId
Set-NodeText $manifest $ns "/o:OfficeApp/o:Version" $AddinVersion
Set-NodeText $manifest $ns "/o:OfficeApp/o:AppDomains/o:AppDomain" $base

Set-DefaultValue $manifest $ns "/o:OfficeApp/o:IconUrl" $icon32Url
Set-DefaultValue $manifest $ns "/o:OfficeApp/o:HighResolutionIconUrl" $icon80Url
Set-DefaultValue $manifest $ns "/o:OfficeApp/o:SupportUrl" $SupportUrl
Set-DefaultValue $manifest $ns "/o:OfficeApp/o:DefaultSettings/o:SourceLocation" $taskpaneUrl

Set-DefaultValue $manifest $ns "//bt:Image[@id='Icon.16x16']" $icon16Url
Set-DefaultValue $manifest $ns "//bt:Image[@id='Icon.32x32']" $icon32Url
Set-DefaultValue $manifest $ns "//bt:Image[@id='Icon.80x80']" $icon80Url
Set-DefaultValue $manifest $ns "//bt:Url[@id='Taskpane.Url']" $taskpaneUrl
Set-DefaultValue $manifest $ns "//bt:Url[@id='GetStarted.LearnMoreUrl']" $SupportUrl

$xml = $manifest.OuterXml
if ($xml -match 'localhost|127\.0\.0\.1') {
  throw "Hosted manifest still contains a loopback URL."
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputPath) | Out-Null
$settings = New-Object System.Xml.XmlWriterSettings
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$settings.Indent = $true
$settings.OmitXmlDeclaration = $false
$writer = [System.Xml.XmlWriter]::Create($OutputPath, $settings)
try {
  $manifest.Save($writer)
} finally {
  $writer.Dispose()
}

Write-Output "Rendered hosted manifest: $OutputPath"
