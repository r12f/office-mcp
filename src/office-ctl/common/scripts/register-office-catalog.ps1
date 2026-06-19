param(
  [string]$CatalogPath = "",
  [string]$RepoRoot = "",
  [string]$BaseUrl = "https://localhost:8765",
  [string]$DaemonStatusCommand = "",
  [string]$TrustedCatalogRegistryKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\office-mcp",
  [switch]$ClearOfficeCache,
  [switch]$SkipRegistry
)

$ErrorActionPreference = "Stop"

function Get-ParentPath {
  param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][int]$Levels
  )

  $current = $Path
  for ($i = 0; $i -lt $Levels; $i++) {
    $current = Split-Path -Parent $current
  }
  return $current
}

if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
  $RepoRoot = Get-ParentPath -Path $PSScriptRoot -Levels 4
}

if ([string]::IsNullOrWhiteSpace($CatalogPath)) {
  $CatalogPath = Join-Path $RepoRoot "addin-catalog"
}

if (-not [string]::IsNullOrWhiteSpace($DaemonStatusCommand)) {
  $statusJson = & $DaemonStatusCommand daemon status
  if ($LASTEXITCODE -ne 0) { throw "Daemon status command failed: $DaemonStatusCommand daemon status" }
  $status = $statusJson | ConvertFrom-Json
  if (-not $status.uiUrl) { throw "Daemon status output does not include uiUrl." }
  $origin = ([Uri]$status.uiUrl).GetLeftPart([System.UriPartial]::Authority)
  if ($origin -notmatch '^https://localhost:[0-9]+$') {
    throw "Daemon status uiUrl must use a local HTTPS localhost origin: $($status.uiUrl)"
  }
  $BaseUrl = $origin
}

if ($BaseUrl -notmatch '^https://localhost:[0-9]+$') {
  throw "BaseUrl must be a local HTTPS origin such as https://localhost:8765."
}

function Set-DefaultValue([xml]$Document, [System.Xml.XmlNamespaceManager]$NamespaceManager, [string]$XPath, [string]$Value) {
  $node = $Document.SelectSingleNode($XPath, $NamespaceManager)
  if (-not $node) { throw "Manifest node not found: $XPath" }
  $node.SetAttribute("DefaultValue", $Value)
}

function Write-LocalManifest {
  param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$DestinationPath,
    [Parameter(Mandatory = $true)][string]$Origin
  )

  [xml]$manifest = Get-Content -Raw -LiteralPath $SourcePath
  $ns = New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
  $ns.AddNamespace("o", "http://schemas.microsoft.com/office/appforoffice/1.1")
  $ns.AddNamespace("bt", "http://schemas.microsoft.com/office/officeappbasictypes/1.0")

  $appDomain = $manifest.SelectSingleNode("/o:OfficeApp/o:AppDomains/o:AppDomain", $ns)
  if (-not $appDomain) { throw "Manifest node not found: /o:OfficeApp/o:AppDomains/o:AppDomain" }
  $appDomain.InnerText = $Origin

  foreach ($xpath in @(
      "/o:OfficeApp/o:IconUrl",
      "//bt:Image[@id='Icon.32x32']"
    )) {
    Set-DefaultValue $manifest $ns $xpath "$Origin/assets/icon-32.png"
  }
  Set-DefaultValue $manifest $ns "//bt:Image[@id='Icon.16x16']" "$Origin/assets/icon-16.png"
  foreach ($xpath in @(
      "/o:OfficeApp/o:HighResolutionIconUrl",
      "//bt:Image[@id='Icon.80x80']"
    )) {
    Set-DefaultValue $manifest $ns $xpath "$Origin/assets/icon-80.png"
  }

  $taskpane = $manifest.SelectSingleNode("/o:OfficeApp/o:DefaultSettings/o:SourceLocation", $ns).GetAttribute("DefaultValue")
  $relativeTaskpane = ([uri]$taskpane).PathAndQuery.TrimStart('/')
  Set-DefaultValue $manifest $ns "/o:OfficeApp/o:DefaultSettings/o:SourceLocation" "$Origin/$relativeTaskpane"
  Set-DefaultValue $manifest $ns "//bt:Url[@id='Taskpane.Url']" "$Origin/$relativeTaskpane"

  $settings = New-Object System.Xml.XmlWriterSettings
  $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
  $settings.Indent = $true
  $settings.OmitXmlDeclaration = $false
  $writer = [System.Xml.XmlWriter]::Create($DestinationPath, $settings)
  try {
    $manifest.Save($writer)
  } finally {
    $writer.Dispose()
  }
}

function Remove-LegacyHostCatalogFolders {
  param(
    [Parameter(Mandatory = $true)][string]$CatalogRoot
  )

  $resolvedCatalogRoot = [System.IO.Path]::GetFullPath($CatalogRoot)
  foreach ($legacyHost in @("word", "excel", "powerpoint")) {
    $legacyPath = Join-Path $resolvedCatalogRoot $legacyHost
    if (-not (Test-Path -LiteralPath $legacyPath)) { continue }

    $resolvedLegacyPath = [System.IO.Path]::GetFullPath($legacyPath)
    $catalogPrefix = $resolvedCatalogRoot.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $resolvedLegacyPath.StartsWith($catalogPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
      throw "Refusing to remove legacy catalog folder outside catalog root: $resolvedLegacyPath"
    }
    Remove-Item -LiteralPath $resolvedLegacyPath -Recurse -Force
  }
}
function Remove-DeveloperDebugRegistration {
  param(
    [Parameter(Mandatory = $true)][string]$AddinId
  )

  $developerRoot = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
  $developerValue = Get-ItemProperty -LiteralPath $developerRoot -Name $AddinId -ErrorAction SilentlyContinue
  if ($developerValue) {
    Remove-ItemProperty -LiteralPath $developerRoot -Name $AddinId -Force
  }

  $developerKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\$AddinId"
  if (Test-Path -LiteralPath $developerKey) {
    Remove-Item -LiteralPath $developerKey -Recurse -Force
  }
}

function Assert-TrustedCatalogRegistryKey {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  if ($Path -notmatch '^HKCU:\\') {
    throw "TrustedCatalogRegistryKey must be under HKCU."
  }
  if ($Path -match '[\r\n]') {
    throw "TrustedCatalogRegistryKey must be a single registry path."
  }
}

function Assert-OfficeHostsClosed {
  param(
    [Parameter(Mandatory = $true)][array]$Hosts
  )

  $running = @()
  foreach ($officeHost in $Hosts) {
    $processName = $officeHost.ProcessName
    if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
      $running += $processName
    }
  }
  if ($running.Count -gt 0) {
    throw "Close Office host processes before clearing add-in cache: $($running -join ', ')."
  }
}

function Remove-OfficeAddinCache {
  param(
    [Parameter(Mandatory = $true)][array]$Hosts
  )

  $wefRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
  if (-not (Test-Path -LiteralPath $wefRoot)) { return }

  foreach ($officeHost in $Hosts) {
    $addinId = $officeHost.AddinId
    $hostName = $officeHost.Name

    $addinInfoRoot = Join-Path $wefRoot "AddinInfo\1\filesystem\$hostName\1"
    if (Test-Path -LiteralPath $addinInfoRoot) {
      Get-ChildItem -LiteralPath $addinInfoRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name.StartsWith($addinId, [System.StringComparison]::OrdinalIgnoreCase) } |
        Remove-Item -Recurse -Force
    }

    Get-ChildItem -LiteralPath $wefRoot -Recurse -File -ErrorAction SilentlyContinue |
      Where-Object { $_.FullName -match '\\Manifests\\' -and $_.Name.StartsWith($addinId, [System.StringComparison]::OrdinalIgnoreCase) } |
      Remove-Item -Force

    $appCommandsRoot = Join-Path $wefRoot "AppCommands"
    if (Test-Path -LiteralPath $appCommandsRoot) {
      Get-ChildItem -LiteralPath $appCommandsRoot -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object {
          ($_.FullName -match '\\TrustedCatalog\\' -and $_.Name.StartsWith($addinId, [System.StringComparison]::OrdinalIgnoreCase)) -or
          ($_.Name -like "$hostName.RibbonCache.*")
        } |
        Remove-Item -Force
    }
  }
}

function ConvertTo-OfficeCatalogUrl {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  $resolvedPath = [System.IO.Path]::GetFullPath($Path)
  if ($resolvedPath.StartsWith("\\", [System.StringComparison]::Ordinal)) {
    return $resolvedPath
  }

  $root = [System.IO.Path]::GetPathRoot($resolvedPath)
  if ([string]::IsNullOrWhiteSpace($root) -or $root.Length -lt 2 -or $root[1] -ne ':') {
    throw "CatalogPath must be an absolute drive path or UNC path: $Path"
  }

  $drive = $root.Substring(0, 1).ToUpperInvariant()
  $relativePath = $resolvedPath.Substring($root.Length).TrimStart('\', '/')
  return "\\localhost\$drive`$\$relativePath"
}

function ConvertTo-OfficeCatalogRegistryUrl {
  param(
    [Parameter(Mandatory = $true)][string]$Path
  )

  return ConvertTo-OfficeCatalogUrl -Path $Path
}

$hosts = @(
  @{ Name = "Word"; ProcessName = "WINWORD"; CatalogFile = "office-mcp-word.xml"; Manifest = Join-Path $RepoRoot "src\office-ctl\word\manifest.xml"; AddinId = "11111111-aaaa-bbbb-cccc-222222222222" },
  @{ Name = "Excel"; ProcessName = "EXCEL"; CatalogFile = "office-mcp-excel.xml"; Manifest = Join-Path $RepoRoot "src\office-ctl\excel\manifest.xml"; AddinId = "33333333-aaaa-bbbb-cccc-444444444444" },
  @{ Name = "PowerPoint"; ProcessName = "POWERPNT"; CatalogFile = "office-mcp-powerpoint.xml"; Manifest = Join-Path $RepoRoot "src\office-ctl\powerpoint\manifest.xml"; AddinId = "44444444-aaaa-bbbb-cccc-555555555555" }
)

New-Item -ItemType Directory -Force -Path $CatalogPath | Out-Null
Remove-LegacyHostCatalogFolders -CatalogRoot $CatalogPath

foreach ($officeHost in $hosts) {
  if (-not (Test-Path -LiteralPath $officeHost.Manifest)) {
    throw "Cannot find $($officeHost.Name) manifest: $($officeHost.Manifest)"
  }
  Write-LocalManifest -SourcePath $officeHost.Manifest -DestinationPath (Join-Path $CatalogPath $officeHost.CatalogFile) -Origin $BaseUrl
}

if (-not $SkipRegistry) {
  Assert-TrustedCatalogRegistryKey -Path $TrustedCatalogRegistryKey
  $catalogUrl = ConvertTo-OfficeCatalogRegistryUrl -Path $CatalogPath
  $key = $TrustedCatalogRegistryKey
  New-Item -Path $key -Force | Out-Null
  Set-ItemProperty -Path $key -Name Id -Value "office-mcp"
  Set-ItemProperty -Path $key -Name Url -Value $catalogUrl
  Set-ItemProperty -Path $key -Name Flags -Value 1 -Type DWord
  foreach ($officeHost in $hosts) {
    Remove-DeveloperDebugRegistration -AddinId $officeHost.AddinId
  }
} else {
  $catalogUrl = ConvertTo-OfficeCatalogUrl -Path $CatalogPath
}

if ($ClearOfficeCache) {
  Assert-OfficeHostsClosed -Hosts $hosts
  Remove-OfficeAddinCache -Hosts $hosts
}

Write-Output "Registered Office trusted catalog: $CatalogPath"
Write-Output "Catalog URL: $catalogUrl"
Write-Output "Manifest origin: $BaseUrl"
Write-Output "Word manifest: $(Join-Path $CatalogPath 'office-mcp-word.xml')"
Write-Output "Excel manifest: $(Join-Path $CatalogPath 'office-mcp-excel.xml')"
Write-Output "PowerPoint manifest: $(Join-Path $CatalogPath 'office-mcp-powerpoint.xml')"
if ($ClearOfficeCache) { Write-Output "Cleared Office WEF add-in cache for office-mcp." }
