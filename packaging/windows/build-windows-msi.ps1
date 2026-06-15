param(
  [string]$Version = "0.1.0",
  [string]$OutputDir = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) "artifacts"),
  [switch]$SkipNpmInstall
)

$ErrorActionPreference = "Stop"

function Assert-LastExitCode([string]$CommandName) {
  if ($LASTEXITCODE -ne 0) {
    throw "$CommandName failed with exit code $LASTEXITCODE"
  }
}

function Reset-DirectoryInside([string]$Path, [string]$Parent) {
  New-Item -ItemType Directory -Force -Path $Parent | Out-Null
  $resolvedParent = (Resolve-Path -LiteralPath $Parent).Path
  if (Test-Path -LiteralPath $Path) {
    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    if (-not $resolvedPath.StartsWith($resolvedParent, [System.StringComparison]::OrdinalIgnoreCase)) {
      throw "Refusing to remove staging directory outside output directory: $resolvedPath"
    }
    Remove-Item -LiteralPath $resolvedPath -Recurse -Force
  }
  New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

function New-WixPayloadFragment([string]$StageRoot, [string]$OutputPath) {
  $resolvedStageRoot = (Resolve-Path -LiteralPath $StageRoot).Path
  $directories = New-Object System.Collections.Generic.List[string]
  $refs = New-Object System.Collections.Generic.List[string]

  Add-WixDirectory -BasePath $resolvedStageRoot -Path $resolvedStageRoot -Directories $directories -Refs $refs

  $content = @()
  $content += '<?xml version="1.0" encoding="UTF-8"?>'
  $content += '<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">'
  $content += '  <Fragment>'
  $content += '    <DirectoryRef Id="INSTALLFOLDER">'
  $content += $directories
  $content += '    </DirectoryRef>'
  $content += '  </Fragment>'
  $content += '  <Fragment>'
  $content += '    <ComponentGroup Id="OfficeMcpPayloadComponents">'
  $content += $refs
  $content += '    </ComponentGroup>'
  $content += '  </Fragment>'
  $content += '</Wix>'
  $content | Set-Content -Encoding UTF8 -Path $OutputPath
}

function Add-WixDirectory(
  [string]$BasePath,
  [string]$Path,
  [System.Collections.Generic.List[string]]$Directories,
  [System.Collections.Generic.List[string]]$Refs
) {
  foreach ($childDir in Get-ChildItem -LiteralPath $Path -Directory | Sort-Object FullName) {
    $dirId = "Dir_" + (Get-StableId $BasePath $childDir.FullName)
    $name = ConvertTo-WixXml $childDir.Name
    $Directories.Add("      <Directory Id=`"$dirId`" Name=`"$name`">")
    Add-WixDirectory -BasePath $BasePath -Path $childDir.FullName -Directories $Directories -Refs $Refs
    $Directories.Add("      </Directory>")
  }

  foreach ($file in Get-ChildItem -LiteralPath $Path -File | Sort-Object FullName) {
    $componentId = "Cmp_" + (Get-StableId $BasePath $file.FullName)
    $fileId = "File_" + (Get-StableId $BasePath ($file.FullName + ":file"))
    $source = ConvertTo-WixXml $file.FullName
    $name = ConvertTo-WixXml $file.Name
    $Directories.Add("      <Component Id=`"$componentId`" Guid=`"*`">")
    $Directories.Add("        <File Id=`"$fileId`" Source=`"$source`" Name=`"$name`" KeyPath=`"yes`" />")
    $Directories.Add("      </Component>")
    $Refs.Add("      <ComponentRef Id=`"$componentId`" />")
  }
}

function Get-StableId([string]$BasePath, [string]$Path) {
  $baseUri = [System.Uri]((Join-Path $BasePath "") -replace '\\', '/')
  $pathUri = [System.Uri]($Path -replace '\\', '/')
  $relative = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($pathUri).ToString()).ToLowerInvariant()
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($relative)
  $sha = [System.Security.Cryptography.SHA1]::Create()
  try {
    return ([System.BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '')
  } finally {
    $sha.Dispose()
  }
}

function ConvertTo-WixXml([string]$Value) {
  return [System.Security.SecurityElement]::Escape($Value)
}

function Assert-MsiStagePayload([string]$StageRoot, [string]$GeneratedWxsPath) {
  $requiredPaths = @(
    "node\node.exe",
    "mcp-server\dist\src\cli.js",
    "mcp-server\dist\src\daemon.js",
    "mcp-server\node_modules\@modelcontextprotocol\sdk\package.json",
    "mcp-server\node_modules\ws\package.json",
    "mcp-server\node_modules\zod\package.json",
    "mcp-server\scripts\export-localhost-dev-cert.ps1",
    "addin\manifest.xml",
    "addin\public\taskpane.html",
    "addin\public\taskpane.css",
    "addin\public\taskpane.js",
    "addin-catalog\manifest.xml",
    "office-mcp-env.ps1",
    "config.toml",
    "office-mcp-daemon.ps1",
    "office-mcp.ps1"
  )

  foreach ($relativePath in $requiredPaths) {
    $path = Join-Path $StageRoot $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
      throw "MSI staging payload is missing required path: $relativePath"
    }
  }

  $daemonLauncher = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp-daemon.ps1")
  if ($daemonLauncher -notmatch "office-mcp-env\.ps1" -or $daemonLauncher -notmatch "node\\node\.exe" -or $daemonLauncher -notmatch "daemon run") {
    throw "Daemon launcher must use the packaged node.exe and run the daemon."
  }

  $cliLauncher = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp.ps1")
  if ($cliLauncher -notmatch "office-mcp-env\.ps1" -or $cliLauncher -notmatch "node\\node\.exe" -or $cliLauncher -notmatch "@args") {
    throw "CLI launcher must use the packaged node.exe and forward arguments."
  }

  $generatedWxs = Get-Content -Raw -LiteralPath $GeneratedWxsPath
  foreach ($needle in @("node.exe", "cli.js", "@modelcontextprotocol", "taskpane.js", "addin-catalog", "office-mcp-env.ps1", "office-mcp-daemon.ps1")) {
    if (-not $generatedWxs.Contains($needle)) {
      throw "Generated WiX payload is missing expected entry: $needle"
    }
  }

  $envScript = Get-Content -Raw -LiteralPath (Join-Path $StageRoot "office-mcp-env.ps1")
  if ($envScript -notmatch "OFFICE_MCP_INSTALL_ROOT") {
    throw "MSI environment script must set OFFICE_MCP_INSTALL_ROOT."
  }
  if ($envScript -notmatch "OFFICE_MCP_CONFIG_PATH") {
    throw "MSI environment script must set OFFICE_MCP_CONFIG_PATH."
  }
  foreach ($envName in @("OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH", "OFFICE_MCP_ADDIN_CHANNEL__PORT", "OFFICE_MCP_MCP_HTTP__PORT")) {
    if ($envScript -notmatch $envName) {
      throw "MSI environment script must set $envName."
    }
  }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$serverRoot = Join-Path $repoRoot "mcp-server"
$addinRoot = Join-Path $repoRoot "addin"
$wxsPath = Join-Path $repoRoot "packaging\wix\Product.wxs"
$outputPath = Join-Path $OutputDir "office-mcp-setup-$Version-x64.msi"
$stageRoot = Join-Path $OutputDir "msi-stage"
$generatedWxsPath = Join-Path $OutputDir "msi-payload.wxs"

if (-not (Test-Path $wxsPath)) {
  throw "Missing WiX source file: $wxsPath"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

Push-Location $serverRoot
try {
  if (-not $SkipNpmInstall) {
    if (Test-Path (Join-Path $serverRoot "node_modules")) {
      npm install
      Assert-LastExitCode "npm install"
    } else {
      npm ci
      Assert-LastExitCode "npm ci"
    }
  }
  npm run check
  Assert-LastExitCode "npm run check"
} finally {
  Pop-Location
}

Push-Location $addinRoot
try {
  if (-not $SkipNpmInstall) {
    npm ci
    Assert-LastExitCode "npm ci"
  }
  npm run check
  Assert-LastExitCode "npm run check"
} finally {
  Pop-Location
}

Reset-DirectoryInside -Path $stageRoot -Parent $OutputDir

$stageServerRoot = Join-Path $stageRoot "mcp-server"
$stageAddinRoot = Join-Path $stageRoot "addin"
$stageCatalogRoot = Join-Path $stageRoot "addin-catalog"
$stageNodeRoot = Join-Path $stageRoot "node"
New-Item -ItemType Directory -Force -Path $stageServerRoot, $stageAddinRoot, $stageCatalogRoot, $stageNodeRoot | Out-Null

$nodeCommand = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $nodeCommand) {
  throw "node.exe must be available on PATH to build the Windows MSI runtime payload."
}
Copy-Item -Force -Path $nodeCommand.Source -Destination (Join-Path $stageNodeRoot "node.exe")

Copy-Item -Force -Path (Join-Path $serverRoot "package.json") -Destination $stageServerRoot
Copy-Item -Force -Path (Join-Path $serverRoot "package-lock.json") -Destination $stageServerRoot
Copy-Item -Recurse -Force -Path (Join-Path $serverRoot "dist") -Destination $stageServerRoot
Copy-Item -Recurse -Force -Path (Join-Path $serverRoot "scripts") -Destination $stageServerRoot

Push-Location $stageServerRoot
try {
  npm install --omit=dev --ignore-scripts
  Assert-LastExitCode "npm install --omit=dev --ignore-scripts"
} finally {
  Pop-Location
}

Copy-Item -Force -Path (Join-Path $addinRoot "manifest.xml") -Destination $stageAddinRoot
Copy-Item -Recurse -Force -Path (Join-Path $addinRoot "public") -Destination $stageAddinRoot
Copy-Item -Force -Path (Join-Path $addinRoot "manifest.xml") -Destination (Join-Path $stageCatalogRoot "manifest.xml")

@'
[addin_channel]
bind = "localhost"
port = 8765
heartbeat_interval_sec = 30
heartbeat_timeout_sec = 10
session_grace_sec = 60
max_pending_per_session = 4
shared_secret = ""
certificate_path = ""

[mcp_http]
bind = "127.0.0.1"
port = 8800
api_key = ""

[limits]
max_response_bytes = 1048576
max_request_bytes = 16777216
max_ws_frame_bytes = 16777216
default_tool_timeout_ms = 30000
requests_per_minute = 120

[audit]
enabled = false
path = ""

[logging]
level = "info"
file = ""
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "config.toml")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pfxPath = Join-Path $installRoot '.office-mcp-localhost.pfx'

$env:OFFICE_MCP_INSTALL_ROOT = $installRoot
$env:OFFICE_MCP_CONFIG_PATH = Join-Path $installRoot 'config.toml'
$env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PATH = $pfxPath
$env:OFFICE_MCP_ADDIN_CHANNEL__CERTIFICATE_PASSPHRASE = 'office-mcp-localhost'
$env:OFFICE_MCP_ADDIN_CHANNEL__BIND = 'localhost'
$env:OFFICE_MCP_ADDIN_CHANNEL__PORT = '8765'
$env:OFFICE_MCP_MCP_HTTP__BIND = '127.0.0.1'
$env:OFFICE_MCP_MCP_HTTP__PORT = '8800'
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "office-mcp-env.ps1")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $installRoot 'office-mcp-env.ps1')
$serverRoot = Join-Path $installRoot 'mcp-server'
$nodeExe = Join-Path $installRoot 'node\node.exe'
$pfxPath = Join-Path $installRoot '.office-mcp-localhost.pfx'

if (-not (Test-Path -LiteralPath $pfxPath)) {
  & (Join-Path $serverRoot 'scripts\export-localhost-dev-cert.ps1') -OutputPath $pfxPath
}

Set-Location $serverRoot
& $nodeExe .\dist\src\cli.js daemon run
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "office-mcp-daemon.ps1")

@'
$ErrorActionPreference = 'Stop'
$installRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $installRoot 'office-mcp-env.ps1')
$serverRoot = Join-Path $installRoot 'mcp-server'
$nodeExe = Join-Path $installRoot 'node\node.exe'
Set-Location $serverRoot
& $nodeExe .\dist\src\cli.js @args
'@ | Set-Content -Encoding ASCII -Path (Join-Path $stageRoot "office-mcp.ps1")

New-WixPayloadFragment -StageRoot $stageRoot -OutputPath $generatedWxsPath
Assert-MsiStagePayload -StageRoot $stageRoot -GeneratedWxsPath $generatedWxsPath

Push-Location $repoRoot
try {
  dotnet tool restore
  Assert-LastExitCode "dotnet tool restore"
  dotnet wix build $wxsPath $generatedWxsPath -arch x64 -d "RepoRoot=$repoRoot" -d "PackageVersion=$Version" -o $outputPath
  Assert-LastExitCode "dotnet wix build"
} finally {
  Pop-Location
}

if (-not (Test-Path $outputPath)) {
  throw "MSI was not created: $outputPath"
}

Write-Output "Built $outputPath"
