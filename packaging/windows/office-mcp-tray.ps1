param(
  [string]$InstallRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path),
  [string]$RuntimePath = $env:OFFICE_MCP_UI_RUNTIME_PATH,
  [string]$ProbeStatePath,
  [switch]$Probe
)

$ErrorActionPreference = "Stop"

$envScript = Join-Path $InstallRoot "office-mcp-env.ps1"
if (Test-Path -LiteralPath $envScript) {
  . $envScript
}

$daemonExe = Join-Path $InstallRoot "office-mcp-daemon.exe"
if (-not (Test-Path -LiteralPath $daemonExe)) {
  throw "Cannot find office-mcp-daemon.exe under $InstallRoot."
}

$argsList = @("tray")
if ($RuntimePath) { $argsList += @("--runtime-path", $RuntimePath) }
if ($ProbeStatePath) { $argsList += @("--probe-state-path", $ProbeStatePath) }
if ($Probe) { $argsList += "--probe" }

& $daemonExe @argsList
exit $LASTEXITCODE
