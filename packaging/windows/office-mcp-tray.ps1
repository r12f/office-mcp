param(
  [string]$InstallRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path),
  [string]$TaskName = "office-mcp"
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$envScript = Join-Path $InstallRoot "office-mcp-env.ps1"
if (Test-Path -LiteralPath $envScript) {
  . $envScript
}

$cli = Join-Path $InstallRoot "office-mcp.ps1"
$nodeCli = Join-Path $InstallRoot "mcp-server\dist\src\cli.js"
$nodeExe = Join-Path $InstallRoot "node\node.exe"
$runtimePath = if ($env:OFFICE_MCP_UI_RUNTIME_PATH) { $env:OFFICE_MCP_UI_RUNTIME_PATH } else { Join-Path $env:LOCALAPPDATA "office-mcp\ui-runtime.json" }

function Invoke-OfficeMcpJson([string[]]$Arguments) {
  try {
    if (Test-Path -LiteralPath $cli) {
      $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cli @Arguments 2>$null
    } elseif ((Test-Path -LiteralPath $nodeExe) -and (Test-Path -LiteralPath $nodeCli)) {
      $output = & $nodeExe $nodeCli @Arguments 2>$null
    } else {
      $output = & node $nodeCli @Arguments 2>$null
    }
    if ($LASTEXITCODE -ne 0 -or -not $output) { return $null }
    return ($output | Out-String | ConvertFrom-Json)
  } catch {
    return $null
  }
}

function Start-OfficeMcpUi {
  $runtime = Get-OfficeMcpUiRuntime
  if ($runtime -and $runtime.uiUrl) {
    Start-Process $runtime.uiUrl
  } elseif (Test-Path -LiteralPath $cli) {
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $cli, "ui")
  }
}

function Get-OfficeMcpUiRuntime {
  try {
    if (-not (Test-Path -LiteralPath $runtimePath)) { return $null }
    return (Get-Content -Raw -LiteralPath $runtimePath | ConvertFrom-Json)
  } catch {
    return $null
  }
}

function Get-OfficeMcpUiState {
  $runtime = Get-OfficeMcpUiRuntime
  if (-not $runtime -or -not $runtime.stateUrl -or -not $runtime.token) { return $null }
  try {
    return Invoke-RestMethod -Uri $runtime.stateUrl -Headers @{ "x-office-mcp-ui-token" = $runtime.token } -TimeoutSec 2
  } catch {
    return $null
  }
}

function Stop-OfficeMcpDaemon {
  if (Test-Path -LiteralPath $cli) {
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $cli, "daemon", "stop")
  }
}

function Start-OfficeMcpDaemon {
  $status = Invoke-OfficeMcpJson @("daemon", "status")
  if ($status -and $status.running) { return }
  if (Test-Path -LiteralPath $cli) {
    Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $cli, "daemon", "start")
  }
}

function Get-OfficeMcpStatusText {
  $status = Invoke-OfficeMcpJson @("daemon", "status")
  if (-not $status -or -not $status.running) {
    return @{ Status = "Down"; Clients = 0; Documents = 0 }
  }
  $uiState = Get-OfficeMcpUiState
  if (-not $uiState) {
    return @{ Status = "Up"; Clients = 0; Documents = 0 }
  }
  $documents = 0
  foreach ($group in @("word", "excel", "powerpoint", "outlook", "other")) {
    if ($uiState.documents.$group) { $documents += @($uiState.documents.$group).Count }
  }
  $clientCount = if ($uiState.clients) { @($uiState.clients).Count } else { 0 }
  $statusText = if ($uiState.daemon.status) { (Get-Culture).TextInfo.ToTitleCase([string]$uiState.daemon.status) } else { "Up" }
  return @{ Status = $statusText; Clients = $clientCount; Documents = $documents }
}

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Text = "Office MCP"
$notifyIcon.Icon = [System.Drawing.SystemIcons]::Application
$notifyIcon.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$statusItem = $menu.Items.Add("Status: Down")
$clientsItem = $menu.Items.Add("Clients: 0")
$documentsItem = $menu.Items.Add("Documents: 0")
$statusItem.Enabled = $false
$clientsItem.Enabled = $false
$documentsItem.Enabled = $false
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$showItem = $menu.Items.Add("Show Office MCP")
$quitItem = $menu.Items.Add("Quit Office MCP")

$showItem.add_Click({ Start-OfficeMcpUi })
$quitItem.add_Click({
  $result = [System.Windows.Forms.MessageBox]::Show("Quit Office MCP and disconnect clients?", "Quit Office MCP", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
  if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    Stop-OfficeMcpDaemon
    $notifyIcon.Visible = $false
    [System.Windows.Forms.Application]::Exit()
  }
})

$notifyIcon.ContextMenuStrip = $menu
$notifyIcon.add_DoubleClick({ Start-OfficeMcpUi })

Start-OfficeMcpDaemon

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 2000
$timer.add_Tick({
  $state = Get-OfficeMcpStatusText
  $statusItem.Text = "Status: $($state.Status)"
  $clientsItem.Text = "Clients: $($state.Clients)"
  $documentsItem.Text = "Documents: $($state.Documents)"
  $notifyIcon.Text = "Office MCP - $($state.Status) - $($state.Clients) clients - $($state.Documents) documents"
})
$timer.Start()

[System.Windows.Forms.Application]::Run()

$timer.Stop()
$notifyIcon.Visible = $false
$notifyIcon.Dispose()
