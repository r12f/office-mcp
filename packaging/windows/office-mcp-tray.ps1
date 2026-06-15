param(
  [string]$InstallRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path),
  [string]$TaskName = "office-mcp",
  [string]$RuntimePath = $env:OFFICE_MCP_UI_RUNTIME_PATH,
  [string]$ProbeStatePath,
  [switch]$Probe,
  [switch]$AllowUntrustedRuntimeCertificate
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

$envScript = Join-Path $InstallRoot "office-mcp-env.ps1"
if (Test-Path -LiteralPath $envScript) {
  . $envScript
}

$cli = Join-Path $InstallRoot "office-mcp.ps1"
$nodeCli = Join-Path $InstallRoot "mcp-server\dist\src\cli.js"
$nodeExe = Join-Path $InstallRoot "node\node.exe"
$runtimePath = if ($RuntimePath) { $RuntimePath } else { Join-Path $env:LOCALAPPDATA "office-mcp\ui-runtime.json" }
$script:LastUiStateError = $null

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
  $script:LastUiStateError = $null
  if ($Probe -and $ProbeStatePath -and (Test-Path -LiteralPath $ProbeStatePath)) {
    try { return (Get-Content -Raw -LiteralPath $ProbeStatePath | ConvertFrom-Json) } catch { $script:LastUiStateError = $_.Exception.Message; return $null }
  }
  $runtime = Get-OfficeMcpUiRuntime
  if (-not $runtime -or -not $runtime.stateUrl -or -not $runtime.token) { return $null }
  try {
    if ($AllowUntrustedRuntimeCertificate) {
      $previousCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
      try {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        $request = [System.Net.HttpWebRequest]::Create([string]$runtime.stateUrl)
        $request.Method = "GET"
        $request.Timeout = 2000
        $request.ReadWriteTimeout = 2000
        $request.Headers.Add("x-office-mcp-ui-token", [string]$runtime.token)
        $response = $request.GetResponse()
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $body = $reader.ReadToEnd()
        $reader.Dispose()
        $response.Dispose()
        return ($body | ConvertFrom-Json)
      } finally {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $previousCallback
      }
    }
    return Invoke-RestMethod -Uri $runtime.stateUrl -Headers @{ "x-office-mcp-ui-token" = $runtime.token } -TimeoutSec 2
  } catch {
    $script:LastUiStateError = $_.Exception.Message
    return $null
  }
}

function Get-OfficeMcpDocumentCount($UiState) {
  if (-not $UiState -or -not $UiState.documents) { return 0 }
  $documents = 0
  foreach ($group in @("word", "excel", "powerpoint", "outlook", "other")) {
    if ($UiState.documents.$group) { $documents += @($UiState.documents.$group).Count }
  }
  return $documents
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
  $uiState = Get-OfficeMcpUiState
  if ($uiState) {
    $documents = Get-OfficeMcpDocumentCount $uiState
    $clientCount = if ($uiState.clients) { @($uiState.clients).Count } else { 0 }
    $statusText = if ($uiState.daemon.status) { (Get-Culture).TextInfo.ToTitleCase([string]$uiState.daemon.status) } else { "Up" }
    return @{ Status = $statusText; Clients = $clientCount; Documents = $documents }
  }
  $status = Invoke-OfficeMcpJson @("daemon", "status")
  if ($status -and $status.running) { return @{ Status = "Up"; Clients = 0; Documents = 0 } }
  return @{ Status = "Down"; Clients = 0; Documents = 0 }
}

function Get-OfficeMcpMenuLabels {
  $labels = @()
  foreach ($item in $menu.Items) {
    if ($item -is [System.Windows.Forms.ToolStripSeparator]) { $labels += "---" }
    else { $labels += $item.Text }
  }
  return $labels
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

if ($Probe) {
  $state = Get-OfficeMcpStatusText
  $statusItem.Text = "Status: $($state.Status)"
  $clientsItem.Text = "Clients: $($state.Clients)"
  $documentsItem.Text = "Documents: $($state.Documents)"
  $notifyIcon.Text = "Office MCP - $($state.Status) - $($state.Clients) clients - $($state.Documents) documents"
  $runtime = Get-OfficeMcpUiRuntime
  $uiState = Get-OfficeMcpUiState
  $evidence = [ordered]@{
    ok = $true
    notify_icon_created = $null -ne $notifyIcon
    context_menu_created = $null -ne $notifyIcon.ContextMenuStrip
    menu_items = @(Get-OfficeMcpMenuLabels)
    status = [ordered]@{ Status = $state.Status; Clients = $state.Clients; Documents = $state.Documents }
    tooltip = $notifyIcon.Text
    runtime_path = $runtimePath
    can_read_runtime = $null -ne $runtime
    state_fetch_ok = $null -ne $uiState
    state_fetch_error = $script:LastUiStateError
    state_client_count = if ($uiState -and $uiState.clients) { @($uiState.clients).Count } else { $null }
    state_document_count = if ($uiState) { Get-OfficeMcpDocumentCount $uiState } else { $null }
  }
  $notifyIcon.Visible = $false
  $notifyIcon.Dispose()
  $evidence | ConvertTo-Json -Depth 6
  return
}

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
