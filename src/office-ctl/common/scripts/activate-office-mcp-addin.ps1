param(
  [string]$HostName = $env:OFFICE_MCP_E2E_HOST,
  [string]$DocumentPath = $env:OFFICE_MCP_E2E_DOCUMENT_PATH,
  [int]$TimeoutSeconds = 30,
  [switch]$DryRun = ($env:OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN -eq "1")
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($HostName)) { throw "OFFICE_MCP_E2E_HOST is required." }
if ([string]::IsNullOrWhiteSpace($DocumentPath)) { throw "OFFICE_MCP_E2E_DOCUMENT_PATH is required." }

$hostKey = $HostName.ToLowerInvariant()
if ($hostKey -notin @("word", "excel", "powerpoint")) {
  throw "Unsupported Office MCP E2E host: $HostName"
}

if ($DryRun) {
  Write-Output (@{
      activated = $true
      dry_run = $true
      host = $hostKey
      document_path = $DocumentPath
      control_name = "Open Control Panel"
    } | ConvertTo-Json -Compress)
  exit 0
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

function Get-OfficeApplication {
  param([Parameter(Mandatory = $true)][string]$HostKey)

  $progId = switch ($HostKey) {
    "word" { "Word.Application" }
    "excel" { "Excel.Application" }
    "powerpoint" { "PowerPoint.Application" }
  }
  return [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
}

function Activate-DriverDocument {
  param(
    [Parameter(Mandatory = $true)]$Application,
    [Parameter(Mandatory = $true)][string]$HostKey,
    [Parameter(Mandatory = $true)][string]$Path
  )

  $target = [System.IO.Path]::GetFullPath($Path)
  if ($HostKey -eq "word") {
    foreach ($document in @($Application.Documents)) {
      if ([System.IO.Path]::GetFullPath($document.FullName) -eq $target) {
        $document.Activate()
        $Application.Visible = $true
        $Application.Activate()
        return
      }
    }
  }
  if ($HostKey -eq "excel") {
    foreach ($workbook in @($Application.Workbooks)) {
      if ([System.IO.Path]::GetFullPath($workbook.FullName) -eq $target) {
        $workbook.Activate()
        $Application.Visible = $true
        return
      }
    }
  }
  if ($HostKey -eq "powerpoint") {
    foreach ($presentation in @($Application.Presentations)) {
      if ([System.IO.Path]::GetFullPath($presentation.FullName) -eq $target) {
        $presentation.Windows.Item(1).Activate()
        return
      }
    }
  }
  throw "Could not find driver-owned $HostKey document: $target"
}

function Find-DescendantByName {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$Name
  )

  $condition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::NameProperty,
    $Name
  )
  return $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}

function Invoke-Control {
  param([Parameter(Mandatory = $true)]$Element)

  $invokePattern = $null
  if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
    $invokePattern.Invoke()
    return
  }

  $rect = $Element.Current.BoundingRectangle
  if ($rect.IsEmpty) { throw "Open Control Panel button has no clickable bounds." }
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
    [int]($rect.Left + ($rect.Width / 2)),
    [int]($rect.Top + ($rect.Height / 2))
  )
  Add-Type -Namespace NativeMethods -Name Mouse -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern void mouse_event(int flags, int dx, int dy, int data, int extraInfo);
"@
  [NativeMethods.Mouse]::mouse_event(0x0002, 0, 0, 0, 0)
  [NativeMethods.Mouse]::mouse_event(0x0004, 0, 0, 0, 0)
}

$app = Get-OfficeApplication -HostKey $hostKey
Activate-DriverDocument -Application $app -HostKey $hostKey -Path $DocumentPath

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
do {
  Start-Sleep -Milliseconds 500
  $desktop = [System.Windows.Automation.AutomationElement]::RootElement
  $button = Find-DescendantByName -Root $desktop -Name "Open Control Panel"
  if ($button) {
    Invoke-Control -Element $button
    Write-Output (@{
        activated = $true
        host = $hostKey
        document_path = $DocumentPath
        control_name = "Open Control Panel"
      } | ConvertTo-Json -Compress)
    exit 0
  }
} while ((Get-Date) -lt $deadline)

throw "Timed out waiting for the Office MCP Control ribbon button."
