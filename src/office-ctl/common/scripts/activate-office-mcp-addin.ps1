param(
  [string]$HostName = $env:OFFICE_MCP_E2E_HOST,
  [string]$DocumentPath = $env:OFFICE_MCP_E2E_DOCUMENT_PATH,
  [int]$TimeoutSeconds = 30,
  [string]$LogPath = $env:OFFICE_MCP_E2E_ACTIVATOR_LOG,
  [switch]$DryRun = ($env:OFFICE_MCP_E2E_ACTIVATOR_DRY_RUN -eq "1")
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($HostName)) { throw "OFFICE_MCP_E2E_HOST is required." }
if ([string]::IsNullOrWhiteSpace($DocumentPath)) { throw "OFFICE_MCP_E2E_DOCUMENT_PATH is required." }

$hostKey = $HostName.ToLowerInvariant()
if ($hostKey -notin @("word", "excel", "powerpoint")) {
  throw "Unsupported Office MCP E2E host: $HostName"
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
  $LogPath = Join-Path $env:TEMP "office-mcp-activator-$hostKey.log"
}

function Write-ActivatorLog {
  param([Parameter(Mandatory = $true)][string]$Message)
  Add-Content -LiteralPath $LogPath -Value "$(Get-Date -Format o) $Message"
}

function Get-CanonicalPath {
  param([Parameter(Mandatory = $true)][string]$Path)
  try { return (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName.ToLowerInvariant() }
  catch { return [System.IO.Path]::GetFullPath($Path).ToLowerInvariant() }
}

Write-ActivatorLog "start host=$hostKey document=$DocumentPath"

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

  $target = Get-CanonicalPath -Path $Path
  if ($HostKey -eq "word") {
    foreach ($document in @($Application.Documents)) {
      if ((Get-CanonicalPath -Path $document.FullName) -eq $target) {
        Write-ActivatorLog "activated word document=$($document.FullName)"
        $document.Activate()
        $Application.Visible = $true
        $Application.Activate()
        return [IntPtr]$Application.ActiveWindow.Hwnd
      }
    }
  }
  if ($HostKey -eq "excel") {
    foreach ($workbook in @($Application.Workbooks)) {
      if ((Get-CanonicalPath -Path $workbook.FullName) -eq $target) {
        Write-ActivatorLog "activated excel workbook=$($workbook.FullName)"
        $workbook.Activate()
        $Application.Visible = $true
        return [IntPtr]$Application.Hwnd
      }
    }
  }
  if ($HostKey -eq "powerpoint") {
    foreach ($presentation in @($Application.Presentations)) {
      if ((Get-CanonicalPath -Path $presentation.FullName) -eq $target) {
        Write-ActivatorLog "activated powerpoint presentation=$($presentation.FullName)"
        $window = $presentation.Windows.Item(1)
        $window.Activate()
        return [IntPtr]$window.HWND
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

function Find-OfficeWindow {
  param([Parameter(Mandatory = $true)][string]$HostKey)
  $processName = switch ($HostKey) {
    "word" { "WINWORD" }
    "excel" { "EXCEL" }
    "powerpoint" { "POWERPNT" }
  }
  $processes = @(Get-Process -Name $processName -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    Sort-Object StartTime -Descending)
  foreach ($process in $processes) {
    $element = [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
    if ($element) {
      Write-ActivatorLog "using window pid=$($process.Id) title=$($process.MainWindowTitle)"
      return $element
    }
  }
  throw "Could not find a visible $HostKey Office window."
}

function Get-OfficeWindowFromHandle {
  param([Parameter(Mandatory = $true)][IntPtr]$Handle)
  if ($Handle -eq [IntPtr]::Zero) { throw "Office window handle is zero." }
  $element = [System.Windows.Automation.AutomationElement]::FromHandle($Handle)
  if (-not $element) { throw "Could not resolve Office window from handle $Handle." }
  Write-ActivatorLog "using driver document window handle=$Handle title=$($element.Current.Name)"
  return $element
}

function Get-VisibleControlNameSample {
  param([Parameter(Mandatory = $true)]$Root)
  $condition = [System.Windows.Automation.Condition]::TrueCondition
  return @($Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition) |
    ForEach-Object { $_.Current.Name } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique -First 80)
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

function Try-InvokeNamedControl {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$Name
  )
  $control = Find-DescendantByName -Root $Root -Name $Name
  if (-not $control) { return $false }
  Write-ActivatorLog "invoking control name=$Name"
  Invoke-Control -Element $control
  Start-Sleep -Milliseconds 500
  return $true
}

$app = Get-OfficeApplication -HostKey $hostKey
$driverWindowHandle = Activate-DriverDocument -Application $app -HostKey $hostKey -Path $DocumentPath

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
do {
  Start-Sleep -Milliseconds 500
  $window = Get-OfficeWindowFromHandle -Handle $driverWindowHandle
  $button = Find-DescendantByName -Root $window -Name "Open Control Panel"
  if ($button) {
    Write-ActivatorLog "found Open Control Panel control"
    Invoke-Control -Element $button
    Write-Output (@{
        activated = $true
        host = $hostKey
        document_path = $DocumentPath
        control_name = "Open Control Panel"
      } | ConvertTo-Json -Compress)
    exit 0
  }
  foreach ($tabName in @("Home", "Insert", "Add-ins", "My Add-ins")) {
    if (Try-InvokeNamedControl -Root $window -Name $tabName) {
      $button = Find-DescendantByName -Root (Get-OfficeWindowFromHandle -Handle $driverWindowHandle) -Name "Open Control Panel"
      if ($button) {
        Write-ActivatorLog "found Open Control Panel control after tab=$tabName"
        Invoke-Control -Element $button
        Write-Output (@{
            activated = $true
            host = $hostKey
            document_path = $DocumentPath
            control_name = "Open Control Panel"
            tab_name = $tabName
          } | ConvertTo-Json -Compress)
        exit 0
      }
    }
  }
} while ((Get-Date) -lt $deadline)

$sample = Get-VisibleControlNameSample -Root (Get-OfficeWindowFromHandle -Handle $driverWindowHandle)
Write-ActivatorLog "timed out; visible control sample=$($sample -join ' | ')"

throw "Timed out waiting for the Office MCP Control ribbon button."
