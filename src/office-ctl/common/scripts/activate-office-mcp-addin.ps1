param(
  [string]$HostName = $env:OFFICE_MCP_E2E_HOST,
  [string]$DocumentPath = $env:OFFICE_MCP_E2E_DOCUMENT_PATH,
  [string]$ManifestPath = $env:OFFICE_MCP_E2E_MANIFEST_PATH,
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
$script:officialDocumentPath = $null

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

function Invoke-OfficialSideload {
  if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    Write-ActivatorLog "official sideload skipped: manifest path missing"
    return
  }
  if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-ActivatorLog "official sideload skipped: manifest not found path=$ManifestPath"
    return
  }
  $appName = switch ($hostKey) {
    "word" { "Word" }
    "excel" { "Excel" }
    "powerpoint" { "PowerPoint" }
  }
  Write-ActivatorLog "official sideload start app=$appName manifest=$ManifestPath document=$DocumentPath"
  try {
    $manifestArg = '"' + $ManifestPath.Replace('"', '\"') + '"'
    $documentArg = '"' + $DocumentPath.Replace('"', '\"') + '"'
    $command = "npx --yes office-addin-dev-settings sideload $manifestArg desktop --app $appName --document $documentArg"
    $stdoutPath = Join-Path $env:TEMP "office-mcp-sideload-$hostKey-$([guid]::NewGuid().ToString('N')).out.log"
    $stderrPath = Join-Path $env:TEMP "office-mcp-sideload-$hostKey-$([guid]::NewGuid().ToString('N')).err.log"
    $process = Start-Process -FilePath "cmd.exe" -ArgumentList @('/d', '/s', '/c', $command) -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -PassThru
    if (-not $process.WaitForExit(15000)) {
      $output = @()
      if (Test-Path -LiteralPath $stdoutPath) { $output += Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue }
      if (Test-Path -LiteralPath $stderrPath) { $output += Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue }
      Write-ActivatorLog "official sideload timed out; leaving process running pid=$($process.Id) output=$($output -join ' | ')"
      return
    }
    $exitCode = $process.ExitCode
    $output = @()
    if (Test-Path -LiteralPath $stdoutPath) { $output += Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $stderrPath) { $output += Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue }
    Write-ActivatorLog "official sideload exit=$exitCode output=$($output -join ' | ')"
    $launchLine = @($output | Where-Object { $_ -match '^Launching .* via (.+)$' } | Select-Object -First 1)
    if ($launchLine -and ($launchLine -match '^Launching .* via (.+)$')) {
      $script:officialDocumentPath = $Matches[1]
      Write-ActivatorLog "official sideload document=$script:officialDocumentPath"
    }
    if ($exitCode -ne 0) {
      if ($script:officialDocumentPath) {
        Write-ActivatorLog "official sideload exit code was unavailable after launch output"
      } else {
        Write-ActivatorLog "official sideload failed: exit=$exitCode"
      }
    }
  } catch {
    Write-ActivatorLog "official sideload failed: $($_.Exception.Message)"
    Write-ActivatorLog "official sideload error detail: $($_ | Out-String)"
  }
}

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

function Get-OfficeStateSnapshot {
  param(
    [Parameter(Mandatory = $true)]$Application,
    [Parameter(Mandatory = $true)][string]$HostKey
  )

  $items = @()
  try {
    if ($HostKey -eq "word") {
      foreach ($document in @($Application.Documents)) {
        $items += "doc:name=$($document.Name);full=$($document.FullName);path=$($document.Path)"
      }
      foreach ($window in @($Application.Windows)) {
        $items += "window:caption=$($window.Caption);full=$($window.Document.FullName)"
      }
    }
    if ($HostKey -eq "excel") {
      foreach ($workbook in @($Application.Workbooks)) {
        $items += "workbook:name=$($workbook.Name);full=$($workbook.FullName);path=$($workbook.Path)"
      }
    }
    if ($HostKey -eq "powerpoint") {
      foreach ($presentation in @($Application.Presentations)) {
        $items += "presentation:name=$($presentation.Name);full=$($presentation.FullName);path=$($presentation.Path)"
      }
    }
  } catch {
    $items += "com-snapshot-error=$($_.Exception.Message)"
  }

  $processName = switch ($HostKey) {
    "word" { "WINWORD" }
    "excel" { "EXCEL" }
    "powerpoint" { "POWERPNT" }
  }
  foreach ($process in @(Get-Process -Name $processName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })) {
    $items += "process:pid=$($process.Id);title=$($process.MainWindowTitle);handle=$($process.MainWindowHandle)"
  }
  return ($items -join " || ")
}

function Wait-ForDriverDocument {
  param(
    [Parameter(Mandatory = $true)]$Application,
    [Parameter(Mandatory = $true)][string]$HostKey,
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)]$Deadline
  )

  $lastError = $null
  do {
    try {
      return Activate-DriverDocument -Application $Application -HostKey $HostKey -Path $Path
    } catch {
      $lastError = $_.Exception.Message
      Write-ActivatorLog "waiting for driver document path=$Path error=$lastError state=$(Get-OfficeStateSnapshot -Application $Application -HostKey $HostKey)"
      Start-Sleep -Milliseconds 500
    }
  } while ((Get-Date) -lt $Deadline)
  throw $lastError
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

function Find-DescendantByAutomationId {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$AutomationId
  )

  $condition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::AutomationIdProperty,
    $AutomationId
  )
  return $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}

function Test-AutomationElementCandidate {
  param(
    [Parameter(Mandatory = $true)]$Element,
    [int]$ProcessId = 0
  )

  try {
    if ($ProcessId -gt 0 -and $Element.Current.ProcessId -ne $ProcessId) { return $false }
    if ($Element.Current.IsOffscreen) { return $false }
    if (-not $Element.Current.IsEnabled) { return $false }
    $controlType = $Element.Current.ControlType.ProgrammaticName
    return $controlType -in @(
      "ControlType.Button",
      "ControlType.MenuItem",
      "ControlType.ListItem",
      "ControlType.Hyperlink",
      "ControlType.SplitButton"
    )
  } catch {
    return $false
  }
}

function Find-DescendantByNameLike {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$Name,
    [int]$ProcessId = 0
  )

  $exact = Find-DescendantByName -Root $Root -Name $Name
  if ($exact -and (Test-AutomationElementCandidate -Element $exact -ProcessId $ProcessId)) { return $exact }

  $condition = [System.Windows.Automation.Condition]::TrueCondition
  foreach ($element in @($Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition))) {
    $currentName = $element.Current.Name
    if (-not [string]::IsNullOrWhiteSpace($currentName) -and $currentName -like "*$Name*" -and (Test-AutomationElementCandidate -Element $element -ProcessId $ProcessId)) {
      return $element
    }
  }
  return $null
}

function Find-GlobalControlByName {
  param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][int]$ProcessId
  )
  return Find-DescendantByNameLike -Root ([System.Windows.Automation.AutomationElement]::RootElement) -Name $Name -ProcessId $ProcessId
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

function Get-OfficeProcessIdFromHandle {
  param([Parameter(Mandatory = $true)][IntPtr]$Handle)
  $element = Get-OfficeWindowFromHandle -Handle $Handle
  return [int]$element.Current.ProcessId
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

  $togglePattern = $null
  if ($Element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
    $togglePattern.Toggle()
    return
  }

  $expandPattern = $null
  if ($Element.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandPattern)) {
    $expandPattern.Expand()
    return
  }

  $rect = $Element.Current.BoundingRectangle
  if ($rect.IsEmpty) { throw "Control has no clickable bounds." }
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
    [int]($rect.Left + ($rect.Width / 2)),
    [int]($rect.Top + ($rect.Height / 2))
  )
  Add-Type -Namespace NativeMethods -Name Mouse -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern void mouse_event(int flags, int dx, int dy, int data, int extraInfo);
"@
  Add-Type -Namespace NativeMethods -Name Window -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetForegroundWindow(System.IntPtr hWnd);
"@ -ErrorAction SilentlyContinue
  try { [NativeMethods.Window]::SetForegroundWindow([IntPtr]$Element.Current.NativeWindowHandle) | Out-Null } catch {}
  [NativeMethods.Mouse]::mouse_event(0x0002, 0, 0, 0, 0)
  [NativeMethods.Mouse]::mouse_event(0x0004, 0, 0, 0, 0)
}

function Try-InvokeNamedControl {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$Name
  )
  $processId = 0
  try { $processId = [int]$Root.Current.ProcessId } catch {}
  $control = Find-DescendantByNameLike -Root $Root -Name $Name -ProcessId $processId
  if (-not $control) { return $false }
  Write-ActivatorLog "invoking control name=$Name"
  Invoke-Control -Element $control
  Start-Sleep -Milliseconds 500
  return $true
}

function Try-InvokeAutomationIdControl {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$AutomationId
  )
  $control = Find-DescendantByAutomationId -Root $Root -AutomationId $AutomationId
  if (-not $control) { return $false }
  Write-ActivatorLog "invoking control automation_id=$AutomationId name=$($control.Current.Name)"
  Invoke-Control -Element $control
  Start-Sleep -Milliseconds 800
  return $true
}

function Wait-ForOpenControlPanel {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Deadline,
    [Parameter(Mandatory = $true)][string]$Source
  )

  do {
    $window = Get-OfficeWindowFromHandle -Handle $WindowHandle
    $button = Find-DescendantByNameLike -Root $window -Name "Open Control Panel" -ProcessId $window.Current.ProcessId
    if ($button) {
      Write-ActivatorLog "found Open Control Panel control source=$Source"
      Invoke-Control -Element $button
      return $true
    }
    Start-Sleep -Milliseconds 500
  } while ((Get-Date) -lt $Deadline)

  Write-ActivatorLog "Open Control Panel did not appear yet source=$Source"
  return $false
}

function Try-EnableOfficeMcpAddin {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Deadline,
    [Parameter(Mandatory = $true)][string]$Source
  )

  $officeProcessId = Get-OfficeProcessIdFromHandle -Handle $WindowHandle
  $addin = Find-DescendantByNameLike -Root $Root -Name "Office MCP Control" -ProcessId $officeProcessId
  if (-not $addin) {
    $addin = Find-GlobalControlByName -Name "Office MCP Control" -ProcessId $officeProcessId
  }
  if (-not $addin) { return $false }

  Write-ActivatorLog "found Office MCP Control source=$Source name=$($addin.Current.Name)"
  Invoke-Control -Element $addin
  Start-Sleep -Milliseconds 800
  if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline $Deadline -Source $Source) { return $true }
  Write-ActivatorLog "Office MCP Control was clicked but Open Control Panel did not appear yet source=$Source"
  return $false
}

function Try-OpenAddinFromCatalog {
  param(
    [Parameter(Mandatory = $true)]$Window,
    [Parameter(Mandatory = $true)]$Deadline
  )

  foreach ($automationId in @("OfficeExtensionsShowAddinFlyout")) {
    if (Try-InvokeAutomationIdControl -Root $Window -AutomationId $automationId) {
      Write-ActivatorLog "catalog fallback invoked automation_id=$automationId"
      Start-Sleep -Milliseconds 800
      $nextWindow = Get-OfficeWindowFromHandle -Handle $driverWindowHandle
      if (Wait-ForOpenControlPanel -WindowHandle $driverWindowHandle -Deadline (Get-Date).AddSeconds(2) -Source "catalog-fallback:$automationId") {
        return $true
      }
      if (Try-EnableOfficeMcpAddin -Root $nextWindow -WindowHandle $driverWindowHandle -Deadline $Deadline -Source "catalog-fallback:$automationId") {
        return $true
      }
      $Window = $nextWindow
    }
  }

  foreach ($name in @("Insert", "Add-ins", "My Add-ins", "Shared Folder", "Office MCP Control", "Add")) {
    if (Try-InvokeNamedControl -Root $Window -Name $name) {
      Write-ActivatorLog "catalog fallback invoked name=$name"
      Start-Sleep -Milliseconds 800
      $nextWindow = Get-OfficeWindowFromHandle -Handle $driverWindowHandle
      if (Wait-ForOpenControlPanel -WindowHandle $driverWindowHandle -Deadline (Get-Date).AddSeconds(2) -Source "catalog-fallback:$name") {
        return $true
      }
      if (Try-EnableOfficeMcpAddin -Root $nextWindow -WindowHandle $driverWindowHandle -Deadline $Deadline -Source "catalog-fallback:$name") {
        return $true
      }
      $Window = $nextWindow
    }
  }
  return $false
}

Invoke-OfficialSideload

$app = Get-OfficeApplication -HostKey $hostKey
$activeDocumentPath = if ([string]::IsNullOrWhiteSpace($script:officialDocumentPath)) { $DocumentPath } else { $script:officialDocumentPath }
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
if (-not [string]::IsNullOrWhiteSpace($script:officialDocumentPath)) {
  try {
    $copyDeadline = (Get-Date).AddSeconds([Math]::Min(5, [Math]::Max(1, $TimeoutSeconds / 4)))
    $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath -Deadline $copyDeadline
  } catch {
    Write-ActivatorLog "official sideload copy was not visible; falling back to original document path=$DocumentPath error=$($_.Exception.Message)"
    $activeDocumentPath = $DocumentPath
    $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath -Deadline $deadline
  }
} else {
  $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath -Deadline $deadline
}

do {
  Start-Sleep -Milliseconds 500
  $window = Get-OfficeWindowFromHandle -Handle $driverWindowHandle
  if (Wait-ForOpenControlPanel -WindowHandle $driverWindowHandle -Deadline (Get-Date).AddSeconds(1) -Source "initial") {
    Write-Output (@{
        activated = $true
        host = $hostKey
        document_path = $activeDocumentPath
        control_name = "Open Control Panel"
      } | ConvertTo-Json -Compress)
    exit 0
  }
  foreach ($tabName in @("Home", "Insert", "Add-ins", "My Add-ins")) {
    if (Try-InvokeNamedControl -Root $window -Name $tabName) {
      if (Wait-ForOpenControlPanel -WindowHandle $driverWindowHandle -Deadline (Get-Date).AddSeconds(2) -Source "tab:$tabName") {
        Write-Output (@{
            activated = $true
            host = $hostKey
            document_path = $activeDocumentPath
            control_name = "Open Control Panel"
            tab_name = $tabName
          } | ConvertTo-Json -Compress)
        exit 0
      }
      $nextWindow = Get-OfficeWindowFromHandle -Handle $driverWindowHandle
      if (Try-EnableOfficeMcpAddin -Root $nextWindow -WindowHandle $driverWindowHandle -Deadline $deadline -Source "tab:$tabName") {
        Write-Output (@{
            activated = $true
            host = $hostKey
            document_path = $activeDocumentPath
            control_name = "Open Control Panel"
            tab_name = $tabName
          } | ConvertTo-Json -Compress)
        exit 0
      }
    }
  }
  if (Try-OpenAddinFromCatalog -Window (Get-OfficeWindowFromHandle -Handle $driverWindowHandle) -Deadline $deadline) {
    Write-Output (@{
        activated = $true
        host = $hostKey
        document_path = $activeDocumentPath
        control_name = "Open Control Panel"
        activation_path = "catalog-fallback"
      } | ConvertTo-Json -Compress)
    exit 0
  }
} while ((Get-Date) -lt $deadline)

$sample = Get-VisibleControlNameSample -Root (Get-OfficeWindowFromHandle -Handle $driverWindowHandle)
Write-ActivatorLog "timed out; visible control sample=$($sample -join ' | ')"

throw "Timed out waiting for the Office MCP Control ribbon button."
