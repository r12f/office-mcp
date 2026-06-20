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
  if ($Path -match '^https?://') { return $Path.Trim().ToLowerInvariant() }
  try { return (Get-Item -LiteralPath $Path -ErrorAction Stop).FullName.ToLowerInvariant() }
  catch { return [System.IO.Path]::GetFullPath($Path).ToLowerInvariant() }
}

Write-ActivatorLog "start host=$hostKey document=$DocumentPath"
$script:officialDocumentPath = $null

function Get-OfficeAppName {
  switch ($hostKey) {
    "word" { return "Word" }
    "excel" { return "Excel" }
    "powerpoint" { return "PowerPoint" }
  }
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

function Invoke-OfficialRegistration {
  if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    Write-ActivatorLog "official registration skipped: manifest path missing"
    return $false
  }
  if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-ActivatorLog "official registration skipped: manifest not found path=$ManifestPath"
    return $false
  }
  $appName = Get-OfficeAppName
  Write-ActivatorLog "official registration start app=$appName manifest=$ManifestPath"
  try {
    $manifestArg = '"' + $ManifestPath.Replace('"', '\"') + '"'
    $command = "npx --yes office-addin-dev-settings register $manifestArg"
    $stdoutPath = Join-Path $env:TEMP "office-mcp-register-$hostKey-$([guid]::NewGuid().ToString('N')).out.log"
    $stderrPath = Join-Path $env:TEMP "office-mcp-register-$hostKey-$([guid]::NewGuid().ToString('N')).err.log"
    $process = Start-Process -FilePath "cmd.exe" -ArgumentList @('/d', '/s', '/c', $command) -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath -PassThru
    if (-not $process.WaitForExit(15000)) {
      $output = @()
      if (Test-Path -LiteralPath $stdoutPath) { $output += Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue }
      if (Test-Path -LiteralPath $stderrPath) { $output += Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue }
      Write-ActivatorLog "official registration timed out; killing process pid=$($process.Id) output=$($output -join ' | ')"
      Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
      return $false
    }
    $exitCode = $process.ExitCode
    $output = @()
    if (Test-Path -LiteralPath $stdoutPath) { $output += Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $stderrPath) { $output += Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue }
    Write-ActivatorLog "official registration exit=$exitCode output=$($output -join ' | ')"
    return ($exitCode -eq 0)
  } catch {
    Write-ActivatorLog "official registration failed: $($_.Exception.Message)"
    Write-ActivatorLog "official registration error detail: $($_ | Out-String)"
    return $false
  }
}

function Invoke-OfficialSideload {
  if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    Write-ActivatorLog "official sideload skipped: manifest path missing"
    return
  }
  if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-ActivatorLog "official sideload skipped: manifest not found path=$ManifestPath"
    return
  }
  $appName = Get-OfficeAppName
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
    $targetName = [System.IO.Path]::GetFileName($Path)
    foreach ($workbook in @($Application.Workbooks)) {
      $workbookFullName = ""
      $workbookName = ""
      try { $workbookFullName = [string]$workbook.FullName } catch {}
      try { $workbookName = [string]$workbook.Name } catch {}
      if ((-not [string]::IsNullOrWhiteSpace($workbookFullName) -and (Get-CanonicalPath -Path $workbookFullName) -eq $target) -or $workbookName -eq $targetName) {
        Write-ActivatorLog "activated excel workbook=$workbookFullName name=$workbookName"
        $workbook.Activate()
        $Application.Visible = $true
        return [IntPtr]$Application.Hwnd
      }
    }
    return Get-ExcelMainWindowHandle -Path $Path
  }
  if ($HostKey -eq "powerpoint") {
    foreach ($presentation in @($Application.Presentations)) {
      if ((Get-CanonicalPath -Path $presentation.FullName) -eq $target) {
        Write-ActivatorLog "activated powerpoint presentation=$($presentation.FullName)"
        try {
          $window = $presentation.Windows.Item(1)
          $window.Activate()
          if ($window.HWND) { return [IntPtr]$window.HWND }
        } catch {
          Write-ActivatorLog "powerpoint presentation window handle unavailable: $($_.Exception.Message)"
        }
        return Get-PowerPointMainWindowHandle -Path $presentation.FullName
      }
    }
  }
  throw "Could not find driver-owned $HostKey document: $target"
}

function Get-ExcelMainWindowHandle {
  param([Parameter(Mandatory = $true)][string]$Path)

  $name = [System.IO.Path]::GetFileName($Path)
  $processes = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
  foreach ($process in $processes) {
    if ($process.MainWindowTitle -like "*$name*") {
      Write-ActivatorLog "using excel main window fallback pid=$($process.Id) title=$($process.MainWindowTitle)"
      return [IntPtr]$process.MainWindowHandle
    }
  }
  throw "Could not find Excel main window for $name."
}

function Get-PowerPointMainWindowHandle {
  param([Parameter(Mandatory = $true)][string]$Path)

  $name = [System.IO.Path]::GetFileName($Path)
  $processes = @(Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
  foreach ($process in $processes) {
    if ($process.MainWindowTitle -like "*$name*") {
      Write-ActivatorLog "using powerpoint main window fallback pid=$($process.Id) title=$($process.MainWindowTitle)"
      return [IntPtr]$process.MainWindowHandle
    }
  }
  if ($processes.Count -eq 1) {
    Write-ActivatorLog "using sole powerpoint main window fallback pid=$($processes[0].Id) title=$($processes[0].MainWindowTitle)"
    return [IntPtr]$processes[0].MainWindowHandle
  }
  throw "Could not find PowerPoint main window for $name."
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
      if ($HostKey -eq "excel") { return Get-ExcelMainWindowHandle -Path $Path }
      Write-ActivatorLog "waiting for driver document path=$Path error=$lastError state=$(Get-OfficeStateSnapshot -Application $Application -HostKey $HostKey)"
      Start-Sleep -Milliseconds 500
    }
  } while ((Get-Date) -lt $Deadline)
  throw $lastError
}

function Close-DriverDocumentIfDifferent {
  param(
    [Parameter(Mandatory = $true)]$Application,
    [Parameter(Mandatory = $true)][string]$HostKey,
    [Parameter(Mandatory = $true)][string]$OriginalPath,
    [Parameter(Mandatory = $true)][string]$ActivePath
  )

  if ((Get-CanonicalPath -Path $OriginalPath) -eq (Get-CanonicalPath -Path $ActivePath)) { return }
  $target = Get-CanonicalPath -Path $OriginalPath
  try {
    if ($HostKey -eq "word") {
      foreach ($document in @($Application.Documents)) {
        if ((Get-CanonicalPath -Path $document.FullName) -eq $target) {
          Write-ActivatorLog "closing original word document after sideload copy=$($document.FullName)"
          $document.Close($false)
        }
      }
    }
    if ($HostKey -eq "excel") {
      foreach ($workbook in @($Application.Workbooks)) {
        if ((Get-CanonicalPath -Path $workbook.FullName) -eq $target) {
          Write-ActivatorLog "closing original excel workbook after sideload copy=$($workbook.FullName)"
          $workbook.Close($false)
        }
      }
    }
    if ($HostKey -eq "powerpoint") {
      foreach ($presentation in @($Application.Presentations)) {
        if ((Get-CanonicalPath -Path $presentation.FullName) -eq $target) {
          Write-ActivatorLog "closing original powerpoint presentation after sideload copy=$($presentation.FullName)"
          $presentation.Close()
        }
      }
    }
  } catch {
    Write-ActivatorLog "failed to close original document after sideload copy: $($_.Exception.Message)"
  }
}

function Open-OfficialSideloadDocument {
  param(
    [Parameter(Mandatory = $true)]$Application,
    [Parameter(Mandatory = $true)][string]$HostKey,
    [Parameter(Mandatory = $true)][string]$Path
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    Write-ActivatorLog "official sideload document open skipped; file not found path=$Path"
    return
  }
  try {
    if ($HostKey -eq "word") {
      $Application.Documents.Open($Path, $false, $false, $false) | Out-Null
    } elseif ($HostKey -eq "excel") {
      $Application.Workbooks.Open($Path) | Out-Null
    } elseif ($HostKey -eq "powerpoint") {
      $Application.Presentations.Open($Path, $false, $false, $true) | Out-Null
    }
    Write-ActivatorLog "official sideload document opened path=$Path"
    return $Application
  } catch {
    Write-ActivatorLog "official sideload document open failed path=$Path error=$($_.Exception.Message)"
    if ($HostKey -eq "excel") {
      try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.Workbooks.Open($Path) | Out-Null
        Write-ActivatorLog "official sideload excel document opened via new application path=$Path"
        return $excel
      } catch {
        Write-ActivatorLog "official sideload excel document new application open failed path=$Path error=$($_.Exception.Message)"
      }
    }
    try {
      Start-Process -FilePath $Path
      Write-ActivatorLog "official sideload document opened via shell path=$Path"
    } catch {
      Write-ActivatorLog "official sideload document shell open failed path=$Path error=$($_.Exception.Message)"
    }
  }
  return $null
}

function Get-ExcelApplicationForWorkbookOpen {
  param($Application)

  foreach ($candidate in @($Application)) {
    if (-not $candidate) { continue }
    try {
      $null = $candidate.Workbooks
      $candidate.Visible = $true
      $candidate.DisplayAlerts = $false
      return $candidate
    } catch {
      Write-ActivatorLog "excel application handle was stale before workbook reopen: $($_.Exception.Message)"
    }
  }

  try {
    $active = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $null = $active.Workbooks
    $active.Visible = $true
    $active.DisplayAlerts = $false
    Write-ActivatorLog "excel application reacquired before workbook reopen"
    return $active
  } catch {
    Write-ActivatorLog "excel active application unavailable before workbook reopen: $($_.Exception.Message)"
  }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  Write-ActivatorLog "excel application created before workbook reopen"
  return $excel
}

function Open-ExcelWorkbookAfterWebExtensionPatch {
  param(
    $Application,
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$Source
  )

  $target = Get-CanonicalPath -Path $Path
  try {
    foreach ($workbook in @($Application.Workbooks)) {
      $workbookPath = ""
      try { $workbookPath = [string]$workbook.FullName } catch {}
      if (-not [string]::IsNullOrWhiteSpace($workbookPath) -and (Get-CanonicalPath -Path $workbookPath) -eq $target) {
        Write-ActivatorLog "excel $Source workbook closed before webextension patch path=$Path"
        $workbook.Close($false)
      }
    }
  } catch {
    Write-ActivatorLog "excel $Source workbook close before webextension patch skipped: $($_.Exception.Message)"
  }

  Ensure-ExcelSideloadWebExtension -WorkbookPath $Path
  $excel = Get-ExcelApplicationForWorkbookOpen -Application $Application
  try {
    $excel.Workbooks.Open($Path) | Out-Null
  } catch {
    Write-ActivatorLog "excel $Source workbook reopen failed on current application; creating replacement application path=$Path error=$($_.Exception.Message)"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.Workbooks.Open($Path) | Out-Null
  }
  $excel.Visible = $true
  Write-ActivatorLog "excel $Source workbook reopened after webextension patch path=$Path"
  return $excel
}

function Reopen-ExcelSideloadDocumentAfterWebExtensionPatch {
  param(
    [Parameter(Mandatory = $true)]$Application,
    [Parameter(Mandatory = $true)][string]$Path
  )

  return Open-ExcelWorkbookAfterWebExtensionPatch -Application $Application -Path $Path -Source "sideload"
}

function Try-OpenExcelPatchedDriverWorkbook {
  param(
    [Parameter(Mandatory = $true)]$Application,
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)]$Deadline
  )

  if ($hostKey -ne "excel") { return $null }
  try {
    $Application = Open-ExcelWorkbookAfterWebExtensionPatch -Application $Application -Path $Path -Source "driver"
    return Wait-ForDriverDocument -Application $Application -HostKey "excel" -Path $Path -Deadline $Deadline
  } catch {
    Write-ActivatorLog "excel driver workbook patch/reopen failed path=$Path error=$($_.Exception.Message)"
    return $null
  }
}

function Set-ZipEntryText {
  param(
    [Parameter(Mandatory = $true)]$Zip,
    [Parameter(Mandatory = $true)][string]$EntryName,
    [Parameter(Mandatory = $true)][string]$Text
  )

  $existing = $Zip.GetEntry($EntryName)
  if ($existing) { $existing.Delete() }
  $entry = $Zip.CreateEntry($EntryName)
  $stream = $entry.Open()
  try {
    $encoding = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($stream, $encoding)
    try { $writer.Write($Text) } finally { $writer.Dispose() }
  } finally {
    $stream.Dispose()
  }
}

function Get-ZipEntryText {
  param(
    [Parameter(Mandatory = $true)]$Zip,
    [Parameter(Mandatory = $true)][string]$EntryName
  )

  $entry = $Zip.GetEntry($EntryName)
  if (-not $entry) { return "" }
  $stream = $entry.Open()
  try {
    $reader = New-Object System.IO.StreamReader($stream)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
  } finally {
    $stream.Dispose()
  }
}

function Ensure-XmlBeforeCloseTag {
  param(
    [Parameter(Mandatory = $true)][string]$Xml,
    [Parameter(Mandatory = $true)][string]$Needle,
    [Parameter(Mandatory = $true)][string]$Insert,
    [Parameter(Mandatory = $true)][string]$CloseTag
  )

  if ($Xml -like "*$Needle*") { return $Xml }
  return $Xml.Replace($CloseTag, "$Insert`r`n$CloseTag")
}

function Ensure-ExcelSideloadWebExtension {
  param([Parameter(Mandatory = $true)][string]$WorkbookPath)

  if ($hostKey -ne "excel") { return }
  if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    Write-ActivatorLog "excel webextension injection skipped; workbook not found path=$WorkbookPath"
    return
  }
  if ([string]::IsNullOrWhiteSpace($ManifestPath) -or -not (Test-Path -LiteralPath $ManifestPath)) {
    Write-ActivatorLog "excel webextension injection skipped; manifest not found path=$ManifestPath"
    return
  }

  try {
    [xml]$manifest = Get-Content -LiteralPath $ManifestPath -Raw
    $addinId = [string]$manifest.OfficeApp.Id
    $addinVersion = [string]$manifest.OfficeApp.Version
    $sourceLocation = [string]$manifest.OfficeApp.DefaultSettings.SourceLocation.DefaultValue
    if ([string]::IsNullOrWhiteSpace($addinId)) { throw "Manifest Id is missing." }
    if ([string]::IsNullOrWhiteSpace($addinVersion)) { $addinVersion = "1.0.0.0" }
    if ([string]::IsNullOrWhiteSpace($sourceLocation)) { throw "Manifest SourceLocation is missing." }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($WorkbookPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
      $contentTypes = Get-ZipEntryText -Zip $zip -EntryName "[Content_Types].xml"
      $contentTypes = Ensure-XmlBeforeCloseTag -Xml $contentTypes -Needle "application/vnd.ms-office.webextensiontaskpanes+xml" -Insert '  <Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>' -CloseTag "</Types>"
      $contentTypes = Ensure-XmlBeforeCloseTag -Xml $contentTypes -Needle "application/vnd.ms-office.webextension+xml" -Insert '  <Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>' -CloseTag "</Types>"
      Set-ZipEntryText -Zip $zip -EntryName "[Content_Types].xml" -Text $contentTypes

      $workbookRels = Get-ZipEntryText -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"
      if ([string]::IsNullOrWhiteSpace($workbookRels)) {
        $workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
      }
      $workbookRels = Ensure-XmlBeforeCloseTag -Xml $workbookRels -Needle "relationships/webextensiontaskpanes" -Insert '  <Relationship Id="rIdOfficeMcpTaskpanes" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="webextensions/taskpanes.xml"/>' -CloseTag "</Relationships>"
      Set-ZipEntryText -Zip $zip -EntryName "xl/_rels/workbook.xml.rels" -Text $workbookRels

      $taskpanes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:taskpane dockstate="right" visibility="1" width="350" row="4">
    <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdOfficeMcpWebExtension"/>
  </wetp:taskpane>
</wetp:taskpanes>
'@
      Set-ZipEntryText -Zip $zip -EntryName "xl/webextensions/taskpanes.xml" -Text $taskpanes

      $taskpaneRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOfficeMcpWebExtension" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>
</Relationships>
'@
      Set-ZipEntryText -Zip $zip -EntryName "xl/webextensions/_rels/taskpanes.xml.rels" -Text $taskpaneRels

      $escapedManifestPath = [System.Security.SecurityElement]::Escape((Get-Item -LiteralPath $ManifestPath).FullName)
      $escapedSourceLocation = [System.Security.SecurityElement]::Escape($sourceLocation)
      $escapedAddinId = [System.Security.SecurityElement]::Escape($addinId)
      $escapedVersion = [System.Security.SecurityElement]::Escape($addinVersion)
      $webExtension = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{$escapedAddinId}">
  <we:reference id="$escapedAddinId" version="$escapedVersion" store="$escapedManifestPath" storeType="FileSystem"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
    <we:property name="Office.AutoShowTaskpaneWithDocument.Url" value="$escapedSourceLocation"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot/>
</we:webextension>
"@
      Set-ZipEntryText -Zip $zip -EntryName "xl/webextensions/webextension1.xml" -Text $webExtension
    } finally {
      $zip.Dispose()
    }
    Write-ActivatorLog "excel webextension injection completed path=$WorkbookPath manifest=$ManifestPath"
  } catch {
    Write-ActivatorLog "excel webextension injection failed path=$WorkbookPath error=$($_.Exception.Message)"
  }
}

function Write-ActivationResult {
  param(
    [Parameter(Mandatory = $true)][string]$DocumentPath,
    [string]$ControlName = "",
    [string]$TabName = "",
    [string]$ActivationPath = "",
    [bool]$ControlOpened = $false
  )

  $result = @{
    activated = $true
    host = $hostKey
    document_path = $DocumentPath
    control_opened = $ControlOpened
  }
  if (-not [string]::IsNullOrWhiteSpace($ControlName)) { $result.control_name = $ControlName }
  if (-not [string]::IsNullOrWhiteSpace($TabName)) { $result.tab_name = $TabName }
  if (-not [string]::IsNullOrWhiteSpace($ActivationPath)) { $result.activation_path = $ActivationPath }
  Write-Output ($result | ConvertTo-Json -Compress)
}

function Test-ActivatorDeadline {
  param([Parameter(Mandatory = $true)]$Deadline)
  if ((Get-Date) -ge $Deadline) {
    Write-ActivatorLog "activator deadline reached"
    return $false
  }
  return $true
}

function Try-OpenControlPanelForDriverDocument {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Deadline,
    [switch]$AllowCatalogFallback
  )

  do {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { break }
    Start-Sleep -Milliseconds 200
    $window = Get-OfficeWindowFromHandle -Handle $WindowHandle
    if ($hostKey -ne "excel" -and (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(1) -Source "initial")) {
      return @{ opened = $true; control_name = "Open Control Panel"; tab_name = ""; activation_path = "" }
    }
    $tabNames = @("Add-ins", "Home", "Insert", "My Add-ins")
    if ($hostKey -eq "excel" -and -not $AllowCatalogFallback) {
      $tabNames = @("Add-ins")
    }
    foreach ($tabName in $tabNames) {
      if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { break }
      if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(1) -Source "current:$tabName") {
        return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "" }
      }
      if (Try-EnableOfficeMcpAddin -Root $window -WindowHandle $WindowHandle -Deadline $Deadline -Source "current:$tabName") {
        return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "" }
      }
      if (Try-InvokeOfficeAddinsRibbon -WindowHandle $WindowHandle -Window $window -Deadline $Deadline -Source "current:$tabName") {
        $nextWindow = Get-OfficeWindowFromHandle -Handle $WindowHandle
        if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(2) -Source "ribbon:$tabName") {
          return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "" }
        }
        if (Try-EnableOfficeMcpAddin -Root $nextWindow -WindowHandle $WindowHandle -Deadline $Deadline -Source "ribbon:$tabName") {
          return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "" }
        }
        if ($AllowCatalogFallback -and (Try-OpenAddinFromCatalog -WindowHandle $WindowHandle -Window $nextWindow -Deadline $Deadline)) {
          return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "catalog-fallback" }
        }
        $window = $nextWindow
      }
      if (Try-InvokeNamedControl -Root $window -Name $tabName) {
        if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(2) -Source "tab:$tabName") {
          return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "" }
        }
        $nextWindow = Get-OfficeWindowFromHandle -Handle $WindowHandle
        if (Try-EnableOfficeMcpAddin -Root $nextWindow -WindowHandle $WindowHandle -Deadline $Deadline -Source "tab:$tabName") {
          return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "" }
        }
        if ($AllowCatalogFallback -and (Try-OpenAddinFromCatalog -WindowHandle $WindowHandle -Window $nextWindow -Deadline $Deadline)) {
          return @{ opened = $true; control_name = "Open Control Panel"; tab_name = $tabName; activation_path = "catalog-fallback" }
        }
        $window = $nextWindow
      }
    }
    if ($AllowCatalogFallback) {
      if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { break }
      if (Try-OpenAddinFromCatalog -WindowHandle $WindowHandle -Window (Get-OfficeWindowFromHandle -Handle $WindowHandle) -Deadline $Deadline) {
        return @{ opened = $true; control_name = "Open Control Panel"; tab_name = ""; activation_path = "catalog-fallback" }
      }
    }
  } while ((Get-Date) -lt $Deadline)

  $sample = @()
  try {
    $sample = Get-VisibleControlNameSample -Root (Get-OfficeWindowFromHandle -Handle $WindowHandle)
  } catch {
    Write-ActivatorLog "control panel visible control sample failed: $($_.Exception.Message)"
  }
  Write-ActivatorLog "control panel best-effort timed out; visible control sample=$($sample -join ' | ')"
  return @{ opened = $false; control_name = ""; tab_name = ""; activation_path = "" }
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
    [Parameter(Mandatory = $true)][string]$AutomationId,
    [int]$ProcessId = 0
  )

  $condition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::AutomationIdProperty,
    $AutomationId
  )
  foreach ($element in @($Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition))) {
    if (Test-AutomationElementCandidate -Element $element -ProcessId $ProcessId) { return $element }
  }
  return $null
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
      "ControlType.SplitButton",
      "ControlType.TabItem"
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

function Find-GlobalControlByAutomationId {
  param(
    [Parameter(Mandatory = $true)][string]$AutomationId,
    [Parameter(Mandatory = $true)][int]$ProcessId
  )
  return Find-DescendantByAutomationId -Root ([System.Windows.Automation.AutomationElement]::RootElement) -AutomationId $AutomationId -ProcessId $ProcessId
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

function Send-ActivatorKey {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)][string]$Key
  )

  Add-Type -Namespace NativeMethods -Name Window -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetForegroundWindow(System.IntPtr hWnd);
"@ -ErrorAction SilentlyContinue
  Add-Type -Namespace NativeMethods -Name Keyboard -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
"@ -ErrorAction SilentlyContinue
  try { [NativeMethods.Window]::SetForegroundWindow($WindowHandle) | Out-Null } catch {}
  if ($Key -eq "{ESC}") {
    [NativeMethods.Keyboard]::keybd_event(0x1B, 0, 0, 0)
    [NativeMethods.Keyboard]::keybd_event(0x1B, 0, 0x0002, 0)
  } else {
    throw "Unsupported activator key: $Key"
  }
  Start-Sleep -Milliseconds 800
}

function Try-InvokeNamedControl {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$Name
  )
  $processId = 0
  try { $processId = [int]$Root.Current.ProcessId } catch {}
  $control = Find-DescendantByNameLike -Root $Root -Name $Name -ProcessId $processId
  if (-not $control) {
    $control = Find-DescendantByNameLike -Root $Root -Name $Name
  }
  if (-not $control) { return $false }
  Write-ActivatorLog "invoking control name=$Name"
  Invoke-Control -Element $control
  Start-Sleep -Milliseconds 500
  return $true
}

function Try-InvokeAutomationIdControl {
  param(
    [Parameter(Mandatory = $true)]$Root,
    [Parameter(Mandatory = $true)][string]$AutomationId,
    [int]$ProcessId = 0
  )
  $control = Find-DescendantByAutomationId -Root $Root -AutomationId $AutomationId -ProcessId $ProcessId
  if (-not $control) { return $false }
  Write-ActivatorLog "invoking control automation_id=$AutomationId name=$($control.Current.Name)"
  Invoke-Control -Element $control
  Start-Sleep -Milliseconds 800
  return $true
}

function Try-InvokeOfficeAddinsRibbon {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Window,
    [Parameter(Mandatory = $true)]$Deadline,
    [Parameter(Mandatory = $true)][string]$Source
  )

  $processId = Get-OfficeProcessIdFromHandle -Handle $WindowHandle
  foreach ($automationId in @("OfficeExtensionsShowAddinFlyout")) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { return $false }
    $control = Find-DescendantByAutomationId -Root $Window -AutomationId $automationId -ProcessId $processId
    if (-not $control) {
      $control = Find-GlobalControlByAutomationId -AutomationId $automationId -ProcessId $processId
    }
    if (-not $control) { continue }
    Write-ActivatorLog "invoking Office Add-ins ribbon automation_id=$automationId name=$($control.Current.Name) source=$Source"
    Invoke-Control -Element $control
    Start-Sleep -Milliseconds 800
    return $true
  }

  foreach ($name in @("Add-ins", "My Add-ins")) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { return $false }
    $control = Find-DescendantByNameLike -Root $Window -Name $name -ProcessId $processId
    if (-not $control) {
      $control = Find-GlobalControlByName -Name $name -ProcessId $processId
    }
    if (-not $control) { continue }
    Write-ActivatorLog "invoking Office Add-ins ribbon name=$name actual_name=$($control.Current.Name) source=$Source"
    Invoke-Control -Element $control
    Start-Sleep -Milliseconds 800
    return $true
  }

  return $false
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
    Start-Sleep -Milliseconds 200
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
    $addin = Find-DescendantByNameLike -Root $Root -Name "Office MCP Control"
  }
  if (-not $addin) {
    $addin = Find-GlobalControlByName -Name "Office MCP Control" -ProcessId $officeProcessId
  }
  if (-not $addin) { return $false }

  Write-ActivatorLog "found Office MCP Control source=$Source name=$($addin.Current.Name)"
  Invoke-Control -Element $addin
  Start-Sleep -Milliseconds 800
  $shortPanelDeadline = (Get-Date).AddSeconds(4)
  if ($shortPanelDeadline -gt $Deadline) { $shortPanelDeadline = $Deadline }
  if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline $shortPanelDeadline -Source $Source) { return $true }
  Write-ActivatorLog "Office MCP Control direct click did not show Open Control Panel; trying catalog confirmation source=$Source"
  $confirmDeadline = (Get-Date).AddSeconds(12)
  if ($confirmDeadline -gt $Deadline) { $confirmDeadline = $Deadline }
  if (Try-ConfirmCatalogAddinInstall -WindowHandle $WindowHandle -Deadline $confirmDeadline -Source $Source) { return $true }
  Write-ActivatorLog "Office MCP Control was clicked but Open Control Panel did not appear yet source=$Source"
  return $false
}

function Try-ConfirmCatalogAddinInstall {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Deadline,
    [Parameter(Mandatory = $true)][string]$Source
  )

  $officeProcessId = Get-OfficeProcessIdFromHandle -Handle $WindowHandle
  foreach ($name in @("Add", "Continue", "Trust this add-in", "Open", "Office MCP Control", "Open Control Panel")) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { break }
    $window = Get-OfficeWindowFromHandle -Handle $WindowHandle
    $control = Find-DescendantByNameLike -Root $window -Name $name -ProcessId $officeProcessId
    if (-not $control) {
      $control = Find-DescendantByNameLike -Root $window -Name $name
    }
    if (-not $control) {
      $control = Find-GlobalControlByName -Name $name -ProcessId $officeProcessId
    }
    if (-not $control) { continue }
    Write-ActivatorLog "catalog install confirm invoked name=$name source=$Source"
    Invoke-Control -Element $control
    Start-Sleep -Milliseconds 1000
    if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(4) -Source "catalog-confirm:$name") { return $true }
    Try-DismissOfficeModalDialog -WindowHandle $WindowHandle -Deadline $Deadline -Source "catalog-confirm:$name" | Out-Null
    Try-DismissCatalogOverlay -WindowHandle $WindowHandle -Deadline $Deadline -Source "catalog-confirm:$name" | Out-Null
    if (Try-OpenControlPanelFromRibbonTabs -WindowHandle $WindowHandle -Deadline $Deadline -Source "catalog-confirm:$name") { return $true }
  }
  return $false
}

function Try-DismissOfficeModalDialog {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Deadline,
    [Parameter(Mandatory = $true)][string]$Source
  )

  $root = [System.Windows.Automation.AutomationElement]::RootElement
  foreach ($message in @("You cannot close Microsoft Excel because a dialog box is open", "You cannot close Microsoft Word because a dialog box is open", "You cannot close Microsoft PowerPoint because a dialog box is open")) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { return $false }
    $dialogText = Find-DescendantByNameLike -Root $root -Name $message
    if (-not $dialogText) { continue }
    foreach ($name in @("OK", "Close")) {
      $control = Find-DescendantByNameLike -Root $root -Name $name
      if (-not $control) { continue }
      Write-ActivatorLog "office modal dialog dismissed name=$name source=$Source"
      Invoke-Control -Element $control
      Start-Sleep -Milliseconds 800
      return $true
    }
  }
  return $false
}

function Try-DismissCatalogOverlay {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Deadline,
    [Parameter(Mandatory = $true)][string]$Source
  )

  Write-ActivatorLog "catalog overlay dismiss sending escape source=$Source"
  Send-ActivatorKey -WindowHandle $WindowHandle -Key "{ESC}"
  Write-ActivatorLog "catalog overlay dismiss sent escape source=$Source"
  Send-ActivatorKey -WindowHandle $WindowHandle -Key "{ESC}"
  Write-ActivatorLog "catalog overlay dismiss sent second escape source=$Source"
  $window = Get-OfficeWindowFromHandle -Handle $WindowHandle
  $officeProcessId = [int]$window.Current.ProcessId
  foreach ($name in @("Back", "Cancel", "Done")) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { return $false }
    $control = Find-DescendantByNameLike -Root $window -Name $name -ProcessId $officeProcessId
    if (-not $control) {
      $control = Find-DescendantByNameLike -Root $window -Name $name
    }
    if (-not $control) { continue }
    Write-ActivatorLog "catalog overlay dismiss invoked name=$name source=$Source"
    Invoke-Control -Element $control
    Start-Sleep -Milliseconds 800
    return $true
  }
  return $false
}

function Try-OpenControlPanelFromRibbonTabs {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Deadline,
    [Parameter(Mandatory = $true)][string]$Source
  )

  $postCatalogTabNames = @("Home", "Insert", "Add-ins", "My Add-ins")
  foreach ($tabName in $postCatalogTabNames) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { break }
    $window = Get-OfficeWindowFromHandle -Handle $WindowHandle
    if (Try-InvokeOfficeAddinsRibbon -WindowHandle $WindowHandle -Window $window -Deadline $Deadline -Source "$Source:ribbon:$tabName") {
      Write-ActivatorLog "post-catalog ribbon scan invoked name=$tabName source=$Source"
      if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(3) -Source "$Source:ribbon:$tabName") { return $true }
    }
    if (-not (Try-InvokeNamedControl -Root $window -Name $tabName)) { continue }
    Write-ActivatorLog "post-catalog tab scan invoked name=$tabName source=$Source"
    if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(3) -Source "$Source:tab:$tabName") { return $true }
  }
  return $false
}

function Try-OpenAddinFromCatalog {
  param(
    [Parameter(Mandatory = $true)][IntPtr]$WindowHandle,
    [Parameter(Mandatory = $true)]$Window,
    [Parameter(Mandatory = $true)]$Deadline
  )

  foreach ($automationId in @("OfficeExtensionsShowAddinFlyout")) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { break }
    $processId = Get-OfficeProcessIdFromHandle -Handle $WindowHandle
    $control = Find-DescendantByAutomationId -Root $Window -AutomationId $automationId -ProcessId $processId
    if (-not $control) {
      $control = Find-GlobalControlByAutomationId -AutomationId $automationId -ProcessId $processId
    }
    if ($control) {
      Write-ActivatorLog "catalog fallback invoked automation_id=$automationId"
      Invoke-Control -Element $control
      Start-Sleep -Milliseconds 800
      $nextWindow = Get-OfficeWindowFromHandle -Handle $WindowHandle
      if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(2) -Source "catalog-fallback:$automationId") {
        return $true
      }
      if (Try-EnableOfficeMcpAddin -Root $nextWindow -WindowHandle $WindowHandle -Deadline $Deadline -Source "catalog-fallback:$automationId") {
        return $true
      }
      $Window = $nextWindow
    }
  }

  foreach ($name in @("Insert", "Add-ins", "My Add-ins", "Shared Folder", "Office MCP Control", "Add")) {
    if (-not (Test-ActivatorDeadline -Deadline $Deadline)) { break }
    if (Try-InvokeNamedControl -Root $Window -Name $name) {
      Write-ActivatorLog "catalog fallback invoked name=$name"
      Start-Sleep -Milliseconds 800
      $nextWindow = Get-OfficeWindowFromHandle -Handle $WindowHandle
      if (Wait-ForOpenControlPanel -WindowHandle $WindowHandle -Deadline (Get-Date).AddSeconds(2) -Source "catalog-fallback:$name") {
        return $true
      }
      if (Try-EnableOfficeMcpAddin -Root $nextWindow -WindowHandle $WindowHandle -Deadline $Deadline -Source "catalog-fallback:$name") {
        return $true
      }
      $Window = $nextWindow
    }
  }
  return $false
}

$app = Get-OfficeApplication -HostKey $hostKey
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$registered = Invoke-OfficialRegistration
if ($registered) {
  try {
    $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $DocumentPath -Deadline $deadline
    Write-ActivatorLog "registered ribbon path active; attempting to open control panel document=$DocumentPath"
    $panel = Try-OpenControlPanelForDriverDocument -WindowHandle $driverWindowHandle -Deadline (Get-Date).AddSeconds([Math]::Min(12, [Math]::Max(3, $TimeoutSeconds / 2))) -AllowCatalogFallback:$false
    if ($panel.opened) {
      Write-ActivationResult -DocumentPath $DocumentPath -ControlName $panel.control_name -TabName $panel.tab_name -ActivationPath "official-registration" -ControlOpened $true
      exit 0
    }
  } catch {
    Write-ActivatorLog "registered ribbon path failed; falling back to sideload document=$DocumentPath error=$($_.Exception.Message)"
  }
}

if ($hostKey -eq "excel") {
  $patchedDriverWindowHandle = Try-OpenExcelPatchedDriverWorkbook -Application $app -Path $DocumentPath -Deadline $deadline
  if ($patchedDriverWindowHandle) {
    Write-ActivatorLog "excel patched driver workbook active; attempting to open control panel document=$DocumentPath"
    $panel = Try-OpenControlPanelForDriverDocument -WindowHandle $patchedDriverWindowHandle -Deadline $deadline -AllowCatalogFallback:$false
    if ($panel.opened) {
      Write-ActivationResult -DocumentPath $DocumentPath -ControlName $panel.control_name -TabName $panel.tab_name -ActivationPath "patched-driver-workbook" -ControlOpened $true
      exit 0
    }
    Write-ActivatorLog "excel patched driver workbook control panel did not open; skipping official sideload fallback to avoid duplicate Excel windows document=$DocumentPath"
    throw "Excel patched driver workbook did not open Office MCP Control."
  }
}

Invoke-OfficialSideload

$activeDocumentPath = if ([string]::IsNullOrWhiteSpace($script:officialDocumentPath)) { $DocumentPath } else { $script:officialDocumentPath }
if (-not [string]::IsNullOrWhiteSpace($script:officialDocumentPath)) {
  try {
    $copyWaitSeconds = [Math]::Min(8, [Math]::Max(3, $TimeoutSeconds / 3))
    if ($hostKey -eq "excel") {
      $app = Reopen-ExcelSideloadDocumentAfterWebExtensionPatch -Application $app -Path $activeDocumentPath
    } else {
      Ensure-ExcelSideloadWebExtension -WorkbookPath $activeDocumentPath
    }
    $copyDeadline = (Get-Date).AddSeconds($copyWaitSeconds)
    try {
      $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath -Deadline $copyDeadline
    } catch {
      Write-ActivatorLog "official sideload copy not active yet; opening document path=$activeDocumentPath error=$($_.Exception.Message)"
      $openedApplication = Open-OfficialSideloadDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath
      if ($openedApplication) { $app = $openedApplication }
      $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath -Deadline $copyDeadline
    }
    Close-DriverDocumentIfDifferent -Application $app -HostKey $hostKey -OriginalPath $DocumentPath -ActivePath $activeDocumentPath
    Write-ActivatorLog "official sideload copy is active; attempting to open control panel document=$activeDocumentPath"
    $panel = Try-OpenControlPanelForDriverDocument -WindowHandle $driverWindowHandle -Deadline $deadline -AllowCatalogFallback
    $activationPath = if ([string]::IsNullOrWhiteSpace($panel.activation_path)) { "official-sideload" } else { "official-sideload:$($panel.activation_path)" }
    if ($panel.opened) {
      Write-ActivationResult -DocumentPath $activeDocumentPath -ControlName $panel.control_name -TabName $panel.tab_name -ActivationPath $activationPath -ControlOpened $true
      exit 0
    }
    Write-ActivatorLog "official sideload control panel did not open; continuing activation fallback document=$activeDocumentPath"
  } catch {
    Write-ActivatorLog "official sideload copy was not visible; falling back to original document path=$DocumentPath error=$($_.Exception.Message)"
    $activeDocumentPath = $DocumentPath
    $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath -Deadline $deadline
  }
} else {
  $driverWindowHandle = Wait-ForDriverDocument -Application $app -HostKey $hostKey -Path $activeDocumentPath -Deadline $deadline
}

$panel = Try-OpenControlPanelForDriverDocument -WindowHandle $driverWindowHandle -Deadline $deadline -AllowCatalogFallback
if ($panel.opened) {
  Write-ActivationResult -DocumentPath $activeDocumentPath -ControlName $panel.control_name -TabName $panel.tab_name -ActivationPath $panel.activation_path -ControlOpened $true
  exit 0
}

throw "Timed out waiting for the Office MCP Control ribbon button."
