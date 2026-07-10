Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Reports Google Chrome (and optionally Microsoft Edge) browsing activity
  outside of a configurable workday window.

.DESCRIPTION
  Queries the Chromium-based SQLite history databases for Chrome and, when
  -IncludeEdge is supplied, Microsoft Edge. The script locates each browser
  profile's History file, copies it to an isolated temp directory (because
  Chromium holds an exclusive lock on the live database), and runs a SQL
  query filtered by a lookback window. Rows whose local visit time falls
  outside the supplied -StartHour/-EndHour window on -Workdays, or on any
  day not listed in -Workdays, are returned as off-hours activity.

  When a queried browser is running, the live History database is locked.
  By default the script skips that browser's profiles and notifies the
  user. Supply -Instant to stop the browser so the query can proceed; the
  script restarts the browser after the query completes. -Instant uses a
  graceful close and restores tabs silently via --restore-last-session.
  Add -Force to use a hard stop instead, which surfaces the browser's
  "Restore pages" prompt on relaunch.

  Requires the PSSQLite PowerShell module. When missing, supply
  -EnsureInstalled to install it from PSGallery automatically.

.PARAMETER StartHour
  Workday start hour (0-23). Visits before this hour on a workday count as
  off-hours. Defaults to 8.

.PARAMETER EndHour
  Workday end hour (0-23), exclusive. Visits at or after this hour on a
  workday count as off-hours. Defaults to 17 (yielding a 9-hour workday).

.PARAMETER Workdays
  String array of weekday abbreviations counted as workdays. Any visit on
  a day not in this list is treated as off-hours regardless of time.
  Defaults to Mon,Tue,Wed,Thu,Fri. Accepts any subset of
  Sun,Mon,Tue,Wed,Thu,Fri,Sat.

.PARAMETER DaysBack
  Lookback window in days. Only visits within this window are evaluated.
  Defaults to 90 (~3 months). Use e.g. 28 for the past 4 weeks.

.PARAMETER IncludeEdge
  Also query Microsoft Edge history databases.

.PARAMETER Profile
  Chrome/Edge profile folder name to scan. Defaults to 'Default'.

.PARAMETER AllProfiles
  Scan 'Default' plus every 'Profile *' folder under the browser's User
  Data directory.

.PARAMETER EnsureInstalled
  Install the PSSQLite module from PSGallery when it is not already
  available.

.PARAMETER Instant
  Stop a running queried browser so its history database can be queried.
  Uses a graceful close and restores tabs silently on relaunch. Prompts
  for confirmation before stopping the browser.

.PARAMETER Force
  Requires -Instant. Forces the browser to stop immediately instead of
  closing it gracefully. The browser is restarted without session
  arguments so the 'Restore pages' prompt appears on relaunch.

.PARAMETER ExportCsv
  Optional path to write the off-hours hit list as CSV.

.PARAMETER DryRun
  Preview which profiles would be queried and which browsers would be
  stopped without making changes.

.PARAMETER PassThru
  Return the off-hours history rows as structured objects.

.EXAMPLE
  PS> ./Find-OffHoursActivity.ps1
  Reports Chrome off-hours visits from the Default profile for the past
  90 days, using the 8-17 workday window on Mon-Fri.

.EXAMPLE
  PS> ./Find-OffHoursActivity.ps1 -IncludeEdge -AllProfiles -DaysBack 28
  Reports Chrome and Edge off-hours visits across all profiles for the
  past 4 weeks.

.EXAMPLE
  PS> ./Find-OffHoursActivity.ps1 -StartHour 9 -EndHour 18 -Workdays Mon,Tue,Wed,Thu,Fri,Sat -Instant -EnsureInstalled
  Reports off-hours activity for a 9-18 workday including Saturday,
  stopping Chrome if it is running (with confirmation) and installing
  PSSQLite if missing.

.EXAMPLE
  PS> ./Find-OffHoursActivity.ps1 -IncludeEdge -Instant -Force -DryRun
  Previews which browsers would be force-stopped and which profiles
  would be queried without making changes.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Workdays', Justification = 'Used by nested off-hours query helper through script scope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Profile', Justification = 'Used by nested browser profile helper through script scope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'AllProfiles', Justification = 'Used by nested browser profile helper through script scope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'EnsureInstalled', Justification = 'Used by nested PSSQLite availability helper through script scope.')]
[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [ValidateRange(0, 23)]
  [int]
  $StartHour = 8,

  [Parameter(Mandatory = $false)]
  [ValidateRange(0, 23)]
  [int]
  $EndHour = 17,

  [Parameter(Mandatory = $false)]
  [ValidateSet('Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat')]
  [string[]]
  $Workdays = @('Mon', 'Tue', 'Wed', 'Thu', 'Fri'),

  [Parameter(Mandatory = $false)]
  [ValidateScript({ $_ -gt 0 })]
  [int]
  $DaysBack = 90,

  [Parameter(Mandatory = $false)]
  [switch]
  $IncludeEdge,

  [Parameter(Mandatory = $false)]
  [string]
  $Profile = 'Default',

  [Parameter(Mandatory = $false)]
  [switch]
  $AllProfiles,

  [Parameter(Mandatory = $false)]
  [switch]
  $EnsureInstalled,

  [Parameter(Mandatory = $false)]
  [switch]
  $Instant,

  [Parameter(Mandatory = $false)]
  [switch]
  $Force,

  [Parameter(Mandatory = $false)]
  [string]
  $ExportCsv,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no browsers will be stopped and no queries will run`n" -Color Yellow
}

if ($Force -and -not $Instant) {
  Write-Log -Message '-Force requires -Instant. Re-run with -Instant to stop the browser and query it.' -Color Red
  exit 1
}

if ($StartHour -ge $EndHour) {
  Write-Log -Message "-StartHour ($StartHour) must be less than -EndHour ($EndHour)." -Color Red
  exit 1
}

$_results = New-Object System.Collections.ArrayList
$_hits = New-Object System.Collections.ArrayList
$_instantConfirmed = @{}

# ---- Script-local helpers ----------------------------------------------------

function Test-PSSQLiteAvailable {
  [OutputType([bool])]
  [CmdletBinding()]
  param()

  if (Get-Module -ListAvailable -Name PSSQLite -ErrorAction SilentlyContinue) {
    return $true
  }

  if (-not $EnsureInstalled) {
    return $false
  }

  if ($DryRun) {
    Write-Log -Message '[DRY RUN] Would install PSSQLite from PSGallery.' -Color Yellow
    Add-OperationResult -Results $_results -Target 'PSSQLite' -Source 'PowerShellGallery' -Action 'Install' -Status 'Skipped' -Detail 'DryRun'
    return $false
  }

  Write-Log -Message 'PSSQLite not found - installing from PSGallery...' -Color Yellow
  try {
    $null = Install-Module -Name PSSQLite -Scope CurrentUser -Force -Repository PSGallery -ErrorAction Stop
    Import-Module PSSQLite -Force -ErrorAction Stop
    Write-Log -Message '  -> PSSQLite installed and imported.' -Color Green
    Add-OperationResult -Results $_results -Target 'PSSQLite' -Source 'PowerShellGallery' -Action 'Install' -Status 'Completed' -Detail 'PSSQLite installed from PSGallery.'
    return $true
  }
  catch {
    Write-Log -Message "  -> FAILED - could not install PSSQLite: $_" -Color Red
    Add-OperationResult -Results $_results -Target 'PSSQLite' -Source 'PowerShellGallery' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
    return $false
  }
}

function Get-BrowserUserDataRoot {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser
  )

  switch ($Browser) {
    'Chrome' { return Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Google\Chrome\User Data' }
    'Edge' { return Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\Edge\User Data' }
  }
}

function Get-BrowserExecutablePath {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser
  )

  switch ($Browser) {
    'Chrome' {
      $_candidates = @(
        (Join-Path -Path $env:ProgramFiles -ChildPath 'Google\Chrome\Application\chrome.exe'),
        (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath 'Google\Chrome\Application\chrome.exe')
      )
    }
    'Edge' {
      $_candidates = @(
        (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path -Path $env:ProgramFiles -ChildPath 'Microsoft\Edge\Application\msedge.exe')
      )
    }
  }

  foreach ($_candidate in $_candidates) {
    if (-not [string]::IsNullOrWhiteSpace($_candidate) -and (Test-Path -LiteralPath $_candidate -PathType Leaf)) {
      return $_candidate
    }
  }
  return $null
}

function Get-BrowserProcessName {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser
  )

  switch ($Browser) {
    'Chrome' { return 'chrome' }
    'Edge' { return 'msedge' }
  }
}

function Test-BrowserRunning {
  [OutputType([bool])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser
  )

  $_procName = Get-BrowserProcessName -Browser $Browser
  return $null -ne (Get-Process -Name $_procName -ErrorAction SilentlyContinue)
}

function Get-BrowserProfiles {
  [OutputType([string[]])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser
  )

  $_userDataRoot = Get-BrowserUserDataRoot -Browser $Browser
  if (-not (Test-Path -LiteralPath $_userDataRoot -PathType Container)) {
    return @()
  }

  if (-not $AllProfiles) {
    return @($Profile)
  }

  $_profiles = New-Object System.Collections.Generic.List[string]
  if (Test-Path -LiteralPath (Join-Path -Path $_userDataRoot -ChildPath 'Default') -PathType Container) {
    [void]$_profiles.Add('Default')
  }

  Get-ChildItem -LiteralPath $_userDataRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like 'Profile *' } |
    ForEach-Object { [void]$_profiles.Add($_.Name) }

  return $_profiles.ToArray()
}

function Get-BrowserHistoryPath {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser,

    [Parameter(Mandatory = $true)]
    [string]$ProfileName
  )

  $_userDataRoot = Get-BrowserUserDataRoot -Browser $Browser
  $_historyPath = Join-Path -Path $_userDataRoot -ChildPath "$ProfileName\History"
  if (Test-Path -LiteralPath $_historyPath -PathType Leaf) {
    return $_historyPath
  }
  return $null
}

function Copy-HistoryDatabase {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$HistoryPath,

    [Parameter(Mandatory = $true)]
    [string]$Browser,

    [Parameter(Mandatory = $true)]
    [string]$ProfileName
  )

  $_tempRoot = Join-Path -Path $env:TEMP -ChildPath 'winkit\Find-OffHoursActivity'
  if (-not (Test-Path -LiteralPath $_tempRoot -PathType Container)) {
    $null = New-Item -Path $_tempRoot -ItemType Directory -Force -ErrorAction SilentlyContinue
  }

  $_safeProfile = $ProfileName -replace '[^A-Za-z0-9._-]', '-'
  $_dbFileName = "{0}-{1}.db" -f $Browser, $_safeProfile
  $_destHistory = Join-Path -Path $_tempRoot -ChildPath $_dbFileName

  if ($DryRun) {
    return $_destHistory
  }

  try {
    Copy-Item -LiteralPath $HistoryPath -Destination $_destHistory -Force -ErrorAction Stop
  }
  catch {
    return $null
  }

  foreach ($_sidecar in @('History-journal', 'History-wal', 'History-shm')) {
    $_src = Join-Path -Path (Split-Path -Path $HistoryPath -Parent) -ChildPath $_sidecar
    if (Test-Path -LiteralPath $_src -PathType Leaf) {
      $_destSidecar = $_destHistory + '-' + ($_sidecar -replace '^History-', '')
      Copy-Item -LiteralPath $_src -Destination $_destSidecar -Force -ErrorAction SilentlyContinue
    }
  }

  return $_destHistory
}

function ConvertFrom-ChromeTime {
  [OutputType([datetime])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [long]$ChromeTime
  )

  $_fileTime = [long]($ChromeTime * 10)
  return [datetime]::FromFileTimeUtc($_fileTime).ToLocalTime()
}

function Stop-BrowserProcess {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser,

    [switch]$Force
  )

  $_procName = Get-BrowserProcessName -Browser $Browser
  $_target = "$Browser ($($_procName).exe)"

  if ($DryRun) {
    if ($Force) {
      Write-Log -Message "[DRY RUN] Would force-stop $Browser." -Color Yellow
      Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Skipped' -Detail 'DryRun - force stop.'
    }
    else {
      Write-Log -Message "[DRY RUN] Would gracefully close $Browser." -Color Yellow
      Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Skipped' -Detail 'DryRun - graceful close.'
    }
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_target, 'Stop browser process')) {
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  if ($Force) {
    try {
      Get-Process -Name $_procName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
      $_deadline = (Get-Date).AddSeconds(3)
      while ((Get-Date) -lt $_deadline -and (Test-BrowserRunning -Browser $Browser)) {
        Start-Sleep -Milliseconds 200
      }
      if (Test-BrowserRunning -Browser $Browser) {
        Write-Log -Message "  -> $Browser still running after force stop attempt." -Color Red
        Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Failed' -Detail 'Force stop did not terminate the process.'
        return $false
      }
      Write-Log -Message "  -> $Browser force-stopped." -Color Green
      Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Stopped' -Detail 'Force-killed to enable history query; browser will offer tab restore on relaunch.'
      return $true
    }
    catch {
      Write-Log -Message "  -> FAILED - could not force-stop ${Browser}: $_" -Color Red
      Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Failed' -Detail $_.Exception.Message
      return $false
    }
  }

  $_mainWindows = @(Get-Process -Name $_procName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
  if ($_mainWindows.Count -eq 0) {
    Write-Log -Message "  -> $Browser has no closeable main window; cannot gracefully close." -Color Red
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Failed' -Detail 'No main window to close. Re-run with -Force for a hard stop.'
    return $false
  }

  foreach ($_w in $_mainWindows) {
    $_w.CloseMainWindow() | Out-Null
  }

  $_deadline = (Get-Date).AddSeconds(5)
  while ((Get-Date) -lt $_deadline -and (Test-BrowserRunning -Browser $Browser)) {
    Start-Sleep -Milliseconds 200
  }

  if (Test-BrowserRunning -Browser $Browser) {
    Write-Log -Message "  -> $Browser still running after graceful close timeout." -Color Red
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Failed' -Detail 'Soft close timed out; browser still running. Re-run with -Force for a hard stop, or close the browser manually.'
    return $false
  }

  Write-Log -Message "  -> $Browser closed gracefully." -Color Green
  Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Stop' -Status 'Stopped' -Detail 'Graceful close succeeded; tabs will be restored via --restore-last-session on relaunch.'
  return $true
}

function Start-BrowserProcess {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser,

    [switch]$Force
  )

  $_procName = Get-BrowserProcessName -Browser $Browser
  $_exePath = Get-BrowserExecutablePath -Browser $Browser
  $_target = "$Browser ($($_procName).exe)"

  if (-not $_exePath) {
    Write-Log -Message "  -> Could not locate $Browser executable; skipping relaunch." -Color Yellow
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Start' -Status 'Skipped' -Detail 'Executable not found.'
    return
  }

  if (Test-BrowserRunning -Browser $Browser) {
    Write-Log -Message "  -> $Browser already running; skipping relaunch." -Color Gray
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Start' -Status 'Skipped' -Detail 'Browser already running.'
    return
  }

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would restart $Browser." -Color Yellow
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Start' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_target, 'Start browser process')) {
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Start' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  try {
    if ($Force) {
      $null = Start-Process -FilePath $_exePath -ErrorAction Stop
    }
    else {
      $null = Start-Process -FilePath $_exePath -ArgumentList '--restore-last-session' -ErrorAction Stop
    }
    Write-Log -Message "  -> $Browser restarted." -Color Green
    if ($Force) {
      Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Start' -Status 'Completed' -Detail 'Relaunched; Restore pages prompt should appear.'
    }
    else {
      Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Start' -Status 'Completed' -Detail 'Relaunched with --restore-last-session; tabs restored silently.'
    }
  }
  catch {
    Write-Log -Message "  -> FAILED - could not restart ${Browser}: $_" -Color Red
    Add-OperationResult -Results $_results -Target $_target -Source 'BrowserProcess' -Action 'Start' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-OffHoursQuery {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Chrome', 'Edge')]
    [string]$Browser,

    [Parameter(Mandatory = $true)]
    [string]$ProfileName,

    [Parameter(Mandatory = $true)]
    [string]$DataSourcePath
  )

  $_cutoff = (Get-Date).AddDays(-$DaysBack)
  $_chromeEpochCutoff = [long]([datetimeoffset]::new($_cutoff.ToUniversalTime()).ToUnixTimeSeconds() * 1000000) + 11644473600000000

  $sql = @'
SELECT u.url,
       u.title,
       u.visit_count,
       v.visit_time,
       v.from_visit,
       v.transition
FROM   visits v
JOIN   urls   u ON v.url = u.id
WHERE  v.visit_time >= @cutoff
ORDER  BY v.visit_time DESC;
'@

  try {
    Import-Module PSSQLite -ErrorAction Stop
    $_parameters = @{ cutoff = $_chromeEpochCutoff }
    $_rows = Invoke-SqliteQuery -DataSource $DataSourcePath -Query $sql -SqlParameters $_parameters -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $_results -Target "$Browser\$ProfileName\History" -Source "$($Browser)History" -Action 'Query' -Status 'Failed' -Detail $_.Exception.Message
    return
  }

  foreach ($_row in $_rows) {
    $_localTime = ConvertFrom-ChromeTime -ChromeTime ([long]$_row.visit_time)
    $_dayAbbrev = $_localTime.ToString('ddd', [System.Globalization.CultureInfo]::InvariantCulture).Substring(0, 3)
    if ($Workdays -notcontains $_dayAbbrev) {
      $_isOffHours = $true
    }
    else {
      $_isOffHours = ($_localTime.Hour -lt $StartHour -or $_localTime.Hour -ge $EndHour)
    }

    if (-not $_isOffHours) { continue }

    [void]$_hits.Add([PSCustomObject]@{
        Browser = $Browser
        Profile = $ProfileName
        VisitTime = $_localTime
        DayOfWeek = $_dayAbbrev
        Hour = $_localTime.Hour
        Url = $_row.url
        Title = $_row.title
        VisitCount = $_row.visit_count
        FromVisit = $_row.from_visit
        Transition = $_row.transition
      })
  }
}

# ---- PSSQLite availability ---------------------------------------------------

$_sqliteAvailable = Test-PSSQLiteAvailable
if (-not $_sqliteAvailable) {
  if ($DryRun) {
    Write-Log -Message '[DRY RUN] PSSQLite not available - query steps would be skipped.' -Color Yellow
  }
  else {
    Write-Log -Message 'PSSQLite module not found. Re-run with -EnsureInstalled to install it from PSGallery.' -Color Red
    exit 1
  }
}

# ---- Browser selection -------------------------------------------------------

$_browsers = @('Chrome')
if ($IncludeEdge) { $_browsers += 'Edge' }

# ---- Main loop ---------------------------------------------------------------

foreach ($_browser in $_browsers) {
  $_profiles = Get-BrowserProfiles -Browser $_browser
  if ($_profiles.Count -eq 0) {
    Write-Log -Message "$_browser : User Data directory not found - skipping." -Color Gray
    Add-OperationResult -Results $_results -Target $_browser -Source "$($_browser)History" -Action 'Resolve' -Status 'Skipped' -SkippedReason 'UserDataNotFound'
    continue
  }

  # ---- Phase 1: Resolve existing history files ------------------------------
  $_profileHistoryPaths = [ordered]@{}
  foreach ($_profileName in $_profiles) {
    $_historyPath = Get-BrowserHistoryPath -Browser $_browser -ProfileName $_profileName
    if (-not $_historyPath) {
      Write-Log -Message "  Skipping $_browser\$_profileName : History file not found." -Color Gray
      Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Skipped' -SkippedReason 'ProfileNotFound'
      continue
    }
    $_profileHistoryPaths[$_profileName] = $_historyPath
  }

  if ($_profileHistoryPaths.Count -eq 0) {
    continue
  }

  # ---- Phase 2: Stop browser if running and -Instant is set -----------------
  $_browserWasStopped = $false
  $_running = Test-BrowserRunning -Browser $_browser
  if ($_running) {
    if ($Instant) {
      if (-not $_instantConfirmed.ContainsKey($_browser)) {
        if ($DryRun) {
          $_instantConfirmed[$_browser] = $false
        }
        else {
          $_caption = "Stop $_browser?"
          if ($Force) {
            $_message = "$_browser is running and its history database is locked. Force-close it so the script can query the database? Unsaved form input may be lost. It will be restarted immediately after the history files are copied and the 'Restore pages' prompt will appear so you can recover your tabs."
          }
          else {
            $_message = "$_browser is running and its history database is locked. Close it gracefully so the script can query the database? It will be restarted immediately after the history files are copied with --restore-last-session so your tabs come back automatically."
          }
          $_instantConfirmed[$_browser] = $PSCmdlet.ShouldContinue($_caption, $_message)
        }
      }

      if ($_instantConfirmed[$_browser] -or $DryRun) {
        if ($DryRun) {
          if ($Force) {
            Write-Log -Message "[DRY RUN] Would force-stop $_browser to copy history databases." -Color Yellow
          }
          else {
            Write-Log -Message "[DRY RUN] Would gracefully close $_browser to copy history databases." -Color Yellow
          }
          $_browserWasStopped = $true
        }
        else {
          if (-not (Test-BrowserRunning -Browser $_browser)) {
            $_stopResult = $true
          }
          else {
            $_stopResult = Stop-BrowserProcess -Browser $_browser -Force:$Force
          }
          if ($_stopResult) {
            $_browserWasStopped = $true
          }
          else {
            Write-Log -Message "  Could not stop $_browser - skipping all profiles." -Color Yellow
            foreach ($_profileName in $_profileHistoryPaths.Keys) {
              Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Skipped' -SkippedReason 'DatabaseLocked' -Detail 'Browser could not be stopped; supply -Instant and optionally -Force.'
            }
            continue
          }
        }
      }
      else {
        Write-Log -Message "  User declined to stop $_browser - skipping all profiles." -Color Yellow
        foreach ($_profileName in $_profileHistoryPaths.Keys) {
          Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Skipped' -SkippedReason 'DatabaseLocked' -Detail 'User declined -Instant stop; browser still running.'
        }
        continue
      }
    }
    else {
      Write-Log -Message "  $_browser is running - skipping all profiles. Re-run with -Instant to stop the browser and query it." -Color Yellow
      foreach ($_profileName in $_profileHistoryPaths.Keys) {
        Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Skipped' -SkippedReason 'DatabaseLocked' -Detail 'Browser running; supply -Instant to stop and query.'
      }
      continue
    }
  }

  # ---- Phase 3: Copy all profile history databases to temp ------------------
  $_copiedDbs = [ordered]@{}
  foreach ($_profileName in $_profileHistoryPaths.Keys) {
    $_historyPath = $_profileHistoryPaths[$_profileName]

    if ($DryRun) {
      Write-Log -Message "[DRY RUN] Would copy $_browser\$_profileName\History to temp." -Color Yellow
      $_copiedDbs[$_profileName] = $null
      continue
    }

    $_tempCopy = Copy-HistoryDatabase -HistoryPath $_historyPath -Browser $_browser -ProfileName $_profileName
    if (-not $_tempCopy) {
      Write-Log -Message "  Skipping $_browser\$_profileName : history database copy failed (locked or unreadable)." -Color Yellow
      Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Skipped' -SkippedReason 'DatabaseLocked' -Detail 'History database copy failed.'
      continue
    }
    $_copiedDbs[$_profileName] = $_tempCopy
  }

  if ($_copiedDbs.Count -eq 0) {
    continue
  }

  # ---- Phase 4: Restart browser immediately so the user can browse ----------
  if ($_browserWasStopped -and -not $DryRun) {
    Write-Log -Message "  Restarting $_browser so you can continue browsing..." -Color Yellow
    Start-BrowserProcess -Browser $_browser -Force:$Force
  }
  elseif ($_browserWasStopped -and $DryRun) {
    Write-Log -Message "[DRY RUN] Would restart $_browser after copying history databases." -Color Yellow
  }

  # ---- Phase 5: Query all copied databases ----------------------------------
  $_browserHitCount = 0
  $_browserQueriedProfiles = 0

  foreach ($_profileName in $_copiedDbs.Keys) {
    $_tempCopy = $_copiedDbs[$_profileName]

    if ($DryRun) {
      Write-Log -Message "[DRY RUN] Would query $_browser\$_profileName for the past $DaysBack day(s)." -Color Yellow
      Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Skipped' -Detail 'DryRun'
      $_browserQueriedProfiles++
      continue
    }

    if (-not $_sqliteAvailable) {
      Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Skipped' -SkippedReason 'PSSQLiteUnavailable'
      continue
    }

    if (-not $_tempCopy) {
      continue
    }

    $_beforeHitCount = $_hits.Count
    Invoke-OffHoursQuery -Browser $_browser -ProfileName $_profileName -DataSourcePath $_tempCopy
    $_profileHitCount = $_hits.Count - $_beforeHitCount
    $_browserHitCount += $_profileHitCount
    $_browserQueriedProfiles++

    if ($_profileHitCount -eq 0) {
      Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Completed' -Detail 'NoOffHoursMatches'
    }
    else {
      Add-OperationResult -Results $_results -Target "$_browser\$_profileName\History" -Source "$($_browser)History" -Action 'Query' -Status 'Completed' -Detail "$_profileHitCount off-hours visits."
    }

    Remove-Item -LiteralPath $_tempCopy -Force -ErrorAction SilentlyContinue
    foreach ($_ext in @('-wal', '-shm', '-journal')) {
      $_sidecar = $_tempCopy + $_ext
      if (Test-Path -LiteralPath $_sidecar -PathType Leaf) {
        Remove-Item -LiteralPath $_sidecar -Force -ErrorAction SilentlyContinue
      }
    }
  }

  Write-Log -Message "$_browser : scanned $_browserQueriedProfiles profile(s), $_browserHitCount off-hours visit(s) in the past $DaysBack day(s)." -Color $(if ($_browserHitCount -gt 0) { 'Cyan' } else { 'Gray' })
}

# ---- Optional CSV export -----------------------------------------------------

if ($ExportCsv) {
  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would export $($_hits.Count) hit(s) to '$ExportCsv'." -Color Yellow
    Add-OperationResult -Results $_results -Target $ExportCsv -Source 'Export' -Action 'WriteCsv' -Status 'Skipped' -Detail 'DryRun'
  }
  else {
    try {
      $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportCsv)
      $_exportDir = Split-Path -Path $_exportPath -Parent
      if (-not [string]::IsNullOrWhiteSpace($_exportDir) -and -not (Test-Path -LiteralPath $_exportDir)) {
        $null = New-Item -Path $_exportDir -ItemType Directory -Force -ErrorAction Stop
      }
      if ($_hits.Count -eq 0) {
        '' | Out-File -FilePath $_exportPath -Encoding UTF8
      }
      else {
        $_hits | Export-Csv -Path $_exportPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
      }
      Write-Log -Message "Exported $($_hits.Count) hit(s) to: $_exportPath" -Color Green
      Add-OperationResult -Results $_results -Target $_exportPath -Source 'Export' -Action 'WriteCsv' -Status 'Completed' -Detail "$($_hits.Count) row(s) exported."
    }
    catch {
      Write-Log -Message "FAILED - could not export CSV: $_" -Color Red
      Add-OperationResult -Results $_results -Target $ExportCsv -Source 'Export' -Action 'WriteCsv' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

# ---- Summary -----------------------------------------------------------------

$_totalHits = $_hits.Count
if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no browsers were stopped and no queries were run." -Color Yellow
}
else {
  $_chromeHits = @($_hits | Where-Object { $_.Browser -eq 'Chrome' }).Count
  $_edgeHits = @($_hits | Where-Object { $_.Browser -eq 'Edge' }).Count
  $_summary = "Off-hours activity: Chrome=$_chromeHits"
  if ($IncludeEdge) { $_summary += " | Edge=$_edgeHits" }
  $_summary += " | Total=$_totalHits"
  Write-Log -Message "`n$_summary" -Color $(if ($_totalHits -gt 0) { 'Cyan' } else { 'Green' })
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Find-OffHoursActivity'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_hits
}
