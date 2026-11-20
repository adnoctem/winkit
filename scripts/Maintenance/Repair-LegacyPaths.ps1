#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.4.1' }

<#
.SYNOPSIS
  Repairs legacy PSFoundation machine and user paths.

.DESCRIPTION
  Moves legacy PSFoundation data, configuration, cache, home, log, Office
  journal, and media-manifest paths to their winkit names. The script covers
  the machine ProgramData paths, the invoking user's profile, and other local
  user profiles when accessible.

  The script refuses to merge when both the legacy and current locations exist.
  Office recovery journals of finished runs move with their folder. A journal of
  an unfinished run, or one that cannot be read, blocks the Office folder: resume
  the run with Install-Office.ps1 or Switch-OfficeVersion.ps1 -Mode Recover and
  -LogRoot pointing at the legacy folder, retire it with -RetirePending and
  -Reason, or move it as it is with -Force. Media content and hashes are not
  changed.

  Run once per machine after upgrading. The script is safe to run again: a
  second run reports that there is nothing left to move.

.PARAMETER ProgramDataRoot
  Machine data root. Defaults to ProgramData. Office paths and the machine
  PSFoundation path are resolved below it.

.PARAMETER UserProfile
  Current user's profile root. Defaults to USERPROFILE.

.PARAMETER ProfilesRoot
  Root containing local user profiles. Other profiles are inspected when
  accessible. Defaults to the parent of USERPROFILE.

.PARAMETER RoamingRoot
  Current user's roaming application-data root. Defaults to APPDATA.

.PARAMETER LocalRoot
  Current user's local application-data root. Defaults to LOCALAPPDATA.

.PARAMETER TempRoot
  Current user's temporary directory. Defaults to TEMP.

.PARAMETER MediaPath
  Office media roots to inspect. The root and its immediate subdirectories are
  checked for the legacy manifest. Defaults to <SystemDrive>\Managed\Media.

.PARAMETER ReportDirectory
  Parent for unique repair report directories. Defaults to
  ProgramData\winkit-reports\LegacyPathRepairs. No report is written during
  -DryRun.

.PARAMETER RetirePending
  Retire every unfinished Office recovery journal in the legacy Office folder
  before moving it, through Substrate's Move-OfficeRecovery. Retirement keeps
  the evidence in the folder's Abandoned subfolder with an audit record; a
  retired run cannot be resumed. Requires -Reason.

.PARAMETER Reason
  Reason recorded in each retirement audit. Required with -RetirePending.

.PARAMETER Force
  Move the legacy Office folder even when it contains unfinished recovery
  journals, without retiring them. Their recorded paths may no longer resolve
  after the move. Prefer -RetirePending.

.PARAMETER DryRun
  Show every planned move and refusal without changing paths or writing a report.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> .\Repair-LegacyPaths.ps1 -DryRun
  Preview every legacy PSFoundation path and media-manifest move.

.EXAMPLE
  PS> .\Repair-LegacyPaths.ps1 -PassThru
  Move legacy paths, write a JSON report, and return per-item results.

.EXAMPLE
  PS> .\Repair-LegacyPaths.ps1 -RetirePending -Reason 'Superseded by the winkit migration' -PassThru
  Retire unfinished Office runs in the legacy folder, then move every legacy path.

.EXAMPLE
  PS> .\Repair-LegacyPaths.ps1 -Force -PassThru
  Move the legacy Office folder even if unfinished recovery journals are present.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - local filesystem operations only.
  SYSTEM-account execution: supported when the target paths are accessible.
  Work is not transactional: each path is moved separately.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [ValidateNotNullOrEmpty()]
  [string]
  $ProgramDataRoot = $env:ProgramData,

  [ValidateNotNullOrEmpty()]
  [string]
  $UserProfile = $env:USERPROFILE,

  [ValidateNotNullOrEmpty()]
  [string]
  $ProfilesRoot = (Split-Path -Path $env:USERPROFILE -Parent),

  [ValidateNotNullOrEmpty()]
  [string]
  $RoamingRoot = $env:APPDATA,

  [ValidateNotNullOrEmpty()]
  [string]
  $LocalRoot = $env:LOCALAPPDATA,

  [ValidateNotNullOrEmpty()]
  [string]
  $TempRoot = $env:TEMP,

  [ValidateNotNullOrEmpty()]
  [string[]]
  $MediaPath = @(Join-Path $env:SystemDrive 'Managed\Media'),

  [ValidateNotNullOrEmpty()]
  [string]
  $ReportDirectory = (Join-Path $env:ProgramData 'winkit-reports\LegacyPathRepairs'),

  [switch]
  $RetirePending,

  [string]
  $Reason,

  [switch]
  $Force,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.4.1 -Force

$ErrorActionPreference = 'Stop'
if ($DryRun) {
  $WhatIfPreference = $true
}

if ($RetirePending -and [string]::IsNullOrWhiteSpace($Reason)) {
  throw '-RetirePending requires -Reason, which is recorded in each retirement audit.'
}

$_startedAt = [DateTime]::UtcNow.ToString('o')
$_runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$_results = New-Object System.Collections.ArrayList
$_programDataRoot = [IO.Path]::GetFullPath($ProgramDataRoot).TrimEnd('\')
$_userProfile = [IO.Path]::GetFullPath($UserProfile).TrimEnd('\')
$_profilesRoot = [IO.Path]::GetFullPath($ProfilesRoot).TrimEnd('\')
$_roamingRoot = [IO.Path]::GetFullPath($RoamingRoot).TrimEnd('\')
$_localRoot = [IO.Path]::GetFullPath($LocalRoot).TrimEnd('\')
$_tempRoot = [IO.Path]::GetFullPath($TempRoot).TrimEnd('\')
$_reportRoot = [IO.Path]::GetFullPath($ReportDirectory).TrimEnd('\')
$_mediaRoots = @()
foreach ($_media in $MediaPath) {
  $_mediaRoots += [IO.Path]::GetFullPath($_media).TrimEnd('\')
}

$_mappings = New-Object System.Collections.ArrayList
$null = $_mappings.Add([PSCustomObject]@{
    Kind        = 'OfficeJournalRoot'
    Legacy      = (Join-Path $_programDataRoot 'PSFoundation-Office')
    Current     = (Join-Path $_programDataRoot 'winkit-Office')
    JournalRoot = $true
  })
$null = $_mappings.Add([PSCustomObject]@{
    Kind        = 'MachineData'
    Legacy      = (Join-Path $_programDataRoot 'PSFoundation')
    Current     = (Join-Path $_programDataRoot 'winkit')
    JournalRoot = $false
  })
$null = $_mappings.Add([PSCustomObject]@{
    Kind        = 'UserHome'
    Legacy      = (Join-Path $_userProfile 'PSFoundation')
    Current     = (Join-Path $_userProfile 'winkit')
    JournalRoot = $false
  })
$null = $_mappings.Add([PSCustomObject]@{
    Kind        = 'UserRoaming'
    Legacy      = (Join-Path $_roamingRoot 'PSFoundation')
    Current     = (Join-Path $_roamingRoot 'winkit')
    JournalRoot = $false
  })
$null = $_mappings.Add([PSCustomObject]@{
    Kind        = 'UserLocal'
    Legacy      = (Join-Path $_localRoot 'PSFoundation')
    Current     = (Join-Path $_localRoot 'winkit')
    JournalRoot = $false
  })
$null = $_mappings.Add([PSCustomObject]@{
    Kind        = 'UserTemp'
    Legacy      = (Join-Path $_tempRoot 'PSFoundation')
    Current     = (Join-Path $_tempRoot 'winkit')
    JournalRoot = $false
  })

if (Test-Path -LiteralPath $_profilesRoot -PathType Container -ErrorAction Stop) {
  try {
    foreach ($_profile in @(Get-ChildItem -LiteralPath $_profilesRoot -Directory -Force -ErrorAction Stop)) {
      if ($_profile.FullName.TrimEnd('\') -ieq $_userProfile -or ($_profile.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        continue
      }

      $null = $_mappings.Add([PSCustomObject]@{
          Kind        = "ProfileHome:$($_profile.Name)"
          Legacy      = (Join-Path $_profile.FullName 'PSFoundation')
          Current     = (Join-Path $_profile.FullName 'winkit')
          JournalRoot = $false
        })
      $null = $_mappings.Add([PSCustomObject]@{
          Kind        = "ProfileRoaming:$($_profile.Name)"
          Legacy      = (Join-Path $_profile.FullName 'AppData\Roaming\PSFoundation')
          Current     = (Join-Path $_profile.FullName 'AppData\Roaming\winkit')
          JournalRoot = $false
        })
      $null = $_mappings.Add([PSCustomObject]@{
          Kind        = "ProfileLocal:$($_profile.Name)"
          Legacy      = (Join-Path $_profile.FullName 'AppData\Local\PSFoundation')
          Current     = (Join-Path $_profile.FullName 'AppData\Local\winkit')
          JournalRoot = $false
        })
      $null = $_mappings.Add([PSCustomObject]@{
          Kind        = "ProfileTemp:$($_profile.Name)"
          Legacy      = (Join-Path $_profile.FullName 'AppData\Local\Temp\PSFoundation')
          Current     = (Join-Path $_profile.FullName 'AppData\Local\Temp\winkit')
          JournalRoot = $false
        })
    }
  }
  catch {
    $null = $_mappings.Add([PSCustomObject]@{
        Kind        = 'ProfilesRoot'
        Legacy      = $_profilesRoot
        Current     = $null
        JournalRoot = $false
        Error       = $_.Exception.Message
      })
  }
}

foreach ($_mapping in $_mappings) {
  if ($_mapping.Legacy -and ($_reportRoot -eq $_mapping.Legacy -or $_reportRoot.StartsWith($_mapping.Legacy + '\', [StringComparison]::OrdinalIgnoreCase))) {
    throw "ReportDirectory must be outside legacy PSFoundation paths: $($_mapping.Legacy)"
  }
}

$_planned = 0
$_moved = 0
$_refused = 0
$_failed = 0
$_skipped = 0
$_manifestPath = $null

function Get-LegacyJournal {
  param (
    [string]
    $Root
  )

  if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
    return @()
  }

  # Every Office run leaves its journal behind. Only a run that did not reach a completed Verify checkpoint, or a
  # journal that cannot be read, may still need recovery; retired runs live in Abandoned and are not listed here.
  $_journals = New-Object System.Collections.ArrayList
  foreach ($_file in @(Get-ChildItem -LiteralPath $Root -File -Force -ErrorAction Stop | Where-Object { $_.Name -match '^[0-9a-fA-F]{32}\.json$' })) {
    $_journal = [PSCustomObject]@{
      RunId          = [IO.Path]::GetFileNameWithoutExtension($_file.Name)
      Path           = $_file.FullName
      Action         = $null
      Phase          = $null
      PhaseCompleted = $false
      Finished       = $false
      Error          = $null
    }

    try {
      $_record = Get-Content -LiteralPath $_file.FullName -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
      $_journal.Action = [string]$_record.Action
      $_journal.Phase = [string]$_record.Phase
      $_journal.PhaseCompleted = $_record.PhaseCompleted -eq $true
      $_journal.Finished = $_journal.Phase -eq 'Verify' -and $_journal.PhaseCompleted
    }
    catch {
      $_journal.Error = $_.Exception.Message
    }

    $null = $_journals.Add($_journal)
  }

  return @($_journals.ToArray())
}

function Format-LegacyJournal {
  param (
    [object]
    $Journal,

    [string]
    $LogRoot
  )

  if ($Journal.Error) {
    return "$($Journal.RunId) (unreadable: $($Journal.Error))"
  }

  # Install and Migrate runs can be resumed by their script; other actions can only be retired.
  $_resume = ''
  $_script = switch ($Journal.Action) {
    'Install' {
      'Install-Office.ps1'
    }
    'Migrate' {
      'Switch-OfficeVersion.ps1'
    }
    default {
      $null
    }
  }
  if ($_script) {
    $_resume = "; resume with .\scripts\Office\$_script -Mode Recover -RunId $($Journal.RunId) -LogRoot '$LogRoot'"
  }

  return "$($Journal.RunId) ($($Journal.Action), stopped in phase $($Journal.Phase)$_resume)"
}

function Add-LegacyRepairResult {
  param (
    [string]
    $Target,

    [string]
    $Kind,

    [string]
    $Status,

    [string]
    $Detail
  )

  Add-OperationResult -Results $_results -Target $Target -Source 'LegacyPaths' -Action 'Repair' -Status $Status -Detail $Detail -Property @{
    Kind = $Kind
  }
}

# ---- Move each legacy path to its current name -----------------------------
foreach ($_mapping in $_mappings) {
  if ($_mapping.PSObject.Properties['Error']) {
    $_failed++
    Add-LegacyRepairResult -Target $_mapping.Legacy -Kind $_mapping.Kind -Status 'Failed' -Detail "Could not enumerate user profiles: $($_mapping.Error)"
    continue
  }

  $_legacyExists = Test-Path -LiteralPath $_mapping.Legacy -ErrorAction Stop
  $_currentExists = Test-Path -LiteralPath $_mapping.Current -ErrorAction Stop

  if ($_legacyExists -and $_currentExists) {
    $_refused++
    Add-LegacyRepairResult -Target $_mapping.Legacy -Kind $_mapping.Kind -Status 'Skipped' -Detail "Refused: both the legacy path '$($_mapping.Legacy)' and the current path '$($_mapping.Current)' exist. Merge them manually; nothing was moved."
    continue
  }

  if (-not $_legacyExists) {
    if ($_currentExists) {
      $_skipped++
      Add-LegacyRepairResult -Target $_mapping.Current -Kind $_mapping.Kind -Status 'Skipped' -Detail 'Already uses the current path.'
    }
    else {
      $_skipped++
      Add-LegacyRepairResult -Target $_mapping.Legacy -Kind $_mapping.Kind -Status 'Skipped' -Detail 'No legacy path was found.'
    }
    continue
  }

  if (-not (Test-Path -LiteralPath $_mapping.Legacy -PathType Container -ErrorAction Stop)) {
    $_refused++
    Add-LegacyRepairResult -Target $_mapping.Legacy -Kind $_mapping.Kind -Status 'Skipped' -Detail 'Refused: the legacy path exists but is not a directory.'
    continue
  }

  $_pendingJournals = @()
  $_retiredNote = ''
  if ($_mapping.JournalRoot) {
    $_pendingJournals = @(Get-LegacyJournal -Root $_mapping.Legacy | Where-Object { -not $_.Finished })

    # Retire unfinished runs through Substrate, which validates and preserves each journal before it moves.
    if ($_pendingJournals.Count -gt 0 -and $RetirePending) {
      $_retired = New-Object System.Collections.ArrayList
      $_retireErrors = New-Object System.Collections.ArrayList
      foreach ($_journal in $_pendingJournals) {
        if ($DryRun) {
          $null = $_retired.Add($_journal.RunId)
          continue
        }

        try {
          $null = Move-OfficeRecovery -RunId $_journal.RunId -Reason $Reason -LogRoot $_mapping.Legacy -Confirm:$false -ErrorAction Stop
          $null = $_retired.Add($_journal.RunId)
        }
        catch {
          $null = $_retireErrors.Add("$($_journal.RunId): $($_.Exception.Message)")
        }
      }

      if ($_retireErrors.Count -gt 0) {
        $_refused++
        Add-LegacyRepairResult -Target $_mapping.Legacy -Kind $_mapping.Kind -Status 'Skipped' -Detail "Refused: unfinished Office run(s) could not be retired: $($_retireErrors -join '; '). Nothing was moved."
        continue
      }

      if ($DryRun) {
        $_retiredNote = " Would retire unfinished Office run(s) first: $($_retired -join ', ')."
      }
      else {
        $_retiredNote = " Retired unfinished Office run(s) first: $($_retired -join ', ')."
        $_pendingJournals = @(Get-LegacyJournal -Root $_mapping.Legacy | Where-Object { -not $_.Finished })
      }
    }

    if ($_pendingJournals.Count -gt 0 -and -not $Force -and -not ($RetirePending -and $DryRun)) {
      $_refused++
      $_pendingNames = @($_pendingJournals | ForEach-Object { Format-LegacyJournal -Journal $_ -LogRoot $_mapping.Legacy }) -join '; '
      Add-LegacyRepairResult -Target $_mapping.Legacy -Kind $_mapping.Kind -Status 'Skipped' -Detail "Refused: unfinished Office run(s) may still need recovery: $_pendingNames. Run again with -RetirePending -Reason '<why>' to retire them, or with -Force to move them as they are."
      continue
    }
  }

  $_planned++
  $_pendingNote = $_retiredNote
  if ($_pendingJournals.Count -gt 0 -and -not ($RetirePending -and $DryRun)) {
    $_pendingNames = @($_pendingJournals | ForEach-Object { $_.RunId }) -join ', '
    $_pendingNote = "$_retiredNote Unfinished journals moved because -Force was supplied: $_pendingNames."
  }

  if ($DryRun) {
    $_skipped++
    Add-LegacyRepairResult -Target $_mapping.Current -Kind $_mapping.Kind -Status 'Skipped' -Detail "DryRun: would move '$($_mapping.Legacy)' to '$($_mapping.Current)'.$($_pendingNote)"
    continue
  }

  if (-not $PSCmdlet.ShouldProcess($_mapping.Current, "Move legacy PSFoundation path '$($_mapping.Legacy)'")) {
    $_skipped++
    Add-LegacyRepairResult -Target $_mapping.Current -Kind $_mapping.Kind -Status 'Skipped' -Detail 'WhatIf: move was declined.'
    continue
  }

  try {
    $_parent = Split-Path -Path $_mapping.Current -Parent
    if (-not (Test-Path -LiteralPath $_parent -PathType Container -ErrorAction Stop)) {
      $null = New-Item -ItemType Directory -Path $_parent -Force -ErrorAction Stop
    }
    if (Test-Path -LiteralPath $_mapping.Current -ErrorAction Stop) {
      throw "The current path appeared before the move: $($_mapping.Current)"
    }

    Move-Item -LiteralPath $_mapping.Legacy -Destination $_mapping.Current -ErrorAction Stop
    $_moved++
    Add-LegacyRepairResult -Target $_mapping.Current -Kind $_mapping.Kind -Status 'Completed' -Detail "Moved '$($_mapping.Legacy)' to '$($_mapping.Current)'.$($_pendingNote)"
  }
  catch {
    $_failed++
    Add-LegacyRepairResult -Target $_mapping.Legacy -Kind $_mapping.Kind -Status 'Failed' -Detail "Move failed: $($_.Exception.Message)"
  }
}

# ---- Rename legacy media manifests -----------------------------------------
$_seenMedia = @{}
foreach ($_mediaRoot in $_mediaRoots) {
  if (-not (Test-Path -LiteralPath $_mediaRoot -PathType Container -ErrorAction Stop)) {
    $_skipped++
    Add-LegacyRepairResult -Target $_mediaRoot -Kind 'MediaRoot' -Status 'Skipped' -Detail 'Media root not found.'
    continue
  }

  $_folders = @($_mediaRoot)
  try {
    foreach ($_directory in @(Get-ChildItem -LiteralPath $_mediaRoot -Directory -Force -ErrorAction Stop)) {
      if ($_directory.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        $_skipped++
        Add-LegacyRepairResult -Target $_directory.FullName -Kind 'MediaRoot' -Status 'Skipped' -Detail 'Reparse-point media directory was not followed.'
        continue
      }
      $_folders += $_directory.FullName
    }
  }
  catch {
    $_failed++
    Add-LegacyRepairResult -Target $_mediaRoot -Kind 'MediaRoot' -Status 'Failed' -Detail "Could not enumerate media directories: $($_.Exception.Message)"
    continue
  }

  foreach ($_folder in $_folders) {
    if ($_seenMedia.ContainsKey($_folder)) {
      continue
    }
    $_seenMedia[$_folder] = $true

    $_legacyManifest = Join-Path $_folder 'psfoundation-office-media.json'
    $_newManifest = Join-Path $_folder 'winkit-office-media.json'
    if (-not (Test-Path -LiteralPath $_legacyManifest -ErrorAction Stop)) {
      if (Test-Path -LiteralPath $_newManifest -ErrorAction Stop) {
        $_skipped++
        Add-LegacyRepairResult -Target $_newManifest -Kind 'MediaManifest' -Status 'Skipped' -Detail 'Already uses the new media manifest name.'
      }
      continue
    }

    if (Test-Path -LiteralPath $_newManifest -ErrorAction Stop) {
      $_refused++
      Add-LegacyRepairResult -Target $_folder -Kind 'MediaManifest' -Status 'Skipped' -Detail 'Refused: both the legacy and new media manifests exist in this folder. Resolve them manually; nothing was renamed.'
      continue
    }

    $_planned++
    if ($DryRun) {
      $_skipped++
      Add-LegacyRepairResult -Target $_newManifest -Kind 'MediaManifest' -Status 'Skipped' -Detail "DryRun: would rename '$_legacyManifest' to '$_newManifest'."
    }
    elseif ($PSCmdlet.ShouldProcess($_newManifest, "Rename legacy Office media manifest '$_legacyManifest'")) {
      try {
        if (Test-Path -LiteralPath $_newManifest -ErrorAction Stop) {
          throw "The new manifest appeared before the rename: $_newManifest"
        }
        Move-Item -LiteralPath $_legacyManifest -Destination $_newManifest -ErrorAction Stop
        $_moved++
        Add-LegacyRepairResult -Target $_newManifest -Kind 'MediaManifest' -Status 'Completed' -Detail "Renamed '$_legacyManifest' to '$_newManifest'."
      }
      catch {
        $_failed++
        Add-LegacyRepairResult -Target $_legacyManifest -Kind 'MediaManifest' -Status 'Failed' -Detail "Rename failed: $($_.Exception.Message)"
      }
    }
    else {
      $_skipped++
      Add-LegacyRepairResult -Target $_newManifest -Kind 'MediaManifest' -Status 'Skipped' -Detail 'WhatIf: rename was declined.'
    }
  }
}

# ---- Summary and report ----------------------------------------------------
if ($_failed -gt 0 -or $_refused -gt 0) {
  $_status = 'Failed'
}
elseif ($DryRun) {
  $_status = 'Preview'
}
elseif ($_moved -gt 0) {
  $_status = 'Completed'
}
else {
  $_status = 'Skipped'
}

if (-not $DryRun) {
  $_manifestPath = Join-Path (Join-Path $_reportRoot $_runId) 'manifest.json'
}

Add-OperationResult -Results $_results -Target 'Legacy PSFoundation paths' -Source 'LegacyPaths' -Action 'RepairLegacyPaths' -Status $_status -Detail "Legacy path repair finished with $_moved move(s), $_refused refusal(s), and $_failed failure(s)." -Property @{
  RunId         = $_runId
  ManifestPath  = $_manifestPath
  Planned       = $_planned
  Moved         = $_moved
  Refused       = $_refused
  Failed        = $_failed
  Skipped       = $_skipped
  Force         = [bool]$Force
  RetirePending = [bool]$RetirePending
  Reason        = $Reason
  DryRun        = [bool]$DryRun
}

if (-not $DryRun) {
  try {
    $_runDirectory = Join-Path $_reportRoot $_runId
    $null = New-Item -ItemType Directory -Path $_runDirectory -Force -ErrorAction Stop
    $_manifest = [ordered]@{
      SchemaVersion = 1
      RunId         = $_runId
      StartedAtUtc  = $_startedAt
      FinishedAtUtc = [DateTime]::UtcNow.ToString('o')
      Actor         = [Security.Principal.WindowsIdentity]::GetCurrent().Name
      Machine       = $env:COMPUTERNAME
      ProgramData   = $_programDataRoot
      UserProfile   = $_userProfile
      ProfilesRoot  = $_profilesRoot
      MediaPath     = @($_mediaRoots)
      Mappings      = @($_mappings | ForEach-Object {
          [PSCustomObject]@{
            Kind        = $_.Kind
            Legacy      = $_.Legacy
            Current     = $_.Current
            JournalRoot = $_.JournalRoot
          }
        })
      Force         = [bool]$Force
      RetirePending = [bool]$RetirePending
      Reason        = $Reason
      Status        = $_status
      Results       = @($_results.ToArray())
      Planned       = $_planned
      Moved         = $_moved
      Refused       = $_refused
      Failed        = $_failed
      Skipped       = $_skipped
    }
    $_json = ConvertTo-Json -InputObject $_manifest -Depth 12 -ErrorAction Stop
    [IO.File]::WriteAllText($_manifestPath, $_json, (New-Object Text.UTF8Encoding($true)))
    $null = Write-OperationResultLog -Results $_results -Path (Join-Path $_runDirectory 'operations.jsonl') -ScriptName 'Repair-LegacyPaths' -RunId $_runId
  }
  catch {
    $_status = 'Failed'
    $_failed++
    $_manifestPath = $null
    Add-OperationResult -Results $_results -Target $_reportRoot -Source 'LegacyPaths' -Action 'Report' -Status 'Failed' -Detail "Could not write the repair report: $($_.Exception.Message)"
    Write-Warning "Could not write the legacy path repair report: $($_.Exception.Message)"
  }
}

Write-Log -Message "Legacy path repair: $_status | Moved: $_moved | Refused: $_refused | Failed: $_failed | Report: $_manifestPath" -Color $(if ($_failed -gt 0 -or $_refused -gt 0) { 'Yellow' } else { 'Green' })
if ($PassThru -or $WhatIfPreference) {
  $_results
}
if ($_failed -gt 0 -or $_refused -gt 0) {
  exit 1
}
exit 0
