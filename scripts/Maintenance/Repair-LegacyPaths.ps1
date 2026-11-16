#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.1' }

<#
.SYNOPSIS
  Repairs legacy PSFoundation Office paths and media manifest names.

.DESCRIPTION
  Moves the legacy ProgramData\PSFoundation-Office journal root to
  ProgramData\winkit-Office and renames psfoundation-office-media.json to
  winkit-office-media.json in prepared Office media folders.

  The script refuses to merge when both the legacy and new locations exist.
  Journals that may still need recovery are refused unless -Force is supplied;
  retire them with Move-OfficeRecovery first. Media content and hashes are not
  changed.

  Run once per machine after upgrading. The script is safe to run again: a
  second run reports that there is nothing left to move.

.PARAMETER LegacyLogRoot
  Legacy Office journal/log root. Defaults to ProgramData\PSFoundation-Office.

.PARAMETER LogRoot
  New Office journal/log root. Defaults to ProgramData\winkit-Office.

.PARAMETER MediaPath
  Office media roots to inspect. The root and its immediate subdirectories are
  checked for the legacy manifest. Defaults to <SystemDrive>\Managed\Media.

.PARAMETER ReportDirectory
  Parent for unique repair report directories. Defaults to
  ProgramData\winkit\LegacyPathRepairs. No report is written during -DryRun.

.PARAMETER Force
  Move the legacy root even when it contains recovery journals. Use only after
  retiring or accepting those journals; their recorded paths may no longer
  resolve after the move.

.PARAMETER DryRun
  Show every planned move and refusal without changing paths or writing a report.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> .\Repair-LegacyPaths.ps1 -DryRun
  Preview the legacy Office path and media-manifest moves.

.EXAMPLE
  PS> .\Repair-LegacyPaths.ps1 -PassThru
  Move legacy paths, write a JSON report, and return per-item results.

.EXAMPLE
  PS> .\Repair-LegacyPaths.ps1 -Force -PassThru
  Move the legacy root even if recovery journals are present.

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
  $LegacyLogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

  [ValidateNotNullOrEmpty()]
  [string]
  $LogRoot = (Join-Path $env:ProgramData 'winkit-Office'),

  [ValidateNotNullOrEmpty()]
  [string[]]
  $MediaPath = @(Join-Path $env:SystemDrive 'Managed\Media'),

  [ValidateNotNullOrEmpty()]
  [string]
  $ReportDirectory = (Join-Path $env:ProgramData 'winkit\LegacyPathRepairs'),

  [switch]
  $Force,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force

$ErrorActionPreference = 'Stop'
if ($DryRun) {
  $WhatIfPreference = $true
}

$_startedAt = [DateTime]::UtcNow.ToString('o')
$_runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$_results = New-Object System.Collections.ArrayList
$_legacyRoot = [IO.Path]::GetFullPath($LegacyLogRoot).TrimEnd('\')
$_newRoot = [IO.Path]::GetFullPath($LogRoot).TrimEnd('\')
$_reportRoot = [IO.Path]::GetFullPath($ReportDirectory).TrimEnd('\')
if ($_reportRoot -eq $_legacyRoot -or $_reportRoot.StartsWith($_legacyRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
  throw "ReportDirectory must be outside the legacy Office root: $_legacyRoot"
}
$_mediaRoots = @()
foreach ($_media in $MediaPath) {
  $_mediaRoots += [IO.Path]::GetFullPath($_media).TrimEnd('\')
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

  return @(Get-ChildItem -LiteralPath $Root -File -Force -ErrorAction Stop |
      Where-Object { $_.Name -match '^[0-9a-fA-F]{32}\.json$' })
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
    $Detail,

    [hashtable]
    $Extra
  )

  $_property = @{
    Kind = $Kind
  }
  if ($Extra) {
    foreach ($_key in $Extra.Keys) {
      $_property[$_key] = $Extra[$_key]
    }
  }

  Add-OperationResult -Results $_results -Target $Target -Source 'LegacyPaths' -Action 'Repair' -Status $Status -Detail $Detail -Property $_property
}

# ---- Plan and move the legacy journal/log root ------------------------------
$_legacyExists = Test-Path -LiteralPath $_legacyRoot -ErrorAction Stop
$_newExists = Test-Path -LiteralPath $_newRoot -ErrorAction Stop

if ($_legacyExists -and $_newExists) {
  $_refused++
  Add-LegacyRepairResult -Target $_legacyRoot -Kind 'LogRoot' -Status 'Skipped' -Detail "Refused: both the legacy root '$_legacyRoot' and the new root '$_newRoot' exist. Merge them manually; nothing was moved."
}
elseif ($_legacyExists) {
  if (-not (Test-Path -LiteralPath $_legacyRoot -PathType Container -ErrorAction Stop)) {
    $_refused++
    Add-LegacyRepairResult -Target $_legacyRoot -Kind 'LogRoot' -Status 'Skipped' -Detail 'Refused: the legacy root exists but is not a directory.'
  }
  else {
    $_pendingJournals = @(Get-LegacyJournal -Root $_legacyRoot)
    if ($_pendingJournals.Count -gt 0 -and -not $Force) {
      $_refused++
      $_pendingNames = @($_pendingJournals | ForEach-Object { $_.Name }) -join ', '
      Add-LegacyRepairResult -Target $_legacyRoot -Kind 'LogRoot' -Status 'Skipped' -Detail "Refused: pending recovery journal(s): $_pendingNames. Retire them with Move-OfficeRecovery first, or use -Force."
    }
    else {
      $_planned++
      $_pendingNote = ''
      if ($_pendingJournals.Count -gt 0) {
        $_pendingNames = @($_pendingJournals | ForEach-Object { $_.Name }) -join ', '
        $_pendingNote = " Pending journals moved because -Force was supplied: $_pendingNames."
      }

      if ($DryRun) {
        $_skipped++
        Add-LegacyRepairResult -Target $_newRoot -Kind 'LogRoot' -Status 'Skipped' -Detail "DryRun: would move '$_legacyRoot' to '$_newRoot'.$($_pendingNote)"
      }
      elseif ($PSCmdlet.ShouldProcess($_newRoot, "Move legacy Office root '$_legacyRoot'")) {
        try {
          $_parent = Split-Path -Path $_newRoot -Parent
          if (-not (Test-Path -LiteralPath $_parent -PathType Container -ErrorAction Stop)) {
            $null = New-Item -ItemType Directory -Path $_parent -Force -ErrorAction Stop
          }
          if (Test-Path -LiteralPath $_newRoot -ErrorAction Stop) {
            throw "The new root appeared before the move: $_newRoot"
          }

          Move-Item -LiteralPath $_legacyRoot -Destination $_newRoot -ErrorAction Stop
          $_moved++
          Add-LegacyRepairResult -Target $_newRoot -Kind 'LogRoot' -Status 'Completed' -Detail "Moved '$_legacyRoot' to '$_newRoot'.$($_pendingNote)"
        }
        catch {
          $_failed++
          Add-LegacyRepairResult -Target $_legacyRoot -Kind 'LogRoot' -Status 'Failed' -Detail "Move failed: $($_.Exception.Message)"
        }
      }
      else {
        $_skipped++
        Add-LegacyRepairResult -Target $_newRoot -Kind 'LogRoot' -Status 'Skipped' -Detail 'WhatIf: move was declined.'
      }
    }
  }
}
elseif ($_newExists) {
  $_skipped++
  Add-LegacyRepairResult -Target $_newRoot -Kind 'LogRoot' -Status 'Skipped' -Detail 'Already uses the new Office journal root.'
}
else {
  $_skipped++
  Add-LegacyRepairResult -Target $_legacyRoot -Kind 'LogRoot' -Status 'Skipped' -Detail 'No legacy Office journal root was found.'
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

Add-OperationResult -Results $_results -Target $_newRoot -Source 'LegacyPaths' -Action 'RepairLegacyPaths' -Status $_status -Detail "Legacy path repair finished with $_moved move(s), $_refused refusal(s), and $_failed failure(s)." -Property @{
  RunId        = $_runId
  ManifestPath = $_manifestPath
  Planned      = $_planned
  Moved        = $_moved
  Refused      = $_refused
  Failed       = $_failed
  Skipped      = $_skipped
  Force        = [bool]$Force
  DryRun       = [bool]$DryRun
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
      LegacyLogRoot = $_legacyRoot
      LogRoot       = $_newRoot
      MediaPath     = @($_mediaRoots)
      Force         = [bool]$Force
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
