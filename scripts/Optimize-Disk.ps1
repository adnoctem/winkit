#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Optimises fixed disk volumes with TRIM/defrag, volume repair, and
  recycle-bin cleanup.

.DESCRIPTION
  Runs maintenance operations on fixed disk volumes: Optimize-Volume
  (TRIM for SSDs, defrag for HDDs — Windows handles media-type
  detection), Repair-Volume (read-only scan), and Clear-RecycleBin.

  By default only the system drive is targeted. Use -All to include
  every fixed volume. Each operation can be skipped individually with
  a -Skip* switch.

  Designed for scheduled-task maintenance scenarios where a
  bi-weekly or monthly run keeps enterprise disks in good condition.

.PARAMETER All
  Operate on all fixed volumes, not just the system drive.

.PARAMETER SkipDefrag
  Skip Optimize-Volume (TRIM/defrag).

.PARAMETER SkipRepair
  Skip Repair-Volume scan.

.PARAMETER SkipRecycleBin
  Skip Clear-RecycleBin.

.PARAMETER DryRun
  Preview which operations would run without executing them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Optimize-Disk.ps1
  TRIM/defrag and scan the system drive, then empty recycle bin.

.EXAMPLE
  PS> ./Optimize-Disk.ps1 -All
  Optimise all fixed volumes.

.EXAMPLE
  PS> ./Optimize-Disk.ps1 -All -SkipRecycleBin
  Run defrag/repair on all fixed drives without emptying recycle bins.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Operate on all fixed volumes, not just the system drive.'
  )]
  [switch]
  $All,

  [Parameter(Mandatory = $false)]
  [switch]
  $SkipDefrag,

  [Parameter(Mandatory = $false)]
  [switch]
  $SkipRepair,

  [Parameter(Mandatory = $false)]
  [switch]
  $SkipRecycleBin,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no disk operations will be executed`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# ---- Resolve volumes ---------------------------------------------------------
if ($All) {
  $_volumes = Get-Volume -DriveType Fixed -ErrorAction Stop |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_.DriveLetter) } |
    Sort-Object -Property DriveLetter
}
else {
  $_volumes = @(Get-Volume -DriveLetter $env:SystemDrive[0] -ErrorAction Stop |
      Where-Object { $_.DriveType -eq 'Fixed' })
}

if (-not $_volumes -or @($_volumes).Count -eq 0) {
  Write-Log -Message 'No fixed volumes found to optimise.' -Color Yellow
  Add-OperationResult -Results $_results -Target 'Disk' -Source 'OptimizeDisk' -Action 'Scan' -Status 'Skipped' -Detail 'No fixed volumes found.'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

Write-Log -Message "Target volumes: $(($_volumes | ForEach-Object { "$($_.DriveLetter): ($($_.FileSystemLabel))" }) -join ', ')" -Color Cyan

# ---- Optimize-Volume (TRIM / defrag) -----------------------------------------
if (-not $SkipDefrag) {
  Write-Log -Message "`n==> Volume optimisation (TRIM / defrag)" -Color Cyan

  foreach ($_vol in $_volumes) {
    $_driveLetter = "$($_.DriveLetter):"
    $_label = if ($_.FileSystemLabel) { "$_driveLetter ($($_.FileSystemLabel))" } else { $_driveLetter }

    if ($DryRun) {
      Write-Log -Message "[DRY RUN] Would optimise $_label" -Color Yellow
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Optimize' -Status 'Skipped' -Detail 'DryRun'
      continue
    }

    if (-not $PSCmdlet.ShouldProcess($_driveLetter, 'Optimise volume')) {
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Optimize' -Status 'Skipped' -Detail 'WhatIf'
      continue
    }

    try {
      Write-Log -Message "Optimising $_label ..." -Color Yellow
      $null = Optimize-Volume -DriveLetter $_.DriveLetter -ReTrim -ErrorAction Stop
      Write-Log -Message "  -> $_label optimised." -Color Green
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Optimize' -Status 'Completed' -Detail 'Volume TRIM/defrag completed.'
    }
    catch {
      Write-Log -Message "  -> FAILED ($_label): $_" -Color Red
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Optimize' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

# ---- Repair-Volume (scan) ----------------------------------------------------
if (-not $SkipRepair) {
  Write-Log -Message "`n==> Volume repair scan" -Color Cyan

  foreach ($_vol in $_volumes) {
    $_driveLetter = "$($_.DriveLetter):"

    if ($DryRun) {
      Write-Log -Message "[DRY RUN] Would scan $_driveLetter" -Color Yellow
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Repair' -Status 'Skipped' -Detail 'DryRun'
      continue
    }

    if (-not $PSCmdlet.ShouldProcess($_driveLetter, 'Scan volume for errors')) {
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Repair' -Status 'Skipped' -Detail 'WhatIf'
      continue
    }

    try {
      Write-Log -Message "Scanning $_driveLetter ..." -Color Yellow
      $null = Repair-Volume -DriveLetter $_.DriveLetter -Scan -ErrorAction Stop
      Write-Log -Message "  -> $_driveLetter scan complete." -Color Green
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Repair' -Status 'Completed' -Detail 'Volume scan completed.'
    }
    catch {
      Write-Log -Message "  -> FAILED ($_driveLetter): $_" -Color Red
      Add-OperationResult -Results $_results -Target $_driveLetter -Source 'OptimizeDisk' -Action 'Repair' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

# ---- Clear-RecycleBin --------------------------------------------------------
if (-not $SkipRecycleBin) {
  Write-Log -Message "`n==> Recycle bin cleanup" -Color Cyan

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would empty recycle bins for: $(($_volumes | ForEach-Object { "$($_.DriveLetter):" }) -join ', ')" -Color Yellow
    Add-OperationResult -Results $_results -Target 'RecycleBin' -Source 'OptimizeDisk' -Action 'Clean' -Status 'Skipped' -Detail 'DryRun'
  }
  elseif ($PSCmdlet.ShouldProcess('Recycle Bin', 'Empty recycle bin')) {
    try {
      foreach ($_vol in $_volumes) {
        Clear-RecycleBin -DriveLetter $_.DriveLetter -Force -ErrorAction Stop
        Write-Log -Message "  -> $($_vol.DriveLetter):\ recycle bin emptied." -Color Green
      }
      Add-OperationResult -Results $_results -Target 'RecycleBin' -Source 'OptimizeDisk' -Action 'Clean' -Status 'Completed' -Detail 'Recycle bins emptied.'
    }
    catch {
      Write-Log -Message "  -> FAILED: $_" -Color Red
      Add-OperationResult -Results $_results -Target 'RecycleBin' -Source 'OptimizeDisk' -Action 'Clean' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

# ---- Summary -----------------------------------------------------------------
$_failedCount = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_completedCount = @($_results | Where-Object { $_.Status -eq 'Completed' }).Count
$_skippedCount = @($_results | Where-Object { $_.Status -ne 'Completed' -and $_.Status -ne 'Failed' }).Count

if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no disk operations were performed" -Color Yellow
}
else {
  Write-Log -Message "`nDisk optimisation: $_completedCount completed | $_skippedCount skipped | $_failedCount failed" -Color $(if ($_failedCount -gt 0) { 'Yellow' } else { 'Green' })
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Optimize-Disk'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failedCount -gt 0) {
  $global:LASTEXITCODE = 1
}
