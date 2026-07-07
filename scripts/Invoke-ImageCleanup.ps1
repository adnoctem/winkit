#Requires -Version 5.0

<#
.SYNOPSIS
  Runs safe, targeted disk and component cleanup operations.

.DESCRIPTION
  Wraps several cleanup operations into a single guarded script:

  - DISM component cleanup (StartComponentCleanup /ResetBase)
  - Windows Update download cache cleanup
  - Temporary file cleanup (%TEMP%, Windows\Temp, Prefetch)

  Each operation is gated by a -Skip* switch. No global recursive
  filesystem sweeps are performed - only explicit, known-safe paths.

  Requires administrator elevation.

.PARAMETER SkipDism
  Skip DISM component store cleanup.

.PARAMETER SkipWindowsUpdate
  Skip Windows Update download cache cleanup.

.PARAMETER SkipTemp
  Skip temporary file and Prefetch cleanup.

.PARAMETER DryRun
  Preview which operations would run without executing them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Invoke-ImageCleanup.ps1

.EXAMPLE
  PS> ./Invoke-ImageCleanup.ps1 -SkipTemp -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [switch]
  $SkipDism,

  [Parameter(Mandatory = $false)]
  [switch]
  $SkipWindowsUpdate,

  [Parameter(Mandatory = $false)]
  [switch]
  $SkipTemp,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru,

  # Internal: set automatically on elevated re-launch. Not for direct use.
  [switch]
  $Elevated
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if (-not (Test-Elevation)) {
  Request-AdministratorPrivilege `
    -BoundParameters    $PSBoundParameters `
    -ArgumentList       $args `
    -IsElevatedRelaunch:$Elevated
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no cleanup operations will be executed`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# ---- DISM component store cleanup -------------------------------------------
if (-not $SkipDism) {
  Write-Log -Message '==> DISM component store cleanup' -Color Cyan

  if ($DryRun) {
    Write-Log -Message '[DRY RUN] Would run DISM component cleanup.' -Color Yellow
    Add-OperationResult -Results $_results -Target 'DISM-ComponentCleanup' -Source 'ImageCleanup' -Action 'Clean' -Status 'Skipped' -Detail 'DryRun'
  }
  elseif ($PSCmdlet.ShouldProcess('DISM component store', 'Clean up superseded components')) {
    try {
      $null = dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase 2>&1
      Write-Log -Message '    DISM component cleanup complete.' -Color Green
      Add-OperationResult -Results $_results -Target 'DISM-ComponentCleanup' -Source 'ImageCleanup' -Action 'Clean' -Status 'Completed' -Detail '/StartComponentCleanup /ResetBase'
    }
    catch {
      Write-Log -Message "    FAILED: $_" -Color Red
      Add-OperationResult -Results $_results -Target 'DISM-ComponentCleanup' -Source 'ImageCleanup' -Action 'Clean' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

# ---- Windows Update cache cleanup -------------------------------------------
if (-not $SkipWindowsUpdate) {
  Write-Log -Message '==> Windows Update cache cleanup' -Color Cyan
  $_wuDownloadDir = "$env:SystemRoot\SoftwareDistribution\Download"

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would clean $_wuDownloadDir" -Color Yellow
    Add-OperationResult -Results $_results -Target 'WindowsUpdateCache' -Source 'ImageCleanup' -Action 'Clean' -Status 'Skipped' -Detail 'DryRun'
  }
  elseif ($PSCmdlet.ShouldProcess($_wuDownloadDir, 'Remove Windows Update download cache')) {
    if (Test-Path -LiteralPath $_wuDownloadDir) {
      try {
        Remove-Item -LiteralPath "$_wuDownloadDir\*" -Recurse -Force -ErrorAction Stop
        Write-Log -Message '    Windows Update cache cleaned.' -Color Green
        Add-OperationResult -Results $_results -Target 'WindowsUpdateCache' -Source 'ImageCleanup' -Action 'Clean' -Status 'Completed' -Detail $_wuDownloadDir
      }
      catch {
        Write-Log -Message "    FAILED: $_" -Color Red
        Add-OperationResult -Results $_results -Target 'WindowsUpdateCache' -Source 'ImageCleanup' -Action 'Clean' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
    else {
      Write-Log -Message '    Windows Update cache directory not found.' -Color Gray
      Add-OperationResult -Results $_results -Target 'WindowsUpdateCache' -Source 'ImageCleanup' -Action 'Clean' -Status 'Skipped' -Detail 'Directory not found'
    }
  }
}

# ---- Temporary file cleanup -------------------------------------------------
if (-not $SkipTemp) {
  Write-Log -Message '==> Temporary file cleanup' -Color Cyan

  $_tempPaths = @(
    $env:TEMP,
    "$env:SystemRoot\Temp",
    "$env:SystemRoot\Prefetch"
  )

  foreach ($_path in $_tempPaths) {
    if (-not (Test-Path -LiteralPath $_path)) { continue }

    if ($DryRun) {
      Write-Log -Message "[DRY RUN] Would clean $_path" -Color Yellow
      Add-OperationResult -Results $_results -Target $_path -Source 'ImageCleanup' -Action 'Clean' -Status 'Skipped' -Detail 'DryRun'
    }
    elseif ($PSCmdlet.ShouldProcess($_path, 'Remove temporary files')) {
      try {
        Remove-Item -LiteralPath "$_path\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log -Message "    Cleaned: $_path" -Color Green
        Add-OperationResult -Results $_results -Target $_path -Source 'ImageCleanup' -Action 'Clean' -Status 'Completed' -Detail 'Temporary files removed.'
      }
      catch {
        Write-Log -Message "    FAILED ($_path): $_" -Color Red
        Add-OperationResult -Results $_results -Target $_path -Source 'ImageCleanup' -Action 'Clean' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
  }
}

# ---- Summary ----------------------------------------------------------------
$_failedCount = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_completedCount = @($_results | Where-Object { $_.Status -eq 'Completed' }).Count
$_skippedCount = @($_results | Where-Object { $_.Status -ne 'Completed' -and $_.Status -ne 'Failed' }).Count

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no cleanup was performed" -Color Yellow }
else { Write-Log -Message "`nCleanup: $_completedCount completed | $_skippedCount skipped | $_failedCount failed" -Color $(if ($_failedCount -gt 0) { 'Yellow' } else { 'Green' }) }

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Invoke-ImageCleanup'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
