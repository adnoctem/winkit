#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Creates a system restore point for pre-configuration safety.

.DESCRIPTION
  Checks whether System Restore is available and enabled on the system
  drive. If enabled, creates a restore point with the description
  "winkit pre-configuration". If disabled, optionally enables it first.

  Useful as a pre-flight safety step before running larger configuration
  batches. Requires administrator elevation.

.PARAMETER EnableIfNeeded
  Enable System Restore on the system drive if it is currently disabled.
  Without this flag, the script fails clearly when System Restore is
  unavailable.

.PARAMETER Description
  Restore point description. Defaults to "winkit pre-configuration".

.PARAMETER DryRun
  Preview whether a restore point could be created without creating one.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./New-SystemRestorePoint.ps1

.EXAMPLE
  PS> ./New-SystemRestorePoint.ps1 -EnableIfNeeded

.EXAMPLE
  PS> ./New-SystemRestorePoint.ps1 -Description "Before Windows Update" -DryRun

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
  $EnableIfNeeded,

  [Parameter(Mandatory = $false)]
  [string]
  $Description = 'winkit pre-configuration',

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no restore point will be created`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_drive = "$env:SystemDrive\"

Write-Log -Message "Checking System Restore availability on $_drive" -Color Yellow

try {
  $_restoreEnabled = (Get-ComputerRestore).Count -gt 0
}
catch {
  Write-Log -Message 'System Restore is not available on this system.' -Color Red
  Add-OperationResult -Results $_results -Target $_drive -Source 'SystemRestore' -Action 'CreateRestorePoint' -Status 'Failed' -Detail 'System Restore unavailable.'
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

if (-not $_restoreEnabled) {
  if ($EnableIfNeeded) {
    if ($DryRun) {
      Write-Log -Message "[DRY RUN] Would enable System Restore on $_drive" -Color Yellow
    }
    else {
      Write-Log -Message 'System Restore is disabled - enabling...' -Color Yellow
      try {
        Enable-ComputerRestore -Drive $_drive -ErrorAction Stop
        Write-Log -Message '  -> System Restore enabled.' -Color Green
        Add-OperationResult -Results $_results -Target $_drive -Source 'SystemRestore' -Action 'Enable' -Status 'Completed' -Detail 'System Restore enabled.'
      }
      catch {
        Write-Log -Message "  -> FAILED - could not enable System Restore: $_" -Color Red
        Add-OperationResult -Results $_results -Target $_drive -Source 'SystemRestore' -Action 'Enable' -Status 'Failed' -Detail $_.Exception.Message
        if ($PassThru -or $DryRun) { $_results }
        exit 1
      }
    }
  }
  else {
    Write-Log -Message 'System Restore is disabled on the system drive.' -Color Red
    Write-Log -Message 'Re-run with -EnableIfNeeded to enable it, or enable it manually first.' -Color Yellow
    Add-OperationResult -Results $_results -Target $_drive -Source 'SystemRestore' -Action 'CreateRestorePoint' -Status 'Skipped' -Detail 'System Restore disabled; supply -EnableIfNeeded.'
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }
}

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would create restore point: '$Description'" -Color Yellow
  Add-OperationResult -Results $_results -Target $Description -Source 'SystemRestore' -Action 'CreateRestorePoint' -Status 'Skipped' -Detail 'DryRun'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if (-not $PSCmdlet.ShouldProcess($_drive, "Create restore point: '$Description'")) {
  Add-OperationResult -Results $_results -Target $Description -Source 'SystemRestore' -Action 'CreateRestorePoint' -Status 'Skipped' -Detail 'WhatIf'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

Write-Log -Message "Creating restore point: '$Description'..." -Color Yellow
try {
  Checkpoint-Computer -Description $Description -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
  Write-Log -Message "  -> Restore point created successfully." -Color Green
  Add-OperationResult -Results $_results -Target $Description -Source 'SystemRestore' -Action 'CreateRestorePoint' -Status 'Completed' -Detail "Restore point '$Description' created."
}
catch {
  Write-Log -Message "  -> FAILED - could not create restore point: $_" -Color Red
  Add-OperationResult -Results $_results -Target $Description -Source 'SystemRestore' -Action 'CreateRestorePoint' -Status 'Failed' -Detail $_.Exception.Message
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'New-SystemRestorePoint'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
