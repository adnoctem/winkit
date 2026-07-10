Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Enables or disables Compact OS on the system drive.

.DESCRIPTION
  Compact OS compresses Windows system binaries to save disk space
  (typically 2-4 GB). Uses the built-in compact.exe tool supported
  by Microsoft since Windows 10.

  Requires administrator elevation. A reboot is not required; changes
  take effect immediately for new file writes and over time as the
  system manages compressed state.

.PARAMETER Undo
  Query the Compact OS state without changing it (alias for -Query).

.PARAMETER Query
  Query the current Compact OS state without making changes.

.PARAMETER DryRun
  Preview what would be done without making changes.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Enable-CompactOS.ps1
  Enables Compact OS on the system drive.

.EXAMPLE
  PS> ./Enable-CompactOS.ps1 -Undo
  Disables Compact OS and returns system files to uncompressed state.

.EXAMPLE
  PS> ./Enable-CompactOS.ps1 -Query

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
  $Undo,

  [Parameter(Mandatory = $false)]
  [switch]
  $Query,

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
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# ---- Query current state ----------------------------------------------------
try {
  $_queryResult = & compact.exe /CompactOS:query 2>&1
  $_queryText = $_queryResult -join "`n"
  Write-Log -Message "Current Compact OS state:" -Color Yellow
  Write-Log -Message "  $_queryText" -Color Gray
}
catch {
  Write-Log -Message "Failed to query Compact OS state: $_" -Color Red
  Add-OperationResult -Results $_results -Target 'CompactOS' -Source 'CompactOS' -Action 'Query' -Status 'Failed' -Detail $_.Exception.Message
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

if ($Query) {
  Add-OperationResult -Results $_results -Target 'CompactOS' -Source 'CompactOS' -Action 'Query' -Status 'Completed' -Detail $_queryText.Trim()
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

$_action = if ($Undo) { 'never' } else { 'always' }
$_label = if ($Undo) { 'Disabling' } else { 'Enabling' }

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would set Compact OS to: $_action" -Color Yellow
  Add-OperationResult -Results $_results -Target 'CompactOS' -Source 'CompactOS' -Action 'Set' -Status 'Skipped' -Detail "DryRun - CompactOS:$_action"
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if (-not $PSCmdlet.ShouldProcess("Compact OS", "Set to '$_action'")) {
  Add-OperationResult -Results $_results -Target 'CompactOS' -Source 'CompactOS' -Action 'Set' -Status 'Skipped' -Detail 'WhatIf'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

Write-Log -Message "$_label Compact OS..." -Color Yellow
try {
  $_setResult = & compact.exe /CompactOS:$_action 2>&1
  Write-Log -Message "  -> $($_setResult -join ' ')" -Color Green
  Add-OperationResult -Results $_results -Target 'CompactOS' -Source 'CompactOS' -Action 'Set' -Status 'Completed' -Detail "CompactOS:$_action"
}
catch {
  Write-Log -Message "  -> FAILED: $_" -Color Red
  Add-OperationResult -Results $_results -Target 'CompactOS' -Source 'CompactOS' -Action 'Set' -Status 'Failed' -Detail $_.Exception.Message
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Enable-CompactOS'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
