#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Clears Windows event logs with explicit confirmation.

.DESCRIPTION
  Clears the specified event logs using wevtutil (the cross-version tool -
  the Clear-EventLog cmdlet does not exist in PowerShell 7). Each log is
  reported with its pre-clear record count, and the operation requires
  explicit confirmation unless -Force is supplied.

  This is a deliberate, standalone operation - it is intentionally NOT part of
  Invoke-ImageCleanup.ps1: clearing event logs removes diagnostic and audit
  trail history, which is a consideration for environments with audit/ISO
  requirements. Only clear logs you explicitly intend to clear.

  Requires administrator elevation.

.PARAMETER LogName
  Event log(s) to clear. Defaults to Application and System.

.PARAMETER Force
  Skip the explicit confirmation prompt.

.PARAMETER DryRun
  Preview which logs would be cleared without clearing anything.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Clear-EventLogs.ps1
  Confirms, then clears the Application and System logs.

.EXAMPLE
  PS> ./Clear-EventLogs.ps1 -LogName 'Application','System','Security' -Force
  Clears the three standard logs without prompting.

.EXAMPLE
  PS> ./Clear-EventLogs.ps1 -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Event log(s) to clear.'
  )]
  [string[]]
  $LogName = @('Application', 'System'),

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Skip the explicit confirmation prompt.'
  )]
  [switch]
  $Force,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Preview which logs would be cleared without clearing anything.'
  )]
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
  Write-Log -Message "DRY RUN - no event logs will be cleared`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

Write-Log -Message 'Checking event logs...' -Color Yellow

$_targets = @()
foreach ($_name in $LogName) {
  $_log = Get-WinEvent -ListLog $_name -ErrorAction SilentlyContinue
  if (-not $_log) {
    Add-OperationResult -Results $_results -Target $_name -Source 'EventLog' -Action 'Clear' -Status 'Skipped' -Detail 'Log not found.'
    continue
  }
  $_targets += [pscustomobject]@{
    Name = $_name
    RecordCount = [long]$_log.RecordCount
  }
}

if ($_targets.Count -eq 0) {
  Write-Log -Message 'No matching event logs found; nothing to clear.' -Color Yellow
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Clear-EventLogs'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

foreach ($_target in $_targets) {
  Write-Log -Message "  $($_target.Name): $($_target.RecordCount) record(s)." -Color Gray
}

if (-not $DryRun -and -not $Force) {
  $answer = Read-Host "This will permanently clear $($_targets.Count) event log(s) - removing their audit/diagnostic history.`nType 'yes' to continue"
  if ($answer -ne 'yes') {
    Write-Log -Message 'Event log clearing declined; no logs were cleared.' -Color Yellow
    foreach ($_target in $_targets) {
      Add-OperationResult -Results $_results -Target $_target.Name -Source 'EventLog' -Action 'Clear' -Status 'Skipped' -Detail 'Declined by user.'
    }
    $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Clear-EventLogs'
    if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }
}

foreach ($_target in $_targets) {
  if ($PSCmdlet.ShouldProcess($_target.Name, 'Clear event log')) {
    try {
      (& wevtutil cl "$($_target.Name)") 2>&1 | Out-Null
      if ($LASTEXITCODE -ne 0) {
        throw "wevtutil exited with code $LASTEXITCODE."
      }
      Add-OperationResult -Results $_results -Target $_target.Name -Source 'EventLog' -Action 'Clear' -Status 'Cleared' -Detail "Cleared ($($_target.RecordCount) record(s) removed)."
    }
    catch {
      Add-OperationResult -Results $_results -Target $_target.Name -Source 'EventLog' -Action 'Clear' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  else {
    Add-OperationResult -Results $_results -Target $_target.Name -Source 'EventLog' -Action 'Clear' -Status 'Skipped' -Detail 'WhatIf'
  }
}

$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_skipped = @($_results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_completed = @($_results | Where-Object { $_.Status -notin @('Failed', 'Skipped') }).Count
$_color = if ($_failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Log -Message "`nEvent log clearing complete. Cleared: $_completed | Skipped: $_skipped | Failed: $_failed" -Color $_color
$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Clear-EventLogs'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failed -gt 0) {
  exit 1
}
exit 0
