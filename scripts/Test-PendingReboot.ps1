#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.1.0' }

<#
.SYNOPSIS
  Checks the local machine for pending-reboot indicators across multiple
  independent sources.

.DESCRIPTION
  Wrapper around PSFoundation's Test-PendingReboot, which checks ten
  independent read-only pending-reboot indicators:

    - CBS RebootPending key and servicing state (RebootInProgress /
      PackagesPending) - set by DISM or feature installs independent of
      Windows Update
    - Session Manager PendingFileRenameOperations and ...Operations2
    - RunOnce DVDRebootSignal
    - Netlogon JoinDomain / AvoidSpnSet (pending domain-join reboot)
    - Windows Update Auto Update RebootRequired / PostRebootReporting
    - SCCM/MECM client SDK (CCM_ClientUtilities.DetermineIfRebootPending)
    - Pending computer rename (ActiveComputerName vs ComputerName)
    - Win32_ComputerSystem.PendingSystemReboot

  The check is independent of PSWindowsUpdate's Get-WURebootStatus, which
  only reflects the Windows Update service state.

  Emits a structured result entry and exits with code 1 when any pending
  reboot indicator is set, making the script usable as a gate before
  operations that require a clean reboot state. No elevation needed.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Test-PendingReboot.ps1
  Reports pending-reboot indicators and exits 1 if any is set.

.EXAMPLE
  PS> ./Test-PendingReboot.ps1 -PassThru
  Returns the result entry.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

$_results = New-Object System.Collections.ArrayList

$pendingState = Test-PendingReboot

if ($pendingState.PendingReboot) {
  Write-Log -Message "Pending reboot detected - indicators: $($pendingState.Indicators -join ', ')" -Color Red
  Add-OperationResult -Results $_results -Target 'PendingReboot' -Source 'PendingReboot' -Action 'Check' -Status 'Detected' -Detail ($pendingState.Indicators -join ', ')
}
else {
  Write-Log -Message 'No pending reboot detected.' -Color Green
  Add-OperationResult -Results $_results -Target 'PendingReboot' -Source 'PendingReboot' -Action 'Check' -Status 'NotDetected' -Detail 'All indicators clear.'
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Test-PendingReboot'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_results
}

if ($pendingState.PendingReboot) {
  exit 1
}
exit 0
