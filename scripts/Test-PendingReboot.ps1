#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Checks the local machine for pending-reboot indicators across multiple
  independent sources.

.DESCRIPTION
  Read-only diagnostic that inspects the standard pending-reboot indicators a
  Windows machine can set, beyond the Windows Update service state alone:

    - CBS RebootPending key (Component Based Servicing)
    - CBS servicing state (RebootInProgress / PackagesPending values) - set
      by DISM or feature installs independent of Windows Update
    - Session Manager PendingFileRenameOperations and ...Operations2 - set by
      arbitrary installers performing a delayed file replace
    - RunOnce DVDRebootSignal
    - Netlogon JoinDomain / AvoidSpnSet - a pending domain-join reboot
    - Windows Update Auto Update RebootRequired / PostRebootReporting
    - SCCM/MECM client SDK (CCM_ClientUtilities.DetermineIfRebootPending)
    - Pending computer rename (ActiveComputerName vs ComputerName)
    - Win32_ComputerSystem.PendingSystemReboot

  Each indicator is reported individually (which indicator triggered, not just
  a boolean), which is useful for diagnostics. The script exits with code 1
  when any pending-reboot indicator is set, making it usable as a gate before
  operations that require a clean reboot state.

  The check is independent of PSWindowsUpdate's Get-WURebootStatus, which only
  reflects the Windows Update service state.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Test-PendingReboot.ps1
  Reports all pending-reboot indicators and exits 1 if any is set.

.EXAMPLE
  PS> ./Test-PendingReboot.ps1 -PassThru
  Returns the per-indicator result entries.

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

function Add-RebootIndicatorResult {
  [CmdletBinding()]
  param(
    [System.Collections.ArrayList]$Results,
    [string]$Name,
    [bool]$Detected,
    [string]$Detail
  )

  Add-OperationResult -Results $Results -Target $Name -Source 'PendingReboot' -Action 'Check' -Status $(if ($Detected) { 'Detected' } else { 'NotDetected' }) -Detail $Detail
}

$_pendingIndicators = New-Object System.Collections.Generic.List[string]

# 1. CBS RebootPending key.
$_cbsPendingKey = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
Add-RebootIndicatorResult -Results $_results -Name 'CBS RebootPending' -Detected $_cbsPendingKey -Detail 'Component Based Servicing RebootPending key.'
if ($_cbsPendingKey) { $_pendingIndicators.Add('CBS RebootPending') }

# 2. CBS servicing state (RebootInProgress / PackagesPending values).
$_cbs = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing' -ErrorAction SilentlyContinue
$_cbsServicingState = ($null -ne $_cbs) -and (($_cbs.PSObject.Properties.Name -contains 'RebootInProgress') -or ($_cbs.PSObject.Properties.Name -contains 'PackagesPending'))
Add-RebootIndicatorResult -Results $_results -Name 'CBS servicing state' -Detected $_cbsServicingState -Detail 'RebootInProgress/PackagesPending values under Component Based Servicing.'
if ($_cbsServicingState) { $_pendingIndicators.Add('CBS servicing state') }

# 3. Session Manager pending file rename operations (both variants).
$_sessionManagerPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
foreach ($_name in @('PendingFileRenameOperations', 'PendingFileRenameOperations2')) {
  $_detected = Test-Path -LiteralPath (Join-Path -Path $_sessionManagerPath -ChildPath $_name)
  Add-RebootIndicatorResult -Results $_results -Name $_name -Detected $_detected -Detail 'Delayed file replace pending under Session Manager.'
  if ($_detected) { $_pendingIndicators.Add($_name) }
}

# 4. RunOnce DVDRebootSignal.
$_runOnce = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -ErrorAction SilentlyContinue
$_dvdRebootSignal = ($null -ne $_runOnce) -and ($_runOnce.PSObject.Properties.Name -contains 'DVDRebootSignal')
Add-RebootIndicatorResult -Results $_results -Name 'DVDRebootSignal' -Detected $_dvdRebootSignal -Detail 'RunOnce DVDRebootSignal value.'
if ($_dvdRebootSignal) { $_pendingIndicators.Add('DVDRebootSignal') }

# 5. Netlogon domain-join signals.
$_netlogon = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -ErrorAction SilentlyContinue
$_netlogonJoin = ($null -ne $_netlogon) -and (($_netlogon.PSObject.Properties.Name -contains 'JoinDomain') -or ($_netlogon.PSObject.Properties.Name -contains 'AvoidSpnSet'))
Add-RebootIndicatorResult -Results $_results -Name 'Netlogon join' -Detected $_netlogonJoin -Detail 'JoinDomain/AvoidSpnSet values under Netlogon.'
if ($_netlogonJoin) { $_pendingIndicators.Add('Netlogon join') }

# 6. Windows Update Auto Update signals.
$_autoUpdate = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update' -ErrorAction SilentlyContinue
$_autoUpdatePending = ($null -ne $_autoUpdate) -and (($_autoUpdate.PSObject.Properties.Name -contains 'RebootRequired') -or ($_autoUpdate.PSObject.Properties.Name -contains 'PostRebootReporting'))
Add-RebootIndicatorResult -Results $_results -Name 'Windows Update' -Detected $_autoUpdatePending -Detail 'RebootRequired/PostRebootReporting values under Auto Update.'
if ($_autoUpdatePending) { $_pendingIndicators.Add('Windows Update') }

# 7. SCCM/MECM client SDK (no-op when the client is not installed).
$_sccmPending = $false
try {
  $_ccm = Get-CimInstance -Namespace 'root\ccm\clientsdk' -ClassName CCM_ClientUtilities -ErrorAction SilentlyContinue
  if ($_ccm) {
    $_sccmPending = [bool]($_ccm.DetermineIfRebootPending()).RebootPending
  }
}
catch {
  $_sccmPending = $false
}
Add-RebootIndicatorResult -Results $_results -Name 'SCCM/MECM client' -Detected $_sccmPending -Detail 'CCM_ClientUtilities.DetermineIfRebootPending.'
if ($_sccmPending) { $_pendingIndicators.Add('SCCM/MECM client') }

# 8. Pending computer rename (ActiveComputerName differs from ComputerName).
$_computerNameKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName'
$_activeName = (Get-ItemProperty "$_computerNameKey\ActiveComputerName" -ErrorAction SilentlyContinue).ComputerName
$_configuredName = (Get-ItemProperty "$_computerNameKey\ComputerName" -ErrorAction SilentlyContinue).ComputerName
$_renamePending = ($_activeName -and $_configuredName -and ($_activeName -ne $_configuredName))
Add-RebootIndicatorResult -Results $_results -Name 'Computer rename' -Detected $_renamePending -Detail "ActiveComputerName '$_activeName' vs ComputerName '$_configuredName'."
if ($_renamePending) { $_pendingIndicators.Add('Computer rename') }

# 9. CIM pending-reboot property.
$_computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
$_cimPending = ($null -ne $_computerSystem) -and [bool]$_computerSystem.PendingSystemReboot
Add-RebootIndicatorResult -Results $_results -Name 'CIM PendingSystemReboot' -Detected $_cimPending -Detail 'Win32_ComputerSystem.PendingSystemReboot.'
if ($_cimPending) { $_pendingIndicators.Add('CIM PendingSystemReboot') }

$_pending = ($_pendingIndicators.Count -gt 0)

if ($_pending) {
  Write-Log -Message "Pending reboot detected - indicators: $($_pendingIndicators -join ', ')" -Color Red
}
else {
  Write-Log -Message 'No pending reboot detected.' -Color Green
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Test-PendingReboot'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_results
}

if ($_pending) {
  exit 1
}
exit 0
