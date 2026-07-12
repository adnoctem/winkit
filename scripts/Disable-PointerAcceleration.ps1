#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Disables or re-enables pointer acceleration ("Enhance pointer precision") in Windows.

.DESCRIPTION
  Configures the HKCU:\Control Panel\Mouse registry values that control pointer
  acceleration.  By default the script disables acceleration; use -Undo to
  restore the Windows defaults.  The -Instant switch applies the change
  immediately via a user32.dll API call - no logout required.  -DryRun
  previews every action without touching the system.

.PARAMETER Undo
  Re-enable pointer acceleration (restore Windows defaults).

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER Instant
  Apply the registry change instantly via SystemParametersInfo.
  NOT supported together with -Undo - restart the machine to revert instead.
.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Disable-PointerAcceleration.ps1
  Disables pointer acceleration via the registry.  A sign-out or restart may
  be required for the change to take effect.

.EXAMPLE
  PS> ./Disable-PointerAcceleration.ps1 -Instant
  Disables pointer acceleration and applies the change immediately - no restart
  needed.

.EXAMPLE
  PS> ./Disable-PointerAcceleration.ps1 -Undo
  Restores the default (enabled) pointer acceleration settings.

.EXAMPLE
  PS> ./Disable-PointerAcceleration.ps1 -DryRun
  Shows which registry values would be modified without making any changes.

.EXAMPLE
  PS> ./Disable-PointerAcceleration.ps1 -Undo -DryRun
  Previews the undo operation without touching the registry.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
# Parameters
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Re-enable pointer acceleration (restore Windows defaults).'
  )]
  [switch]
  $Undo,

  [Parameter(
    Position = 1,
    Mandatory = $false,
    HelpMessage = 'Preview changes without applying them.'
  )]
  [switch]
  $DryRun,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Apply the registry change instantly via SystemParametersInfo. NOT supported together with -Undo - restart the machine to revert instead.'
  )]
  [switch]
  $Instant,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

# Validation
if ($Undo -and $Instant) {
  Write-Log -Message '-Instant cannot be combined with -Undo. The instantaneous change is forwarded directly to the mouse driver and cannot be reverted programmatically - please restart your machine to undo pointer acceleration settings.' -Color Red
  exit 1
}

# When -DryRun is active, enable WhatIf for all downstream lib calls.
# Set-RegistryValue (and the helpers it calls) all declare SupportsShouldProcess,
# so setting $WhatIfPreference causes them to report intent without mutating.
if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

# Registry target and values
# MouseSpeed / MouseThreshold1 / MouseThreshold2 are REG_SZ values under
# HKCU:\Control Panel\Mouse that together control "Enhance pointer precision".
#   Disabled (acceleration off): all three set to "0"
#   Enabled  (Windows default):  "1", "6", "10"
$registryPath = 'HKCU:\Control Panel\Mouse'
$mouseValues = @(
  @{
    Name = 'MouseSpeed'
    Off = '0'
    On = '1'
  }
  @{
    Name = 'MouseThreshold1'
    Off = '0'
    On = '6'
  }
  @{
    Name = 'MouseThreshold2'
    Off = '0'
    On = '10'
  }
)

$targetLabel = if ($Undo) { 'Enabling' } else { 'Disabling' }
$anyChanges = $false
$results = New-Object System.Collections.ArrayList

# Apply registry values
foreach ($entry in $mouseValues) {
  $targetValue = if ($Undo) { $entry.On } else { $entry.Off }
  $target = "$registryPath\$($entry.Name)"

  Write-Log -Message "$targetLabel pointer acceleration: $($entry.Name) = '$targetValue'" -Color Yellow

  if ($DryRun) {
    Write-Log -Message "  -> Would set $target = '$targetValue' (String)" -Color Gray
    Add-OperationResult -Results $results -Target $target -Source 'Registry' -Action 'SetValue' -Status 'Skipped' -Detail 'DryRun'
    continue
  }

  $result = Set-RegistryValue -Path $registryPath -Name $entry.Name -Value $targetValue -Type String

  if ($result) {
    Write-Log -Message "  -> $($result.Status)" -Color Gray
    Add-OperationResult -Results $results -Target $target -Source 'Registry' -Action 'SetValue' -Status $result.Status -Detail 'Pointer acceleration registry value.'
    if ($result.Status -in @('Created', 'Updated')) {
      $anyChanges = $true
    }
  }
  else {
    Write-Log -Message "  -> FAILED - could not write '$($entry.Name)'" -Color Red
    Add-OperationResult -Results $results -Target $target -Source 'Registry' -Action 'SetValue' -Status 'Failed' -Detail 'Could not write pointer acceleration registry value.'
  }
}

# Instant apply (only for disabling, never for undo)
if ($Instant -and -not $DryRun -and -not $Undo) {
  Write-Log -Message "`nApplying changes instantly via SystemParametersInfo ..." -Color Yellow

  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class NativeMethods {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);
}
'@

  # SPI_SETMOUSE = 0x0004, SPIF_SENDCHANGE = 0x0002
  $applied = [NativeMethods]::SystemParametersInfo(0x0004, 0, [IntPtr]::Zero, 0x0002)
  if ($applied) {
    Write-Log -Message '  -> Done - pointer acceleration settings applied immediately.' -Color Green
    Add-OperationResult -Results $results -Target 'SystemParametersInfo/SPI_SETMOUSE' -Source 'User32' -Action 'ApplyInstant' -Status 'Completed' -Detail 'Pointer acceleration settings applied immediately.'
  }
  else {
    Write-Log -Message '  -> Warning: SystemParametersInfo returned false - the change may not have been applied instantly. Try signing out and back in.' -Color Yellow
    Add-OperationResult -Results $results -Target 'SystemParametersInfo/SPI_SETMOUSE' -Source 'User32' -Action 'ApplyInstant' -Status 'Failed' -Detail 'SystemParametersInfo returned false.'
  }
}
elseif ($Instant -and $DryRun) {
  Write-Log -Message "`n[DRY RUN] Would call SystemParametersInfo to apply changes instantly." -Color Yellow
  Add-OperationResult -Results $results -Target 'SystemParametersInfo/SPI_SETMOUSE' -Source 'User32' -Action 'ApplyInstant' -Status 'Skipped' -Detail 'DryRun'
}

# Summary
if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow
}
elseif ($anyChanges) {
  if ($Undo) {
    Write-Log -Message "`nPointer acceleration has been re-enabled." -Color Green
  }
  else {
    Write-Log -Message "`nPointer acceleration has been disabled." -Color Green
    if (-not $Instant) {
      Write-Log -Message 'A sign-out or restart may be required for the change to take effect.' -Color Yellow
    }
  }
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}

$_operationLog = Write-OperationResultLog -Results $results -ScriptName 'Disable-PointerAcceleration'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $results
}
