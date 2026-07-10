Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Disables Windows Remote Assistance policy settings.

.DESCRIPTION
  Applies policy registry values that prevent unsolicited Remote Assistance
  requests and block "Get Help" / Quick Assist-style remote sessions. Separate
  from Remote Desktop behaviour - this only controls Remote Assistance.

.PARAMETER Undo
  Re-enable Remote Assistance and remove the policy values managed by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Disable-RemoteAssistance.ps1

.EXAMPLE
  PS> ./Disable-RemoteAssistance.ps1 -Undo -DryRun

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

$raSettings = @(
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services'
    Name = 'fAllowUnsolicited'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Disallow unsolicited Remote Assistance requests.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services'
    Name = 'fAllowToGetHelp'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Disallow Get Help / Quick Assist remote control.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services'
    Name = 'fAllowFullControl'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Disallow full-control Remote Assistance sessions.'
  }
)

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $raSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }

  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel Remote Assistance setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel Remote Assistance setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray
      continue
    }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }

  if ($result) {
    Write-Log -Message "  -> $($result.Status)" -Color Gray
    if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true }
  }
  else {
    Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red
  }
}

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nRemote Assistance settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }

$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $raSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Disable-RemoteAssistance'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
