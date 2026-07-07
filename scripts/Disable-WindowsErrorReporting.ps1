#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Disables Windows Error Reporting and related diagnostic feedback.

.DESCRIPTION
  Applies policy registry values that suppress Windows Error Reporting,
  Watson crash dumps, and related diagnostic feedback prompts. Use -Undo
  to re-enable. WER is useful during diagnostics - this script should be
  applied deliberately, not as a blanket default.

.PARAMETER Undo
  Re-enable Windows Error Reporting and remove the policy values managed
  by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Disable-WindowsErrorReporting.ps1

.EXAMPLE
  PS> ./Disable-WindowsErrorReporting.ps1 -Undo -DryRun

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

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$werSettings = @(
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\Windows Error Reporting'
    Name = 'Disabled'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Description = 'Disable Windows Error Reporting.'
  }
  @{
    Path = 'HKLM:\Software\Microsoft\Windows\Windows Error Reporting'
    Name = 'Disabled'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Description = 'Disable Windows Error Reporting user override.'
  }
  @{
    Path = 'HKLM:\Software\Microsoft\Windows\Windows Error Reporting'
    Name = 'DontShowUI'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Description = 'Suppress error-reporting UI prompts.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\PCHealth\ErrorReporting'
    Name = 'DoReport'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Disable PCHealth error reporting.'
  }
)

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $werSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }

  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel WER setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel WER setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
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
elseif ($anyChanges) { Write-Log -Message "`nWindows Error Reporting settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }

$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $werSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Disable-WindowsErrorReporting'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
