#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Configures Windows Update policy defaults.
.DESCRIPTION
  Applies Windows Update, Delivery Optimization, restart, and driver-update
  policy values using winkit registry helpers. The default profile keeps
  automatic updates enabled while reducing peer sharing, unexpected restarts,
  and driver delivery through quality updates.
.PARAMETER Undo
  Restore defaults or remove values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default update settings JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> ./Configure-Updates.ps1
  Applies the default Windows Update policy profile.
.EXAMPLE
  PS> ./Configure-Updates.ps1 -DryRun
  Previews update policy values without writing them.
.EXAMPLE
  PS> ./Configure-Updates.ps1 -Undo
  Removes Windows Update values managed by this script.
.EXAMPLE
  PS> ./Configure-Updates.ps1 -ExportConfig -ExportPath '.\updates.json'
  Exports the default update policy template.
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
    HelpMessage = 'Restore defaults or remove values managed by this script.'
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
    HelpMessage = 'JSON file containing setting overrides.'
  )]
  [string]
  $Config,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export the default update settings JSON and exit.'
  )]
  [switch]
  $ExportConfig,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export current registry values as reusable JSON config and exit.'
  )]
  [switch]
  $ExportCurrentState,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'File path used with -ExportConfig.'
  )]
  [string]
  $ExportPath,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) { $WhatIfPreference = $true; Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow }

$results = New-Object System.Collections.ArrayList

$updateSettings = @(
  @{
    Path = 'Registry::HKEY_USERS\S-1-5-20\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Settings'
    Name = 'DownloadMode'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'DeliveryOptimization'
    Description = 'Disable sharing downloaded updates with other PCs for NETWORK SERVICE.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
    Name = 'IsContinuousInnovationOptedIn'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'InnovationOptIn'
    Description = 'Disable getting latest updates as soon as available.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
    Name = 'NoAutoRebootWithLoggedOnUsers'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'RestartBehavior'
    Description = 'Prevent automatic restarts while users are logged on.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    Name = 'ExcludeWUDriversInQualityUpdate'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'DriverUpdates'
    Description = 'Exclude drivers from Windows quality updates.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState'
    Name = 'ExcludeWUDrivers'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'DriverUpdates'
    Description = 'Exclude drivers policy state.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
    Name = 'NoAutoUpdate'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Schedule'
    Description = 'Keep automatic updates enabled while using configured schedule values.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
    Name = 'AUOptions'
    Preferred = 4
    Default = $null
    Type = 'DWord'
    Group = 'Schedule'
    Description = 'Auto download and schedule install.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
    Name = 'ScheduledInstallDay'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Schedule'
    Description = 'Install updates every day when scheduled.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
    Name = 'ScheduledInstallTime'
    Preferred = 3
    Default = $null
    Type = 'DWord'
    Group = 'Schedule'
    Description = 'Scheduled install hour, 24-hour local time.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $updateSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current update settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $updateSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default update settings exported to: $_exportPath" -Color Green
  }
  else { $updateSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try { $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop; if ($_overrides -isnot [array]) { $_overrides = @($_overrides) } }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $updateSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
foreach ($entry in $updateSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  $target = "$($entry.Path)\$($entry.Name)"
  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel update setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $target" -Color Gray
      Add-OperationResult -Results $results -Target $target -Source 'Registry' -Action 'RemoveValue' -Status 'Skipped' -Detail 'DryRun'
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
    $action = 'RemoveValue'
  }
  else {
    Write-Log -Message "$targetLabel update setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would set $target = '$targetValue' ($($entry.Type))" -Color Gray
      Add-OperationResult -Results $results -Target $target -Source 'Registry' -Action 'SetValue' -Status 'Skipped' -Detail 'DryRun'
      continue
    }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
    $action = 'SetValue'
  }
  if ($result) {
    Write-Log -Message "  -> $($result.Status)" -Color Gray
    Add-OperationResult -Results $results -Target $target -Source 'Registry' -Action $action -Status $result.Status -Detail $entry.Description
  }
  else {
    Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red
    Add-OperationResult -Results $results -Target $target -Source 'Registry' -Action $action -Status 'Failed' -Detail "Could not process '$($entry.Name)'."
  }
}

$_changed = @($results | Where-Object { $_.Status -in @('Created', 'Updated', 'Removed') }).Count
$_skipped = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_failed = @($results | Where-Object { $_.Status -eq 'Failed' }).Count

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($_changed -gt 0) { Write-Log -Message "`nWindows Update settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }

Write-Log -Message "Update settings complete. Changed: $_changed | Skipped: $_skipped | Failed: $_failed" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
$_operationLog = Write-OperationResultLog -Results $results -ScriptName 'Configure-Updates'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $results
}
