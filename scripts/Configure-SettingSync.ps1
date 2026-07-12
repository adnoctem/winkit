#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Configures Windows setting sync policies.

.DESCRIPTION
  Applies machine-wide policy registry values that control Windows
  setting synchronisation. Covers the overall sync toggle and
  individual sync categories (app data, browser data, themes,
  passwords, language preferences).

  Keep separate from privacy and notification scripts because sync
  behaviour is often site-policy dependent and benefits from being
  managed independently.

.PARAMETER Undo
  Restore Windows defaults for the values managed by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER Config
  JSON file containing setting overrides.

.PARAMETER ExportConfig
  Export the default setting sync policy JSON and exit.

.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.

.PARAMETER ExportPath
  File path used with -ExportConfig.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Configure-SettingSync.ps1

.EXAMPLE
  PS> ./Configure-SettingSync.ps1 -DryRun

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
  [string]
  $Config,

  [Parameter(Mandatory = $false)]
  [switch]
  $ExportConfig,

  [Parameter(Mandatory = $false)]
  [switch]
  $ExportCurrentState,

  [Parameter(Mandatory = $false)]
  [string]
  $ExportPath,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$syncSettings = @(
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable Windows setting sync. (2 = disabled)'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableSettingSyncUserOverride'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Description = 'Prevent users from re-enabling setting sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableAppSyncSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable app data sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableApplicationSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable application setting sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableCredentialsSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable credential/password sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableDesktopThemeSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable theme/desktop sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableStartLayoutSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable Start layout sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableWebBrowserSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable web browser data sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableLanguageSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable language preference sync.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\SettingSync'
    Name = 'DisableWindowsSettingSync'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Description = 'Disable Windows setting sync (broad).'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $syncSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current setting sync policy exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $syncSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default setting sync policy exported to: $_exportPath" -Color Green
  }
  else { $syncSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try { $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop; if ($_overrides -isnot [array]) { $_overrides = @($_overrides) } }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $syncSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $syncSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($null -eq $targetValue) {
    if ($Undo) { Write-Log -Message "$targetLabel setting sync policy: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow }
    else { Write-Log -Message "Skipping setting sync policy '$($entry.Name)' - Preferred is null. Enable through JSON config if desired." -Color Gray; continue }
    if ($DryRun) { Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray; continue }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel setting sync policy: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray; continue }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }
  if ($result) { Write-Log -Message "  -> $($result.Status)" -Color Gray; if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true } }
  else { Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red }
}

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nSetting sync policy has been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }

$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $syncSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-SettingSync'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
