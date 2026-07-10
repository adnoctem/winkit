Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Disables Windows Game DVR and Game Bar protocol integration.

.DESCRIPTION
  Applies registry values that disable Game DVR capture and redirect Game Bar
  protocol handlers away from the Xbox Game Bar application. Uses the winkit
  registry helpers for idempotent writes. -Undo removes the values managed by
  this script, matching the delete-on-undo behavior documented in the source
  material.

.PARAMETER Undo
  Remove the Game DVR and Game Bar registry values managed by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER Config
  Path to a JSON file containing an array of setting overrides. Each entry in
  the JSON needs at minimum a "Name" field matching a known registry value;
  "Preferred" and/or "Default" fields replace the corresponding built-in
  values. For default registry values, include the full "Path" and use an empty
  string for "Name".

.PARAMETER ExportConfig
  Export the default Game DVR settings as JSON to the console. Use -ExportPath
  to write to a file instead. Cannot be combined with -DryRun.

.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  When used together with -ExportConfig, writes the JSON to this file path
  instead of printing to the console.

.EXAMPLE
  PS> ./Disable-GameDVR.ps1
  Disables Game DVR and Game Bar protocol integration.

.EXAMPLE
  PS> ./Disable-GameDVR.ps1 -Undo
  Removes the registry values managed by this script.

.EXAMPLE
  PS> ./Disable-GameDVR.ps1 -DryRun
  Shows which registry values would be modified without making any changes.

.EXAMPLE
  PS> ./Disable-GameDVR.ps1 -ExportConfig -ExportPath '.\game-dvr-settings.json'
  Exports the default Game DVR settings template to .\game-dvr-settings.json.

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
    HelpMessage = 'Remove the Game DVR and Game Bar registry values managed by this script.'
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
    HelpMessage = 'Path to a JSON config file that overrides individual Game DVR settings.'
  )]
  [string]
  $Config,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export the default Game DVR settings to the console or to a file with -ExportPath.'
  )]
  [switch]
  $ExportConfig,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export current Game DVR registry values to reusable JSON config.'
  )]
  [switch]
  $ExportCurrentState,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'File path for -ExportConfig. When omitted the settings are printed to the console.'
  )]
  [string]
  $ExportPath,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$gameConfigStoreKey = 'HKCU:\System\GameConfigStore'
$gameDvrUserKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR'
$gameDvrPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\GameDVR'

# Write protocol handlers to HKLM:\Software\Classes instead of HKCR's merged
# view so the target is explicit and machine-wide.
$gameBarClassesKey = 'HKLM:\SOFTWARE\Classes\ms-gamebar'
$gameBarServicesClassesKey = 'HKLM:\SOFTWARE\Classes\ms-gamebarservices'

$gameDvrSettings = @(
  @{
    Path = $gameConfigStoreKey
    Name = 'GameDVR_Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'DisableDVR: Disable Game DVR for the current user'
  }
  @{
    Path = $gameDvrUserKey
    Name = 'AppCaptureEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'DisableDVR: Disable app capture for the current user'
  }
  @{
    Path = $gameDvrPolicyKey
    Name = 'AllowGameDVR'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'DisableDVR: Disable Game DVR machine-wide'
  }
  @{
    Path = 'HKCU:\SOFTWARE\Microsoft\GameBar'
    Name = 'UseNexusForGameBarEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'DisableGameBarIntegration: Disable Nexus/Game Bar integration'
  }

  @{
    Path = $gameBarClassesKey
    Name = ''
    Preferred = 'URL:ms-gamebar'
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Register ms-gamebar protocol label'
  }
  @{
    Path = $gameBarClassesKey
    Name = 'URL Protocol'
    Preferred = ''
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Preserve ms-gamebar URL protocol marker'
  }
  @{
    Path = $gameBarClassesKey
    Name = 'NoOpenWith'
    Preferred = ''
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Hide ms-gamebar protocol from Open With'
  }
  @{
    Path = "$gameBarClassesKey\shell\open\command"
    Name = ''
    Preferred = '%SystemRoot%\System32\systray.exe'
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Redirect ms-gamebar launch command'
  }

  @{
    Path = $gameBarServicesClassesKey
    Name = ''
    Preferred = 'URL:ms-gamebarservices'
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Register ms-gamebarservices protocol label'
  }
  @{
    Path = $gameBarServicesClassesKey
    Name = 'URL Protocol'
    Preferred = ''
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Preserve ms-gamebarservices URL protocol marker'
  }
  @{
    Path = $gameBarServicesClassesKey
    Name = 'NoOpenWith'
    Preferred = ''
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Hide ms-gamebarservices protocol from Open With'
  }
  @{
    Path = "$gameBarServicesClassesKey\shell\open\command"
    Name = ''
    Preferred = '%SystemRoot%\System32\systray.exe'
    Default = $null
    Type = 'String'
    Description = 'DisableGameBarIntegration: Redirect ms-gamebarservices launch command'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $gameDvrSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current Game DVR settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) {
    Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red
    exit 1
  }

  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $gameDvrSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default Game DVR settings exported to: $_exportPath" -Color Green
  }
  else {
    $gameDvrSettings | ConvertTo-Json -Depth 3
  }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) {
    Write-Log -Message '-Config requires a path to a JSON file.' -Color Red
    exit 1
  }

  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)

  if (-not (Test-Path -LiteralPath $_configPath)) {
    Write-Log -Message "Config file not found: '$_configPath'" -Color Red
    exit 1
  }

  try {
    $_jsonContent = Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop
    $_overrides = ConvertFrom-Json -InputObject $_jsonContent -ErrorAction Stop

    if ($_overrides -isnot [array]) {
      $_overrides = @($_overrides)
    }
  }
  catch {
    Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red
    exit 1
  }

  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $gameDvrSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$targetLabel = if ($Undo) { 'Removing' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $gameDvrSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  $displayName = if ($entry.Name -eq '') { '(default)' } else { $entry.Name }

  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel Game DVR setting: Remove '$displayName' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$displayName" -Color Gray
      continue
    }

    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel Game DVR setting: $displayName = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would set $($entry.Path)\$displayName = '$targetValue' ($($entry.Type))" -Color Gray
      continue
    }

    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }

  if ($result) {
    Write-Log -Message "  -> $($result.Status)" -Color Gray
    if ($result.Status -in @('Created', 'Updated', 'Removed')) {
      $anyChanges = $true
    }
  }
  else {
    Write-Log -Message "  -> FAILED - could not process '$displayName'" -Color Red
  }
}

if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow
}
elseif ($anyChanges) {
  if ($Undo) {
    Write-Log -Message "`nGame DVR and Game Bar settings managed by this script have been removed." -Color Green
  }
  else {
    Write-Log -Message "`nGame DVR and Game Bar integration have been disabled." -Color Green
  }

  Write-Log -Message 'Sign out and back in, or restart Windows, for all changes to take effect.' -Color Yellow
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $gameDvrSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Disable-GameDVR'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
