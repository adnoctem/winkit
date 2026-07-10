Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Disables Windows diagnostic tracking transport and telemetry collection policy.
.DESCRIPTION
  Applies service and policy registry values for the Diagnostic Tracking
  service (DiagTrack), dmwappushsvc, and Windows DataCollection policy. This
  intentionally avoids IFEO debugger/taskkill tricks and does not delete tasks,
  hosts entries, or telemetry files.
.PARAMETER Undo
  Remove values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default diagnostic tracking settings JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.EXAMPLE
  PS> ./Disable-DiagnosticTracking.ps1
  Applies diagnostic tracking service and telemetry policy values.
.EXAMPLE
  PS> ./Disable-DiagnosticTracking.ps1 -DryRun
  Previews diagnostic tracking changes without writing registry values.
.EXAMPLE
  PS> ./Disable-DiagnosticTracking.ps1 -Undo
  Removes diagnostic tracking values managed by this script.
.EXAMPLE
  PS> ./Disable-DiagnosticTracking.ps1 -ExportConfig -ExportPath '.\diagnostic-tracking.json'
  Exports the default diagnostic tracking profile.
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
    HelpMessage = 'Remove values managed by this script.'
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
    HelpMessage = 'Export the default diagnostic tracking settings JSON and exit.'
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

# -----------------------------------------------------------------------------

if ($DryRun) { $WhatIfPreference = $true; Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow }

$diagnosticSettings = @(
  @{
    Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\DiagTrack'
    Name = 'Start'
    Preferred = 4
    Default = $null
    Type = 'DWord'
    Group = 'Services'
    Description = 'Disable Diagnostic Tracking service (DiagTrack).'
  }
  @{
    Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\dmwappushsvc'
    Name = 'Start'
    Preferred = 4
    Default = $null
    Type = 'DWord'
    Group = 'Services'
    Description = 'Disable dmwappushsvc telemetry service.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection'
    Name = 'AllowTelemetry'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'DataCollection'
    Description = 'Set telemetry collection to Security/Off where supported.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'
    Name = 'AllowTelemetry'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'DataCollection'
    Description = 'Set policy telemetry collection to Security/Off where supported.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'
    Name = 'MaxTelemetryAllowed'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'DataCollection'
    Description = 'Set maximum telemetry allowed to zero.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows'
    Name = 'CEIPEnable'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'CustomerExperience'
    Description = 'Disable Windows Customer Experience Improvement Program.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Microsoft\SQMClient\Windows'
    Name = 'CEIPEnable'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'CustomerExperience'
    Description = 'Disable local CEIP preference.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $diagnosticSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current diagnostic tracking settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $diagnosticSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default diagnostic tracking settings exported to: $_exportPath" -Color Green
  }
  else { $diagnosticSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try { $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop; if ($_overrides -isnot [array]) { $_overrides = @($_overrides) } }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $diagnosticSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$targetLabel = if ($Undo) { 'Removing' } else { 'Applying' }
$anyChanges = $false
foreach ($entry in $diagnosticSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel diagnostic tracking setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray; continue }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel diagnostic tracking setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray; continue }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }
  if ($result) { Write-Log -Message "  -> $($result.Status)" -Color Gray; if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true } }
  else { Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red }
}

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nDiagnostic tracking settings have been processed. Restart Windows for service start changes to take effect." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $diagnosticSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Disable-DiagnosticTracking'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
