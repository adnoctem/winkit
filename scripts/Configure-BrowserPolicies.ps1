Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Configures browser policy defaults for Microsoft Edge and optional third-party browsers.
.DESCRIPTION
  Applies browser policy registry values using winkit helpers. The default
  profile configures Microsoft Edge Chromium only. Chrome and Firefox entries
  are exported with Preferred = null so users can opt in through JSON config
  without the default run unexpectedly managing third-party browser policies.
.PARAMETER Undo
  Remove values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default browser policy JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.EXAMPLE
  PS> ./Configure-BrowserPolicies.ps1
  Applies the default Microsoft Edge policy profile.
.EXAMPLE
  PS> ./Configure-BrowserPolicies.ps1 -DryRun
  Previews browser policy changes without writing registry values.
.EXAMPLE
  PS> ./Configure-BrowserPolicies.ps1 -Undo
  Removes browser policy values managed by this script.
.EXAMPLE
  PS> ./Configure-BrowserPolicies.ps1 -ExportConfig -ExportPath '.\browser-policies.json'
  Exports the default policy template, including opt-in Chrome and Firefox entries.
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
    HelpMessage = 'Export the default browser policy JSON and exit.'
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

$browserSettings = @(
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    Name = 'BlockThirdPartyCookies'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'Edge'
    Description = 'Block third-party cookies in Microsoft Edge.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    Name = 'AutofillCreditCardEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Edge'
    Description = 'Disable credit card autofill in Microsoft Edge.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    Name = 'SyncDisabled'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'Edge'
    Description = 'Disable Microsoft Edge sync.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    Name = 'BackgroundModeEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Edge'
    Description = 'Disable Microsoft Edge background mode.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    Name = 'PersonalizationReportingEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Edge'
    Description = 'Disable Microsoft Edge personalization reporting.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    Name = 'DiagnosticData'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Edge'
    Description = 'Disable Microsoft Edge diagnostic data.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    Name = 'ChromeCleanupEnabled'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'ChromeOptional'
    Description = 'Optional: disable Chrome cleanup.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    Name = 'ChromeCleanupReportingEnabled'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'ChromeOptional'
    Description = 'Optional: disable Chrome cleanup reporting.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Google\Chrome'
    Name = 'MetricsReportingEnabled'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'ChromeOptional'
    Description = 'Optional: disable Chrome metrics reporting.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Mozilla\Firefox'
    Name = 'DisableTelemetry'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'FirefoxOptional'
    Description = 'Optional: disable Firefox telemetry.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Mozilla\Firefox'
    Name = 'DisableDefaultBrowserAgent'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'FirefoxOptional'
    Description = 'Optional: disable Firefox default browser agent.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $browserSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current browser policy settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $browserSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default browser policy settings exported to: $_exportPath" -Color Green
  }
  else { $browserSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try { $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop; if ($_overrides -isnot [array]) { $_overrides = @($_overrides) } }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $browserSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$targetLabel = if ($Undo) { 'Removing' } else { 'Applying' }
$anyChanges = $false
foreach ($entry in $browserSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($null -eq $targetValue) {
    if ($Undo) { Write-Log -Message "$targetLabel browser policy: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow }
    else { Write-Log -Message "Skipping browser policy '$($entry.Name)' - Preferred is null. Enable through JSON config if desired." -Color Gray; continue }
    if ($DryRun) { Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray; continue }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel browser policy: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray; continue }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }
  if ($result) { Write-Log -Message "  -> $($result.Status)" -Color Gray; if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true } }
  else { Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red }
}

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nBrowser policy settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $browserSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-BrowserPolicies'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
