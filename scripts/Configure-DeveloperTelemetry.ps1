#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Configures developer and workstation telemetry policy defaults.
.DESCRIPTION
  Applies telemetry and feedback policy defaults for Office 16.0 and Visual
  Studio, with optional vendor/workstation preference entries exported but
  disabled by default. Optional entries use Preferred = null so the default run
  only manages conservative policy values unless a JSON config opts into more.
.PARAMETER Undo
  Remove values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed telemetry settings
  there instead of HKCU. HKLM policy values still target the live machine hive.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default developer telemetry settings JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.EXAMPLE
  PS> ./Configure-DeveloperTelemetry.ps1
  Applies the default Office and Visual Studio telemetry policy profile.
.EXAMPLE
  PS> ./Configure-DeveloperTelemetry.ps1 -DryRun
  Previews developer telemetry policy changes without writing registry values.
.EXAMPLE
  PS> ./Configure-DeveloperTelemetry.ps1 -SysPrep
  Writes HKCU-backed Office/Visual Studio defaults to the default user profile hive.
.EXAMPLE
  PS> ./Configure-DeveloperTelemetry.ps1 -ExportConfig -ExportPath '.\developer-telemetry.json'
  Exports the full profile, including optional NVIDIA and preference entries.
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
    HelpMessage = 'Mount the default user profile hive and write HKCU-backed telemetry settings there.'
  )]
  [switch]
  $SysPrep,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'JSON file containing setting overrides.'
  )]
  [string]
  $Config,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export the default developer telemetry settings JSON and exit.'
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

$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }

$developerSettings = @(
  @{
    Path = "$regHive\SOFTWARE\Policies\Microsoft\Office\16.0\OSM"
    Name = 'EnableLogging'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'OfficePolicy'
    Description = 'Disable Office 16.0 OSM logging policy.'
  }
  @{
    Path = "$regHive\SOFTWARE\Policies\Microsoft\Office\16.0\OSM"
    Name = 'EnableUpload'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'OfficePolicy'
    Description = 'Disable Office 16.0 OSM upload policy.'
  }
  @{
    Path = "$regHive\SOFTWARE\Microsoft\Office\16.0\Common\ClientTelemetry"
    Name = 'DisableTelemetry'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'OfficePreferenceOptional'
    Description = 'Optional: disable Office client telemetry preference.'
  }
  @{
    Path = "$regHive\SOFTWARE\Microsoft\Office\16.0\Common\ClientTelemetry"
    Name = 'VerboseLogging'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'OfficePreferenceOptional'
    Description = 'Optional: disable Office verbose telemetry logging.'
  }
  @{
    Path = "$regHive\SOFTWARE\Microsoft\Office\16.0\Common\Feedback"
    Name = 'Enabled'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'OfficePreferenceOptional'
    Description = 'Optional: disable Office feedback.'
  }
  @{
    Path = "$regHive\SOFTWARE\Microsoft\Office\16.0\Common"
    Name = 'QMEnable'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'OfficePreferenceOptional'
    Description = 'Optional: disable Office quality monitoring.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Feedback'
    Name = 'DisableFeedbackDialog'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'VisualStudioPolicy'
    Description = 'Disable Visual Studio feedback dialog.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Feedback'
    Name = 'DisableEmailInput'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'VisualStudioPolicy'
    Description = 'Disable Visual Studio feedback email input.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Feedback'
    Name = 'DisableScreenshotCapture'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'VisualStudioPolicy'
    Description = 'Disable Visual Studio feedback screenshot capture.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\SQM'
    Name = 'OptIn'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'VisualStudioPolicy'
    Description = 'Disable Visual Studio SQM opt-in policy.'
  }
  @{
    Path = "$regHive\Software\Microsoft\VisualStudio\Telemetry"
    Name = 'TurnOffSwitch'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'VisualStudioPreferenceOptional'
    Description = 'Optional: disable current-user Visual Studio telemetry switch.'
  }
  @{
    Path = 'HKLM:\Software\Microsoft\VSCommon\16.0\SQM'
    Name = 'OptIn'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'VisualStudioPreferenceOptional'
    Description = 'Optional: disable VSCommon 16.0 SQM opt-in.'
  }
  @{
    Path = 'HKLM:\Software\Wow6432Node\Microsoft\VSCommon\16.0\SQM'
    Name = 'OptIn'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'VisualStudioPreferenceOptional'
    Description = 'Optional: disable Wow6432Node VSCommon 16.0 SQM opt-in.'
  }
  @{
    Path = 'HKLM:\Software\NVIDIA Corporation\Global\FTS'
    Name = 'EnableRID44231'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'NvidiaOptional'
    Description = 'Optional: disable NVIDIA telemetry RID44231.'
  }
  @{
    Path = 'HKLM:\Software\NVIDIA Corporation\Global\FTS'
    Name = 'EnableRID64640'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'NvidiaOptional'
    Description = 'Optional: disable NVIDIA telemetry RID64640.'
  }
  @{
    Path = 'HKLM:\Software\NVIDIA Corporation\Global\FTS'
    Name = 'EnableRID66610'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'NvidiaOptional'
    Description = 'Optional: disable NVIDIA telemetry RID66610.'
  }
  @{
    Path = 'HKLM:\Software\NVIDIA Corporation\NvControlPanel2\Client'
    Name = 'OptInOrOutPreference'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'NvidiaOptional'
    Description = 'Optional: opt out of NVIDIA Control Panel telemetry.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $developerSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current developer telemetry settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $developerSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default developer telemetry settings exported to: $_exportPath" -Color Green
  }
  else { $developerSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try { $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop; if ($_overrides -isnot [array]) { $_overrides = @($_overrides) } }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $developerSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; $mountResult = Mount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup
  if (-not $mountResult) { Write-Log -Message 'Failed to mount the default user hive.' -Color Red; exit 1 }
}

$targetLabel = if ($Undo) { 'Removing' } else { 'Applying' }
$anyChanges = $false
foreach ($entry in $developerSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($null -eq $targetValue) {
    if ($Undo) { Write-Log -Message "$targetLabel developer telemetry setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow }
    else { Write-Log -Message "Skipping developer telemetry setting '$($entry.Name)' - Preferred is null. Enable through JSON config if desired." -Color Gray; continue }
    if ($DryRun) { Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray; continue }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel developer telemetry setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray; continue }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }
  if ($result) { Write-Log -Message "  -> $($result.Status)" -Color Gray; if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true } }
  else { Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red }
}

if ($SysPrep) { $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; Dismount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup }

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nDeveloper telemetry settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $developerSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-DeveloperTelemetry'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
