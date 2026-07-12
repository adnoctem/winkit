#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Configures Windows visual appearance: theme mode, transparency, and animations.

.DESCRIPTION
  Applies registry values that control the Windows visual appearance -
  dark/light app and system mode, transparency effects, and animation
  behaviour. These are per-user settings persisted under HKCU.

  Use -SysPrep to write to the default user profile hive so new user
  profiles inherit the configured appearance defaults.

.PARAMETER Undo
  Restore Windows defaults for the values managed by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed appearance
  settings there instead of HKCU.

.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.

.PARAMETER ExportConfig
  Export the default appearance settings JSON and exit.

.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.

.PARAMETER ExportPath
  File path used with -ExportConfig.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Configure-Appearance.ps1

.EXAMPLE
  PS> ./Configure-Appearance.ps1 -SysPrep -DryRun

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
  $SysPrep,

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

# Elevation guard: -SysPrep mounts the default user hive (requires admin)
if ($SysPrep -and -not (Test-Elevation)) {
  Write-Error '-SysPrep requires administrator privileges. Run elevated or omit -SysPrep for current-user only.'
  exit 1
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }

$appearanceSettings = @(
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    Name = 'AppsUseLightTheme'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'App mode: dark (0), light (1).'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    Name = 'SystemUsesLightTheme'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'System mode: dark (0), light (1).'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    Name = 'EnableTransparency'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Transparency effects: off (0), on (1).'
  }
  @{
    Path = "$regHive\Control Panel\Desktop\WindowMetrics"
    Name = 'MinAnimate'
    Preferred = 0
    Default = 1
    Type = 'String'
    Description = 'Window animation effects: off (0), on (1).'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Name = 'TaskbarAnimations'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Taskbar animations: off (0), on (1).'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $appearanceSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current appearance settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $appearanceSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default appearance settings exported to: $_exportPath" -Color Green
  }
  else { $appearanceSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try { $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop; if ($_overrides -isnot [array]) { $_overrides = @($_overrides) } }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $appearanceSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; $mountResult = Mount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup
  if (-not $mountResult) { Write-Log -Message 'Failed to mount the default user hive.' -Color Red; exit 1 }
}

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $appearanceSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($null -eq $targetValue) {
    if ($Undo) { Write-Log -Message "$targetLabel appearance setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow }
    else { Write-Log -Message "Skipping appearance setting '$($entry.Name)' - Preferred is null. Enable through JSON config if desired." -Color Gray; continue }
    if ($DryRun) { Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray; continue }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel appearance setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray; continue }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }
  if ($result) { Write-Log -Message "  -> $($result.Status)" -Color Gray; if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true } }
  else { Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red }
}

if ($SysPrep) { $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; Dismount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup }

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nAppearance settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }

$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $appearanceSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-Appearance'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
