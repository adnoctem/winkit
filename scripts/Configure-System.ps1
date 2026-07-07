#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Configures Windows system, power, storage, accessibility, and shell defaults.
.DESCRIPTION
  Applies a compact system configuration profile using winkit registry helpers.
  This script groups broad system toggles that are too small to justify many
  single-purpose scripts, including power behavior, automatic device encryption,
  accessibility shortcuts, Storage Sense, shell behavior, and pointer defaults.
  Build-gated settings are skipped automatically on unsupported Windows builds.
.PARAMETER Undo
  Restore defaults or remove values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed system settings
  there instead of HKCU. HKLM values still target the live machine hive.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default system settings JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.EXAMPLE
  PS> ./Configure-System.ps1
  Applies the default system profile.
.EXAMPLE
  PS> ./Configure-System.ps1 -DryRun
  Shows system registry changes without applying them.
.EXAMPLE
  PS> ./Configure-System.ps1 -Undo
  Restores defaults or removes values managed by this script.
.EXAMPLE
  PS> ./Configure-System.ps1 -ExportConfig -ExportPath '.\system.json'
  Exports the default system settings template.
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
    HelpMessage = 'Mount the default user profile hive and write HKCU-backed system settings there.'
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
    HelpMessage = 'Export the default system settings JSON and exit.'
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
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) { $WhatIfPreference = $true; Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow }

$currentBuild = Get-OSBuildNumber
$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }

$systemSettings = @(
  @{
    Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power'
    Name = 'HiberbootEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Power'
    Description = 'Disable Fast Startup.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Power\PowerSettings\f15576e8-98b7-4186-b944-eafa664402d9'
    Name = 'ACSettingIndex'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Power'
    Description = 'Disable Modern Standby networking on AC power.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Power\PowerSettings\f15576e8-98b7-4186-b944-eafa664402d9'
    Name = 'DCSettingIndex'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Power'
    Description = 'Disable Modern Standby networking on battery.'
  }
  @{
    Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerThrottling'
    Name = 'PowerThrottlingOff'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 14393
    Group = 'Power'
    Description = 'Disable CPU PowerThrottling.'
  }
  @{
    Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker'
    Name = 'PreventDeviceEncryption'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Security'
    Description = 'Disable automatic BitLocker device encryption.'
  }
  @{
    Path = "$regHive\Control Panel\Accessibility\StickyKeys"
    Name = 'Flags'
    Preferred = '506'
    Default = $null
    Type = 'String'
    MinBuild = 26100
    Group = 'Security'
    Description = 'Disable Sticky Keys keyboard shortcut.'
  }
  @{
    Path = "$regHive\SOFTWARE\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy"
    Name = '01'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Storage'
    Description = 'Disable Storage Sense automatic cleanup.'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\CDP"
    Name = 'DragTrayEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 26200
    Group = 'Shell'
    Description = 'Disable Drag Tray sharing UI.'
  }
  @{
    Path = "$regHive\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
    Name = ''
    Preferred = ''
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Group = 'Shell'
    Description = 'Use classic Windows 10 context menu style.'
  }
  @{
    Path = "$regHive\Control Panel\Mouse"
    Name = 'MouseSpeed'
    Preferred = '0'
    Default = $null
    Type = 'String'
    Group = 'Pointer'
    Description = 'Disable pointer acceleration: speed.'
  }
  @{
    Path = "$regHive\Control Panel\Mouse"
    Name = 'MouseThreshold1'
    Preferred = '0'
    Default = $null
    Type = 'String'
    Group = 'Pointer'
    Description = 'Disable pointer acceleration: threshold 1.'
  }
  @{
    Path = "$regHive\Control Panel\Mouse"
    Name = 'MouseThreshold2'
    Preferred = '0'
    Default = $null
    Type = 'String'
    Group = 'Pointer'
    Description = 'Disable pointer acceleration: threshold 2.'
  }
  @{
    Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching'
    Name = 'SearchOrderConfig'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Drivers'
    Description = 'Disable automatic driver installation from Windows Update.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $systemSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current system settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $systemSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default system settings exported to: $_exportPath" -Color Green
  }
  else { $systemSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try {
    $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop
    if ($_overrides -isnot [array]) { $_overrides = @($_overrides) }
  }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $systemSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$_eligibleSettings = @($systemSettings | Where-Object { (-not $_.ContainsKey('MinBuild') -or $currentBuild -ge $_.MinBuild) -and (-not $_.ContainsKey('MaxBuild') -or $currentBuild -le $_.MaxBuild) })
foreach ($_entry in @($systemSettings | Where-Object { ($_.ContainsKey('MinBuild') -and $currentBuild -lt $_.MinBuild) -or ($_.ContainsKey('MaxBuild') -and $currentBuild -gt $_.MaxBuild) })) {
  Write-Log -Message "Skipping system setting '$($_entry.Name)' on build $currentBuild - $($_entry.Description)" -Color Yellow
}

if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; $mountResult = Mount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup
  if (-not $mountResult) { Write-Log -Message 'Failed to mount the default user hive.' -Color Red; exit 1 }
}

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false
foreach ($entry in $_eligibleSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  $displayName = if ($entry.Name -eq '') { '(default)' } else { $entry.Name }
  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel system setting: Remove '$displayName' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would remove $($entry.Path)\$displayName" -Color Gray; continue }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel system setting: $displayName = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would set $($entry.Path)\$displayName = '$targetValue' ($($entry.Type))" -Color Gray; continue }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }
  if ($result) { Write-Log -Message "  -> $($result.Status)" -Color Gray; if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true } }
  else { Write-Log -Message "  -> FAILED - could not process '$displayName'" -Color Red }
}

if ($SysPrep) { $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; Dismount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup }

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nSystem settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $systemSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-System'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
