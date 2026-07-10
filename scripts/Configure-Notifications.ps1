Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Configures Windows notification and suggestion behaviour.

.DESCRIPTION
  Applies registry values that suppress toast notifications, backup
  and account reminders, tile notifications, and suggested content
  delivery. Covers both machine-wide policy and per-user preference
  values.

.PARAMETER Undo
  Restore Windows defaults for the values managed by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed notification
  settings there instead of HKCU.

.PARAMETER Config
  JSON file containing setting overrides.

.PARAMETER ExportConfig
  Export the default notification settings JSON and exit.

.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.

.PARAMETER ExportPath
  File path used with -ExportConfig.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Configure-Notifications.ps1

.EXAMPLE
  PS> ./Configure-Notifications.ps1 -DryRun

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

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }

$notificationSettings = @(
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\PushNotifications"
    Name = 'ToastEnabled'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Toast notifications: off (0), on (1).'
  }
  @{
    Path = "HKLM:\Software\Policies\Microsoft\Windows\CloudContent"
    Name = 'DisableSoftLanding'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Description = 'Disable Windows welcome experience prompts.'
  }
  @{
    Path = "HKLM:\Software\Policies\Microsoft\Windows\CloudContent"
    Name = 'DisableWindowsConsumerFeatures'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Description = 'Disable consumer feature suggestions.'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    Name = 'SubscribedContent-338389Enabled'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Disable suggested content in Settings.'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    Name = 'SubscribedContent-338393Enabled'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Disable tips and suggestions on the lock screen.'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    Name = 'SubscribedContent-353696Enabled'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Disable lock screen facts and tips.'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings"
    Name = 'NOC_GLOBAL_SETTING_ALLOW_CRITICAL_TOASTS_ABOVE_LOCK'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Disable critical notifications above the lock screen.'
  }
  @{
    Path = "$regHive\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.BackgroundAccess"
    Name = 'Enabled'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Disable background access toast prompts.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $notificationSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current notification settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportConfig) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $notificationSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default notification settings exported to: $_exportPath" -Color Green
  }
  else { $notificationSettings | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) { Write-Log -Message '-Config requires a path to a JSON file.' -Color Red; exit 1 }
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) { Write-Log -Message "Config file not found: '$_configPath'" -Color Red; exit 1 }
  try { $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop; if ($_overrides -isnot [array]) { $_overrides = @($_overrides) } }
  catch { Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red; exit 1 }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $notificationSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; $mountResult = Mount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup
  if (-not $mountResult) { Write-Log -Message 'Failed to mount the default user hive.' -Color Red; exit 1 }
}

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $notificationSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($null -eq $targetValue) {
    if ($Undo) { Write-Log -Message "$targetLabel notification setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow }
    else { Write-Log -Message "Skipping notification setting '$($entry.Name)' - Preferred is null. Enable through JSON config if desired." -Color Gray; continue }
    if ($DryRun) { Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray; continue }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel notification setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) { Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray; continue }
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }
  if ($result) { Write-Log -Message "  -> $($result.Status)" -Color Gray; if ($result.Status -in @('Created', 'Updated', 'Removed')) { $anyChanges = $true } }
  else { Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red }
}

if ($SysPrep) { $_whatIfBackup = $WhatIfPreference; $WhatIfPreference = $false; Dismount-DefaultUserHive; $WhatIfPreference = $_whatIfBackup }

if ($DryRun) { Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow }
elseif ($anyChanges) { Write-Log -Message "`nNotification settings have been processed." -Color Green }
else { Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green }

$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $notificationSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-Notifications'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
