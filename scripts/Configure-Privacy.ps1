#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Configures Windows privacy, suggestions, activity history, clipboard, and input defaults.
.DESCRIPTION
  Applies a broad privacy-oriented registry profile using winkit helpers. This
  script intentionally excludes app permission defaults and diagnostic tracking
  service transport, which are handled by dedicated scripts. Automatic Windows
  Content Delivery and suggested app provisioning are handled separately by
  Disable-ContentDelivery.ps1.
.PARAMETER Undo
  Restore defaults or remove privacy values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed privacy settings
  there instead of HKCU. HKLM policy values are still written machine-wide.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default privacy settings JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.EXAMPLE
  PS> ./Configure-Privacy.ps1
  Applies the default privacy profile.
.EXAMPLE
  PS> ./Configure-Privacy.ps1 -DryRun
  Shows every registry value that would be written or removed.
.EXAMPLE
  PS> ./Configure-Privacy.ps1 -Undo
  Removes values whose Windows default is represented by absence.
.EXAMPLE
  PS> ./Configure-Privacy.ps1 -SysPrep
  Writes HKCU-backed privacy defaults to the default user profile hive for new users.
.EXAMPLE
  PS> ./Configure-Privacy.ps1 -ExportConfig -ExportPath '.\privacy.json'
  Exports the default privacy settings template to JSON.
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
    HelpMessage = 'Restore defaults or remove privacy values managed by this script.'
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
    HelpMessage = 'Mount the default user profile hive and write HKCU-backed privacy settings there.'
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
    HelpMessage = 'Export the default privacy settings JSON and exit.'
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

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }
$cloudContentKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent'
$advertisingKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
$privacyKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Privacy"
$speechKey = "$regHive\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy"
$inputTipcKey = "$regHive\Software\Microsoft\Input\TIPC"
$inputPersonalizationKey = "$regHive\Software\Microsoft\InputPersonalization"
$trainedDataKey = "$inputPersonalizationKey\TrainedDataStore"
$personalizationKey = "$regHive\Software\Microsoft\Personalization\Settings"
$explorerAdvancedKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$accountNotificationsKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\SystemSettings\AccountNotifications"
$userProfileEngagementKey = "$regHive\SOFTWARE\Microsoft\Windows\CurrentVersion\UserProfileEngagement"
$suggestedToastKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.Suggested"
$backupToastKey = "$regHive\SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.BackupReminder"
$mobilityKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Mobility"
$siufRulesKey = "$regHive\SOFTWARE\Microsoft\Siuf\Rules"
$activityPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'
$clipboardPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'
$locationPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors'
$appPrivacyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy'

$privacySettings = @(
  @{
    Path = $advertisingKey
    Name = 'Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Advertising'
    Description = 'Disable advertising ID for the current/default user.'
  }
  @{
    Path = $privacyKey
    Name = 'TailoredExperiencesWithDiagnosticDataEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Advertising'
    Description = 'Disable tailored experiences with diagnostic data.'
  }
  @{
    Path = $cloudContentKey
    Name = 'DisableConsumerAccountStateContent'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'Advertising'
    Description = 'Hide Microsoft account / 365 promotional state content.'
  }
  @{
    Path = $cloudContentKey
    Name = 'DisableThirdPartySuggestions'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'Advertising'
    Description = 'Disable third-party suggestions.'
  }

  @{
    Path = $accountNotificationsKey
    Name = 'EnableAccountNotifications'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable account notifications in Settings.'
  }
  @{
    Path = $userProfileEngagementKey
    Name = 'ScoobeSystemSettingEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable SCoobe / welcome engagement prompts.'
  }
  @{
    Path = $explorerAdvancedKey
    Name = 'ShowSyncProviderNotifications'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable sync provider notifications in Explorer.'
  }
  @{
    Path = $suggestedToastKey
    Name = 'Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable suggested system toast notifications.'
  }
  @{
    Path = $backupToastKey
    Name = 'Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable backup reminder toast notifications.'
  }
  @{
    Path = $mobilityKey
    Name = 'OptedIn'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable mobility suggestion opt-in.'
  }
  @{
    Path = $explorerAdvancedKey
    Name = 'Start_IrisRecommendations'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable Start Iris recommendations.'
  }
  @{
    Path = $explorerAdvancedKey
    Name = 'Start_AccountNotifications'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SuggestedContent'
    Description = 'Disable Start account notifications.'
  }

  @{
    Path = $siufRulesKey
    Name = 'NumberOfSIUFInPeriod'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Feedback'
    Description = 'Disable feedback prompt frequency.'
  }
  @{
    Path = $siufRulesKey
    Name = 'PeriodInNanoSeconds'
    Preferred = $null
    Default = $null
    Type = 'DWord'
    Group = 'Feedback'
    Description = 'Remove feedback prompt period.'
  }

  @{
    Path = $activityPolicyKey
    Name = 'PublishUserActivities'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ActivityHistory'
    Description = 'Disable publishing user activities.'
  }
  @{
    Path = $activityPolicyKey
    Name = 'UploadUserActivities'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ActivityHistory'
    Description = 'Disable uploading user activities.'
  }
  @{
    Path = $activityPolicyKey
    Name = 'EnableActivityFeed'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ActivityHistory'
    Description = 'Disable activity feed.'
  }
  @{
    Path = $clipboardPolicyKey
    Name = 'AllowClipboardHistory'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ClipboardAndCDP'
    Description = 'Disable clipboard history.'
  }
  @{
    Path = $clipboardPolicyKey
    Name = 'AllowCrossDeviceClipboard'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ClipboardAndCDP'
    Description = 'Disable cross-device clipboard.'
  }
  @{
    Path = $clipboardPolicyKey
    Name = 'CdpEnableMmx'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ClipboardAndCDP'
    Description = 'Disable CDP cross-device shared experiences.'
  }

  @{
    Path = $speechKey
    Name = 'HasAccepted'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'InputPersonalization'
    Description = 'Disable online speech privacy acceptance.'
  }
  @{
    Path = $inputTipcKey
    Name = 'Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'InputPersonalization'
    Description = 'Disable typing personalization collection.'
  }
  @{
    Path = $inputPersonalizationKey
    Name = 'RestrictImplicitInkCollection'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'InputPersonalization'
    Description = 'Restrict implicit ink collection.'
  }
  @{
    Path = $inputPersonalizationKey
    Name = 'RestrictImplicitTextCollection'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'InputPersonalization'
    Description = 'Restrict implicit text collection.'
  }
  @{
    Path = $trainedDataKey
    Name = 'HarvestContacts'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'InputPersonalization'
    Description = 'Disable contact harvesting for trained data.'
  }
  @{
    Path = $personalizationKey
    Name = 'AcceptedPrivacyPolicy'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'InputPersonalization'
    Description = 'Disable input personalization privacy acceptance.'
  }

  @{
    Path = $locationPolicyKey
    Name = 'DisableLocation'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'Location'
    Description = 'Disable Windows location services through policy.'
  }
  @{
    Path = $appPrivacyKey
    Name = 'LetAppsRunInBackground'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Group = 'AppBackground'
    Description = 'Deny default background app access.'
  }

  # -- Legacy privacy toggles --
  @{
    Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    Name = 'DisallowRunExecutionFromZipFile'
    Preferred = 1
    Default = 0
    Type = 'DWord'
    Group = 'LegacyPrivacy'
    Description = 'Disallow running executables directly from ZIP files.'
  }
  @{
    Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    Name = 'DisallowRunExecutionFromZipFile\Unblock'
    Preferred = 1
    Default = 0
    Type = 'DWord'
    Group = 'LegacyPrivacy'
    Description = 'Disallow unblocking ZIP-executed files via Mark-of-the-Web removal.'
  }
  @{
    Path = 'HKCU:\Software\Microsoft\MediaPlayer\Preferences'
    Name = 'UsageTracking'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Group = 'LegacyPrivacy'
    Description = 'Disable Windows Media Player usage tracking.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\WindowsMediaDRM'
    Name = 'DisableOnline'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'LegacyPrivacy'
    Description = 'Disable WMDRM online feature access.'
  }
  @{
    Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    Name = 'EnableBalloonTips'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Group = 'LegacyPrivacy'
    Description = 'Disable balloon tip notifications.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\System'
    Name = 'DisableCredUI'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'LegacyPrivacy'
    Description = 'Disable credential password reveal UI.'
  }
  @{
    Path = 'HKLM:\Software\Policies\Microsoft\Windows\FindMyDevice'
    Name = 'AllowFindMyDevice'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'LegacyPrivacy'
    Description = 'Disable Find My Device feature.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $privacySettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current privacy settings exported to: $_exportPath" -Color Green
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
    $privacySettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default privacy settings exported to: $_exportPath" -Color Green
  }
  else {
    $privacySettings | ConvertTo-Json -Depth 3
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
    $_overrides = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop
    if ($_overrides -isnot [array]) { $_overrides = @($_overrides) }
  }
  catch {
    Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red
    exit 1
  }
  Write-Log -Message "Merging config: $_configPath`n" -Color Yellow
  Merge-ObjectArrays -Base $privacySettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference
  $WhatIfPreference = $false
  $mountResult = Mount-DefaultUserHive
  $WhatIfPreference = $_whatIfBackup
  if (-not $mountResult) {
    Write-Log -Message 'Failed to mount the default user hive.' -Color Red
    exit 1
  }
}

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $privacySettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if (($Undo -and $null -eq $entry.Default) -or ((-not $Undo) -and $null -eq $entry.Preferred)) {
    Write-Log -Message "$targetLabel privacy setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel privacy setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
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

if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference
  $WhatIfPreference = $false
  Dismount-DefaultUserHive
  $WhatIfPreference = $_whatIfBackup
}

if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow
}
elseif ($anyChanges) {
  Write-Log -Message "`nPrivacy settings have been processed." -Color Green
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $privacySettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-Privacy'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
