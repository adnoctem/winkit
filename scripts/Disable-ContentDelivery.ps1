#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Disables Windows automatic Content Delivery app suggestions and Store app provisioning.

.DESCRIPTION
  Applies the ContentDeliveryManager, CloudContent, and Windows Store policy
  values that suppress suggested app provisioning, sponsored content delivery,
  OEM/preinstalled app reprovisioning, and automatic Store app downloads. This
  script deliberately owns these values separately from Configure-Privacy.ps1
  because they affect Windows content/app delivery behavior rather than general
  privacy toggles. Use -Undo to remove the managed values, -SysPrep to target
  the default user hive for HKCU-backed settings, and -ExportConfig to generate
  a JSON override template.

.PARAMETER Undo
  Remove Content Delivery values managed by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed Content Delivery
  values there instead of HKCU. Machine-wide policy values are still written to
  HKLM. Requires elevation.

.PARAMETER Config
  JSON file containing setting overrides. Each entry needs at minimum a "Name"
  field matching a built-in setting; Preferred and Default can be overridden.

.PARAMETER ExportConfig
  Export the default Content Delivery settings JSON and exit.

.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig. When omitted, JSON is written to stdout.

.EXAMPLE
  PS> ./Disable-ContentDelivery.ps1
  Disables suggested app provisioning and automatic Store app downloads.

.EXAMPLE
  PS> ./Disable-ContentDelivery.ps1 -DryRun
  Shows every registry value that would be written or removed.

.EXAMPLE
  PS> ./Disable-ContentDelivery.ps1 -Undo
  Removes the Content Delivery values managed by this script.

.EXAMPLE
  PS> ./Disable-ContentDelivery.ps1 -SysPrep
  Writes HKCU-backed settings into the default user profile hive so new users
  inherit the Content Delivery suppression profile.

.EXAMPLE
  PS> ./Disable-ContentDelivery.ps1 -ExportConfig -ExportPath '.\content-delivery.json'
  Exports the default settings template to a JSON file.

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
    HelpMessage = 'Remove Content Delivery values managed by this script.'
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
    HelpMessage = 'Mount the default user profile hive and write HKCU-backed Content Delivery values there.'
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
    HelpMessage = 'Export the default Content Delivery settings JSON and exit.'
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

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }
$contentDeliveryKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
$cloudContentKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent'
$windowsStorePolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore'

$contentDeliverySettings = @(
  @{
    Path = $contentDeliveryKey
    Name = 'FeatureManagementEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable Content Delivery feature management for suggested app provisioning.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'ContentDeliveryAllowed'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable Content Delivery suggested app provisioning.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'OemPreInstalledAppsEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable OEM suggested app provisioning.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'PreInstalledAppsEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable preinstalled suggested app provisioning.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'PreInstalledAppsEverEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable preinstalled suggested app reprovisioning state.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContentEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable generic subscribed content provisioning.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-310093Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable suggested content subscription 310093.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-338387Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable lock screen tips subscription.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-338388Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable suggested content subscription 338388.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-338389Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable suggested content subscription 338389.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-338393Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable suggested content subscription 338393.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-353694Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable suggested content subscription 353694.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-353696Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable suggested content subscription 353696.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SubscribedContent-353698Enabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SubscribedContent'
    Description = 'Disable suggested content subscription 353698.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SilentInstalledAppsEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable silent suggested app installs.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SystemPaneSuggestionsEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable Settings/System pane suggestions.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'SoftLandingEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable soft landing tips.'
  }
  @{
    Path = $contentDeliveryKey
    Name = 'RotatingLockScreenOverlayEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'ContentDelivery'
    Description = 'Disable rotating lock screen overlay.'
  }
  @{
    Path = $cloudContentKey
    Name = 'DisableWindowsConsumerFeatures'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'Policy'
    Description = 'Disable Windows consumer feature suggestions.'
  }
  @{
    Path = $windowsStorePolicyKey
    Name = 'AutoDownload'
    Preferred = 2
    Default = $null
    Type = 'DWord'
    Group = 'Policy'
    Description = 'Disable automatic Store app downloads through policy.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $contentDeliverySettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current content delivery settings exported to: $_exportPath" -Color Green
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
    $contentDeliverySettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default Content Delivery settings exported to: $_exportPath" -Color Green
  }
  else {
    $contentDeliverySettings | ConvertTo-Json -Depth 3
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
  Merge-ObjectArrays -Base $contentDeliverySettings -Overrides $_overrides
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

$targetLabel = if ($Undo) { 'Removing' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $contentDeliverySettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if (($Undo -and $null -eq $entry.Default) -or ((-not $Undo) -and $null -eq $entry.Preferred)) {
    Write-Log -Message "$targetLabel Content Delivery setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel Content Delivery setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
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
  Write-Log -Message "`nContent Delivery settings have been processed." -Color Green
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $contentDeliverySettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Disable-ContentDelivery'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
