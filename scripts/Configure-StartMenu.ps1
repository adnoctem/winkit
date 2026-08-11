#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Configures Start menu and Windows Search integration settings.
.DESCRIPTION
  Applies Start menu and Search policy/preference values using winkit registry
  helpers. This includes recommended content, Phone Link integration, Bing/web
  search, local search history, and search highlights. Build-gated settings are
  skipped automatically when the running Windows build does not support them.
.PARAMETER Undo
  Restore defaults or remove Start/Search values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed settings there
  instead of HKCU. Machine-wide policy values are still written to HKLM.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportTemplate
  Export the default Start menu settings as a JSON config template and exit. Use
  -ExportPath to write it to a file instead of printing to the console.
.PARAMETER ExportState
  Export the current Start menu settings as reusable JSON config and exit.
.PARAMETER ExportPath
  File path for -ExportTemplate or -ExportState. When omitted, the JSON is
  written to the console.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> ./Configure-StartMenu.ps1
  Applies the default Start menu and Search profile.
.EXAMPLE
  PS> ./Configure-StartMenu.ps1 -DryRun
  Previews Start/Search registry changes without applying them.
.EXAMPLE
  PS> ./Configure-StartMenu.ps1 -Undo
  Removes or restores values managed by this script.
.EXAMPLE
  PS> ./Configure-StartMenu.ps1 -SysPrep
  Writes HKCU-backed Start/Search defaults to the default user profile hive.
.EXAMPLE
  PS> ./Configure-StartMenu.ps1 -ExportTemplate -ExportPath '.\start-menu.json'
  Exports the default Start menu settings template.
.EXAMPLE
  PS> ./Configure-StartMenu.ps1 -ExportState -ExportPath '.\start-menu-current.json'
  Exports the current Start menu settings to a JSON file.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Restore defaults or remove Start/Search values managed by this script.'
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
    HelpMessage = 'Mount the default user profile hive and write HKCU-backed settings there.'
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
    HelpMessage = 'Export the default Start menu settings as a JSON config template and exit.'
  )]
  [switch]
  $ExportTemplate,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export the current Start menu settings as reusable JSON config and exit.'
  )]
  [switch]
  $ExportState,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'File path for -ExportTemplate or -ExportState. When omitted, the JSON is written to the console.'
  )]
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

$currentBuild = Get-OSBuildNumber
$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }
$userExplorerPolicyKey = "$regHive\Software\Policies\Microsoft\Windows\Explorer"
$userExplorerAdvancedKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$userSearchSettingsKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\SearchSettings"
$userStartCompanionKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Start\Companions\Microsoft.YourPhone_8wekyb3d8bbwe"
$userPoliciesExplorerKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"

$machineSearchPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search'
$machineExplorerPolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer'
$machinePolicyStartKey = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Start'
$machinePolicyEducationKey = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Education'

$startMenuSettings = @(
  @{
    Path = $userExplorerPolicyKey
    Name = 'DisableSearchBoxSuggestions'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Group = 'SearchWeb'
    Description = 'Disable Bing web suggestions in Start/Search.'
  }
  @{
    Path = $machineSearchPolicyKey
    Name = 'AllowCortana'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SearchWeb'
    Description = 'Disable Cortana through Windows Search policy.'
  }
  @{
    Path = $machineSearchPolicyKey
    Name = 'CortanaConsent'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SearchWeb'
    Description = 'Deny Cortana consent through policy.'
  }
  @{
    Path = $machineSearchPolicyKey
    Name = 'ConnectedSearchUseWeb'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SearchWeb'
    Description = 'Disable connected web search usage.'
  }
  @{
    Path = $machineSearchPolicyKey
    Name = 'AllowSearchToUseLocation'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SearchWeb'
    Description = 'Prevent Search from using location.'
  }
  @{
    Path = $userSearchSettingsKey
    Name = 'IsDeviceSearchHistoryEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'SearchHistory'
    Description = 'Disable local device search history.'
  }
  @{
    Path = $userSearchSettingsKey
    Name = 'IsDynamicSearchBoxEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Group = 'SearchHighlights'
    Description = 'Disable Search Highlights.'
  }
  @{
    Path = $userStartCompanionKey
    Name = 'IsEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Group = 'StartCompanions'
    Description = 'Disable Phone Link Start menu companion.'
  }
  @{
    Path = $userPoliciesExplorerKey
    Name = 'NoStartMenuMorePrograms'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 26200
    Group = 'StartCompanions'
    Description = 'Hide All Apps section in Start menu.'
  }
  @{
    Path = $machineExplorerPolicyKey
    Name = 'HideRecommendedSection'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Group = 'StartRecommended'
    Description = 'Hide Recommended section through Explorer policy.'
  }
  @{
    Path = $machinePolicyStartKey
    Name = 'HideRecommendedSection'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Group = 'StartRecommended'
    Description = 'Hide Recommended section through PolicyManager Start CSP.'
  }
  @{
    Path = $machinePolicyEducationKey
    Name = 'IsEducationEnvironment'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Group = 'StartRecommended'
    Description = 'Use education environment mode to suppress Recommended content.'
  }
  @{
    Path = $userExplorerAdvancedKey
    Name = 'Start_Layout'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Group = 'StartRecommended'
    Description = 'Use more pins / less recommendations Start layout.'
  }
)

if ($ExportState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportState.' -Color Red; exit 1 }
  if ($ExportTemplate) { Write-Log -Message '-ExportTemplate cannot be combined with -ExportState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $startMenuSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current Start menu settings exported to: $_exportPath" -Color Green
  }
  else { $_currentState | ConvertTo-Json -Depth 3 }
  exit 0
}

if ($ExportTemplate) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportTemplate.' -Color Red; exit 1 }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $startMenuSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default Start menu settings exported to: $_exportPath" -Color Green
  }
  else {
    $startMenuSettings | ConvertTo-Json -Depth 3
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
  Merge-ObjectArrays -Base $startMenuSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$_eligibleSettings = @($startMenuSettings | Where-Object {
    (-not $_.ContainsKey('MinBuild') -or $currentBuild -ge $_.MinBuild) -and
    (-not $_.ContainsKey('MaxBuild') -or $currentBuild -le $_.MaxBuild)
  })
$_skippedSettings = @($startMenuSettings | Where-Object {
    ($_.ContainsKey('MinBuild') -and $currentBuild -lt $_.MinBuild) -or
    ($_.ContainsKey('MaxBuild') -and $currentBuild -gt $_.MaxBuild)
  })
foreach ($_entry in $_skippedSettings) {
  Write-Log -Message "Skipping Start menu setting '$($_entry.Name)' on build $currentBuild - $($_entry.Description)" -Color Yellow
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

foreach ($entry in $_eligibleSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel Start menu setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel Start menu setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
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
  Write-Log -Message "`nStart menu and Search settings have been processed." -Color Green
  Write-Log -Message 'Sign out and back in, or restart Windows, for all changes to take effect.' -Color Yellow
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $startMenuSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-StartMenu'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
