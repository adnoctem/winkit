#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Configures Windows Explorer (File Explorer) folder options to a preferred state.

.DESCRIPTION
  Applies a curated set of registry values under HKCU:\Software\Microsoft\Windows\
  CurrentVersion\Explorer that control Explorer's behaviour - folder view, navigation
  pane, search preferences, and privacy settings.  Uses the winkit registry helpers
  for idempotent, safe writes.  -Undo restores Windows defaults.  -Instant restarts
  Explorer so the changes take effect immediately.  -SysPrep mounts the
  default user profile hive (C:\Users\Default\NTUSER.DAT) and writes to it
  for system imaging scenarios.

.PARAMETER Undo
  Restore Explorer settings to Windows defaults.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER Instant
  Restart Windows Explorer after applying changes so they take effect immediately.
  Safe to use with -Undo.

.PARAMETER Config
  Path to a JSON file containing an array of setting overrides.  Each entry in
  the JSON needs at minimum a "Name" field (matching the desired setting);
  "Preferred" and/or "Default" fields replace the corresponding values in the
  built-in defaults.  Entries that do not match any known setting are skipped
  with a warning.

.PARAMETER ExportConfig
  Export the default Explorer settings as JSON to the console.  Use -ExportPath
  to write to a file instead.  Cannot be combined with -DryRun.

.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  When used together with -ExportConfig, writes the JSON to this file path
  instead of printing to the console.

.PARAMETER SysPrep
  Mount the default user profile hive (C:\Users\Default\NTUSER.DAT) and write
  Explorer settings there instead of HKCU.  Use this when preparing a system
  image (sysprep / audit mode) so that newly created user profiles inherit the
  configured Explorer defaults.  Cannot be combined with -Instant.  Requires
  elevation.

.EXAMPLE
  PS> ./Configure-Explorer.ps1
  Applies preferred Explorer settings via the registry.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -Instant
  Applies preferred Explorer settings and restarts Explorer immediately.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -Undo
  Restores the default Windows Explorer settings.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -Undo -Instant
  Restores defaults and restarts Explorer immediately.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -DryRun
  Shows which registry values would be modified without making any changes.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -Undo -DryRun
  Previews the undo operation without touching the registry.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -ExportConfig
  Prints the default settings template to the console and exits.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -ExportConfig -ExportPath '.\my-explorer.json'
  Exports the default settings template to .\my-explorer.json and exits.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -Config '.\my-explorer.json'
  Applies Explorer settings, overriding built-in defaults with values from the
  JSON file.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -Config '.\my-explorer.json' -DryRun
  Previews which registry values would be set after merging the JSON overrides.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -Config '.\my-explorer.json' -Instant
  Applies settings (with JSON overrides) and restarts Explorer immediately.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -SysPrep
  Mounts the default user profile hive and writes preferred Explorer settings
  so they are inherited by new user profiles (sysprep / imaging scenario).

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -SysPrep -Undo
  Restores Windows Explorer defaults in the default user profile hive.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -SysPrep -Config '.\my-explorer.json'
  Writes Explorer settings (with JSON overrides) to the default user profile.

.EXAMPLE
  PS> ./Configure-Explorer.ps1 -SysPrep -DryRun
  Previews which registry values would be written to the default user profile.

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
    HelpMessage = 'Restore Explorer settings to Windows defaults.'
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
    HelpMessage = 'Restart Windows Explorer after applying changes so they take effect immediately.'
  )]
  [switch]
  $Instant,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Mount the default user profile hive and write settings there (sysprep / imaging mode). Cannot be used with -Instant.'
  )]
  [switch]
  $SysPrep,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Path to a JSON config file that overrides individual Explorer settings.'
  )]
  [string]
  $Config,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export the default Explorer settings to the console (or to a file with -ExportPath).'
  )]
  [switch]
  $ExportConfig,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export current Explorer registry values to reusable JSON config.'
  )]
  [switch]
  $ExportCurrentState,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'File path for -ExportConfig.  When omitted the settings are printed to the console.'
  )]
  [string]
  $ExportPath,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

# When -DryRun is active, enable WhatIf for all downstream lib calls.
if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

# -- SysPrep validation ------------------------------------------------
if ($SysPrep -and $Instant) {
  Write-Log -Message '-SysPrep cannot be combined with -Instant. SysPrep mode writes to the default user profile hive (mounted from disk); restarting Explorer would only affect the current user session.' -Color Red
  exit 1
}

# Registry target keys - when -SysPrep is active, redirect to the mounted
# default user hive (Registry::HKEY_USERS\DefaultUser) instead of HKCU
$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }
$explorerKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Explorer"
$advancedKey = "$explorerKey\Advanced"
$cabinetStateKey = "$explorerKey\CabinetState"
$searchPrefKey = "$explorerKey\Search\Preferences"
$classesKey = "$regHive\Software\Classes"
$galleryUserKey = "$classesKey\CLSID\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c}"
$homeUserKey = "$classesKey\CLSID\{f874310e-b6b7-47dc-bc84-b9e6b38f5903}"
$oneDriveUserKey = "$classesKey\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
$navPaneGalleryKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\NavPane\ShowGallery'
$navPaneHomeKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\NavPane\ShowHome'
$navPaneOneDriveKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\NavPane\ShowOneDrive'
$currentBuild = Get-OSBuildNumber

# Explorer settings - each entry holds the registry path, value name, our
# preferred value (Preferred), the Windows default (Default), and the
# registry type.  A Default of $null means "remove the value" on undo.
$explorerSettings = @(
  # -- General tab --
  @{
    Path = $explorerKey
    Name = 'LaunchTo'
    Preferred = 2
    Default = 1
    Type = 'DWord'
    Description = 'Open File Explorer to: Home'
  }
  @{
    Path = $advancedKey
    Name = 'SeparateProcess'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Browse folders: Open each folder in the same window'
  }
  @{
    Path = $explorerKey
    Name = 'ShowRecent'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Privacy: Hide recently used files in Quick Access'
  }
  @{
    Path = $explorerKey
    Name = 'ShowFrequent'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Privacy: Hide frequently used folders in Quick Access'
  }

  # -- View tab --
  @{
    Path = $advancedKey
    Name = 'IconsOnly'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Always show icons, never thumbnails: OFF'
  }
  @{
    Path = $advancedKey
    Name = 'ShowTypeOverlay'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Display file icon on thumbnails: ON'
  }
  @{
    Path = $cabinetStateKey
    Name = 'FullPath'
    Preferred = 1
    Default = 0
    Type = 'DWord'
    Description = 'Display the full path in the title bar: ON'
  }
  @{
    Path = $advancedKey
    Name = 'Hidden'
    Preferred = 1
    Default = 2
    Type = 'DWord'
    Description = 'Show hidden files, folders, and drives'
  }
  @{
    Path = $advancedKey
    Name = 'HideDrivesWithNoMedia'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Hide empty drives: OFF'
  }
  @{
    Path = $advancedKey
    Name = 'HideFileExt'
    Preferred = 0
    Default = 1
    Type = 'DWord'
    Description = 'Hide extensions for known file types: OFF'
  }
  @{
    Path = $advancedKey
    Name = 'ShowSuperHidden'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Hide protected operating system files: ON'
  }
  @{
    Path = $advancedKey
    Name = 'ShowCompColor'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Show encrypted or compressed NTFS files in color: OFF'
  }
  @{
    Path = $advancedKey
    Name = 'ShowInfoTip'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Show pop-up description for folder and desktop items: ON'
  }
  @{
    Path = $advancedKey
    Name = 'ShowStatusBar'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Show status bar: ON'
  }
  @{
    Path = $advancedKey
    Name = 'AutoCheckSelect'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Use check boxes to select items: OFF'
  }
  @{
    Path = $advancedKey
    Name = 'UseCompactMode'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Decrease space between items (compact view): OFF'
  }
  @{
    Path = $advancedKey
    Name = 'FolderContentsInfoTip'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Display file size information in folder tips: ON'
  }
  @{
    Path = $advancedKey
    Name = 'HideMergeConflicts'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Hide folder merge conflicts: ON'
  }
  @{
    Path = $advancedKey
    Name = 'ShowSyncProviderNotifications'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Show sync provider notifications: ON'
  }
  @{
    Path = $advancedKey
    Name = 'ShowPreviewHandlers'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Show preview handlers in preview pane: ON'
  }
  @{
    Path = $advancedKey
    Name = 'SharingWizardOn'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Use Sharing Wizard: ON'
  }
  @{
    Path = $advancedKey
    Name = 'PersistBrowsers'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Restore previous folder windows at logon: OFF'
  }
  @{
    Path = $advancedKey
    Name = 'TypeAhead'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'When typing into list view: Select the typed item'
  }

  # -- Navigation pane --
  @{
    Path = $advancedKey
    Name = 'NavPaneExpandToCurrentFolder'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Expand to open folder: OFF'
  }
  @{
    Path = $advancedKey
    Name = 'NavPaneShowAllFolders'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Show all folders: OFF'
  }
  @{
    Path = $explorerKey
    Name = 'ShowDriveLettersFirst'
    Preferred = 4
    Default = 0
    Type = 'DWord'
    Description = 'Show drive letters before drive labels'
  }
  @{
    Path = $galleryUserKey
    Name = 'System.IsPinnedToNameSpaceTree'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Hide Gallery from the navigation pane'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'CheckedValue'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: checked value'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'DefaultValue'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: default value'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'HKeyRoot'
    Preferred = 2147483649
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: HKCU root'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'Id'
    Preferred = 13
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: id'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'RegPath'
    Preferred = 'Software\\Classes\\CLSID\\{e88865ea-0e1c-4e20-9aa6-edcd0212c87c}'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: registry path'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'Text'
    Preferred = 'Show Gallery'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: display text'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'Type'
    Preferred = 'checkbox'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: option type'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'UncheckedValue'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: unchecked value'
  }
  @{
    Path = $navPaneGalleryKey
    Name = 'ValueName'
    Preferred = 'System.IsPinnedToNameSpaceTree'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Gallery nav pane option metadata: value name'
  }
  @{
    Path = $homeUserKey
    Name = ''
    Preferred = 'CLSID_MSGraphHomeFolder'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Home navigation CLSID label'
  }
  @{
    Path = $homeUserKey
    Name = 'System.IsPinnedToNameSpaceTree'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Hide Home from the navigation pane'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'CheckedValue'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: checked value'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'DefaultValue'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: default value'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'HKeyRoot'
    Preferred = 2147483649
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: HKCU root'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'Id'
    Preferred = 13
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: id'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'RegPath'
    Preferred = 'Software\\Classes\\CLSID\\{f874310e-b6b7-47dc-bc84-b9e6b38f5903}'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: registry path'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'Text'
    Preferred = 'Show Home'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: display text'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'Type'
    Preferred = 'checkbox'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: option type'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'UncheckedValue'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: unchecked value'
  }
  @{
    Path = $navPaneHomeKey
    Name = 'ValueName'
    Preferred = 'System.IsPinnedToNameSpaceTree'
    Default = $null
    Type = 'String'
    MinBuild = 22000
    Description = 'Home nav pane option metadata: value name'
  }
  @{
    Path = $oneDriveUserKey
    Name = 'System.IsPinnedToNameSpaceTree'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Hide OneDrive from the navigation pane'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'CheckedValue'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    Description = 'OneDrive nav pane option metadata: checked value'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'DefaultValue'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'OneDrive nav pane option metadata: default value'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'HKeyRoot'
    Preferred = 2147483649
    Default = $null
    Type = 'DWord'
    Description = 'OneDrive nav pane option metadata: HKCU root'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'Id'
    Preferred = 13
    Default = $null
    Type = 'DWord'
    Description = 'OneDrive nav pane option metadata: id'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'RegPath'
    Preferred = 'Software\\Classes\\CLSID\\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
    Default = $null
    Type = 'String'
    Description = 'OneDrive nav pane option metadata: registry path'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'Text'
    Preferred = 'Show OneDrive'
    Default = $null
    Type = 'String'
    Description = 'OneDrive nav pane option metadata: display text'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'Type'
    Preferred = 'checkbox'
    Default = $null
    Type = 'String'
    Description = 'OneDrive nav pane option metadata: option type'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'UncheckedValue'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'OneDrive nav pane option metadata: unchecked value'
  }
  @{
    Path = $navPaneOneDriveKey
    Name = 'ValueName'
    Preferred = 'System.IsPinnedToNameSpaceTree'
    Default = $null
    Type = 'String'
    Description = 'OneDrive nav pane option metadata: value name'
  }

  # -- Search tab --
  @{
    Path = $searchPrefKey
    Name = 'WholeFileSystem'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = "Don't use index when searching system files: OFF"
  }
  @{
    Path = $searchPrefKey
    Name = 'SystemFolders'
    Preferred = 1
    Default = 1
    Type = 'DWord'
    Description = 'Include system directories in non-indexed searches: ON'
  }
  @{
    Path = $searchPrefKey
    Name = 'ArchivedFiles'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    Description = 'Include compressed files (ZIP, CAB): OFF'
  }
  @{
    Path = $advancedKey
    Name = 'Start_SearchFiles'
    Preferred = 2
    Default = 2
    Type = 'DWord'
    Description = 'Always search file names and contents: OFF'
  }

  # -- Navigation pane: context menu entries --
  @{
    Path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
    Name = '{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'
    Preferred = ''
    Default = $null
    Type = 'String'
    Description = 'Remove "Share" from the context menu (Win10/Win11).'
  }
  @{
    Path = 'HKCU:\Software\Classes\CLSID\{e2bf9676-5f8f-435c-97eb-11607a5bedf7}'
    Name = 'System.IsPinnedToNameSpaceTree'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Hide Share entry from navigation pane.'
  }
  @{
    Path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
    Name = '{3e22fe5c-40ed-41f4-b551-9c6a7b5b2e07}'
    Preferred = ''
    Default = $null
    Type = 'String'
    Description = 'Remove "Include in library" from the context menu.'
  }
  @{
    Path = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Blocked'
    Name = '{F81E9010-6EA4-11CE-A7FF-00AA003CA9F6}'
    Preferred = ''
    Default = $null
    Type = 'String'
    Description = 'Remove "Give access to" from the context menu.'
  }

  # -- Duplicate drive visibility --
  @{
    Path = "$classesKey\CLSID\{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}"
    Name = 'System.IsPinnedToNameSpaceTree'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Description = 'Hide duplicate removable drives from navigation pane.'
  }
)

# -- Export config (print to console or write to disk and exit) ----------
if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $explorerSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current Explorer settings exported to: $_exportPath" -Color Green
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
    $explorerSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default Explorer settings exported to: $_exportPath" -Color Green
  }
  else {
    $explorerSettings | ConvertTo-Json -Depth 3
  }
  exit 0
}

# -- Merge user config (override built-in defaults with JSON values) ------
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

  Merge-ObjectArrays -Base $explorerSettings -Overrides $_overrides

  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

# -- SysPrep: mount the default user hive -----------------------------
if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference
  $WhatIfPreference = $false  # mount is infrastructure, not a config change
  $mountResult = Mount-DefaultUserHive
  $WhatIfPreference = $_whatIfBackup
  if (-not $mountResult) {
    Write-Log -Message 'Failed to mount the default user hive. Ensure you are running elevated and C:\Users\Default\NTUSER.DAT exists.' -Color Red
    exit 1
  }
}

$_eligibleSettings = @($explorerSettings | Where-Object {
    (-not $_.ContainsKey('MinBuild') -or $currentBuild -ge $_.MinBuild) -and
    (-not $_.ContainsKey('MaxBuild') -or $currentBuild -le $_.MaxBuild)
  })

$_skippedSettings = @($explorerSettings | Where-Object {
    ($_.ContainsKey('MinBuild') -and $currentBuild -lt $_.MinBuild) -or
    ($_.ContainsKey('MaxBuild') -and $currentBuild -gt $_.MaxBuild)
  })

foreach ($_entry in $_skippedSettings) {
  Write-Log -Message "Skipping Explorer setting '$($_entry.Name)' on build $currentBuild - $($_entry.Description)" -Color Yellow
}

$targetLabel = if ($Undo) { 'Restoring' } else { 'Applying' }
$anyChanges = $false

# Apply registry values
foreach ($entry in $_eligibleSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }

  if ($Undo -and $null -eq $entry.Default) {
    # Default is absence of the value - remove it
    Write-Log -Message "$targetLabel Explorer setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel Explorer setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    $result = Set-RegistryValue -Path $entry.Path -Name $entry.Name -Value $targetValue -Type $entry.Type
  }

  if ($result) {
    Write-Log -Message "  -> $($result.Status)" -Color Gray
    if ($result.Status -in @('Created', 'Updated', 'Removed')) {
      $anyChanges = $true
    }
  }
  else {
    Write-Log -Message "  -> FAILED - could not process '$($entry.Name)'" -Color Red
  }
}

# -- SysPrep: dismount the default user hive --------------------------
if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference
  $WhatIfPreference = $false
  Dismount-DefaultUserHive
  $WhatIfPreference = $_whatIfBackup
}

# Instant apply - restart Explorer
if ($Instant -and -not $DryRun) {
  Write-Log -Message "`nRestarting Windows Explorer ..." -Color Yellow
  Stop-Process -Name explorer -Force
  Start-Process explorer
  Write-Log -Message '  -> Done - Explorer restarted.' -Color Green
}
elseif ($Instant -and $DryRun) {
  Write-Log -Message "`n[DRY RUN] Would restart Windows Explorer." -Color Yellow
}

# Summary
if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow
}
elseif ($anyChanges) {
  if ($Undo) {
    Write-Log -Message "`nExplorer settings have been restored to Windows defaults." -Color Green
  }
  else {
    Write-Log -Message "`nExplorer settings have been applied." -Color Green
  }
  if ($SysPrep) {
    Write-Log -Message 'Settings were written to the default user profile hive - new user profiles will inherit them.' -Color Yellow
  }
  elseif (-not $Instant) {
    Write-Log -Message 'Use -Instant or restart Explorer for changes to take effect.' -Color Yellow
  }
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $_eligibleSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-Explorer'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
