#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Configures Windows taskbar layout, buttons, search, widgets, and behavior.
.DESCRIPTION
  Applies a curated taskbar profile using winkit registry helpers. The exported
  JSON profile represents each registry value once, including mutually exclusive
  modes such as search display, taskbar icon-only combining, and multi-monitor
  behavior.
  Use -Undo to restore known Windows defaults or remove values where the source
  material only provided delete-on-undo behavior. Build-gated settings are
  skipped automatically on unsupported Windows releases.
.PARAMETER Undo
  Restore defaults or remove values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER Instant
  Restart Windows Explorer after applying changes.
.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed taskbar settings
  there instead of HKCU. HKLM policy values still target the live machine hive.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default taskbar settings JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.EXAMPLE
  PS> ./Configure-Taskbar.ps1
  Applies the default taskbar profile.
.EXAMPLE
  PS> ./Configure-Taskbar.ps1 -DryRun
  Shows taskbar registry changes without applying them.
.EXAMPLE
  PS> ./Configure-Taskbar.ps1 -Instant
  Applies taskbar settings and restarts Explorer so shell changes are visible.
.EXAMPLE
  PS> ./Configure-Taskbar.ps1 -Undo
  Restores known defaults or removes managed values.
.EXAMPLE
  PS> ./Configure-Taskbar.ps1 -ExportConfig -ExportPath '.\taskbar.json'
  Exports the default taskbar settings template.
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
    HelpMessage = 'Restart Windows Explorer after applying changes.'
  )]
  [switch]
  $Instant,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Mount the default user profile hive and write HKCU-backed taskbar settings there.'
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
    HelpMessage = 'Export the default taskbar settings JSON and exit.'
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

if ($SysPrep -and $Instant) {
  Write-Log -Message '-SysPrep cannot be combined with -Instant.' -Color Red
  exit 1
}

$currentBuild = Get-OSBuildNumber
$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }
$advancedKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$searchKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Search"
$policyExplorerKey = "$regHive\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
$widgetsPolicyManagerKey = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\default\NewsAndInterests\AllowNewsAndInterests'
$widgetsDshKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh'

$taskbarSettings = @(
  @{
    Path = $advancedKey
    Name = 'TaskbarAl'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Alignment'
    Description = 'Taskbar alignment: left. Windows centered default is 1.'
  }
  @{
    Path = $searchKey
    Name = 'SearchboxTaskbarMode'
    Preferred = 0
    Default = 2
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Search'
    Description = 'Taskbar search display mode. 0 hide, 1 icon, 2 box, 3 label.'
  }
  @{
    Path = $advancedKey
    Name = 'ShowTaskViewButton'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Buttons'
    Description = 'Hide Task View button.'
  }
  @{
    Path = $advancedKey
    Name = 'TaskbarMn'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MaxBuild = 22621
    Group = 'Buttons'
    Description = 'Hide Chat / Meet Now button.'
  }
  @{
    Path = $policyExplorerKey
    Name = 'HideSCAMeetNow'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MaxBuild = 22621
    Group = 'Buttons'
    Description = 'Hide Meet Now policy button.'
  }
  @{
    Path = $widgetsPolicyManagerKey
    Name = 'value'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Buttons'
    Description = 'Disable Widgets / News and Interests through PolicyManager.'
  }
  @{
    Path = $widgetsDshKey
    Name = 'AllowNewsAndInterests'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    Group = 'Buttons'
    Description = 'Disable Widgets / News and Interests policy.'
  }
  @{
    Path = "$advancedKey\TaskbarDeveloperSettings"
    Name = 'TaskbarEndTask'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22631
    Group = 'Behavior'
    Description = 'Show End Task in taskbar context menu.'
  }
  @{
    Path = $advancedKey
    Name = 'LastActiveClick'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Behavior'
    Description = 'Enable last active click behavior for grouped apps.'
  }
  @{
    Path = $advancedKey
    Name = 'TaskbarGlomLevel'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Combining'
    Description = 'Main taskbar combining. 0 always combine and hide labels, 1 when full, 2 never.'
  }
  @{
    Path = $advancedKey
    Name = 'MMTaskbarGlomLevel'
    Preferred = 0
    Default = 0
    Type = 'DWord'
    MinBuild = 22000
    Group = 'Combining'
    Description = 'Secondary taskbar combining. 0 always combine and hide labels, 1 when full, 2 never.'
  }
  @{
    Path = $advancedKey
    Name = 'MMTaskbarMode'
    Preferred = 1
    Default = 0
    Type = 'DWord'
    MinBuild = 22000
    Group = 'MultiMonitor'
    Description = 'Multi-monitor taskbar mode. 0 all, 1 main and active, 2 active only.'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $taskbarSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current taskbar settings exported to: $_exportPath" -Color Green
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
    $taskbarSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default taskbar settings exported to: $_exportPath" -Color Green
  }
  else {
    $taskbarSettings | ConvertTo-Json -Depth 3
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
  Merge-ObjectArrays -Base $taskbarSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$_eligibleSettings = @($taskbarSettings | Where-Object {
    (-not $_.ContainsKey('MinBuild') -or $currentBuild -ge $_.MinBuild) -and
    (-not $_.ContainsKey('MaxBuild') -or $currentBuild -le $_.MaxBuild)
  })
$_skippedSettings = @($taskbarSettings | Where-Object {
    ($_.ContainsKey('MinBuild') -and $currentBuild -lt $_.MinBuild) -or
    ($_.ContainsKey('MaxBuild') -and $currentBuild -gt $_.MaxBuild)
  })

foreach ($_entry in $_skippedSettings) {
  Write-Log -Message "Skipping taskbar setting '$($_entry.Name)' on build $currentBuild - $($_entry.Description)" -Color Yellow
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
    Write-Log -Message "$targetLabel taskbar setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel taskbar setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
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

if ($Instant -and -not $DryRun) {
  Write-Log -Message "`nRestarting Windows Explorer ..." -Color Yellow
  Stop-Process -Name explorer -Force
  Start-Process explorer
  Write-Log -Message '  -> Done - Explorer restarted.' -Color Green
}
elseif ($Instant -and $DryRun) {
  Write-Log -Message "`n[DRY RUN] Would restart Windows Explorer." -Color Yellow
}

if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow
}
elseif ($anyChanges) {
  Write-Log -Message "`nTaskbar settings have been processed." -Color Green
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $_eligibleSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-Taskbar'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
