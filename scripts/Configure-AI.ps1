#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Configures Windows AI-related features to the winkit preferred state.

.DESCRIPTION
  Applies a curated set of registry values that disable or constrain Windows 11
  AI integrations such as Copilot, Recall, Click To Do, Edge AI, Paint AI,
  Notepad AI and the Windows AI service auto-start behavior. Uses the winkit
  registry helpers for idempotent writes. -Undo removes the configured values,
  matching the delete-on-undo behavior documented in the source material.
  -SysPrep mounts the default user profile hive and writes HKCU-backed settings
  there for system imaging scenarios.

.PARAMETER Undo
  Remove the AI policy/configuration values managed by this script.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER SysPrep
  Mount the default user profile hive (C:\Users\Default\NTUSER.DAT) and write
  HKCU-backed settings there instead of HKCU. Machine-wide HKLM settings are
  still written to HKLM. Requires elevation.

.PARAMETER Config
  Path to a JSON file containing an array of setting overrides. Each entry in
  the JSON needs at minimum a "Name" field matching a known registry value;
  "Preferred" and/or "Default" fields replace the corresponding built-in
  values. Entries that do not match any known setting are skipped.

.PARAMETER ExportConfig
  Export the default AI settings as JSON to the console. Use -ExportPath to
  write to a file instead. Cannot be combined with -DryRun.

.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  When used together with -ExportConfig, writes the JSON to this file path
  instead of printing to the console.

.EXAMPLE
  PS> ./Configure-AI.ps1
  Applies preferred AI-related Windows policy settings.

.EXAMPLE
  PS> ./Configure-AI.ps1 -Undo
  Removes AI-related policy settings managed by this script.

.EXAMPLE
  PS> ./Configure-AI.ps1 -DryRun
  Shows which registry values would be modified without making any changes.

.EXAMPLE
  PS> ./Configure-AI.ps1 -ExportConfig -ExportPath '.\ai-settings.json'
  Exports the default AI settings template to .\ai-settings.json and exits.

.EXAMPLE
  PS> ./Configure-AI.ps1 -SysPrep
  Writes HKCU-backed AI settings to the default user profile hive so new users
  inherit them.

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
    HelpMessage = 'Remove the AI policy/configuration values managed by this script.'
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
    HelpMessage = 'Path to a JSON config file that overrides individual AI settings.'
  )]
  [string]
  $Config,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export the default AI settings to the console or to a file with -ExportPath.'
  )]
  [switch]
  $ExportConfig,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export current AI registry values to reusable JSON config.'
  )]
  [switch]
  $ExportCurrentState,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'File path for -ExportConfig. When omitted the settings are printed to the console.'
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

$currentBuild = Get-OSBuildNumber
$currentUserHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }

$aiUserPolicyKey = "$currentUserHive\Software\Policies\Microsoft\Windows\WindowsAI"
$copilotUserExplorerKey = "$currentUserHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$copilotUserPolicyKey = "$currentUserHive\Software\Policies\Microsoft\Windows\WindowsCopilot"

$aiMachinePolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI'
$aiServiceKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\WSAIFabricSvc'
$copilotMachinePolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot'
$edgePolicyKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
$notepadPolicyKey = 'HKLM:\SOFTWARE\Policies\WindowsNotepad'
$paintPolicyKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Paint'

$aiSettings = @(
  @{
    Path = $aiServiceKey
    Name = 'Start'
    Preferred = 3
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableAISvcAutoStart: Set Windows AI service to manual start'
  }

  @{
    Path = $aiUserPolicyKey
    Name = 'DisableClickToDo'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableClickToDo: Disable Click To Do for the current/default user'
  }
  @{
    Path = $aiMachinePolicyKey
    Name = 'DisableClickToDo'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableClickToDo: Disable Click To Do machine-wide'
  }

  @{
    Path = $copilotUserExplorerKey
    Name = 'ShowCopilotButton'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableCopilot: Hide Copilot button in Explorer/taskbar UI'
  }
  @{
    Path = $copilotUserPolicyKey
    Name = 'TurnOffWindowsCopilot'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableCopilot: Turn off Copilot for the current/default user'
  }
  @{
    Path = $copilotMachinePolicyKey
    Name = 'TurnOffWindowsCopilot'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableCopilot: Turn off Copilot machine-wide'
  }

  @{
    Path = $edgePolicyKey
    Name = 'CopilotCDPPageContext'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Disable Copilot CDP page context'
  }
  @{
    Path = $edgePolicyKey
    Name = 'CopilotPageContext'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Disable Copilot page context'
  }
  @{
    Path = $edgePolicyKey
    Name = 'HubsSidebarEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Disable Edge hubs sidebar'
  }
  @{
    Path = $edgePolicyKey
    Name = 'EdgeEntraCopilotPageContext'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Disable Entra Copilot page context'
  }
  @{
    Path = $edgePolicyKey
    Name = 'EdgeHistoryAISearchEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Disable AI history search'
  }
  @{
    Path = $edgePolicyKey
    Name = 'ComposeInlineEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Disable inline compose'
  }
  @{
    Path = $edgePolicyKey
    Name = 'GenAILocalFoundationalModelSettings'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Restrict local generative AI model settings'
  }
  @{
    Path = $edgePolicyKey
    Name = 'NewTabPageBingChatEnabled'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableEdgeAI: Disable Bing Chat on new tab page'
  }

  @{
    Path = $notepadPolicyKey
    Name = 'DisableAIFeatures'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableNotepadAI: Disable Notepad AI features'
  }

  @{
    Path = $paintPolicyKey
    Name = 'DisableCocreator'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisablePaintAI: Disable Cocreator'
  }
  @{
    Path = $paintPolicyKey
    Name = 'DisableGenerativeFill'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisablePaintAI: Disable Generative Fill'
  }
  @{
    Path = $paintPolicyKey
    Name = 'DisableImageCreator'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisablePaintAI: Disable Image Creator'
  }
  @{
    Path = $paintPolicyKey
    Name = 'DisableGenerativeErase'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisablePaintAI: Disable Generative Erase'
  }
  @{
    Path = $paintPolicyKey
    Name = 'DisableRemoveBackground'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisablePaintAI: Disable Remove Background'
  }

  @{
    Path = $aiUserPolicyKey
    Name = 'DisableAIDataAnalysis'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableRecall: Disable AI data analysis for the current/default user'
  }
  @{
    Path = $aiMachinePolicyKey
    Name = 'DisableAIDataAnalysis'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableRecall: Disable AI data analysis machine-wide'
  }
  @{
    Path = $aiMachinePolicyKey
    Name = 'AllowRecallEnablement'
    Preferred = 0
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableRecall: Prevent Recall enablement'
  }
  @{
    Path = $aiMachinePolicyKey
    Name = 'TurnOffSavingSnapshots'
    Preferred = 1
    Default = $null
    Type = 'DWord'
    MinBuild = 22621
    Description = 'DisableRecall: Turn off saving snapshots'
  }
)

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $aiSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current AI settings exported to: $_exportPath" -Color Green
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
    $aiSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default AI settings exported to: $_exportPath" -Color Green
  }
  else {
    $aiSettings | ConvertTo-Json -Depth 3
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
  Merge-ObjectArrays -Base $aiSettings -Overrides $_overrides
  Write-Log -Message "  -> $($_overrides.Count) override(s) processed`n" -Color Gray
}

$_eligibleSettings = @($aiSettings | Where-Object {
    (-not $_.ContainsKey('MinBuild') -or $currentBuild -ge $_.MinBuild) -and
    (-not $_.ContainsKey('MaxBuild') -or $currentBuild -le $_.MaxBuild)
  })

$_skippedSettings = @($aiSettings | Where-Object {
    ($_.ContainsKey('MinBuild') -and $currentBuild -lt $_.MinBuild) -or
    ($_.ContainsKey('MaxBuild') -and $currentBuild -gt $_.MaxBuild)
  })

foreach ($_entry in $_skippedSettings) {
  Write-Log -Message "Skipping AI setting '$($_entry.Name)' on build $currentBuild - $($_entry.Description)" -Color Yellow
}

if ($_eligibleSettings.Count -eq 0) {
  Write-Log -Message "No AI settings apply to Windows build $currentBuild." -Color Yellow
  exit 0
}

if ($SysPrep) {
  $_whatIfBackup = $WhatIfPreference
  $WhatIfPreference = $false
  $mountResult = Mount-DefaultUserHive
  $WhatIfPreference = $_whatIfBackup

  if (-not $mountResult) {
    Write-Log -Message 'Failed to mount the default user hive. Ensure you are running elevated and C:\Users\Default\NTUSER.DAT exists.' -Color Red
    exit 1
  }
}

$targetLabel = if ($Undo) { 'Removing' } else { 'Applying' }
$anyChanges = $false

foreach ($entry in $_eligibleSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }

  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel AI setting: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }

    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel AI setting: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would set $($entry.Path)\$($entry.Name) = '$targetValue' ($($entry.Type))" -Color Gray
      continue
    }

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
  if ($Undo) {
    Write-Log -Message "`nAI settings managed by this script have been removed." -Color Green
  }
  else {
    Write-Log -Message "`nAI settings have been applied." -Color Green
  }

  if ($SysPrep) {
    Write-Log -Message 'HKCU-backed settings were written to the default user profile hive - new user profiles will inherit them.' -Color Yellow
  }
  else {
    Write-Log -Message 'Restart Windows or sign out and back in for all policy changes to take effect.' -Color Yellow
  }
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $_eligibleSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Configure-AI'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
