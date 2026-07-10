Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Sets default Windows app permission consent values.
.DESCRIPTION
  Applies default ConsentStore values for sensitive app capabilities such as
  camera, microphone, location, contacts, calendar, radios, and broad file
  system access. This is intentionally separate from Configure-Privacy.ps1
  because these defaults are more restrictive and can affect app workflows.
  Both machine-default and current/default user consent locations are managed.
.PARAMETER Undo
  Remove permission default values managed by this script.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER SysPrep
  Mount the default user profile hive and write HKCU-backed permission defaults
  there instead of HKCU. HKLM ConsentStore defaults still target the live machine.
.PARAMETER Config
  JSON file containing setting overrides. Entries match built-in settings by
  Name and Path and can override Preferred or Default values.
.PARAMETER ExportConfig
  Export the default permission settings JSON and exit.
.PARAMETER ExportCurrentState
  Export current registry values as reusable JSON config and exit.
.PARAMETER ExportPath
  File path used with -ExportConfig.
.EXAMPLE
  PS> ./Set-AppPermissionDefaults.ps1
  Applies the default app permission-deny profile.
.EXAMPLE
  PS> ./Set-AppPermissionDefaults.ps1 -DryRun
  Shows every ConsentStore value that would be written.
.EXAMPLE
  PS> ./Set-AppPermissionDefaults.ps1 -Undo
  Removes the permission default values managed by this script.
.EXAMPLE
  PS> ./Set-AppPermissionDefaults.ps1 -SysPrep
  Writes HKCU-backed permission defaults into the default user profile hive.
.EXAMPLE
  PS> ./Set-AppPermissionDefaults.ps1 -ExportConfig -ExportPath '.\app-permissions.json'
  Exports the default permission profile to JSON.
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
    HelpMessage = 'Remove permission default values managed by this script.'
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
    HelpMessage = 'Mount the default user profile hive and write HKCU-backed permission defaults there.'
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
    HelpMessage = 'Export the default permission settings JSON and exit.'
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

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$regHive = if ($SysPrep) { 'Registry::HKEY_USERS\DefaultUser' } else { 'HKCU:' }
$userConsentRoot = "$regHive\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore"
$machineConsentRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore'

$capabilities = @(
  @{
    Capability = 'webcam'
    Group = 'Camera'
    Description = 'camera access'
  }
  @{
    Capability = 'microphone'
    Group = 'Microphone'
    Description = 'microphone access'
  }
  @{
    Capability = 'location'
    Group = 'Location'
    Description = 'location access'
  }
  @{
    Capability = 'contacts'
    Group = 'Contacts'
    Description = 'contacts access'
  }
  @{
    Capability = 'appointments'
    Group = 'Calendar'
    Description = 'calendar access'
  }
  @{
    Capability = 'phoneCallHistory'
    Group = 'CallHistory'
    Description = 'call history access'
  }
  @{
    Capability = 'email'
    Group = 'Email'
    Description = 'email access'
  }
  @{
    Capability = 'radios'
    Group = 'Radios'
    Description = 'radio control access'
  }
  @{
    Capability = 'documentsLibrary'
    Group = 'DocumentsLibrary'
    Description = 'Documents library access'
  }
  @{
    Capability = 'picturesLibrary'
    Group = 'PicturesLibrary'
    Description = 'Pictures library access'
  }
  @{
    Capability = 'videosLibrary'
    Group = 'VideosLibrary'
    Description = 'Videos library access'
  }
  @{
    Capability = 'broadFileSystemAccess'
    Group = 'BroadFileSystem'
    Description = 'broad file system access'
  }
)

$permissionSettings = foreach ($capability in $capabilities) {
  @{
    Path = "$machineConsentRoot\$($capability.Capability)"
    Name = 'Value'
    Preferred = 'Deny'
    Default = $null
    Type = 'String'
    Group = $capability.Group
    Description = "Deny machine-default $($capability.Description)."
  }
  @{
    Path = "$userConsentRoot\$($capability.Capability)"
    Name = 'Value'
    Preferred = 'Deny'
    Default = $null
    Type = 'String'
    Group = $capability.Group
    Description = "Deny current/default user $($capability.Description)."
  }
}

if ($ExportCurrentState) {
  if ($DryRun) { Write-Log -Message '-DryRun cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($ExportConfig) { Write-Log -Message '-ExportConfig cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  if ($Undo) { Write-Log -Message '-Undo cannot be combined with -ExportCurrentState.' -Color Red; exit 1 }
  $_currentState = Export-RegistrySettingState -Settings $permissionSettings
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_currentState | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current app permission settings exported to: $_exportPath" -Color Green
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
    $permissionSettings | ConvertTo-Json -Depth 3 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default app permission settings exported to: $_exportPath" -Color Green
  }
  else {
    $permissionSettings | ConvertTo-Json -Depth 3
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
  Merge-ObjectArrays -Base $permissionSettings -Overrides $_overrides
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

foreach ($entry in $permissionSettings) {
  $targetValue = if ($Undo) { $entry.Default } else { $entry.Preferred }
  if ($Undo -and $null -eq $entry.Default) {
    Write-Log -Message "$targetLabel app permission default: Remove '$($entry.Name)' - $($entry.Description)" -Color Yellow
    if ($DryRun) {
      Write-Log -Message "  -> Would remove $($entry.Path)\$($entry.Name)" -Color Gray
      continue
    }
    $result = Remove-RegistryValue -Path $entry.Path -Name $entry.Name
  }
  else {
    Write-Log -Message "$targetLabel app permission default: $($entry.Name) = '$targetValue' - $($entry.Description)" -Color Yellow
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
  Write-Log -Message "`nApp permission defaults have been processed." -Color Green
}
else {
  Write-Log -Message "`nAll registry values were already at the desired target - nothing to do." -Color Green
}
$_operationResults = @(ConvertTo-RegistrySettingResult -Settings $permissionSettings -Undo:$Undo -DryRun:$DryRun)
$_operationLog = Write-OperationResultLog -Results $_operationResults -ScriptName 'Set-AppPermissionDefaults'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_operationResults
}
