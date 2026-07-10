Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Sets Windows Terminal as the default console host experience.
.DESCRIPTION
  Configures the Windows 11 console delegation registry values so new cmd.exe,
  powershell.exe, pwsh.exe, and other console launches open in Windows Terminal
  instead of the legacy conhost window.

  The setting is per-user. By default, this script applies it to the current
  user, all loaded real user hives, and the default user profile hive so future
  user profiles inherit the same behavior. Non-elevated sessions can still set
  the current user; machine-wide/default-user scopes are skipped when elevation
  is required.
.PARAMETER EnsureInstalled
  Install Windows Terminal through winget when it is not already installed.
.PARAMETER CurrentUser
  Apply the setting to the current user's HKCU hive.
.PARAMETER AllExistingUsers
  Apply the setting to all loaded real user hives under HKEY_USERS.
.PARAMETER DefaultUser
  Mount C:\Users\Default\NTUSER.DAT and apply the setting for future profiles.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER PassThru
  Return structured result objects.
.EXAMPLE
  PS> ./Set-TerminalExperienceDefaults.ps1
  Applies Windows Terminal defaults to current, loaded, and future users where
  permissions allow.
.EXAMPLE
  PS> ./Set-TerminalExperienceDefaults.ps1 -EnsureInstalled -DryRun
  Shows the install and registry actions without changing the system.
.EXAMPLE
  PS> ./Set-TerminalExperienceDefaults.ps1 -CurrentUser
  Applies only to the current user.
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
    HelpMessage = 'Install Windows Terminal through winget when it is not already installed.'
  )]
  [switch]
  $EnsureInstalled,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Apply the setting to the current user HKCU hive.'
  )]
  [switch]
  $CurrentUser,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Apply the setting to all loaded real user hives under HKEY_USERS.'
  )]
  [switch]
  $AllExistingUsers,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Mount C:\Users\Default\NTUSER.DAT and apply the setting for future profiles.'
  )]
  [switch]
  $DefaultUser,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Preview changes without applying them.'
  )]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# -----------------------------------------------------------------------------

if (-not $CurrentUser -and -not $AllExistingUsers -and -not $DefaultUser) {
  $CurrentUser = $true
  $AllExistingUsers = $true
  $DefaultUser = $true
}

# Elevation guard: -AllExistingUsers / -DefaultUser mount user hives (requires admin)
if (($AllExistingUsers -or $DefaultUser) -and -not (Test-Elevation)) {
  Write-Error '-AllExistingUsers and -DefaultUser require administrator privileges. Run elevated or use -CurrentUser only.'
  exit 1
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no terminal defaults will be changed`n" -Color Yellow
}

$terminalConsoleGuid = '{2EACA947-7F5F-4CFA-BA87-8F7FBEEFBE69}'
$terminalTerminalGuid = '{E12CFF52-A866-4C77-9A90-F570A7AA2C6B}'
$results = New-Object System.Collections.ArrayList

function Set-TerminalRegistryValue {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$Scope,
    [string]$Path,
    [string]$Name,
    [object]$Value,
    [string]$Type
  )

  $_target = "$Path\$Name"
  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $_target -Scope $Scope -Action 'SetValue' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_target, "Set terminal default value to $Value")) {
    Add-OperationResult -Results $Results -Target $_target -Scope $Scope -Action 'SetValue' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  try {
    $_result = Set-RegistryValue -Path $Path -Name $Name -Value $Value -Type $Type
    if ($_result) {
      Add-OperationResult -Results $Results -Target $_target -Scope $Scope -Action 'SetValue' -Status $_result.Status -Detail 'Terminal delegation default applied.'
    }
    else {
      Add-OperationResult -Results $Results -Target $_target -Scope $Scope -Action 'SetValue' -Status 'Failed' -Detail 'Set-RegistryValue returned no result.'
    }
  }
  catch {
    Add-OperationResult -Results $Results -Target $_target -Scope $Scope -Action 'SetValue' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Set-TerminalExperienceHive {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$Scope,
    [string]$BasePath
  )

  $_startupKey = "$BasePath\Console\%%Startup"
  Set-TerminalRegistryValue -Results $Results -Scope $Scope -Path $_startupKey -Name 'DelegationConsole' -Value $terminalConsoleGuid -Type String -WhatIf:$WhatIfPreference
  Set-TerminalRegistryValue -Results $Results -Scope $Scope -Path $_startupKey -Name 'DelegationTerminal' -Value $terminalTerminalGuid -Type String -WhatIf:$WhatIfPreference
  Set-TerminalRegistryValue -Results $Results -Scope $Scope -Path "$BasePath\Console" -Name 'ForceV2' -Value 1 -Type DWord -WhatIf:$WhatIfPreference
}

function Install-TerminalExperiencePackage {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results
  )

  $_installed = if (Test-Elevation) {
    @(Get-AppxPackage -AllUsers -Name 'Microsoft.WindowsTerminal' -ErrorAction SilentlyContinue)
  }
  else {
    @(Get-AppxPackage -Name 'Microsoft.WindowsTerminal' -ErrorAction SilentlyContinue)
  }
  if ($_installed.Count -gt 0) {
    Add-OperationResult -Results $Results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Skipped' -Detail 'AlreadyInstalled'
    return
  }

  $_winget = Get-Command -Name winget.exe -ErrorAction SilentlyContinue
  if (-not $_winget) {
    Add-OperationResult -Results $Results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Skipped' -Detail 'WinGetUnavailable'
    return
  }

  if ($DryRun) {
    Add-OperationResult -Results $Results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess('Microsoft.WindowsTerminal', 'Install Windows Terminal with winget')) {
    Add-OperationResult -Results $Results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  $_arguments = @(
    'install',
    '--id',
    'Microsoft.WindowsTerminal',
    '--exact',
    '--silent',
    '--accept-source-agreements',
    '--accept-package-agreements',
    '--scope',
    'machine'
  )

  try {
    $_process = Start-Process -FilePath $_winget.Source -ArgumentList $_arguments -Wait -PassThru -NoNewWindow -ErrorAction Stop
    if ($_process.ExitCode -eq 0) {
      Add-OperationResult -Results $Results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Installed' -Detail 'Windows Terminal installed with winget.'
    }
    else {
      Add-OperationResult -Results $Results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Failed' -Detail "winget exited with $($_process.ExitCode)."
    }
  }
  catch {
    Add-OperationResult -Results $Results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
  }
}

if ($EnsureInstalled) {
  Install-TerminalExperiencePackage -Results $results -WhatIf:$WhatIfPreference
}

if ($CurrentUser) {
  Set-TerminalExperienceHive -Results $results -Scope 'CurrentUser' -BasePath 'HKCU:' -WhatIf:$WhatIfPreference
}

if ($AllExistingUsers) {
  $_userHives = @(Get-ChildItem -LiteralPath 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -match '\\S-1-5-21-\d+-\d+-\d+-\d+$' })

  if ($_userHives.Count -eq 0) {
    Add-OperationResult -Results $results -Target 'Registry::HKEY_USERS' -Scope 'AllExistingUsers' -Action 'SetValue' -Status 'Skipped' -Detail 'NoLoadedUserHives'
  }

  foreach ($_hive in $_userHives) {
    $_sid = Split-Path -Path $_hive.Name -Leaf
    Set-TerminalExperienceHive -Results $results -Scope $_sid -BasePath "Registry::$($_hive.Name)" -WhatIf:$WhatIfPreference
  }
}

if ($DefaultUser) {
  if ($DryRun) {
    Add-OperationResult -Results $results -Target 'Registry::HKEY_USERS\DefaultUser\Console\%%Startup' -Scope 'DefaultUser' -Action 'SetValue' -Status 'Skipped' -Detail 'DryRun'
  }
  elseif (-not $PSCmdlet.ShouldProcess('Registry::HKEY_USERS\DefaultUser', 'Load default user hive and set terminal defaults')) {
    Add-OperationResult -Results $results -Target 'Registry::HKEY_USERS\DefaultUser' -Scope 'DefaultUser' -Action 'SetValue' -Status 'Skipped' -Detail 'WhatIf'
  }
  else {
    $_whatIfBackup = $WhatIfPreference
    $WhatIfPreference = $false
    $_mountResult = Mount-DefaultUserHive -Confirm:$false
    $WhatIfPreference = $_whatIfBackup

    if ($_mountResult) {
      try {
        Set-TerminalExperienceHive -Results $results -Scope 'DefaultUser' -BasePath 'Registry::HKEY_USERS\DefaultUser' -WhatIf:$WhatIfPreference
      }
      finally {
        $_whatIfBackup = $WhatIfPreference
        $WhatIfPreference = $false
        Dismount-DefaultUserHive -Confirm:$false
        $WhatIfPreference = $_whatIfBackup
      }
    }
    else {
      Add-OperationResult -Results $results -Target 'Registry::HKEY_USERS\DefaultUser' -Scope 'DefaultUser' -Action 'MountHive' -Status 'Failed' -Detail 'Default user hive could not be mounted.'
    }
  }
}

$_applied = @($results | Where-Object { $_.Status -in @('Applied', 'Created', 'Updated', 'Installed') }).Count
$_skipped = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_failed = @($results | Where-Object { $_.Status -eq 'Failed' }).Count
Write-Log -Message "Terminal defaults complete. Applied: $_applied | Skipped: $_skipped | Failed: $_failed" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
$_operationLog = Write-OperationResultLog -Results $results -ScriptName 'Set-TerminalExperienceDefaults'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}
Write-Log -Message 'Close and reopen console windows before verifying the new default terminal host.' -Color Gray

if ($PassThru -or $DryRun) {
  $results
}
