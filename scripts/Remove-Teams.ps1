#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Removes Microsoft Teams (classic Win32 and new MSIX-based client) with
  cleanup of cache, startup entries, shortcuts, and reinstall prevention.

.DESCRIPTION
  Port of asheroto/UninstallTeams (v1.2.5), shaped like the sibling
  Remove-OneDrive.ps1: stops Teams processes, uninstalls through the registry
  uninstall keys (MSI GUIDs are re-issued cleanly via msiexec /x), the bundled
  Teams Installer, and Update.exe from both install locations, removes Teams
  Appx packages, deletes application cache folders, removes startup entries
  (including the AutorunsDisabled keys where Task Manager relocates disabled
  entries), removes leftover uninstall registry keys and Desktop/Start Menu
  shortcuts, and optionally applies policy values that prevent Teams from
  silently reinstalling itself.

  The Chat widget (Win+C) is the main silent-reinstall vector: leaving it
  enabled means a user pressing Win+C can bring Teams back. -BlockReinstall
  disables the Chat widget (ChatIcon = Disabled, current user and machine)
  together with the Office-driven install prevention
  (PreventTeamsInstall = 1), mirroring Remove-OneDrive.ps1's -BlockReinstall
  precedent. The individual -DisableChatWidget and -DisableOfficeTeamsInstall
  switches allow granular control.

  Requires administrator elevation (machine-wide Appx removal, HKLM policy
  values, Program Files cleanup). The script removes only application cache
  and configuration data - not user documents - so no -Force gate is needed
  for the folder cleanup.

  Unlike the upstream script, this version does not include a self-update
  mechanism: winkit scripts are versioned and distributed with the repository
  as a whole.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER BlockReinstall
  Apply policy values that prevent Teams from silently reinstalling: disable
  the Chat widget (ChatIcon, current user and machine) and disable the
  Office-driven Teams install (PreventTeamsInstall).

.PARAMETER DisableChatWidget
  Only disable the Chat widget reinstall vector (ChatIcon = Disabled, current
  user and machine).

.PARAMETER DisableOfficeTeamsInstall
  Only disable the Office-driven Teams install vector (PreventTeamsInstall).

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Remove-Teams.ps1 -DryRun
  Previews the removal without changing anything.

.EXAMPLE
  PS> ./Remove-Teams.ps1 -BlockReinstall
  Removes Teams and applies both reinstall-prevention policies.

.EXAMPLE
  PS> ./Remove-Teams.ps1 -DisableChatWidget
  Removes Teams and disables only the Chat widget reinstall vector.

.LINK
  https://github.com/adnoctem/winkit
  https://github.com/asheroto/UninstallTeams

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Preview changes without applying them.'
  )]
  [switch]
  $DryRun,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Apply policy values that prevent Teams from silently reinstalling.'
  )]
  [switch]
  $BlockReinstall,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Disable the Chat widget (Win+C) reinstall vector.'
  )]
  [switch]
  $DisableChatWidget,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Disable the Office-driven Teams install vector.'
  )]
  [switch]
  $DisableOfficeTeamsInstall,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no Teams changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

function Invoke-TeamsUninstaller {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  $_found = $false

  # 1. Registry uninstall keys, accumulated across all display-name variants
  #    (Microsoft Teams / MSTeams / Teams Machine-Wide). The upstream script
  #    overwrites the match list with each call; here all matches are kept.
  $_registryMatches = @(Find-Win32Program -Name '*Teams*' |
      Where-Object { $_.DisplayName -match 'Microsoft Teams|MSTeams|Teams Machine-Wide' -and ($_.UninstallString -or $_.QuietUninstallString) })

  foreach ($_program in $_registryMatches) {
    $_found = $true
    $_uninstallString = if ($_program.QuietUninstallString) { $_program.QuietUninstallString } else { $_program.UninstallString }

    if ($_uninstallString -match 'msiexec\.exe\s*/[XxIi]\{([^\}]+)\}') {
      # MSI case: re-issue a clean msiexec /x {GUID} /qn instead of trusting
      # the stored string's formatting/quoting.
      $_productGuid = $Matches[1]
      if ($PSCmdlet.ShouldProcess($_program.DisplayName, "Run msiexec /x {$_productGuid} /qn")) {
        try {
          $_process = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x {$_productGuid} /qn" -Wait -PassThru -ErrorAction Stop
          Add-OperationResult -Results $Results -Target $_program.DisplayName -Action 'Uninstall' -Status "ExitCode:$($_process.ExitCode)" -Detail 'Teams MSI uninstall executed.'
        }
        catch {
          Add-OperationResult -Results $Results -Target $_program.DisplayName -Action 'Uninstall' -Status 'Failed' -Detail $_.Exception.Message
        }
      }
      else {
        Add-OperationResult -Results $Results -Target $_program.DisplayName -Action 'Uninstall' -Status 'Skipped' -Detail 'WhatIf'
      }
    }
    else {
      # Non-MSI case: delegate to PSFoundation, which handles quoting correctly.
      $_uninstallResult = Uninstall-Win32Program -InputObject $_program -Quiet -Force -DryRun:$DryRun -WhatIf:$WhatIfPreference
      Add-OperationResult -Results $Results -Target $_program.DisplayName -Action 'UninstallRegistryEntry' -Status $_uninstallResult.Status -Detail $_uninstallResult.Error
    }
  }

  # 2. Bundled Teams Installer.
  $_teamsInstaller = Join-Path ${env:ProgramFiles(x86)} 'Teams Installer\Teams.exe'
  if (Test-Path -LiteralPath $_teamsInstaller) {
    $_found = $true
    if ($PSCmdlet.ShouldProcess($_teamsInstaller, 'Run Teams.exe --uninstall')) {
      try {
        $_process = Start-Process -FilePath $_teamsInstaller -ArgumentList '--uninstall' -Wait -PassThru -ErrorAction Stop
        Add-OperationResult -Results $Results -Target $_teamsInstaller -Action 'Uninstall' -Status "ExitCode:$($_process.ExitCode)" -Detail 'Teams Installer uninstall executed.'
      }
      catch {
        Add-OperationResult -Results $Results -Target $_teamsInstaller -Action 'Uninstall' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
    else {
      Add-OperationResult -Results $Results -Target $_teamsInstaller -Action 'Uninstall' -Status 'Skipped' -Detail 'WhatIf'
    }
  }

  # 3. Update.exe from both the per-user and machine-wide install locations.
  $_updatePaths = @(
    (Join-Path -Path $env:APPDATA -ChildPath 'Microsoft\Teams\Update.exe'),
    (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath 'Microsoft\Teams\current\Update.exe')
  ) | Where-Object { Test-Path -LiteralPath $_ }

  foreach ($_path in $_updatePaths) {
    $_found = $true
    if ($PSCmdlet.ShouldProcess($_path, 'Run Update.exe -uninstall -s')) {
      try {
        $_process = Start-Process -FilePath $_path -ArgumentList '-uninstall -s' -Wait -PassThru -ErrorAction Stop
        Add-OperationResult -Results $Results -Target $_path -Action 'Uninstall' -Status "ExitCode:$($_process.ExitCode)" -Detail 'Teams Update.exe uninstall executed.'
      }
      catch {
        Add-OperationResult -Results $Results -Target $_path -Action 'Uninstall' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
    else {
      Add-OperationResult -Results $Results -Target $_path -Action 'Uninstall' -Status 'Skipped' -Detail 'WhatIf'
    }
  }

  if (-not $_found) {
    Add-OperationResult -Results $Results -Target 'Teams' -Action 'Uninstall' -Status 'Skipped' -Detail 'No Teams uninstaller was found.'
  }
}

function Remove-TeamsAppxPackage {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  $_removals = @(
    @{ Pattern = 'Microsoft Teams*'; AllUsers = $false },
    @{ Pattern = 'Microsoft Teams*'; AllUsers = $true },
    @{ Pattern = 'MSTeams*'; AllUsers = $true }
  )

  foreach ($_removal in $_removals) {
    $_packages = @(if ($_removal.AllUsers) { Get-AppxPackage -Name $_removal.Pattern -AllUsers -ErrorAction SilentlyContinue } else { Get-AppxPackage -Name $_removal.Pattern -ErrorAction SilentlyContinue })
    foreach ($_package in $_packages) {
      $_target = "$($_package.Name) ($($_package.PackageFullName))"
      if ($PSCmdlet.ShouldProcess($_target, 'Remove Teams Appx package')) {
        try {
          if ($_removal.AllUsers) {
            Remove-AppxPackage -Package $_package.PackageFullName -AllUsers -ErrorAction Stop
          }
          else {
            Remove-AppxPackage -Package $_package.PackageFullName -ErrorAction Stop
          }
          Add-OperationResult -Results $Results -Target $_target -Action 'RemoveAppxPackage' -Status 'Removed' -Detail 'Teams Appx package removed.'
        }
        catch {
          Add-OperationResult -Results $Results -Target $_target -Action 'RemoveAppxPackage' -Status 'Failed' -Detail $_.Exception.Message
        }
      }
      else {
        Add-OperationResult -Results $Results -Target $_target -Action 'RemoveAppxPackage' -Status 'Skipped' -Detail 'WhatIf'
      }
    }
  }
}

function Remove-TeamsCacheFolder {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  $_folders = @(
    (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft Teams'),
    (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\Teams'),
    (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\TeamsMeetingAddin'),
    (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\TeamsPresenceAddin')
  ) | Where-Object { Test-Path -LiteralPath $_ }

  foreach ($_folder in $_folders) {
    if ($PSCmdlet.ShouldProcess($_folder, 'Remove Teams cache folder')) {
      try {
        if (-not $DryRun) {
          Remove-Item -LiteralPath $_folder -Recurse -Force -ErrorAction Stop
        }
        Add-OperationResult -Results $Results -Target $_folder -Action 'RemoveCacheFolder' -Status 'Removed' -Detail 'Teams cache folder removed.'
      }
      catch {
        Add-OperationResult -Results $Results -Target $_folder -Action 'RemoveCacheFolder' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
    else {
      Add-OperationResult -Results $Results -Target $_folder -Action 'RemoveCacheFolder' -Status 'Skipped' -Detail 'WhatIf'
    }
  }
}

function Remove-TeamsStartupEntry {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper is guarded by the parent script ShouldProcess/DryRun flow.')]
  [CmdletBinding()]
  param([System.Collections.ArrayList]$Results)

  $_runKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled',
    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run\AutorunsDisabled'
  )
  $_names = @('Teams', 'TeamsMachineUninstallerLocalAppData', 'TeamsMachineUninstallerProgramData', 'com.squirrel.Teams.Teams', 'TeamsMachineInstaller')

  foreach ($_key in $_runKeys) {
    foreach ($_name in $_names) {
      if ($DryRun) {
        Add-OperationResult -Results $Results -Target "$_key\$_name" -Action 'RemoveStartupEntry' -Status 'Skipped' -Detail 'DryRun'
        continue
      }

      $_result = Remove-RegistryValue -Path $_key -Name $_name
      $_status = if ($_result) { $_result.Status } else { 'Failed' }
      Add-OperationResult -Results $Results -Target "$_key\$_name" -Action 'RemoveStartupEntry' -Status $_status -Detail 'Teams startup entry removed when present.'
    }
  }
}

function Remove-TeamsUninstallRegistryKey {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  $_uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Teams',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams',
    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Teams'
  ) | Where-Object { Test-Path -LiteralPath $_ }

  foreach ($_path in $_uninstallPaths) {
    if ($PSCmdlet.ShouldProcess($_path, 'Remove Teams uninstall registry key')) {
      try {
        if (-not $DryRun) {
          Remove-Item -LiteralPath $_path -Recurse -Force -ErrorAction Stop
        }
        Add-OperationResult -Results $Results -Target $_path -Action 'RemoveUninstallRegistryKey' -Status 'Removed' -Detail 'Leftover Teams uninstall registry key removed.'
      }
      catch {
        Add-OperationResult -Results $Results -Target $_path -Action 'RemoveUninstallRegistryKey' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
    else {
      Add-OperationResult -Results $Results -Target $_path -Action 'RemoveUninstallRegistryKey' -Status 'Skipped' -Detail 'WhatIf'
    }
  }
}

function Remove-TeamsShortcut {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  foreach ($_name in @('Microsoft Teams', 'Microsoft Teams classic (work or school)')) {
    foreach ($_location in @('Desktop', 'StartMenu')) {
      $_userPath = [Environment]::GetFolderPath($_location)
      $_publicPath = [Environment]::GetFolderPath("Common$_location")

      foreach ($_base in @($_userPath, $_publicPath)) {
        $_shortcut = Join-Path -Path $_base -ChildPath "$_name.lnk"
        if (-not (Test-Path -LiteralPath $_shortcut)) {
          continue
        }

        if ($PSCmdlet.ShouldProcess($_shortcut, 'Remove Teams shortcut')) {
          try {
            if (-not $DryRun) {
              Remove-Item -LiteralPath $_shortcut -Force -ErrorAction Stop
            }
            Add-OperationResult -Results $Results -Target $_shortcut -Action 'RemoveShortcut' -Status 'Removed' -Detail 'Teams shortcut removed.'
          }
          catch {
            Add-OperationResult -Results $Results -Target $_shortcut -Action 'RemoveShortcut' -Status 'Failed' -Detail $_.Exception.Message
          }
        }
        else {
          Add-OperationResult -Results $Results -Target $_shortcut -Action 'RemoveShortcut' -Status 'Skipped' -Detail 'WhatIf'
        }
      }
    }
  }
}

function Set-TeamsPolicy {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper is guarded by the parent script ShouldProcess/DryRun flow.')]
  [CmdletBinding()]
  param(
    [System.Collections.ArrayList]$Results,
    [bool]$ChatWidget,
    [bool]$OfficeTeamsInstall
  )

  $_settings = @()
  if ($ChatWidget) {
    $_settings += @(
      @{
        Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat'
        Name = 'ChatIcon'
        Value = 3
        Type = 'DWord'
      },
      @{
        Path = 'HKCU:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat'
        Name = 'ChatIcon'
        Value = 3
        Type = 'DWord'
      }
    )
  }
  if ($OfficeTeamsInstall) {
    $_settings += @(
      @{
        Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate'
        Name = 'PreventTeamsInstall'
        Value = 1
        Type = 'DWord'
      }
    )
  }

  foreach ($_setting in $_settings) {
    if ($DryRun) {
      Add-OperationResult -Results $Results -Target "$($_setting.Path)\$($_setting.Name)" -Action 'SetPolicy' -Status 'Skipped' -Detail 'DryRun'
      continue
    }

    $_result = Set-RegistryValue -Path $_setting.Path -Name $_setting.Name -Value $_setting.Value -Type $_setting.Type
    $_status = if ($_result) { $_result.Status } else { 'Failed' }
    Add-OperationResult -Results $Results -Target "$($_setting.Path)\$($_setting.Name)" -Action 'SetPolicy' -Status $_status -Detail 'Teams reinstall prevention policy.'
  }
}

function Test-TeamsReinstallStatus {
  [CmdletBinding()]
  param([System.Collections.ArrayList]$Results)

  $_chatPathMachine = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat'
  $_chatPathUser = 'HKCU:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat'
  $_officePath = 'HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate'

  $_chatIconUser = (Get-ItemProperty -Path $_chatPathUser -Name 'ChatIcon' -ErrorAction SilentlyContinue).ChatIcon
  $_chatIconMachine = (Get-ItemProperty -Path $_chatPathMachine -Name 'ChatIcon' -ErrorAction SilentlyContinue).ChatIcon
  $_preventTeamsInstall = (Get-ItemProperty -Path $_officePath -Name 'PreventTeamsInstall' -ErrorAction SilentlyContinue).PreventTeamsInstall

  # Effective state: current-user overrides machine-wide unless unset.
  $_chatEffective = if ($null -ne $_chatIconUser) { $_chatIconUser } elseif ($null -ne $_chatIconMachine) { $_chatIconMachine } else { $null }
  $_chatBlocked = ($_chatEffective -eq 3)
  $_officeBlocked = ($_preventTeamsInstall -eq 1)

  Add-OperationResult -Results $Results -Target 'ChatWidget' -Source 'Teams' -Action 'Verify' -Status $(if ($_chatBlocked) { 'Completed' } else { 'Warn' }) -Detail "Effective ChatIcon: $(if ($null -ne $_chatEffective) { $_chatEffective } else { 'Unset' }) (3 = disabled; unset means the default applies - which has varied across Windows builds)."
  Add-OperationResult -Results $Results -Target 'OfficeTeamsInstall' -Source 'Teams' -Action 'Verify' -Status $(if ($_officeBlocked) { 'Completed' } else { 'Warn' }) -Detail "PreventTeamsInstall: $(if ($null -ne $_preventTeamsInstall) { $_preventTeamsInstall } else { 'Unset' }) (1 = disabled)."

  if (-not $_chatBlocked -or -not $_officeBlocked) {
    Write-Log -Message 'Warning: Teams can still be reinstalled. Use -BlockReinstall to disable the Chat widget and the Office-driven Teams install.' -Color Yellow
  }
}

Write-Log -Message 'Stopping Teams processes...' -Color Yellow
$_processes = @(Get-Process -Name 'Microsoft Teams*', 'Teams Machine-Wide*', 'MSTeams*', 'ms-teams*', 'Teams*' -ErrorAction SilentlyContinue | Sort-Object Id -Unique)
foreach ($_process in $_processes) {
  if ($PSCmdlet.ShouldProcess($_process.ProcessName, 'Stop Teams process')) {
    try {
      Stop-Process -Id $_process.Id -Force -ErrorAction Stop
      Add-OperationResult -Results $_results -Target $_process.ProcessName -Action 'StopProcess' -Status 'Stopped' -Detail "PID $($_process.Id)"
    }
    catch {
      Add-OperationResult -Results $_results -Target $_process.ProcessName -Action 'StopProcess' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

Write-Log -Message 'Running Teams uninstallers...' -Color Yellow
Invoke-TeamsUninstaller -Results $_results

Write-Log -Message 'Removing Teams Appx packages...' -Color Yellow
Remove-TeamsAppxPackage -Results $_results

Write-Log -Message 'Removing Teams cache folders...' -Color Yellow
Remove-TeamsCacheFolder -Results $_results

Write-Log -Message 'Removing Teams startup entries...' -Color Yellow
Remove-TeamsStartupEntry -Results $_results

Write-Log -Message 'Removing leftover Teams uninstall registry keys...' -Color Yellow
Remove-TeamsUninstallRegistryKey -Results $_results

Write-Log -Message 'Removing Teams shortcuts...' -Color Yellow
Remove-TeamsShortcut -Results $_results

if ($BlockReinstall -or $DisableChatWidget -or $DisableOfficeTeamsInstall) {
  Write-Log -Message 'Applying Teams reinstall prevention policies...' -Color Yellow
  Set-TeamsPolicy -Results $_results -ChatWidget ($BlockReinstall -or $DisableChatWidget) -OfficeTeamsInstall ($BlockReinstall -or $DisableOfficeTeamsInstall)
}

Write-Log -Message 'Checking Teams reinstall status...' -Color Yellow
Test-TeamsReinstallStatus -Results $_results

$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_skipped = @($_results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_completed = @($_results | Where-Object { $_.Status -notin @('Failed', 'Skipped') }).Count
$_color = if ($_failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Log -Message "`nTeams removal workflow complete. Completed: $_completed | Skipped: $_skipped | Failed: $_failed" -Color $_color
$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Remove-Teams'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
