#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Removes Microsoft OneDrive with guarded cleanup steps.
.DESCRIPTION
  Stops OneDrive processes, runs known OneDrive uninstallers, applies policy and
  Explorer sidebar cleanup, removes startup entries, and optionally removes
  OneDrive scheduled tasks or migrates known folders back to the local profile.
  The script does not reboot, does not require Safe Mode, and does not delete
  user data unless explicit switches are supplied.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER Instant
  Restart Explorer after applying shell visibility changes.
.PARAMETER BackupPath
  Destination used when -MigrateKnownFolders backs up files before moving them.
.PARAMETER NoBackup
  Skip backup creation during -MigrateKnownFolders.
.PARAMETER MigrateKnownFolders
  Move Desktop, Documents, and Pictures content out of the OneDrive folder and
  restore known folder registry values to local profile paths. Cloud-only files
  are skipped to avoid data loss.
.PARAMETER RemoveScheduledTasks
  Remove scheduled tasks whose task name starts with OneDrive.
.PARAMETER RemoveUserFolder
  Remove the OneDrive folder only when it is empty.
.PARAMETER BlockReinstall
  Apply OneDrive policy values that prevent file sync and consumer reinstall
  behavior.
.PARAMETER Force
  Continue optional operations that normally require explicit confirmation.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> ./Remove-OneDrive.ps1 -DryRun
.EXAMPLE
  PS> ./Remove-OneDrive.ps1 -BlockReinstall -RemoveScheduledTasks -Instant
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'BackupPath', Justification = 'Used by nested known-folder migration helper through script scope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'NoBackup', Justification = 'Used by nested known-folder migration helper through script scope.')]
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
    HelpMessage = 'Restart Explorer after applying shell visibility changes.'
  )]
  [switch]
  $Instant,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Destination used when -MigrateKnownFolders backs up files before moving them.'
  )]
  [string]
  $BackupPath = (Join-Path -Path $env:USERPROFILE -ChildPath ('OneDrive-Backup-{0:yyyyMMdd-HHmmss}' -f (Get-Date))),

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Skip backup creation during -MigrateKnownFolders.'
  )]
  [switch]
  $NoBackup,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Move Desktop, Documents, and Pictures content out of the OneDrive folder.'
  )]
  [switch]
  $MigrateKnownFolders,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Remove scheduled tasks whose task name starts with OneDrive.'
  )]
  [switch]
  $RemoveScheduledTasks,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Remove the OneDrive folder only when it is empty.'
  )]
  [switch]
  $RemoveUserFolder,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Apply OneDrive policy values that prevent file sync and consumer reinstall.'
  )]
  [switch]
  $BlockReinstall,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Continue optional operations that normally require explicit confirmation.'
  )]
  [switch]
  $Force,

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
  Write-Log -Message "DRY RUN - no OneDrive changes will be applied`n" -Color Yellow
}

function Get-OneDrivePath {
  [CmdletBinding()]
  param()

  $_paths = @(
    $env:OneDrive,
    $env:OneDriveConsumer,
    $env:OneDriveCommercial,
    (Join-Path -Path $env:USERPROFILE -ChildPath 'OneDrive'),
    (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\OneDrive'),
    (Join-Path -Path $env:ProgramData -ChildPath 'Microsoft OneDrive')
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

  $_paths | Sort-Object -Unique | Where-Object { Test-Path -LiteralPath $_ }
}

function Test-OneDriveCloudOnlyFile {
  [CmdletBinding()]
  param([Parameter(Mandatory = $true)][string]$Path)

  try {
    $_item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    return (($_item.Attributes -band [System.IO.FileAttributes]::Offline) -ne 0)
  }
  catch {
    return $true
  }
}

function Invoke-OneDriveUninstaller {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  $_setupPaths = @(
    (Join-Path -Path $env:SystemRoot -ChildPath 'System32\OneDriveSetup.exe'),
    (Join-Path -Path $env:SystemRoot -ChildPath 'SysWOW64\OneDriveSetup.exe'),
    (Join-Path -Path $env:ProgramFiles -ChildPath 'Microsoft Office\root\Integration\Addons\OneDriveSetup.exe')
  )

  if (${env:ProgramFiles(x86)}) {
    $_setupPaths += (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath 'Microsoft Office\root\Integration\Addons\OneDriveSetup.exe')
  }

  $_setupPaths = @($_setupPaths | Sort-Object -Unique | Where-Object { Test-Path -LiteralPath $_ })

  foreach ($_path in $_setupPaths) {
    if ($PSCmdlet.ShouldProcess($_path, 'Run OneDriveSetup.exe /uninstall')) {
      try {
        $_process = Start-Process -FilePath $_path -ArgumentList '/uninstall' -Wait -PassThru -ErrorAction Stop
        Add-OperationResult -Results $Results -Target $_path -Action 'Uninstall' -Status "ExitCode:$($_process.ExitCode)" -Detail 'OneDrive setup uninstaller executed.'
      }
      catch {
        Add-OperationResult -Results $Results -Target $_path -Action 'Uninstall' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
    else {
      Add-OperationResult -Results $Results -Target $_path -Action 'Uninstall' -Status 'Skipped' -Detail 'WhatIf'
    }
  }

  $_registryMatches = @(Find-Win32Program -Name '*OneDrive*' -IncludeSystemComponent |
      Where-Object { $_.UninstallString -or $_.QuietUninstallString })

  foreach ($_program in $_registryMatches) {
    $_uninstallResult = Uninstall-Win32Program -InputObject $_program -Quiet -Force -DryRun:$DryRun -WhatIf:$WhatIfPreference
    Add-OperationResult -Results $Results -Target $_program.DisplayName -Action 'UninstallRegistryEntry' -Status $_uninstallResult.Status -Detail $_uninstallResult.Error
  }

  if ($_setupPaths.Count -eq 0 -and $_registryMatches.Count -eq 0) {
    Add-OperationResult -Results $Results -Target 'OneDrive' -Action 'Uninstall' -Status 'Skipped' -Detail 'No OneDrive uninstaller was found.'
  }
}

function Set-OneDrivePolicy {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper is guarded by the parent script ShouldProcess/DryRun flow.')]
  [CmdletBinding()]
  param([System.Collections.ArrayList]$Results)

  $_settings = @(
    @{
      Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive'
      Name = 'DisableFileSync'
      Value = 1
      Type = 'DWord'
    },
    @{
      Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive'
      Name = 'DisableFileSyncNGSC'
      Value = 1
      Type = 'DWord'
    },
    @{
      Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent'
      Name = 'DisableWindowsConsumerFeatures'
      Value = 1
      Type = 'DWord'
    }
  )

  foreach ($_setting in $_settings) {
    if ($DryRun) {
      Add-OperationResult -Results $Results -Target "$($_setting.Path)\$($_setting.Name)" -Action 'SetPolicy' -Status 'Skipped' -Detail 'DryRun'
      continue
    }

    $_result = Set-RegistryValue -Path $_setting.Path -Name $_setting.Name -Value $_setting.Value -Type $_setting.Type
    $_status = if ($_result) { $_result.Status } else { 'Failed' }
    Add-OperationResult -Results $Results -Target "$($_setting.Path)\$($_setting.Name)" -Action 'SetPolicy' -Status $_status -Detail 'OneDrive reinstall/sync prevention policy.'
  }
}

function Hide-OneDriveFromExplorer {
  [CmdletBinding()]
  param([System.Collections.ArrayList]$Results)

  $_keys = @(
    'HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}',
    'HKLM:\SOFTWARE\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}',
    'HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
  )

  foreach ($_key in $_keys) {
    if ($DryRun) {
      Add-OperationResult -Results $Results -Target "$_key\System.IsPinnedToNameSpaceTree" -Action 'HideExplorerEntry' -Status 'Skipped' -Detail 'DryRun'
      continue
    }

    $_result = Set-RegistryValue -Path $_key -Name 'System.IsPinnedToNameSpaceTree' -Value 0 -Type DWord
    $_status = if ($_result) { $_result.Status } else { 'Failed' }
    Add-OperationResult -Results $Results -Target "$_key\System.IsPinnedToNameSpaceTree" -Action 'HideExplorerEntry' -Status $_status -Detail 'OneDrive navigation pane entry hidden.'
  }
}

function Remove-OneDriveStartupEntry {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper is guarded by the parent script ShouldProcess/DryRun flow.')]
  [CmdletBinding()]
  param([System.Collections.ArrayList]$Results)

  $_runKeys = @(
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
  )
  $_names = @('OneDrive', 'OneDriveSetup')

  foreach ($_key in $_runKeys) {
    foreach ($_name in $_names) {
      if ($DryRun) {
        Add-OperationResult -Results $Results -Target "$_key\$_name" -Action 'RemoveStartupEntry' -Status 'Skipped' -Detail 'DryRun'
        continue
      }

      $_result = Remove-RegistryValue -Path $_key -Name $_name
      $_status = if ($_result) { $_result.Status } else { 'Failed' }
      Add-OperationResult -Results $Results -Target "$_key\$_name" -Action 'RemoveStartupEntry' -Status $_status -Detail 'OneDrive startup entry removed when present.'
    }
  }
}

function Move-OneDriveKnownFolder {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [System.Collections.ArrayList]$Results,
    [string]$OneDriveRoot,
    [string]$FolderName,
    [string]$Destination
  )

  $_source = Join-Path -Path $OneDriveRoot -ChildPath $FolderName
  if (-not (Test-Path -LiteralPath $_source)) {
    Add-OperationResult -Results $Results -Target $_source -Action 'MigrateKnownFolder' -Status 'Skipped' -Detail 'Folder not found.'
    return
  }

  if (-not $NoBackup) {
    $_backupDestination = Join-Path -Path $BackupPath -ChildPath $FolderName
    if ($PSCmdlet.ShouldProcess($_source, "Back up to $_backupDestination")) {
      try {
        if (-not $DryRun -and -not (Test-Path -LiteralPath $BackupPath)) {
          $null = New-Item -Path $BackupPath -ItemType Directory -Force -ErrorAction Stop
        }
        if (-not $DryRun) {
          Copy-Item -LiteralPath $_source -Destination $_backupDestination -Recurse -Force -ErrorAction Stop
        }
        Add-OperationResult -Results $Results -Target $_source -Action 'BackupKnownFolder' -Status 'Completed' -Detail $_backupDestination
      }
      catch {
        Add-OperationResult -Results $Results -Target $_source -Action 'BackupKnownFolder' -Status 'Failed' -Detail $_.Exception.Message
        return
      }
    }
  }

  if (-not (Test-Path -LiteralPath $Destination) -and -not $DryRun) {
    $null = New-Item -Path $Destination -ItemType Directory -Force -ErrorAction Stop
  }

  $_items = @(Get-ChildItem -LiteralPath $_source -Force -ErrorAction SilentlyContinue)
  foreach ($_item in $_items) {
    if (Test-OneDriveCloudOnlyFile -Path $_item.FullName) {
      Add-OperationResult -Results $Results -Target $_item.FullName -Action 'MigrateKnownFolder' -Status 'Skipped' -Detail 'Cloud-only or inaccessible item.'
      continue
    }

    $_target = Join-Path -Path $Destination -ChildPath $_item.Name
    if ($PSCmdlet.ShouldProcess($_item.FullName, "Move to $_target")) {
      try {
        if (-not $DryRun) {
          Move-Item -LiteralPath $_item.FullName -Destination $_target -Force -ErrorAction Stop
        }
        Add-OperationResult -Results $Results -Target $_item.FullName -Action 'MigrateKnownFolder' -Status 'Moved' -Detail $_target
      }
      catch {
        Add-OperationResult -Results $Results -Target $_item.FullName -Action 'MigrateKnownFolder' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
  }
}

function Restore-KnownFolderRegistry {
  [CmdletBinding()]
  param([System.Collections.ArrayList]$Results)

  $_userShellFolders = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
  $_shellFolders = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
  $_settings = @(
    @{
      Path = $_userShellFolders
      Name = 'Desktop'
      Value = '%USERPROFILE%\Desktop'
      Type = 'ExpandString'
    },
    @{
      Path = $_userShellFolders
      Name = 'Personal'
      Value = '%USERPROFILE%\Documents'
      Type = 'ExpandString'
    },
    @{
      Path = $_userShellFolders
      Name = 'My Pictures'
      Value = '%USERPROFILE%\Pictures'
      Type = 'ExpandString'
    },
    @{
      Path = $_shellFolders
      Name = 'Desktop'
      Value = (Join-Path -Path $env:USERPROFILE -ChildPath 'Desktop')
      Type = 'String'
    },
    @{
      Path = $_shellFolders
      Name = 'Personal'
      Value = (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents')
      Type = 'String'
    },
    @{
      Path = $_shellFolders
      Name = 'My Pictures'
      Value = (Join-Path -Path $env:USERPROFILE -ChildPath 'Pictures')
      Type = 'String'
    }
  )

  foreach ($_setting in $_settings) {
    if ($DryRun) {
      Add-OperationResult -Results $Results -Target "$($_setting.Path)\$($_setting.Name)" -Action 'RestoreKnownFolderRegistry' -Status 'Skipped' -Detail 'DryRun'
      continue
    }
    $_result = Set-RegistryValue -Path $_setting.Path -Name $_setting.Name -Value $_setting.Value -Type $_setting.Type
    $_status = if ($_result) { $_result.Status } else { 'Failed' }
    Add-OperationResult -Results $Results -Target "$($_setting.Path)\$($_setting.Name)" -Action 'RestoreKnownFolderRegistry' -Status $_status -Detail $_setting.Value
  }
}

$_results = New-Object System.Collections.ArrayList

Write-Log -Message 'Stopping OneDrive processes...' -Color Yellow
$_processes = @(Get-Process -Name 'OneDrive', 'OneDriveStandaloneUpdater', 'FileCoAuth' -ErrorAction SilentlyContinue)
foreach ($_process in $_processes) {
  if ($PSCmdlet.ShouldProcess($_process.ProcessName, 'Stop OneDrive process')) {
    try {
      Stop-Process -Id $_process.Id -Force -ErrorAction Stop
      Add-OperationResult -Results $_results -Target $_process.ProcessName -Action 'StopProcess' -Status 'Stopped' -Detail "PID $($_process.Id)"
    }
    catch {
      Add-OperationResult -Results $_results -Target $_process.ProcessName -Action 'StopProcess' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

Write-Log -Message 'Running OneDrive uninstallers...' -Color Yellow
Invoke-OneDriveUninstaller -Results $_results -WhatIf:$WhatIfPreference

if ($BlockReinstall) {
  Write-Log -Message 'Applying OneDrive reinstall/sync prevention policies...' -Color Yellow
  Set-OneDrivePolicy -Results $_results
}

Write-Log -Message 'Hiding OneDrive from Explorer and removing startup entries...' -Color Yellow
Hide-OneDriveFromExplorer -Results $_results
Remove-OneDriveStartupEntry -Results $_results

if ($MigrateKnownFolders) {
  $_oneDriveRoots = @(Get-OneDrivePath | Where-Object { $_ -like "$env:USERPROFILE*" })
  foreach ($_root in $_oneDriveRoots) {
    Move-OneDriveKnownFolder -Results $_results -OneDriveRoot $_root -FolderName 'Desktop' -Destination (Join-Path -Path $env:USERPROFILE -ChildPath 'Desktop') -WhatIf:$WhatIfPreference
    Move-OneDriveKnownFolder -Results $_results -OneDriveRoot $_root -FolderName 'Documents' -Destination (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents') -WhatIf:$WhatIfPreference
    Move-OneDriveKnownFolder -Results $_results -OneDriveRoot $_root -FolderName 'Pictures' -Destination (Join-Path -Path $env:USERPROFILE -ChildPath 'Pictures') -WhatIf:$WhatIfPreference
  }
  Restore-KnownFolderRegistry -Results $_results
}

if ($RemoveScheduledTasks) {
  Write-Log -Message 'Removing OneDrive scheduled tasks...' -Color Yellow
  $_tasks = @(Get-ScheduledTask -TaskName 'OneDrive*' -ErrorAction SilentlyContinue)
  foreach ($_task in $_tasks) {
    if ($PSCmdlet.ShouldProcess($_task.TaskName, 'Unregister OneDrive scheduled task')) {
      try {
        if (-not $DryRun) {
          Unregister-ScheduledTask -TaskName $_task.TaskName -TaskPath $_task.TaskPath -Confirm:$false -ErrorAction Stop
        }
        Add-OperationResult -Results $_results -Target "$($_task.TaskPath)$($_task.TaskName)" -Action 'RemoveScheduledTask' -Status 'Removed' -Detail ''
      }
      catch {
        Add-OperationResult -Results $_results -Target "$($_task.TaskPath)$($_task.TaskName)" -Action 'RemoveScheduledTask' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
  }
}

if ($RemoveUserFolder) {
  $_oneDriveRoots = @(Get-OneDrivePath | Where-Object { $_ -like "$env:USERPROFILE*" })
  foreach ($_root in $_oneDriveRoots) {
    $_resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_root)
    $_profile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($env:USERPROFILE)
    if (-not $_resolved.StartsWith($_profile, [System.StringComparison]::OrdinalIgnoreCase)) {
      Add-OperationResult -Results $_results -Target $_root -Action 'RemoveUserFolder' -Status 'Skipped' -Detail 'Path is outside the current user profile.'
      continue
    }

    $_children = @(Get-ChildItem -LiteralPath $_resolved -Force -ErrorAction SilentlyContinue)
    if ($_children.Count -gt 0 -and -not $Force) {
      Add-OperationResult -Results $_results -Target $_resolved -Action 'RemoveUserFolder' -Status 'Skipped' -Detail 'Folder is not empty. Re-run with -Force only after validating contents.'
      continue
    }

    if ($PSCmdlet.ShouldProcess($_resolved, 'Remove OneDrive user folder')) {
      try {
        if (-not $DryRun) {
          Remove-Item -LiteralPath $_resolved -Force -Recurse:$Force -ErrorAction Stop
        }
        Add-OperationResult -Results $_results -Target $_resolved -Action 'RemoveUserFolder' -Status 'Removed' -Detail ''
      }
      catch {
        Add-OperationResult -Results $_results -Target $_resolved -Action 'RemoveUserFolder' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
  }
}

if ($Instant) {
  if ($DryRun) {
    Add-OperationResult -Results $_results -Target 'explorer.exe' -Action 'RestartExplorer' -Status 'Skipped' -Detail 'DryRun'
  }
  elseif ($PSCmdlet.ShouldProcess('explorer.exe', 'Restart Explorer')) {
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Process explorer
    Add-OperationResult -Results $_results -Target 'explorer.exe' -Action 'RestartExplorer' -Status 'Completed' -Detail ''
  }
}

$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_skipped = @($_results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_completed = @($_results | Where-Object { $_.Status -notin @('Failed', 'Skipped') }).Count
$_color = if ($_failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Log -Message "`nOneDrive removal workflow complete. Completed: $_completed | Skipped: $_skipped | Failed: $_failed" -Color $_color
$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Remove-OneDrive'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
