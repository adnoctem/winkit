#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.4.0' }

<#
.SYNOPSIS
  Backs up KeePass (KDBX) databases to every attached drive marked for them.
.DESCRIPTION
  Intended for offline backup media such as USB sticks and external disks: run it
  on a schedule and whichever marked drive happens to be attached receives a
  backup. When no marked drive is attached, the run is skipped without error.

  A DRIVE ONLY RECEIVES BACKUPS ONCE IT IS MARKED. Mark it with
  New-DriveMarker.ps1, which writes a hidden marker file to the drive root holding
  a unique ID, and pass that ID as -MarkerId:

    PS> .\New-DriveMarker.ps1 -DriveRoot F:\

  The ID is what makes the marker safe. Matching on the file name alone would let
  any stick carrying a file of that name silently receive copies of the
  databases. Several drives can share one ID (pass -MarkerId to New-DriveMarker.ps1)
  so that one scheduled task serves a set of rotating backup media.

  Every attached local volume (removable and fixed; network drives are excluded)
  is checked for a marker whose ID matches -MarkerId. A drive that holds one of the
  source paths is skipped, because a copy on the same disk is not a separate
  backup. Each matching drive is then backed up with Backup-KdbxDatabase.ps1 into
  -Destination, so the copy is verified, never overwrites an existing backup, and
  records a SHA256SUMS manifest. See that script for the details.
.PARAMETER MarkerId
  The ID written by New-DriveMarker.ps1. Only drives whose marker holds this ID
  receive backups.
.PARAMETER DatabasePath
  One or more live KeePass database files to snapshot.
.PARAMETER BackupPath
  Folder whose *.kdbx files are backed up, typically KeePassXC's backup folder.
.PARAMETER Destination
  Folder on each marked drive that receives the backup, relative to the drive root.
  Defaults to KeePass-Backups.
.PARAMETER MarkerName
  File name of the hidden marker at the drive root. Defaults to .backup.marker.
.PARAMETER DriveRoot
  Roots to search instead of all attached local volumes, for example a volume
  mounted into a folder.
.PARAMETER Verify
  Verify each marked drive's backup against its SHA256SUMS manifest instead of copying.
.PARAMETER DryRun
  Report what would be copied without writing anything.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> .\New-DriveMarker.ps1 -DriveRoot F:\
  PS> .\Backup-KdbxDatabaseToDrive.ps1 -MarkerId <ID printed above> -DatabasePath "$env:USERPROFILE\Documents\Passwords.kdbx" -BackupPath "$env:USERPROFILE\Documents\KeePass Backups"
  Marks a USB stick once, then backs up to it whenever it is attached.
.EXAMPLE
  PS> .\Backup-KdbxDatabaseToDrive.ps1 -MarkerId 5f0c7a3e-2b9d-4e1a-9c67-0d8e4b2f1a93 -Verify
  Verifies the backup on every attached marked drive.
.EXAMPLE
  PS> $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\winkit\scripts\Maintenance\Backup-KdbxDatabaseToDrive.ps1" -MarkerId 5f0c7a3e-2b9d-4e1a-9c67-0d8e4b2f1a93 -DatabasePath "C:\Users\me\Documents\Passwords.kdbx"'
  PS> Register-ScheduledTask -TaskName 'Backup KeePass databases to offline drive' -Action $action -Trigger (New-ScheduledTaskTrigger -AtLogOn) -User "$env:USERDOMAIN\$env:USERNAME"
  Backs up to an attached marked drive at every logon.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - file operations only.
  SYSTEM-account execution: supported for local and removable drives; the source databases must be readable by the account.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [guid]
  $MarkerId,

  [string[]]
  $DatabasePath,

  [string]
  $BackupPath,

  [ValidateScript({
      if ([System.IO.Path]::IsPathRooted($_) -or $_ -match '(^|[\\/])\.\.([\\/]|$)') { throw 'Destination must be a relative path inside the drive, without "..".' }
      $true
    })]
  [ValidateNotNullOrEmpty()]
  [string]
  $Destination = 'KeePass-Backups',

  [ValidatePattern('^[^\\/:*?"<>|]+$')]
  [string]
  $MarkerName = '.backup.marker',

  [string[]]
  $DriveRoot,

  [switch]
  $Verify,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - nothing will be written`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_emitResults = $PassThru -or $DryRun
$_backupScript = Join-Path -Path $PSScriptRoot -ChildPath 'Backup-KdbxDatabase.ps1'

function Complete-DriveBackup {
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Backup-KdbxDatabaseToDrive'
  if ($_operationLog) {
    Write-Log -Message "Operation log: $_operationLog" -Color Gray
  }
  if ($_emitResults) {
    $_results
  }
  if (@($_results | Where-Object { $_.Status -in @('Failed', 'Conflict') }).Count -gt 0) {
    exit 1
  }
  exit 0
}

if (-not $Verify -and -not $DatabasePath -and -not $BackupPath) {
  Write-Log -Message 'FAILED - Specify -DatabasePath, -BackupPath, or both.' -Color Red
  Add-OperationResult -Results $_results -Target 'Parameters' -Source 'KdbxDriveBackup' -Action 'Validate' -Status 'Failed' -Detail 'Specify -DatabasePath, -BackupPath, or both.'
  Complete-DriveBackup
}

# ---- Find marked drives ------------------------------------------------------

if ($DriveRoot) {
  $_roots = @($DriveRoot | ForEach-Object { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_) })
}
else {
  # DriveType 2 = removable, 3 = local fixed. Network drives (4) are excluded: a
  # disconnected share can hang enumeration and is not offline media anyway.
  $_roots = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 2 OR DriveType = 3' -ErrorAction Stop | ForEach-Object { "$($_.DeviceID)\" })
}

$_sources = @(@($DatabasePath) + @($BackupPath) | Where-Object { $_ } | ForEach-Object { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_) })

Write-Log -Message "Looking for drives marked with $MarkerName" -Color Cyan

$_matched = New-Object System.Collections.Generic.List[string]
foreach ($_root in $_roots) {
  $_marker = Join-Path -Path $_root -ChildPath $MarkerName
  try {
    if (-not (Test-Path -LiteralPath $_marker -PathType Leaf)) { continue }
    $_firstLine = [string](Get-Content -LiteralPath $_marker -TotalCount 1 -ErrorAction Stop)
  }
  catch {
    continue
  }

  $_found = [guid]::Empty
  if (-not [guid]::TryParse($_firstLine.Trim().TrimStart([char]0xFEFF), [ref]$_found)) {
    $_detail = 'Marker file found, but its first line is not an ID. Recreate it with New-DriveMarker.ps1.'
    Write-Log -Message "  [Warn] ${_root}: $_detail" -Color Yellow
    Add-OperationResult -Results $_results -Target $_root -Source 'KdbxDriveBackup' -Action 'Discover' -Status 'Warn' -Detail $_detail
    continue
  }
  if ($_found -ne $MarkerId) {
    $_detail = 'Marker ID does not match -MarkerId; this drive is marked for a different backup.'
    Write-Log -Message "  [Skipped] ${_root}: $_detail" -Color Gray
    Add-OperationResult -Results $_results -Target $_root -Source 'KdbxDriveBackup' -Action 'Discover' -Status 'Skipped' -Detail $_detail
    continue
  }

  $_rootWithSeparator = $_root.TrimEnd('\') + '\'
  $_hostsSource = @($_sources | Where-Object { $_.StartsWith($_rootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase) })
  if ($_hostsSource.Count -gt 0) {
    $_detail = "Marked, but it holds a source ($($_hostsSource[0])); a copy on the same disk is not a separate backup."
    Write-Log -Message "  [Skipped] ${_root}: $_detail" -Color Yellow
    Add-OperationResult -Results $_results -Target $_root -Source 'KdbxDriveBackup' -Action 'Discover' -Status 'Skipped' -Detail $_detail
    continue
  }

  $_matched.Add($_root)
}

if ($_matched.Count -eq 0) {
  $_detail = 'No attached drive is marked for this backup. Mark one with New-DriveMarker.ps1.'
  Write-Log -Message $_detail -Color Yellow
  Add-OperationResult -Results $_results -Target 'Drives' -Source 'KdbxDriveBackup' -Action 'Discover' -Status 'Skipped' -Detail $_detail
  Complete-DriveBackup
}

# ---- Back up to each marked drive ----------------------------------------------

foreach ($_root in $_matched) {
  $_destination = Join-Path -Path $_root -ChildPath $Destination
  Write-Log -Message "`nMarked drive: $_root" -Color Cyan

  $_arguments = @{ Destination = $_destination; PassThru = $true }
  if ($DatabasePath) { $_arguments['DatabasePath'] = $DatabasePath }
  if ($BackupPath) { $_arguments['BackupPath'] = $BackupPath }
  if ($Verify) { $_arguments['Verify'] = $true }
  if ($DryRun) { $_arguments['DryRun'] = $true }
  if ($WhatIfPreference -and -not $DryRun) { $_arguments['WhatIf'] = $true }

  $global:LASTEXITCODE = 0
  $_driveResults = @(& $_backupScript @_arguments)
  $_exitCode = $LASTEXITCODE
  foreach ($_result in $_driveResults) { [void]$_results.Add($_result) }

  $_status = if ($_exitCode -eq 0) { 'Completed' } else { 'Failed' }
  $_action = if ($Verify) { 'Verify' } else { 'Backup' }
  Add-OperationResult -Results $_results -Target $_root -Source 'KdbxDriveBackup' -Action $_action -Status $_status -Detail "Backup-KdbxDatabase.ps1 exited with code $_exitCode for $_destination."
}

Complete-DriveBackup
