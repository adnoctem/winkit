#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.4.0' }

<#
.SYNOPSIS
  Marks a drive to receive backups from Backup-KdbxDatabaseToDrive.ps1.
.DESCRIPTION
  Writes a marker file to the root of a drive, such as a USB stick or external
  disk. Its first line is a unique ID and its second line a comment naming its
  purpose. Backup-KdbxDatabaseToDrive.ps1 only backs up to drives whose marker
  holds the ID it is given, so a stick that merely carries a file of the same name
  receives nothing.

  The marker is hidden, so it stays out of the way when the drive is also used for
  ordinary file transfers. Hidden files are still found by name, so this does not
  affect the backup.

  Without -MarkerId a new ID is generated. Pass an existing ID to mark several
  drives for the same backup, for example a set of rotating offline media served by
  one scheduled task.

  An existing marker is only replaced with -Force: changing its ID silently stops
  that drive from receiving backups until the task is updated. Marking the drive
  Windows runs from produces a warning, since a backup there is not offline.

  The ID is printed together with a ready-to-use Backup-KdbxDatabaseToDrive.ps1
  command.
.PARAMETER DriveRoot
  Root of the drive to mark, for example F:\, or the folder a volume is mounted to.
.PARAMETER MarkerId
  ID to write. Defaults to a newly generated one.
.PARAMETER MarkerName
  File name of the marker. Defaults to .backup.marker; must match the -MarkerName
  used by Backup-KdbxDatabaseToDrive.ps1.
.PARAMETER Force
  Replace an existing marker.
.PARAMETER DryRun
  Report the marker that would be written without writing it.
.PARAMETER PassThru
  Return structured operation results, including the MarkerId.
.EXAMPLE
  PS> .\New-DriveMarker.ps1 -DriveRoot F:\
  Marks F:\ with a new ID and prints it.
.EXAMPLE
  PS> .\New-DriveMarker.ps1 -DriveRoot G:\ -MarkerId 5f0c7a3e-2b9d-4e1a-9c67-0d8e4b2f1a93
  Marks a second drive for the same backup as an existing one.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - file operations only.
  SYSTEM-account execution: supported; the account needs write access to the drive root.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $DriveRoot,

  [guid]
  $MarkerId = [guid]::NewGuid(),

  [ValidatePattern('^[^\\/:*?"<>|]+$')]
  [string]
  $MarkerName = '.backup.marker',

  [switch]
  $Force,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - the marker will not be written`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_emitResults = $PassThru -or $DryRun
$_root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DriveRoot)
$_markerPath = Join-Path -Path $_root -ChildPath $MarkerName
$_properties = @{ MarkerId = "$MarkerId"; Path = $_markerPath }

function Complete-Marker {
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'New-DriveMarker'
  if ($_operationLog) {
    Write-Log -Message "Operation log: $_operationLog" -Color Gray
  }
  if ($_emitResults) {
    $_results
  }
  if (@($_results | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) {
    exit 1
  }
  exit 0
}

if (-not (Test-Path -LiteralPath $_root -PathType Container)) {
  $_detail = "Drive root not found: $_root"
  Write-Log -Message "FAILED - $_detail" -Color Red
  Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Failed' -Detail $_detail
  Complete-Marker
}

if ([System.IO.Path]::GetPathRoot($_root) -ieq [System.IO.Path]::GetPathRoot($env:SystemRoot)) {
  $_detail = "$_root is on the system drive; a backup there is not an offline copy."
  Write-Log -Message "WARNING - $_detail" -Color Yellow
  Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Warn' -Detail $_detail
}

if (Test-Path -LiteralPath $_markerPath -PathType Leaf) {
  $_existing = [guid]::Empty
  $_firstLine = [string](Get-Content -LiteralPath $_markerPath -TotalCount 1)
  $_parsed = [guid]::TryParse($_firstLine.Trim().TrimStart([char]0xFEFF), [ref]$_existing)

  if ($_parsed -and $_existing -eq $MarkerId) {
    Write-Log -Message "$_root is already marked with ID $MarkerId." -Color Gray
    Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Skipped' -Detail 'Already marked with this ID.' -Property $_properties
    Complete-Marker
  }
  if (-not $Force) {
    $_current = if ($_parsed) { "ID $_existing" } else { 'an unreadable ID' }
    $_detail = "$_root is already marked with $_current. Use -Force to replace it, or pass -MarkerId to reuse the existing ID."
    Write-Log -Message "FAILED - $_detail" -Color Red
    Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Failed' -Detail $_detail
    Complete-Marker
  }
}

$_command = ".\Backup-KdbxDatabaseToDrive.ps1 -MarkerId $MarkerId -DatabasePath <database.kdbx> -BackupPath <backup folder>"
if ($MarkerName -ne '.backup.marker') {
  $_command += " -MarkerName '$MarkerName'"
}

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would write $_markerPath with ID $MarkerId." -Color Yellow
  Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Skipped' -Detail "DryRun: would write ID $MarkerId." -Property $_properties
}
elseif ($PSCmdlet.ShouldProcess($_markerPath, "Write drive marker with ID $MarkerId")) {
  try {
    $_content = "$MarkerId`r`n# winkit drive marker: this drive receives Backup-KdbxDatabaseToDrive.ps1 backups. Do not edit or delete.`r`n"
    # Writing over a hidden file throws, so clear the attribute before replacing one.
    if (Test-Path -LiteralPath $_markerPath -PathType Leaf) {
      (Get-Item -LiteralPath $_markerPath -Force).Attributes = [System.IO.FileAttributes]::Normal
    }
    [System.IO.File]::WriteAllText($_markerPath, $_content, (New-Object System.Text.UTF8Encoding($false)))
    # Hidden keeps the marker out of the way when the drive is used for other transfers.
    (Get-Item -LiteralPath $_markerPath -Force).Attributes = [System.IO.FileAttributes]::Hidden
    Write-Log -Message "Marked $_root" -Color Green
    Write-Log -Message "  Marker ID: $MarkerId" -Color Green
    Write-Log -Message "  Back up to this drive with:" -Color Gray
    Write-Log -Message "    $_command" -Color Gray
    Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Completed' -Detail "Marked with ID $MarkerId." -Property $_properties
  }
  catch {
    Write-Log -Message "FAILED - $($_.Exception.Message)" -Color Red
    Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Failed' -Detail $_.Exception.Message
  }
}
else {
  Add-OperationResult -Results $_results -Target $_markerPath -Source 'DriveMarker' -Action 'Create' -Status 'Skipped' -Detail 'WhatIf' -Property $_properties
}

Complete-Marker
