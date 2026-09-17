#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.4.0' }

<#
.SYNOPSIS
  Backs up KeePass (KDBX) databases and their backup folder to another location.
.DESCRIPTION
  Copies KeePass databases to a destination folder such as a mounted cloud
  drive, laying the destination out as:

    <Destination>\backups\     the files from -BackupPath (e.g. KeePassXC's
                               "backup before saving" copies), under their own names
    <Destination>\snapshots\   copies of each -DatabasePath file, named after the
                               database's last save time
    <Destination>\SHA256SUMS   a sha256sum-compatible manifest of every copied file

  The backup folder only holds versions that a later save replaced, so the newest
  data lives solely in the live database until the next save. -DatabasePath
  snapshots close that gap; a snapshot is skipped while the database is unchanged
  since its latest snapshot.

  Nothing at the destination is ever overwritten or deleted. Every file is copied
  under a temporary name, verified by SHA-256 against the source, and only then
  renamed into place, so an interrupted copy never leaves a file that looks like a
  valid backup. A live database that changes while it is copied is retried once.

  A destination file with the same name but different content is a conflict: the
  destination copy is kept and the run exits 1. Backup files never change after
  they are written, so a conflict usually means the source was altered - for
  example encrypted in place by ransomware - and must not replace a good copy.

  Every file is checked for the KDBX signature. A live database without one fails
  the run and is not snapshotted; a backup file without one is still copied, with
  a warning, so no data is dropped.

  The destination folder is created only when its parent already exists. When a
  mounted drive is missing (for example G:\My Drive), the run fails instead of
  silently writing the "offsite" copy to the local disk.

  -Verify re-hashes the destination against SHA256SUMS and copies nothing, to
  detect bit rot or tampering on backup media.
.PARAMETER Destination
  Folder that receives the backup. Its parent folder must exist.
.PARAMETER DatabasePath
  One or more live KeePass database files to snapshot.
.PARAMETER BackupPath
  Folder whose *.kdbx files are backed up, typically the folder KeePassXC writes
  "backup before saving" copies to.
.PARAMETER Verify
  Verify the destination against its SHA256SUMS manifest instead of copying.
.PARAMETER DryRun
  Report what would be copied without writing anything.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> .\Backup-KdbxDatabase.ps1 -Destination 'G:\My Drive\Backups\KeePass' -DatabasePath "$env:USERPROFILE\Documents\Passwords.kdbx" -BackupPath "$env:USERPROFILE\Documents\KeePass Backups"
  Backs up the database and its backup folder to a mounted Google Drive.
.EXAMPLE
  PS> .\Backup-KdbxDatabase.ps1 -Destination 'G:\My Drive\Backups\KeePass' -Verify
  Checks every file in the destination against its recorded hash.
.EXAMPLE
  PS> $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\winkit\scripts\Maintenance\Backup-KdbxDatabase.ps1" -Destination "G:\My Drive\Backups\KeePass" -DatabasePath "C:\Users\me\Documents\Passwords.kdbx"'
  PS> Register-ScheduledTask -TaskName 'Backup KeePass databases' -Action $action -Trigger (New-ScheduledTaskTrigger -AtLogOn) -User "$env:USERDOMAIN\$env:USERNAME"
  Runs the backup at every logon as the signed-in user.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - file operations only.
  SYSTEM-account execution: works for local and removable destinations, but run as the signed-in user for cloud-drive mounts (such as Google Drive for desktop), which SYSTEM cannot see.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $Destination,

  [string[]]
  $DatabasePath,

  [string]
  $BackupPath,

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
$_destination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination)
$_manifestPath = Join-Path $_destination 'SHA256SUMS'
$_utf8 = New-Object System.Text.UTF8Encoding($false)
$_kdbxSignature = [byte[]](0x03, 0xD9, 0xA2, 0x9A, 0x67, 0xFB, 0x4B, 0xB5)

function Complete-Backup {
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Backup-KdbxDatabase'
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

function Add-BackupResult {
  param ([string]$Target, [string]$Action, [string]$Status, [string]$Detail)

  $_color = switch ($Status) {
    'Completed' { 'Green' }
    'Skipped' { 'Gray' }
    'Warn' { 'Yellow' }
    default { 'Red' }
  }
  Write-Log -Message "  [$Status] ${Target}: $Detail" -Color $_color
  Add-OperationResult -Results $_results -Target $Target -Source 'KdbxBackup' -Action $Action -Status $Status -Detail $Detail
}

function Open-SharedRead {
  # KeePass may hold the database open; share read, write and delete so the
  # backup never blocks or is blocked by a save in progress.
  param ([string]$Path)
  return [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
}

function Get-Sha256 {
  param ([string]$Path)

  $_stream = Open-SharedRead -Path $Path
  $_sha = [System.Security.Cryptography.SHA256]::Create()
  try {
    return ([BitConverter]::ToString($_sha.ComputeHash($_stream)) -replace '-', '').ToLowerInvariant()
  }
  finally {
    $_sha.Dispose()
    $_stream.Dispose()
  }
}

function Test-KdbxSignature {
  param ([string]$Path)

  $_stream = Open-SharedRead -Path $Path
  try {
    $_header = New-Object byte[] 8
    $_read = $_stream.Read($_header, 0, 8)
  }
  finally {
    $_stream.Dispose()
  }
  if ($_read -lt 8) { return $false }
  for ($_i = 0; $_i -lt 8; $_i++) {
    if ($_header[$_i] -ne $_kdbxSignature[$_i]) { return $false }
  }
  return $true
}

function Read-BackupManifest {
  $_entries = @{}
  if (Test-Path -LiteralPath $_manifestPath -PathType Leaf) {
    foreach ($_line in [System.IO.File]::ReadAllLines($_manifestPath)) {
      if ($_line -match '^([0-9a-fA-F]{64}) [ *](.+)$') {
        $_entries[$Matches[2].ToLowerInvariant()] = @{ Hash = $Matches[1].ToLowerInvariant(); Path = $Matches[2] }
      }
    }
  }
  return $_entries
}

function Add-ManifestEntry {
  param ([string]$RelativePath, [string]$Hash)

  if ($_manifest.ContainsKey($RelativePath.ToLowerInvariant())) { return }
  [System.IO.File]::AppendAllText($_manifestPath, "$Hash *$RelativePath`r`n", $_utf8)
  $_manifest[$RelativePath.ToLowerInvariant()] = @{ Hash = $Hash; Path = $RelativePath }
}

function Copy-BackupFile {
  # Copies one file without ever replacing an existing one. Returns a status
  # object; 'Changed' means the source changed during the copy.
  [CmdletBinding(SupportsShouldProcess = $true)]
  param ([string]$SourcePath, [string]$RelativePath)

  $_target = Join-Path $_destination ($RelativePath -replace '/', '\')
  $_hash = Get-Sha256 -Path $SourcePath

  if (Test-Path -LiteralPath $_target -PathType Leaf) {
    if ((Get-Sha256 -Path $_target) -eq $_hash) {
      if (-not $DryRun -and $PSCmdlet.ShouldProcess($_manifestPath, "Record $RelativePath")) {
        Add-ManifestEntry -RelativePath $RelativePath -Hash $_hash
      }
      return @{ Status = 'Skipped'; Detail = 'Already backed up.' }
    }
    return @{ Status = 'Conflict'; Detail = 'A different file with this name already exists at the destination. It was kept; the source may have been altered.' }
  }

  if ($DryRun) {
    return @{ Status = 'Skipped'; Detail = 'DryRun: would copy.' }
  }
  if (-not $PSCmdlet.ShouldProcess($_target, "Copy $SourcePath")) {
    return @{ Status = 'Skipped'; Detail = 'WhatIf' }
  }

  $null = New-Item -ItemType Directory -Path (Split-Path -Path $_target -Parent) -Force
  $_lastWrite = [System.IO.File]::GetLastWriteTimeUtc($SourcePath)
  $_temporary = "$_target.$([guid]::NewGuid().ToString('N')).tmp"
  try {
    $_source = Open-SharedRead -Path $SourcePath
    try {
      $_output = [System.IO.File]::Open($_temporary, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
      $_sha = [System.Security.Cryptography.SHA256]::Create()
      try {
        $_buffer = New-Object byte[] 81920
        while (($_read = $_source.Read($_buffer, 0, $_buffer.Length)) -gt 0) {
          $_output.Write($_buffer, 0, $_read)
          [void]$_sha.TransformBlock($_buffer, 0, $_read, $null, 0)
        }
        [void]$_sha.TransformFinalBlock((New-Object byte[] 0), 0, 0)
        $_output.Flush($true)
        $_copiedHash = ([BitConverter]::ToString($_sha.Hash) -replace '-', '').ToLowerInvariant()
      }
      finally {
        $_sha.Dispose()
        $_output.Dispose()
      }
    }
    finally {
      $_source.Dispose()
    }

    if ((Get-Sha256 -Path $SourcePath) -ne $_hash -or $_copiedHash -ne $_hash) {
      return @{ Status = 'Changed'; Detail = 'The source changed while it was being copied.' }
    }
    if ((Get-Sha256 -Path $_temporary) -ne $_hash) {
      return @{ Status = 'Failed'; Detail = 'The copy did not verify against the source hash.' }
    }

    [System.IO.File]::Move($_temporary, $_target)
    [System.IO.File]::SetLastWriteTimeUtc($_target, $_lastWrite)
    Add-ManifestEntry -RelativePath $RelativePath -Hash $_hash
    return @{ Status = 'Completed'; Detail = "Copied and verified (SHA-256 $($_hash.Substring(0, 12)))." }
  }
  catch {
    return @{ Status = 'Failed'; Detail = $_.Exception.Message }
  }
  finally {
    if (Test-Path -LiteralPath $_temporary) {
      Remove-Item -LiteralPath $_temporary -Force -ErrorAction SilentlyContinue
    }
  }
}

# ---- Verify ------------------------------------------------------------------

if ($Verify) {
  Write-Log -Message "Verifying KeePass backup at $_destination" -Color Cyan

  if (-not (Test-Path -LiteralPath $_manifestPath -PathType Leaf)) {
    Add-BackupResult -Target $_destination -Action 'Verify' -Status 'Failed' -Detail 'No SHA256SUMS manifest found.'
    Complete-Backup
  }

  $_manifest = Read-BackupManifest
  $_verified = 0
  foreach ($_entry in $_manifest.Values) {
    $_file = Join-Path $_destination ($_entry.Path -replace '/', '\')
    if (-not (Test-Path -LiteralPath $_file -PathType Leaf)) {
      Add-BackupResult -Target $_entry.Path -Action 'Verify' -Status 'Failed' -Detail 'Recorded in SHA256SUMS but missing.'
    }
    elseif ((Get-Sha256 -Path $_file) -ne $_entry.Hash) {
      Add-BackupResult -Target $_entry.Path -Action 'Verify' -Status 'Failed' -Detail 'Hash mismatch - the file is corrupted or was modified.'
    }
    else {
      $_verified++
    }
  }

  foreach ($_folder in @('backups', 'snapshots')) {
    $_directory = Join-Path $_destination $_folder
    if (-not (Test-Path -LiteralPath $_directory -PathType Container)) { continue }
    foreach ($_file in @(Get-ChildItem -LiteralPath $_directory -Filter '*.kdbx' -File)) {
      $_relative = "$_folder/$($_file.Name)"
      if (-not $_manifest.ContainsKey($_relative.ToLowerInvariant())) {
        Add-BackupResult -Target $_relative -Action 'Verify' -Status 'Warn' -Detail 'Present but not recorded in SHA256SUMS; its integrity cannot be checked.'
      }
    }
  }

  Add-BackupResult -Target $_destination -Action 'Verify' -Status 'Completed' -Detail "$_verified of $($_manifest.Count) recorded file(s) verified."
  Complete-Backup
}

# ---- Validate inputs -----------------------------------------------------------

if (-not $DatabasePath -and -not $BackupPath) {
  Add-BackupResult -Target 'Parameters' -Action 'Validate' -Status 'Failed' -Detail 'Specify -DatabasePath, -BackupPath, or both.'
  Complete-Backup
}

Write-Log -Message "Backing up KeePass databases to $_destination" -Color Cyan

if (-not (Test-Path -LiteralPath $_destination -PathType Container)) {
  $_parent = Split-Path -Path $_destination -Parent
  if (-not $_parent -or -not (Test-Path -LiteralPath $_parent -PathType Container)) {
    Add-BackupResult -Target $_destination -Action 'Validate' -Status 'Failed' -Detail "The destination's parent folder does not exist. Is the drive mounted?"
    Complete-Backup
  }
  if ($DryRun) {
    Add-BackupResult -Target $_destination -Action 'CreateDestination' -Status 'Skipped' -Detail 'DryRun: would create the destination folder.'
  }
  elseif ($PSCmdlet.ShouldProcess($_destination, 'Create backup destination')) {
    $null = New-Item -ItemType Directory -Path $_destination -Force
  }
}

$_manifest = Read-BackupManifest

# ---- Live database snapshots -----------------------------------------------------

if ($DatabasePath) {
  $_databases = @($DatabasePath | ForEach-Object { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_) })
  $_duplicates = @($_databases | Group-Object { [System.IO.Path]::GetFileNameWithoutExtension($_) } | Where-Object { $_.Count -gt 1 })

  foreach ($_database in $_databases) {
    $_baseName = [System.IO.Path]::GetFileNameWithoutExtension($_database)

    if (@($_duplicates | Where-Object { $_.Name -eq $_baseName }).Count -gt 0) {
      Add-BackupResult -Target $_database -Action 'Snapshot' -Status 'Failed' -Detail "Another -DatabasePath is also named '$_baseName'; their snapshots would collide. Rename one database."
      continue
    }
    if (-not (Test-Path -LiteralPath $_database -PathType Leaf)) {
      Add-BackupResult -Target $_database -Action 'Snapshot' -Status 'Failed' -Detail 'Database file not found.'
      continue
    }
    if (-not (Test-KdbxSignature -Path $_database)) {
      Add-BackupResult -Target $_database -Action 'Snapshot' -Status 'Failed' -Detail 'Not a valid KDBX database (signature mismatch). No snapshot was taken.'
      continue
    }

    $_pattern = '^' + [regex]::Escape($_baseName) + '_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.kdbx$'
    for ($_attempt = 1; $_attempt -le 2; $_attempt++) {
      $_latest = $null
      $_snapshotDirectory = Join-Path $_destination 'snapshots'
      if (Test-Path -LiteralPath $_snapshotDirectory -PathType Container) {
        $_latest = @(Get-ChildItem -LiteralPath $_snapshotDirectory -File | Where-Object { $_.Name -match $_pattern } | Sort-Object Name -Descending) | Select-Object -First 1
      }
      if ($_latest -and (Get-Sha256 -Path $_latest.FullName) -eq (Get-Sha256 -Path $_database)) {
        $_outcome = @{ Status = 'Skipped'; Detail = "Unchanged since snapshot $($_latest.Name)." }
        break
      }

      $_savedAt = (Get-Item -LiteralPath $_database).LastWriteTime.ToString('yyyy-MM-dd_HH-mm-ss', [Globalization.CultureInfo]::InvariantCulture)
      $_outcome = Copy-BackupFile -SourcePath $_database -RelativePath "snapshots/${_baseName}_$_savedAt.kdbx"
      if ($_outcome.Status -ne 'Changed') { break }
      if ($_attempt -eq 2) {
        $_outcome = @{ Status = 'Failed'; Detail = 'The database changed during two copy attempts (saved while backing up?). Run the backup again.' }
      }
    }
    Add-BackupResult -Target $_database -Action 'Snapshot' -Status $_outcome.Status -Detail $_outcome.Detail
  }
}

# ---- Backup folder ------------------------------------------------------------------

if ($BackupPath) {
  $_backupFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BackupPath)
  if (-not (Test-Path -LiteralPath $_backupFolder -PathType Container)) {
    Add-BackupResult -Target $_backupFolder -Action 'Backup' -Status 'Failed' -Detail 'Backup folder not found.'
  }
  else {
    $_files = @(Get-ChildItem -LiteralPath $_backupFolder -Filter '*.kdbx' -File | Sort-Object Name)
    if ($_files.Count -eq 0) {
      Add-BackupResult -Target $_backupFolder -Action 'Backup' -Status 'Skipped' -Detail 'No *.kdbx files found.'
    }
    foreach ($_file in $_files) {
      if (-not (Test-KdbxSignature -Path $_file.FullName)) {
        Add-BackupResult -Target $_file.Name -Action 'Backup' -Status 'Warn' -Detail 'Not a valid KDBX database (signature mismatch); copying it anyway so no data is dropped.'
      }
      $_outcome = Copy-BackupFile -SourcePath $_file.FullName -RelativePath "backups/$($_file.Name)"
      if ($_outcome.Status -eq 'Changed') {
        $_outcome = @{ Status = 'Failed'; Detail = 'The backup file changed while it was being copied; backup files are not expected to change.' }
      }
      Add-BackupResult -Target $_file.Name -Action 'Backup' -Status $_outcome.Status -Detail $_outcome.Detail
    }
  }
}

$_copied = @($_results | Where-Object { $_.Status -eq 'Completed' }).Count
$_conflicts = @($_results | Where-Object { $_.Status -eq 'Conflict' }).Count
$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_color = if ($_conflicts + $_failed -gt 0) { 'Red' } else { 'Green' }
Write-Log -Message "Done: $_copied copied, $_conflicts conflict(s), $_failed failure(s)." -Color $_color

Complete-Backup
