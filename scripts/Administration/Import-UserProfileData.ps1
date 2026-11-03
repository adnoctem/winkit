#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.1' }

<#
.SYNOPSIS
  Imports staged user data into an existing, signed-out Windows profile.
.DESCRIPTION
  Copies ordinary data folders from an extracted backup using robocopy, checks every
  included file with SHA-256, and applies the target profile's inherited permissions
  and user ownership to the selected destination trees. Existing custom permissions
  within those trees are reset, including on destination-only files. Other profile
  folders and the profile registry are not changed.

  Resolve the target account by name or SID. It must already have a local, unloaded
  profile. Source and destination must be separate local NTFS directories. Keep the
  user signed out and the staging data unchanged for the entire run.

  ABB export/download and extraction are preparation steps. This script does not
  contact Synology or restore a disk image. AppData, profile hives, PSTs and OSTs are
  excluded. Source junctions and symbolic links are skipped; destination reparse
  points, encrypted files, and offline/cloud placeholders stop the run.

  The default folders are Desktop, Documents, Downloads, Pictures, Music, Videos,
  Favorites, Contacts, Links, Saved Games, and Searches. Missing default folders are
  reported and skipped. Explicitly selected or mapped source folders must exist.
  Paths are literal on-disk names, not localized Explorer display names.
.PARAMETER SourcePath
  Extracted old profile directory, for example C:\Migration\User1. Select the
  directory containing Documents and Desktop, not the ZIP file or disk root.
.PARAMETER TargetUser
  Qualified account name (NEWDOMAIN\User1, COMPUTER\User1) or user SID. The user must
  have signed in once to create their profile, then signed out.
.PARAMETER Folder
  Relative data directories to copy to the same relative path in the target profile.
  Replaces the default list. AppData and system/profile directories are not allowed.
.PARAMETER FolderMap
  Additional source-relative to target-relative directory mappings, for example
  @{ 'Business Data' = 'Documents\Business Data' }. Mappings may not overlap each
  other or selected Folder paths. Use -Folder @() for mappings only.
  Redirected/OneDrive folders are not discovered automatically: specify their actual
  local relative destination explicitly, with sync stopped and files hydrated.
.PARAMETER Overwrite
  Allow differing files at matching destination paths to be replaced. Without this
  switch, any content conflict fails before copying. Destination-only files remain.
.PARAMETER RemoveSource
  After ALL folders pass copy, content, permission, and audit checks, delete only
  verified included source files and empty selected directories. Excluded files,
  links, unselected folders, and SourcePath itself remain. Staging is retained by default.
.PARAMETER ReportDirectory
  Parent for unique run directories holding manifest.json, operations.jsonl and native
  logs. Defaults to ProgramData\winkit\ProfileImports. Must be outside source and target.
  Each run directory is restricted to the operator, Administrators, and SYSTEM.
.PARAMETER ReserveBytes
  Free bytes to retain after the conservative full-source-size estimate. Default: 256 MB.
.PARAMETER DryRun
  Validate, hash files, and report the plan without copying, changing ACLs, or writing logs.
.PARAMETER PassThru
  Return structured results, including the report path and per-folder status.
.EXAMPLE
  PS> .\Import-UserProfileData.ps1 -SourcePath C:\Migration\User1 -TargetUser 'NEWDOMAIN\User1' -DryRun
  Preview the standard data folders for an existing signed-out target user.
.EXAMPLE
  PS> .\Import-UserProfileData.ps1 -SourcePath C:\Migration\User1 -TargetUser 'NEWDOMAIN\User1' -Folder Desktop,Documents,Downloads -PassThru
  Copy the selected data, normalize their permissions, and retain staging.
.EXAMPLE
  PS> .\Import-UserProfileData.ps1 -SourcePath C:\Migration\User1 -TargetUser 'NEWDOMAIN\User1' -Folder @() -FolderMap @{ 'Business Data' = 'Documents\Business Data' } -RemoveSource -PassThru
  Import an explicitly mapped data directory and clean up only verified staged files.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - local filesystem, registry, CIM, and native Windows tools.
  SYSTEM-account execution: supported if the target account resolves and files are accessible.
  Work is not transactional: failed copies or ACL operations may leave partial changes.
  Do not use on a live profile, a writable shared staging tree, or an active sync folder.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $SourcePath,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $TargetUser,

  [AllowEmptyCollection()]
  [string[]]
  $Folder = @(
    'Desktop',
    'Documents',
    'Downloads',
    'Pictures',
    'Music',
    'Videos',
    'Favorites',
    'Contacts',
    'Links',
    'Saved Games',
    'Searches'
  ),

  [System.Collections.IDictionary]
  $FolderMap = @{},

  [switch]
  $Overwrite,

  [switch]
  $RemoveSource,

  [string]
  $ReportDirectory = (Join-Path $env:ProgramData 'winkit\ProfileImports'),

  [ValidateRange(0, [long]::MaxValue)]
  [long]
  $ReserveBytes = 256MB,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force

$ErrorActionPreference = 'Stop'
if ($DryRun) {
  $WhatIfPreference = $true
}

$_results = New-Object System.Collections.ArrayList
$_plan = New-Object System.Collections.ArrayList
$_runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$_runDirectory = $null
$_manifestPath = $null
$_failure = $null
$_status = 'Skipped'
$_source = $null
$_profile = $null
$_startedAt = [DateTime]::UtcNow.ToString('o')
$_actor = [Security.Principal.WindowsIdentity]::GetCurrent().Name

# Path checks use literal local paths and reject every existing reparse ancestor.
function Resolve-ImportPath {
  param (
    [string]
    $Path
  )

  $_resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
  if ($_resolved -notmatch '^[A-Za-z]:\\' -or $_resolved.Substring(2).Contains(':')) {
    throw "Use a local drive path without alternate data streams: $Path"
  }

  $_resolved = [IO.Path]::GetFullPath($_resolved).TrimEnd('\')
  if ($_resolved.Length -le 2) {
    throw 'A drive root cannot be used as an import directory.'
  }

  Assert-ImportPath -Path $_resolved
  return $_resolved
}

function Assert-ImportPath {
  param (
    [string]
    $Path
  )

  $_current = $Path
  while ($_current -and $_current.Length -gt 3) {
    if (Test-Path -LiteralPath $_current) {
      $_item = Get-Item -LiteralPath $_current -Force
      if (($_item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Reparse-point path is not allowed: $_current"
      }
    }

    $_current = Split-Path -Path $_current -Parent
  }
}

function Test-ImportPathWithin {
  param (
    [string]
    $Path,

    [string]
    $Root
  )

  return $Path.Equals($Root, [StringComparison]::OrdinalIgnoreCase) -or
  $Path.StartsWith($Root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Join-ImportDataPath {
  param (
    [string]
    $Root,

    [string]
    $RelativePath
  )

  $_parts = @($RelativePath -split '[\\/]')
  foreach ($_part in $_parts) {
    if ([string]::IsNullOrWhiteSpace($_part) -or $_part -in @('.', '..') -or
      $_part -match '[<>:"|?*]' -or $_part -match '[. ]$' -or
      $_part -match '^(AppData|Application Data|Local Settings|Windows|Program Files( \(x86\))?|ProgramData)$' -or
      $_part -match '^(NTUSER\.|UsrClass\.)') {
      throw "Not a relative user-data directory: '$RelativePath'."
    }
  }

  $_joined = [IO.Path]::GetFullPath((Join-Path $Root ($_parts -join '\')))
  if (-not (Test-ImportPathWithin -Path $_joined -Root $Root) -or $_joined -eq $Root) {
    throw "Data directory escapes its root: $RelativePath"
  }

  return $_joined
}

function Resolve-ImportIdentity {
  param (
    [string]
    $User
  )

  if ($User -match '^S-1-') {
    $_sid = New-Object Security.Principal.SecurityIdentifier($User)
    $null = $_sid.Translate([Security.Principal.NTAccount])
  }
  else {
    $_account = New-Object Security.Principal.NTAccount($User)
    $_sid = $_account.Translate([Security.Principal.SecurityIdentifier])
  }

  return [PSCustomObject]@{
    SID  = $_sid.Value
    User = $_sid.Translate([Security.Principal.NTAccount]).Value
  }
}

function Get-ImportProfile {
  param (
    [string]
    $User
  )

  $_identity = Resolve-ImportIdentity -User $User
  $_profiles = @(Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($_identity.SID)'")
  if ($_profiles.Count -ne 1 -or $_profiles[0].Special -or ($_profiles[0].Status -band 9) -ne 0) {
    throw "The target must have one ordinary local profile. Sign in once, then sign out: $User"
  }

  $_profilePath = Resolve-ImportPath -Path $_profiles[0].LocalPath
  $_registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{0}' -f $_identity.SID
  $_registered = (Get-ItemProperty -LiteralPath $_registryPath -Name ProfileImagePath).ProfileImagePath
  $_registered = Resolve-ImportPath -Path ([Environment]::ExpandEnvironmentVariables($_registered))
  if ($_registered -ne $_profilePath -or -not (Test-Path -LiteralPath $_profilePath -PathType Container) -or
    -not (Test-Path -LiteralPath (Join-Path $_profilePath 'NTUSER.DAT') -PathType Leaf)) {
    throw 'CIM and ProfileList must identify the same existing profile directory.'
  }

  if ($_profiles[0].Loaded -or
    (Test-Path -LiteralPath "Registry::HKEY_USERS\$($_identity.SID)") -or
    (Test-Path -LiteralPath "Registry::HKEY_USERS\$($_identity.SID)_Classes")) {
    throw "The target profile is loaded. Sign the user out and stop background tasks using it: $User"
  }

  return [PSCustomObject]@{
    SID  = $_identity.SID
    Path = $_profilePath
    User = $_identity.User
  }
}

function Assert-ImportProfileUnchanged {
  param (
    [object]
    $TargetProfile
  )

  $_current = Get-ImportProfile -User $TargetProfile.SID
  if ($_current.Path -ne $TargetProfile.Path -or $_current.SID -ne $TargetProfile.SID) {
    throw 'The target profile changed during the import.'
  }
}

function Get-ImportVolume {
  param (
    [string]
    $Path
  )

  $_drive = New-Object IO.DriveInfo([IO.Path]::GetPathRoot($Path))
  if (-not $_drive.IsReady -or $_drive.DriveType -ne [IO.DriveType]::Fixed -or $_drive.DriveFormat -ne 'NTFS') {
    throw "Import paths must be on ready local fixed NTFS volumes: $Path"
  }

  return $_drive
}

function Get-ImportInventory {
  param (
    [string]
    $Root,

    [switch]
    $Destination
  )

  $_files = New-Object System.Collections.ArrayList
  $_directories = New-Object System.Collections.ArrayList
  $_excluded = New-Object System.Collections.ArrayList
  $_pending = New-Object 'System.Collections.Generic.Queue[string]'
  $_pending.Enqueue($Root)

  while ($_pending.Count -gt 0) {
    $_directory = $_pending.Dequeue()
    Assert-ImportPath -Path $_directory
    $null = $_directories.Add($_directory)

    foreach ($_item in @(Get-ChildItem -LiteralPath $_directory -Force)) {
      $_attributes = $_item.Attributes
      $_relative = $_item.FullName.Substring($Root.Length).TrimStart('\')
      if (($_attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        if ($Destination -or $_item.LinkType -notin @('Junction', 'SymbolicLink')) {
          throw "Unsupported reparse point or cloud placeholder: $($_item.FullName)"
        }

        $null = $_excluded.Add([PSCustomObject]@{
            Path        = $_item.FullName
            Reason      = 'Link'
            IsDirectory = [bool]$_item.PSIsContainer
          })
        continue
      }

      if (($_attributes -band [IO.FileAttributes]::Encrypted) -ne 0 -or
        ($_attributes -band [IO.FileAttributes]::Offline) -ne 0) {
        throw "Decrypt or hydrate this item before importing: $($_item.FullName)"
      }

      if ($_item.PSIsContainer) {
        if (-not $Destination -and $_item.Name -in @('AppData', 'Application Data', 'Local Settings')) {
          $null = $_excluded.Add([PSCustomObject]@{
              Path        = $_item.FullName
              Reason      = 'ProfileData'
              IsDirectory = $true
            })
          continue
        }

        $_pending.Enqueue($_item.FullName)
        continue
      }

      if (-not $Destination -and ($_item.Extension -in @('.pst', '.ost') -or $_item.Name -match '^(NTUSER\.|UsrClass\.)')) {
        $null = $_excluded.Add([PSCustomObject]@{
            Path        = $_item.FullName
            Reason      = 'MailOrRegistry'
            IsDirectory = $false
          })
        continue
      }

      $null = $_files.Add([PSCustomObject]@{
          RelativePath = $_relative
          Length       = [long]$_item.Length
          SHA256       = $null
        })
    }
  }

  return [PSCustomObject]@{
    Files       = @($_files.ToArray())
    Directories = @($_directories.ToArray())
    Excluded    = @($_excluded.ToArray())
  }
}

function Assert-ImportPermission {
  param (
    [string]
    $Path,

    [string]
    $SID,

    [switch]
    $RequireOwner,

    [switch]
    $RequireInheritance
  )

  $_acl = Get-Acl -LiteralPath $Path
  $_rules = @($_acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
  if (@($_rules | Where-Object AccessControlType -EQ Deny).Count -gt 0) {
    throw "A deny ACL prevents conservative permission verification: $Path"
  }

  $_rights = [Security.AccessControl.FileSystemRights]0
  foreach ($_rule in $_rules) {
    if ($_rule.IdentityReference.Value -eq $SID -and
      ($_rule.PropagationFlags -band [Security.AccessControl.PropagationFlags]::InheritOnly) -eq 0) {
      $_inherits = ($_rule.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ContainerInherit) -ne 0 -and
      ($_rule.InheritanceFlags -band [Security.AccessControl.InheritanceFlags]::ObjectInherit) -ne 0
      if (-not $RequireInheritance -or $_inherits) {
        $_rights = $_rights -bor $_rule.FileSystemRights
      }
    }
  }

  $_needed = [Security.AccessControl.FileSystemRights]::Modify
  if (($_rights -band $_needed) -ne $_needed) {
    throw "The target SID lacks a verified Modify grant on '$Path'. Repair the target profile permissions first."
  }

  if ($RequireOwner -and $_acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -ne $SID) {
    throw "Owner verification failed on '$Path'."
  }
}

function Invoke-ImportPermission {
  param (
    [string]
    $Path,

    [string]
    $SID,

    [string]
    $LogPrefix
  )

  # The caller checks the entire destination tree for links before recursive icacls.
  $_commands = @(
    @('/reset', '/T', '/C', '/L', '/Q'),
    @('/setowner', "*$SID", '/T', '/C', '/L', '/Q')
  )
  for ($_index = 0; $_index -lt $_commands.Count; $_index++) {
    $_arguments = @($Path) + $_commands[$_index]
    $_process = Invoke-SafeProcess -FilePath "$env:SystemRoot\System32\icacls.exe" -ArgumentList $_arguments -AsResult
    [IO.File]::WriteAllText("$LogPrefix-$_index.log", $_process.StdOut + $_process.StdErr)
    if ($_process.TimedOut -or $_process.Cancelled -or $_process.ExitCode -ne 0) {
      throw "icacls failed on '$Path' (exit $($_process.ExitCode)); inspect $LogPrefix-$_index.log."
    }
  }
}

function Assert-ImportContent {
  param (
    [object]
    $Entry
  )

  $_current = Get-ImportInventory -Root $Entry.Source
  if ($_current.Files.Count -ne $Entry.Inventory.Files.Count -or
    $_current.Directories.Count -ne $Entry.Inventory.Directories.Count -or
    (@($_current.Files.RelativePath | Sort-Object) -join "`n") -cne
    (@($Entry.Inventory.Files.RelativePath | Sort-Object) -join "`n")) {
    throw "Source inventory changed during import: $($Entry.Source)"
  }

  foreach ($_directory in $Entry.Inventory.Directories) {
    $_relative = $_directory.Substring($Entry.Source.Length).TrimStart('\')
    $_targetDirectory = $Entry.Destination
    if ($_relative) {
      $_targetDirectory = Join-Path $Entry.Destination $_relative
    }

    Assert-ImportPath -Path $_targetDirectory
    if (-not (Test-Path -LiteralPath $_targetDirectory -PathType Container)) {
      throw "Expected destination directory is missing: $_targetDirectory"
    }
  }

  foreach ($_file in $Entry.Inventory.Files) {
    $_sourceFile = Join-Path $Entry.Source $_file.RelativePath
    $_destinationFile = Join-Path $Entry.Destination $_file.RelativePath
    foreach ($_path in @($_sourceFile, $_destinationFile)) {
      Assert-ImportPath -Path $_path
      $_item = Get-Item -LiteralPath $_path -Force
      if ($_item.PSIsContainer -or $_item.Length -ne $_file.Length -or
        (Get-FileHash -LiteralPath $_path -Algorithm SHA256).Hash -ne $_file.SHA256) {
        throw "SHA-256/length verification failed: $_path"
      }
    }
  }
}

function Write-ImportAudit {
  param (
    [object]
    $Manifest,

    [System.Collections.ArrayList]
    $Results,

    [string]
    $Directory
  )

  Assert-ImportPath -Path $Directory
  $_json = ConvertTo-Json -InputObject $Manifest -Depth 12
  [IO.File]::WriteAllText((Join-Path $Directory 'manifest.json'), $_json, (New-Object Text.UTF8Encoding($true)))
  $null = Write-OperationResultLog -Results $Results -Path (Join-Path $Directory 'operations.jsonl') `
    -ScriptName 'Import-UserProfileData' -RunId $Manifest.RunId
}

function Initialize-ImportReport {
  param (
    [string]
    $Path
  )

  $null = New-Item -ItemType Directory -Path $Path
  $_acl = New-Object Security.AccessControl.DirectorySecurity
  $_acl.SetAccessRuleProtection($true, $false)
  $_identities = @(
    'S-1-5-18',
    'S-1-5-32-544',
    [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
  )
  foreach ($_identity in @($_identities | Select-Object -Unique)) {
    $_sid = New-Object Security.Principal.SecurityIdentifier($_identity)
    $_rule = New-Object Security.AccessControl.FileSystemAccessRule(
      $_sid,
      [Security.AccessControl.FileSystemRights]::FullControl,
      [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
      [Security.AccessControl.PropagationFlags]::None,
      [Security.AccessControl.AccessControlType]::Allow
    )
    $_acl.AddAccessRule($_rule)
  }

  Set-Acl -LiteralPath $Path -AclObject $_acl
}

try {
  if (-not (Test-Elevation)) {
    throw 'Run this script in an elevated PowerShell session.'
  }

  $_source = Resolve-ImportPath -Path $SourcePath
  if (-not (Test-Path -LiteralPath $_source -PathType Container)) {
    throw "Extract the backup first; SourcePath must be a directory: $_source"
  }

  $_profile = Get-ImportProfile -User $TargetUser
  $_reportRoot = Resolve-ImportPath -Path $ReportDirectory
  foreach ($_pair in @(@($_source, $_profile.Path), @($_source, $_reportRoot), @($_profile.Path, $_reportRoot))) {
    if ((Test-ImportPathWithin -Path $_pair[0] -Root $_pair[1]) -or
      (Test-ImportPathWithin -Path $_pair[1] -Root $_pair[0])) {
      throw 'Source, target profile, and report directories must not overlap.'
    }
  }

  $null = Get-ImportVolume -Path $_source
  $_volume = Get-ImportVolume -Path $_profile.Path
  $null = Get-ImportVolume -Path $_reportRoot
  Assert-ImportPermission -Path $_profile.Path -SID $_profile.SID -RequireInheritance

  $_mappings = New-Object System.Collections.ArrayList
  foreach ($_name in $Folder) {
    $null = $_mappings.Add([PSCustomObject]@{
        Source   = $_name
        Target   = $_name
        Required = $PSBoundParameters.ContainsKey('Folder')
      })
  }
  foreach ($_name in @($FolderMap.Keys | Sort-Object)) {
    $null = $_mappings.Add([PSCustomObject]@{
        Source   = [string]$_name
        Target   = [string]$FolderMap[$_name]
        Required = $true
      })
  }

  $_bytes = [long]0
  $_fileCount = 0
  foreach ($_mapping in $_mappings) {
    $_from = Join-ImportDataPath -Root $_source -RelativePath $_mapping.Source
    $_to = Join-ImportDataPath -Root $_profile.Path -RelativePath $_mapping.Target
    Assert-ImportPath -Path $_from
    Assert-ImportPath -Path $_to
    if (-not (Test-Path -LiteralPath $_from -PathType Container)) {
      if ($_mapping.Required) {
        throw "Selected source directory does not exist: $_from"
      }

      Add-OperationResult -Results $_results -Target $_from -Action 'Plan' -Status 'Skipped' -Detail 'Default folder is absent.'
      continue
    }

    foreach ($_existing in $_plan) {
      if ((Test-ImportPathWithin -Path $_from -Root $_existing.Source) -or
        (Test-ImportPathWithin -Path $_existing.Source -Root $_from) -or
        (Test-ImportPathWithin -Path $_to -Root $_existing.Destination) -or
        (Test-ImportPathWithin -Path $_existing.Destination -Root $_to)) {
        throw 'Selected source and destination mappings must be distinct and non-overlapping.'
      }
    }

    Write-Log -Message "Inventory and hash: $_from -> $_to" -Color Cyan
    $_inventory = Get-ImportInventory -Root $_from
    if (Test-Path -LiteralPath $_to) {
      if (-not (Test-Path -LiteralPath $_to -PathType Container)) {
        throw "A file occupies the destination directory: $_to"
      }

      $null = Get-ImportInventory -Root $_to -Destination
    }

    $_parent = Split-Path -Path $_to -Parent
    while (-not (Test-Path -LiteralPath $_parent)) {
      $_parent = Split-Path -Path $_parent -Parent
    }
    Assert-ImportPermission -Path $_parent -SID $_profile.SID -RequireInheritance

    foreach ($_directory in $_inventory.Directories) {
      $_relative = $_directory.Substring($_from.Length).TrimStart('\')
      $_targetDirectory = $_to
      if ($_relative) {
        $_targetDirectory = Join-Path $_to $_relative
      }

      if (Test-Path -LiteralPath $_targetDirectory -PathType Leaf) {
        throw "A destination file conflicts with a source directory: $_targetDirectory"
      }
    }

    foreach ($_file in $_inventory.Files) {
      $_sourceFile = Join-Path $_from $_file.RelativePath
      $_file.SHA256 = (Get-FileHash -LiteralPath $_sourceFile -Algorithm SHA256).Hash
      $_targetFile = Join-Path $_to $_file.RelativePath
      if (Test-Path -LiteralPath $_targetFile) {
        $_targetItem = Get-Item -LiteralPath $_targetFile -Force
        if ($_targetItem.PSIsContainer) {
          throw "A destination directory conflicts with a source file: $_targetFile"
        }

        if (-not $Overwrite -and ($_targetItem.Length -ne $_file.Length -or
            (Get-FileHash -LiteralPath $_targetFile -Algorithm SHA256).Hash -ne $_file.SHA256)) {
          throw "Destination content conflict: $_targetFile. Review it before using -Overwrite."
        }
      }

      $_bytes += $_file.Length
      $_fileCount++
    }

    $null = $_plan.Add([PSCustomObject]@{
        Source      = $_from
        Destination = $_to
        Inventory   = $_inventory
        Status      = 'Planned'
        ExitCode    = $null
        CopyLog     = $null
      })
    Add-OperationResult -Results $_results -Target $_to -Source $_from -Action 'Plan' -Status 'Completed' -Property @{
      Files    = $_inventory.Files.Count
      Excluded = $_inventory.Excluded.Count
    }
  }

  if ($_plan.Count -eq 0) {
    throw 'No selected source folders exist. Point SourcePath at the extracted old user directory.'
  }
  if ($_volume.AvailableFreeSpace -lt ($_bytes + $ReserveBytes)) {
    throw "Insufficient target disk space: need at least $($_bytes + $ReserveBytes) free bytes."
  }

  $_manifest = [ordered]@{
    SchemaVersion = 1
    RunId         = $_runId
    StartedAtUtc  = $_startedAt
    FinishedAtUtc = $null
    Actor         = $_actor
    Machine       = $env:COMPUTERNAME
    SourcePath    = $_source
    Target        = $_profile
    Files         = $_fileCount
    Bytes         = $_bytes
    Overwrite     = [bool]$Overwrite
    RemoveSource  = [bool]$RemoveSource
    Status        = 'Planned'
    Error         = $null
    Plan          = @($_plan.ToArray())
  }

  if ($PSCmdlet.ShouldProcess($_profile.Path, "Import $_fileCount files and reset permissions/ownership in $($_plan.Count) selected data folders")) {
    Assert-ImportProfileUnchanged -TargetProfile $_profile
    $_runDirectory = Join-Path $_reportRoot $_runId
    Assert-ImportPath -Path $_runDirectory
    Initialize-ImportReport -Path $_runDirectory
    $_manifestPath = Join-Path $_runDirectory 'manifest.json'
    $_manifest.Status = 'Started'
    Write-ImportAudit -Manifest $_manifest -Results $_results -Directory $_runDirectory

    $_index = 0
    foreach ($_entry in $_plan) {
      Assert-ImportProfileUnchanged -TargetProfile $_profile
      Assert-ImportPath -Path $_entry.Source
      Assert-ImportPath -Path $_entry.Destination
      # Recheck destination links immediately before invoking a recursive native tool.
      if (Test-Path -LiteralPath $_entry.Destination) {
        $null = Get-ImportInventory -Root $_entry.Destination -Destination
      }

      $_index++
      $_entry.CopyLog = Join-Path $_runDirectory "robocopy-$_index.log"
      $_arguments = @(
        $_entry.Source,
        $_entry.Destination,
        '/E',
        '/COPY:DAT',
        '/DCOPY:DAT',
        '/XJ',
        '/XJD',
        '/XJF',
        '/R:1',
        '/W:1',
        '/IS',
        '/IT',
        '/IM',
        '/NP',
        "/UNILOG:$($_entry.CopyLog)",
        '/XF',
        '*.pst',
        '*.ost',
        'NTUSER.*',
        'UsrClass.*'
      )
      $_arguments += @($_entry.Inventory.Excluded | Where-Object { -not $_.IsDirectory } | ForEach-Object { $_.Path })
      $_excludedDirectories = @($_entry.Inventory.Excluded | Where-Object IsDirectory | ForEach-Object { $_.Path })
      if ($_excludedDirectories.Count -gt 0) {
        $_arguments += @('/XD') + $_excludedDirectories
      }

      # Keep argument growth bounded rather than silently dropping exclusions.
      if (($_arguments -join ' ').Length -gt 24000) {
        throw 'Too many excluded paths for a safe robocopy command; split the selected folders into smaller imports.'
      }

      Write-Log -Message "Copying $($_entry.Source)" -Color Cyan
      $_entry.Status = 'Copying'
      Write-ImportAudit -Manifest $_manifest -Results $_results -Directory $_runDirectory
      $_copy = Invoke-SafeProcess -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList $_arguments -AsResult
      $_entry.ExitCode = $_copy.ExitCode
      $_meaning = Convert-RobocopyExitCode -ExitCode $_copy.ExitCode
      if ($_copy.TimedOut -or $_copy.Cancelled -or $_copy.ExitCode -lt 0 -or $_copy.ExitCode -ge 8) {
        throw "Robocopy failed: $_meaning (exit $($_copy.ExitCode)); see $($_entry.CopyLog)."
      }

      $_destinationInventory = Get-ImportInventory -Root $_entry.Destination -Destination
      Assert-ImportContent -Entry $_entry
      Assert-ImportProfileUnchanged -TargetProfile $_profile
      Write-Log -Message "Applying target profile permissions: $($_entry.Destination)" -Color Cyan
      Invoke-ImportPermission -Path $_entry.Destination -SID $_profile.SID -LogPrefix (Join-Path $_runDirectory "permissions-$_index")
      foreach ($_path in $_destinationInventory.Directories) {
        Assert-ImportPermission -Path $_path -SID $_profile.SID -RequireOwner -RequireInheritance
      }
      foreach ($_file in $_destinationInventory.Files) {
        Assert-ImportPermission -Path (Join-Path $_entry.Destination $_file.RelativePath) -SID $_profile.SID -RequireOwner
      }

      $_entry.Status = 'Verified'
      Add-OperationResult -Results $_results -Target $_entry.Destination -Source $_entry.Source -Action 'Import' -Status 'Completed' `
        -Changed $true -Detail 'SHA-256, lengths, owner, and Modify permissions verified.' -Property @{
        Files        = $_entry.Inventory.Files.Count
        RobocopyCode = $_entry.ExitCode
        RobocopyText = $_meaning
        CopyLog      = $_entry.CopyLog
      }
      Write-ImportAudit -Manifest $_manifest -Results $_results -Directory $_runDirectory
    }

    $_manifest.Status = 'Verified'
    Write-ImportAudit -Manifest $_manifest -Results $_results -Directory $_runDirectory

    if ($RemoveSource -and $PSCmdlet.ShouldProcess($_source, 'Delete only verified imported source files and empty selected directories')) {
      Assert-ImportProfileUnchanged -TargetProfile $_profile
      foreach ($_entry in $_plan) {
        Assert-ImportContent -Entry $_entry
      }

      foreach ($_entry in $_plan) {
        foreach ($_file in $_entry.Inventory.Files) {
          $_sourceFile = Join-Path $_entry.Source $_file.RelativePath
          $_targetFile = Join-Path $_entry.Destination $_file.RelativePath
          if (-not (Test-ImportPathWithin -Path $_sourceFile -Root $_source) -or
            -not (Test-ImportPathWithin -Path $_targetFile -Root $_profile.Path)) {
            throw 'Cleanup path is outside the validated source or target.'
          }

          Assert-ImportPath -Path $_sourceFile
          Assert-ImportPath -Path $_targetFile
          if ((Get-FileHash -LiteralPath $_sourceFile -Algorithm SHA256).Hash -ne $_file.SHA256 -or
            (Get-FileHash -LiteralPath $_targetFile -Algorithm SHA256).Hash -ne $_file.SHA256) {
            throw "Content changed before cleanup: $_sourceFile"
          }

          [IO.File]::Delete($_sourceFile)
          Add-OperationResult -Results $_results -Target $_sourceFile -Action 'Cleanup' -Status 'Completed' -Changed $true
        }

        foreach ($_directory in @($_entry.Inventory.Directories | Sort-Object Length -Descending)) {
          if (-not (Test-ImportPathWithin -Path $_directory -Root $_source) -or $_directory -eq $_source) {
            throw 'Cleanup directory is outside the selected staging subtree.'
          }

          Assert-ImportPath -Path $_directory
          if (@(Get-ChildItem -LiteralPath $_directory -Force).Count -eq 0) {
            [IO.Directory]::Delete($_directory, $false)
          }
        }
      }
    }

    $_status = 'Completed'
  }
  elseif ($WhatIfPreference) {
    $_status = 'Preview'
  }
}
catch {
  $_failure = $_.Exception.Message
  $_status = 'Failed'
  Write-Log -Message "Profile import failed: $_failure" -Color Red
}

$_summary = @{
  Target       = $TargetUser
  Source       = $SourcePath
  Action       = 'ImportProfileData'
  Status       = $_status
  ErrorMessage = $_failure
  Property     = @{
    RunId        = $_runId
    ManifestPath = $_manifestPath
    StagingPath  = $_source
    Plan         = @($_plan.ToArray())
  }
}
Add-OperationResult -Results $_results @_summary

if ($_runDirectory) {
  try {
    $_manifest.Status = $_status
    $_manifest.Error = $_failure
    $_manifest.FinishedAtUtc = [DateTime]::UtcNow.ToString('o')
    Write-ImportAudit -Manifest $_manifest -Results $_results -Directory $_runDirectory
  }
  catch {
    $_failure = "Could not finalize the audit: $($_.Exception.Message)"
    $_status = 'Failed'
    $_summaryResult = $_results | Where-Object Action -EQ 'ImportProfileData' | Select-Object -Last 1
    $_summaryResult.Status = 'Failed'
    $_summaryResult.Error = $_failure
    Add-OperationResult -Results $_results -Target $_runDirectory -Action 'Audit' -Status 'Failed' -ErrorMessage $_failure
    Write-Log -Message $_failure -Color Red
  }
}

Write-Log -Message "Profile import: $_status | Report: $_manifestPath | Staging: $_source" -Color Cyan
if ($PassThru -or $WhatIfPreference) {
  $_results
}
if ($_failure) {
  exit 1
}
