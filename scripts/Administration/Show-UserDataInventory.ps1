#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.4.0' }

<#
.SYNOPSIS
  Inventories a user's data locations and generates a migration plan.

.DESCRIPTION
  Reads the selected user's known folders, public folders and shortcuts, home
  drive, network drives, outside-profile data, and path lengths. It can consume a
  Show-OutlookAccount report to add PST/OST files and signatures to the
  manual-copy plan.

  The home drive is read from the user's signed-in session (Volatile
  Environment) and, for domain accounts, from Active Directory (homeDrive,
  homeDirectory), which also names the logon script. Persistent mappings come
  from the user's registry. Drives connected in the current session are read only
  when the user runs the inventory without elevation, because other sessions
  cannot see them. A home drive is mapped to -HomeDriveTarget and staged as
  HomeDrive; other drives are listed as decisions.

  The script is read-only. It runs in the user's session without elevation for
  the current user. Another local profile requires administrator rights; its
  registry hive is loaded only when the profile is not in use.

  The generated plan lists the folders to download from ABB and their staging
  names, the folder mappings for Import-UserProfileData, manual copies, notes about
  unusual folder locations, and open decisions. Known folders with data are
  mapped from where they actually are, including redirected folders and folders
  nested in other known folders; empty ones are noted and -Folder selects a
  subset. Public folders are listed as decisions unless named with
  -IncludePublic. Program shortcuts on the public desktop are noted as recreated
  by the new computer's installers; its other files are listed for a manual
  copy. Other data outside the profile is reported but only mapped when named
  with -Include. System, vendor, and winkit-managed folders are not scanned
  unless named with -Include.

  Unless -DryRun is supplied, a unique run directory receives manifest.json with
  the full inventory and folder-map.json for Import-UserProfileData -FolderMapPath.

.PARAMETER User
  User to inventory. Defaults to the current user. Accepts an account name or
  SID. Another local profile requires administrator rights.

.PARAMETER TargetRoot
  New profile folder, for example C:\Users\newuser. Every mapped folder, the home
  drive included, is checked for destination paths that are too long at its
  mapped target below this folder. A folder below the same profiles folder also
  names the new user for the suggested import call, in the old user's domain.
  Without it, the old profile folder stands in, with a warning.

.PARAMETER OutlookReport
  Show-OutlookAccount JSON report or a directory containing reports. Without it,
  the plan states that Outlook was not inventoried.

.PARAMETER OutlookReportMaxAge
  Maximum accepted report age in days. Defaults to 7.

.PARAMETER Include
  Local directories outside the profile to migrate, for example C:\asfibu. Each is
  scanned even when it is excluded by default, downloaded as its own folder, and
  mapped to Documents\<folder name>. Edit folder-map.json to change the target.

.PARAMETER Folder
  Known folders to map, for example Desktop, Documents. Defaults to every known
  folder that holds data. Same names as Import-UserProfileData -Folder.

.PARAMETER IncludePublic
  Public folders to import below the matching user folder, for example
  Documents, Downloads. Public folders are shared by every user of the old
  computer, so by default they are listed as decisions only. The public desktop
  is never imported.

.PARAMETER StagingRoot
  Local directory that will hold every extracted download, used for the staging
  paths in the plan and the suggested import call. Defaults to
  C:\Managed\Migration on the system drive.

.PARAMETER HomeDriveTarget
  Target folder, relative to the new profile, for the user's home drive. Defaults
  to Private.

.PARAMETER MeasureNetwork
  Measure size and file counts on the home drive and other network drives through
  their UNC paths. Off by default because it can be slow.

.PARAMETER ReportDirectory
  Parent for unique inventory reports. Defaults to
  ProgramData\winkit\UserDataInventory. No report is written during -DryRun.

.PARAMETER DryRun
  Inventory and print the plan without writing a report.

.PARAMETER PassThru
  Return the inventory object.

.EXAMPLE
  PS> .\Show-UserDataInventory.ps1 -DryRun
  Inventory the current user and print the migration plan.

.EXAMPLE
  PS> .\Show-UserDataInventory.ps1 -User 'CONTOSO\User1' -TargetRoot 'C:\Users\User1' -OutlookReport 'C:\Reports'
  Inventory another local profile and consume the newest matching Outlook report.

.EXAMPLE
  PS> .\Show-UserDataInventory.ps1 -User 'CONTOSO\User1' -Folder Desktop, Documents, Downloads -IncludePublic Documents
  Map only three known folders and the public documents.

.EXAMPLE
  PS> .\Show-UserDataInventory.ps1 -User 'CONTOSO\User1' -Include 'C:\Business' -StagingRoot 'C:\Managed\Migration'
  Also migrate C:\Business and write folder-map.json for Import-UserProfileData.

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - no Outlook COM or desktop session is required.
  SYSTEM-account execution: not supported; inventory a real user profile.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [string]
  $User,

  [string]
  $TargetRoot,

  [string]
  $OutlookReport,

  [ValidateRange(1, 365)]
  [int]
  $OutlookReportMaxAge = 7,

  [string[]]
  $Include = @(),

  [ValidateSet(
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
  )]
  [string[]]
  $Folder = @(),

  [ValidateSet(
    'Documents',
    'Downloads',
    'Pictures',
    'Music',
    'Videos'
  )]
  [string[]]
  $IncludePublic = @(),

  [string]
  $StagingRoot = (Join-Path $env:SystemDrive 'Managed\Migration'),

  [ValidateNotNullOrEmpty()]
  [string]
  $HomeDriveTarget = 'Private',

  [switch]
  $MeasureNetwork,

  [string]
  $ReportDirectory = (Join-Path $env:ProgramData 'winkit\UserDataInventory'),

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.4.0 -Force

$ErrorActionPreference = 'Stop'
if ($DryRun) {
  $WhatIfPreference = $true
}

if ([string]::IsNullOrWhiteSpace($User)) {
  $User = [Security.Principal.WindowsIdentity]::GetCurrent().Name
}

function ConvertTo-DataInventoryJson {
  param (
    [object]
    $InputObject,

    [int]
    $Depth = 20
  )

  $_json = ConvertTo-Json -InputObject $InputObject -Depth $Depth -ErrorAction Stop

  # Windows PowerShell 5.1 escapes ' < > & as \u00XX. They are valid JSON literals, so restore them for
  # readability. An escaped backslash before the sequence (\\u0027) is literal text and stays untouched.
  $_pattern = '(?<!\\)((?:\\\\)*)\\u00(27|3[cCeE]|26)'
  $_evaluator = {
    param (
      $Match
    )

    return $Match.Groups[1].Value + [string][char][Convert]::ToInt32($Match.Groups[2].Value, 16)
  }

  return [regex]::Replace($_json, $_pattern, $_evaluator)
}

function Test-DataInventoryAdministrator {
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-DataInventoryAccount {
  param (
    [string]
    $Account
  )

  if ($Account -match '^S-1-5-') {
    $sid = New-Object Security.Principal.SecurityIdentifier($Account)
    $name = $sid.Translate([Security.Principal.NTAccount]).Value
  }
  else {
    $ntAccount = New-Object Security.Principal.NTAccount($Account)
    $sid = $ntAccount.Translate([Security.Principal.SecurityIdentifier])
    $name = $ntAccount.Value
  }

  return [PSCustomObject]@{
    Name = $name
    SID  = $sid
  }
}

function Resolve-DataInventoryUser {
  param (
    [string]
    $Account
  )

  $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $accountInfo = Resolve-DataInventoryAccount -Account $Account
  $isCurrent = $accountInfo.SID.Value -eq $currentIdentity.User.Value

  if ($isCurrent) {
    return [PSCustomObject]@{
      Name        = $currentIdentity.Name
      SID         = $currentIdentity.User.Value
      ProfilePath = $env:USERPROFILE
      Registry    = 'HKCU:'
      LoadedHive  = $null
      IsCurrent   = $true
      IsAdmin     = (Test-DataInventoryAdministrator)
    }
  }

  if (-not (Test-DataInventoryAdministrator)) {
    throw "Run elevated to inventory another user's profile: $Account"
  }

  $_profile = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($accountInfo.SID.Value)'" -ErrorAction Stop
  if ($null -eq $_profile -or [string]::IsNullOrWhiteSpace($_profile.LocalPath)) {
    throw "No local profile was found for $Account."
  }

  $loadedHive = $null
  $registry = "Registry::HKEY_USERS\$($accountInfo.SID.Value)"
  if (-not $_profile.Loaded) {
    $loadedHive = 'winkit-inventory-' + [guid]::NewGuid().ToString('N')
    $hiveResult = Invoke-SafeProcess -FilePath "$env:SystemRoot\System32\reg.exe" -ArgumentList @('load', "HKU\$loadedHive", (Join-Path $_profile.LocalPath 'NTUSER.DAT')) -AsResult -ErrorAction Stop
    if ($hiveResult.ExitCode -ne 0 -or $hiveResult.TimedOut -or $hiveResult.Cancelled) {
      throw "Could not load the profile registry hive for $Account."
    }
    $registry = "Registry::HKEY_USERS\$loadedHive"
  }

  return [PSCustomObject]@{
    Name        = $accountInfo.Name
    SID         = $accountInfo.SID.Value
    ProfilePath = $_profile.LocalPath
    Registry    = $registry
    LoadedHive  = $loadedHive
    IsCurrent   = $false
    IsAdmin     = $true
  }
}

function Test-DataInventoryPathWithin {
  param (
    [string]
    $Path,

    [string]
    $Root
  )

  # True only for a strict descendant; equal paths are not nested.
  return $Path.StartsWith($Root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Format-DataInventoryLiteral {
  param (
    [string]
    $Value
  )

  return "'" + $Value.Replace("'", "''") + "'"
}

function Expand-DataInventoryPath {
  param (
    [string]
    $Path,

    [string]
    $ProfilePath
  )

  $_expanded = $Path
  $_expanded = $_expanded.Replace('%USERPROFILE%', $ProfilePath)
  $_expanded = $_expanded.Replace('%APPDATA%', (Join-Path $ProfilePath 'AppData\Roaming'))
  $_expanded = $_expanded.Replace('%LOCALAPPDATA%', (Join-Path $ProfilePath 'AppData\Local'))
  return [Environment]::ExpandEnvironmentVariables($_expanded)
}

function Get-DataInventoryTreeSummary {
  param (
    [string]
    $Path,

    [string[]]
    $ExcludeDirectory = @(),

    [int]
    $ReparseItemLimit = 50
  )

  $result = [PSCustomObject]@{
    Path          = $Path
    Exists        = $false
    Files         = 0
    Directories   = 0
    Bytes         = [long]0
    Newest        = $null
    ReparsePoints = 0
    ReparseItems  = @()
    Offline       = 0
    Encrypted     = 0
    Error         = $null
  }

  if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
    return $result
  }

  $result.Exists = $true
  $_newest = $null
  $_reparseItems = New-Object System.Collections.ArrayList
  $_pending = New-Object 'System.Collections.Generic.Queue[string]'
  $_pending.Enqueue($Path)

  # Walk the tree without following links; excluded (nested, separately counted) folders are skipped.
  while ($_pending.Count -gt 0) {
    $_directory = $_pending.Dequeue()
    try {
      $_items = @(Get-ChildItem -LiteralPath $_directory -Force -ErrorAction Stop)
    }
    catch {
      if ($null -eq $result.Error) {
        $result.Error = $_.Exception.Message
      }

      continue
    }

    foreach ($_item in $_items) {
      # Explorer metadata is never imported, so a folder holding only desktop.ini is empty.
      if (-not $_item.PSIsContainer -and $_item.Name -in @('desktop.ini', 'Thumbs.db')) {
        continue
      }

      if ($_item.PSIsContainer) {
        $_isExcluded = @($ExcludeDirectory | Where-Object { $_.TrimEnd('\').Equals($_item.FullName, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
        if ($_isExcluded) {
          continue
        }
      }

      if ($_item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        $result.ReparsePoints++
        if ($_reparseItems.Count -lt $ReparseItemLimit) {
          $null = $_reparseItems.Add([PSCustomObject]@{
              Path     = $_item.FullName
              LinkType = [string]$_item.LinkType
              Target   = (@($_item.Target) | Where-Object { $_ }) -join '; '
            })
        }

        continue
      }

      if ($_item.PSIsContainer) {
        $result.Directories++
        $_pending.Enqueue($_item.FullName)
      }
      else {
        $result.Files++
        $result.Bytes += [long]$_item.Length
      }

      if ($null -eq $_newest -or $_item.LastWriteTimeUtc -gt $_newest) {
        $_newest = $_item.LastWriteTimeUtc
      }

      if ($_item.Attributes -band [IO.FileAttributes]::Offline) {
        $result.Offline++
      }

      if ($_item.Attributes -band [IO.FileAttributes]::Encrypted) {
        $result.Encrypted++
      }
    }
  }

  # ISO 8601 text; Windows PowerShell 5.1 would otherwise serialize dates as /Date(...)/.
  if ($null -ne $_newest) {
    $result.Newest = $_newest.ToString('o')
  }

  $result.ReparseItems = @($_reparseItems.ToArray())
  return $result
}

function Get-DataInventoryKnownFolder {
  param (
    [string]
    $RegistryRoot,

    [string]
    $ProfilePath
  )

  $_definitions = @(
    @{ Name = 'Desktop'; Value = 'Desktop' }
    @{ Name = 'Documents'; Value = 'Personal' }
    @{ Name = 'Downloads'; Value = '{374DE290-123F-4565-9164-39C4925E467B}' }
    @{ Name = 'Pictures'; Value = 'My Pictures' }
    @{ Name = 'Music'; Value = 'My Music' }
    @{ Name = 'Videos'; Value = 'My Video' }
    @{ Name = 'Favorites'; Value = 'Favorites' }
    @{ Name = 'Contacts'; Value = 'Contacts' }
    @{ Name = 'Links'; Value = 'Links' }
    @{ Name = 'Saved Games'; Value = '{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}' }
    @{ Name = 'Searches'; Value = '{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}' }
  )

  $_key = Join-Path $RegistryRoot 'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
  $_values = Get-ItemProperty -LiteralPath $_key -ErrorAction SilentlyContinue

  # Resolve every location first so a known folder stored inside another is counted only once.
  $_located = New-Object System.Collections.ArrayList
  foreach ($_definition in $_definitions) {
    $_raw = $null
    if ($null -ne $_values) {
      $_property = $_values.PSObject.Properties | Where-Object { $_.Name -ieq $_definition.Value } | Select-Object -First 1
      if ($null -ne $_property) {
        $_raw = [string]$_property.Value
      }
    }
    if ([string]::IsNullOrWhiteSpace($_raw)) {
      continue
    }

    $null = $_located.Add([PSCustomObject]@{
        Name = $_definition.Name
        Raw  = $_raw
        Path = (Expand-DataInventoryPath -Path $_raw -ProfilePath $ProfilePath).TrimEnd('\')
      })
  }

  $_result = New-Object System.Collections.ArrayList
  foreach ($_location in $_located) {
    $_path = $_location.Path
    $_raw = $_location.Raw
    $_nested = @($_located | Where-Object { Test-DataInventoryPathWithin -Path $_.Path -Root $_path })
    $_summary = Get-DataInventoryTreeSummary -Path $_path -ExcludeDirectory @($_nested | ForEach-Object { $_.Path })
    $_isUnc = $_path.StartsWith('\\')
    $_isOutsideProfile = -not $_path.StartsWith($ProfilePath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
    $_isOneDrive = $_path -match '(^|\\)OneDrive( - |\\)'
    $_isReparse = $false
    if (Test-Path -LiteralPath $_path -ErrorAction SilentlyContinue) {
      $_isReparse = [bool]((Get-Item -LiteralPath $_path -Force -ErrorAction SilentlyContinue).Attributes -band [IO.FileAttributes]::ReparsePoint)
    }

    $null = $_result.Add([PSCustomObject]@{
        Name             = $_location.Name
        Path             = $_path
        Raw              = $_raw
        Exists           = $_summary.Exists
        Files            = $_summary.Files
        Directories      = $_summary.Directories
        Bytes            = $_summary.Bytes
        Newest           = $_summary.Newest
        NestedFolders    = @($_nested | ForEach-Object { $_.Name })
        ReparsePoints    = $_summary.ReparsePoints
        ReparseItems     = $_summary.ReparseItems
        Offline          = $_summary.Offline
        Encrypted        = $_summary.Encrypted
        IsUnc            = $_isUnc
        IsOutsideProfile = $_isOutsideProfile
        IsOneDrive       = $_isOneDrive
        IsReparse        = $_isReparse
        Error            = $_summary.Error
      })
  }

  return @($_result.ToArray())
}

function Get-DataInventoryShortcut {
  param (
    [string]
    $Path
  )

  $_target = $null
  try {
    if ([IO.Path]::GetExtension($Path) -eq '.url') {
      $_content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
      if ($_content -match '(?im)^\s*URL\s*=\s*(.+?)\s*$') {
        $_target = $Matches[1]
      }
    }
    elseif ([IO.Path]::GetExtension($Path) -eq '.lnk') {
      $_shell = New-Object -ComObject WScript.Shell
      try {
        $_target = [string]($_shell.CreateShortcut($Path).TargetPath)
      }
      finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($_shell)
      }
    }
  }
  catch {
    return [PSCustomObject]@{
      Path      = $Path
      Target    = $null
      IsNetwork = $false
      Error     = $_.Exception.Message
    }
  }

  $_isNetwork = $false
  if (-not [string]::IsNullOrWhiteSpace($_target)) {
    if ($_target.StartsWith('\\')) {
      $_isNetwork = $true
    }
    elseif ($_target -match '^[A-Za-z]:\\') {
      try {
        $_isNetwork = (New-Object IO.DriveInfo([IO.Path]::GetPathRoot($_target))).DriveType -eq [IO.DriveType]::Network
      }
      catch {
        $_isNetwork = $false
      }
    }
  }

  return [PSCustomObject]@{
    Path      = $Path
    Target    = $_target
    IsNetwork = $_isNetwork
    Error     = $null
  }
}

function Get-DataInventoryPublicFolder {
  param (
    [string]
    $PublicRoot,

    [bool]
    $IncludePermissions
  )

  $_result = New-Object System.Collections.ArrayList
  foreach ($_name in @('Desktop', 'Documents', 'Downloads', 'Pictures', 'Music', 'Videos')) {
    $_path = Join-Path $PublicRoot $_name
    if (-not (Test-Path -LiteralPath $_path -ErrorAction SilentlyContinue)) {
      continue
    }

    $_summary = Get-DataInventoryTreeSummary -Path $_path
    $_files = @(Get-ChildItem -LiteralPath $_path -Recurse -File -Force -ErrorAction SilentlyContinue)
    $_shortcuts = @($_files | Where-Object { $_.Extension -in @('.lnk', '.url') } | ForEach-Object {
        Get-DataInventoryShortcut -Path $_.FullName
      })
    $_permissions = $null
    if ($IncludePermissions) {
      try {
        $_permissions = (Get-Acl -LiteralPath $_path -ErrorAction Stop).AccessToString
      }
      catch {
        $_permissions = $_.Exception.Message
      }
    }

    $null = $_result.Add([PSCustomObject]@{
        Name         = $_name
        Path         = $_path
        Files        = $_summary.Files
        Directories  = $_summary.Directories
        Bytes        = $_summary.Bytes
        Newest       = $_summary.Newest
        ReparseItems = $_summary.ReparseItems
        Shortcuts    = @($_shortcuts)
        Permissions  = $_permissions
        Error        = $_summary.Error
      })
  }

  return @($_result.ToArray())
}

function Get-DataInventoryDirectoryAccount {
  param (
    [string]
    $Sid
  )

  # ADSI needs no RSAT module. The client timeout keeps an unreachable domain from stalling the run.
  $_result = [PSCustomObject]@{
    HomeDrive     = $null
    HomeDirectory = $null
    ScriptPath    = $null
    Error         = $null
  }

  try {
    $_searcher = New-Object DirectoryServices.DirectorySearcher
    $_searcher.Filter = "(&(objectCategory=person)(objectClass=user)(objectSid=$Sid))"
    $_searcher.ClientTimeout = [TimeSpan]::FromSeconds(15)
    $_searcher.ServerTimeLimit = [TimeSpan]::FromSeconds(15)
    foreach ($_property in @('homeDrive', 'homeDirectory', 'scriptPath')) {
      $null = $_searcher.PropertiesToLoad.Add($_property)
    }

    $_found = $_searcher.FindOne()
    if ($null -eq $_found) {
      $_result.Error = "No directory account was found for $Sid."
      return $_result
    }

    $_result.HomeDrive = [string](@($_found.Properties['homedrive']) | Select-Object -First 1)
    $_result.HomeDirectory = [string](@($_found.Properties['homedirectory']) | Select-Object -First 1)
    $_result.ScriptPath = [string](@($_found.Properties['scriptpath']) | Select-Object -First 1)
  }
  catch {
    $_result.Error = $_.Exception.Message
  }

  return $_result
}

function Get-DataInventoryNetworkDrive {
  param (
    [string]
    $RegistryRoot,

    [object]
    $InventoryUser,

    [bool]
    $Measure
  )

  $_drives = New-Object System.Collections.ArrayList
  $_notes = New-Object System.Collections.ArrayList

  # 1. Home drive from the user's signed-in session; present only while the user is signed in.
  $_volatile = Get-ItemProperty -LiteralPath (Join-Path $RegistryRoot 'Volatile Environment') -ErrorAction SilentlyContinue
  if ($null -ne $_volatile -and -not [string]::IsNullOrWhiteSpace([string]$_volatile.HOMESHARE)) {
    $null = $_drives.Add([PSCustomObject]@{
        Letter = ([string]$_volatile.HOMEDRIVE).TrimEnd(':')
        Target = ([string]$_volatile.HOMESHARE).TrimEnd('\')
        Source = 'HomeDrive'
        Origin = 'Volatile Environment'
      })
  }

  # 2. Home drive and logon script from Active Directory, for domain accounts only.
  $_isDomainAccount = $InventoryUser.SID -like 'S-1-5-21-*' -and
  -not ([string]$InventoryUser.Name).StartsWith("$env:COMPUTERNAME\", [StringComparison]::OrdinalIgnoreCase)
  if ($_isDomainAccount) {
    $_account = Get-DataInventoryDirectoryAccount -Sid $InventoryUser.SID
    if ($_account.Error) {
      $null = $_notes.Add("Active Directory was not read for the home drive and logon script: $($_account.Error)")
    }

    $_homeDirectory = ([string]$_account.HomeDirectory).TrimEnd('\')
    $_known = @($_drives | Where-Object { $_.Target.Equals($_homeDirectory, [StringComparison]::OrdinalIgnoreCase) })
    if ($_homeDirectory.StartsWith('\\') -and $_known.Count -eq 0) {
      $null = $_drives.Add([PSCustomObject]@{
          Letter = ([string]$_account.HomeDrive).TrimEnd(':')
          Target = $_homeDirectory
          Source = 'HomeDrive'
          Origin = 'Active Directory'
        })
    }

    if (-not [string]::IsNullOrWhiteSpace($_account.ScriptPath)) {
      $null = $_drives.Add([PSCustomObject]@{
          Letter = $null
          Target = [string]$_account.ScriptPath
          Source = 'LogonScript'
          Origin = 'Active Directory'
        })
    }
  }

  # 3. Persistent mappings stored in the user's hive.
  $_key = Join-Path $RegistryRoot 'Network'
  if (Test-Path -LiteralPath $_key -ErrorAction SilentlyContinue) {
    foreach ($_child in @(Get-ChildItem -LiteralPath $_key -ErrorAction SilentlyContinue)) {
      $_properties = Get-ItemProperty -LiteralPath $_child.PSPath -ErrorAction SilentlyContinue
      if ($null -eq $_properties -or [string]::IsNullOrWhiteSpace([string]$_properties.RemotePath)) {
        continue
      }

      $null = $_drives.Add([PSCustomObject]@{
          Letter = $_child.PSChildName
          Target = ([string]$_properties.RemotePath).TrimEnd('\')
          Source = 'Persistent'
          Origin = 'Registry'
        })
    }
  }

  # 4. Connected drives belong to this process's logon session; they are the user's only when the user
  # runs the inventory without elevation, or when linked connections share them with the elevated token.
  $_isElevated = Test-DataInventoryAdministrator
  $_linked = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue
  $_hasLinkedConnections = $null -ne $_linked -and [int]$_linked.EnableLinkedConnections -eq 1
  if (-not $InventoryUser.IsCurrent) {
    $null = $_notes.Add("Connected drives were not read: they belong to the session running the inventory, not to $($InventoryUser.Name). Home drives, persistent mappings, and the logon script were read instead.")
  }
  elseif ($_isElevated -and -not $_hasLinkedConnections) {
    $null = $_notes.Add('Connected drives were not read: an elevated session cannot see the drives mapped by the non-elevated session unless EnableLinkedConnections is set. Home drives, persistent mappings, and the logon script were read instead.')
  }
  else {
    foreach ($_psDrive in @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.DisplayRoot -like '\\*' })) {
      $null = $_drives.Add([PSCustomObject]@{
          Letter = $_psDrive.Name
          Target = ([string]$_psDrive.DisplayRoot).TrimEnd('\')
          Source = 'Connected'
          Origin = 'Session'
        })
    }
  }

  # One entry per letter (or per target for letterless entries); earlier sources win.
  $_result = New-Object System.Collections.ArrayList
  $_seen = @{}
  foreach ($_drive in $_drives) {
    $_identity = 'target:' + $_drive.Target.ToUpperInvariant()
    if (-not [string]::IsNullOrWhiteSpace($_drive.Letter)) {
      $_identity = 'letter:' + $_drive.Letter.ToUpperInvariant()
    }

    if ($_seen.ContainsKey($_identity)) {
      continue
    }

    $_seen[$_identity] = $true

    # Measure through the UNC target; the letter may not be mapped in this session.
    $_summary = $null
    $_canMeasure = $_drive.Source -ne 'LogonScript' -and $_drive.Target.StartsWith('\\')
    if ($Measure -and $_canMeasure) {
      $_summary = Get-DataInventoryTreeSummary -Path $_drive.Target
    }

    $_files = $null
    $_bytes = $null
    $_newest = $null
    $_error = $null
    if ($null -ne $_summary) {
      $_files = $_summary.Files
      $_bytes = $_summary.Bytes
      $_newest = $_summary.Newest
      $_error = $_summary.Error
      if (-not $_summary.Exists) {
        $_error = "Not reachable from this session: $($_drive.Target)"
      }
    }

    $null = $_result.Add([PSCustomObject]@{
        Letter     = $_drive.Letter
        Target     = $_drive.Target
        Persistent = $_drive.Source -eq 'Persistent'
        Source     = $_drive.Source
        Origin     = $_drive.Origin
        Files      = $_files
        Bytes      = $_bytes
        Newest     = $_newest
        Measured   = $null -ne $_summary
        Error      = $_error
      })
  }

  return [PSCustomObject]@{
    Drives = @($_result.ToArray())
    Notes  = @($_notes.ToArray())
  }
}

function Get-DataInventoryOutsideProfile {
  param (
    [string]
    $ProfilePath,

    [string[]]
    $Include = @(),

    [bool]
    $ScanVolumes,

    [string]
    $ManagedRoot = (Join-Path $env:SystemDrive 'Managed')
  )

  # System, vendor, and winkit-managed folders hold no user data. Names starting with $ are
  # Windows servicing folders such as $Recycle.Bin and $WinREAgent. -Include overrides both.
  $_excludedNames = @(
    'Windows',
    'Windows.old',
    'Program Files',
    'Program Files (x86)',
    'ProgramData',
    'Users',
    'System Volume Information',
    'Recovery',
    'PerfLogs',
    'Documents and Settings',
    'Config.Msi',
    'MSOCache',
    'OneDriveTemp',
    'inetpub',
    'Intel',
    'AMD',
    'NVIDIA'
  )
  $_managedRoot = $ManagedRoot.TrimEnd('\')

  $_programs = @()
  foreach ($_key in @(
      'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
      'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
      'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )) {
    try {
      $_programs += @(Get-ItemProperty -Path $_key -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName })
    }
    catch {
      Write-Verbose "Could not read uninstall entries from $_key."
    }
  }

  # Collect top-level folders of fixed volumes, then add included folders that the scan skipped.
  $_candidates = New-Object System.Collections.ArrayList
  if ($ScanVolumes) {
    # Get-Volume has no -DriveType parameter; DriveType is an output property.
    $_volumes = @(
      Get-Volume -ErrorAction SilentlyContinue |
        Where-Object { $_.DriveLetter -and [string]$_.DriveType -eq 'Fixed' }
    )

    foreach ($_volume in $_volumes) {
      $_root = $_volume.DriveLetter + ':\'
      foreach ($_directory in @(Get-ChildItem -LiteralPath $_root -Directory -Force -ErrorAction SilentlyContinue)) {
        $_path = $_directory.FullName.TrimEnd('\')
        $_isIncluded = @($Include | Where-Object { $_.Equals($_path, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
        $_isExcluded = $_directory.Name -in $_excludedNames -or
        $_directory.Name.StartsWith('$') -or
        $_path.Equals($_managedRoot, [StringComparison]::OrdinalIgnoreCase)

        if ($_isExcluded -and -not $_isIncluded) {
          continue
        }

        if ($_directory.Attributes -band [IO.FileAttributes]::ReparsePoint) {
          continue
        }

        if (Test-DataInventoryPathWithin -Path $_path -Root $ProfilePath) {
          continue
        }

        $null = $_candidates.Add($_path)
      }
    }
  }

  foreach ($_path in $Include) {
    if (@($_candidates | Where-Object { $_.Equals($_path, [StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) {
      $null = $_candidates.Add($_path)
    }
  }

  $_result = New-Object System.Collections.ArrayList
  foreach ($_path in $_candidates) {
    $_summary = Get-DataInventoryTreeSummary -Path $_path
    $_isIncluded = @($Include | Where-Object { $_.Equals($_path, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
    $_owners = @(
      $_programs |
        Where-Object {
          $_location = ([string]$_.InstallLocation).TrimEnd('\')
          -not [string]::IsNullOrWhiteSpace($_location) -and (
            $_location.Equals($_path, [StringComparison]::OrdinalIgnoreCase) -or
            (Test-DataInventoryPathWithin -Path $_location -Root $_path) -or
            (Test-DataInventoryPathWithin -Path $_path -Root $_location)
          )
        } |
        Select-Object -ExpandProperty DisplayName -Unique
    )

    $null = $_result.Add([PSCustomObject]@{
        Path         = $_path
        Files        = $_summary.Files
        Bytes        = $_summary.Bytes
        Newest       = $_summary.Newest
        ReparseItems = $_summary.ReparseItems
        Programs     = @($_owners)
        Included     = $_isIncluded
        Error        = $_summary.Error
      })
  }

  return @($_result.ToArray())
}

function Resolve-DataInventoryInclude {
  param (
    [string[]]
    $Path,

    [string]
    $ProfilePath,

    [string]
    $PublicRoot
  )

  $_resolved = New-Object System.Collections.ArrayList
  foreach ($_entry in @($Path | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
    $_full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_entry)
    if ($_full -notmatch '^[A-Za-z]:\\') {
      throw "-Include needs a local drive path: $_entry"
    }

    $_full = [IO.Path]::GetFullPath($_full).TrimEnd('\')
    if ($_full.Length -le 2) {
      throw "-Include cannot name a drive root: $_entry"
    }

    if (-not (Test-Path -LiteralPath $_full -PathType Container)) {
      throw "-Include directory not found: $_full"
    }

    foreach ($_covered in @($ProfilePath, $PublicRoot)) {
      $_coveredRoot = $_covered.TrimEnd('\')
      if ($_full.Equals($_coveredRoot, [StringComparison]::OrdinalIgnoreCase) -or
        (Test-DataInventoryPathWithin -Path $_full -Root $_coveredRoot)) {
        throw "-Include is for folders outside the profile and the public folders: $_full"
      }
    }

    if (@($_resolved | Where-Object { $_.Equals($_full, [StringComparison]::OrdinalIgnoreCase) }).Count -eq 0) {
      $null = $_resolved.Add($_full)
    }
  }

  return @($_resolved.ToArray())
}

function ConvertFrom-DataInventoryRobocopyList {
  param (
    [string]
    $Text,

    [string]
    $SourceRoot,

    [string]
    $DestinationRoot
  )

  # robocopy /L /FP prints one full source path per line: a drive path or a UNC path, directories with a trailing \.
  $_sourceRoot = $SourceRoot.TrimEnd('\')
  $_destinationRoot = $DestinationRoot.TrimEnd('\')
  $_seen = @{}
  $_longPaths = New-Object System.Collections.ArrayList
  foreach ($_line in ($Text -split "`r?`n")) {
    if ($_line -notmatch '((?:[A-Za-z]:|\\\\[^\\\s]+\\[^\\]+)\\[^\r\n]*?)\s*$') {
      continue
    }

    $_sourcePath = $Matches[1].TrimEnd()
    $_isDirectory = $false
    if ($_sourcePath.EndsWith('\')) {
      $_isDirectory = $true
      $_sourcePath = $_sourcePath.TrimEnd('\')
    }

    $_isBelowRoot = $_sourcePath.Equals($_sourceRoot, [StringComparison]::OrdinalIgnoreCase) -or
    (Test-DataInventoryPathWithin -Path $_sourcePath -Root $_sourceRoot)
    if (-not $_isBelowRoot) {
      continue
    }

    $_relative = $_sourcePath.Substring($_sourceRoot.Length).TrimStart('\')
    $_destinationPath = $_destinationRoot
    if ($_relative) {
      $_destinationPath = Join-Path $_destinationRoot $_relative
    }

    # Same limits as Import-UserProfileData: longer than 247 (directory) or 259 (file) characters fails.
    $_limit = 259
    if ($_isDirectory) {
      $_limit = 247
    }

    if ($_destinationPath.Length -le $_limit -or $_seen.ContainsKey($_destinationPath)) {
      continue
    }

    $_seen[$_destinationPath] = $true
    $_wordPair = $false
    if (-not $_isDirectory -and [IO.Path]::GetExtension($_sourcePath) -in @('.htm', '.html')) {
      $_base = [IO.Path]::GetFileNameWithoutExtension($_sourcePath)
      $_parent = Split-Path -Path $_sourcePath -Parent
      foreach ($_suffix in @('-Dateien', '_files', '-filer', '_Dateien')) {
        if (Test-Path -LiteralPath (Join-Path $_parent ($_base + $_suffix)) -PathType Container -ErrorAction SilentlyContinue) {
          $_wordPair = $true
          break
        }
      }
    }

    $null = $_longPaths.Add([PSCustomObject]@{
        Source      = $_sourcePath
        Destination = $_destinationPath
        Length      = $_destinationPath.Length
        Limit       = $_limit
        IsDirectory = $_isDirectory
        WordWebPair = $_wordPair
      })
  }

  return @($_longPaths.ToArray())
}

function Get-DataInventoryLongPath {
  param (
    [string]
    $SourceRoot,

    [string]
    $DestinationRoot,

    [string[]]
    $ExcludeDirectory = @()
  )

  $_sourceRoot = $SourceRoot.TrimEnd('\')
  $_destinationRoot = $_sourceRoot
  if (-not [string]::IsNullOrWhiteSpace($DestinationRoot)) {
    $_destinationRoot = $DestinationRoot
  }

  $_destinationRoot = [IO.Path]::GetFullPath($_destinationRoot).TrimEnd('\')

  $_arguments = @(
    $_sourceRoot,
    $_destinationRoot,
    '/L',
    '/E',
    '/NJH',
    '/NJS',
    '/FP',
    '/BYTES',
    '/NC',
    '/NS',
    '/NP',
    '/XJ',
    '/R:0',
    '/W:0'
  )
  if ($ExcludeDirectory.Count -gt 0) {
    $_arguments += @('/XD') + $ExcludeDirectory
  }

  $_result = Invoke-SafeProcess -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList $_arguments -AsResult -ErrorAction Stop
  if ($_result.TimedOut -or $_result.Cancelled -or $null -eq $_result.ExitCode -or $_result.ExitCode -lt 0 -or $_result.ExitCode -ge 8) {
    throw "Path-length scan failed for ${_sourceRoot}: robocopy exit $($_result.ExitCode)"
  }

  return @(ConvertFrom-DataInventoryRobocopyList -Text $_result.StdOut -SourceRoot $_sourceRoot -DestinationRoot $_destinationRoot)
}

function Resolve-DataInventoryTarget {
  param (
    [string]
    $TargetRoot,

    [object]
    $InventoryUser
  )

  $_profileRoot = ([string]$InventoryUser.ProfilePath).TrimEnd('\')
  $_result = [PSCustomObject]@{
    Root    = $_profileRoot
    Assumed = $true
    User    = $null
    Notes   = @()
  }

  if ([string]::IsNullOrWhiteSpace($TargetRoot)) {
    $_result.Notes = @("-TargetRoot was not given: destination path lengths assume the new profile is $_profileRoot, the old user's folder, and the new user is unknown. Replace <target> in the import call.")
    return $_result
  }

  # The target names a folder on the new computer; its drive need not exist here, so avoid provider resolution.
  $_root = $TargetRoot
  if (-not [IO.Path]::IsPathRooted($_root)) {
    $_root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_root)
  }

  $_root = [IO.Path]::GetFullPath($_root).TrimEnd('\')
  $_result.Root = $_root
  $_result.Assumed = $false

  # A target below the same profiles folder names the new user; the domain is assumed to be the old user's.
  $_profilesFolder = Split-Path -Path $_profileRoot -Parent
  $_targetParent = Split-Path -Path $_root -Parent
  $_accountParts = ([string]$InventoryUser.Name).Split('\')
  if ($_targetParent -and $_targetParent.Equals($_profilesFolder, [StringComparison]::OrdinalIgnoreCase) -and $_accountParts.Count -eq 2) {
    $_result.User = '{0}\{1}' -f $_accountParts[0], (Split-Path -Path $_root -Leaf)
    $_result.Notes = @("The new user $($_result.User) in the import call is derived from -TargetRoot $_root and the old user's domain. Check it before the import.")
  }
  else {
    $_result.Notes = @("The new user cannot be determined from -TargetRoot $_root because it is not a profile folder below $_profilesFolder. Replace <target> in the import call.")
  }

  return $_result
}

function Get-DataInventoryOutlookReport {
  param (
    [string]
    $Path,

    [string]
    $Computer,

    [string]
    $UserName,

    [int]
    $MaxAgeDays,

    [string]
    $SignaturePath
  )

  $_result = [PSCustomObject]@{
    Path      = $null
    Valid     = $false
    Warnings  = @()
    DataFiles = @()
    Signature = $null
    Error     = $null
  }
  if ([string]::IsNullOrWhiteSpace($Path)) {
    $_result.Warnings += 'Outlook not inventoried: run Show-OutlookAccount.ps1 first and pass -OutlookReport.'
    return $_result
  }

  $_resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
  if (Test-Path -LiteralPath $_resolved -PathType Container) {
    $_candidate = Get-ChildItem -LiteralPath $_resolved -Filter 'Show-OutlookAccount*.json' -File -ErrorAction SilentlyContinue |
      Sort-Object -Property LastWriteTimeUtc -Descending |
      Select-Object -First 1
    if ($null -eq $_candidate) {
      $_result.Warnings += "No Show-OutlookAccount report was found in ${_resolved}."
      return $_result
    }
    $_resolved = $_candidate.FullName
  }

  $_result.Path = $_resolved
  try {
    $_report = Get-Content -LiteralPath $_resolved -Raw -Encoding UTF8 | ConvertFrom-Json
  }
  catch {
    $_result.Error = "Could not read the Outlook report: $($_.Exception.Message)"
    return $_result
  }

  if ([string]$_report.Script -ne 'Show-OutlookAccount') {
    $_result.Error = 'The report was not produced by Show-OutlookAccount.'
    return $_result
  }
  if ([string]$_report.ComputerName -ne $Computer) {
    $_result.Warnings += "The Outlook report is from computer '$($_report.ComputerName)', not '$Computer'."
  }
  if ([string]$_report.UserName -ne $UserName) {
    $_result.Warnings += "The Outlook report is for user '$($_report.UserName)', not '$UserName'."
  }

  $_finished = $null
  if (-not [string]::IsNullOrWhiteSpace([string]$_report.FinishedAt)) {
    try {
      $_finished = [DateTimeOffset]::Parse([string]$_report.FinishedAt, [Globalization.CultureInfo]::InvariantCulture).LocalDateTime
    }
    catch {
      Write-Verbose 'Could not parse the Outlook report timestamp.'
    }
  }
  if ($null -ne $_finished -and ((Get-Date) - $_finished).TotalDays -gt $MaxAgeDays) {
    $_result.Warnings += "The Outlook report is older than $MaxAgeDays day(s)."
  }

  $_files = New-Object System.Collections.ArrayList
  foreach ($_store in @($_report.Stores)) {
    if ([string]::IsNullOrWhiteSpace([string]$_store.FilePath)) {
      continue
    }
    $null = $_files.Add([PSCustomObject]@{
        DisplayName    = [string]$_store.DisplayName
        Path           = [string]$_store.FilePath
        SizeBytes      = $_store.SizeBytes
        Format         = [string]$_store.DataFileFormat
        ExchangeBacked = [bool]$_store.IsExchangeBacked
      })
  }

  $_result.DataFiles = @($_files.ToArray())
  $_result.Signature = $SignaturePath
  $_result.Valid = ($_result.Warnings.Count -eq 0 -and $null -eq $_result.Error)
  return $_result
}

function Add-DataInventoryRoot {
  param (
    [System.Collections.ArrayList]
    $Roots,

    [string]
    $Path,

    [string]
    $Reason,

    [string]
    $StagingName
  )

  $_existing = @($Roots | Where-Object { $_.Path.Equals($Path, [StringComparison]::OrdinalIgnoreCase) }) | Select-Object -First 1
  if ($null -ne $_existing) {
    if ($Reason -notin $_existing.Reasons) {
      $null = $_existing.Reasons.Add($Reason)
    }

    return $_existing.StagingName
  }

  # Each download is extracted as <StagingRoot>\<folder name>; equal names get a numbered suffix.
  $_leaf = Split-Path -Path $Path -Leaf
  if (-not [string]::IsNullOrWhiteSpace($StagingName)) {
    $_leaf = $StagingName
  }

  $_name = $_leaf
  $_index = 1
  while (@($Roots | Where-Object { $_.StagingName -ieq $_name }).Count -gt 0) {
    $_index++
    $_name = '{0} ({1})' -f $_leaf, $_index
  }

  $_reasons = New-Object System.Collections.ArrayList
  $null = $_reasons.Add($Reason)
  $null = $Roots.Add([PSCustomObject]@{
      Path        = $Path
      Leaf        = $_leaf
      StagingName = $_name
      Reasons     = $_reasons
    })

  return $_name
}

function Get-DataInventoryMapping {
  param (
    [object[]]
    $KnownFolder,

    [object[]]
    $PublicFolder,

    [object[]]
    $OutsideProfile,

    [string]
    $ProfilePath,

    [string]
    $PublicRoot,

    [object[]]
    $NetworkDrive = @(),

    [string[]]
    $DriveNote = @(),

    [string]
    $HomeDriveTarget = 'Private',

    [string[]]
    $Folder = @(),

    [string[]]
    $IncludePublic = @(),

    [int]
    $NearlyEmptyFileCount = 5
  )

  $_roots = New-Object System.Collections.ArrayList
  $_mappings = New-Object System.Collections.ArrayList
  $_manual = New-Object System.Collections.ArrayList
  $_notes = New-Object System.Collections.ArrayList
  $_decisions = New-Object System.Collections.ArrayList
  $_profileRoot = $ProfilePath.TrimEnd('\')
  $_mappedOrigins = @{}

  # A home drive holds the user's own data, so it is mapped; other drives and logon scripts need a decision.
  foreach ($_note in $DriveNote) {
    $null = $_notes.Add($_note)
  }

  foreach ($_drive in $NetworkDrive) {
    $_label = $_drive.Target
    if (-not [string]::IsNullOrWhiteSpace($_drive.Letter)) {
      $_label = "$($_drive.Letter): ($($_drive.Target))"
    }

    $_size = 'use -MeasureNetwork for its size'
    if ($_drive.Measured -and -not $_drive.Error) {
      $_size = '{0} files, {1:N1} MB' -f $_drive.Files, ($_drive.Bytes / 1MB)
    }
    elseif ($_drive.Error) {
      $_size = "not measured: $($_drive.Error)"
    }

    if ($_drive.Source -eq 'LogonScript') {
      $null = $_decisions.Add("Logon script $($_drive.Target) may connect further drives at sign-in. Review it.")
      continue
    }

    if ($_drive.Source -ne 'HomeDrive') {
      $null = $_decisions.Add("Network drive $_label from $($_drive.Source). Decide whether to stage it; $_size.")
      continue
    }

    $_rootName = Add-DataInventoryRoot -Roots $_roots -Path $_drive.Target -Reason "Home drive $_label" -StagingName 'HomeDrive'
    $null = $_mappings.Add([PSCustomObject]@{
        Source = $_rootName
        Target = $HomeDriveTarget
        Origin = $_drive.Target
        Kind   = 'HomeDrive'
      })
    $null = $_notes.Add("Home drive $_label holds the user's own data ($_size). ABB does not back up the file server: copy the share into the '$_rootName' staging folder before the import. It is imported into $HomeDriveTarget.")
  }

  # Known folders are imported from where they actually are, into their standard location. Empty folders are
  # only noted: a download may leave them out, and the import refuses a missing mapped source.
  $_unselected = New-Object System.Collections.ArrayList
  foreach ($_item in $KnownFolder) {
    if (-not $_item.Exists) {
      continue
    }

    if ($Folder.Count -gt 0 -and $_item.Name -notin $Folder) {
      if ($_item.Files -gt 0) {
        $null = $_unselected.Add($_item.Name)
      }

      continue
    }

    if ($_item.Files -eq 0) {
      $_nestedText = ''
      $_nestedNames = @($_item.NestedFolders | Where-Object { $_ })
      if ($_nestedNames.Count -gt 0) {
        $_nestedText = ' apart from ' + ($_nestedNames -join ', ')
      }

      $null = $_notes.Add("Known folder $($_item.Name) is empty$_nestedText and is not mapped.")
      continue
    }

    if ($_item.Files -le $NearlyEmptyFileCount) {
      $null = $_notes.Add("Known folder $($_item.Name) holds only $($_item.Files) file(s). It is mapped; check whether it is worth migrating.")
    }

    if ($_item.IsUnc -or $_item.IsOneDrive -or $_item.IsReparse) {
      $null = $_decisions.Add("Known folder $($_item.Name) is redirected to a network, OneDrive, or linked location: $($_item.Path). Stage it separately.")
      continue
    }

    $_origin = [IO.Path]::GetFullPath($_item.Path).TrimEnd('\')
    $_key = $_origin.ToUpperInvariant()
    if ($_mappedOrigins.ContainsKey($_key)) {
      $null = $_notes.Add("Known folder $($_item.Name) shares $_origin with $($_mappedOrigins[$_key]); it is imported once, into $($_mappedOrigins[$_key]).")
      continue
    }

    if ($_origin.Equals($_profileRoot, [StringComparison]::OrdinalIgnoreCase)) {
      $null = $_decisions.Add("Known folder $($_item.Name) points at the profile root $_origin. Map its files manually.")
      continue
    }

    $_mappedOrigins[$_key] = $_item.Name
    if ($_item.IsOutsideProfile) {
      $_rootName = Add-DataInventoryRoot -Roots $_roots -Path $_origin -Reason "Known folder $($_item.Name)"
      $_source = $_rootName
      $null = $_notes.Add("Known folder $($_item.Name) is redirected outside the profile to $_origin. Download it separately; it is imported into $($_item.Name).")
    }
    else {
      $_rootName = Add-DataInventoryRoot -Roots $_roots -Path $_profileRoot -Reason 'User profile'
      $_source = Join-Path $_rootName $_origin.Substring($_profileRoot.Length).TrimStart('\')
    }

    $null = $_mappings.Add([PSCustomObject]@{
        Source = $_source
        Target = $_item.Name
        Origin = $_origin
        Kind   = 'KnownFolder'
      })
  }

  if ($_unselected.Count -gt 0) {
    $null = $_notes.Add("Known folders with data that -Folder did not select and are not mapped: $($_unselected -join ', ').")
  }

  # Public folders are shared by every user of the old computer, so they are imported only when named with
  # -IncludePublic. The desktop belongs on the new computer's public desktop and is never imported.
  $_publicRoot = [IO.Path]::GetFullPath($PublicRoot).TrimEnd('\')
  foreach ($_item in $PublicFolder) {
    foreach ($_shortcut in @($_item.Shortcuts)) {
      if ($_shortcut.IsNetwork) {
        $null = $_decisions.Add("Public shortcut targets a network path: $($_shortcut.Path) -> $($_shortcut.Target).")
      }
    }

    if ($_item.Files -eq 0) {
      continue
    }

    if ($_item.Name -eq 'Desktop') {
      # Shortcuts to local programs come back with the new computer's installers; real files need a copy.
      $_programShortcuts = @(
        $_item.Shortcuts |
          Where-Object {
            -not $_.IsNetwork -and [IO.Path]::GetExtension($_.Path) -eq '.lnk' -and
            ([string]::IsNullOrWhiteSpace($_.Target) -or $_.Target -match '^[A-Za-z]:\\')
          }
      )
      $_networkShortcuts = @($_item.Shortcuts | Where-Object { $_.IsNetwork })
      $_otherFiles = $_item.Files - $_programShortcuts.Count - $_networkShortcuts.Count

      if ($_programShortcuts.Count -gt 0) {
        $_names = @($_programShortcuts | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_.Path) }) -join ', '
        $null = $_notes.Add("The public desktop holds $($_programShortcuts.Count) shortcut(s) to local programs ($_names). They are recreated by the installers on the new computer and are not copied.")
      }

      if ($_otherFiles -gt 0) {
        $_publicName = Add-DataInventoryRoot -Roots $_roots -Path $_publicRoot -Reason 'Public Desktop'
        $null = $_manual.Add([PSCustomObject]@{
            Path   = $_item.Path
            Reason = "Public desktop: copy its $_otherFiles file(s) that are not program shortcuts into %PUBLIC%\Desktop on the new computer"
          })
        $null = $_notes.Add("The public desktop is shared by all users and is not imported into the profile. Copy the files from $_publicName\Desktop that are not program shortcuts into %PUBLIC%\Desktop on the new computer.")
      }

      continue
    }

    $_origin = [IO.Path]::GetFullPath($_item.Path).TrimEnd('\')
    $_target = Join-Path $_item.Name "Public $($_item.Name)"
    if ($_item.Name -notin $IncludePublic) {
      $_newest = ''
      if ($_item.Newest) {
        $_newest = ", newest change $($_item.Newest)"
      }

      $null = $_decisions.Add(("Public {0} holds {1} files ({2:N1} MB{3}) shared by all users of the old computer. Use -IncludePublic {0} to import it into {4}." -f $_item.Name, $_item.Files, ($_item.Bytes / 1MB), $_newest, $_target))
      continue
    }

    $_publicName = Add-DataInventoryRoot -Roots $_roots -Path $_publicRoot -Reason "Public $($_item.Name)"
    $null = $_mappings.Add([PSCustomObject]@{
        Source = Join-Path $_publicName $_item.Name
        Target = $_target
        Origin = $_origin
        Kind   = 'Public'
      })
  }

  # Other data outside the profile is mapped only when named with -Include.
  foreach ($_item in $OutsideProfile) {
    $_origin = [IO.Path]::GetFullPath($_item.Path).TrimEnd('\')
    if ($_mappedOrigins.ContainsKey($_origin.ToUpperInvariant())) {
      continue
    }

    if (-not $_item.Included) {
      if ($_item.Files -eq 0) {
        continue
      }

      $_detail = '{0} files, {1:N1} MB' -f $_item.Files, ($_item.Bytes / 1MB)
      $_programs = @($_item.Programs | Where-Object { $_ })
      if ($_programs.Count -gt 0) {
        $_detail += '; probable programs: ' + ($_programs -join ', ')
      }

      $null = $_decisions.Add("Data outside the profile: $_origin ($_detail). Use -Include $(Format-DataInventoryLiteral -Value $_origin) to migrate it.")
      continue
    }

    $_rootName = Add-DataInventoryRoot -Roots $_roots -Path $_origin -Reason 'Included with -Include'
    $_target = Join-Path 'Documents' (Split-Path -Path $_origin -Leaf)
    $null = $_mappings.Add([PSCustomObject]@{
        Source = $_rootName
        Target = $_target
        Origin = $_origin
        Kind   = 'Included'
      })
    $null = $_notes.Add("$_origin is imported into $_target. Edit folder-map.json to choose another target.")
  }

  # Explain folders stored inside another mapped folder; the import excludes them from the outer copy.
  foreach ($_inner in $_mappings) {
    $_containing = @(
      $_mappings |
        Where-Object { $_ -ne $_inner -and (Test-DataInventoryPathWithin -Path $_inner.Origin -Root $_.Origin) } |
        Sort-Object { $_.Origin.Length } -Descending
    )
    $_outer = $null
    if ($_containing.Count -gt 0) {
      $_outer = $_containing[0]
    }

    if ($null -ne $_outer) {
      $null = $_notes.Add("$($_inner.Target) is stored inside $($_outer.Target) at $($_inner.Origin). It is imported into $($_inner.Target) and left out of the $($_outer.Target) import.")
    }
    elseif ($_inner.Kind -eq 'KnownFolder' -and -not ($_inner.Source -split '\\', 2)[-1].Equals($_inner.Target, [StringComparison]::OrdinalIgnoreCase) -and
      (Test-DataInventoryPathWithin -Path $_inner.Origin -Root $_profileRoot)) {
      $null = $_notes.Add("Known folder $($_inner.Target) is stored at $($_inner.Origin). It is imported into $($_inner.Target).")
    }
  }

  foreach ($_root in $_roots) {
    if ($_root.StagingName -ne $_root.Leaf) {
      $null = $_notes.Add("Another download is also named '$($_root.Leaf)'. Extract $($_root.Path) as '$($_root.StagingName)'.")
    }
  }

  return [PSCustomObject]@{
    Roots      = @($_roots.ToArray())
    Mappings   = @($_mappings.ToArray())
    ManualCopy = @($_manual.ToArray())
    Notes      = @($_notes.ToArray())
    Decisions  = @($_decisions.ToArray())
  }
}

function Get-DataInventoryPlan {
  param (
    [object]
    $Mapping,

    [object[]]
    $LongPath,

    [object]
    $Outlook,

    [string]
    $StagingRoot,

    [string]
    $FolderMapPath,

    [string]
    $TargetUser
  )

  $_downloads = New-Object System.Collections.ArrayList
  foreach ($_root in $Mapping.Roots) {
    $null = $_downloads.Add([PSCustomObject]@{
        Path        = $_root.Path
        StagingPath = Join-Path $StagingRoot $_root.StagingName
        Reason      = ($_root.Reasons -join '; ')
      })
  }

  $_manual = New-Object System.Collections.ArrayList
  foreach ($_item in $Mapping.ManualCopy) {
    $null = $_manual.Add($_item)
  }

  $_decisions = New-Object System.Collections.ArrayList
  foreach ($_decision in $Mapping.Decisions) {
    $null = $_decisions.Add($_decision)
  }

  foreach ($_item in $LongPath) {
    $null = $_decisions.Add("Path too long for the target: $($_item.Destination) ($($_item.Length)/$($_item.Limit)).")
    if ($_item.WordWebPair) {
      $null = $_decisions.Add("Word web-page pair needs manual renaming: $($_item.Source).")
    }
  }

  if ($null -ne $Outlook) {
    foreach ($_item in $Outlook.DataFiles) {
      $null = $_manual.Add([PSCustomObject]@{
          Path   = $_item.Path
          Reason = "Outlook $($_item.Format) data file"
        })
    }

    if (-not [string]::IsNullOrWhiteSpace($Outlook.Signature)) {
      $null = $_manual.Add([PSCustomObject]@{
          Path   = $Outlook.Signature
          Reason = 'Outlook signatures'
        })
    }

    foreach ($_warning in $Outlook.Warnings) {
      $null = $_decisions.Add($_warning)
    }

    if ($Outlook.Error) {
      $null = $_decisions.Add($Outlook.Error)
    }
  }
  else {
    $null = $_decisions.Add('Outlook not inventoried: run Show-OutlookAccount.ps1 first and pass -OutlookReport.')
  }

  $_notes = New-Object System.Collections.ArrayList
  foreach ($_note in $Mapping.Notes) {
    $null = $_notes.Add($_note)
  }

  $_suggestedCall = $null
  if ($Mapping.Mappings.Count -gt 0) {
    $null = $_notes.Add("Mapping sources are relative to the staging root, not to $StagingRoot itself. Extract each download to its AbbDownload staging path, or keep the same folder names below any other root and pass that root to Import-UserProfileData -SourcePath.")
    $_mapPath = $FolderMapPath
    if ([string]::IsNullOrWhiteSpace($_mapPath)) {
      $_mapPath = '<folder-map.json>'
    }

    $_sourceLiteral = Format-DataInventoryLiteral -Value $StagingRoot
    $_targetUser = $TargetUser
    if ([string]::IsNullOrWhiteSpace($_targetUser)) {
      $_targetUser = '<target>'
    }

    $_targetLiteral = Format-DataInventoryLiteral -Value $_targetUser
    $_mapLiteral = Format-DataInventoryLiteral -Value $_mapPath
    $_suggestedCall = ".\scripts\Administration\Import-UserProfileData.ps1 -SourcePath $_sourceLiteral -TargetUser $_targetLiteral -FolderMapPath $_mapLiteral -DryRun"
  }

  return [PSCustomObject]@{
    AbbDownload         = @($_downloads.ToArray())
    Mappings            = @($Mapping.Mappings)
    ManualCopy          = @($_manual.ToArray())
    Notes               = @($_notes.ToArray())
    Decisions           = @($_decisions.ToArray())
    SuggestedImportCall = $_suggestedCall
  }
}

function ConvertTo-DataInventoryFolderMap {
  param (
    [object]
    $Plan,

    [string]
    $RunId,

    [string]
    $Computer,

    [string]
    $UserName
  )

  $_mappings = New-Object System.Collections.ArrayList
  foreach ($_item in $Plan.Mappings) {
    $null = $_mappings.Add([ordered]@{
        Source = $_item.Source
        Target = $_item.Target
        Origin = $_item.Origin
        Kind   = $_item.Kind
      })
  }

  return [ordered]@{
    SchemaVersion  = 1
    GeneratedBy    = 'Show-UserDataInventory'
    RunId          = $RunId
    GeneratedAtUtc = [DateTime]::UtcNow.ToString('o')
    Computer       = $Computer
    User           = $UserName
    Mappings       = @($_mappings.ToArray())
    Notes          = @($Plan.Notes)
  }
}

$_inventoryUser = Resolve-DataInventoryUser -Account $User
$_startedAt = [DateTime]::UtcNow.ToString('o')
$_runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$_reportRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportDirectory)
$_runDirectory = Join-Path $_reportRoot $_runId
$_stagingRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($StagingRoot).TrimEnd('\')
$_manifestPath = $null
$_folderMapPath = $null
$_inventory = $null

try {
  $_include = @(Resolve-DataInventoryInclude -Path $Include -ProfilePath $_inventoryUser.ProfilePath -PublicRoot $env:PUBLIC)

  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Reading known folders'
  $_known = @(Get-DataInventoryKnownFolder -RegistryRoot $_inventoryUser.Registry -ProfilePath $_inventoryUser.ProfilePath)
  $_public = @(Get-DataInventoryPublicFolder -PublicRoot $env:PUBLIC -IncludePermissions $_inventoryUser.IsAdmin)

  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Reading home and network drives'
  $_driveInventory = Get-DataInventoryNetworkDrive -RegistryRoot $_inventoryUser.Registry -InventoryUser $_inventoryUser -Measure $MeasureNetwork.IsPresent
  $_network = @($_driveInventory.Drives)

  if (-not $_inventoryUser.IsAdmin) {
    Write-Warning 'Run elevated to inspect fixed drives outside the profile.'
  }

  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Scanning folders outside the profile'
  $_outside = @(Get-DataInventoryOutsideProfile -ProfilePath $_inventoryUser.ProfilePath -Include $_include -ScanVolumes $_inventoryUser.IsAdmin)

  $_mapping = Get-DataInventoryMapping `
    -KnownFolder $_known `
    -PublicFolder $_public `
    -OutsideProfile $_outside `
    -ProfilePath $_inventoryUser.ProfilePath `
    -PublicRoot $env:PUBLIC `
    -NetworkDrive $_network `
    -DriveNote $_driveInventory.Notes `
    -HomeDriveTarget $HomeDriveTarget `
    -Folder $Folder `
    -IncludePublic $IncludePublic

  # Destination lengths are checked against the new profile; without -TargetRoot the old profile path stands in.
  $_target = Resolve-DataInventoryTarget -TargetRoot $TargetRoot -InventoryUser $_inventoryUser
  $_mapping.Notes = @($_mapping.Notes) + @($_target.Notes)
  if ($_target.Assumed) {
    Write-Warning $_target.Notes[0]
    $_mapping.Decisions = @($_mapping.Decisions) + "Pass -TargetRoot with the new profile folder, for example -TargetRoot 'C:\Users\<new user>', for exact destination path lengths."
  }

  # Every mapped tree, the home drive included, is checked at its mapped target; measured drives as if imported
  # into Documents\<share name>. Nested mapped folders are checked on their own.
  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Checking destination path lengths'
  $_checks = New-Object System.Collections.ArrayList
  foreach ($_item in $_mapping.Mappings) {
    $_nested = @(
      $_mapping.Mappings |
        Where-Object { $_ -ne $_item -and (Test-DataInventoryPathWithin -Path $_.Origin -Root $_item.Origin) } |
        ForEach-Object { $_.Origin }
    )
    $null = $_checks.Add([PSCustomObject]@{
        Origin      = $_item.Origin
        Destination = Join-Path $_target.Root $_item.Target
        Exclude     = $_nested
      })
  }

  foreach ($_drive in @($_network | Where-Object { $_.Measured -and -not $_.Error -and $_.Source -ne 'HomeDrive' })) {
    $null = $_checks.Add([PSCustomObject]@{
        Origin      = $_drive.Target
        Destination = Join-Path $_target.Root (Join-Path 'Documents' (Split-Path -Path $_drive.Target -Leaf))
        Exclude     = @()
      })
  }

  $_long = New-Object System.Collections.ArrayList
  foreach ($_check in $_checks) {
    try {
      foreach ($_entry in @(Get-DataInventoryLongPath -SourceRoot $_check.Origin -DestinationRoot $_check.Destination -ExcludeDirectory $_check.Exclude)) {
        $null = $_long.Add($_entry)
      }
    }
    catch {
      Write-Warning "Path lengths were not checked for $($_check.Origin): $($_.Exception.Message)"
      $_mapping.Decisions = @($_mapping.Decisions) + "Path lengths were not checked for $($_check.Origin): $($_.Exception.Message)"
    }
  }

  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Reading Outlook evidence'
  $_signaturePath = Join-Path $_inventoryUser.ProfilePath 'AppData\Roaming\Microsoft\Signatures'
  $_outlookReport = @{
    Path          = $OutlookReport
    Computer      = $env:COMPUTERNAME
    UserName      = $_inventoryUser.Name
    MaxAgeDays    = $OutlookReportMaxAge
    SignaturePath = $_signaturePath
  }
  $_outlook = Get-DataInventoryOutlookReport @_outlookReport

  if (-not $DryRun) {
    $_folderMapPath = Join-Path $_runDirectory 'folder-map.json'
  }

  $_planParameters = @{
    Mapping       = $_mapping
    LongPath      = @($_long.ToArray())
    Outlook       = $_outlook
    StagingRoot   = $_stagingRoot
    FolderMapPath = $_folderMapPath
    TargetUser    = $_target.User
  }
  $_plan = Get-DataInventoryPlan @_planParameters

  $_inventory = [PSCustomObject]@{
    PSTypeName        = 'Winkit.DataInventory'
    SchemaVersion     = 2
    RunId             = $_runId
    StartedAtUtc      = $_startedAt
    FinishedAtUtc     = [DateTime]::UtcNow.ToString('o')
    Computer          = $env:COMPUTERNAME
    User              = $_inventoryUser.Name
    UserSID           = $_inventoryUser.SID
    ProfilePath       = $_inventoryUser.ProfilePath
    TargetRoot        = $_target.Root
    TargetRootAssumed = $_target.Assumed
    TargetUser        = $_target.User
    StagingRoot       = $_stagingRoot
    Include           = $_include
    OutlookReport     = $_outlook
    KnownFolders      = $_known
    PublicFolders     = $_public
    NetworkDrives     = $_network
    OutsideProfile    = $_outside
    LongPaths         = @($_long.ToArray())
    Plan              = $_plan
    ReportPath        = $null
    FolderMapPath     = $_folderMapPath
  }

  $_hasError = @($_known | Where-Object Error).Count -gt 0 -or @($_outside | Where-Object Error).Count -gt 0
  $_status = 'Completed'
  if ($_hasError) {
    $_status = 'CompletedWithWarnings'
  }

  # Write the inventory manifest and the folder map consumed by Import-UserProfileData -FolderMapPath.
  if (-not $DryRun) {
    $null = New-Item -ItemType Directory -Path $_runDirectory -Force -ErrorAction Stop
    $_manifestPath = Join-Path $_runDirectory 'manifest.json'
    $_inventory.ReportPath = $_manifestPath
    $_encoding = New-Object Text.UTF8Encoding($true)

    $_json = ConvertTo-DataInventoryJson -InputObject $_inventory -Depth 20
    [IO.File]::WriteAllText($_manifestPath, $_json, $_encoding)

    $_folderMap = ConvertTo-DataInventoryFolderMap -Plan $_plan -RunId $_runId -Computer $env:COMPUTERNAME -UserName $_inventoryUser.Name
    $_folderMapJson = ConvertTo-DataInventoryJson -InputObject $_folderMap -Depth 6
    [IO.File]::WriteAllText($_folderMapPath, $_folderMapJson, $_encoding)

    $_operationParameters = @{
      Target   = $_inventoryUser.ProfilePath
      Source   = 'UserData'
      Action   = 'Inventory'
      Status   = $_status
      Detail   = 'User data inventory completed.'
      Property = @{
        ReportPath    = $_manifestPath
        FolderMapPath = $_folderMapPath
      }
    }
    $_operation = New-OperationResult @_operationParameters

    $_logParameters = @{
      Results    = @($_operation)
      Path       = Join-Path $_runDirectory 'operations.jsonl'
      ScriptName = 'Show-UserDataInventory'
      RunId      = $_runId
    }
    $null = Write-OperationResultLog @_logParameters
  }

  Write-Log -Message "User data inventory: $_status | Known folders: $($_known.Count) | Public: $($_public.Count) | Network: $($_network.Count) | Outside: $($_outside.Count) | Long paths: $($_long.Count) | Report: $_manifestPath" -Color Green
  Write-Log -Message "Plan: download $($_plan.AbbDownload.Count) folder(s), $($_plan.Mappings.Count) mapping(s), manual copy $($_plan.ManualCopy.Count), notes $($_plan.Notes.Count), decisions $($_plan.Decisions.Count). Folder map: $_folderMapPath" -Color Cyan
  if ($PassThru -or $DryRun) {
    $_inventory
  }
}
catch {
  Write-Log -Message "User data inventory failed: $($_.Exception.Message)" -Color Red
  exit 1
}
finally {
  Write-Progress -Id 71 -Activity 'User data inventory' -Completed
  if ($_inventoryUser.LoadedHive) {
    $null = Invoke-SafeProcess -FilePath "$env:SystemRoot\System32\reg.exe" -ArgumentList @('unload', "HKU\$($_inventoryUser.LoadedHive)") -AsResult -ErrorAction SilentlyContinue
  }
}
exit 0
