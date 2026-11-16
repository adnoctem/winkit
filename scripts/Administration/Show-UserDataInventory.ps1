#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.4.0' }

<#
.SYNOPSIS
  Inventories a user's data locations and generates a migration plan.

.DESCRIPTION
  Reads the selected user's known folders, public folders and shortcuts,
  persistent network drives, outside-profile data, and path lengths. It can
  consume a Show-OutlookAccount report to add PST/OST files and signatures to
  the manual-copy plan.

  The script is read-only. It runs in the user's session without elevation for
  the current user. Another local profile requires administrator rights; its
  registry hive is loaded only when the profile is not in use.

  The generated plan lists ABB download locations, suggested
  Import-UserProfileData calls, manual copies, and decisions. A JSON report is
  written to a unique run directory unless -DryRun is supplied.

.PARAMETER User
  User to inventory. Defaults to the current user. Accepts an account name or
  SID. Another local profile requires administrator rights.

.PARAMETER TargetRoot
  Future profile root used for destination path-length checks, for example
  C:\Users\newuser. Without it, absolute source lengths are reported.

.PARAMETER OutlookReport
  Show-OutlookAccount JSON report or a directory containing reports. Without it,
  the plan states that Outlook was not inventoried.

.PARAMETER OutlookReportMaxAge
  Maximum accepted report age in days. Defaults to 7.

.PARAMETER MeasureNetwork
  Measure size and file counts on mapped network drives. Off by default because
  it can be slow.

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
    $Path
  )

  $result = [PSCustomObject]@{
    Path          = $Path
    Exists        = $false
    Files         = 0
    Directories   = 0
    Bytes         = [long]0
    Newest        = $null
    ReparsePoints = 0
    Offline       = 0
    Encrypted     = 0
    Error         = $null
  }

  if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) {
    return $result
  }

  $result.Exists = $true
  try {
    foreach ($item in @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Stop)) {
      if ($item.PSIsContainer) {
        $result.Directories++
      }
      else {
        $result.Files++
        $result.Bytes += [long]$item.Length
      }
      if ($null -eq $result.Newest -or $item.LastWriteTimeUtc -gt $result.Newest) {
        $result.Newest = $item.LastWriteTimeUtc
      }
      if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        $result.ReparsePoints++
      }
      if ($item.Attributes -band [IO.FileAttributes]::Offline) {
        $result.Offline++
      }
      if ($item.Attributes -band [IO.FileAttributes]::Encrypted) {
        $result.Encrypted++
      }
    }
  }
  catch {
    $result.Error = $_.Exception.Message
  }

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
  $_result = New-Object System.Collections.ArrayList
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

    $_path = Expand-DataInventoryPath -Path $_raw -ProfilePath $ProfilePath
    $_summary = Get-DataInventoryTreeSummary -Path $_path
    $_isUnc = $_path.StartsWith('\\')
    $_isOutsideProfile = -not $_path.StartsWith($ProfilePath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)
    $_isOneDrive = $_path -match '(^|\\)OneDrive( - |\\)'
    $_isReparse = $false
    if (Test-Path -LiteralPath $_path -ErrorAction SilentlyContinue) {
      $_isReparse = [bool]((Get-Item -LiteralPath $_path -Force -ErrorAction SilentlyContinue).Attributes -band [IO.FileAttributes]::ReparsePoint)
    }

    $null = $_result.Add([PSCustomObject]@{
        Name             = $_definition.Name
        Path             = $_path
        Raw              = $_raw
        Exists           = $_summary.Exists
        Files            = $_summary.Files
        Directories      = $_summary.Directories
        Bytes            = $_summary.Bytes
        Newest           = $_summary.Newest
        ReparsePoints    = $_summary.ReparsePoints
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
        Name        = $_name
        Path        = $_path
        Files       = $_summary.Files
        Directories = $_summary.Directories
        Bytes       = $_summary.Bytes
        Newest      = $_summary.Newest
        Shortcuts   = @($_shortcuts)
        Permissions = $_permissions
        Error       = $_summary.Error
      })
  }

  return @($_result.ToArray())
}

function Get-DataInventoryNetworkDrive {
  param (
    [string]
    $RegistryRoot,

    [bool]
    $Measure
  )

  $_drives = @{}
  $_key = Join-Path $RegistryRoot 'Network'
  if (Test-Path -LiteralPath $_key -ErrorAction SilentlyContinue) {
    foreach ($_child in @(Get-ChildItem -LiteralPath $_key -ErrorAction SilentlyContinue)) {
      $_properties = Get-ItemProperty -LiteralPath $_child.PSPath -ErrorAction SilentlyContinue
      if ($null -eq $_properties -or [string]::IsNullOrWhiteSpace([string]$_properties.RemotePath)) {
        continue
      }
      $_drives[$_child.PSChildName.ToUpperInvariant()] = [PSCustomObject]@{
        Letter     = $_child.PSChildName
        Target     = [string]$_properties.RemotePath
        Persistent = $true
        Source     = 'Persistent'
      }
    }
  }

  foreach ($_drive in @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object { $_.DisplayRoot -like '\\*' })) {
    $_letter = $_drive.Name.ToUpperInvariant()
    if (-not $_drives.ContainsKey($_letter)) {
      $_drives[$_letter] = [PSCustomObject]@{
        Letter     = $_drive.Name
        Target     = [string]$_drive.DisplayRoot
        Persistent = $false
        Source     = 'Connected'
      }
    }
  }

  $_result = New-Object System.Collections.ArrayList
  foreach ($_drive in $_drives.Values) {
    $_summary = $null
    if ($Measure) {
      $_summary = Get-DataInventoryTreeSummary -Path ($_drive.Letter + ':\')
    }
    $null = $_result.Add([PSCustomObject]@{
        Letter     = $_drive.Letter
        Target     = $_drive.Target
        Persistent = $_drive.Persistent
        Source     = $_drive.Source
        Files      = if ($_summary) { $_summary.Files } else { $null }
        Bytes      = if ($_summary) { $_summary.Bytes } else { $null }
        Newest     = if ($_summary) { $_summary.Newest } else { $null }
        Measured   = [bool]$Measure
      })
  }

  return @($_result.ToArray())
}

function Get-DataInventoryOutsideProfile {
  param (
    [string]
    $ProfilePath
  )

  $_excluded = @(
    'Windows', 'Program Files', 'Program Files (x86)', 'ProgramData', 'Users',
    '$Recycle.Bin', 'System Volume Information', 'Recovery', 'PerfLogs',
    'Documents and Settings', 'Config.Msi', 'MSOCache'
  )
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

  $_result = New-Object System.Collections.ArrayList
  foreach ($_volume in @(Get-Volume -DriveType Fixed -ErrorAction SilentlyContinue | Where-Object { $_.DriveLetter })) {
    $_root = $_volume.DriveLetter + ':\'
    foreach ($_directory in @(Get-ChildItem -LiteralPath $_root -Directory -Force -ErrorAction SilentlyContinue)) {
      if ($_directory.Name -in $_excluded -or ($_directory.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        continue
      }
      if ($_directory.FullName.StartsWith($ProfilePath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        continue
      }

      $_summary = Get-DataInventoryTreeSummary -Path $_directory.FullName
      $_owners = @($_programs | Where-Object {
          -not [string]::IsNullOrWhiteSpace([string]$_.InstallLocation) -and (
            $_.InstallLocation.StartsWith($_directory.FullName, [StringComparison]::OrdinalIgnoreCase) -or
            $_directory.FullName.StartsWith($_.InstallLocation, [StringComparison]::OrdinalIgnoreCase)
          )
        } | Select-Object -ExpandProperty DisplayName -Unique)
      $null = $_result.Add([PSCustomObject]@{
          Path     = $_directory.FullName
          Files    = $_summary.Files
          Bytes    = $_summary.Bytes
          Newest   = $_summary.Newest
          Programs = @($_owners)
          Error    = $_summary.Error
        })
    }
  }

  return @($_result.ToArray())
}

function Get-DataInventoryLongPath {
  param (
    [string]
    $SourceRoot,

    [string]
    $TargetRoot
  )

  $_destinationRoot = $SourceRoot
  if (-not [string]::IsNullOrWhiteSpace($TargetRoot)) {
    $_destinationRoot = $TargetRoot
  }
  $_destinationRoot = [IO.Path]::GetFullPath($_destinationRoot).TrimEnd('\')

  $_arguments = @(
    $SourceRoot,
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
    '/XJ'
  )
  $_result = Invoke-SafeProcess -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList $_arguments -AsResult -ErrorAction Stop
  if ($_result.TimedOut -or $_result.Cancelled -or $null -eq $_result.ExitCode -or $_result.ExitCode -lt 0 -or $_result.ExitCode -ge 8) {
    throw "Path-length scan failed for ${SourceRoot}: $($_result.ExitCode)"
  }

  $_sourceRoot = [IO.Path]::GetFullPath($SourceRoot).TrimEnd('\')
  $_seen = @{}
  $_longPaths = New-Object System.Collections.ArrayList
  foreach ($_line in ($_result.StdOut -split "`r?`n")) {
    if ($_line -notmatch '([A-Za-z]:\\[^\r\n]+?)\s*$') {
      continue
    }

    $_sourcePath = $Matches[1].TrimEnd()
    $_isDirectory = $false
    if ($_sourcePath.EndsWith('\')) {
      $_isDirectory = $true
      $_sourcePath = $_sourcePath.TrimEnd('\')
    }
    if (-not $_sourcePath.StartsWith($_sourceRoot, [StringComparison]::OrdinalIgnoreCase)) {
      continue
    }

    $_relative = $_sourcePath.Substring($_sourceRoot.Length).TrimStart('\')
    $_destinationPath = $_destinationRoot
    if ($_relative) {
      $_destinationPath = Join-Path $_destinationRoot $_relative
    }
    $_limit = 259
    if ($_isDirectory) {
      $_limit = 247
    }
    if ($_destinationPath.Length -lt $_limit -or $_seen.ContainsKey($_destinationPath)) {
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

function Get-DataInventoryPlan {
  param (
    [object[]]
    $KnownFolder,

    [object[]]
    $PublicFolder,

    [object[]]
    $NetworkDrive,

    [object[]]
    $OutsideProfile,

    [object[]]
    $LongPath,

    [object]
    $Outlook
  )

  $_folder = New-Object System.Collections.ArrayList
  $_folderMap = [ordered]@{}
  $_abb = New-Object System.Collections.ArrayList
  $_manual = New-Object System.Collections.ArrayList
  $_decisions = New-Object System.Collections.ArrayList

  foreach ($_item in $KnownFolder) {
    if (-not $_item.Exists -or $_item.Bytes -eq 0) {
      continue
    }
    if ($_item.IsUnc -or $_item.IsOneDrive -or $_item.IsReparse) {
      $_decisions += "Known folder $($_item.Name) is redirected or linked: $($_item.Path)."
      continue
    }
    if ($_item.IsOutsideProfile) {
      $_folderMap[$_item.Name] = (Join-Path 'Private' $_item.Name)
      $_decisions += "Known folder $($_item.Name) is outside the profile: $($_item.Path). Map it explicitly."
      continue
    }
    $null = $_folder.Add($_item.Name)
    $null = $_abb.Add([PSCustomObject]@{ Path = $_item.Path; Reason = "Known folder $($_item.Name)" })
  }

  foreach ($_item in $PublicFolder) {
    if ($_item.Files -eq 0) {
      continue
    }
    $null = $_abb.Add([PSCustomObject]@{ Path = $_item.Path; Reason = "Public $($_item.Name)" })
    if ($_item.Name -in @('Desktop', 'Documents')) {
      $_folderMap["Public $($_item.Name)"] = (Join-Path $_item.Name "Public $($_item.Name)")
      $_decisions += "Public $($_item.Name) contains data; map it into the target profile or copy it manually."
    }
    else {
      $_decisions += "Public $($_item.Name) contains data; decide whether it belongs to this migration."
    }
    foreach ($_shortcut in $_item.Shortcuts) {
      if ($_shortcut.IsNetwork) {
        $_decisions += "Public shortcut targets a network path: $($_shortcut.Path) -> $($_shortcut.Target)."
      }
    }
  }

  foreach ($_item in $OutsideProfile) {
    if ($_item.Files -eq 0) {
      continue
    }
    $null = $_abb.Add([PSCustomObject]@{ Path = $_item.Path; Reason = 'Outside profile data' })
    $_leaf = Split-Path -Path $_item.Path -Leaf
    $_folderMap[$_leaf] = (Join-Path 'Private' $_leaf)
    $_decisions += "Data outside the profile: $($_item.Path) (probable programs: $($_item.Programs -join ', '))."
  }

  foreach ($_item in $NetworkDrive) {
    $_decisions += "Network drive $($_item.Letter): $($_item.Target). Decide whether to stage it; use -MeasureNetwork for sizes."
  }

  foreach ($_item in $LongPath) {
    $_decisions += "Path too long for the target: $($_item.Destination) ($($_item.Length)/$($_item.Limit))."
    if ($_item.WordWebPair) {
      $_decisions += "Word web-page pair needs manual renaming: $($_item.Source)."
    }
  }

  if ($null -ne $Outlook) {
    foreach ($_item in $Outlook.DataFiles) {
      $null = $_manual.Add([PSCustomObject]@{ Path = $_item.Path; Reason = "Outlook $($_item.Format) data file" })
    }
    if (-not [string]::IsNullOrWhiteSpace($Outlook.Signature)) {
      $null = $_manual.Add([PSCustomObject]@{ Path = $Outlook.Signature; Reason = 'Outlook signatures' })
    }
    foreach ($_warning in $Outlook.Warnings) {
      $_decisions += $warning
    }
    if ($Outlook.Error) {
      $_decisions += $Outlook.Error
    }
  }
  else {
    $_decisions += 'Outlook not inventoried: run Show-OutlookAccount.ps1 first and pass -OutlookReport.'
  }

  return [PSCustomObject]@{
    AbbDownload         = @($_abb.ToArray())
    Folder              = @($_folder.ToArray())
    FolderMap           = $_folderMap
    ManualCopy          = @($_manual.ToArray())
    Decisions           = @($_decisions)
    SuggestedImportCall = if ($_folder.Count -gt 0) {
      ".\scripts\Administration\Import-UserProfileData.ps1 -SourcePath 'C:\Migration\<user>' -TargetUser '<target>' -Folder $($_folder -join ',')"
    }
    else { $null }
  }
}

$_inventoryUser = Resolve-DataInventoryUser -Account $User
$_startedAt = [DateTime]::UtcNow.ToString('o')
$_runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$_reportRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportDirectory)
$_manifestPath = $null
$_inventory = $null

try {
  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Reading known folders'
  $_known = @(Get-DataInventoryKnownFolder -RegistryRoot $_inventoryUser.Registry -ProfilePath $_inventoryUser.ProfilePath)
  $_public = @(Get-DataInventoryPublicFolder -PublicRoot $env:PUBLIC -IncludePermissions $_inventoryUser.IsAdmin)
  $_network = @(Get-DataInventoryNetworkDrive -RegistryRoot $_inventoryUser.Registry -Measure $MeasureNetwork.IsPresent)
  $_outside = @()
  if ($_inventoryUser.IsAdmin) {
    Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Scanning fixed drives'
    $_outside = @(Get-DataInventoryOutsideProfile -ProfilePath $_inventoryUser.ProfilePath)
  }
  else {
    Write-Warning 'Run elevated to inspect fixed drives outside the profile.'
  }

  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Checking destination path lengths'
  $_long = New-Object System.Collections.ArrayList
  foreach ($_item in $_known) {
    if ($_item.Exists) {
      foreach ($_entry in @(Get-DataInventoryLongPath -SourceRoot $_item.Path -TargetRoot $TargetRoot)) {
        $null = $_long.Add($_entry)
      }
    }
  }
  foreach ($_item in $_outside) {
    foreach ($_entry in @(Get-DataInventoryLongPath -SourceRoot $_item.Path -TargetRoot $TargetRoot)) {
      $null = $_long.Add($_entry)
    }
  }

  Write-Progress -Id 71 -Activity 'User data inventory' -Status 'Reading Outlook evidence'
  $_signaturePath = $null
  $_signaturePath = Join-Path $_inventoryUser.ProfilePath 'AppData\Roaming\Microsoft\Signatures'
  $_outlook = Get-DataInventoryOutlookReport -Path $OutlookReport -Computer $env:COMPUTERNAME -UserName $_inventoryUser.Name -MaxAgeDays $OutlookReportMaxAge -SignaturePath $_signaturePath

  $_plan = Get-DataInventoryPlan -KnownFolder $_known -PublicFolder $_public -NetworkDrive $_network -OutsideProfile $_outside -LongPath @($_long.ToArray()) -Outlook $_outlook

  $_inventory = [PSCustomObject]@{
    PSTypeName     = 'Winkit.DataInventory'
    SchemaVersion  = 1
    RunId          = $_runId
    StartedAtUtc   = $_startedAt
    FinishedAtUtc  = [DateTime]::UtcNow.ToString('o')
    Computer       = $env:COMPUTERNAME
    User           = $_inventoryUser.Name
    UserSID        = $_inventoryUser.SID
    ProfilePath    = $_inventoryUser.ProfilePath
    TargetRoot     = $TargetRoot
    OutlookReport  = $_outlook
    KnownFolders   = $_known
    PublicFolders  = $_public
    NetworkDrives  = $_network
    OutsideProfile = $_outside
    LongPaths      = @($_long.ToArray())
    Plan           = $_plan
    ReportPath     = $null
  }

  $_hasError = @($_known | Where-Object Error).Count -gt 0 -or @($_outside | Where-Object Error).Count -gt 0
  $_status = 'Completed'
  if ($_hasError) {
    $_status = 'CompletedWithWarnings'
  }

  if (-not $DryRun) {
    $_runDirectory = Join-Path $_reportRoot $_runId
    $null = New-Item -ItemType Directory -Path $_runDirectory -Force -ErrorAction Stop
    $_manifestPath = Join-Path $_runDirectory 'manifest.json'
    $_inventory.ReportPath = $_manifestPath
    $_json = ConvertTo-Json -InputObject $_inventory -Depth 20 -ErrorAction Stop
    [IO.File]::WriteAllText($_manifestPath, $_json, (New-Object Text.UTF8Encoding($true)))
    $null = Write-OperationResultLog -Results @(New-OperationResult -Target $_inventoryUser.ProfilePath -Source 'UserData' -Action 'Inventory' -Status $_status -Detail 'User data inventory completed.' -Property @{ ReportPath = $_manifestPath }) -Path (Join-Path $_runDirectory 'operations.jsonl') -ScriptName 'Show-UserDataInventory' -RunId $_runId
  }

  Write-Log -Message "User data inventory: $_status | Known folders: $($_known.Count) | Public: $($_public.Count) | Network: $($_network.Count) | Outside: $($_outside.Count) | Long paths: $($_long.Count) | Report: $_manifestPath" -Color Green
  Write-Log -Message "Plan: ABB download $($_plan.AbbDownload.Count), manual copy $($_plan.ManualCopy.Count), decisions $($_plan.Decisions.Count)." -Color Cyan
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
