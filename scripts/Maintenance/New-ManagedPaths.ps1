#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.4.0' }

<#
.SYNOPSIS
  Creates and repairs the standard C:\Managed folder tree.

.DESCRIPTION
  Creates Archives, Backups, Checkpoints, Media, Reports, Tools, Audit,
  Installables, and Migration below the managed root with explicit,
  non-inherited permissions.

  Administrators and SYSTEM receive Full Control on every folder. Tools, Media,
  and Installables also grant Users Read & Execute. Archives, Backups,
  Checkpoints, Reports, Audit, and Migration grant no access to Users.

  Existing folders are reported first. Their permissions are changed only with
  -RepairPermissions. -GrantUser and -Subfolder create a per-user folder with
  Modify access for that account and traverse access on its parent.

.PARAMETER Path
  Managed root. Defaults to C:\Managed.

.PARAMETER ReportDirectory
  Parent for unique report directories. Defaults to
  ProgramData\winkit\ManagedPathReports. No report is written during -DryRun.

.PARAMETER GrantUser
  Account that receives Modify access to a per-user subfolder. Requires
  -Subfolder.

.PARAMETER Subfolder
  Managed folder that receives the per-user child folder.

.PARAMETER RepairPermissions
  Replace non-conforming permissions on existing managed folders. Without this
  switch, differences are reported but not changed.

.PARAMETER DryRun
  Show planned folder creation and permission changes without writing anything.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> .\New-ManagedPaths.ps1 -DryRun
  Preview the managed tree and permission changes.

.EXAMPLE
  PS> .\New-ManagedPaths.ps1 -PassThru
  Create the managed tree and write a JSON report.

.EXAMPLE
  PS> .\New-ManagedPaths.ps1 -GrantUser 'CONTOSO\User1' -Subfolder Archives -PassThru
  Create Archives\User1 with Modify access for that account.

.EXAMPLE
  PS> .\New-ManagedPaths.ps1 -RepairPermissions -PassThru
  Repair non-conforming permissions on an existing managed tree.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - local filesystem ACL operations only.
  SYSTEM-account execution: supported.
  Work is not transactional: each folder and ACL is applied separately.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [ValidateNotNullOrEmpty()]
  [string]
  $Path = (Join-Path $env:SystemDrive 'Managed'),

  [ValidateNotNullOrEmpty()]
  [string]
  $ReportDirectory = (Join-Path $env:ProgramData 'winkit\ManagedPathReports'),

  [string]
  $GrantUser,

  [ValidateSet('Archives', 'Backups', 'Checkpoints', 'Media', 'Reports', 'Tools', 'Audit', 'Installables', 'Migration')]
  [string]
  $Subfolder,

  [switch]
  $RepairPermissions,

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

if ([bool]$GrantUser -ne [bool]$Subfolder) {
  throw 'GrantUser and Subfolder must be supplied together.'
}

$_startedAt = [DateTime]::UtcNow.ToString('o')
$_runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
$_root = [IO.Path]::GetFullPath($Path).TrimEnd('\')
$_reportRoot = [IO.Path]::GetFullPath($ReportDirectory).TrimEnd('\')
$_results = New-Object System.Collections.ArrayList
$_created = 0
$_repaired = 0
$_failed = 0
$_needsRepair = 0
$_skipped = 0
$_manifestPath = $null
$_grantSid = $null
$_grantLeaf = $null

$_folders = @(
  @{ Name = 'Archives'; Users = 'None' }
  @{ Name = 'Backups'; Users = 'None' }
  @{ Name = 'Checkpoints'; Users = 'None' }
  @{ Name = 'Media'; Users = 'ReadAndExecute' }
  @{ Name = 'Reports'; Users = 'None' }
  @{ Name = 'Tools'; Users = 'ReadAndExecute' }
  @{ Name = 'Audit'; Users = 'None' }
  @{ Name = 'Installables'; Users = 'ReadAndExecute' }
  @{ Name = 'Migration'; Users = 'None' }
)

if ($GrantUser) {
  try {
    if ($GrantUser -match '^S-1-5-') {
      $_grantSid = New-Object Security.Principal.SecurityIdentifier($GrantUser)
      $_grantLeaf = ($_grantSid.Translate([Security.Principal.NTAccount])).Value.Split('\')[-1]
    }
    else {
      $_account = New-Object Security.Principal.NTAccount($GrantUser)
      $_grantSid = $_account.Translate([Security.Principal.SecurityIdentifier])
      $_grantLeaf = $GrantUser.Split('\')[-1]
    }
  }
  catch {
    throw "Could not resolve GrantUser '$GrantUser': $($_.Exception.Message)"
  }
}

if ($_grantLeaf -and $_grantLeaf -notmatch '^[^\\/:*?"<>|]+$') {
  throw "The resolved GrantUser leaf name is not a safe folder name: $_grantLeaf"
}

function New-ManagedPathAcl {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Builds an in-memory ACL; the caller handles ShouldProcess.')]
  param (
    [ValidateSet('None', 'ReadAndExecute')]
    [string]
    $UsersAccess = 'None',

    [bool]
    $UsersInherit = $true,

    [Security.Principal.SecurityIdentifier]
    $GrantSid,

    [ValidateSet('None', 'Modify', 'Traverse')]
    [string]
    $GrantAccess = 'None',

    [bool]
    $GrantInherit = $true
  )

  $_acl = New-Object Security.AccessControl.DirectorySecurity
  $_acl.SetAccessRuleProtection($true, $false)
  $_inherit = [Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [Security.AccessControl.InheritanceFlags]::ObjectInherit
  $_none = [Security.AccessControl.PropagationFlags]::None
  $_acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule([Security.Principal.SecurityIdentifier]'S-1-5-32-544', 'FullControl', $_inherit, $_none, 'Allow')))
  $_acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule([Security.Principal.SecurityIdentifier]'S-1-5-18', 'FullControl', $_inherit, $_none, 'Allow')))

  if ($UsersAccess -ne 'None') {
    $_userFlags = $_inherit
    if (-not $UsersInherit) {
      $_userFlags = [Security.AccessControl.InheritanceFlags]::None
    }
    $_acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule([Security.Principal.SecurityIdentifier]'S-1-5-32-545', $UsersAccess, $_userFlags, $_none, 'Allow')))
  }

  if ($null -ne $GrantSid -and $GrantAccess -ne 'None') {
    $_grantFlags = $_inherit
    if (-not $GrantInherit) {
      $_grantFlags = [Security.AccessControl.InheritanceFlags]::None
    }
    $_acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($GrantSid, $GrantAccess, $_grantFlags, $_none, 'Allow')))
  }

  return $_acl
}

function Get-ManagedPathAclSignature {
  param (
    [Security.AccessControl.DirectorySecurity]
    $Acl
  )

  $_rules = @($Acl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]) | ForEach-Object {
      '{0}|{1}|{2}|{3}|{4}' -f $_.IdentityReference.Value, [int]$_.FileSystemRights, $_.AccessControlType, [int]$_.InheritanceFlags, [int]$_.PropagationFlags
    } | Sort-Object)
  return ($_rules -join ';')
}

function Test-ManagedPathAcl {
  param (
    [string]
    $Folder,

    [Security.AccessControl.DirectorySecurity]
    $Desired
  )

  $_current = Get-Acl -LiteralPath $Folder -ErrorAction Stop
  if (-not $_current.AreAccessRulesProtected) {
    return $false
  }
  return ((Get-ManagedPathAclSignature -Acl $_current) -eq (Get-ManagedPathAclSignature -Acl $Desired))
}

function Add-ManagedPathResult {
  param (
    [string]
    $Target,

    [string]
    $Kind,

    [string]
    $Status,

    [string]
    $Detail
  )

  Add-OperationResult -Results $_results -Target $Target -Source 'ManagedPaths' -Action 'Configure' -Status $Status -Detail $Detail -Property @{
    Kind = $Kind
  }
}

function Set-ManagedPath {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'The caller handles ShouldProcess before invoking this helper.')]
  param (
    [string]
    $Folder,

    [Security.AccessControl.DirectorySecurity]
    $Acl,

    [bool]
    $Exists
  )

  if (-not $Exists) {
    $null = New-Item -ItemType Directory -Path $Folder -Force -ErrorAction Stop
  }
  Set-Acl -LiteralPath $Folder -AclObject $Acl -ErrorAction Stop
}

# ---- Root and standard folders ---------------------------------------------
$_rootExists = Test-Path -LiteralPath $_root -PathType Container -ErrorAction Stop
$_rootAcl = New-ManagedPathAcl -UsersAccess 'ReadAndExecute' -UsersInherit $false -GrantSid $_grantSid -GrantAccess $(if ($_grantSid) { 'Traverse' } else { 'None' }) -GrantInherit $false

if (-not $_rootExists) {
  if ($DryRun) {
    $_skipped++
    Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Skipped' -Detail 'DryRun: would create the managed root and apply its permissions.'
  }
  elseif ($PSCmdlet.ShouldProcess($_root, 'Create managed root')) {
    try {
      Set-ManagedPath -Folder $_root -Acl $_rootAcl -Exists $false
      $_created++
      Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Completed' -Detail 'Created the managed root with explicit permissions.'
    }
    catch {
      $_failed++
      Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Failed' -Detail "Could not create the managed root: $($_.Exception.Message)"
    }
  }
  else {
    $_skipped++
    Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Skipped' -Detail 'WhatIf: root creation was declined.'
  }
}
elseif (Test-ManagedPathAcl -Folder $_root -Desired $_rootAcl) {
  $_skipped++
  Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Skipped' -Detail 'Root permissions already conform.'
}
elseif ($RepairPermissions -or $DryRun) {
  if ($DryRun) {
    $_skipped++
    $_needsRepair++
    Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Skipped' -Detail 'DryRun: root permissions differ and would be repaired.'
  }
  elseif ($PSCmdlet.ShouldProcess($_root, 'Repair managed root permissions')) {
    try {
      Set-Acl -LiteralPath $_root -AclObject $_rootAcl -ErrorAction Stop
      $_repaired++
      Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Completed' -Detail 'Repaired root permissions.'
    }
    catch {
      $_failed++
      Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Failed' -Detail "Could not repair root permissions: $($_.Exception.Message)"
    }
  }
  else {
    $_skipped++
    Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Skipped' -Detail 'WhatIf: root permission repair was declined.'
  }
}
else {
  $_needsRepair++
  Add-ManagedPathResult -Target $_root -Kind 'Root' -Status 'Skipped' -Detail 'Root permissions differ. Use -RepairPermissions to change them.'
}

foreach ($_folder in $_folders) {
  $_folderPath = Join-Path $_root $_folder.Name
  $_isGrantParent = $_grantSid -and $Subfolder -eq $_folder.Name
  $_folderAcl = New-ManagedPathAcl -UsersAccess $_folder.Users -GrantSid $(if ($_isGrantParent) { $_grantSid } else { $null }) -GrantAccess $(if ($_isGrantParent) { 'Traverse' } else { 'None' }) -GrantInherit $false
  $_exists = Test-Path -LiteralPath $_folderPath -PathType Container -ErrorAction Stop

  if (-not $_exists) {
    if ($DryRun) {
      $_skipped++
      Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Skipped' -Detail 'DryRun: would create the folder and apply its permissions.'
    }
    elseif ($PSCmdlet.ShouldProcess($_folderPath, "Create managed folder '$($_folder.Name)'")) {
      try {
        Set-ManagedPath -Folder $_folderPath -Acl $_folderAcl -Exists $false
        $_created++
        Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Completed' -Detail 'Created with explicit permissions.'
      }
      catch {
        $_failed++
        Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Failed' -Detail "Could not create the folder: $($_.Exception.Message)"
      }
    }
    else {
      $_skipped++
      Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Skipped' -Detail 'WhatIf: folder creation was declined.'
    }
  }
  elseif (Test-ManagedPathAcl -Folder $_folderPath -Desired $_folderAcl) {
    $_skipped++
    Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Skipped' -Detail 'Permissions already conform.'
  }
  elseif ($RepairPermissions -or $DryRun) {
    if ($DryRun) {
      $_skipped++
      $_needsRepair++
      Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Skipped' -Detail 'DryRun: permissions differ and would be repaired.'
    }
    elseif ($PSCmdlet.ShouldProcess($_folderPath, "Repair managed folder permissions '$($_folder.Name)'")) {
      try {
        Set-Acl -LiteralPath $_folderPath -AclObject $_folderAcl -ErrorAction Stop
        $_repaired++
        Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Completed' -Detail 'Repaired folder permissions.'
      }
      catch {
        $_failed++
        Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Failed' -Detail "Could not repair folder permissions: $($_.Exception.Message)"
      }
    }
    else {
      $_skipped++
      Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Skipped' -Detail 'WhatIf: permission repair was declined.'
    }
  }
  else {
    $_needsRepair++
    Add-ManagedPathResult -Target $_folderPath -Kind $_folder.Name -Status 'Skipped' -Detail 'Permissions differ. Use -RepairPermissions to change them.'
  }
}

# ---- Per-user subfolder ----------------------------------------------------
if ($_grantSid) {
  $_parent = Join-Path $_root $Subfolder
  $_child = Join-Path $_parent $_grantLeaf
  $_childAcl = New-ManagedPathAcl -GrantSid $_grantSid -GrantAccess 'Modify' -GrantInherit $true
  $_childExists = Test-Path -LiteralPath $_child -PathType Container -ErrorAction Stop

  if (-not $_childExists) {
    if ($DryRun) {
      $_skipped++
      Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Skipped' -Detail "DryRun: would create '$_child' with Modify access for $($_grantSid.Value)."
    }
    elseif ($PSCmdlet.ShouldProcess($_child, "Create managed user folder '$_grantLeaf'")) {
      try {
        Set-ManagedPath -Folder $_child -Acl $_childAcl -Exists $false
        $_created++
        Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Completed' -Detail "Created with Modify access for $($_grantSid.Value)."
      }
      catch {
        $_failed++
        Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Failed' -Detail "Could not create the per-user folder: $($_.Exception.Message)"
      }
    }
    else {
      $_skipped++
      Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Skipped' -Detail 'WhatIf: per-user folder creation was declined.'
    }
  }
  elseif (Test-ManagedPathAcl -Folder $_child -Desired $_childAcl) {
    $_skipped++
    Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Skipped' -Detail 'Per-user permissions already conform.'
  }
  elseif ($RepairPermissions -or $DryRun) {
    if ($DryRun) {
      $_skipped++
      $_needsRepair++
      Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Skipped' -Detail 'DryRun: per-user permissions differ and would be repaired.'
    }
    elseif ($PSCmdlet.ShouldProcess($_child, "Repair managed user folder permissions '$_grantLeaf'")) {
      try {
        Set-Acl -LiteralPath $_child -AclObject $_childAcl -ErrorAction Stop
        $_repaired++
        Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Completed' -Detail 'Repaired per-user permissions.'
      }
      catch {
        $_failed++
        Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Failed' -Detail "Could not repair per-user permissions: $($_.Exception.Message)"
      }
    }
    else {
      $_skipped++
      Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Skipped' -Detail 'WhatIf: per-user permission repair was declined.'
    }
  }
  else {
    $_needsRepair++
    Add-ManagedPathResult -Target $_child -Kind "Grant:$Subfolder" -Status 'Skipped' -Detail 'Per-user permissions differ. Use -RepairPermissions to change them.'
  }
}

# ---- Summary and report ----------------------------------------------------
if ($_failed -gt 0) {
  $_status = 'Failed'
}
elseif ($DryRun) {
  $_status = 'Preview'
}
elseif ($_created -gt 0 -or $_repaired -gt 0) {
  $_status = 'Completed'
}
else {
  $_status = 'Skipped'
}

if (-not $DryRun) {
  $_manifestPath = Join-Path (Join-Path $_reportRoot $_runId) 'manifest.json'
}

Add-OperationResult -Results $_results -Target $_root -Source 'ManagedPaths' -Action 'NewManagedPaths' -Status $_status -Detail "Managed path configuration finished with $_created creation(s), $_repaired repair(s), $_needsRepair permission difference(s), and $_failed failure(s)." -Property @{
  RunId        = $_runId
  ManifestPath = $_manifestPath
  Created      = $_created
  Repaired     = $_repaired
  NeedsRepair  = $_needsRepair
  Failed       = $_failed
  Skipped      = $_skipped
  GrantUser    = $_grantSid.Value
  GrantLeaf    = $_grantLeaf
  Subfolder    = $Subfolder
  DryRun       = [bool]$DryRun
}

if (-not $DryRun) {
  try {
    $_runDirectory = Join-Path $_reportRoot $_runId
    $null = New-Item -ItemType Directory -Path $_runDirectory -Force -ErrorAction Stop
    $_manifest = [ordered]@{
      SchemaVersion     = 1
      RunId             = $_runId
      StartedAtUtc      = $_startedAt
      FinishedAtUtc     = [DateTime]::UtcNow.ToString('o')
      Actor             = [Security.Principal.WindowsIdentity]::GetCurrent().Name
      Machine           = $env:COMPUTERNAME
      Path              = $_root
      GrantUser         = $_grantSid.Value
      GrantLeaf         = $_grantLeaf
      Subfolder         = $Subfolder
      RepairPermissions = [bool]$RepairPermissions
      Status            = $_status
      Results           = @($_results.ToArray())
      Created           = $_created
      Repaired          = $_repaired
      NeedsRepair       = $_needsRepair
      Failed            = $_failed
      Skipped           = $_skipped
    }
    $_json = ConvertTo-Json -InputObject $_manifest -Depth 12 -ErrorAction Stop
    [IO.File]::WriteAllText($_manifestPath, $_json, (New-Object Text.UTF8Encoding($true)))
    $null = Write-OperationResultLog -Results $_results -Path (Join-Path $_runDirectory 'operations.jsonl') -ScriptName 'New-ManagedPaths' -RunId $_runId
  }
  catch {
    $_status = 'Failed'
    $_failed++
    $_manifestPath = $null
    Add-OperationResult -Results $_results -Target $_reportRoot -Source 'ManagedPaths' -Action 'Report' -Status 'Failed' -Detail "Could not write the managed path report: $($_.Exception.Message)"
    Write-Warning "Could not write the managed path report: $($_.Exception.Message)"
  }
}

Write-Log -Message "Managed paths: $_status | Created: $_created | Repaired: $_repaired | NeedsRepair: $_needsRepair | Failed: $_failed | Report: $_manifestPath" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
if ($PassThru -or $WhatIfPreference) {
  $_results
}
if ($_failed -gt 0) {
  exit 1
}
exit 0
