#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.4.0' }

<#
.SYNOPSIS
  Imports a GPO backup into local Group Policy or into a domain GPO.
.DESCRIPTION
  Applies a backup produced by Build-GroupPolicyBackup.ps1 (or any GPMC/LGPO
  backup holding registry policy).

  -Target Local (default) imports the backup into this machine's local Group
  Policy with LGPO.exe /g. It requires elevation and relaunches elevated when
  needed. LGPO.exe must carry a valid Microsoft Authenticode signature; if it
  is missing, PSFoundation's Install-LGPO is used to fetch it.

  -Target Domain imports the backup into a domain GPO with Import-GPO, which
  replaces every setting in the target GPO. It requires the Group Policy
  Management module (RSAT) and domain permissions, not local elevation. The
  per-domain values - target GPO name, domain, domain controller, and an
  optional OU link - are parameters. After importing, the GPO is checked for a
  principal holding the Apply Group Policy permission: a GPO without one
  applies to no computer or user, silently.

  The backup is validated before anything is applied: its registry.pol files
  must be readable and the Registry client-side extension must be registered
  for every side that carries settings. Otherwise clients would ignore them.

  When -BackupId is omitted, -Path must contain exactly one backup. With several
  backups the import stops and lists them rather than guessing which to apply.
.PARAMETER Path
  Backup root directory containing {BackupId} folders. Defaults to
  dist/GP/gpo-backups.
.PARAMETER BackupId
  ID of the backup to import (the {GUID} folder name).
.PARAMETER Target
  Local (default) or Domain.
.PARAMETER LgpoPath
  Local only. Path to LGPO.exe. Defaults to %ProgramData%\winkit\tools\LGPO.exe.
.PARAMETER DisplayName
  Domain only. Name of the target GPO. Defaults to the name recorded in the backup.
.PARAMETER Domain
  Domain only. DNS name of the target domain. Defaults to the current domain.
.PARAMETER Server
  Domain only. Domain controller to contact. Defaults to the PDC emulator.
.PARAMETER CreateIfNeeded
  Domain only. Create the target GPO when it does not exist.
.PARAMETER LinkTarget
  Domain only. Distinguished name of an OU, domain, or site to link the GPO to.
  An existing link is left unchanged.
.PARAMETER DryRun
  Validate the backup and report the planned import without applying anything.
  Does not require elevation or RSAT.
.PARAMETER PassThru
  Return structured operation results.
.PARAMETER Elevated
  Internal: set automatically on the elevated relaunch. Not for direct use.
.EXAMPLE
  PS> .\Import-GroupPolicyBackup.ps1 -DryRun
  Validates the single backup in dist/GP/gpo-backups and previews a local import.
.EXAMPLE
  PS> .\Import-GroupPolicyBackup.ps1 -BackupId 63553067-09A4-4D72-AE28-DD382A02E9FD
  Imports the given backup into local Group Policy.
.EXAMPLE
  PS> .\Import-GroupPolicyBackup.ps1 -Target Domain -DisplayName 'winkit Baseline' -CreateIfNeeded -LinkTarget 'OU=Workstations,DC=example,DC=com'
  Imports the backup into a domain GPO, creating it if needed, and links it to an OU.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported for -Target Local; -Target Domain needs the RSAT Group Policy module, which Server Core can install.
  SYSTEM-account execution: -Target Local is supported; -Target Domain needs an account with rights on the target GPO.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [string]
  $Path = (Join-Path $PSScriptRoot '..\..\dist\GP\gpo-backups'),

  [guid]
  $BackupId,

  [ValidateSet('Local', 'Domain')]
  [string]
  $Target = 'Local',

  [string]
  $LgpoPath = (Join-Path $env:ProgramData 'winkit\tools\LGPO.exe'),

  [string]
  $DisplayName,

  [string]
  $Domain,

  [string]
  $Server,

  [switch]
  $CreateIfNeeded,

  [string]
  $LinkTarget,

  [switch]
  $DryRun,

  [switch]
  $PassThru,

  # Internal: set automatically on elevated re-launch. Not for direct use.
  [switch]
  $Elevated
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - the backup is validated but not applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_registryCse = '{35378EAC-683F-11D2-A89A-00C04FBBCFA2}'

function Complete-Import {
  param ([int]$ExitCode = 0)

  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Import-GroupPolicyBackup'
  if ($_operationLog) {
    Write-Log -Message "Operation log: $_operationLog" -Color Gray
  }
  if ($PassThru -or $DryRun) {
    $_results
  }
  if ($ExitCode -eq 0 -and @($_results | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) {
    $ExitCode = 1
  }
  exit $ExitCode
}

function Add-ImportFailure {
  param ([string]$Target, [string]$Action, [string]$Detail)
  Write-Log -Message "FAILED - $Detail" -Color Red
  Add-OperationResult -Results $_results -Target $Target -Source 'GroupPolicyImport' -Action $Action -Status 'Failed' -Detail $Detail
}

# ---- Parameter combinations ----------------------------------------------------

$_domainOnly = @('DisplayName', 'Domain', 'Server', 'CreateIfNeeded', 'LinkTarget') | Where-Object { $PSBoundParameters.ContainsKey($_) }
$_localOnly = @('LgpoPath') | Where-Object { $PSBoundParameters.ContainsKey($_) }
if ($Target -eq 'Local' -and $_domainOnly) {
  Add-ImportFailure -Target 'Parameters' -Action 'Validate' -Detail "-$($_domainOnly -join ', -') only apply to -Target Domain."
  Complete-Import
}
if ($Target -eq 'Domain' -and $_localOnly) {
  Add-ImportFailure -Target 'Parameters' -Action 'Validate' -Detail '-LgpoPath only applies to -Target Local.'
  Complete-Import
}

if ($Target -eq 'Local' -and -not $DryRun -and -not (Test-Elevation)) {
  Request-AdministratorPrivilege `
    -BoundParameters    $PSBoundParameters `
    -ArgumentList       $args `
    -IsElevatedRelaunch:$Elevated
}

# ---- Select the backup -----------------------------------------------------------

$_root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (-not (Test-Path -LiteralPath $_root -PathType Container)) {
  Add-ImportFailure -Target $_root -Action 'Select' -Detail "Backup directory not found: $_root"
  Complete-Import
}

$_backups = @(
  Get-ChildItem -LiteralPath $_root -Directory |
    Where-Object { $_.Name -match '^\{[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'bkupInfo.xml') -PathType Leaf) } |
    ForEach-Object {
      $_directory = $_
      try {
        [xml]$_info = [System.IO.File]::ReadAllText((Join-Path $_directory.FullName 'bkupInfo.xml'))
        [PSCustomObject]@{
          Id          = [guid]$_directory.Name
          Directory   = $_directory.FullName
          DisplayName = ($_info.GetElementsByTagName('GPODisplayName') | Select-Object -First 1).InnerText
          BackupTime  = ($_info.GetElementsByTagName('BackupTime') | Select-Object -First 1).InnerText
        }
      }
      catch {
        Write-Log -Message "  Ignoring $($_directory.Name): unreadable bkupInfo.xml ($($_.Exception.Message))." -Color Yellow
      }
    }
)

$_candidates = ($_backups | Sort-Object BackupTime | ForEach-Object { "$($_.Id) '$($_.DisplayName)' ($($_.BackupTime) UTC)" }) -join '; '
if ($PSBoundParameters.ContainsKey('BackupId')) {
  $_backup = $_backups | Where-Object { $_.Id -eq $BackupId } | Select-Object -First 1
  if (-not $_backup) {
    Add-ImportFailure -Target "$BackupId" -Action 'Select' -Detail "Backup $BackupId not found in $_root. Available: $(if ($_candidates) { $_candidates } else { 'none' })."
    Complete-Import
  }
}
elseif ($_backups.Count -eq 0) {
  Add-ImportFailure -Target $_root -Action 'Select' -Detail "No GPO backups found in $_root."
  Complete-Import
}
elseif ($_backups.Count -gt 1) {
  Add-ImportFailure -Target $_root -Action 'Select' -Detail "$($_backups.Count) backups found; pass -BackupId to choose one: $_candidates."
  Complete-Import
}
else {
  $_backup = $_backups[0]
}

Write-Log -Message "Importing GPO backup $($_backup.Id)" -Color Cyan
Write-Log -Message "  Name   : $($_backup.DisplayName)" -Color Gray
Write-Log -Message "  Built  : $($_backup.BackupTime) UTC" -Color Gray
Write-Log -Message "  Target : $Target" -Color Gray

# ---- Validate the backup -----------------------------------------------------------

$_counts = @{ Machine = 0; User = 0 }
try {
  [xml]$_backupXml = [System.IO.File]::ReadAllText((Join-Path $_backup.Directory 'Backup.xml'))
  foreach ($_side in @('Machine', 'User')) {
    $_extensions = $_backupXml.GetElementsByTagName("$($_side)ExtensionGuids") | Select-Object -First 1
    $_registered = $_extensions -and $_extensions.InnerText.IndexOf($_registryCse, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    $_polPath = Join-Path $_backup.Directory "DomainSysvol\GPO\$_side\registry.pol"
    $_hasPol = Test-Path -LiteralPath $_polPath -PathType Leaf

    if ($_hasPol) {
      $_counts[$_side] = @(ConvertFrom-RegistryPolicy -Path $_polPath).Count
    }
    if ($_registered -and -not $_hasPol) {
      throw "Backup.xml registers the Registry extension for $_side, but $_side\registry.pol is missing."
    }
    if ($_hasPol -and $_counts[$_side] -gt 0 -and -not $_registered) {
      throw "$_side\registry.pol holds $($_counts[$_side]) record(s), but Backup.xml does not register the Registry extension for $_side - clients would ignore them."
    }
  }
}
catch {
  Add-ImportFailure -Target "$($_backup.Id)" -Action 'Validate' -Detail "Invalid backup: $($_.Exception.Message)"
  Complete-Import
}

if (($_counts.Machine + $_counts.User) -eq 0) {
  Add-ImportFailure -Target "$($_backup.Id)" -Action 'Validate' -Detail 'The backup contains no registry policy records.'
  Complete-Import
}

$_summary = "$($_counts.Machine) machine, $($_counts.User) user record(s)"
Write-Log -Message "  Records: $_summary" -Color Gray
Add-OperationResult -Results $_results -Target "$($_backup.Id)" -Source 'GroupPolicyImport' -Action 'Validate' -Status 'Completed' -Detail $_summary -Property @{
  DisplayName    = $_backup.DisplayName
  MachineRecords = $_counts.Machine
  UserRecords    = $_counts.User
}

# ---- Local import ------------------------------------------------------------------

if ($Target -eq 'Local') {
  $_lgpo = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LgpoPath)

  if ($DryRun) {
    if (-not (Test-Path -LiteralPath $_lgpo -PathType Leaf)) {
      $_detail = "LGPO.exe not found at $_lgpo; a real run tries Install-LGPO first."
      Write-Log -Message "  WARNING - $_detail" -Color Yellow
      Add-OperationResult -Results $_results -Target $_lgpo -Source 'GroupPolicyImport' -Action 'ResolveLgpo' -Status 'Warn' -Detail $_detail
    }
    Write-Log -Message "[DRY RUN] Would import $_summary into local Group Policy with LGPO.exe /g." -Color Yellow
    Add-OperationResult -Results $_results -Target 'LocalGroupPolicy' -Source 'GroupPolicyImport' -Action 'Import' -Status 'Skipped' -Detail "DryRun: $_summary from backup $($_backup.Id)."
    Complete-Import
  }

  try {
    if (-not (Test-Path -LiteralPath $_lgpo -PathType Leaf)) {
      Write-Log -Message "  LGPO.exe not found at $_lgpo; running Install-LGPO." -Color Yellow
      $_lgpo = Install-LGPO -Destination (Split-Path -Path $_lgpo -Parent) -ErrorAction Stop
    }
  }
  catch {
    Add-ImportFailure -Target $_lgpo -Action 'ResolveLgpo' -Detail "LGPO.exe not found and Install-LGPO failed: $($_.Exception.Message) Place a Microsoft-signed LGPO.exe at $_lgpo or pass -LgpoPath."
    Complete-Import
  }

  $_signature = Get-AuthenticodeSignature -FilePath $_lgpo
  if ($_signature.Status -ne 'Valid' -or $_signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
    Add-ImportFailure -Target $_lgpo -Action 'ResolveLgpo' -Detail "Refusing to run ${_lgpo}: expected a valid Microsoft Authenticode signature, got '$($_signature.Status)'."
    Complete-Import
  }

  if ($PSCmdlet.ShouldProcess('local Group Policy', "Import backup $($_backup.Id) '$($_backup.DisplayName)' ($_summary)")) {
    # LGPO /g imports every backup under the given path, so stage the selected
    # backup on its own rather than pointing LGPO at the shared backup root.
    $_stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) "winkit-gpo-import-$([guid]::NewGuid().ToString('N'))"
    try {
      $null = New-Item -ItemType Directory -Path $_stagingRoot -Force
      Copy-Item -LiteralPath $_backup.Directory -Destination $_stagingRoot -Recurse -ErrorAction Stop
      $_apply = Invoke-LGPO -PolicyPath $_stagingRoot -LgpoExe $_lgpo -WhatIf:$false -Confirm:$false
      if ($_apply.Success) {
        Write-Log -Message "Imported $_summary into local Group Policy." -Color Green
        Add-OperationResult -Results $_results -Target 'LocalGroupPolicy' -Source 'GroupPolicyImport' -Action 'Import' -Status 'Completed' -Detail "Backup $($_backup.Id): $_summary." -Property @{
          BackupId = "$($_backup.Id)"
          ExitCode = $_apply.ExitCode
        }
      }
      else {
        $_output = (@($_apply.StdErr, $_apply.StdOut) | Where-Object { $_ }) -join ' '
        Add-ImportFailure -Target 'LocalGroupPolicy' -Action 'Import' -Detail "LGPO.exe exited with code $($_apply.ExitCode): $($_output.Trim())"
      }
    }
    catch {
      Add-ImportFailure -Target 'LocalGroupPolicy' -Action 'Import' -Detail $_.Exception.Message
    }
    finally {
      if (Test-Path -LiteralPath $_stagingRoot) {
        Remove-Item -LiteralPath $_stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
      }
    }
  }
  else {
    Add-OperationResult -Results $_results -Target 'LocalGroupPolicy' -Source 'GroupPolicyImport' -Action 'Import' -Status 'Skipped' -Detail 'WhatIf'
  }

  Complete-Import
}

# ---- Domain import -------------------------------------------------------------------

$_gpoName = if ($DisplayName) { $DisplayName } else { $_backup.DisplayName }
$_domainParams = @{}
if ($Domain) { $_domainParams['Domain'] = $Domain }
if ($Server) { $_domainParams['Server'] = $Server }
$_where = if ($Domain) { " in $Domain" } else { '' }

if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
  $_detail = 'The GroupPolicy module (RSAT Group Policy Management Tools) is not installed.'
  if (-not $DryRun) {
    Add-ImportFailure -Target 'GroupPolicy' -Action 'Import' -Detail $_detail
    Complete-Import
  }
  Write-Log -Message "  WARNING - $_detail A real run would fail." -Color Yellow
  Add-OperationResult -Results $_results -Target 'GroupPolicy' -Source 'GroupPolicyImport' -Action 'Import' -Status 'Warn' -Detail $_detail
}

if ($DryRun) {
  $_create = if ($CreateIfNeeded) { ', creating it if needed' } else { '' }
  Write-Log -Message "[DRY RUN] Would replace all settings in GPO '$_gpoName'$_where with backup $($_backup.Id)$_create." -Color Yellow
  Add-OperationResult -Results $_results -Target $_gpoName -Source 'GroupPolicyImport' -Action 'Import' -Status 'Skipped' -Detail "DryRun: Import-GPO from backup $($_backup.Id)$_create."
  if ($LinkTarget) {
    Write-Log -Message "[DRY RUN] Would link '$_gpoName' to $LinkTarget." -Color Yellow
    Add-OperationResult -Results $_results -Target $LinkTarget -Source 'GroupPolicyImport' -Action 'Link' -Status 'Skipped' -Detail "DryRun: link '$_gpoName'."
  }
  Complete-Import
}

if (-not $PSCmdlet.ShouldProcess("GPO '$_gpoName'$_where", "Replace all settings with backup $($_backup.Id) ($_summary)")) {
  Add-OperationResult -Results $_results -Target $_gpoName -Source 'GroupPolicyImport' -Action 'Import' -Status 'Skipped' -Detail 'WhatIf'
  Complete-Import
}

try {
  Import-Module GroupPolicy -ErrorAction Stop
  $null = Import-GPO -BackupId $_backup.Id -Path $_root -TargetName $_gpoName -CreateIfNeeded:$CreateIfNeeded @_domainParams -ErrorAction Stop
  Write-Log -Message "Imported $_summary into GPO '$_gpoName'$_where." -Color Green
  Add-OperationResult -Results $_results -Target $_gpoName -Source 'GroupPolicyImport' -Action 'Import' -Status 'Completed' -Detail "Backup $($_backup.Id): $_summary." -Property @{
    BackupId = "$($_backup.Id)"
  }
}
catch {
  Add-ImportFailure -Target $_gpoName -Action 'Import' -Detail "Import-GPO failed: $($_.Exception.Message)"
  Complete-Import
}

try {
  $_appliers = @(Get-GPPermission -Name $_gpoName -All @_domainParams -ErrorAction Stop | Where-Object { "$($_.Permission)" -eq 'GpoApply' })
  if ($_appliers.Count -eq 0) {
    $_detail = "No principal holds the Apply Group Policy permission on '$_gpoName', so it applies to no computer or user. Grant it, e.g. Set-GPPermission -Name '$_gpoName' -TargetName 'Authenticated Users' -TargetType Group -PermissionLevel GpoApply."
    Write-Log -Message "  WARNING - $_detail" -Color Yellow
    Add-OperationResult -Results $_results -Target $_gpoName -Source 'GroupPolicyImport' -Action 'CheckPermission' -Status 'Warn' -Detail $_detail
  }
  else {
    $_trustees = ($_appliers | ForEach-Object { $_.Trustee.Name }) -join ', '
    Add-OperationResult -Results $_results -Target $_gpoName -Source 'GroupPolicyImport' -Action 'CheckPermission' -Status 'Completed' -Detail "Applies to: $_trustees."
  }
}
catch {
  $_detail = "Could not read permissions on '${_gpoName}': $($_.Exception.Message)"
  Write-Log -Message "  WARNING - $_detail" -Color Yellow
  Add-OperationResult -Results $_results -Target $_gpoName -Source 'GroupPolicyImport' -Action 'CheckPermission' -Status 'Warn' -Detail $_detail
}

if ($LinkTarget) {
  try {
    $_existing = @((Get-GPInheritance -Target $LinkTarget @_domainParams -ErrorAction Stop).GpoLinks | Where-Object { $_.DisplayName -eq $_gpoName })
    if ($_existing.Count -gt 0) {
      Write-Log -Message "  '$_gpoName' is already linked to $LinkTarget." -Color Gray
      Add-OperationResult -Results $_results -Target $LinkTarget -Source 'GroupPolicyImport' -Action 'Link' -Status 'Skipped' -Detail 'Already linked.'
    }
    elseif ($PSCmdlet.ShouldProcess($LinkTarget, "Link GPO '$_gpoName'")) {
      $null = New-GPLink -Name $_gpoName -Target $LinkTarget @_domainParams -ErrorAction Stop
      Write-Log -Message "Linked '$_gpoName' to $LinkTarget." -Color Green
      Add-OperationResult -Results $_results -Target $LinkTarget -Source 'GroupPolicyImport' -Action 'Link' -Status 'Completed' -Detail "Linked '$_gpoName'."
    }
    else {
      Add-OperationResult -Results $_results -Target $LinkTarget -Source 'GroupPolicyImport' -Action 'Link' -Status 'Skipped' -Detail 'WhatIf'
    }
  }
  catch {
    Add-ImportFailure -Target $LinkTarget -Action 'Link' -Detail "Linking failed: $($_.Exception.Message)"
  }
}

Complete-Import
