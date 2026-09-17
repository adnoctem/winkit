#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.4.0' }

<#
.SYNOPSIS
  Compiles winkit policy text sources into a domain-importable GPO backup.
.DESCRIPTION
  Reads LGPO text policy sources (resources/policies/*.txt), compiles them into
  Machine and User registry.pol files with PSFoundation's
  ConvertTo-RegistryPolicy, and wraps them in a GPO backup stamped from the
  vendored skeleton templates (resources/policies/skeleton/). The result imports
  into a domain with Import-GPO, or locally with Import-GroupPolicyBackup.ps1.

  LGPO.exe is not required. The source parser follows LGPO's own text format
  rules (verified against LGPO 3.0), with one deliberate difference: input that
  LGPO silently corrupts - out-of-range integers, non-ASCII data in a file
  without a byte-order mark - fails the build instead.

  Scope is taken from each entry's Computer/User line. (LGPO.exe /r ignores it
  and writes both scopes into a single file, which is why this script does not
  delegate to it.)

  Every build gets a fresh GPO GUID and backup ID. Domain provenance fields are
  stamped with synthetic values under the reserved .invalid TLD; Import-GPO
  imports into whatever domain it targets, so they are informational only.

  Sources are validated and the backup is assembled in a temporary staging
  folder first. Nothing is written to -OutputPath unless every source compiles
  and the stamped backup passes validation.
.PARAMETER SourcePath
  Directory containing the policy text sources. Defaults to resources/policies.
.PARAMETER OutputPath
  Directory that receives the backup folder and its manifest.xml entry. Defaults
  to dist/GP/gpo-backups. Existing backups in the directory are kept.
.PARAMETER DisplayName
  GPO display name recorded in the backup. Defaults to 'winkit Baseline'.
.PARAMETER SkeletonPath
  Directory holding the Backup.xml and bkupInfo.xml templates. Defaults to
  resources/policies/skeleton.
.PARAMETER Exclude
  Wildcard patterns for source file names to skip. Defaults to '*-example.txt'
  so the documentation example is never compiled into a baseline.
.PARAMETER DryRun
  Compile and validate in staging without writing to -OutputPath.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> .\Build-GroupPolicyBackup.ps1
  Compiles resources/policies/*.txt into dist/GP/gpo-backups.
.EXAMPLE
  PS> .\Build-GroupPolicyBackup.ps1 -DisplayName 'winkit Baseline 2026-09' -DryRun
  Validates the sources and previews the backup without writing it.
.EXAMPLE
  PS> .\Build-GroupPolicyBackup.ps1 -SourcePath .\dist\GP\exports -OutputPath .\dist\GP\review
  Compiles an Export-LocalPolicy capture into a separate review folder.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - pure file processing, no GUI or Group Policy tooling required.
  SYSTEM-account execution: supported; no elevation is needed unless -OutputPath requires it.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Build-GroupPolicyBackup is the intended entry verb for this build pipeline.')]
[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [string]
  $SourcePath = (Join-Path $PSScriptRoot '..\..\resources\policies'),

  [string]
  $OutputPath = (Join-Path $PSScriptRoot '..\..\dist\GP\gpo-backups'),

  [ValidateScript({
      if ($_.Length -gt 255) { throw 'DisplayName must be 255 characters or fewer.' }
      if ($_ -match '[\x00-\x1F]|\]\]>|\{\{|\}\}') { throw 'DisplayName must not contain control characters, "]]>", "{{" or "}}".' }
      $true
    })]
  [ValidateNotNullOrEmpty()]
  [string]
  $DisplayName = 'winkit Baseline',

  [string]
  $SkeletonPath = (Join-Path $PSScriptRoot '..\..\resources\policies\skeleton'),

  [string[]]
  $Exclude = @('*-example.txt'),

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - the backup is compiled and validated in staging only`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# Synthetic provenance. Import-GPO targets the domain passed to it, so these only
# have to be well-formed; .invalid is reserved and can never resolve.
$_provenance = @{
  GPODomain           = 'winkit.invalid'
  GPODomainController = 'build.winkit.invalid'
  NetBIOSDomainName   = 'WINKIT'
  DomainSid           = 'S-1-5-21-1000000000-2000000000-3000000000'
}

# Registry client-side extension plus the Administrative Templates tool
# extension for each side. Without these the GPO carries registry.pol files that
# clients never process.
$_registryCse = '{35378EAC-683F-11D2-A89A-00C04FBBCFA2}'
$_machineToolExtension = '{D02B1F72-3407-48AE-BA88-E8213C6761F1}'
$_userToolExtension = '{D02B1F73-3407-48AE-BA88-E8213C6761F1}'

function ConvertTo-LgpoInteger {
  param (
    [string]
    $Text,

    [ValidateSet(32, 64)]
    [int]
    $Bits
  )

  $_max = if ($Bits -eq 32) { [decimal][uint32]::MaxValue } else { [decimal][uint64]::MaxValue }
  $_modulus = $_max + 1

  if ($Text -match '^0[xX]([0-9a-fA-F]+)$') {
    $_digits = $Matches[1].TrimStart('0')
    if ($_digits.Length -gt ($Bits / 4)) {
      throw "Hexadecimal value '$Text' does not fit in $Bits bits."
    }
    if ($_digits.Length -eq 0) { $_digits = '0' }
    $_value = [decimal][uint64]::Parse($_digits, [Globalization.NumberStyles]::HexNumber, [Globalization.CultureInfo]::InvariantCulture)
  }
  elseif ($Text -match '^-?\d+$') {
    $_value = [decimal]::Parse($Text, [Globalization.CultureInfo]::InvariantCulture)
    # LGPO wraps out-of-range values silently (DWORD:4294967296 becomes 0);
    # refuse them instead. Negative values map to two's complement as in LGPO.
    if ($_value -gt $_max -or $_value -lt (0 - ($_modulus / 2))) {
      throw "Value '$Text' is out of range for a $Bits-bit integer."
    }
    if ($_value -lt 0) { $_value = $_modulus + $_value }
  }
  else {
    throw "Value '$Text' is not a decimal or 0x-prefixed hexadecimal integer."
  }

  if ($Bits -eq 32) { return [uint32]$_value }
  return [uint64]$_value
}

function ConvertFrom-LgpoString {
  # LGPO text escaping: "\\" is a backslash and "\0" is a NUL character, which
  # separates MULTISZ strings. Any other backslash sequence is kept literally.
  param (
    [string]
    $Text,

    [switch]
    $Multi
  )

  $_builder = New-Object System.Text.StringBuilder
  $_parts = New-Object System.Collections.Generic.List[string]

  for ($_i = 0; $_i -lt $Text.Length; $_i++) {
    $_char = $Text[$_i]
    if ($_char -eq '\' -and ($_i + 1) -lt $Text.Length) {
      $_next = $Text[$_i + 1]
      if ($_next -eq '\') {
        [void]$_builder.Append('\')
        $_i++
        continue
      }
      if ($_next -eq '0') {
        if (-not $Multi) {
          # LGPO stores an embedded NUL here, which truncates the string for
          # most readers - almost always an unescaped path such as C:\0data.
          throw 'SZ/EXSZ value contains "\0", which LGPO stores as a NUL character. Write "\\0" for a literal backslash followed by 0.'
        }
        $_parts.Add($_builder.ToString())
        [void]$_builder.Clear()
        $_i++
        continue
      }
    }
    [void]$_builder.Append($_char)
  }

  if (-not $Multi) {
    return $_builder.ToString()
  }

  if ($Text.Length -eq 0) {
    return , [string[]]@()
  }
  $_parts.Add($_builder.ToString())
  if ($_parts.Contains('')) {
    throw 'MULTISZ values must not contain empty strings.'
  }
  return , [string[]]$_parts.ToArray()
}

function ConvertFrom-LgpoText {
  param (
    [string]
    $Path
  )

  $_name = Split-Path -Path $Path -Leaf
  $_bytes = [System.IO.File]::ReadAllBytes($Path)
  $_hasByteOrderMark = $true

  if ($_bytes.Length -ge 3 -and $_bytes[0] -eq 0xEF -and $_bytes[1] -eq 0xBB -and $_bytes[2] -eq 0xBF) {
    $_text = [System.Text.Encoding]::UTF8.GetString($_bytes, 3, $_bytes.Length - 3)
  }
  elseif ($_bytes.Length -ge 2 -and $_bytes[0] -eq 0xFF -and $_bytes[1] -eq 0xFE) {
    $_text = [System.Text.Encoding]::Unicode.GetString($_bytes, 2, $_bytes.Length - 2)
  }
  elseif ($_bytes.Length -ge 2 -and $_bytes[0] -eq 0xFE -and $_bytes[1] -eq 0xFF) {
    throw "${_name}: UTF-16 big-endian is not supported by LGPO. Save the file as UTF-8 with BOM."
  }
  else {
    # Byte-for-byte decode; non-ASCII data lines are rejected below because
    # LGPO reads BOM-less files as ANSI and would store different text.
    $_text = [System.Text.Encoding]::GetEncoding(28591).GetString($_bytes)
    $_hasByteOrderMark = $false
  }

  $_lines = $_text -split "`r?`n"
  $_entries = New-Object System.Collections.Generic.List[object]
  $_index = 0

  while ($_index -lt $_lines.Count) {
    $_line = $_lines[$_index]
    # Only empty lines and column-0 comments are skipped: LGPO rejects
    # whitespace-only lines and indented comments as format errors.
    if ($_line.Length -eq 0 -or $_line.StartsWith(';')) {
      $_index++
      continue
    }

    $_lineNumber = $_index + 1
    if ($_line -ieq 'Computer') { $_scope = 'Machine' }
    elseif ($_line -ieq 'User') { $_scope = 'User' }
    else {
      throw "${_name}:${_lineNumber}: expected 'Computer' or 'User', found '$_line'."
    }

    if (($_index + 3) -ge $_lines.Count) {
      throw "${_name}:${_lineNumber}: incomplete entry - expected key, value name, and action lines."
    }

    $_key = $_lines[$_index + 1]
    $_valueName = $_lines[$_index + 2]
    $_action = $_lines[$_index + 3]

    if (-not $_hasByteOrderMark) {
      foreach ($_dataLine in @($_key, $_valueName, $_action)) {
        if ($_dataLine -match '[^\x00-\x7F]') {
          throw "${_name}:${_lineNumber}: non-ASCII data in a file without a byte-order mark. LGPO reads such files as ANSI; save it as UTF-8 with BOM."
        }
      }
    }

    if ([string]::IsNullOrWhiteSpace($_key) -or $_key -ne $_key.Trim()) {
      throw "${_name}:$($_lineNumber + 1): the registry key must be non-empty without surrounding whitespace."
    }

    $_record = $null
    try {
      switch -CaseSensitive -Regex ($_action) {
        '^DWORD:(.*)$' { $_record = @{ ValueName = $_valueName; Type = 4; Data = (ConvertTo-LgpoInteger -Text $Matches[1] -Bits 32) }; break }
        '^QWORD:(.*)$' { $_record = @{ ValueName = $_valueName; Type = 11; Data = (ConvertTo-LgpoInteger -Text $Matches[1] -Bits 64) }; break }
        '^SZ:(.*)$' { $_record = @{ ValueName = $_valueName; Type = 1; Data = (ConvertFrom-LgpoString -Text $Matches[1]) }; break }
        '^EXSZ:(.*)$' { $_record = @{ ValueName = $_valueName; Type = 2; Data = (ConvertFrom-LgpoString -Text $Matches[1]) }; break }
        '^MULTISZ:(.*)$' { $_record = @{ ValueName = $_valueName; Type = 7; Data = (ConvertFrom-LgpoString -Text $Matches[1] -Multi) }; break }
        '^BINARY:(.*)$' {
          $_hex = $Matches[1]
          $_data = New-Object System.Collections.Generic.List[byte]
          if ($_hex.Length -gt 0) {
            foreach ($_pair in $_hex.Split(',')) {
              if ($_pair -notmatch '^[0-9a-fA-F]{1,2}$') { throw "invalid BINARY byte '$_pair'." }
              $_data.Add([Convert]::ToByte($_pair, 16))
            }
          }
          $_record = @{ ValueName = $_valueName; Type = 3; Data = [byte[]]$_data.ToArray() }
          break
        }
        '^DELETE$' { $_record = @{ ValueName = "**del.$_valueName"; Type = 1; Data = ' ' }; break }
        '^DELETEALLVALUES$' {
          if ($_valueName -ne '*') { throw "DELETEALLVALUES requires '*' as the value name." }
          $_record = @{ ValueName = '**delvals.'; Type = 1; Data = ' ' }
          break
        }
        '^CREATEKEY$' {
          if ($_valueName -ne '*') { throw "CREATEKEY requires '*' as the value name." }
          $_record = @{ ValueName = ''; Type = 0; Data = $null }
          break
        }
        '^DELETEKEYS$' { $_record = @{ ValueName = '**DeleteKeys'; Type = 1; Data = $_valueName }; break }
        '^CLEAR$' { $_record = 'Clear'; break }
        default { throw "unrecognized action '$_action'. Actions are case-sensitive (e.g. DWORD:1, SZ:text, DELETE)." }
      }
    }
    catch {
      throw "${_name}:$($_lineNumber + 3): $($_.Exception.Message)"
    }

    # CLEAR marks a value "not configured"; a freshly built registry.pol simply omits it.
    if ($_record -is [hashtable]) {
      $_entries.Add([PSCustomObject]@{
          Scope     = $_scope
          Key       = $_key
          ValueName = $_record.ValueName
          Type      = [uint32]$_record.Type
          Data      = $_record.Data
          Source    = $_name
          Line      = $_lineNumber
        })
    }

    $_index += 4
  }

  return , $_entries.ToArray()
}

function Expand-GpoTemplate {
  param (
    [string]
    $Template,

    [hashtable]
    $Values,

    [string]
    $Name
  )

  $_unknown = @([regex]::Matches($Template, '\{\{(\w+)\}\}') | ForEach-Object { $_.Groups[1].Value } | Where-Object { -not $Values.ContainsKey($_) } | Select-Object -Unique)
  if ($_unknown.Count -gt 0) {
    throw "Template '$Name' uses placeholder(s) this script does not provide: $($_unknown -join ', ')."
  }

  # Single pass, so a substituted value can never be re-expanded as a placeholder.
  $_evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($m) [string]$Values[$m.Groups[1].Value] }
  return [regex]::Replace($Template, '\{\{(\w+)\}\}', $_evaluator)
}

function Get-PolicyDataFingerprint {
  param ($Entry)

  $_data = $Entry.Data
  if ($_data -is [byte[]]) { $_data = ($_data | ForEach-Object { '{0:x2}' -f $_ }) -join '' }
  elseif ($_data -is [array]) { $_data = $_data -join [char]0 }
  return "$($Entry.Type):$_data"
}

# ---- Resolve inputs ---------------------------------------------------------

$_sourceRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourcePath)
$_outputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$_skeletonRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SkeletonPath)

$_templatePaths = @{
  Backup     = Join-Path $_skeletonRoot 'Backup.xml'
  BackupInfo = Join-Path $_skeletonRoot 'bkupInfo.xml'
}

$_inputProblems = New-Object System.Collections.Generic.List[string]
if (-not (Test-Path -LiteralPath $_sourceRoot -PathType Container)) {
  $_inputProblems.Add("Source directory not found: $_sourceRoot")
}
foreach ($_template in $_templatePaths.Values) {
  if (-not (Test-Path -LiteralPath $_template -PathType Leaf)) {
    $_inputProblems.Add("Skeleton template not found: $_template")
  }
}

if ($_inputProblems.Count -gt 0) {
  foreach ($_problem in $_inputProblems) {
    Write-Log -Message $_problem -Color Red
    Add-OperationResult -Results $_results -Target 'Inputs' -Source 'GroupPolicyBackup' -Action 'Validate' -Status 'Failed' -Detail $_problem
  }
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

Write-Log -Message 'Compiling Group Policy backup' -Color Cyan
Write-Log -Message "  Sources : $_sourceRoot" -Color Gray
Write-Log -Message "  Output  : $_outputRoot" -Color Gray
Write-Log -Message "  Name    : $DisplayName" -Color Gray

# ---- Parse sources ----------------------------------------------------------

$_sourceFiles = New-Object System.Collections.Generic.List[object]
foreach ($_file in @(Get-ChildItem -LiteralPath $_sourceRoot -Filter '*.txt' -File | Sort-Object Name)) {
  $_excludedBy = $Exclude | Where-Object { $_file.Name -like $_ } | Select-Object -First 1
  if ($_excludedBy) {
    Write-Log -Message "  Skipping $($_file.Name) (matches exclude pattern '$_excludedBy')." -Color Gray
    Add-OperationResult -Results $_results -Target $_file.Name -Source 'GroupPolicyBackup' -Action 'Compile' -Status 'Skipped' -Detail "Excluded by pattern '$_excludedBy'."
    continue
  }
  $_sourceFiles.Add($_file)
}

if ($_sourceFiles.Count -eq 0) {
  $_detail = "No policy sources to compile in $_sourceRoot."
  Write-Log -Message $_detail -Color Red
  Add-OperationResult -Results $_results -Target 'Sources' -Source 'GroupPolicyBackup' -Action 'Compile' -Status 'Failed' -Detail $_detail
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

$_allEntries = New-Object System.Collections.Generic.List[object]
foreach ($_file in $_sourceFiles) {
  try {
    $_parsed = ConvertFrom-LgpoText -Path $_file.FullName
    foreach ($_entry in $_parsed) { $_allEntries.Add($_entry) }
    $_machineCount = @($_parsed | Where-Object { $_.Scope -eq 'Machine' }).Count
    $_userCount = @($_parsed | Where-Object { $_.Scope -eq 'User' }).Count
    Write-Log -Message "  $($_file.Name): $_machineCount machine, $_userCount user record(s)" -Color Gray
    Add-OperationResult -Results $_results -Target $_file.Name -Source 'GroupPolicyBackup' -Action 'Compile' -Status 'Completed' -Detail "$_machineCount machine, $_userCount user record(s)." -Property @{
      MachineRecords = $_machineCount
      UserRecords    = $_userCount
    }
  }
  catch {
    Write-Log -Message "  FAILED - $($_.Exception.Message)" -Color Red
    Add-OperationResult -Results $_results -Target $_file.Name -Source 'GroupPolicyBackup' -Action 'Compile' -Status 'Failed' -Detail $_.Exception.Message
  }
}

if (@($_results | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) {
  Write-Log -Message 'Build aborted: fix the source errors above. No backup was written.' -Color Red
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

# Registry.pol is applied in order, so a later file silently wins. Surface that.
$_seen = @{}
foreach ($_entry in $_allEntries) {
  if ($_entry.ValueName.StartsWith('**') -or $_entry.ValueName -eq '') { continue }
  $_identity = "$($_entry.Scope)|$($_entry.Key.ToLowerInvariant())|$($_entry.ValueName.ToLowerInvariant())"
  $_fingerprint = Get-PolicyDataFingerprint -Entry $_entry
  if ($_seen.ContainsKey($_identity)) {
    $_first = $_seen[$_identity]
    if ($_first.Source -ne $_entry.Source -and $_first.Fingerprint -ne $_fingerprint) {
      $_detail = "$($_entry.Scope) $($_entry.Key)\$($_entry.ValueName) is set differently in $($_first.Source):$($_first.Line) and $($_entry.Source):$($_entry.Line); the later file wins."
      Write-Log -Message "  WARNING - $_detail" -Color Yellow
      Add-OperationResult -Results $_results -Target "$($_entry.Key)\$($_entry.ValueName)" -Source 'GroupPolicyBackup' -Action 'Compile' -Status 'Warn' -Detail $_detail
    }
  }
  $_seen[$_identity] = @{ Source = $_entry.Source; Line = $_entry.Line; Fingerprint = $_fingerprint }
}

$_machineEntries = @($_allEntries | Where-Object { $_.Scope -eq 'Machine' })
$_userEntries = @($_allEntries | Where-Object { $_.Scope -eq 'User' })

if (($_machineEntries.Count + $_userEntries.Count) -eq 0) {
  $_detail = 'The sources contain no policy records (only comments or CLEAR actions).'
  Write-Log -Message $_detail -Color Red
  Add-OperationResult -Results $_results -Target 'Sources' -Source 'GroupPolicyBackup' -Action 'Compile' -Status 'Failed' -Detail $_detail
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

# ---- Assemble in staging ----------------------------------------------------

$_backupId = '{' + [guid]::NewGuid().ToString().ToUpperInvariant() + '}'
$_gpoGuid = '{' + [guid]::NewGuid().ToString().ToUpperInvariant() + '}'
$_stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) "winkit-gpo-build-$([guid]::NewGuid().ToString('N'))"
$_stagedBackup = Join-Path $_stagingRoot $_backupId
$_utf8 = New-Object System.Text.UTF8Encoding($false)
$_buildFailed = $false

try {
  $_sides = @(
    @{ Name = 'Machine'; Entries = $_machineEntries; PathVariable = '%GPO_MACH_FSPATH%'; ToolExtension = $_machineToolExtension },
    @{ Name = 'User'; Entries = $_userEntries; PathVariable = '%GPO_USER_FSPATH%'; ToolExtension = $_userToolExtension }
  )

  $_values = @{
    GPOGuid        = $_gpoGuid
    BackupId       = $_backupId
    GPODomainGuid  = '{' + [guid]::NewGuid().ToString() + '}'
    BackupTime     = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    GPODisplayName = $DisplayName
    Comment        = "Built by winkit Build-GroupPolicyBackup from: $(($_sourceFiles | ForEach-Object { $_.Name }) -join ', ')"
    DomainSidBytes = (($_provenance.DomainSid -split '-' | Select-Object -Last 3 | ForEach-Object { [BitConverter]::GetBytes([uint32]$_) | ForEach-Object { '{0:x2}' -f $_ } }) -join ' ')
  }
  foreach ($_key in $_provenance.Keys) { $_values[$_key] = $_provenance[$_key] }

  $_fileEntries = New-Object System.Collections.Generic.List[string]
  foreach ($_side in $_sides) {
    $_sideDirectory = Join-Path $_stagedBackup "DomainSysvol\GPO\$($_side.Name)"
    # Staging lives in TEMP and is always written, including in a dry run.
    $null = New-Item -ItemType Directory -Path $_sideDirectory -Force -WhatIf:$false

    if ($_side.Entries.Count -gt 0) {
      $_polPath = Join-Path $_sideDirectory 'registry.pol'
      $_records = @($_side.Entries | ForEach-Object { [PSCustomObject]@{ Key = $_.Key; ValueName = $_.ValueName; Type = $_.Type; Data = $_.Data } })
      ConvertTo-RegistryPolicy -InputObject $_records -Path $_polPath -WhatIf:$false -Confirm:$false

      $_values["$($_side.Name)VersionNumber"] = '65537'
      $_values["$($_side.Name)ExtensionGuids"] = "<![CDATA[[$_registryCse$($_side.ToolExtension)]]]>"
      $_fileEntries.Add("      <FSObjectFile bkp:Path=`"$($_side.PathVariable)\registry.pol`" bkp:SourceExpandedPath=`"\\$($_provenance.GPODomainController)\sysvol\$($_provenance.GPODomain)\Policies\$_gpoGuid\$($_side.Name)\registry.pol`" bkp:Location=`"DomainSysvol\GPO\$($_side.Name)\registry.pol`" />")
    }
    else {
      $_values["$($_side.Name)VersionNumber"] = '0'
      $_values["$($_side.Name)ExtensionGuids"] = ''
    }
  }
  $_values['RegistryFileEntries'] = $_fileEntries -join "`r`n"

  $_backupXml = Expand-GpoTemplate -Template ([System.IO.File]::ReadAllText($_templatePaths.Backup)) -Values $_values -Name 'Backup.xml'
  $_backupInfoXml = Expand-GpoTemplate -Template ([System.IO.File]::ReadAllText($_templatePaths.BackupInfo)) -Values $_values -Name 'bkupInfo.xml'
  [System.IO.File]::WriteAllText((Join-Path $_stagedBackup 'Backup.xml'), $_backupXml, $_utf8)
  [System.IO.File]::WriteAllText((Join-Path $_stagedBackup 'bkupInfo.xml'), $_backupInfoXml, $_utf8)

  # ---- Validate the staged backup ------------------------------------------
  foreach ($_document in @('Backup.xml', 'bkupInfo.xml')) {
    $_content = [System.IO.File]::ReadAllText((Join-Path $_stagedBackup $_document))
    if ($_content -match '\{\{\w+\}\}') {
      throw "$_document still contains an unsubstituted placeholder."
    }
    try { $null = [xml]$_content }
    catch { throw "$_document is not well-formed XML after stamping: $($_.Exception.Message)" }
  }

  [xml]$_stampedBackup = $_backupXml
  $_descriptorHex = $_stampedBackup.GroupPolicyBackupScheme.GroupPolicyObject.GroupPolicyCoreSettings.SecurityDescriptor.Trim()
  $_descriptorBytes = [byte[]]($_descriptorHex -split '\s+' | ForEach-Object { [Convert]::ToByte($_, 16) })
  $null = New-Object System.Security.AccessControl.RawSecurityDescriptor -ArgumentList $_descriptorBytes, 0

  foreach ($_side in $_sides) {
    if ($_side.Entries.Count -eq 0) { continue }
    $_readBack = @(ConvertFrom-RegistryPolicy -Path (Join-Path $_stagedBackup "DomainSysvol\GPO\$($_side.Name)\registry.pol"))
    if ($_readBack.Count -ne $_side.Entries.Count) {
      throw "$($_side.Name) registry.pol holds $($_readBack.Count) record(s), expected $($_side.Entries.Count)."
    }
  }

  Write-Log -Message "  Staged and validated backup $_backupId ($($_machineEntries.Count) machine, $($_userEntries.Count) user record(s))." -Color Gray

  # ---- Publish ---------------------------------------------------------------
  $_publishedPath = Join-Path $_outputRoot $_backupId
  $_backupProperty = @{
    BackupId       = $_backupId
    GpoGuid        = $_gpoGuid
    DisplayName    = $DisplayName
    Path           = $_publishedPath
    MachineRecords = $_machineEntries.Count
    UserRecords    = $_userEntries.Count
  }

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would write $_publishedPath and add it to manifest.xml." -Color Yellow
    Add-OperationResult -Results $_results -Target $_backupId -Source 'GroupPolicyBackup' -Action 'Build' -Status 'Skipped' -Detail 'DryRun: compiled and validated in staging; nothing written.' -Property $_backupProperty
  }
  elseif ($PSCmdlet.ShouldProcess($_publishedPath, "Write GPO backup '$DisplayName'")) {
    $null = New-Item -ItemType Directory -Path $_outputRoot -Force
    Move-Item -LiteralPath $_stagedBackup -Destination $_publishedPath -ErrorAction Stop

    $_manifestPath = Join-Path $_outputRoot 'manifest.xml'
    $_manifestNamespace = 'http://www.microsoft.com/GroupPolicy/GPOOperations/Manifest'
    $_manifest = New-Object System.Xml.XmlDocument
    if (Test-Path -LiteralPath $_manifestPath -PathType Leaf) {
      $_manifest.Load($_manifestPath)
    }
    else {
      $_manifest.LoadXml("<Backups xmlns=`"$_manifestNamespace`" xmlns:mfst=`"$_manifestNamespace`" mfst:version=`"1.0`" />")
    }
    [xml]$_instance = $_backupInfoXml
    [void]$_manifest.DocumentElement.AppendChild($_manifest.ImportNode($_instance.DocumentElement, $true))

    $_settings = New-Object System.Xml.XmlWriterSettings
    $_settings.Encoding = $_utf8
    $_settings.OmitXmlDeclaration = $true
    $_writer = [System.Xml.XmlWriter]::Create($_manifestPath, $_settings)
    try { $_manifest.Save($_writer) }
    finally { $_writer.Dispose() }

    Write-Log -Message "Backup written: $_publishedPath" -Color Green
    Add-OperationResult -Results $_results -Target $_backupId -Source 'GroupPolicyBackup' -Action 'Build' -Status 'Completed' -Detail $_publishedPath -Property $_backupProperty
  }
  else {
    Add-OperationResult -Results $_results -Target $_backupId -Source 'GroupPolicyBackup' -Action 'Build' -Status 'Skipped' -Detail 'WhatIf' -Property $_backupProperty
  }
}
catch {
  $_buildFailed = $true
  Write-Log -Message "FAILED - $($_.Exception.Message)" -Color Red
  Add-OperationResult -Results $_results -Target $_backupId -Source 'GroupPolicyBackup' -Action 'Build' -Status 'Failed' -Detail $_.Exception.Message
}
finally {
  if (Test-Path -LiteralPath $_stagingRoot) {
    Remove-Item -LiteralPath $_stagingRoot -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
  }
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Build-GroupPolicyBackup'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_buildFailed) {
  exit 1
}
