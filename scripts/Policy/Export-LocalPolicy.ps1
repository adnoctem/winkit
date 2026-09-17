#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.4.0' }

<#
.SYNOPSIS
  Exports the local Group Policy registry settings as an LGPO text policy source.
.DESCRIPTION
  Reads the local Machine and User registry.pol files with PSFoundation's
  ConvertFrom-RegistryPolicy and writes them as a single LGPO text file - the
  source format compiled by Build-GroupPolicyBackup.ps1. This turns a
  hand-configured golden image into a reviewable, version-controlled policy
  source.

  Everything present in local policy is captured, including settings that did
  not come from the golden-image configuration. Entries are grouped under their
  top-most captured registry key and every group is reported, so nothing enters
  a baseline unnoticed. Review and prune the file before committing it.

  The metadata header (Title, Owner, Justification, ADMX, Upstream) cannot be
  derived from registry.pol and is emitted as a TODO stub to complete by hand.

  Some registry.pol records have no LGPO text form: the **DeleteValues,
  **SecureKey and **soft. directives, registry types other than SZ, EXSZ,
  MULTISZ, BINARY, DWORD and QWORD, and strings containing line breaks. These
  are written as commented-out blocks and reported as warnings, because a baseline
  compiled from the file will not contain them.

  The file is written as UTF-8 with a byte-order mark: LGPO reads BOM-less files
  as ANSI, which corrupts non-ASCII values. LGPO.exe itself is not required.
.PARAMETER Path
  Destination text file. Defaults to a timestamped file under dist/GP/exports,
  outside the tracked policy sources.
.PARAMETER Scope
  Which local policy to export: All (default), Machine, or User.
.PARAMETER PolicyRoot
  Local Group Policy directory holding Machine\registry.pol and
  User\registry.pol. Defaults to %SystemRoot%\System32\GroupPolicy.
.PARAMETER Force
  Overwrite an existing destination file.
.PARAMETER DryRun
  Report what would be exported without writing the file.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> .\Export-LocalPolicy.ps1
  Exports machine and user local policy to dist/GP/exports.
.EXAMPLE
  PS> .\Export-LocalPolicy.ps1 -Scope Machine -Path .\resources\policies\10-golden-image.txt
  Exports machine policy directly as a new policy source.
.EXAMPLE
  PS> .\Export-LocalPolicy.ps1 -DryRun
  Lists the captured registry roots without writing anything.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - reads local policy files directly, no Group Policy editor required.
  SYSTEM-account execution: supported; reading local policy does not require elevation.
  Multiple local GPOs (Administrators, Non-Administrators, per-user) are not exported.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [string]
  $Path = (Join-Path $PSScriptRoot "..\..\dist\GP\exports\local-policy-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"),

  [ValidateSet('All', 'Machine', 'User')]
  [string]
  $Scope = 'All',

  [string]
  $PolicyRoot = (Join-Path $env:SystemRoot 'System32\GroupPolicy'),

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
  Write-Log -Message "DRY RUN - the export file will not be written`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

function ConvertTo-LgpoString {
  param ([string]$Text)
  return $Text.Replace('\', '\\')
}

function ConvertTo-LgpoTextBlock {
  # Returns the four LGPO text lines for a registry.pol record, or the reason
  # the record has no LGPO text form.
  param ($Entry, [string]$ScopeLine)

  $_key = $Entry.Key
  $_name = $Entry.ValueName
  $_data = $Entry.Data

  if ($_key -match '[\r\n]' -or $_name -match '[\r\n]') {
    return @{ Reason = 'the key or value name contains a line break' }
  }
  if ($_key -ne $_key.Trim()) {
    return @{ Reason = 'the key has leading or trailing whitespace' }
  }

  if ($_name.StartsWith('**del.')) {
    return @{ Lines = @($ScopeLine, $_key, $_name.Substring(6), 'DELETE') }
  }
  if ($_name -eq '**delvals.') {
    return @{ Lines = @($ScopeLine, $_key, '*', 'DELETEALLVALUES') }
  }
  if ($_name -eq '**DeleteKeys') {
    if ([string]$_data -match '[\r\n]') { return @{ Reason = 'the **DeleteKeys list contains a line break' } }
    return @{ Lines = @($ScopeLine, $_key, [string]$_data, 'DELETEKEYS') }
  }
  if ($_name.StartsWith('**')) {
    return @{ Reason = "directive '$_name' has no LGPO text form" }
  }
  if ($_name -eq '' -and $Entry.Type -eq 0 -and ($null -eq $_data -or @($_data).Count -eq 0)) {
    return @{ Lines = @($ScopeLine, $_key, '*', 'CREATEKEY') }
  }

  switch ([int]$Entry.Type) {
    { $_ -in 1, 2 } {
      if ($_data -match '[\r\n]') { return @{ Reason = 'the string value contains a line break' } }
      if ([string]$_data -match '\x00') { return @{ Reason = 'the string value contains an embedded NUL character' } }
      $_prefix = if ([int]$Entry.Type -eq 1) { 'SZ' } else { 'EXSZ' }
      return @{ Lines = @($ScopeLine, $_key, $_name, "${_prefix}:$(ConvertTo-LgpoString ([string]$_data))") }
    }
    7 {
      $_items = @($_data | Where-Object { $null -ne $_ })
      if (@($_items | Where-Object { $_ -match '[\r\n]' }).Count -gt 0) { return @{ Reason = 'a multi-string item contains a line break' } }
      $_joined = (@($_items | ForEach-Object { ConvertTo-LgpoString ([string]$_) })) -join '\0'
      return @{ Lines = @($ScopeLine, $_key, $_name, "MULTISZ:$_joined") }
    }
    3 {
      $_hex = (@($_data) | ForEach-Object { '{0:x2}' -f $_ }) -join ','
      return @{ Lines = @($ScopeLine, $_key, $_name, "BINARY:$_hex") }
    }
    4 { return @{ Lines = @($ScopeLine, $_key, $_name, "DWORD:$([uint32]$_data)") } }
    11 { return @{ Lines = @($ScopeLine, $_key, $_name, "QWORD:$([uint64]$_data)") } }
    default { return @{ Reason = "registry type $($Entry.Type) has no LGPO text form" } }
  }
}

# ---- Read local policy -------------------------------------------------------

$_policyRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PolicyRoot)
$_outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

$_scopes = @(
  @{ Name = 'Machine'; ScopeLine = 'Computer'; File = Join-Path $_policyRoot 'Machine\registry.pol' },
  @{ Name = 'User'; ScopeLine = 'User'; File = Join-Path $_policyRoot 'User\registry.pol' }
) | Where-Object { $Scope -eq 'All' -or $_.Name -eq $Scope }

Write-Log -Message 'Exporting local Group Policy registry settings' -Color Cyan
Write-Log -Message "  Policy root : $_policyRoot" -Color Gray

$_groups = New-Object System.Collections.Generic.List[object]
$_exported = 0
$_notExportable = 0

foreach ($_scope in $_scopes) {
  if (-not (Test-Path -LiteralPath $_scope.File -PathType Leaf)) {
    Write-Log -Message "  $($_scope.Name): no registry.pol - nothing configured." -Color Gray
    continue
  }

  try {
    $_entries = @(ConvertFrom-RegistryPolicy -Path $_scope.File)
  }
  catch {
    Write-Log -Message "  FAILED - $($_scope.Name): $($_.Exception.Message)" -Color Red
    Add-OperationResult -Results $_results -Target $_scope.File -Source 'LocalPolicyExport' -Action 'Read' -Status 'Failed' -Detail $_.Exception.Message
    continue
  }

  # Group each entry under its top-most captured ancestor key. Relative order
  # within a key is preserved, which is what registry.pol processing depends on.
  $_keys = @($_entries | ForEach-Object { $_.Key } | Select-Object -Unique)
  $_rootOf = @{}
  foreach ($_key in $_keys) {
    $_root = $_key
    foreach ($_candidate in $_keys) {
      if ($_candidate.Length -lt $_root.Length -and $_key.StartsWith("$_candidate\", [System.StringComparison]::OrdinalIgnoreCase)) {
        $_root = $_candidate
      }
    }
    $_rootOf[$_key.ToLowerInvariant()] = $_root
  }

  $_scopeGroups = [ordered]@{}
  foreach ($_entry in $_entries) {
    $_root = $_rootOf[$_entry.Key.ToLowerInvariant()]
    $_groupKey = $_root.ToLowerInvariant()
    if (-not $_scopeGroups.Contains($_groupKey)) {
      $_scopeGroups[$_groupKey] = [PSCustomObject]@{
        Scope         = $_scope.Name
        ScopeLine     = $_scope.ScopeLine
        Root          = $_root
        Blocks        = New-Object System.Collections.Generic.List[object]
        Exportable    = 0
        NotExportable = 0
      }
    }
    $_group = $_scopeGroups[$_groupKey]
    $_block = ConvertTo-LgpoTextBlock -Entry $_entry -ScopeLine $_scope.ScopeLine
    $_group.Blocks.Add([PSCustomObject]@{ Entry = $_entry; Lines = $_block.Lines; Reason = $_block.Reason })
    if ($_block.Lines) { $_group.Exportable++ } else { $_group.NotExportable++ }
  }

  foreach ($_group in $_scopeGroups.Values) {
    $_groups.Add($_group)
    $_exported += $_group.Exportable
    $_notExportable += $_group.NotExportable
  }
}

# ---- Report ------------------------------------------------------------------

if ($_groups.Count -gt 0) {
  Write-Log -Message "`nCaptured registry roots:" -Color Cyan
  foreach ($_group in $_groups) {
    $_suffix = if ($_group.NotExportable -gt 0) { ", $($_group.NotExportable) not exportable" } else { '' }
    Write-Log -Message ("  {0,-7} {1} ({2} entries{3})" -f $_group.Scope, $_group.Root, $_group.Exportable, $_suffix) -Color $(if ($_group.NotExportable -gt 0) { 'Yellow' } else { 'Gray' })
    Add-OperationResult -Results $_results -Target $_group.Root -Source 'LocalPolicyExport' -Action 'Capture' -Status 'Completed' -Detail "$($_group.Scope): $($_group.Exportable) entries." -Property @{
      Scope         = $_group.Scope
      Entries       = $_group.Exportable
      NotExportable = $_group.NotExportable
    }
    foreach ($_block in @($_group.Blocks | Where-Object { -not $_.Lines })) {
      $_detail = "$($_group.Scope) $($_block.Entry.Key) '$($_block.Entry.ValueName)' not exported: $($_block.Reason)."
      Write-Log -Message "    WARNING - $_detail" -Color Yellow
      Add-OperationResult -Results $_results -Target "$($_block.Entry.Key)\$($_block.Entry.ValueName)" -Source 'LocalPolicyExport' -Action 'Capture' -Status 'Warn' -Detail $_detail
    }
  }
}

if ($_exported -eq 0) {
  $_detail = 'No exportable local policy entries were found; no file written.'
  Write-Log -Message "`n$_detail" -Color Yellow
  Add-OperationResult -Results $_results -Target $_outputPath -Source 'LocalPolicyExport' -Action 'Write' -Status 'Skipped' -Detail $_detail
}
else {
  # ---- Compose -----------------------------------------------------------------
  $_scopeLabel = (@($_groups | ForEach-Object { $_.Scope } | Select-Object -Unique)) -join ', '
  $_text = New-Object System.Collections.Generic.List[string]
  $_text.Add("; Version: $(Get-Date -Format 'yyyy-MM-dd')-01")
  $_text.Add('; Title: TODO')
  $_text.Add('; Owner: TODO')
  $_text.Add('; Justification: TODO')
  $_text.Add('; ADMX: TODO')
  $_text.Add('; Upstream: TODO')
  $_text.Add(';')
  $_text.Add("; Exported by Export-LocalPolicy on $([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)) (scope: $_scopeLabel).")
  $_text.Add('; This captures ALL local policy, including settings that did not come from')
  $_text.Add('; the golden-image configuration. Review and prune every group before committing.')

  foreach ($_group in $_groups) {
    $_text.Add('')
    $_text.Add("; ---- $($_group.ScopeLine): $($_group.Root) ----")
    foreach ($_block in $_group.Blocks) {
      $_text.Add('')
      if ($_block.Lines) {
        foreach ($_line in $_block.Lines) { $_text.Add($_line) }
      }
      else {
        $_text.Add("; NOT EXPORTED - $($_block.Reason). A baseline built from this file will not contain it.")
        $_text.Add("; $($_group.ScopeLine)")
        $_text.Add("; $($_block.Entry.Key -replace '[\r\n]', ' ')")
        $_text.Add("; $($_block.Entry.ValueName -replace '[\r\n]', ' ')")
      }
    }
  }

  $_summary = "$_exported entries in $($_groups.Count) group(s)" + $(if ($_notExportable -gt 0) { ", $_notExportable not exportable" } else { '' })

  if ($DryRun) {
    Write-Log -Message "`n[DRY RUN] Would write $_summary to $_outputPath" -Color Yellow
    Add-OperationResult -Results $_results -Target $_outputPath -Source 'LocalPolicyExport' -Action 'Write' -Status 'Skipped' -Detail "DryRun: $_summary."
  }
  elseif ((Test-Path -LiteralPath $_outputPath) -and -not $Force) {
    $_detail = "Destination exists: $_outputPath. Use -Force to overwrite it."
    Write-Log -Message "`nFAILED - $_detail" -Color Red
    Add-OperationResult -Results $_results -Target $_outputPath -Source 'LocalPolicyExport' -Action 'Write' -Status 'Failed' -Detail $_detail
  }
  elseif ($PSCmdlet.ShouldProcess($_outputPath, "Write local policy export ($_summary)")) {
    try {
      $_directory = Split-Path -Path $_outputPath -Parent
      if ($_directory) { $null = New-Item -ItemType Directory -Path $_directory -Force }
      $_encoding = New-Object System.Text.UTF8Encoding($true)
      [System.IO.File]::WriteAllText($_outputPath, (($_text -join "`r`n") + "`r`n"), $_encoding)
      Write-Log -Message "`nExported $_summary to $_outputPath" -Color Green
      Write-Log -Message 'Complete the TODO header and review every group before committing.' -Color Yellow
      Add-OperationResult -Results $_results -Target $_outputPath -Source 'LocalPolicyExport' -Action 'Write' -Status 'Completed' -Detail $_summary -Property @{
        Entries       = $_exported
        NotExportable = $_notExportable
        Groups        = $_groups.Count
      }
    }
    catch {
      Write-Log -Message "`nFAILED - $($_.Exception.Message)" -Color Red
      Add-OperationResult -Results $_results -Target $_outputPath -Source 'LocalPolicyExport' -Action 'Write' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  else {
    Add-OperationResult -Results $_results -Target $_outputPath -Source 'LocalPolicyExport' -Action 'Write' -Status 'Skipped' -Detail 'WhatIf'
  }
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Export-LocalPolicy'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if (@($_results | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) {
  exit 1
}
