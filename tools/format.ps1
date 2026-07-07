#Requires -Version 5.0

<#
.SYNOPSIS
  Formats repository PowerShell source files with PSScriptAnalyzer.

.DESCRIPTION
  Runs Invoke-Formatter over .ps1, .psm1, and .psd1 files using the repository
  PSScriptAnalyzerSettings.psd1 file. By default, files are rewritten in place
  with UTF-8 without BOM and CRLF line endings.

  The script intentionally delegates normal whitespace, brace, indentation, and
  casing rules to PSScriptAnalyzer. Its only repository-specific post-processing
  is expanding single-line hashtable literals with multiple key/value pairs into
  a readable block form. That keeps configure-script settings arrays close to a
  JSON-like layout while avoiding column alignment such as:

    Rules               = @{
    VeryLongPropertyKey = $true

  The desired style is single-space assignment:

    Rules = @{
      VeryLongPropertyKey = $true
    }

  Use -NoExpandHashtables to run only Invoke-Formatter without that expansion
  pass. Use -Check to report files that would change without writing them,
  suitable for pre-commit hooks and CI jobs.

.PARAMETER Path
  Root paths to scan. Defaults to the repository root.

.PARAMETER Settings
  PSScriptAnalyzer settings file. Defaults to PSScriptAnalyzerSettings.psd1 in
  the repository root.

.PARAMETER Check
  Report formatting drift without modifying files. Exits with code 1 when any
  file would be changed.

.PARAMETER IncludeSecrets
  Include files under the secrets directory. Excluded by default.

.PARAMETER NoExpandHashtables
  Disable the repository-specific post-pass that expands single-line hashtable
  literals with multiple key/value pairs into multi-line blocks.

.EXAMPLE
  PS> ./format.ps1
  Formats all repository PowerShell files in place.

.EXAMPLE
  PS> ./format.ps1 -Check
  Verifies formatting and exits non-zero if any file would change.

.EXAMPLE
  PS> ./format.ps1 -Path ./lib,./scripts
  Formats only the library and script directories.

.EXAMPLE
  PS> ./format.ps1 -Check -NoExpandHashtables
  Checks the raw PSScriptAnalyzer formatter output without expanding single-line
  hashtables.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding()]
param (
  [Parameter(Position = 0)]
  [string[]]$Path = @(Split-Path -Path $PSScriptRoot -Parent),

  [string]$Settings = (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'PSScriptAnalyzerSettings.psd1'),

  [switch]$Check,

  [switch]$IncludeSecrets,

  [switch]$NoExpandHashtables
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
  Write-Error 'PSScriptAnalyzer is not installed. Install it with: Install-Module PSScriptAnalyzer'
  exit 1
}

Import-Module PSScriptAnalyzer -ErrorAction Stop

$settingsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Settings)
if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) {
  Write-Error "Settings file not found: $settingsPath"
  exit 1
}

$extensions = @('.ps1', '.psm1', '.psd1')
$excludedDirectories = @('.git', '.idea')
if (-not $IncludeSecrets) {
  $excludedDirectories += 'secrets'
}

$rootFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Split-Path -Path $PSScriptRoot -Parent))
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$changed = New-Object System.Collections.Generic.List[string]
$processed = 0

function Expand-SingleLineHashtable {
  <#
    .SYNOPSIS
      Expands compact multi-entry hashtable literals into block form.

    .DESCRIPTION
      Invoke-Formatter normalizes whitespace inside hashtables, but it does not
      currently expand compact literals such as
      `@{ Path = 'HKCU:\x'; Name = 'Foo'; Type = 'DWord' }`. This helper uses
      the PowerShell parser to find only syntactically valid single-line
      hashtable literals with more than one key/value pair and rewrites them as
      one property per line.

      The helper does not align assignment operators. PSScriptAnalyzer's
      PSAlignAssignmentStatement rule can align hashtable values, but for this
      repository that makes large settings objects harder to read when one key
      is much longer than the others.
  #>

  param (
    [Parameter(Mandatory = $true)]
    [string]$Text
  )

  $tokens = $null
  $errors = $null
  $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors)
  if ($errors.Count -gt 0) {
    return $Text
  }

  $hashtables = @($ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.HashtableAst] -and
        $node.KeyValuePairs.Count -gt 1 -and
        $node.Extent.Text.IndexOf("`n") -lt 0
      }, $true) | Sort-Object { $_.Extent.StartOffset } -Descending)

  $result = $Text
  foreach ($hashtable in $hashtables) {
    $start = $hashtable.Extent.StartOffset
    $lineStart = $result.LastIndexOf("`n", [Math]::Max(0, $start - 1))
    if ($lineStart -lt 0) {
      $lineStart = 0
    }
    else {
      $lineStart++
    }

    $indent = $result.Substring($lineStart, $start - $lineStart)
    if ($indent -match '\S') {
      $indent = ''
    }

    $childIndent = "$indent  "
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('@{')

    foreach ($pair in $hashtable.KeyValuePairs) {
      $key = $pair.Item1.Extent.Text
      $value = $pair.Item2.Extent.Text
      [void]$lines.Add("$childIndent$key = $value")
    }

    [void]$lines.Add("$indent}")
    $replacement = $lines -join "`n"
    $result = $result.Remove($hashtable.Extent.StartOffset, $hashtable.Extent.EndOffset - $hashtable.Extent.StartOffset).Insert($hashtable.Extent.StartOffset, $replacement)
  }

  return $result
}

function Test-FormatterExcludedPath {
  param (
    [Parameter(Mandatory = $true)]
    [string]$FilePath
  )

  $relative = $FilePath
  if ($FilePath.StartsWith($rootFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    $relative = $FilePath.Substring($rootFullPath.Length).TrimStart('\', '/')
  }

  foreach ($directory in $excludedDirectories) {
    if ($relative -eq $directory -or $relative.StartsWith("$directory\", [System.StringComparison]::OrdinalIgnoreCase) -or $relative.StartsWith("$directory/", [System.StringComparison]::OrdinalIgnoreCase)) {
      return $true
    }
  }

  return $false
}

$files = foreach ($entry in $Path) {
  $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($entry)
  if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) {
    Get-Item -LiteralPath $resolvedPath
  }
  elseif (Test-Path -LiteralPath $resolvedPath -PathType Container) {
    Get-ChildItem -LiteralPath $resolvedPath -Recurse -File
  }
  else {
    Write-Error "Path not found: $entry"
  }
}

$files = @($files |
    Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
    Where-Object { -not (Test-FormatterExcludedPath -FilePath $_.FullName) } |
    Sort-Object -Property FullName -Unique)

foreach ($file in $files) {
  $processed++
  $source = [System.IO.File]::ReadAllText($file.FullName)
  $normalizedSource = $source -replace "`r`n|`r|`n", "`n"
  $formatted = Invoke-Formatter -ScriptDefinition $normalizedSource -Settings $settingsPath
  if (-not $NoExpandHashtables) {
    $formatted = Expand-SingleLineHashtable -Text $formatted
  }

  if ($formatted -ne $normalizedSource -or $source -ne $normalizedSource) {
    [void]$changed.Add($file.FullName)
    if (-not $Check) {
      $formattedCrlf = $formatted -replace "`r`n", "`n" -replace "`n", "`r`n"
      [System.IO.File]::WriteAllText($file.FullName, $formattedCrlf, $utf8NoBom)
      Write-Output "Formatted: $($file.FullName)"
    }
  }
}

if ($Check) {
  if ($changed.Count -gt 0) {
    Write-Output "Formatting required for $($changed.Count) file(s):"
    $changed | ForEach-Object { Write-Output "  $_" }
    exit 1
  }

  Write-Output "Formatting check passed for $processed file(s)."
  exit 0
}

Write-Output "Formatting complete. Processed: $processed | Changed: $($changed.Count)"
