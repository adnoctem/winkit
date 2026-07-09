#Requires -Version 5.0

<#
.SYNOPSIS
  Formats repository PowerShell source files with PSScriptAnalyzer.

.DESCRIPTION
  Runs Invoke-Formatter over .ps1, .psm1, and .psd1 files using the repository
  PSScriptAnalyzerSettings.psd1 file. The script delegates whitespace, brace,
  indentation, and casing rules entirely to PSScriptAnalyzer and performs no
  repository-specific post-processing beyond encoding and line-ending
  normalization on write.

  Files are rewritten in place with UTF-8 with BOM (required for reliable
  parsing under Windows PowerShell 5.1) and CRLF line endings.

  Use -Check to report files that would change without writing them, suitable
  for pre-commit hooks and CI jobs.

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

.EXAMPLE
  PS> ./format.ps1
  Formats all repository PowerShell files in place.

.EXAMPLE
  PS> ./format.ps1 -Check
  Verifies formatting and exits non-zero if any file would change.

.EXAMPLE
  PS> ./format.ps1 -Path ./lib,./scripts
  Formats only the library and script directories.

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

  [switch]$IncludeSecrets
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
$excludedDirectories = @('.git', '.idea', 'dist', 'build')
if (-not $IncludeSecrets) {
  $excludedDirectories += 'secrets'
}

$rootFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Split-Path -Path $PSScriptRoot -Parent))
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$changed = New-Object System.Collections.Generic.List[string]
$processed = 0

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

  $_formattedTrimmed = $formatted -replace '\s+$', ''
  $_normalizedTrimmed = $normalizedSource -replace '\s+$', ''

  $_needsFormat = ($_formattedTrimmed -ne $_normalizedTrimmed)
  $_needsLineEndingFix = ($source -match '(?<!\r)\n|\r(?!\n)')

  if ($_needsFormat -or $_needsLineEndingFix) {
    [void]$changed.Add($file.FullName)
    if (-not $Check) {
      $formattedCrlf = $formatted -replace "`r`n", "`n" -replace "`n", "`r`n"
      [System.IO.File]::WriteAllText($file.FullName, $formattedCrlf, $utf8Bom)
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
