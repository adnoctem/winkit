#Requires -Version 5.0

<#
.SYNOPSIS
  Runs PSScriptAnalyzer checks for repository PowerShell sources.

.DESCRIPTION
  By default, scans the entire repository root for .ps1, .psm1, and .psd1 files
  and invokes PSScriptAnalyzer once per file. The .git, .idea, dist, build, and
  secrets directories are excluded by default so generated, vendored, or
  sensitive content is never analyzed.

  PSScriptAnalyzer 1.25.0 expects a single string for -Path and rejects an
  array such as `-Path ./lib,./scripts`, so results are collected per file and
  printed together.

  The script exits with code 1 when any analyzer findings remain, making it
  suitable for pre-commit and CI usage.

.PARAMETER Path
  Files or directories to analyze. Defaults to the repository root, which is
  scanned recursively for PowerShell sources with the standard exclusions
  applied.

.PARAMETER Settings
  PSScriptAnalyzer settings file. Defaults to PSScriptAnalyzerSettings.psd1 in
  the repository root.

.EXAMPLE
  PS> ./lint.ps1
  Runs analyzer checks over all repository PowerShell sources.

.EXAMPLE
  PS> ./lint.ps1 -Path ./lib,./scripts/Windows
  Runs analyzer checks over selected paths.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding()]
param (
  [string[]]$Path = @(Split-Path -Path $PSScriptRoot -Parent),

  [string]$Settings = (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'PSScriptAnalyzerSettings.psd1')
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
  Write-Error 'PSScriptAnalyzer is not installed. Install it with: Install-Module PSScriptAnalyzer'
  exit 1
}

$settingsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Settings)
if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) {
  Write-Error "Settings file not found: $settingsPath"
  exit 1
}

$extensions = @('.ps1', '.psm1', '.psd1')
$excludedDirectories = @('.git', '.idea', 'dist', 'build', 'secrets')
$rootFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Split-Path -Path $PSScriptRoot -Parent))

function Test-LintExcludedPath {
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
    Write-Warning "Skipping missing analyzer path: $entry"
  }
}

$files = @($files |
    Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
    Where-Object { -not (Test-LintExcludedPath -FilePath $_.FullName) } |
    Sort-Object -Property FullName -Unique)

if ($files.Count -eq 0) {
  Write-Output 'PSScriptAnalyzer passed: no PowerShell files to analyze.'
  exit 0
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($file in $files) {
  $_analysis = @(Invoke-ScriptAnalyzer -Path $file.FullName -Settings $settingsPath)
  foreach ($_result in $_analysis) {
    [void]$results.Add($_result)
  }
}

if ($results.Count -gt 0) {
  $results |
    Select-Object RuleName, Severity, ScriptName, Line, Message |
    Format-Table -AutoSize
  exit 1
}

Write-Output "PSScriptAnalyzer passed for $($files.Count) file(s)."
exit 0
