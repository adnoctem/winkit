#Requires -Version 5.0

<#
.SYNOPSIS
  Runs PSScriptAnalyzer checks for repository PowerShell sources.

.DESCRIPTION
  Invokes PSScriptAnalyzer once per configured path because PSScriptAnalyzer
  1.25.0 expects a single string for -Path and rejects an array such as
  `-Path ./lib,./scripts`. Results from each invocation are collected and
  printed together.

  The script exits with code 1 when any analyzer findings remain, making it
  suitable for pre-commit and CI usage.

.PARAMETER Path
  Files or directories to analyze. Defaults to lib, scripts, tools/build.ps1,
  and tools/format.ps1 when they exist.

.PARAMETER Settings
  PSScriptAnalyzer settings file. Defaults to PSScriptAnalyzerSettings.psd1 in
  the repository root.

.EXAMPLE
  PS> ./lint.ps1
  Runs analyzer checks over the default repository PowerShell paths.

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
  [string[]]$Path = @(
    (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'lib'),
    (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'scripts'),
    (Join-Path -Path $PSScriptRoot -ChildPath 'build.ps1'),
    (Join-Path -Path $PSScriptRoot -ChildPath 'format.ps1')
  ),

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

$results = New-Object System.Collections.Generic.List[object]

foreach ($entry in $Path) {
  $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($entry)
  if (-not (Test-Path -LiteralPath $resolvedPath)) {
    Write-Warning "Skipping missing analyzer path: $entry"
    continue
  }

  $_analysis = @(Invoke-ScriptAnalyzer -Path $resolvedPath -Recurse -Settings $settingsPath)
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

Write-Output "PSScriptAnalyzer passed for $($Path.Count) path(s)."
exit 0
