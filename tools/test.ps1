#Requires -Version 5.1

<#
.SYNOPSIS
  Runs the winkit Pester test suite.
.DESCRIPTION
  Runs all Pester tests under the repository tests directory. By default only
  the logic tests that need no external services run. Pass -Outlook to also
  run the Outlook integration suite, which needs an installed Outlook and
  creates a disposable scratch PST store for the duration of the run.

  Logic tests for the shared PSFoundation module run in the PSFoundation
  repository instead.
.PARAMETER Outlook
  Enable the Outlook integration tests (tests/Office/Outlook.Integration.Tests.ps1).
.EXAMPLE
  PS> ./tools/test.ps1
  Runs the winkit logic test suite.
.EXAMPLE
  PS> ./tools/test.ps1 -Outlook
  Runs the logic tests plus the Outlook integration tests.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding()]
param (
  [switch]
  $Outlook
)

$ErrorActionPreference = 'Stop'

$repositoryRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Split-Path -Path $PSScriptRoot -Parent))
$testsPath = Join-Path -Path $repositoryRoot -ChildPath 'tests'

if (-not (Test-Path -LiteralPath $testsPath -PathType Container)) {
  Write-Output "No test directory found: $testsPath"
  exit 0
}

$env:WINKIT_TEST_OUTLOOK = if ($Outlook) { '1' } else { '0' }
if ($Outlook) {
  Write-Output 'Outlook integration tests enabled.'
}

$_testFiles = @(Get-ChildItem -LiteralPath $testsPath -Recurse -Filter '*.Tests.ps1' -File)
if (-not $Outlook) {
  $_testFiles = @($_testFiles | Where-Object { $_.Name -notlike '*Integration*' })
}

if ($_testFiles.Count -eq 0) {
  Write-Output 'No tests to run. Add Pester test files under tests/, or pass -Outlook for the integration suite.'
  exit 0
}

$configuration = New-PesterConfiguration
$configuration.Run.Path = $testsPath
$configuration.Run.Exit = $true
$configuration.Output.Verbosity = 'Normal'

if (-not $Outlook) {
  $configuration.Run.ExcludePath = @('*Integration*')
}

Invoke-Pester -Configuration $configuration
exit $LASTEXITCODE
