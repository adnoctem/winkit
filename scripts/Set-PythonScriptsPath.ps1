#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Adds the active Python Scripts directory to the user or machine PATH.

.DESCRIPTION
  Resolves the active Python installation's Scripts directory using
  `python -c "import sysconfig; ..."` and appends that directory to the user or
  machine PATH if it is not already present. This is useful for Store / Python
  Install Manager layouts where `python -m pip install pre-commit` succeeds but
  console scripts such as pre-commit.exe are not discoverable in new shells.

.PARAMETER Scope
  PATH scope to update. Defaults to User. Machine requires elevation.

.PARAMETER Python
  Python launcher or executable to query. Defaults to python.

.PARAMETER DryRun
  Show the PATH change without writing it.

.PARAMETER PassThru
  Return the resolved Scripts path.

.EXAMPLE
  PS> ./Set-PythonScriptsPath.ps1
  Adds the active Python Scripts directory to the current user's PATH.

.EXAMPLE
  PS> ./Set-PythonScriptsPath.ps1 -Scope Machine
  Adds the active Python Scripts directory to the machine PATH. Requires elevation.

.EXAMPLE
  PS> ./Set-PythonScriptsPath.ps1 -DryRun -PassThru
  Shows and returns the Scripts directory without changing PATH.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [ValidateSet('User', 'Machine')]
  [string]$Scope = 'User',

  [string]$Python = 'python',

  [switch]$DryRun,

  [switch]$PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

# Elevation guard: -Scope Machine requires admin
if ($Scope -eq 'Machine' -and -not (Test-Elevation)) {
  Write-Error '-Scope Machine requires administrator privileges. Use -Scope User (default) for per-user PATH.'
  exit 1
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - PATH will not be modified`n" -Color Yellow
}

try {
  $pythonCommand = Get-Command -Name $Python -ErrorAction Stop
}
catch {
  Write-Log -Message "Python executable not found: $Python" -Color Red
  exit 1
}

try {
  $scriptsPath = & $pythonCommand.Source -c "import sysconfig; print(sysconfig.get_path('scripts'))" 2>$null
  $scriptsPath = $scriptsPath | Select-Object -First 1
}
catch {
  Write-Log -Message "Failed to resolve Python Scripts path from '$($pythonCommand.Source)': $_" -Color Red
  exit 1
}

if ([string]::IsNullOrWhiteSpace($scriptsPath)) {
  Write-Log -Message "Python did not report a Scripts path." -Color Red
  exit 1
}

$scriptsPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($scriptsPath)
if (-not (Test-Path -LiteralPath $scriptsPath -PathType Container)) {
  Write-Log -Message "Python Scripts path does not exist: $scriptsPath" -Color Red
  exit 1
}

$currentPath = [Environment]::GetEnvironmentVariable('Path', $Scope)
if ([string]::IsNullOrWhiteSpace($currentPath)) {
  $pathEntries = @()
}
else {
  $pathEntries = @($currentPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

$alreadyPresent = @($pathEntries | Where-Object { $_.TrimEnd('\') -ieq $scriptsPath.TrimEnd('\') }).Count -gt 0
if ($alreadyPresent) {
  Write-Log -Message "Python Scripts path is already present in $Scope PATH: $scriptsPath" -Color Green
}
else {
  $newPath = if ($pathEntries.Count -gt 0) {
    (($pathEntries + $scriptsPath) -join ';')
  }
  else {
    $scriptsPath
  }

  if ($PSCmdlet.ShouldProcess("$Scope PATH", "Append '$scriptsPath'")) {
    [Environment]::SetEnvironmentVariable('Path', $newPath, $Scope)
    Write-Log -Message "Added Python Scripts path to $Scope PATH: $scriptsPath" -Color Green
  }
}

if (-not $DryRun -and (($env:Path -split ';') -notcontains $scriptsPath)) {
  $env:Path = "$env:Path;$scriptsPath"
}

if ($PassThru) {
  $scriptsPath
}
