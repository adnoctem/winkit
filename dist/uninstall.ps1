#Requires -Version 5.0

<#
.SYNOPSIS
  Removes this registered winkit installation using its local maintenance engine.
.DESCRIPTION
  Runs offline from an installed release. The installation path and scope are
  bound to the adjacent ownership manifest, not inherited environment settings.
  WINKIT_DRY_RUN previews removal. WINKIT_NON_INTERACTIVE bypasses the explicit
  confirmation; WINKIT_PASS_THRU returns the maintenance result. Shared modules
  and package providers remain installed. Added/modified files block removal.
.EXAMPLE
  PS> & 'C:\Program Files\winkit\dist\uninstall.ps1'
  Requests confirmation before removing this machine-wide installation.
.EXAMPLE
  PS> $env:WINKIT_DRY_RUN = '1'
  PS> & "$env:LOCALAPPDATA\Programs\winkit\dist\uninstall.ps1"
  Previews removal without downloads, prompts, or changes.
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core support: Yes.
  SYSTEM-account suitability: AllUsers only; CurrentUser refers to that account.
  This standalone distribution entry point has no public script parameters.
#>

& {
  param (
    [string]
    $EntryPath,

    [object[]]
    $SuppliedArguments
  )

  Set-StrictMode -Version 2.0
  $ErrorActionPreference = 'Stop'
  if ($SuppliedArguments.Count) {
    throw 'Configure removal with WINKIT_* environment variables, not parameters.'
  }
  if (-not $EntryPath) {
    throw 'Run the installed local uninstall.ps1 file. For web removal use WINKIT_UNINSTALL=1 with dist/install.ps1.'
  }

  $directory = Split-Path $EntryPath -Parent
  $root = Split-Path $directory -Parent
  $engine = Join-Path $directory 'install.ps1'
  $state = Get-Content -LiteralPath (Join-Path $root '.winkit-install.json') -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
  if ($state.schemaVersion -ne 2 -or $state.installPath -ne $root -or $state.scope -notin @('CurrentUser', 'AllUsers')) {
    throw 'This uninstaller is not inside a recognized winkit installation.'
  }
  $engineFile = @($state.files | Where-Object Path -EQ 'dist\install.ps1')
  if ($engineFile.Count -ne 1 -or (Get-FileHash -LiteralPath $engine -Algorithm SHA256).Hash -ne $engineFile[0].Sha256) {
    throw 'The local maintenance engine does not match the ownership manifest.'
  }

  $preview = $false
  $nonInteractive = $false
  foreach ($setting in @('WINKIT_DRY_RUN', 'WINKIT_NON_INTERACTIVE')) {
    $value = [Environment]::GetEnvironmentVariable($setting, 'Process')
    if (-not [string]::IsNullOrWhiteSpace($value)) {
      $value = $value.Trim().ToLowerInvariant()
      if ($value -notin @('1', 'true', 'yes', 'on', '0', 'false', 'no', 'off')) {
        throw "$setting must be 1/true/yes/on or 0/false/no/off."
      }
      $enabled = $value -in @('1', 'true', 'yes', 'on')
      if ($setting -eq 'WINKIT_DRY_RUN') { $preview = $enabled }
      else { $nonInteractive = $enabled }
    }
  }
  $preview = $preview -or $WhatIfPreference
  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  $elevationRequired = $state.scope -eq 'AllUsers' -and -not $preview -and
  -not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  $nativeHostRequired = [Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess
  if ($elevationRequired -or $nativeHostRequired) {
    if ($elevationRequired -and $nonInteractive) {
      throw 'Non-interactive AllUsers removal requires an already elevated session.'
    }
    $systemDirectory = if ($nativeHostRequired) { 'Sysnative' } else { 'System32' }
    $hostPath = Join-Path $env:SystemRoot "$systemDirectory\WindowsPowerShell\v1.0\powershell.exe"
    $start = @{
      FilePath     = $hostPath
      ArgumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $EntryPath))
      Wait         = $true
      PassThru     = $true
      # The removal confirmation must remain visible in the interactive host.
      WindowStyle  = 'Normal'
    }
    if ($elevationRequired) {
      $start.Verb = 'RunAs'
    }
    $savedPreview = [Environment]::GetEnvironmentVariable('WINKIT_DRY_RUN', 'Process')
    try {
      if ($preview) {
        $env:WINKIT_DRY_RUN = '1'
      }
      $process = Start-Process @start
    }
    finally {
      [Environment]::SetEnvironmentVariable('WINKIT_DRY_RUN', $savedPreview, 'Process')
    }
    if ($process.ExitCode -ne 0) {
      throw "Uninstaller exited with code $($process.ExitCode)."
    }
    return
  }

  $overrides = @{
    WINKIT_UNINSTALL    = '1'
    WINKIT_INSTALL_PATH = $root
    WINKIT_SCOPE        = $state.scope
    WINKIT_REPOSITORY   = $null
    WINKIT_VERSION      = $null
    WINKIT_FORCE        = $null
    WINKIT_NO_PATH      = $null
    WINKIT_CHANGE_SCOPE = $null
  }
  $saved = @{}
  try {
    foreach ($name in $overrides.Keys) {
      $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
      [Environment]::SetEnvironmentVariable($name, $overrides[$name], 'Process')
    }
    & $engine
  }
  finally {
    foreach ($name in $saved.Keys) {
      [Environment]::SetEnvironmentVariable($name, $saved[$name], 'Process')
    }
  }
} $PSCommandPath @($args)
