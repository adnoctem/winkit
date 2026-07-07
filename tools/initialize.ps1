<#
  Initial setup script used to download PowerShell module dependencies
  defined in the project manifest and set up the project for local use.

  .PARAMETER Force
    Reinstall all modules even if the required version is already present.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'This script is intended for interactive use and Write-Host is appropriate for user feedback.')]

[CmdletBinding()]
param(
  [switch]$Force
)

# ---- Ensure NuGet provider (required by PowerShellGet) ----------------------
$null = Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue

# ---- Configure module -------------------------------------------------------
$RepositoryRoot = Split-Path -Path $PSScriptRoot -Parent
$manifestPath = Join-Path -Path $RepositoryRoot -ChildPath 'lib/winkit.psd1'
$manifest = Test-ModuleManifest -Path $manifestPath -ErrorAction Stop

Write-Host "Using manifest: $manifest"

foreach ($mod in $manifest.RequiredModules) {
  if ($mod -is [string]) {
    $name = $mod
    $minVer = $null
    $exactVer = $null
  }
  else {
    $name = $mod.Name
    $minVer = $mod.Version          # from ModuleVersion key (minimum)
    $exactVer = $mod.RequiredVersion # from RequiredVersion key (exact)
  }

  Write-Host "Ensuring module '$name' is installed.." -ForegroundColor Yellow

  $installed = Get-Module -ListAvailable -Name $name |
    Sort-Object Version -Descending |
    Select-Object -First 1

  # ---- Determine whether current install satisfies the requirement ----
  $satisfied = $false
  if ($installed -and -not $Force) {
    if ($exactVer) {
      $satisfied = $installed.Version -eq $exactVer
    }
    elseif ($minVer) {
      $satisfied = $installed.Version -ge $minVer
    }
    # No version constraint → latest is always acceptable
  }

  if ($satisfied) {
    Write-Host "    -> OK (found $($installed.Version))" -ForegroundColor Green
    continue
  }

  # ---- Resolve target version ----
  if ($exactVer) {
    $targetVer = $exactVer
  }
  elseif ($minVer) {
    # Install the exact minimum version from the manifest. The satisfaction
    # check above uses -ge so newer compatible versions already installed
    # are accepted (not downgraded).
    $targetVer = $minVer
  }
  else {
    $target = Find-Module -Name $name |
      Sort-Object Version -Descending |
      Select-Object -First 1
    if (-not $target) {
      Write-Host "    -> '$name' not found in gallery." -ForegroundColor Red
      continue
    }
    $targetVer = $target.Version
  }

  Write-Host "    -> Installing $name $targetVer ..." -ForegroundColor Yellow

  $installParams = @{
    Name = $name
    RequiredVersion = $targetVer.ToString()
    Scope = 'CurrentUser'
    Force = $true
    AllowClobber = $true
    SkipPublisherCheck = $true
  }

  Install-Module @installParams
  Write-Host "    -> Installed $name $targetVer" -ForegroundColor Green
}

# ---------------------------------------------------------------
Write-Host "Successfully processed all RequiredModules!" -ForegroundColor Yellow
