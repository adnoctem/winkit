<#
  Initial setup script used to install PowerShell module dependencies
  declared in the project's requirements.psd1 manifest.

  .PARAMETER Force
    Reinstall all modules even if the required version is already present.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'This script is intended for interactive use and Write-Host is appropriate for user feedback.')]

[CmdletBinding()]
param(
  [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ---- Ensure NuGet provider (required by PowerShellGet) ----------------------
$null = Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue

# ---- Read requirements ------------------------------------------------------
$RepositoryRoot = Split-Path -Path $PSScriptRoot -Parent
$requirementsPath = Join-Path -Path $RepositoryRoot -ChildPath 'requirements.psd1'

if (-not (Test-Path -LiteralPath $requirementsPath)) {
  Write-Error "Requirements file not found: $requirementsPath"
  exit 1
}

$requirements = Import-PowerShellDataFile -Path $requirementsPath -ErrorAction Stop

foreach ($name in $requirements.Keys) {
  $targetVersion = $requirements[$name]

  Write-Host "Ensuring module '$name' (>= $targetVersion) is installed.." -ForegroundColor Yellow

  $installed = Get-Module -ListAvailable -Name $name |
    Sort-Object Version -Descending |
    Select-Object -First 1

  if ($installed -and -not $Force) {
    if ($installed.Version -ge [version]$targetVersion) {
      Write-Host "    -> OK (found $($installed.Version))" -ForegroundColor Green
      continue
    }
  }

  Write-Host "    -> Installing $name $targetVersion ..." -ForegroundColor Yellow

  $installParams = @{
    Name = $name
    RequiredVersion = $targetVersion
    Scope = 'CurrentUser'
    Force = $true
    AllowClobber = $true
    SkipPublisherCheck = $true
    ErrorAction = 'Stop'
  }

  try {
    Install-Module @installParams
    Write-Host "    -> Installed $name $targetVersion" -ForegroundColor Green
  }
  catch {
    Write-Host "    -> Failed to install $name : $_" -ForegroundColor Red
  }
}

Write-Host "Successfully processed all required modules!" -ForegroundColor Yellow
