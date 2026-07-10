Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Installs a Visual C++ Redistributable silently.

.DESCRIPTION
  Downloads and silently installs the requested VC++ runtime, skipping if
  it is already present (detected via the registry). Architecture is
  auto-detected; ARM64 packages exist only for 14.0 â€" older versions fall
  back to x64.

  Requires administrator elevation because the installer writes to system
  locations.

.PARAMETER Version
  VC++ version to install. One of: 14.0, 12.0, 11.0, 10.0, 9.0, 8.0.
  14.0 covers the unified 2015-2022 runtime.

.PARAMETER DryRun
  Preview the install steps without executing them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Install-VCRedistributables.ps1 -Version 14.0

.EXAMPLE
  PS> ./Install-VCRedistributables.ps1 -Version 12.0 -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [ValidateSet('14.0', '12.0', '11.0', '10.0', '9.0', '8.0')]
  [string]
  $Version,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$ProgressPreference = 'SilentlyContinue'
$_results = New-Object System.Collections.ArrayList

$major = ($Version -split '\.')[0]

$arch = switch ($env:PROCESSOR_ARCHITECTURE) {
  'AMD64' { 'x64' }
  'ARM64' { 'arm64' }
  default { 'x86' }
}
if ($arch -eq 'arm64' -and $major -ne '14') {
  Write-Log -Message "ARM64 not available for VC++ $Version; falling back to x64." -Color Yellow
  $arch = 'x64'
}

# ---- Already installed? -----------------------------------------------------
$regPath = "HKLM:\SOFTWARE\Microsoft\VisualStudio\$major.0\VC\Runtimes\$arch"
if (Test-Path -LiteralPath $regPath) {
  $installed = (Get-ItemProperty -Path $regPath -Name 'Installed' -ErrorAction SilentlyContinue).Installed
  if ($installed -eq 1) {
    Write-Log -Message "VC++ $Version ($arch) is already installed." -Color Green
    Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Skipped' -Detail 'AlreadyInstalled'
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }
}

$urls = @{
  '14.0' = @{
    x64 = 'https://aka.ms/vc14/vc_redist.x64.exe'
    x86 = 'https://aka.ms/vc14/vc_redist.x86.exe'
    arm64 = 'https://aka.ms/vc14/vc_redist.arm64.exe'
  }
  '12.0' = @{
    x64 = 'https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x64.exe'
    x86 = 'https://download.microsoft.com/download/2/E/6/2E61CFA4-993B-4DD4-91DA-3737CD5CD6E3/vcredist_x86.exe'
  }
  '11.0' = @{
    x64 = 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe'
    x86 = 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe'
  }
  '10.0' = @{
    x64 = 'https://download.microsoft.com/download/A/8/0/A80747C3-41BD-45DF-B505-E9710D2744E0/vcredist_x64.exe'
    x86 = 'https://download.microsoft.com/download/C/6/D/C6D0FD4E-9E53-4897-9B91-836EBA2FA716/vcredist_x86.exe'
  }
  '9.0' = @{
    x64 = 'https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe'
    x86 = 'https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe'
  }
  '8.0' = @{
    x64 = 'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.EXE'
    x86 = 'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x86.EXE'
  }
}

$downloadUrl = $urls[$Version][$arch]
if (-not $downloadUrl) {
  Write-Log -Message "No download URL for VC++ $Version ($arch)." -Color Red
  Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Failed' -Detail "No URL for $Version-$arch"
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would install VC++ $Version ($arch)." -Color Yellow
  Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Skipped' -Detail 'DryRun'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

$work = Join-Path $env:TEMP "vcredist-install-$(New-Guid)"
$null = New-Item -Path $work -ItemType Directory -Force -ErrorAction SilentlyContinue
try {
  $installer = Join-Path $work "vcredist_$arch.exe"
  Write-Log -Message "Downloading VC++ $Version ($arch)..." -Color Gray
  Invoke-WebRequest -Uri $downloadUrl -OutFile $installer

  if (-not $PSCmdlet.ShouldProcess("VC++ $Version ($arch)", 'Install silently')) {
    Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  Write-Log -Message "Installing VC++ $Version ($arch)..." -Color Yellow
  $proc = Start-Process -FilePath $installer -ArgumentList '/install', '/quiet', '/norestart' -Wait -PassThru

  if ($proc.ExitCode -notin @(0, 3010)) {
    Write-Log -Message "  -> VC++ installation failed with exit code $($proc.ExitCode)." -Color Red
    Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Failed' -Detail "Exit code $($proc.ExitCode)"
    exit $proc.ExitCode
  }

  if ($proc.ExitCode -eq 3010) {
    Write-Log -Message "  -> VC++ $Version ($arch) installed; reboot required to complete." -Color Yellow
    Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Completed' -Detail 'Installation requires reboot.'
  }
  else {
    Write-Log -Message "  -> VC++ $Version ($arch) installed successfully." -Color Green
    Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Completed' -Detail 'Installed successfully.'
  }
}
catch {
  Write-Log -Message "  -> FAILED - could not install VC++ ${Version}: $_" -Color Red
  Add-OperationResult -Results $_results -Target "VCRedist-$Version-$arch" -Source 'VCRedist' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
}
finally {
  Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-VCRedistributables'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

exit 0
