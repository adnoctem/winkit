#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Installs or updates PowerShell 7+ (pwsh) via winget.

.DESCRIPTION
  Checks whether PowerShell 7 is already available and, if not, uses
  winget to install the latest stable release. Designed for first-run
  machine bootstrap â€" the script is deliberately light and relies on
  winget being present.

  Requires administrator elevation because winget installs at machine
  scope.

.PARAMETER DryRun
  Preview the install step without executing it.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Install-PowerShellCore.ps1

.EXAMPLE
  PS> ./Install-PowerShellCore.ps1 -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - no Desktop Experience dependency.
  SYSTEM-account execution: works; installation involves no per-user registration.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# ---- OS compatibility gate ---------------------------------------------------
# PowerShell 7.x supports Windows 10 1607 / Server 2016 (build 14393) and newer.
$minSupportedBuild = 14393
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
$isServerOs = ($null -ne $osInfo -and [int]$osInfo.ProductType -ge 2)
if ($osInfo -and [int]$osInfo.OperatingSystemSKU -ge 112 -and [int]$osInfo.OperatingSystemSKU -le 115) {
  # Server multi-session editions report as Server but behave like a workstation.
  $isServerOs = $false
}
if ((Get-OSBuildNumber) -lt $minSupportedBuild) {
  $osVersion = Get-OSVersionInfo
  $osLabel = if ($isServerOs) { 'Windows Server' } else { 'Windows' }
  Write-Log -Message "Unsupported OS: this script requires $osLabel build $minSupportedBuild or newer (current: $($osVersion.DisplayVersion), build $($osVersion.CurrentBuild))." -Color Red
  Add-OperationResult -Results $_results -Target 'PowerShellCore' -Source 'Winget' -Action 'Validate' -Status 'Failed' -Detail "Unsupported OS - build $($osVersion.CurrentBuild), minimum $minSupportedBuild."
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

if (Get-Command pwsh -ErrorAction SilentlyContinue) {
  $_installedVersion = & pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 2>$null
  Write-Log -Message "PowerShell ($_installedVersion) is already installed." -Color Green
  Add-OperationResult -Results $_results -Target 'PowerShellCore' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'AlreadyInstalled'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if ($DryRun) {
  Write-Log -Message '[DRY RUN] Would install PowerShell 7 via winget.' -Color Yellow
  Add-OperationResult -Results $_results -Target 'PowerShellCore' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'DryRun'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if (-not $PSCmdlet.ShouldProcess('PowerShell 7', 'Install via winget')) {
  Add-OperationResult -Results $_results -Target 'PowerShellCore' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

Write-Log -Message 'Installing PowerShell 7 via winget...' -Color Yellow
try {
  $proc = Start-Process -FilePath 'winget.exe' -ArgumentList @(
    'install', '--id', 'Microsoft.PowerShell', '--exact',
    '--silent', '--accept-package-agreements', '--accept-source-agreements'
  ) -Wait -PassThru -ErrorAction Stop

  if ($proc.ExitCode -eq 0) {
    Write-Log -Message '  -> PowerShell 7 installed successfully.' -Color Green
    Add-OperationResult -Results $_results -Target 'PowerShellCore' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Installed via winget.'
  }
  else {
    Write-Log -Message "  -> Winget exited with code $($proc.ExitCode)." -Color Yellow
    Add-OperationResult -Results $_results -Target 'PowerShellCore' -Source 'Winget' -Action 'Install' -Status "ExitCode:$($proc.ExitCode)" -Detail 'Winget install returned non-zero exit code.'
  }
}
catch {
  Write-Log -Message "  -> FAILED - could not install PowerShell 7: $_" -Color Red
  Add-OperationResult -Results $_results -Target 'PowerShellCore' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-PowerShellCore'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

exit 0
