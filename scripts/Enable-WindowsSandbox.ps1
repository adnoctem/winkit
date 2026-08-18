#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Enables the Windows Sandbox optional feature.

.DESCRIPTION
  Enables the Containers-DisposableClientVM Windows feature, which provides
  the Windows Sandbox isolated desktop environment. A reboot is required
  after enabling. Checks whether the feature is already enabled and skips
  if so.

  Requires administrator elevation.

.PARAMETER DryRun
  Preview the enable operation without applying it.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Enable-WindowsSandbox.ps1

.EXAMPLE
  PS> ./Enable-WindowsSandbox.ps1 -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
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
# Windows Sandbox requires Windows 10 1809 / Server 2019 (build 17763) or
# newer, on the Professional, Enterprise or Education edition.
$minSupportedBuild = 17763
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
  Add-OperationResult -Results $_results -Target 'WindowsSandbox' -Source 'WindowsFeature' -Action 'Validate' -Status 'Failed' -Detail "Unsupported OS - build $($osVersion.CurrentBuild), minimum $minSupportedBuild."
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}
$edition = Get-OSEdition
if ($edition -notin @('Professional', 'Enterprise', 'Education')) {
  Write-Log -Message "Unsupported edition: Windows Sandbox requires Professional, Enterprise or Education (current: $edition)." -Color Red
  Add-OperationResult -Results $_results -Target 'WindowsSandbox' -Source 'WindowsFeature' -Action 'Validate' -Status 'Failed' -Detail "Unsupported edition - $edition."
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

$_featureName = 'Containers-DisposableClientVM'

$state = Get-WindowsOptionalFeature -Online -FeatureName $_featureName -ErrorAction SilentlyContinue
if ($state -and $state.State -eq 'Enabled') {
  Write-Log -Message 'Windows Sandbox is already enabled.' -Color Green
  Add-OperationResult -Results $_results -Target $_featureName -Source 'WindowsFeature' -Action 'Enable' -Status 'Skipped' -Detail 'AlreadyEnabled'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would enable Windows Sandbox." -Color Yellow
  Write-Log -Message '  -> A reboot will be required after enabling.' -Color Gray
  Add-OperationResult -Results $_results -Target $_featureName -Source 'WindowsFeature' -Action 'Enable' -Status 'Skipped' -Detail 'DryRun'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if (-not $PSCmdlet.ShouldProcess($_featureName, 'Enable Windows optional feature')) {
  Add-OperationResult -Results $_results -Target $_featureName -Source 'WindowsFeature' -Action 'Enable' -Status 'Skipped' -Detail 'WhatIf'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

Write-Log -Message 'Enabling Windows Sandbox...' -Color Yellow
try {
  $null = Enable-WindowsOptionalFeature -Online -FeatureName $_featureName -NoRestart -ErrorAction Stop
  Write-Log -Message "  -> Windows Sandbox enabled (pending reboot)." -Color Green
  Add-OperationResult -Results $_results -Target $_featureName -Source 'WindowsFeature' -Action 'Enable' -Status 'Completed' -Detail 'Reboot required.'
}
catch {
  Write-Log -Message "  -> FAILED - could not enable Windows Sandbox: $_" -Color Red
  Add-OperationResult -Results $_results -Target $_featureName -Source 'WindowsFeature' -Action 'Enable' -Status 'Failed' -Detail $_.Exception.Message
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Enable-WindowsSandbox'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
