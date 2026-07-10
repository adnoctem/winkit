Import-Module PSFoundation -Force

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
  Author: Maximilian Gindorfer <info@mvprowess.com>
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

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
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
