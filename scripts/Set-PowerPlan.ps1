#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Activates a Windows power plan and sets fast process-kill timers.

.DESCRIPTION
  Activates the requested power scheme via powercfg and writes
  controlled registry values for service process timeouts. The
  Ultimate Performance scheme is revealed automatically when it is
  hidden.

  This script is opt-in and not included in any default optimizer
  profile. Power plans are a genuine performance tradeoff —
  especially on battery-powered machines.

.PARAMETER Plan
  Power plan to activate. Defaults to UltimatePerformance.
  Supported: UltimatePerformance, HighPerformance, Balanced, PowerSaver.

.PARAMETER Undo
  Restore the Balanced power plan and default process-kill timers.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Set-PowerPlan.ps1
  Activates Ultimate Performance and sets fast process-kill timers.

.EXAMPLE
  PS> ./Set-PowerPlan.ps1 -Plan HighPerformance
  Activates High Performance instead.

.EXAMPLE
  PS> ./Set-PowerPlan.ps1 -Undo
  Restores Balanced plan and default timeout values.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Power plan to activate.'
  )]
  [ValidateSet('UltimatePerformance', 'HighPerformance', 'Balanced', 'PowerSaver')]
  [string]
  $Plan = 'UltimatePerformance',

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Restore Balanced plan and default timers.'
  )]
  [switch]
  $Undo,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Preview changes without applying them.'
  )]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

$_schemeGuids = @{
  'UltimatePerformance' = 'e9a42b02-d5df-448d-aa00-03f14749eb61'
  'HighPerformance' = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
  'Balanced' = '381b4222-f694-41f0-9685-ff5bb260df2e'
  'PowerSaver' = 'a1841308-3541-4fab-bc81-f71556f20b4a'
}

$_controlKey = 'HKLM:\SYSTEM\CurrentControlSet\Control'
$_timeoutDefaults = @{
  'WaitToKillServiceTimeout' = 5000
  'WaitToKillAppTimeout' = 5000
  'HungAppTimeout' = 5000
}
$_timeoutValues = @{
  'WaitToKillServiceTimeout' = 2000
  'WaitToKillAppTimeout' = 2000
  'HungAppTimeout' = 2000
}

# ---- Undo --------------------------------------------------------------------
if ($Undo) {
  Write-Log -Message 'Restoring Balanced power plan and default process-kill timers.' -Color Yellow

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would restore Balanced plan and default timers." -Color Yellow
    Add-OperationResult -Results $_results -Target 'PowerPlan' -Source 'PowerPlan' -Action 'Restore' -Status 'Skipped' -Detail 'DryRun'
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }

  if (-not $PSCmdlet.ShouldProcess('Power plan', 'Restore Balanced plan and default timers')) {
    Add-OperationResult -Results $_results -Target 'PowerPlan' -Source 'PowerPlan' -Action 'Restore' -Status 'Skipped' -Detail 'WhatIf'
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }

  $_balancedGuid = $_schemeGuids['Balanced']
  try {
    $null = powercfg /setactive $_balancedGuid 2>&1
    Write-Log -Message '  -> Balanced power plan activated.' -Color Green
  }
  catch {
    Write-Log -Message "  -> FAILED to activate Balanced plan: $_" -Color Red
    Add-OperationResult -Results $_results -Target 'PowerPlan' -Source 'PowerPlan' -Action 'Restore' -Status 'Failed' -Detail $_.Exception.Message
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }

  foreach ($_name in $_timeoutDefaults.Keys) {
    $_defaultVal = $_timeoutDefaults[$_name]
    try {
      Set-ItemProperty -LiteralPath $_controlKey -Name $_name -Value $_defaultVal -Type String -ErrorAction Stop
      Write-Log -Message "  -> $_name = $_defaultVal" -Color Green
    }
    catch {
      Write-Log -Message "  -> FAILED to set $_name = $_defaultVal: $_" -Color Red
    }
  }

  Write-Log -Message '  -> Default timers restored.' -Color Green
  Add-OperationResult -Results $_results -Target 'PowerPlan' -Source 'PowerPlan' -Action 'Restore' -Status 'Completed' -Detail 'Balanced plan + default timers restored.'

  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Set-PowerPlan'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

# ---- Reveal Ultimate Performance if needed -----------------------------------
$_targetGuid = $_schemeGuids[$Plan]

if ($Plan -eq 'UltimatePerformance') {
  $_existingSchemes = powercfg /list 2>&1 | Out-String
  if ($_existingSchemes -notmatch [regex]::Escape($_targetGuid)) {
    Write-Log -Message 'Ultimate Performance scheme is hidden - revealing it.' -Color Yellow

    if ($DryRun) {
      Write-Log -Message "[DRY RUN] Would reveal Ultimate Performance via powercfg." -Color Yellow
    }
    else {
      try {
        $null = powercfg -duplicatescheme $_targetGuid 2>&1
        Write-Log -Message '  -> Ultimate Performance revealed.' -Color Green
        Add-OperationResult -Results $_results -Target 'UltimatePerformance' -Source 'PowerPlan' -Action 'Reveal' -Status 'Completed' -Detail 'Duplicate scheme created.'
      }
      catch {
        Write-Log -Message "  -> FAILED to reveal Ultimate Performance: $_" -Color Red
        Add-OperationResult -Results $_results -Target 'UltimatePerformance' -Source 'PowerPlan' -Action 'Reveal' -Status 'Failed' -Detail $_.Exception.Message
        if ($PassThru -or $DryRun) { $_results }
        exit 1
      }
    }
  }
}

# ---- Activate plan -----------------------------------------------------------
Write-Log -Message "Activating power plan: $Plan" -Color Yellow

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would activate plan GUID $_targetGuid" -Color Yellow
}
else {
  try {
    $null = powercfg /setactive $_targetGuid 2>&1
    Write-Log -Message "  -> $Plan power plan activated." -Color Green
    Add-OperationResult -Results $_results -Target $Plan -Source 'PowerPlan' -Action 'Activate' -Status 'Completed' -Detail "Power plan activated: $Plan"
  }
  catch {
    Write-Log -Message "  -> FAILED to activate $Plan plan: $_" -Color Red
    Add-OperationResult -Results $_results -Target $Plan -Source 'PowerPlan' -Action 'Activate' -Status 'Failed' -Detail $_.Exception.Message
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }
}

# ---- Set process-kill timers -------------------------------------------------
Write-Log -Message 'Setting fast process-kill timers.' -Color Yellow

if ($DryRun) {
  foreach ($_name in $_timeoutValues.Keys) {
    Write-Log -Message "[DRY RUN] Would set $_name = $($_timeoutValues[$_name])" -Color Yellow
  }
}
else {
  foreach ($_name in $_timeoutValues.Keys) {
    $_val = $_timeoutValues[$_name]
    try {
      Set-ItemProperty -LiteralPath $_controlKey -Name $_name -Value $_val -Type String -ErrorAction Stop
      Write-Log -Message "  -> $_name = $_val" -Color Green
    }
    catch {
      Write-Log -Message "  -> FAILED to set $_name = $_val: $_" -Color Red
    }
  }
  Write-Log -Message '  -> Process-kill timers applied.' -Color Green
}

if (-not $DryRun) {
  Write-Log -Message "`nPower plan '$Plan' is now active. A restart is not required." -Color Green
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Set-PowerPlan'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
