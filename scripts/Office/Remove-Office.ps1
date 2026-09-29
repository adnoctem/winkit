#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.7.4' }

<#
.SYNOPSIS
  Inventories or removes explicitly selected Click-to-Run Office products.
.DESCRIPTION
  Uses PSFoundation 1.7.4 for removal planning, confirmation, execution, and
  verification. Never defaults to removing all Office. Selected products already
  absent are a successful no-op. Unselected products must remain unchanged.
  Standalone MSI removal is unsupported; use an approved migration where applicable.
  No automatic reboot or custom deletion of user data, profiles, or licensing keys.
.PARAMETER Mode
  Mandatory Check or Remove. Check without a selection returns inventory.
.PARAMETER RemoveProductId
  Exact Click-to-Run product IDs. Required for Remove; optional for a Check plan.
  Removal includes the selected products' installed languages.
.PARAMETER OdtPath
  Existing Microsoft-signed Office Deployment Tool setup.exe. Required for
  execution and preparation; never downloaded implicitly.
.PARAMETER LogRoot
  Protected local journal/log directory. Defaults to ProgramData\PSFoundation-Office.
  Recover must use the directory containing the original run's journal.
.PARAMETER ForceCloseApps
  Explicitly allow closing Office applications across sessions; unsaved work can be lost.
.PARAMETER DryRun
  Validate and preview without changing Office, downloading, or writing logs/journals.
.PARAMETER PassThru
  Return structured results, including module reason codes and recovery information.
.EXAMPLE
  .\Remove-Office.ps1 -Mode Check -PassThru
.EXAMPLE
  .\Remove-Office.ps1 -Mode Check -RemoveProductId O365ProPlusRetail -PassThru
.EXAMPLE
  .\Remove-Office.ps1 -Mode Remove -RemoveProductId O365ProPlusRetail -OdtPath C:\ODT\setup.exe -DryRun
.EXAMPLE
  .\Remove-Office.ps1 -Mode Remove -RemoveProductId O365ProPlusRetail -OdtPath C:\ODT\setup.exe -Confirm:$false -PassThru
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not supported.
  SYSTEM-account execution: supported where PSFoundation permits the host and
  media is accessible. Microsoft 365 activation requires the licensed user's session.
  Current native verification and recovery limits are documented in README.md.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'PSFoundation owns confirmation and execution; WhatIf, DryRun, and explicit Confirm are forwarded.')]
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [Parameter(Mandatory = $true)]
  [ValidateSet('Check', 'Remove')]
  [string]
  $Mode,

  [ValidateNotNullOrEmpty()]
  [string[]]
  $RemoveProductId = @(),

  [ValidateNotNullOrEmpty()]
  [string]
  $OdtPath,

  [ValidateNotNullOrEmpty()]
  [string]
  $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

  [switch]
  $ForceCloseApps,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------
# Validate the operation before calling any deployment API.
# -----------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

if ($DryRun) {
  $WhatIfPreference = $true
}

$_result = $null
$_exitCode = 0
$_phase = 'Validate'
$_mayHaveChanged = $false
$_preview = [bool]$WhatIfPreference

try {
  Write-Log -Message "Office $Mode - validating options..." -Color Cyan
  Write-Progress -Id 51 -Activity 'Office removal' -Status 'Validating options' -PercentComplete -1
  if ($Mode -eq 'Check') {
    foreach ($_name in @('OdtPath', 'LogRoot', 'ForceCloseApps')) {
      if ($PSBoundParameters.ContainsKey($_name)) {
        throw "Parameter $_name is not valid in Mode Check."
      }
    }
  }
  elseif (-not $RemoveProductId.Count -or -not $OdtPath) {
    throw 'Remove requires RemoveProductId and OdtPath.'
  }

  if ($OdtPath) {
    $OdtPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OdtPath)
  }

  if ($LogRoot) {
    $LogRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogRoot)
  }

  $_phase = 'Plan'
  Write-Log -Message 'Reading installed Office products and configuration...' -Color Cyan
  Write-Progress -Id 51 -Activity 'Office removal' -Status 'Reading Office inventory' -PercentComplete -1
  $_inventory = Get-OfficeInventory
  $_plan = $null

  if ($RemoveProductId.Count) {
    Write-Progress -Id 51 -Activity 'Office removal' -Status 'Evaluating selected products and removal eligibility' -PercentComplete -1
    $_plan = Get-OfficeDeploymentPlan -Action Remove -RemoveProductId $RemoveProductId -Inventory $_inventory
  }

  if ($Mode -eq 'Check') {
    $_status = 'Completed'
    $_reason = 'InventoryRead'

    if ($_plan) {
      foreach ($_warning in $_plan.Warnings) {
        Write-Warning $_warning
      }

      $_reason = 'PlanReady'

      if (-not $_plan.Eligible) {
        $_status = 'Blocked'
        $_reason = $_plan.Blockers -join ', '
        $_exitCode = 1
      }
    }

    $_property = @{
      SchemaVersion   = 1
      Phase           = 'Check'
      ReasonCode      = $_reason
      Inventory       = $_inventory
      Plan            = $_plan
      Changed         = $false
      ChangeKnown     = $true
      RebootRequired  = $false
      WrapperExitCode = $_exitCode
    }

    $_result = New-OperationResult `
      -Target $env:COMPUTERNAME `
      -Source Office `
      -Action Check `
      -Status $_status `
      -Property $_property
    Write-Log -Message ("Click-to-Run: {0}; MSI: {1}" -f ($_inventory.Products.ProductId -join ', '), ($_inventory.Msi.Name -join ', ')) -Color Gray
  }
  else {
    # PSFoundation rechecks the selection and confirms the complete removal scope.
    $_execution = @{
      OdtPath        = $OdtPath
      ForceCloseApps = [bool]$ForceCloseApps
      DryRun         = [bool]$DryRun
      WhatIf         = $_preview
      ErrorAction    = 'Stop'
    }

    if ($PSBoundParameters.ContainsKey('Confirm')) {
      $_execution.Confirm = [bool]$PSBoundParameters['Confirm']
    }

    if ($PSBoundParameters.ContainsKey('Verbose')) {
      $_execution.Verbose = [bool]$PSBoundParameters['Verbose']
    }

    $_phase = 'Remove'
    $_mayHaveChanged = -not $_preview
    $_progressStatus = if ($_preview) { 'Previewing Office removal' } else { 'Validating and removing Office; confirmation may be required' }
    Write-Log -Message $_progressStatus -Color Cyan
    Write-Progress -Id 51 -Activity 'Office removal' -Status $_progressStatus -PercentComplete -1
    $_result = Uninstall-Office -Plan $_plan -LogRoot $LogRoot @_execution
    $_exitCode = [int]$_result.WrapperExitCode
  }

}
catch {
  $_exitCode = 1
  $_reason = $_.Exception.Data['OfficeReason']

  if (-not $_reason) {
    $_reason = 'WrapperFailed'
  }

  $_property = @{
    SchemaVersion   = 1
    Phase           = $_phase
    ReasonCode      = $_reason
    Error           = $_.Exception.Message
    Changed         = $false
    ChangeKnown     = (-not $_mayHaveChanged)
    RebootRequired  = $false
    WrapperExitCode = 1
  }

  if ($_mayHaveChanged) {
    $_property.Changed = $null
  }

  $_result = New-OperationResult `
    -Target $env:COMPUTERNAME `
    -Source Office `
    -Action $Mode `
    -Status Failed `
    -Property $_property
}

finally {
  Write-Progress -Id 51 -Activity 'Office removal' -Completed
}

# -----------------------------------------------------------------------------
# Keep the module's detailed outcome intact for unattended callers.
# -----------------------------------------------------------------------------

$_color = 'Green'

if ($_exitCode -ne 0) {
  $_color = 'Yellow'
}

if ($_exitCode -eq 1) {
  $_color = 'Red'
}

Write-Log -Message ("Office {0}: {1} ({2})" -f $Mode, $_result.Status, $_result.ReasonCode) -Color $_color

if ($_result.Error) {
  Write-Log -Message $_result.Error -Color Red
}

if ($_result.RecoveryPath) {
  Write-Log -Message ("Recovery journal: {0}" -f $_result.RecoveryPath) -Color Gray
}

if ($PassThru -or $_preview) {
  $_result
}

if ($_exitCode) {
  exit $_exitCode
}
