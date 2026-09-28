#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.6.0' }

<#
.SYNOPSIS
  Plans, prepares, installs, or resumes an Office deployment.
.DESCRIPTION
  Delegates deployment to PSFoundation 1.6.0. Mode is mandatory.
  Check is read-only; with a target it returns the complete plan and blockers.
  Prepare publishes verified media. Install uses the module's validated plan.
  Recover resumes only the original installation journal.
  Conflicting installations are not replaced; a verified compliant target is a no-op.
  No automatic reboot, rollback, or Outlook profile conversion is provided.
  Preserve backups and installation media; test deployments on a recoverable pilot.
.PARAMETER Mode
  Check, Prepare, Install, or Recover. No operation is selected implicitly.
.PARAMETER TargetProductId
  Required for Prepare and deployment. Optional in Check to assess a target plan.
  Product IDs and compatibility are validated by PSFoundation.
.PARAMETER Architecture
  Target Office architecture, 32 or 64. Defaults to 64.
.PARAMETER Channel
  Derived from the volume product; Microsoft 365 defaults to Current.
.PARAMETER Language
  Ordered Office language IDs. Defaults to en-us through PSFoundation.
  The first language is primary. Cannot be combined with AutoSourceLocales.
.PARAMETER AutoSourceLocales
  Explicitly discover locales instead of defaulting to en-us. Ambiguous discovery fails.
.PARAMETER LocaleSource
  InstalledOffice or OperatingSystem. Requires AutoSourceLocales.
  OperatingSystem means the machine installation UI language, not the operator's culture.
.PARAMETER Version
  Optional exact 16.0 build. Prepare resolves a build; deployment uses the media manifest.
.PARAMETER ExcludeApp
  Applications to exclude. Values are validated by PSFoundation.
.PARAMETER ExcludePublisher
  Add Publisher to ExcludeApp.
.PARAMETER SourcePath
  Absolute or relative local/UNC package directory. Prepare needs an existing parent.
  Deployment needs compatible schema-2 media unless the target is already compliant.
.PARAMETER ProductKey
  Optional SecureString volume key. Never stored in a journal; supply it again for recovery.
.PARAMETER RunId
  Required for Recover. The 32-digit hexadecimal identifier of the original operation.
  Recovery uses its recorded target, languages, media, and authority.
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
  .\Install-Office.ps1 -Mode Check -PassThru
.EXAMPLE
  .\Install-Office.ps1 -Mode Prepare -TargetProductId Standard2024Volume -SourcePath C:\Media\Office2024 -OdtPath C:\ODT\setup.exe -Language en-us,de-de
.EXAMPLE
  .\Install-Office.ps1 -Mode Install -TargetProductId Standard2024Volume -SourcePath C:\Media\Office2024 -OdtPath C:\ODT\setup.exe -DryRun
.EXAMPLE
  .\Install-Office.ps1 -Mode Recover -RunId 0123456789abcdef0123456789abcdef -OdtPath C:\ODT\setup.exe -DryRun
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
  [ValidateSet('Check', 'Prepare', 'Install', 'Recover')]
  [string]
  $Mode,

  [ValidateNotNullOrEmpty()]
  [string]
  $TargetProductId,

  [ValidateSet('32', '64')]
  [string]
  $Architecture = '64',

  [ValidateNotNullOrEmpty()]
  [string]
  $Channel,

  [ValidateNotNullOrEmpty()]
  [string[]]
  $Language,

  [switch]
  $AutoSourceLocales,

  [ValidateSet('InstalledOffice', 'OperatingSystem')]
  [string]
  $LocaleSource,

  [ValidatePattern('^16\.0\.\d+\.\d+$')]
  [string]
  $Version,

  [string[]]
  $ExcludeApp = @(),

  [switch]
  $ExcludePublisher,

  [ValidateNotNullOrEmpty()]
  [string]
  $SourcePath,

  [ValidateNotNullOrEmpty()]
  [string]
  $OdtPath,

  [ValidateNotNullOrEmpty()]
  [string]
  $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

  [switch]
  $ForceCloseApps,

  [Security.SecureString]
  $ProductKey,

  [ValidatePattern('^[a-fA-F0-9]{32}$')]
  [string]
  $RunId,

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
  $_configurationNames = @(
    'TargetProductId',
    'Architecture',
    'Channel',
    'Language',
    'AutoSourceLocales',
    'LocaleSource',
    'Version',
    'ExcludeApp',
    'ExcludePublisher'
  )

  $_operationNames = $_configurationNames + @(
    'SourcePath', 'OdtPath', 'LogRoot', 'ForceCloseApps', 'ProductKey', 'RunId'
  )

  $_allowed = switch ($Mode) {
    'Check' { $_configurationNames + @('SourcePath') }
    'Prepare' { $_configurationNames + @('SourcePath', 'OdtPath') }
    'Install' { $_configurationNames + @('SourcePath', 'OdtPath', 'LogRoot', 'ForceCloseApps', 'ProductKey') }
    'Recover' { @('RunId', 'OdtPath', 'LogRoot', 'ForceCloseApps', 'ProductKey') }
  }

  foreach ($_name in $_operationNames) {
    if ($PSBoundParameters.ContainsKey($_name) -and $_name -notin $_allowed) {
      throw "Parameter $_name is not valid in Mode $Mode."
    }
  }

  if ($Mode -ne 'Check' -and -not $OdtPath) {
    throw "Mode $Mode requires OdtPath."
  }

  if ($Mode -eq 'Recover' -and -not $RunId) {
    throw 'Recover requires the original RunId.'
  }

  if ($Mode -in @('Prepare', 'Install') -and -not $TargetProductId) {
    throw "Mode $Mode requires TargetProductId."
  }

  if ($Mode -eq 'Prepare' -and -not $SourcePath) {
    throw 'Prepare requires SourcePath.'
  }

  if ($Mode -eq 'Check' -and -not $TargetProductId) {
    foreach ($_name in ($_configurationNames + @('SourcePath'))) {
      if ($PSBoundParameters.ContainsKey($_name)) {
        throw "Check requires TargetProductId when $_name is supplied."
      }
    }
  }

  if ($SourcePath) {
    $SourcePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourcePath)
  }

  if ($OdtPath) {
    $OdtPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OdtPath)
  }

  if ($LogRoot) {
    $LogRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogRoot)
  }

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

  # ---------------------------------------------------------------------------
  # Recovery uses recorded intent; supplied target/removal overrides were rejected.
  # ---------------------------------------------------------------------------

  if ($Mode -eq 'Recover') {
    $_phase = 'Recover'
    $_recovery = Get-OfficeDeploymentRecovery -RunId $RunId -LogRoot $LogRoot
    $_mayHaveChanged = -not $_preview
    $_result = Resume-OfficeInstallation -Recovery $_recovery -ProductKey $ProductKey @_execution
    $_exitCode = [int]$_result.WrapperExitCode
  }
  else {
    $_configuration = $null

    if ($TargetProductId) {
      $_configurationArgs = @{ TargetProductId = $TargetProductId }

      foreach ($_name in ($_configurationNames | Where-Object { $_ -notin @('TargetProductId', 'ExcludePublisher') })) {
        if ($PSBoundParameters.ContainsKey($_name)) {
          $_configurationArgs[$_name] = $PSBoundParameters[$_name]
        }
      }

      if ($ExcludePublisher) {
        $_configurationArgs.ExcludeApp = @($ExcludeApp) + @('Publisher')
      }

      $_phase = 'Configure'
      $_configuration = New-OfficeDeploymentConfiguration @_configurationArgs
    }

    if ($Mode -eq 'Prepare') {
      # Preparation has a media-assessment contract, not an execution result.
      $_phase = 'Prepare'
      $_existing = Test-Path -LiteralPath $SourcePath
      $_prepareArgs = @{
        Configuration = $_configuration
        SourcePath    = $SourcePath
        OdtPath       = $OdtPath
        DryRun        = [bool]$DryRun
        WhatIf        = $_preview
        ErrorAction   = 'Stop'
      }

      if ($PSBoundParameters.ContainsKey('Confirm')) {
        $_prepareArgs.Confirm = [bool]$PSBoundParameters['Confirm']
      }

      if ($PSBoundParameters.ContainsKey('Verbose')) {
        $_prepareArgs.Verbose = [bool]$PSBoundParameters['Verbose']
      }

      $_mayHaveChanged = -not $_preview
      $_media = Save-OfficeDeploymentMedia @_prepareArgs
      $_status = 'Completed'
      $_reason = 'MediaPrepared'
      $_changed = -not $_existing

      if ($_media.Status -eq 'Preview') {
        $_status = 'Preview'
        $_reason = 'NotExecuted'
        $_changed = $false
      }
      elseif (-not $_media.Valid) {
        throw 'PSFoundation did not return verified media.'
      }
      elseif ($_existing) {
        $_reason = 'AlreadyPrepared'
      }

      $_property = @{
        SchemaVersion   = 1
        Phase           = 'Prepare'
        ReasonCode      = $_reason
        Configuration   = $_configuration
        Media           = $_media
        Changed         = $_changed
        ChangeKnown     = $true
        RebootRequired  = $false
        WrapperExitCode = 0
      }

      $_result = New-OperationResult `
        -Target $SourcePath `
        -Source Office `
        -Action Prepare `
        -Status $_status `
        -Property $_property
    }
    else {
      $_phase = 'Plan'
      $_inventory = Get-OfficeInventory
      $_plan = $null
      $_activation = $null

      if ($_configuration) {
        $_planArgs = @{
          Action        = 'Install'
          Configuration = $_configuration
          Inventory     = $_inventory
        }

        if ($SourcePath) {
          $_planArgs.SourcePath = $SourcePath
        }

        $_plan = Get-OfficeDeploymentPlan @_planArgs
      }

      if ($Mode -eq 'Check') {
        $_status = 'Completed'
        $_reason = 'InventoryRead'

        if ($_plan) {
          foreach ($_warning in $_plan.Warnings) {
            Write-Warning $_warning
          }

          $_activation = Get-OfficeActivationStatus -TargetProductId $TargetProductId
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
          Activation      = $_activation
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
        # The module revalidates inventory/media and owns the confirmation boundary.
        $_phase = 'Install'
        $_mayHaveChanged = -not $_preview
        $_result = Install-Office -Plan $_plan -LogRoot $LogRoot -ProductKey $ProductKey @_execution
        $_exitCode = [int]$_result.WrapperExitCode
      }
    }
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
  # The module owns staging, journal persistence, native execution, and cleanup.
  $ProductKey = $null
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
