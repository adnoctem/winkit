#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Orchestrates Windows configuration scripts according to a named profile.

.DESCRIPTION
  Single entry point for running winkit configuration scripts as a set.
  Rather than maintaining parallel launcher scripts per machine type, this
  orchestrator holds a manifest mapping profiles (Desktop / Server / DC /
  Minimal) to the ordered list of config scripts that profile should run.

  Execution is sequential. Machine configuration has implicit ordering and
  shared-resource contention (registry, AppX service, DISM, default-user
  hive mounting) that makes blanket parallelism unsafe. Each step is run
  in isolation so one failure does not abort the whole run unless -StopOnError
  is specified.

  resolved from the scripts/ directory.

.PARAMETER Profile
  Which configuration set to run. If omitted, the profile is auto-detected
  from the OS product type (workstation -> Desktop, server -> Server,
  domain controller -> DC).

.PARAMETER StopOnError
  Abort the entire run if any single step fails. Default is to log the
  failure, continue, and report a summary at the end.

.PARAMETER ListOnly
  Print the resolved script list for the selected profile and exit without
  running anything. Useful for verifying profile membership.

.PARAMETER DryRun
  Preview which scripts would run without executing them.

.PARAMETER PassThru
  Return structured operation results for each step.

.EXAMPLE
  PS> ./Invoke-Optimizer.ps1
  Auto-detects the profile and runs the appropriate set.

.EXAMPLE
  PS> ./Invoke-Optimizer.ps1 -Profile Desktop -DryRun
  Shows what the Desktop profile would run without making changes.

.EXAMPLE
  PS> ./Invoke-Optimizer.ps1 -Profile Desktop -ListOnly
  Lists the Desktop profile scripts in execution order and exits.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [ValidateSet('Desktop', 'Server', 'DC', 'Minimal')]
  [string]
  $Profile,

  [Parameter(Mandatory = $false)]
  [switch]
  $StopOnError,

  [Parameter(Mandatory = $false)]
  [switch]
  $ListOnly,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru,

  # Internal: set automatically on elevated re-launch. Not for direct use.
  [switch]
  $Elevated
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if (-not (Test-Elevation)) {
  Request-AdministratorPrivilege `
    -BoundParameters    $PSBoundParameters `
    -ArgumentList       $args `
    -IsElevatedRelaunch:$Elevated
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no configuration scripts will be executed`n" -Color Yellow
}

# ---- Profile manifest --------------------------------------------------------
# $Common scripts are applied to every non-Minimal profile as a shared baseline.
# Adding a new config script means adding its base name here under the right
# profiles - not editing launcher files.
$Common = @(
  'Disable-DiagnosticTracking'
  'Configure-Updates'
)

$profileSets = @{
  Minimal = $Common
  Desktop = $Common + @(
    'Disable-ContentDelivery'
    'Set-AppPermissionDefaults'
    'Configure-BrowserPolicies'
    'Configure-Privacy'
    'Configure-AI'
    'Configure-System'
    'Configure-StartMenu'
    'Configure-Taskbar'
    'Configure-Explorer'
    'Disable-GameDVR'
    'Disable-PointerAcceleration'
    'Remove-Bloatware'
    'Remove-OneDrive'
    'Set-TerminalExperienceDefaults'
  )
  Server = $Common + @(
    'Disable-ContentDelivery'
    'Set-AppPermissionDefaults'
    'Configure-BrowserPolicies'
    'Configure-System'
  )
  DC = $Common
}

# ---- Profile resolution / auto-detection -------------------------------------
$_resolvedProfile = $Profile
if (-not $_resolvedProfile) {
  $_productType = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).ProductType
  $_resolvedProfile = switch ($_productType) {
    1 { 'Desktop' }
    2 { 'DC' }
    3 { 'Server' }
    default { 'Minimal' }
  }
  Write-Log -Message "Auto-detected profile '$_resolvedProfile' from OS ProductType $_productType." -Color Gray
}

$_scriptNames = $profileSets[$_resolvedProfile]
if (-not $_scriptNames -or $_scriptNames.Count -eq 0) {
  Write-Log -Message "Profile '$_resolvedProfile' has no scripts defined. Nothing to do." -Color Yellow
  exit 0
}

# ---- De-duplicate while preserving order -------------------------------------
$_seen = New-Object System.Collections.Generic.HashSet[string]
$_ordered = foreach ($_n in $_scriptNames) { if ($_seen.Add($_n)) { $_n } }

$_scriptsDir = $PSScriptRoot
$_resolved = New-Object System.Collections.ArrayList
foreach ($_name in $_ordered) {
  $_path = Join-Path $_scriptsDir "$_name.ps1"
  if (-not (Test-Path -LiteralPath $_path -PathType Leaf)) {
    Write-Log -Message "Script '$_name' listed in profile '$_resolvedProfile' not found at $_path - skipping." -Color Yellow
    continue
  }
  [void]$_resolved.Add([PSCustomObject]@{
      Name = $_name
      Path = $_path
    })
}

if ($_resolved.Count -eq 0) {
  Write-Log -Message "No resolved scripts found for profile '$_resolvedProfile'. Nothing to do." -Color Yellow
  exit 0
}

# ---- List only ---------------------------------------------------------------
if ($ListOnly) {
  Write-Log -Message "Profile '$_resolvedProfile' runs these scripts in order:" -Color Cyan
  foreach ($_script in $_resolved) {
    Write-Log -Message "  - $($_script.Name)" -Color Gray
  }
  exit 0
}

# ---- Execution ---------------------------------------------------------------
$_results = New-Object System.Collections.ArrayList
$_stepNumber = 0

foreach ($_script in $_resolved) {
  $_stepNumber++
  $_label = "[$_stepNumber/$($_resolved.Count)] $($_script.Name)"

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would execute: $_label" -Color Yellow
    Add-OperationResult -Results $_results -Target $_script.Name -Source 'Optimizer' -Action 'Run' -Status 'Skipped' -Detail 'DryRun'
    continue
  }

  Write-Log -Message "==> $_label" -Color Cyan
  $_stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  try {
    $_childParams = @{}
    foreach ($_key in $PSBoundParameters.Keys) {
      if ($_key -notin @('Profile', 'StopOnError', 'ListOnly', 'DryRun', 'PassThru')) {
        $_childParams[$_key] = $PSBoundParameters[$_key]
      }
    }
    if ($DryRun -and -not $_childParams.ContainsKey('WhatIf')) {
      $_childParams['WhatIf'] = $true
    }
    $_childParams['ErrorAction'] = 'Stop'

    & $_script.Path @_childParams

    $_stopwatch.Stop()
    Add-OperationResult -Results $_results -Target $_script.Name -Source 'Optimizer' -Action 'Run' -Status 'Completed' -Detail "$($_script.Name) completed in $([math]::Round($_stopwatch.Elapsed.TotalSeconds, 1))s"
    Write-Log -Message "    done ($([math]::Round($_stopwatch.Elapsed.TotalSeconds, 1))s)" -Color Green
  }
  catch {
    $_stopwatch.Stop()
    Add-OperationResult -Results $_results -Target $_script.Name -Source 'Optimizer' -Action 'Run' -Status 'Failed' -Detail $_.Exception.Message
    Write-Log -Message "    FAILED: $($_script.Name)" -Color Red

    if ($StopOnError) {
      Write-Log -Message "Stopping - -StopOnError is set and step '$($_script.Name)' failed." -Color Red
      break
    }
  }
}

# ---- Summary -----------------------------------------------------------------
Write-Log -Message "`n=== Optimizer summary (profile: $_resolvedProfile) ===" -Color Cyan

foreach ($_result in $_results) {
  $_status = if ($_result.Status -eq 'Completed') { 'OK' } else { if ($_result.Status -eq 'Failed') { 'FAIL' } else { 'SKIP' } }
  $_color = if ($_result.Status -eq 'Completed') { 'Green' } elseif ($_result.Status -eq 'Failed') { 'Red' } else { 'Gray' }
  Write-Log -Message "  [$_status] $($_result.Target)" -Color $_color
}

$_failedCount = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_completedCount = @($_results | Where-Object { $_.Status -eq 'Completed' }).Count
$_skippedCount = @($_results | Where-Object { $_.Status -ne 'Completed' -and $_.Status -ne 'Failed' }).Count

if ($_failedCount -gt 0) {
  Write-Log -Message "`n$_completedCount succeeded | $_skippedCount skipped | $_failedCount failed" -Color Yellow
}
else {
  Write-Log -Message "`n$_completedCount succeeded | $_skippedCount skipped" -Color Green
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Invoke-Optimizer'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failedCount -gt 0) {
  $global:LASTEXITCODE = 1
}
else {
  $global:LASTEXITCODE = 0
}
exit $LASTEXITCODE
