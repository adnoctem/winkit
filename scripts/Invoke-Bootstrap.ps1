Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Bootstraps a fresh Windows machine with winkit prerequisites and PATH
  registration.

.DESCRIPTION
  Runs the bootstrap installers in dependency order â€" winget first, then
  PowerShell 7, then VC++ redistributables â€" then adds the bin/ directory
  to the system PATH so the .cmd launchers (optimizer, server-optimizer,
  ir.cmd) can be invoked by name from anywhere.

  Requires administrator elevation. Intended as the first command run on
  a newly imaged machine.

.PARAMETER VCRedistVersions
  VC++ redistributable versions to install. Defaults to 14.0 (the unified
  2015-2022 runtime modern software needs). Pass multiple to install several.

.PARAMETER SkipPath
  Don't modify the system PATH. Useful for re-runs or when PATH is managed
  elsewhere.

.PARAMETER BinDirectory
  Path to the bin/ directory to add to PATH. Defaults to the bin/ folder
  resolved relative to this script.

.PARAMETER DryRun
  Preview which steps would run without executing them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Invoke-Bootstrap.ps1
  Full bootstrap: winget, PowerShell 7, VC++ 14.0, and PATH registration.

.EXAMPLE
  PS> ./Invoke-Bootstrap.ps1 -VCRedistVersions 14.0,12.0 -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification = 'VCRedistVersions is a VC++ runtime version array, not a credential.')]
[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string[]]
  $VCRedistVersions = @('14.0'),

  [Parameter(Mandatory = $false)]
  [switch]
  $SkipPath,

  [Parameter(Mandatory = $false)]
  [string]
  $BinDirectory = (Join-Path $PSScriptRoot '..\bin'),

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

# -----------------------------------------------------------------------------

if (-not (Test-Elevation)) {
  Request-AdministratorPrivilege `
    -BoundParameters    $PSBoundParameters `
    -ArgumentList       $args `
    -IsElevatedRelaunch:$Elevated
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# ---- Script-local: Add-SystemPath -------------------------------------------
# CHANGE-NOTE: if winkit exposes a canonical Add-SystemPath function, prefer
# that and replace this local helper.
function Add-SystemPath {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param ([Parameter(Mandatory = $true)][string]$Directory)

  $resolved = (Resolve-Path -LiteralPath $Directory -ErrorAction Stop).Path.TrimEnd('\')
  $key = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
  $current = (Get-ItemProperty -Path $key -Name 'Path').Path
  $entries = $current -split ';' | Where-Object { $_ -ne '' }

  if ($entries -contains $resolved) {
    Write-Log -Message "  PATH already contains '$resolved'; no change." -Color Gray
    return $true
  }

  if ($PSCmdlet.ShouldProcess("system PATH", "append '$resolved'")) {
    $new = ($entries + $resolved) -join ';'
    Set-ItemProperty -Path $key -Name 'Path' -Value $new -Type ExpandString
    $env:PATH = "$env:PATH;$resolved"
    Write-Log -Message "  Added '$resolved' to system PATH." -Color Green

    try {
      if (-not ('Win32.NativeMethods' -as [type])) {
        Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
public static extern System.IntPtr SendMessageTimeout(System.IntPtr hWnd, uint Msg, System.UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out System.UIntPtr lpdwResult);
'@
      }
      $HWND_BROADCAST = [System.IntPtr]0xffff
      $WM_SETTINGCHANGE = 0x1a
      $result = [System.UIntPtr]::Zero
      [Win32.NativeMethods]::SendMessageTimeout($HWND_BROADCAST, $WM_SETTINGCHANGE, [System.UIntPtr]::Zero, 'Environment', 2, 5000, [ref]$result) | Out-Null
    }
    catch {
      Write-Log -Message "  PATH set, but environment broadcast failed (new shells will still pick it up): $($_.Exception.Message)" -Color Yellow
    }
    return $true
  }
  return $false
}

# ---- Step runner ------------------------------------------------------------
function Invoke-BootstrapStep {
  param([string]$Name, [string]$ScriptName, [hashtable]$Arguments = @{})
  $path = Join-Path $PSScriptRoot $ScriptName
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
    Write-Log -Message "  Installer '$ScriptName' not found at $path - skipping." -Color Yellow
    Add-OperationResult -Results $_results -Target $Name -Source 'Bootstrap' -Action 'Install' -Status 'Skipped' -SkippedReason 'ScriptNotFound'
    return $false
  }

  Write-Log -Message "==> $Name" -Color Cyan
  try {
    $childArgs = $Arguments.Clone()
    if ($WhatIfPreference) { $childArgs['WhatIf'] = $true }
    if ($VerbosePreference -eq 'Continue') { $childArgs['Verbose'] = $true }
    if ($DryRun) { $childArgs['DryRun'] = $true }
    & $path @childArgs
    Write-Log -Message "    $Name complete." -Color Green
    return $true
  }
  catch {
    Write-Log -Message "    $Name FAILED: $($_.Exception.Message)" -Color Red
    return $false
  }
}

# ---- Execution â€" dependency order matters -----------------------------------
Write-Log -Message 'Starting winkit bootstrap.' -Color Cyan

$_stepResults = @{}

# 1. winget â€" foundation for everything else.
$_stepResults['winget'] = Invoke-BootstrapStep -Name 'Install winget' -ScriptName 'Install-WinGet.ps1'

# 2. PowerShell 7 â€" depends on winget.
if ($_stepResults['winget'] -or $DryRun) {
  $_stepResults['pwsh'] = Invoke-BootstrapStep -Name 'Install PowerShell 7' -ScriptName 'Install-PowerShellCore.ps1'
}
else {
  Write-Log -Message 'Skipping PowerShell 7 â€" winget step did not succeed.' -Color Yellow
  $_stepResults['pwsh'] = $false
}

# 3. VC++ redistributables â€" independent of the above.
foreach ($v in $VCRedistVersions) {
  $_stepResults["vcredist-$v"] = Invoke-BootstrapStep -Name "Install VC++ $v" -ScriptName 'Install-VCRedistributables.ps1' -Arguments @{ Version = $v }
}

# 4. PATH registration.
if (-not $SkipPath) {
  Write-Log -Message '==> Register bin/ on system PATH' -Color Cyan
  try {
    $_pathResult = Add-SystemPath -Directory $BinDirectory
    $_stepResults['path'] = $_pathResult
    Add-OperationResult -Results $_results -Target 'SystemPath' -Source 'Bootstrap' -Action 'SetPath' -Status $(if ($_pathResult) { 'Completed' } else { 'Skipped' }) -Detail $BinDirectory
  }
  catch {
    Write-Log -Message "    PATH registration FAILED: $($_.Exception.Message)" -Color Red
    $_stepResults['path'] = $false
    Add-OperationResult -Results $_results -Target 'SystemPath' -Source 'Bootstrap' -Action 'SetPath' -Status 'Failed' -Detail $_.Exception.Message
  }
}

# ---- Summary -----------------------------------------------------------------
Write-Log -Message "`n=== Bootstrap summary ===" -Color Cyan
foreach ($step in $_stepResults.Keys) {
  $ok = $_stepResults[$step]
  $status = if ($ok) { 'OK' } else { 'FAIL' }
  $color = if ($ok) { 'Green' } else { 'Red' }
  Write-Log -Message "  [$status] $step" -Color $color
}

$failed = @($_stepResults.Keys | Where-Object { -not $_stepResults[$_] })
if ($failed.Count -gt 0) {
  Write-Log -Message "`n$($failed.Count) step(s) failed: $($failed -join ', ')" -Color Red
  $global:LASTEXITCODE = 1
}
else {
  Write-Log -Message "`nBootstrap complete. Open a new terminal for PATH changes to take effect." -Color Green
  $global:LASTEXITCODE = 0
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Invoke-Bootstrap'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

exit $LASTEXITCODE
