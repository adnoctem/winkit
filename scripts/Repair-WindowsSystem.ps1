#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Runs Windows repair steps as independently-selectable operations.

.DESCRIPTION
  "When Windows is acting up, run this" - a modular repair flow derived from
  Repair-WindowsSystem.ps1 in LeDragoX/Win-Debloat-Tools (itself attributed
  to ChrisTitusTech/win10script), split into independently-selectable steps
  matching winkit's per-concern granularity.

  Steps (select via -Steps; defaults to all except the two opt-in steps):

    PowerPlan        - restore the default power schemes (powercfg).
    Store            - reset the Microsoft Store (wsreset).
    Taskbar          - re-register the ShellExperienceHost / StartMenu
                       packages for the current user.
    TestModeWatermark- disable testsigning (removes the "Test Mode"
                       watermark).
    BITS             - reset stuck Background Intelligent Transfer jobs.
    SFC              - sfc /scannow (REPAIR variant - deliberately distinct
                       from Test-SystemFileIntegrity.ps1, which only
                       verifies).
    DISM             - dism /Online /Cleanup-Image /RestoreHealth.
    Network          - release/renew, flush DNS, Winsock and IP stack reset.

  Opt-in steps (not run by default - see each step's warning):
    Hosts            - resets the hosts file to the stock Microsoft default.
                       REFUSES when the winkit-managed blocklist marker block
                       is present: use Set-HostsBlocklist.ps1 -Undo to remove
                       the blocklist deliberately, then re-run the step.
    AppX             - re-registers currently-installed AppX packages for the
                       current user. Does NOT reinstall removed packages, but
                       on a machine winkit has debloated this step is usually
                       unnecessary and can mask whether Remove-Bloatware.ps1's
                       work actually stuck - only use it to fix apps that are
                       present but fail to launch.

  SFC and DISM can take a long time; both render progress and are bounded by
  -TimeoutMinutes. Requires administrator elevation.

.PARAMETER Steps
  Repair step(s) to run. Defaults to PowerPlan, Store, Taskbar,
  TestModeWatermark, BITS, SFC, DISM, Network (Hosts and AppX are opt-in).

.PARAMETER TimeoutMinutes
  Watchdog ceiling for the SFC and DISM steps. Defaults to 60.

.PARAMETER DryRun
  Preview the steps without running them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Repair-WindowsSystem.ps1
  Runs the default repair steps.

.EXAMPLE
  PS> ./Repair-WindowsSystem.ps1 -Steps SFC, DISM, Network
  Runs only the three named steps.

.EXAMPLE
  PS> ./Repair-WindowsSystem.ps1 -Steps Hosts
  Resets the hosts file (refuses if the winkit blocklist marker is present).

.LINK
  https://github.com/adnoctem/winkit
  https://github.com/LeDragoX/Win-Debloat-Tools
  https://github.com/ChrisTitusTech/win10script

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Carriage-return progress lines require Write-Host for in-place console updates.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'TimeoutMinutes', Justification = 'Consumed by the step functions through script scope.')]

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Repair step(s) to run.'
  )]
  [ValidateSet('PowerPlan', 'Store', 'Taskbar', 'TestModeWatermark', 'BITS', 'SFC', 'DISM', 'Network', 'Hosts', 'AppX')]
  [string[]]
  $Steps = @('PowerPlan', 'Store', 'Taskbar', 'TestModeWatermark', 'BITS', 'SFC', 'DISM', 'Network'),

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 240)]
  [int]
  $TimeoutMinutes = 60,

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

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no repair steps will run`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

$hostsMarkerStart = '# ---- winkit managed blocklist ----'
$hostsFile = "$env:SystemRoot\System32\drivers\etc\hosts"

# ---- Step implementations ----------------------------------------------------

function Invoke-WaitForNativeProcess {
  <#
    Runs a native executable and polls until it exits, rendering progress.
    Returns the exit code, or $null on watchdog timeout.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath,

    [Parameter(Mandatory = $true)]
    [string[]]$ArgumentList,

    [Parameter(Mandatory = $true)]
    [string]$Activity,

    [int]$TimeoutMinutesValue = 60
  )

  $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -NoNewWindow
  $startedAt = Get-Date
  $deadline = $startedAt.AddMinutes($TimeoutMinutesValue)

  while (-not $process.HasExited) {
    if ((Get-Date) -gt $deadline) {
      return $null
    }
    $elapsedTime = (Get-Date) - $startedAt
    $elapsedSeconds = [int]$elapsedTime.TotalSeconds
    $pct = [int][math]::Min(($elapsedSeconds / ($TimeoutMinutesValue * 60)) * 100, 100)
    Write-Progress -Activity $Activity -Status "elapsed: $elapsedSeconds s / $TimeoutMinutesValue min" -PercentComplete $pct
    Start-Sleep -Seconds 5
  }
  Write-Progress -Activity $Activity -Completed

  $process.ExitCode
}

function Invoke-PowerPlanStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('PowerPlan', 'Restore default power schemes')) {
    Add-OperationResult -Results $Results -Target 'PowerPlan' -Source 'Repair' -Action 'PowerPlan' -Status 'Skipped' -Detail 'WhatIf'
    return
  }
  try {
    (& powercfg.exe /restoredefaultschemes) 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
      throw "powercfg exited with code $LASTEXITCODE."
    }
    Add-OperationResult -Results $Results -Target 'PowerPlan' -Source 'Repair' -Action 'PowerPlan' -Status 'Completed' -Detail 'Default power schemes restored.'
  }
  catch {
    Add-OperationResult -Results $Results -Target 'PowerPlan' -Source 'Repair' -Action 'PowerPlan' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-StoreStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('Store', 'Reset the Microsoft Store (wsreset)')) {
    Add-OperationResult -Results $Results -Target 'Store' -Source 'Repair' -Action 'Store' -Status 'Skipped' -Detail 'WhatIf'
    return
  }
  try {
    $null = Start-Process -FilePath "$env:SystemRoot\System32\wsreset.exe" -Wait -PassThru -ErrorAction Stop
    Add-OperationResult -Results $Results -Target 'Store' -Source 'Repair' -Action 'Store' -Status 'Completed' -Detail 'Microsoft Store reset triggered (may appear to do nothing - that is normal).'
  }
  catch {
    Add-OperationResult -Results $Results -Target 'Store' -Source 'Repair' -Action 'Store' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-TaskbarStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('Taskbar', 'Re-register shell packages')) {
    Add-OperationResult -Results $Results -Target 'Taskbar' -Source 'Repair' -Action 'Taskbar' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  $packages = @(Get-AppxPackage -Name 'Microsoft.Windows.ShellExperienceHost', 'Microsoft.Windows.StartMenuExperienceHost' -ErrorAction SilentlyContinue)
  if ($packages.Count -eq 0) {
    Add-OperationResult -Results $Results -Target 'Taskbar' -Source 'Repair' -Action 'Taskbar' -Status 'Skipped' -Detail 'No shell packages present to re-register.'
    return
  }

  $failed = 0
  foreach ($package in $packages) {
    try {
      $manifest = Join-Path -Path $package.InstallLocation -ChildPath 'AppXManifest.xml'
      if (Test-Path -LiteralPath $manifest) {
        $null = Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction Stop
        Add-OperationResult -Results $Results -Target $package.Name -Source 'Repair' -Action 'Taskbar' -Status 'Completed' -Detail 'Shell package re-registered.'
      }
      else {
        $failed++
        Add-OperationResult -Results $Results -Target $package.Name -Source 'Repair' -Action 'Taskbar' -Status 'Warn' -Detail 'AppXManifest.xml not found in install location.'
      }
    }
    catch {
      $failed++
      Add-OperationResult -Results $Results -Target $package.Name -Source 'Repair' -Action 'Taskbar' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  if ($failed -gt 0) {
    Write-Log -Message 'Taskbar step completed with failures - a sign-out/in may be required for the shell to fully reload.' -Color Yellow
  }
}

function Invoke-TestModeWatermarkStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('TestModeWatermark', 'Disable testsigning')) {
    Add-OperationResult -Results $Results -Target 'TestModeWatermark' -Source 'Repair' -Action 'TestModeWatermark' -Status 'Skipped' -Detail 'WhatIf'
    return
  }
  try {
    (& bcdedit.exe /set testsigning off) 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
      throw "bcdedit exited with code $LASTEXITCODE."
    }
    Add-OperationResult -Results $Results -Target 'TestModeWatermark' -Source 'Repair' -Action 'TestModeWatermark' -Status 'Completed' -Detail 'Testsigning disabled - the Test Mode watermark disappears after a reboot.'
  }
  catch {
    Add-OperationResult -Results $Results -Target 'TestModeWatermark' -Source 'Repair' -Action 'TestModeWatermark' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-BitsStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('BITS', 'Reset Background Intelligent Transfer jobs')) {
    Add-OperationResult -Results $Results -Target 'BITS' -Source 'Repair' -Action 'BITS' -Status 'Skipped' -Detail 'WhatIf'
    return
  }
  try {
    (& bitsadmin.exe /reset /allusers) 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
      throw "bitsadmin exited with code $LASTEXITCODE."
    }
    Add-OperationResult -Results $Results -Target 'BITS' -Source 'Repair' -Action 'BITS' -Status 'Completed' -Detail 'BITS transfer queue reset.'
  }
  catch {
    Add-OperationResult -Results $Results -Target 'BITS' -Source 'Repair' -Action 'BITS' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-SfcStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('SFC', 'Run sfc /scannow (repair)')) {
    Add-OperationResult -Results $Results -Target 'SFC' -Source 'Repair' -Action 'SFC' -Status 'Skipped' -Detail 'WhatIf'
    return
  }
  Write-Log -Message "Running sfc /scannow (watchdog $TimeoutMinutes minutes)..." -Color Cyan
  $exitCode = Invoke-WaitForNativeProcess -FilePath "$env:SystemRoot\System32\sfc.exe" -ArgumentList @('/scannow') -Activity 'sfc /scannow' -TimeoutMinutesValue $TimeoutMinutes

  if ($null -eq $exitCode) {
    Add-OperationResult -Results $Results -Target 'SFC' -Source 'Repair' -Action 'SFC' -Status 'Unknown' -Detail "sfc did not finish within $TimeoutMinutes minutes."
    return
  }

  $status = if ($exitCode -eq 0) { 'Completed' } else { 'Failed' }
  Add-OperationResult -Results $Results -Target 'SFC' -Source 'Repair' -Action 'SFC' -Status $status -Detail "sfc /scannow exit code $exitCode (0 = no violations found)."
}

function Invoke-DismStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('DISM', 'Run DISM /RestoreHealth')) {
    Add-OperationResult -Results $Results -Target 'DISM' -Source 'Repair' -Action 'DISM' -Status 'Skipped' -Detail 'WhatIf'
    return
  }
  Write-Log -Message "Running dism /Online /Cleanup-Image /RestoreHealth (watchdog $TimeoutMinutes minutes)..." -Color Cyan
  $exitCode = Invoke-WaitForNativeProcess -FilePath "$env:SystemRoot\System32\dism.exe" -ArgumentList @('/Online', '/Cleanup-Image', '/RestoreHealth') -Activity 'DISM /RestoreHealth' -TimeoutMinutesValue $TimeoutMinutes

  if ($null -eq $exitCode) {
    Add-OperationResult -Results $Results -Target 'DISM' -Source 'Repair' -Action 'DISM' -Status 'Unknown' -Detail "dism did not finish within $TimeoutMinutes minutes."
    return
  }

  $status = if ($exitCode -eq 0) { 'Completed' } else { 'Failed' }
  Add-OperationResult -Results $Results -Target 'DISM' -Source 'Repair' -Action 'DISM' -Status $status -Detail "dism /RestoreHealth exit code $exitCode."
}

function Invoke-NetworkStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('Network', 'Reset the network stack')) {
    Add-OperationResult -Results $Results -Target 'Network' -Source 'Repair' -Action 'Network' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  $commands = @(
    @{ Name = 'ipconfig /release'; Executable = 'ipconfig.exe'; Arguments = @('/release') },
    @{ Name = 'ipconfig /renew'; Executable = 'ipconfig.exe'; Arguments = @('/renew') },
    @{ Name = 'ipconfig /flushdns'; Executable = 'ipconfig.exe'; Arguments = @('/flushdns') },
    @{ Name = 'netsh winsock reset'; Executable = 'netsh.exe'; Arguments = @('winsock', 'reset') },
    @{ Name = 'netsh int ip reset'; Executable = 'netsh.exe'; Arguments = @('int', 'ip', 'reset') }
  )

  $failed = 0
  foreach ($command in $commands) {
    try {
      (& $command.Executable @($command.Arguments)) 2>&1 | Out-Null
      if ($LASTEXITCODE -ne 0) {
        throw "exit code $LASTEXITCODE."
      }
      Add-OperationResult -Results $Results -Target $command.Name -Source 'Repair' -Action 'Network' -Status 'Completed' -Detail 'Network command completed.'
    }
    catch {
      $failed++
      Add-OperationResult -Results $Results -Target $command.Name -Source 'Repair' -Action 'Network' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  if ($failed -gt 0) {
    Write-Log -Message 'Network step completed with failures - a reboot is recommended to fully apply the stack reset.' -Color Yellow
  }
}

function Invoke-HostsStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('Hosts', 'Reset the hosts file to the stock default')) {
    Add-OperationResult -Results $Results -Target 'Hosts' -Source 'Repair' -Action 'Hosts' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  if (-not (Test-Path -LiteralPath $hostsFile -PathType Leaf)) {
    Add-OperationResult -Results $Results -Target 'Hosts' -Source 'Repair' -Action 'Hosts' -Status 'Skipped' -Detail 'Hosts file not found.'
    return
  }

  $hostsContent = Get-Content -LiteralPath $hostsFile -Raw -ErrorAction SilentlyContinue
  if ($hostsContent -and ($hostsContent -match [regex]::Escape($hostsMarkerStart))) {
    Add-OperationResult -Results $Results -Target 'Hosts' -Source 'Repair' -Action 'Hosts' -Status 'Skipped' -Detail 'The hosts file contains the winkit-managed blocklist. Use Set-HostsBlocklist.ps1 -Undo to remove it deliberately, then re-run this step.'
    return
  }

  $stockHosts = @'
# Copyright (c) 1993-2009 Microsoft Corp.
#
# This is a sample HOSTS file used by Microsoft TCP/IP for Windows.
#
# This file contains the mappings of IP addresses to host names. Each
# entry should be kept on an individual line. The IP address should
# be placed in the first column followed by the corresponding host name.
# The IP address and the host name should be separated by at least one
# space.
#
# Additionally, comments (such as these) may be inserted on individual
# lines or following the machine name denoted by a '#' symbol.
#
# For example:
#
#      102.54.94.97     rhino.acme.com          # source server
#       38.25.63.10     x.acme.com              # x client host

# localhost name resolution is handled within DNS itself.
#	127.0.0.1       localhost
#	::1             localhost
'@

  try {
    Set-Content -LiteralPath $hostsFile -Value $stockHosts -Encoding Ascii -ErrorAction Stop
    Add-OperationResult -Results $Results -Target 'Hosts' -Source 'Repair' -Action 'Hosts' -Status 'Completed' -Detail 'Hosts file reset to the stock default.'
  }
  catch {
    Add-OperationResult -Results $Results -Target 'Hosts' -Source 'Repair' -Action 'Hosts' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-AppxStep {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([System.Collections.ArrayList]$Results)

  if (-not $PSCmdlet.ShouldProcess('AppX', 'Re-register installed AppX packages')) {
    Add-OperationResult -Results $Results -Target 'AppX' -Source 'Repair' -Action 'AppX' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  Write-Log -Message 'Note: AppX re-registration only touches packages that are currently installed. On a machine winkit has debloated this step is usually unnecessary - it does not reinstall removed packages, and can mask whether Remove-Bloatware.ps1''s work actually stuck.' -Color Yellow

  $packages = @(Get-AppxPackage -ErrorAction SilentlyContinue)
  if ($packages.Count -eq 0) {
    Add-OperationResult -Results $Results -Target 'AppX' -Source 'Repair' -Action 'AppX' -Status 'Skipped' -Detail 'No AppX packages present for the current user.'
    return
  }

  $failed = 0
  foreach ($package in $packages) {
    $manifest = Join-Path -Path $package.InstallLocation -ChildPath 'AppXManifest.xml'
    if (-not (Test-Path -LiteralPath $manifest)) {
      continue
    }
    try {
      $null = Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction Stop
      Add-OperationResult -Results $Results -Target $package.Name -Source 'Repair' -Action 'AppX' -Status 'Completed' -Detail 'Package re-registered.'
    }
    catch {
      $failed++
      Add-OperationResult -Results $Results -Target $package.Name -Source 'Repair' -Action 'AppX' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  if ($failed -gt 0) {
    Write-Log -Message 'AppX step completed with failures - some packages could not be re-registered.' -Color Yellow
  }
}

# ---- Main flow ---------------------------------------------------------------

Write-Log -Message "Running repair steps: $($Steps -join ', ')" -Color Cyan

foreach ($step in $Steps) {
  switch ($step) {
    'PowerPlan' { Invoke-PowerPlanStep -Results $_results }
    'Store' { Invoke-StoreStep -Results $_results }
    'Taskbar' { Invoke-TaskbarStep -Results $_results }
    'TestModeWatermark' { Invoke-TestModeWatermarkStep -Results $_results }
    'BITS' { Invoke-BitsStep -Results $_results }
    'SFC' { Invoke-SfcStep -Results $_results }
    'DISM' { Invoke-DismStep -Results $_results }
    'Network' { Invoke-NetworkStep -Results $_results }
    'Hosts' { Invoke-HostsStep -Results $_results }
    'AppX' { Invoke-AppxStep -Results $_results }
  }
}

$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Unknown' }).Count
$_skipped = @($_results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_completed = @($_results | Where-Object { $_.Status -notin @('Failed', 'Skipped', 'Unknown') }).Count
$_color = if ($_failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Log -Message "`nRepair workflow complete. Completed: $_completed | Skipped: $_skipped | Failed: $_failed" -Color $_color
$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Repair-WindowsSystem'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failed -gt 0) {
  exit 1
}
exit 0
