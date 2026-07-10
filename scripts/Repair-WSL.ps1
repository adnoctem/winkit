Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Repairs common WSL networking issues by resetting Windows network state.
.DESCRIPTION
  Runs the WSL shutdown and Windows network reset commands commonly used to
  recover WSL internet connectivity after reboot or adapter changes:

  - wsl --shutdown
  - netsh winsock reset
  - netsh int ip reset all
  - netsh winhttp reset proxy
  - ipconfig /flushdns

  The network reset may require a restart before all changes take effect. Use
  -Reboot to restart immediately after the repair commands complete.

  Use -JetBrainsIDEs to also apply the optional WSL/JetBrains IDE performance
  workaround: add the WSL firewall allow rule, disable Public-profile JetBrains
  firewall rules, and add Defender exclusions for JetBrains, Docker, WSL, and
  virtual disk files.

  Requires administrator elevation.
.PARAMETER JetBrainsIDEs
  Apply optional Windows Defender and firewall optimizations for JetBrains IDEs
  working with WSL projects.
.PARAMETER Distro
  WSL distro name used to build optional \\wsl$ and \\wsl.localhost source-path
  Defender exclusions. When omitted, the first listed WSL distro is used.
.PARAMETER LinuxUsername
  Linux username used to build optional WSL source-path Defender exclusions.
  When omitted, the script tries to detect the user with wsl.exe.
.PARAMETER Reboot
  Restart Windows immediately after the repair commands complete successfully.
.PARAMETER DryRun
  Preview the repair commands without executing them.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> .\Repair-WSL.ps1
.EXAMPLE
  PS> .\Repair-WSL.ps1 -DryRun
.EXAMPLE
  PS> .\Repair-WSL.ps1 -Reboot
.EXAMPLE
  PS> .\Repair-WSL.ps1 -JetBrainsIDEs -Distro Ubuntu -LinuxUsername markus -DryRun
.LINK
  https://github.com/microsoft/WSL/issues/3438#issuecomment-410518578
.LINK
  https://github.com/microsoft/WSL/issues/8995
.LINK
  https://gist.github.com/dkorobtsov/963f3b90418e51d12aecb1eaf6106958
.LINK
  https://www.jetbrains.com/help/idea/how-to-use-wsl-development-environment-in-product.html#debugging_system_settings
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT

  Workaround reference:
  microsoft/WSL#3438 - "No internet access after reboot"
  Originally traced to @avxkim / WSL issue #3438 and later repeated in
  multiple WSL networking duplicates, including microsoft/WSL#4275.

  JetBrains IDEs WSL performance workaround reference:
  JetBrains themselves recommend adding the WSL firewall allow rule, disabling
  Public-profile JetBrains firewall rules, and adding Defender exclusions for
  JetBrains, Docker, WSL, and virtual disk files. See the JetBrains help article
  for details: https://www.jetbrains.com/help/
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Parameter(Mandatory = $false)]
  [switch]
  $JetBrainsIDEs,

  [Parameter(Mandatory = $false)]
  [string]
  $Distro,

  [Parameter(Mandatory = $false)]
  [string]
  $LinuxUsername,

  [Parameter(Mandatory = $false)]
  [switch]
  $Reboot,

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

if (-not $DryRun -and -not (Test-Elevation)) {
  Request-AdministratorPrivilege `
    -BoundParameters    $PSBoundParameters `
    -ArgumentList       $args `
    -IsElevatedRelaunch:$Elevated
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no WSL repair or JetBrains optimization changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_repairSteps = @(
  [PSCustomObject]@{
    Target = 'WSL'
    Action = 'Shutdown'
    FilePath = 'wsl.exe'
    ArgumentList = @('--shutdown')
    Detail = 'wsl --shutdown'
  }
  [PSCustomObject]@{
    Target = 'Winsock'
    Action = 'Reset'
    FilePath = 'netsh.exe'
    ArgumentList = @('winsock', 'reset')
    Detail = 'netsh winsock reset'
  }
  [PSCustomObject]@{
    Target = 'TCPIP'
    Action = 'Reset'
    FilePath = 'netsh.exe'
    ArgumentList = @('int', 'ip', 'reset', 'all')
    Detail = 'netsh int ip reset all'
  }
  [PSCustomObject]@{
    Target = 'WinHTTPProxy'
    Action = 'Reset'
    FilePath = 'netsh.exe'
    ArgumentList = @('winhttp', 'reset', 'proxy')
    Detail = 'netsh winhttp reset proxy'
  }
  [PSCustomObject]@{
    Target = 'DNSResolverCache'
    Action = 'Flush'
    FilePath = 'ipconfig.exe'
    ArgumentList = @('/flushdns')
    Detail = 'ipconfig /flushdns'
  }
)

function Invoke-WSLRepairCommand {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Step,

    [Parameter(Mandatory = $true)]
    [System.Collections.IList]
    $Results
  )

  if ($WhatIfPreference) {
    Write-Log -Message "[DRY RUN] Would run: $($Step.Detail)" -Color Yellow
    Add-OperationResult -Results $Results -Target $Step.Target -Source 'WSLRepair' -Action $Step.Action -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($Step.Target, $Step.Detail)) {
    Add-OperationResult -Results $Results -Target $Step.Target -Source 'WSLRepair' -Action $Step.Action -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  $_command = Get-Command -Name $Step.FilePath -ErrorAction SilentlyContinue
  if (-not $_command) {
    Write-Log -Message "    SKIPPED - command not found: $($Step.FilePath)" -Color Yellow
    Add-OperationResult -Results $Results -Target $Step.Target -Source 'WSLRepair' -Action $Step.Action -Status 'Skipped' -Detail "Command not found: $($Step.FilePath)"
    return
  }

  Write-Log -Message "==> $($Step.Detail)" -Color Cyan
  $_output = & $_command.Source @($Step.ArgumentList) 2>&1
  $_exitCode = if ($null -ne $global:LASTEXITCODE) { $global:LASTEXITCODE } else { 0 }

  if ($_exitCode -eq 0) {
    Write-Log -Message '    completed.' -Color Green
    Add-OperationResult -Results $Results -Target $Step.Target -Source 'WSLRepair' -Action $Step.Action -Status 'Completed' -Detail $Step.Detail
    return
  }

  $_detail = @($_output) -join "`n"
  if ([string]::IsNullOrWhiteSpace($_detail)) {
    $_detail = "Exit code $_exitCode"
  }

  Write-Log -Message "    FAILED - $($Step.Detail) exited with $_exitCode" -Color Red
  Add-OperationResult -Results $Results -Target $Step.Target -Source 'WSLRepair' -Action $Step.Action -Status 'Failed' -Detail $_detail
}

function Resolve-WSLDistro {
  param (
    [string]
    $PreferredDistro
  )

  if (-not [string]::IsNullOrWhiteSpace($PreferredDistro)) {
    return $PreferredDistro
  }

  $_command = Get-Command -Name 'wsl.exe' -ErrorAction SilentlyContinue
  if (-not $_command) { return $null }

  try {
    $_distros = & $_command.Source --list --quiet 2>$null
    if ($LASTEXITCODE -ne 0) { return $null }

    foreach ($_distro in @($_distros)) {
      $_name = ([string]$_distro).Replace("`0", '').Trim()
      if (-not [string]::IsNullOrWhiteSpace($_name)) {
        return $_name
      }
    }
  }
  catch {
    return $null
  }

  return $null
}

function Resolve-WSLLinuxUsername {
  param (
    [string]
    $PreferredLinuxUsername,

    [string]
    $ResolvedDistro
  )

  if (-not [string]::IsNullOrWhiteSpace($PreferredLinuxUsername)) {
    return $PreferredLinuxUsername
  }

  if ([string]::IsNullOrWhiteSpace($ResolvedDistro)) { return $null }

  $_command = Get-Command -Name 'wsl.exe' -ErrorAction SilentlyContinue
  if (-not $_command) { return $null }

  try {
    $_username = & $_command.Source -d $ResolvedDistro --exec sh -lc 'whoami' 2>$null
    if ($LASTEXITCODE -ne 0) { return $null }
    return ([string]@($_username)[0]).Replace("`0", '').Trim()
  }
  catch {
    return $null
  }
}

function Add-JetBrainsWSLOptimization {
  [CmdletBinding()]
  param (
    [System.Collections.IList]
    $Results,

    [string]
    $PreferredDistro,

    [string]
    $PreferredLinuxUsername
  )

  Write-Log -Message '==> JetBrains IDE / WSL performance optimization' -Color Cyan

  [void]$Results.Add((Enable-WSLFirewallRule -WhatIf:$WhatIfPreference))
  foreach ($_result in @(Disable-JetBrainsFirewallRule -WhatIf:$WhatIfPreference)) {
    [void]$Results.Add($_result)
  }

  $_distro = Resolve-WSLDistro -PreferredDistro $PreferredDistro
  $_linuxUsername = Resolve-WSLLinuxUsername -PreferredLinuxUsername $PreferredLinuxUsername -ResolvedDistro $_distro

  $_paths = New-Object System.Collections.Generic.List[string]
  if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    $_paths.Add((Join-Path -Path $env:LOCALAPPDATA -ChildPath 'JetBrains'))
  }
  $_paths.Add('C:\Program Files\Docker')
  $_paths.Add('C:\Program Files\JetBrains')

  if (-not [string]::IsNullOrWhiteSpace($_distro) -and -not [string]::IsNullOrWhiteSpace($_linuxUsername)) {
    $_paths.Add("\\wsl$\$_distro\home\$_linuxUsername\src")
    $_paths.Add("\\wsl.localhost\$_distro\home\$_linuxUsername\src")
  }
  else {
    Add-OperationResult -Results $Results -Target 'WSLSourcePaths' -Source 'Defender' -Action 'AddExclusion' -Status 'Skipped' -Detail 'DistroOrLinuxUsernameUnavailable'
  }

  foreach ($_path in $_paths) {
    [void]$Results.Add((Add-DefenderExclusion -Type Path -Value $_path -WhatIf:$WhatIfPreference))
  }

  foreach ($_extension in @('vhd', 'vhdx')) {
    [void]$Results.Add((Add-DefenderExclusion -Type Extension -Value $_extension -WhatIf:$WhatIfPreference))
  }

  $_processes = @(
    'phpstorm64.exe',
    'idea64.exe',
    'pycharm64.exe',
    'rubymine64.exe',
    'webstorm64.exe',
    'datagrip64.exe',
    'goland64.exe',
    'rider64.exe',
    'fsnotifier.exe',
    'jcef_helper.exe',
    'jetbrains-toolbox.exe',
    'docker.exe',
    'com.docker.*.*',
    'Desktop Docker.exe',
    'wsl.exe',
    'wslhost.exe',
    'vmmemWSL'
  )

  foreach ($_process in $_processes) {
    [void]$Results.Add((Add-DefenderExclusion -Type Process -Value $_process -WhatIf:$WhatIfPreference))
  }
}

foreach ($_step in $_repairSteps) {
  Invoke-WSLRepairCommand -Step $_step -Results $_results -WhatIf:$WhatIfPreference
}

if ($JetBrainsIDEs) {
  Add-JetBrainsWSLOptimization -Results $_results -PreferredDistro $Distro -PreferredLinuxUsername $LinuxUsername
}

$_failedCount = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_completedCount = @($_results | Where-Object { $_.Status -eq 'Completed' }).Count
$_skippedCount = @($_results | Where-Object { $_.Status -ne 'Completed' -and $_.Status -ne 'Failed' }).Count

if ($Reboot) {
  if ($_failedCount -gt 0) {
    Write-Log -Message 'Skipping reboot because one or more WSL repair commands failed.' -Color Yellow
    Add-OperationResult -Results $_results -Target 'Computer' -Source 'Shutdown' -Action 'Restart' -Status 'Skipped' -Detail 'RepairFailed'
  }
  elseif ($DryRun) {
    Write-Log -Message '[DRY RUN] Would restart Windows immediately.' -Color Yellow
    Add-OperationResult -Results $_results -Target 'Computer' -Source 'Shutdown' -Action 'Restart' -Status 'Skipped' -Detail 'DryRun'
  }
  elseif ($PSCmdlet.ShouldProcess('Computer', 'Restart immediately with shutdown /r /f /t 0')) {
    Add-OperationResult -Results $_results -Target 'Computer' -Source 'Shutdown' -Action 'Restart' -Status 'Completed' -Detail 'shutdown /r /f /t 0'
    $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Repair-WSL'
    if ($_operationLog) {
      Write-Log -Message "Operation log: $_operationLog" -Color Gray
    }

    Write-Log -Message 'Restarting Windows now...' -Color Yellow
    shutdown.exe /r /f /t 0
    return
  }
  else {
    Add-OperationResult -Results $_results -Target 'Computer' -Source 'Shutdown' -Action 'Restart' -Status 'Skipped' -Detail 'WhatIf'
  }
}
else {
  Write-Log -Message 'A reboot may be required for the network reset to fully apply. Re-run with -Reboot to restart automatically.' -Color Gray
}

if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no WSL repair or JetBrains optimization changes were applied" -Color Yellow
}
else {
  Write-Log -Message "`nWSL repair: $_completedCount completed | $_skippedCount skipped | $_failedCount failed" -Color $(if ($_failedCount -gt 0) { 'Yellow' } else { 'Green' })
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Repair-WSL'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failedCount -gt 0) {
  exit 1
}
