#Requires -Version 5.0

<#
.SYNOPSIS
  Collects incident-response evidence from a Windows endpoint.

.DESCRIPTION
  Single entry point for IR evidence collection. Runs a curated set of
  collection steps organised into profiles (Quick / Full). Each step
  gathers host information, event logs, Defender state, scheduled tasks,
  registry persistence, WMI subscriptions, and other forensic artefacts.

  Output is written to a timestamped case directory optionally compressed
   into a ZIP archive. Designed for offline triage on an infected endpoint -
   run from USB with administrator rights while the machine is network-isolated.

.PARAMETER Profile
  Which collection set to run. Quick is the default (initial triage, fast).
  Full adds deeper event logs, file timeline scans, and ZIP compression.

.PARAMETER OutputPath
  Root directory for the IR case folder. Defaults to the current working
  directory. A timestamped subfolder IR-<COMPUTER>-<TIMESTAMP> is created
  inside it.

.PARAMETER StopOnError
  Abort the collection if any single step fails. Default is to log the failure
  and continue.

.PARAMETER ListOnly
  Print the resolved step list for the selected profile and exit without
  collecting anything.

.PARAMETER DryRun
  Preview which steps would run without executing collection commands.

.PARAMETER PassThru
  Return structured operation results for each step.

.EXAMPLE
  PS> ./Invoke-InformationRetrieval.ps1
  Runs Quick profile IR collection in the current directory.

.EXAMPLE
  PS> ./Invoke-InformationRetrieval.ps1 -Profile Full -OutputPath 'D:\'
  Runs Full profile IR collection, writing the case folder to D:\.

.EXAMPLE
  PS> ./Invoke-InformationRetrieval.ps1 -Profile Quick -ListOnly
  Lists the Quick profile steps in execution order and exits.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [ValidateSet('Quick', 'Full')]
  [string]
  $Profile = 'Quick',

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath,

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

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if (-not (Test-Elevation)) {
  Request-AdministratorPrivilege `
    -BoundParameters    $PSBoundParameters `
    -ArgumentList       $args `
    -IsElevatedRelaunch:$Elevated
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no collection commands will be executed`n" -Color Yellow
}

# ---- Output directory setup --------------------------------------------------
$_outputRoot = if ($OutputPath) { $OutputPath } else { Get-Location }
$_computer = $env:COMPUTERNAME
$_stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$_caseDir = Join-Path $_outputRoot "IR-$_computer-$_stamp"
$_evtxDir = Join-Path $_caseDir 'EVTX'
$_textDir = Join-Path $_caseDir 'Text'
$_artifactsDir = Join-Path $_caseDir 'Artifacts'
$_registryDir = Join-Path $_artifactsDir 'Registry'

if (-not $DryRun) {
  foreach ($_dir in @($_caseDir, $_evtxDir, $_textDir, $_artifactsDir, $_registryDir)) {
    if (-not (Test-Path -LiteralPath $_dir)) {
      $null = New-Item -Path $_dir -ItemType Directory -Force -ErrorAction SilentlyContinue
    }
  }
}

# ---- Profile manifest --------------------------------------------------------
$Common = @(
  @{
    Name = 'Collection metadata'
    Script = {
      @"
CaseDir: $_caseDir
Computer: $_computer
Collection time local: $(Get-Date -Format o)
Collection time UTC: $((Get-Date).ToUniversalTime().ToString("o"))
Profile: Quick
"@ | Out-File -LiteralPath (Join-Path $_textDir 'Collection-Info.txt') -Encoding UTF8
    }
  }
  @{
    Name = 'Basic host information'
    Script = {
      $null = Invoke-SafeProcess -FilePath 'systeminfo.exe' -OutputPath (Join-Path $_textDir 'SystemInfo.txt')
      $null = Invoke-SafeProcess -FilePath 'whoami.exe' -ArgumentList @('/all') -OutputPath (Join-Path $_textDir 'WhoAmI-All.txt')
      $null = Invoke-SafeProcess -FilePath 'ipconfig.exe' -ArgumentList @('/all') -OutputPath (Join-Path $_textDir 'IPConfig.txt')
      $null = Invoke-SafeProcess -FilePath 'route.exe' -ArgumentList @('print') -OutputPath (Join-Path $_textDir 'RoutePrint.txt')
      $null = Invoke-SafeProcess -FilePath 'arp.exe' -ArgumentList @('-a') -OutputPath (Join-Path $_textDir 'ARP.txt')
    }
  }
  @{
    Name = 'Defender state and detections'
    Script = {
      Get-MpThreatDetection | Format-List * | Out-File -LiteralPath (Join-Path $_textDir 'Defender-ThreatDetections.txt') -Encoding UTF8
      Get-MpThreat | Format-List * | Out-File -LiteralPath (Join-Path $_textDir 'Defender-Threats.txt') -Encoding UTF8
      Get-MpPreference | Format-List * | Out-File -LiteralPath (Join-Path $_textDir 'Defender-Preferences.txt') -Encoding UTF8
      Get-MpComputerStatus | Format-List * | Out-File -LiteralPath (Join-Path $_textDir 'Defender-Status.txt') -Encoding UTF8
    }
  }
  @{
    Name = 'Export key event logs'
    Script = {
      Export-EventLog -LogName 'Microsoft-Windows-Windows Defender/Operational' -OutputPath (Join-Path $_evtxDir 'Microsoft-Windows-Windows_Defender_Operational.evtx') -MissingLogPath (Join-Path $_textDir 'Missing-EventLogs.txt')
      Export-EventLog -LogName 'Microsoft-Windows-PowerShell/Operational' -OutputPath (Join-Path $_evtxDir 'Microsoft-Windows-PowerShell_Operational.evtx') -MissingLogPath (Join-Path $_textDir 'Missing-EventLogs.txt')
      Export-EventLog -LogName 'Security' -OutputPath (Join-Path $_evtxDir 'Security.evtx') -MissingLogPath (Join-Path $_textDir 'Missing-EventLogs.txt')
      Export-EventLog -LogName 'System' -OutputPath (Join-Path $_evtxDir 'System.evtx') -MissingLogPath (Join-Path $_textDir 'Missing-EventLogs.txt')
    }
  }
  @{
    Name = 'Scheduled tasks'
    Script = {
      $null = Invoke-SafeProcess -FilePath 'schtasks.exe' -ArgumentList @('/query', '/fo', 'LIST', '/v') -OutputPath (Join-Path $_textDir 'ScheduledTasks-Full.txt')
      Get-ScheduledTaskAction | Export-Csv -LiteralPath (Join-Path $_textDir 'ScheduledTasks-Actions.csv') -NoTypeInformation -Encoding UTF8
      Get-ScheduledTaskAction -SuspiciousOnly | Export-Csv -LiteralPath (Join-Path $_textDir 'ScheduledTasks-Suspicious.csv') -NoTypeInformation -Encoding UTF8
    }
  }
  @{
    Name = 'Registry persistence - Run keys'
    Script = {
      Export-RegistryKey -Key 'HKLM\Software\Microsoft\Windows\CurrentVersion\Run' -OutputPath (Join-Path $_registryDir 'HKLM-Run.reg')
      Export-RegistryKey -Key 'HKCU\Software\Microsoft\Windows\CurrentVersion\Run' -OutputPath (Join-Path $_registryDir 'HKCU-Run.reg')
      Export-RegistryKey -Key 'HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run' -OutputPath (Join-Path $_registryDir 'HKLM-WOW6432-Run.reg')
    }
  }
)

$Quick = $Common + @(
  @{
    Name = 'Process and network state'
    Script = {
      $null = Invoke-SafeProcess -FilePath 'tasklist.exe' -ArgumentList @('/v') -OutputPath (Join-Path $_textDir 'Tasklist.txt')
      $null = Invoke-SafeProcess -FilePath 'netstat.exe' -ArgumentList @('-ano') -OutputPath (Join-Path $_textDir 'Netstat.txt')
      $null = Invoke-SafeProcess -FilePath 'wmic.exe' -ArgumentList @('process', 'get', 'ProcessId,ParentProcessId,ExecutablePath,CommandLine', '/format:list') -OutputPath (Join-Path $_textDir 'Processes-CommandLine.txt')
    }
  }
  @{
    Name = 'Users and groups'
    Script = {
      $null = Invoke-SafeProcess -FilePath 'net.exe' -ArgumentList @('user') -OutputPath (Join-Path $_textDir 'LocalUsers.txt')
      $null = Invoke-SafeProcess -FilePath 'net.exe' -ArgumentList @('localgroup', 'administrators') -OutputPath (Join-Path $_textDir 'LocalAdministrators.txt')
      $null = Invoke-SafeProcess -FilePath 'query.exe' -ArgumentList @('user') -OutputPath (Join-Path $_textDir 'LoggedOnUsers.txt')
    }
  }
  @{
    Name = 'Services'
    Script = {
      Get-Service | Sort-Object Name | Format-Table -AutoSize | Out-File -LiteralPath (Join-Path $_textDir 'Services.txt') -Encoding UTF8
      Get-CimInstance Win32_Service |
        Select-Object Name, DisplayName, State, StartMode, PathName, StartName |
        Sort-Object Name |
        Export-Csv -LiteralPath (Join-Path $_textDir 'Services-Detailed.csv') -NoTypeInformation -Encoding UTF8
    }
  }
)

$Full = $Quick + @(
  @{
    Name = 'Full event log export'
    Script = {
      $logs = @(
        'Application'
        'Microsoft-Windows-TaskScheduler/Operational'
        'Microsoft-Windows-WMI-Activity/Operational'
        'Microsoft-Windows-Shell-Core/Operational'
        'Microsoft-Windows-AppLocker/EXE and DLL'
        'Microsoft-Windows-AppLocker/MSI and Script'
        'Microsoft-Windows-CodeIntegrity/Operational'
        'Microsoft-Windows-Sysmon/Operational'
      )
      foreach ($l in $logs) {
        $safe = $l -replace '[\\/:*?"<>|]+', '_'
        Export-EventLog -LogName $l -OutputPath (Join-Path $_evtxDir "$safe.evtx") -MissingLogPath (Join-Path $_textDir 'Missing-EventLogs.txt')
      }
    }
  }
  @{
    Name = 'Registry persistence - Services and WINEVT'
    Script = {
      Export-RegistryKey -Key 'HKLM\SYSTEM\CurrentControlSet\Services' -OutputPath (Join-Path $_registryDir 'HKLM-Services.reg')
      Export-RegistryKey -Key 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT' -OutputPath (Join-Path $_registryDir 'HKLM-WINEVT.reg')
    }
  }
  @{
    Name = 'WMI persistence'
    Script = {
      $wmi = Get-WMIPersistence
      $wmi.EventFilters | Format-List * | Out-File -LiteralPath (Join-Path $_textDir 'WMI-EventFilters.txt') -Encoding UTF8
      $wmi.CommandLineConsumers | Format-List * | Out-File -LiteralPath (Join-Path $_textDir 'WMI-CommandLineConsumers.txt') -Encoding UTF8
      $wmi.Bindings | Format-List * | Out-File -LiteralPath (Join-Path $_textDir 'WMI-FilterToConsumerBindings.txt') -Encoding UTF8
    }
  }
  @{
    Name = 'ZIP archive'
    Script = {
      $zip = "$_caseDir.zip"
      if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
      Compress-Archive -Path $_caseDir -DestinationPath $zip -Force
      Write-Log -Message "  -> Created ZIP: $zip" -Color Green
    }
  }
)

$profileSets = @{
  Quick = $Quick
  Full = $Full
}

# ---- Profile resolution ------------------------------------------------------
$_resolvedProfile = $Profile
$_scriptNames = $profileSets[$_resolvedProfile]
if (-not $_scriptNames -or $_scriptNames.Count -eq 0) {
  Write-Log -Message "Profile '$_resolvedProfile' has no steps defined. Nothing to do." -Color Yellow
  exit 0
}

if ($ListOnly) {
  Write-Log -Message "Profile '$_resolvedProfile' runs these steps in order:" -Color Cyan
  foreach ($_step in $_scriptNames) {
    Write-Log -Message "  - $($_step.Name)" -Color Gray
  }
  exit 0
}

# ---- Execution ---------------------------------------------------------------
$_results = New-Object System.Collections.ArrayList
$_stepNumber = 0

foreach ($_step in $_scriptNames) {
  $_stepNumber++
  $_label = "[$_stepNumber/$($_scriptNames.Count)] $($_step.Name)"

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would collect: $_label" -Color Yellow
    Add-OperationResult -Results $_results -Target $_step.Name -Source 'InformationRetrieval' -Action 'Collect' -Status 'Skipped' -Detail 'DryRun'
    continue
  }

  Write-Log -Message "==> $_label" -Color Cyan
  $_stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  try {
    & $_step.Script

    $_stopwatch.Stop()
    Add-OperationResult -Results $_results -Target $_step.Name -Source 'InformationRetrieval' -Action 'Collect' -Status 'Completed' -Detail "Completed in $([math]::Round($_stopwatch.Elapsed.TotalSeconds, 1))s"
    Write-Log -Message "    done ($([math]::Round($_stopwatch.Elapsed.TotalSeconds, 1))s)" -Color Green
  }
  catch {
    $_stopwatch.Stop()
    Add-OperationResult -Results $_results -Target $_step.Name -Source 'InformationRetrieval' -Action 'Collect' -Status 'Failed' -Detail $_.Exception.Message
    Write-Log -Message "    FAILED: $($_step.Name)" -Color Red

    if ($StopOnError) {
      Write-Log -Message "Stopping - -StopOnError is set and step '$($_step.Name)' failed." -Color Red
      break
    }
  }
}

# ---- Summary -----------------------------------------------------------------
Write-Log -Message "`n=== Information retrieval summary (profile: $_resolvedProfile) ===" -Color Cyan

foreach ($_result in $_results) {
  $_status = if ($_result.Status -eq 'Completed') { 'OK' } else { if ($_result.Status -eq 'Failed') { 'FAIL' } else { 'SKIP' } }
  $_color = if ($_result.Status -eq 'Completed') { 'Green' } elseif ($_result.Status -eq 'Failed') { 'Red' } else { 'Gray' }
  Write-Log -Message "  [$_status] $($_result.Target)" -Color $_color
}

$_failedCount = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_completedCount = @($_results | Where-Object { $_.Status -eq 'Completed' }).Count
$_skippedCount = @($_results | Where-Object { $_.Status -ne 'Completed' -and $_.Status -ne 'Failed' }).Count

if (-not $DryRun) {
  Write-Log -Message "`nOutput: $_caseDir" -Color Cyan
}

if ($_failedCount -gt 0) {
  Write-Log -Message "`n$_completedCount succeeded | $_skippedCount skipped | $_failedCount failed" -Color Yellow
}
else {
  Write-Log -Message "`n$_completedCount succeeded | $_skippedCount skipped" -Color Green
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Invoke-InformationRetrieval'
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
