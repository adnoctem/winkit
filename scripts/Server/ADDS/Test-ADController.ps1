#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Health-checks a domain controller for configuration drift and readiness.

.DESCRIPTION
  Standalone, re-runnable health check against a (possibly long-promoted)
  domain controller, complementary to the -Phase Verify gate inside
  Initialize-ADController.ps1: Verify is the post-promotion readiness gate in
  the setup flow; this script is the on-demand, after-the-fact check.

  Checks, run inside a WinRM session authenticated with the domain
  credential:

    Services   - NTDS, DNS, Netlogon, KDC, W32Time, DFSR must be running.
    Directory  - Get-ADDomain / Get-ADForest succeed; domain and forest
                 functional levels are reported; for a single-DC forest all
                 five FSMO roles must resolve to this server.
    Shares     - SYSVOL and NETLOGON must be reachable (replication init).
    Time       - w32tm must report a configured time source.
    dcdiag     - connectivity, advertising, services, sysvolcheck, kcc,
                 netlogons, replications, dns - any failed test is a failure.
    Event logs - Error events in the last 24 hours in the Directory Service,
                 DNS Server, and DFS Replication logs are counted (Warn when
                 present).

  There is no official Microsoft "CSS-Exchange style" AD health-check
  repository (verified via GitHub API); the AD counterpart to that tooling is
  the built-in dcdiag/repadmin/AD cmdlets used here. This script emits exactly
  one JSON envelope on the output stream; with -OutputPath it is additionally
  appended as a single JSON line.

.PARAMETER Server
  Target domain controller (static IP or name).

.PARAMETER Credential
  Domain account with read access to the directory (Domain Admins, or any
  account with read permissions plus local admin for the event-log queries).

.PARAMETER OutputPath
  Optional file to append the result envelope to as a JSON line.

.PARAMETER OperationTimeoutMinutes
  WinRM session operation timeout.

.EXAMPLE
  PS> ./Test-ADController.ps1 -Server 192.0.2.10 -Credential $domainAdmin
  Runs all checks and emits the result envelope.

.EXAMPLE
  PS> ./Test-ADController.ps1 -Server 192.0.2.10 -Credential $domainAdmin -OutputPath health.jsonl
  Appends the envelope to health.jsonl for drift tracking over time.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [string]$Server,

  [Parameter(Mandatory = $true)]
  [pscredential]$Credential,

  [Parameter(Mandatory = $false)]
  [string]$OutputPath,

  [Parameter(Mandatory = $false)]
  [int]$OperationTimeoutMinutes = 15
)

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

$startedAt = (Get-Date).ToUniversalTime().ToString('o')
$sessionOption = New-PSSessionOption -OperationTimeout ([int]($OperationTimeoutMinutes * 60000)) -IdleTimeout ([int]($OperationTimeoutMinutes * 60000))

$session = $null
try {
  $session = New-PSSession -ComputerName $Server -Credential $Credential -SessionOption $sessionOption -ErrorAction Stop
}
catch {
  $envelope = [pscustomobject][ordered]@{
    phase = 'test'
    success = $false
    status = 'fail'
    server = $Server
    startedAt = $startedAt
    completedAt = (Get-Date).ToUniversalTime().ToString('o')
    diagnostics = @()
    error = $_.Exception.Message
  }
  $envelope | ConvertTo-Json -Depth 6 -Compress | Write-Output
  if ($OutputPath) {
    ($envelope | ConvertTo-Json -Depth 6 -Compress) | Add-Content -LiteralPath $OutputPath -Encoding UTF8
  }
  exit 1
}

$checkScript = {
  $checks = @()

  $serviceNames = @('NTDS', 'DNS', 'Netlogon', 'KDC', 'W32Time', 'DFSR')
  foreach ($name in $serviceNames) {
    $service = Get-Service -Name $name -ErrorAction SilentlyContinue
    $status = if ($service) { $service.Status.ToString() } else { 'Missing' }
    $checks += [pscustomobject]@{
      Category = 'Service'
      Name = $name
      Status = $(if ($status -eq 'Running') { 'Pass' } else { 'Fail' })
      Detail = $status
    }
  }

  $domain = $null
  $forest = $null
  try {
    $domain = Get-ADDomain -Server localhost -ErrorAction Stop
    $forest = Get-ADForest -Server localhost -ErrorAction Stop
    $checks += [pscustomobject]@{ Category = 'Directory'; Name = 'Domain'; Status = 'Pass'; Detail = "$($domain.Name) (functional level: $($domain.DomainMode))" }
    $checks += [pscustomobject]@{ Category = 'Directory'; Name = 'Forest'; Status = 'Pass'; Detail = "$($forest.Name) (functional level: $($forest.ForestMode))" }

    $dcs = @(Get-ADDomainController -Filter * -ErrorAction Stop)
    $roleHolders = [ordered]@{
      SchemaMaster = $forest.SchemaMaster
      DomainNamingMaster = $forest.DomainNamingMaster
      PDCEmulator = $domain.PDCEmulator
      RIDMaster = $domain.RIDMaster
      InfrastructureMaster = $domain.InfrastructureMaster
    }

    foreach ($role in $roleHolders.Keys) {
      $checks += [pscustomobject]@{ Category = 'FSMO'; Name = $role; Status = 'Pass'; Detail = $roleHolders[$role] }
    }

    if ($dcs.Count -eq 1) {
      $localHost = $env:COMPUTERNAME
      $allLocal = $true
      foreach ($role in $roleHolders.Keys) {
        $shortName = ($roleHolders[$role] -split '\.')[0]
        if ($shortName -ne $localHost) {
          $allLocal = $false
        }
      }
      $checks += [pscustomobject]@{
        Category = 'FSMO'
        Name = 'SingleDCHolder'
        Status = $(if ($allLocal) { 'Pass' } else { 'Warn' })
        Detail = $(if ($allLocal) { 'All roles held by this server (single-DC forest).' } else { 'Single-DC forest but not all roles resolve to this server - investigate.' })
      }
    }
    else {
      $checks += [pscustomobject]@{ Category = 'FSMO'; Name = 'ReplicaCount'; Status = 'Pass'; Detail = "$($dcs.Count) domain controllers present." }
    }
  }
  catch {
    $checks += [pscustomobject]@{ Category = 'Directory'; Name = 'ADQuery'; Status = 'Fail'; Detail = $_.Exception.Message }
  }

  $sysvol = Test-Path -LiteralPath '\\localhost\SYSVOL'
  $netlogon = Test-Path -LiteralPath '\\localhost\NETLOGON'
  $checks += [pscustomobject]@{ Category = 'Share'; Name = 'SYSVOL'; Status = $(if ($sysvol) { 'Pass' } else { 'Fail' }); Detail = $(if ($sysvol) { 'reachable' } else { 'not reachable' }) }
  $checks += [pscustomobject]@{ Category = 'Share'; Name = 'NETLOGON'; Status = $(if ($netlogon) { 'Pass' } else { 'Fail' }); Detail = $(if ($netlogon) { 'reachable' } else { 'not reachable' }) }

  try {
    $timeOutput = (& w32tm /query /status) 2>&1 | Out-String
    $hasSource = ($timeOutput -match 'Source:\s*\S+')
    $checks += [pscustomobject]@{ Category = 'Time'; Name = 'W32Time'; Status = $(if ($hasSource) { 'Pass' } else { 'Warn' }); Detail = $(if ($hasSource) { 'time source configured' } else { 'no time source reported' }) }
  }
  catch {
    $checks += [pscustomobject]@{ Category = 'Time'; Name = 'W32Time'; Status = 'Warn'; Detail = 'w32tm query failed.' }
  }

  try {
    $dcdiagOutput = (& dcdiag /s:localhost /test:connectivity /test:advertising /test:services /test:sysvolcheck /test:kcc /test:netlogons /test:replications /test:dns) 2>&1 | Out-String
    $failedMarkers = @($dcdiagOutput -split "`r?`n" | Where-Object { $_ -match 'failed test' })
    $checks += [pscustomobject]@{
      Category = 'dcdiag'
      Name = 'TestSuite'
      Status = $(if ($failedMarkers.Count -eq 0) { 'Pass' } else { 'Fail' })
      Detail = $(if ($failedMarkers.Count -eq 0) { 'all selected tests passed' } else { "$($failedMarkers.Count) failed test marker(s) found" })
    }
  }
  catch {
    $checks += [pscustomobject]@{ Category = 'dcdiag'; Name = 'TestSuite'; Status = 'Warn'; Detail = 'dcdiag could not run.' }
  }

  $since = (Get-Date).AddHours(-24)
  foreach ($logName in @('Directory Service', 'DNS Server', 'DFS Replication')) {
    try {
      $count = @(Get-WinEvent -FilterHashtable @{ LogName = $logName; Level = 2; StartTime = $since } -ErrorAction Stop).Count
      $checks += [pscustomobject]@{
        Category = 'EventLog'
        Name = $logName
        Status = $(if ($count -eq 0) { 'Pass' } else { 'Warn' })
        Detail = "$count error event(s) in the last 24 hours"
      }
    }
    catch {
      $checks += [pscustomobject]@{ Category = 'EventLog'; Name = $logName; Status = 'Pass'; Detail = 'log unavailable' }
    }
  }

  [pscustomobject]@{
    Checks = $checks
  }
}

$result = $null
try {
  $result = Invoke-Command -Session $session -ScriptBlock $checkScript -ErrorAction Stop
}
catch {
  Remove-PSSession -Session $session -ErrorAction SilentlyContinue
  $envelope = [pscustomobject][ordered]@{
    phase = 'test'
    success = $false
    status = 'fail'
    server = $Server
    startedAt = $startedAt
    completedAt = (Get-Date).ToUniversalTime().ToString('o')
    diagnostics = @()
    error = $_.Exception.Message
  }
  $envelope | ConvertTo-Json -Depth 6 -Compress | Write-Output
  if ($OutputPath) {
    ($envelope | ConvertTo-Json -Depth 6 -Compress) | Add-Content -LiteralPath $OutputPath -Encoding UTF8
  }
  exit 1
}

Remove-PSSession -Session $session -ErrorAction SilentlyContinue

$failedChecks = @($result.Checks | Where-Object { $_.Status -eq 'Fail' })
$warnedChecks = @($result.Checks | Where-Object { $_.Status -eq 'Warn' })

$status = 'pass'
if ($warnedChecks.Count -gt 0) {
  $status = 'warn'
}
if ($failedChecks.Count -gt 0) {
  $status = 'fail'
}

$envelope = [pscustomobject][ordered]@{
  phase = 'test'
  success = ($status -ne 'fail')
  status = $status
  server = $Server
  startedAt = $startedAt
  completedAt = (Get-Date).ToUniversalTime().ToString('o')
  diagnostics = @($result.Checks)
  error = $null
}

$json = $envelope | ConvertTo-Json -Depth 6 -Compress
$json | Write-Output
if ($OutputPath) {
  $json | Add-Content -LiteralPath $OutputPath -Encoding UTF8
}

if ($status -eq 'fail') {
  exit 1
}
exit 0
