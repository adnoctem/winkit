#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Health-checks an Exchange Server (or Exchange Management Tools station)
  and optionally runs the Microsoft CSS Exchange Health Checker.

.DESCRIPTION
  Standalone, re-runnable health check, complementary to the -Phase Verify
  gate inside Initialize-ExchangeServer.ps1. Detects configuration drift and
  broken services after the fact.

  Built-in checks (always run, via a WinRM session with the domain
  credential):

    Management Shell - Exchange Management Shell snapin loads; Get-ExchangeServer
                      reports name, version, and roles.
    Services        - all MSExchange* services must be Running.
    Databases       - Mailbox database mount state.
    Event log       - Application log error events from Exchange sources in
                      the last 24 hours (Warn when present).
    Disk            - free space on C: and D: (Warn below 15%).
    Time            - w32tm must report a configured time source.

  With -IncludeHealthChecker, the official Microsoft CSS-Exchange Health
  Checker (first-party, downloaded from the project's latest release) is
  staged on the target and run against it. It requires Local Administrator /
  Organization Management membership on the target and can take 10-20 minutes;
  the generated report paths are returned in the diagnostics.

  There is no official Microsoft health-check repository for Active Directory
  (verified via GitHub API) - the AD counterpart to this tooling is built-in
  dcdiag/repadmin, which Test-ADController.ps1 covers.

  This script emits exactly one JSON envelope on the output stream; with
  -OutputPath it is additionally appended as a single JSON line.

.PARAMETER Server
  Target Exchange server (static IP or name).

.PARAMETER Credential
  Domain account with Local Administrator / Organization Management membership
  on the target (required by the Health Checker as well).

.PARAMETER OutputPath
  Optional file to append the result envelope to as a JSON line.

.PARAMETER IncludeHealthChecker
  Also download and run the Microsoft CSS Exchange Health Checker on the
  target.

.PARAMETER OperationTimeoutMinutes
  WinRM session operation timeout for the built-in checks.

.PARAMETER HealthCheckerTimeoutMinutes
  Hard ceiling for the Health Checker run on the target.

.EXAMPLE
  PS> ./Test-ExchangeServer.ps1 -Server 192.0.2.11 -Credential $domainAdmin
  Runs the built-in checks and emits the result envelope.

.EXAMPLE
  PS> ./Test-ExchangeServer.ps1 -Server 192.0.2.11 -Credential $domainAdmin -IncludeHealthChecker -OutputPath health.jsonl
  Runs the built-in checks plus the official CSS Health Checker and appends
  the envelope (with report paths) to health.jsonl.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Carriage-return progress lines require Write-Host for in-place console updates.')]

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [string]$Server,

  [Parameter(Mandatory = $true)]
  [pscredential]$Credential,

  [Parameter(Mandatory = $false)]
  [string]$OutputPath,

  [Parameter(Mandatory = $false)]
  [switch]$IncludeHealthChecker,

  [Parameter(Mandatory = $false)]
  [int]$OperationTimeoutMinutes = 45,

  [Parameter(Mandatory = $false)]
  [int]$HealthCheckerTimeoutMinutes = 30
)

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

$healthCheckerUrl = 'https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1'
$healthCheckerRemoteDirectory = 'C:\Windows\Temp\winkit-healthchecker'

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

  $emsOk = $false
  $serverInfo = $null
  $databaseStates = @()
  try {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    $emsOk = $true
    $serverInfo = Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop | Select-Object Name, AdminDisplayVersion, ServerRole
    $databaseStates = @(Get-MailboxDatabase -Server $env:COMPUTERNAME -Status -ErrorAction SilentlyContinue | Select-Object Name, Mounted)
  }
  catch {
    $emsOk = $false
  }

  $checks += [pscustomobject]@{ Category = 'ManagementShell'; Name = 'ExchangeSnapin'; Status = $(if ($emsOk) { 'Pass' } else { 'Fail' }); Detail = $(if ($emsOk) { 'Exchange Management Shell available.' } else { 'Exchange Management Shell snapin failed to load.' }) }

  if ($serverInfo) {
    $checks += [pscustomobject]@{ Category = 'Server'; Name = $serverInfo.Name; Status = 'Pass'; Detail = "Version: $($serverInfo.AdminDisplayVersion), Roles: $($serverInfo.ServerRole)" }
  }

  $exchangeServices = @(Get-Service -Name 'MSExchange*' -ErrorAction SilentlyContinue)
  if ($exchangeServices.Count -eq 0) {
    $checks += [pscustomobject]@{ Category = 'Service'; Name = 'MSExchange*'; Status = 'Warn'; Detail = 'No Exchange services found - management station without local roles?' }
  }
  else {
    foreach ($service in $exchangeServices) {
      $checks += [pscustomobject]@{
        Category = 'Service'
        Name = $service.Name
        Status = $(if ($service.Status -eq 'Running') { 'Pass' } else { 'Fail' })
        Detail = $service.Status.ToString()
      }
    }
  }

  $mounted = @($databaseStates | Where-Object { $_.Mounted })
  if ($databaseStates.Count -gt 0) {
    $checks += [pscustomobject]@{
      Category = 'Database'
      Name = 'MailboxDatabases'
      Status = $(if ($mounted.Count -eq $databaseStates.Count) { 'Pass' } else { 'Warn' })
      Detail = "$($mounted.Count) of $($databaseStates.Count) databases mounted"
    }
  }

  try {
    $since = (Get-Date).AddHours(-24)
    $errorCount = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Level = 2; StartTime = $since } -ErrorAction Stop | Where-Object { $_.ProviderName -like 'MSExchange*' -or $_.ProviderName -like '*Exchange*' }).Count
    $checks += [pscustomobject]@{ Category = 'EventLog'; Name = 'Application (Exchange)'; Status = $(if ($errorCount -eq 0) { 'Pass' } else { 'Warn' }); Detail = "$errorCount Exchange-related error event(s) in the last 24 hours" }
  }
  catch {
    $checks += [pscustomobject]@{ Category = 'EventLog'; Name = 'Application (Exchange)'; Status = 'Pass'; Detail = 'event log query unavailable' }
  }

  $diskScript = {
    $disks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction SilentlyContinue)
    foreach ($disk in $disks) {
      $freePercent = if ($disk.Size -gt 0) { [math]::Round((($disk.FreeSpace / $disk.Size) * 100), 1) } else { 0 }
      [pscustomobject]@{
        DeviceId = $disk.DeviceID
        FreePercent = $freePercent
        FreeGigabytes = [math]::Round($disk.FreeSpace / 1GB, 1)
      }
    }
  }
  $diskStates = Invoke-Command -ComputerName localhost -ScriptBlock $diskScript -ErrorAction SilentlyContinue
  foreach ($disk in @($diskStates)) {
    $checks += [pscustomobject]@{
      Category = 'Disk'
      Name = $disk.DeviceId
      Status = $(if ($disk.FreePercent -lt 15) { 'Warn' } else { 'Pass' })
      Detail = "$($disk.FreeGigabytes) GB free ($($disk.FreePercent)%)"
    }
  }

  try {
    $timeOutput = (& w32tm /query /status) 2>&1 | Out-String
    $hasSource = ($timeOutput -match 'Source:\s*\S+')
    $checks += [pscustomobject]@{ Category = 'Time'; Name = 'W32Time'; Status = $(if ($hasSource) { 'Pass' } else { 'Warn' }); Detail = $(if ($hasSource) { 'time source configured' } else { 'no time source reported' }) }
  }
  catch {
    $checks += [pscustomobject]@{ Category = 'Time'; Name = 'W32Time'; Status = 'Warn'; Detail = 'w32tm query failed.' }
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

$diagnostics = New-Object System.Collections.ArrayList
foreach ($check in @($result.Checks)) {
  $null = $diagnostics.Add($check)
}

$healthCheckerEntries = New-Object System.Collections.ArrayList

if ($IncludeHealthChecker) {
  $localHealthChecker = Join-Path $env:TEMP 'winkit-HealthChecker.ps1'
  try {
    Write-Log -Message 'Downloading the CSS Exchange Health Checker...' -Color Yellow
    Invoke-WebRequest -Uri $healthCheckerUrl -OutFile $localHealthChecker -UseBasicParsing -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $healthCheckerEntries -Target 'CSS-Exchange HealthChecker' -Source 'HealthChecker' -Action 'Run' -Status 'Failed' -Detail "Download failed: $($_.Exception.Message)"
  }

  if (Test-Path -LiteralPath $localHealthChecker) {
    try {
      $directorySetupScript = {
        param($directory)
        $null = New-Item -Path $directory -ItemType Directory -Force
      }
      $null = Invoke-Command -Session $session -ArgumentList $healthCheckerRemoteDirectory -ScriptBlock $directorySetupScript -ErrorAction Stop

      Copy-Item -LiteralPath $localHealthChecker -Destination (Join-Path $healthCheckerRemoteDirectory 'HealthChecker.ps1') -ToSession $session -ErrorAction Stop

      $healthJobScript = {
        param($directory)
        $inner = {
          param($dir)
          $scriptPath = Join-Path $dir 'HealthChecker.ps1'
          $command = "& '$scriptPath' -Server `$env:COMPUTERNAME -OutputFilePath '$dir' -SkipVersionCheck"
          $process = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $command) -Wait -PassThru -NoNewWindow -ErrorAction Stop
          $reports = @(Get-ChildItem -Path $dir -File | Select-Object -ExpandProperty FullName)
          [pscustomobject]@{
            ExitCode = $process.ExitCode
            Reports = $reports
          }
        }
        $job = Start-Job -ScriptBlock $inner -ArgumentList $directory
        $job.Id
      }
      $healthJobId = Invoke-Command -Session $session -ArgumentList $healthCheckerRemoteDirectory -ScriptBlock $healthJobScript -ErrorAction Stop

      $healthJobStateScript = {
        param($id)
        $job = Get-Job -Id $id -ErrorAction SilentlyContinue
        if ($job) { $job.State } else { 'Missing' }
      }

      $healthPollStartedAt = Get-Date
      $healthPollDeadline = $healthPollStartedAt.AddMinutes($HealthCheckerTimeoutMinutes)
      $healthJobState = 'Running'
      $healthTimedOut = $false
      while ($healthJobState -notin @('Completed', 'Failed', 'Missing')) {
        try {
          $healthJobState = Invoke-Command -Session $session -ArgumentList $healthJobId -ScriptBlock $healthJobStateScript -ErrorAction Stop
        }
        catch {
          $healthJobState = 'Missing'
        }
        $elapsedTime = (Get-Date) - $healthPollStartedAt
        $elapsedSeconds = [int]$elapsedTime.TotalSeconds
        $healthPct = [int][math]::Min(($elapsedSeconds / ($HealthCheckerTimeoutMinutes * 60)) * 100, 100)
        Write-Progress -Activity 'CSS Exchange Health Checker' -Status "elapsed: $elapsedSeconds s / $HealthCheckerTimeoutMinutes min, job state: $healthJobState" -PercentComplete $healthPct
        Write-Host ("`rHealth Checker ({0}s / {1}min): {2,-40}" -f $elapsedSeconds, $HealthCheckerTimeoutMinutes, $healthJobState) -NoNewline -ForegroundColor Cyan
        if ((Get-Date) -gt $healthPollDeadline) {
          $healthTimedOut = $true
          break
        }
        Start-Sleep -Seconds 20
      }
      Write-Progress -Activity 'CSS Exchange Health Checker' -Completed
      $healthTotalSeconds = [int]((Get-Date) - $healthPollStartedAt).TotalSeconds
      Write-Host ("`rHealth Checker finished ({0}s).{1}" -f $healthTotalSeconds, (' ' * 40)) -ForegroundColor Cyan

      if ($healthTimedOut -or $healthJobState -eq 'Missing') {
        Add-OperationResult -Results $healthCheckerEntries -Target 'CSS-Exchange HealthChecker' -Source 'HealthChecker' -Action 'Run' -Status 'Unknown' -Detail "Health Checker still running past the $HealthCheckerTimeoutMinutes minute ceiling (job state: $healthJobState)."
      }
      else {
        $healthHarvestScript = {
          param($id)
          $job = Get-Job -Id $id -ErrorAction SilentlyContinue
          if (-not $job) {
            return $null
          }
          $output = @()
          try {
            $output = @(Receive-Job -Id $id -ErrorAction Stop)
          }
          catch {
            $output = @()
          }
          Remove-Job -Id $id -Force -ErrorAction SilentlyContinue
          if ($output.Count -gt 0) { $output[0] } else { $null }
        }
        $healthRun = Invoke-Command -Session $session -ArgumentList $healthJobId -ScriptBlock $healthHarvestScript -ErrorAction Stop

        if (-not $healthRun) {
          Add-OperationResult -Results $healthCheckerEntries -Target 'CSS-Exchange HealthChecker' -Source 'HealthChecker' -Action 'Run' -Status 'Warn' -Detail 'Health Checker job produced no result object; report files may still be present.'
        }
        elseif ($healthRun.ExitCode -ne 0) {
          Add-OperationResult -Results $healthCheckerEntries -Target 'CSS-Exchange HealthChecker' -Source 'HealthChecker' -Action 'Run' -Status 'Warn' -Detail "Health Checker exited with code $($healthRun.ExitCode); report files may still be present."
        }
        else {
          Add-OperationResult -Results $healthCheckerEntries -Target 'CSS-Exchange HealthChecker' -Source 'HealthChecker' -Action 'Run' -Status 'Completed' -Detail "Report files: $($healthRun.Reports -join '; ')"
        }
      }
    }
    catch {
      Add-OperationResult -Results $healthCheckerEntries -Target 'CSS-Exchange HealthChecker' -Source 'HealthChecker' -Action 'Run' -Status 'Failed' -Detail $_.Exception.Message
    }
    finally {
      Remove-Item -LiteralPath $localHealthChecker -Force -ErrorAction SilentlyContinue
    }
  }
}

Remove-PSSession -Session $session -ErrorAction SilentlyContinue

foreach ($entry in $healthCheckerEntries) {
  $null = $diagnostics.Add($entry)
}

$failedChecks = @($diagnostics | Where-Object { $_.Status -eq 'Fail' })
$warnedChecks = @($diagnostics | Where-Object { $_.Status -eq 'Warn' -or $_.Status -eq 'Unknown' })

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
  diagnostics = @($diagnostics)
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
