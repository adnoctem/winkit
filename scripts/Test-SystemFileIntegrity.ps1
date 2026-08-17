#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Verifies Windows system file integrity via sfc /verifyOnly.

.DESCRIPTION
  Runs the System File Checker in /verifyOnly mode (non-mutating: it checks
  integrity without repairing anything) and reports the outcome as a
  structured result. The exit code is the first-pass signal; the tail of
  %windir%\Logs\CBS\CBS.log - where SFC's actual findings live - is folded
  into the result as the source of truth.

  Deliberately does NOT run /scannow: a full scan-and-repair pass has a
  significant runtime and disruption cost and should only ever be run as a
  separate, explicitly chosen action.

  Requires administrator elevation. The verification can take several minutes
  on large systems; a watchdog bounds the wait and the script reports
  'Unknown' if sfc does not finish in time.

.PARAMETER TimeoutMinutes
  Watchdog ceiling for the sfc /verifyOnly run. Defaults to 30 minutes.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Test-SystemFileIntegrity.ps1
  Runs sfc /verifyOnly and reports whether integrity violations were found.

.EXAMPLE
  PS> ./Test-SystemFileIntegrity.ps1 -PassThru
  Returns the structured result including the CBS.log excerpt.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 240)]
  [int]
  $TimeoutMinutes = 30,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

$_results = New-Object System.Collections.ArrayList

$outFile = Join-Path $env:TEMP "sfc-verify-$(New-Guid).log"
$errFile = Join-Path $env:TEMP "sfc-verify-$(New-Guid).err.log"

try {
  Write-Log -Message "Running sfc /verifyOnly (can take several minutes, watchdog $TimeoutMinutes minutes)..." -Color Cyan

  $process = Start-Process -FilePath "$env:SystemRoot\System32\sfc.exe" -ArgumentList '/verifyOnly' -PassThru -NoNewWindow -RedirectStandardOutput $outFile -RedirectStandardError $errFile

  $startedAt = Get-Date
  $deadline = $startedAt.AddMinutes($TimeoutMinutes)
  $timedOut = $false
  while (-not $process.HasExited) {
    if ((Get-Date) -gt $deadline) {
      $timedOut = $true
      break
    }
    $elapsedTime = (Get-Date) - $startedAt
    $elapsedSeconds = [int]$elapsedTime.TotalSeconds
    $pct = [int][math]::Min(($elapsedSeconds / ($TimeoutMinutes * 60)) * 100, 100)
    Write-Progress -Activity 'sfc /verifyOnly' -Status "elapsed: $elapsedSeconds s / $TimeoutMinutes min" -PercentComplete $pct
    Start-Sleep -Seconds 5
  }
  Write-Progress -Activity 'sfc /verifyOnly' -Completed
  $totalSeconds = [int]((Get-Date) - $startedAt).TotalSeconds

  if ($timedOut) {
    Add-OperationResult -Results $_results -Target 'sfc /verifyOnly' -Source 'SFC' -Action 'Verify' -Status 'Unknown' -Detail "sfc did not finish within $TimeoutMinutes minutes ($totalSeconds s elapsed)."
    Write-Log -Message "sfc /verifyOnly did not finish within $TimeoutMinutes minutes." -Color Red
    if ($PassThru) { $_results }
    exit 1
  }

  $exitCode = $process.ExitCode

  $output = ''
  if (Test-Path -LiteralPath $outFile) {
    $output = Get-Content -LiteralPath $outFile -Raw
  }
  if (Test-Path -LiteralPath $errFile) {
    $output += (Get-Content -LiteralPath $errFile -Raw)
  }

  # CBS.log tail: the actual source of truth for what sfc found.
  $cbsLogPath = Join-Path $env:SystemRoot 'Logs\CBS\CBS.log'
  $cbsTail = @()
  if (Test-Path -LiteralPath $cbsLogPath) {
    $cbsTail = @(Get-Content -LiteralPath $cbsLogPath -Tail 30)
  }
  $cbsText = $cbsTail -join "`n"

  $foundViolations = (($output -match 'found corrupt files') -or ($cbsText -match 'found corrupt files'))
  $noViolations = (($output -match 'did not find any integrity violations') -or ($cbsText -match 'did not find any integrity violations'))
  $couldNotRun = (($output -match 'could not perform the requested operation') -or ($cbsText -match 'could not perform the requested operation'))

  if ($foundViolations) {
    $status = 'Failed'
    $detail = "Integrity violations found (exit code $exitCode, $totalSeconds s)."
  }
  elseif ($noViolations) {
    $status = 'Completed'
    $detail = "No integrity violations found (exit code $exitCode, $totalSeconds s)."
  }
  elseif ($couldNotRun) {
    $status = 'Failed'
    $detail = "sfc could not perform the requested operation (exit code $exitCode)."
  }
  elseif ($exitCode -eq 0) {
    $status = 'Completed'
    $detail = "sfc completed with exit code 0 ($totalSeconds s)."
  }
  else {
    $status = 'Failed'
    $detail = "sfc completed with exit code $exitCode ($totalSeconds s)."
  }

  $excerpt = if ($cbsTail.Count -gt 0) { ($cbsTail -join "`n") } else { '(no CBS.log content)' }
  Add-OperationResult -Results $_results -Target 'sfc /verifyOnly' -Source 'SFC' -Action 'Verify' -Status $status -Detail $detail -Property @{ LogExcerpt = $excerpt }

  if ($status -eq 'Completed') {
    Write-Log -Message $detail -Color Green
  }
  else {
    Write-Log -Message $detail -Color Red
  }
}
catch {
  Add-OperationResult -Results $_results -Target 'sfc /verifyOnly' -Source 'SFC' -Action 'Verify' -Status 'Failed' -Detail $_.Exception.Message
  Write-Log -Message "sfc /verifyOnly failed: $($_.Exception.Message)" -Color Red
}
finally {
  Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Test-SystemFileIntegrity'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_results
}

$failed = @($_results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Unknown' })
if ($failed.Count -gt 0) {
  exit 1
}
exit 0
