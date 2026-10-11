#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.1.1' }

<#
.SYNOPSIS
  Verifies Windows system file integrity via sfc /verifyOnly.

.DESCRIPTION
  Wrapper around AdNoctem.Substrate.PowerShell's Test-SystemFileIntegrity, which runs the
  System File Checker in /verifyOnly mode (non-mutating: it checks integrity
  without repairing anything) and synthesizes a structured verdict from the
  sfc output, the exit code, and the tail of %windir%\Logs\CBS\CBS.log - the
  actual source of truth for what sfc found.

  Deliberately does NOT run /scannow: a full scan-and-repair pass has a
  significant runtime and disruption cost and should only ever run as a
  separate, explicitly chosen action (see Repair-WindowsSystem.ps1).

  Requires administrator elevation. The verification can take several
  minutes and runs synchronously.

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
  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.1.1 -Force

$_results = New-Object System.Collections.ArrayList

Write-Log -Message 'Running sfc /verifyOnly (can take several minutes)...' -Color Cyan
$result = Test-SystemFileIntegrity

Add-OperationResult -Results $_results -Target 'sfc /verifyOnly' -Source 'SFC' -Action 'Verify' -Status $result.Status -Detail $result.Detail -Property @{ LogExcerpt = $result.LogExcerpt }

if ($result.Status -eq 'Completed') {
  Write-Log -Message $result.Detail -Color Green
}
else {
  Write-Log -Message $result.Detail -Color Red
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Test-SystemFileIntegrity' -Name 'winkit'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_results
}

if ($result.Status -ne 'Completed') {
  exit 1
}
exit 0
