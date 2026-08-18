#Requires -Version 5.0
#Requires -Module ActiveDirectory
#Requires -Module GroupPolicy
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Checks GPO replication health across all domain controllers.

.DESCRIPTION
  For every domain controller, compares each GPO's DSVersion (the version
  stored in Active Directory) against SysVolVersion (the version stored on
  the DC's SYSVOL share). A mismatch between the two on any DC is the
  standard diagnostic signal for GPO replication problems.

  Read-only diagnostic - nothing is modified, making it safe to run against
  any DC, including one you prefer not to touch otherwise.

  Emits structured report objects to the success stream; -OutputPath writes
  them as CSV (Export-AD.ps1 convention). No elevation needed. Requires the
  GroupPolicy and ActiveDirectory modules (RSAT).

.PARAMETER Server
  Optional domain controller to target for the DC enumeration.

.PARAMETER OutputPath
  Optional CSV file to write the report to.

.EXAMPLE
  PS> ./Get-ADGPOReplication.ps1
  Reports DSVersion vs SysVolVersion for every GPO on every DC.

.EXAMPLE
  PS> ./Get-ADGPOReplication.ps1 -OutputPath gpo-replication.csv

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string]
  $Server,

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath
)

Import-Module PSFoundation -Force

$commonParams = @{}
if ($Server) {
  $commonParams.Server = $Server
}

Write-Log -Message 'Enumerating domain controllers...' -Color Cyan
try {
  $domainControllers = @(Get-ADDomainController -Filter * @commonParams -ErrorAction Stop)
}
catch {
  Write-Log -Message "Domain controller query failed: $($_.Exception.Message)" -Color Red
  exit 1
}

if ($domainControllers.Count -eq 0) {
  Write-Log -Message 'No domain controllers found.' -Color Red
  exit 1
}

$report = @()
$failedDcs = 0
$outOfSync = 0

foreach ($dc in $domainControllers) {
  Write-Log -Message "Checking GPO replication on $($dc.Name)..." -Color Cyan

  $gpos = @()
  try {
    $gpos = @(Get-GPO -All -Server $dc.Name -ErrorAction Stop)
  }
  catch {
    $failedDcs++
    Write-Log -Message "  Could not query GPOs on $($dc.Name): $($_.Exception.Message)" -Color Red
    continue
  }

  foreach ($gpo in $gpos) {
    $inSync = ($gpo.DSVersion -eq $gpo.SysVolVersion)
    if (-not $inSync) {
      $outOfSync++
    }
    $report += [pscustomobject]@{
      DomainController = $dc.Name
      GpoName = $gpo.DisplayName
      GpoId = $gpo.Id
      DSVersion = $gpo.DSVersion
      SysVolVersion = $gpo.SysVolVersion
      InSync = $inSync
    }
  }
}

$report = @($report | Sort-Object GpoName, DomainController)

if ($report.Count -eq 0) {
  Write-Log -Message 'No GPOs found (or none could be queried).' -Color Yellow
  exit 1
}

$report | Write-Output

if ($OutputPath) {
  $report | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
  Write-Log -Message "Report written to $OutputPath" -Color Gray
}

if ($outOfSync -gt 0) {
  Write-Log -Message "$outOfSync GPO/DC combination(s) out of sync - investigate GPO replication." -Color Red
  exit 1
}

Write-Log -Message "All $($report.Count) GPO/DC combinations in sync." -Color Green
exit 0
