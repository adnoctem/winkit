#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.1.0' }

<#
.SYNOPSIS
  Reports certificates that expire within a given number of days.

.DESCRIPTION
  Standing certificate-expiry report: wraps PSFoundation's
  Get-CertificateInventory to list certificates in a store that expire
  within -DaysLeft days (optionally excluding already-expired ones via
  -NotExpired). Certificate-expiry monitoring closes a classic operational
  gap - an expired internal certificate causing an outage is a common
  failure mode.

  Emits structured report objects to the success stream; -OutputPath writes
  them as CSV. No elevation needed for the LocalMachine\My store of the
  current user context (other stores may require elevation).

.PARAMETER DaysLeft
  Only report certificates expiring within this many days. Defaults to 30.

.PARAMETER StorePath
  Certificate store to query. Defaults to Cert:\LocalMachine\My. Passed
  through to Get-CertificateInventory's validated StorePath parameter.

.PARAMETER NotExpired
  Exclude already-expired certificates.

.PARAMETER OutputPath
  Optional CSV file to write the report to.

.EXAMPLE
  PS> ./Get-CertificateReport.ps1
  Reports certificates in the local machine My store expiring within 30 days.

.EXAMPLE
  PS> ./Get-CertificateReport.ps1 -DaysLeft 90 -StorePath Cert:\LocalMachine\CA -NotExpired -OutputPath certs-expiring.csv

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 3650)]
  [int]
  $DaysLeft = 30,

  [Parameter(Mandatory = $false)]
  [string]
  $StorePath = 'Cert:\LocalMachine\My',

  [Parameter(Mandatory = $false)]
  [switch]
  $NotExpired,

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath
)

Import-Module PSFoundation -Force

Write-Log -Message "Querying certificates in $StorePath expiring within $DaysLeft day(s)..." -Color Cyan

try {
  $certificates = @(Get-CertificateInventory -StorePath $StorePath -DaysLeft $DaysLeft -NotExpired:$NotExpired -ErrorAction Stop)
}
catch {
  Write-Log -Message "Certificate query failed: $($_.Exception.Message)" -Color Red
  exit 1
}

$report = @($certificates | Sort-Object DaysUntilExpired)

if ($report.Count -eq 0) {
  Write-Log -Message "No certificates expiring within $DaysLeft day(s) found in $StorePath." -Color Green
  exit 0
}

$report | Write-Output

if ($OutputPath) {
  $report | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
  Write-Log -Message "Report written to $OutputPath" -Color Gray
}

Write-Log -Message "$($report.Count) certificate(s) expiring within $DaysLeft day(s)." -Color Yellow
exit 0
