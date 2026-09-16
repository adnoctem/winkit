#Requires -Version 5.0
#Requires -Module ActiveDirectory
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Reports AD users whose account (not password) is about to expire.

.DESCRIPTION
  Lists users whose AD account-expiration date (AccountExpirationDate -
  typically set for contractor/temporary accounts with a defined end date)
  falls within the next -DaysUntilExpiry days, via Search-ADAccount
  -AccountExpiring.

  This is a DIFFERENT mechanism from password expiration and is deliberately
  NOT merged with Get-ADPasswordExpiry.ps1: that script reports password
  expiry, this one reports account expiry - complementary, audit-relevant
  reports.

  Emits structured report objects to the success stream; -OutputPath writes
  them as CSV (Export-AD.ps1 convention). No elevation needed. Email
  delivery is left to Send-SMTPMessage.ps1 composition, not embedded in the
  report logic.

.PARAMETER DaysUntilExpiry
  Report accounts expiring within this many days. Defaults to 30.

.PARAMETER Server
  Optional domain controller to query.

.PARAMETER OutputPath
  Optional CSV file to write the report to.

.EXAMPLE
  PS> ./Get-ADExpiringAccount.ps1
  Lists users whose account expires within 30 days.

.EXAMPLE
  PS> ./Get-ADExpiringAccount.ps1 -DaysUntilExpiry 14 -Server dc01.company.com -OutputPath expiring-accounts.csv

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
  $DaysUntilExpiry = 30,

  [Parameter(Mandatory = $false)]
  [string]
  $Server,

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath
)

Import-Module PSFoundation -Force

Write-Log -Message "Querying accounts expiring within $DaysUntilExpiry day(s)..." -Color Cyan

$searchParams = @{
  AccountExpiring = $true
  TimeSpan = (New-TimeSpan -Days $DaysUntilExpiry)
  UsersOnly = $true
}
if ($Server) {
  $searchParams.Server = $Server
}

try {
  $accounts = @(Search-ADAccount @searchParams -ErrorAction Stop)
}
catch {
  Write-Log -Message "AD query failed: $($_.Exception.Message)" -Color Red
  exit 1
}

$now = Get-Date
$report = @()
foreach ($account in $accounts) {
  if (-not $account.AccountExpirationDate) {
    continue
  }
  $report += [pscustomobject]@{
    SamAccountName = $account.SamAccountName
    Name = $account.Name
    AccountExpirationDate = $account.AccountExpirationDate
    DaysUntilExpiration = [math]::Round(($account.AccountExpirationDate - $now).TotalDays, 1)
    Enabled = $account.Enabled
  }
}

$report = @($report | Sort-Object AccountExpirationDate)

if ($report.Count -eq 0) {
  Write-Log -Message "No accounts found expiring within $DaysUntilExpiry day(s)." -Color Green
  exit 0
}

$report | Write-Output

if ($OutputPath) {
  $report | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
  Write-Log -Message "Report written to $OutputPath" -Color Gray
}

Write-Log -Message "$($report.Count) account(s) expiring within $DaysUntilExpiry day(s)." -Color Yellow
exit 0
