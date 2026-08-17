#Requires -Version 5.0
#Requires -Module ActiveDirectory
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Reports AD users whose password expires within a given number of days.

.DESCRIPTION
  Lists enabled AD users whose password expires within -DaysUntilExpiry days,
  using the computed msDS-UserPasswordExpiryTimeComputed attribute as the
  authoritative expiry source. That attribute reflects the user's EFFECTIVE
  password policy - including Fine-Grained Password Policies (FGPP) - rather
  than assuming the default domain password policy applies to everyone, which
  would be silently wrong for FGPP-covered users. Users with
  PasswordNeverExpires are excluded.

  Emits structured report objects to the success stream; -OutputPath writes
  them as CSV (Export-AD.ps1 convention). No elevation needed. For email
  delivery, pipe the objects or the CSV into Send-SMTPMessage.ps1.

.PARAMETER DaysUntilExpiry
  Report users expiring within this many days. Defaults to 14.

.PARAMETER Server
  Optional domain controller to query.

.PARAMETER OutputPath
  Optional CSV file to write the report to.

.EXAMPLE
  PS> ./Get-ADPasswordExpiry.ps1
  Lists users whose password expires within 14 days.

.EXAMPLE
  PS> ./Get-ADPasswordExpiry.ps1 -DaysUntilExpiry 30 -Server dc01.company.com -OutputPath expiring.csv
  Writes the 30-day expiry report to expiring.csv.

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
  $DaysUntilExpiry = 14,

  [Parameter(Mandatory = $false)]
  [string]
  $Server,

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath
)

Import-Module PSFoundation -Force

$now = Get-Date
$cutoff = $now.AddDays($DaysUntilExpiry)

Write-Log -Message "Querying users with password expiry within $DaysUntilExpiry day(s)..." -Color Cyan

$commonParams = @{}
if ($Server) {
  $commonParams.Server = $Server
}

try {
  $users = @(Get-ADUser -Filter 'Enabled -eq $true' -Properties DisplayName, Mail, PasswordLastSet, PasswordNeverExpires, 'msDS-UserPasswordExpiryTimeComputed' @commonParams -ResultSetSize $null -ErrorAction Stop)
}
catch {
  Write-Log -Message "AD query failed: $($_.Exception.Message)" -Color Red
  exit 1
}

$report = @()
foreach ($user in $users) {
  if ($user.PasswordNeverExpires) {
    continue
  }

  $expiryTicks = [long]$user.'msDS-UserPasswordExpiryTimeComputed'
  if ($expiryTicks -le 0) {
    continue
  }

  $expiryDate = [datetime]::FromFileTime($expiryTicks)
  if ($expiryDate -gt $cutoff) {
    continue
  }

  $report += [pscustomobject]@{
    SamAccountName = $user.SamAccountName
    DisplayName = $user.DisplayName
    Mail = $user.Mail
    PasswordLastSet = $user.PasswordLastSet
    ExpiryDate = $expiryDate
    DaysUntilExpiry = [math]::Round(($expiryDate - $now).TotalDays, 1)
  }
}

$report = @($report | Sort-Object ExpiryDate)

if ($report.Count -eq 0) {
  Write-Log -Message "No users found with password expiry within $DaysUntilExpiry day(s)." -Color Green
  exit 0
}

$report | Write-Output

if ($OutputPath) {
  $report | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
  Write-Log -Message "Report written to $OutputPath" -Color Gray
}

Write-Log -Message "$($report.Count) user(s) expiring within $DaysUntilExpiry day(s)." -Color Yellow
exit 0
