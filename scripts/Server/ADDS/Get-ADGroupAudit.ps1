#Requires -Version 5.0
#Requires -Module ActiveDirectory
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Reports last-logon times for members of privileged AD groups.

.DESCRIPTION
  Enumerates the members of the given security groups and reports the most
  recent logon for each member across ALL domain controllers. The LastLogon
  attribute is DC-local and not replicated, so taking the maximum raw value
  across every DC is the only correct way to determine a user's last logon -
  querying a single DC silently underreports it.

  Adapted from Get-DomainAdmins.ps1 in nickrod518/PowerShell-Scripts,
  generalized to accept any group name (Domain Admins, Enterprise Admins,
  Schema Admins, or any other privileged group). Directly useful for
  privileged-account review/attestation (e.g. "show me when each Domain Admin
  last logged in").

  Emits structured report objects to the success stream; -OutputPath writes
  them as CSV (Export-AD.ps1 convention). No elevation needed. Query volume is
  members x domain controllers - fine for privileged groups.

.PARAMETER Group
  Security group(s) to audit. Defaults to Domain Admins, Enterprise Admins,
  and Schema Admins.

.PARAMETER Server
  Optional domain controller to target for the membership and DC-enumeration
  queries. LastLogon is always read from every enumerated DC.

.PARAMETER OutputPath
  Optional CSV file to write the report to.

.EXAMPLE
  PS> ./Get-ADGroupAudit.ps1
  Audits the default privileged groups.

.EXAMPLE
  PS> ./Get-ADGroupAudit.ps1 -Group 'Domain Admins','Backup Operators' -Server dc01.company.com -OutputPath group-audit.csv

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string[]]
  $Group = @('Domain Admins', 'Enterprise Admins', 'Schema Admins'),

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
$failedGroups = 0

foreach ($groupName in $Group) {
  Write-Log -Message "Auditing group: $groupName" -Color Cyan

  $members = @()
  try {
    $members = @(Get-ADGroupMember -Identity $groupName @commonParams -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
  }
  catch {
    $failedGroups++
    Write-Log -Message "  Could not query group '$groupName': $($_.Exception.Message)" -Color Red
    continue
  }

  if ($members.Count -eq 0) {
    Write-Log -Message "  No user members found in '$groupName'." -Color Gray
    continue
  }

  foreach ($member in $members) {
    $maxTicks = [long]0
    foreach ($dc in $domainControllers) {
      try {
        $dcUser = Get-ADUser -Identity $member.DistinguishedName -Properties LastLogon -Server $dc.Name -ErrorAction Stop
        $lastLogonTicks = [long]$dcUser.LastLogon
        if ($lastLogonTicks -gt $maxTicks) {
          $maxTicks = $lastLogonTicks
        }
      }
      catch {
        Write-Log -Message "  Could not query LastLogon for $($member.Name) on $($dc.Name): $($_.Exception.Message)" -Color Gray
      }
    }

    if ($maxTicks -le 0) {
      $lastLogon = $null
      $daysSince = $null
    }
    else {
      $lastLogon = [datetime]::FromFileTime($maxTicks)
      $daysSince = [math]::Round(((Get-Date) - $lastLogon).TotalDays, 1)
    }

    $report += [pscustomobject]@{
      Group = $groupName
      SamAccountName = $member.SamAccountName
      Name = $member.Name
      LastLogon = $lastLogon
      DaysSinceLastLogon = $daysSince
    }
  }
}

$report = @($report | Sort-Object Group, LastLogon -Descending)

if ($report.Count -eq 0) {
  Write-Log -Message 'No members found in the audited group(s).' -Color Yellow
  exit 1
}

$report | Write-Output

if ($OutputPath) {
  $report | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
  Write-Log -Message "Report written to $OutputPath" -Color Gray
}

Write-Log -Message "$($report.Count) member record(s) across $($Group.Count) group(s)." -Color Yellow

if ($failedGroups -gt 0) {
  exit 1
}
exit 0
