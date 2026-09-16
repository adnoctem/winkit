#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Reports AD attributes for a domain user via the Global Catalog.

.DESCRIPTION
  Queries the Global Catalog with a DirectorySearcher (no ActiveDirectory
  module dependency - works from any machine with LDAP reachability) and
  returns a structured report for one user: account identifiers, mail and
  proxy addresses, department/title, manager (with resolved name), direct
  reports, group memberships, and password age with a staleness flag.

  Derived from the domain-user reporting path of Get-Account.ps1 in
  stevencohn/WindowsPowerShell; the reference's presentation-only local-user
  reporting is deliberately not ported. Structured output keeps the report
  pipeable into further automation (e.g. a stale-password report across
  accounts).

  The output object is written to the success stream; -OutputPath additionally
  writes it as JSON. No elevation needed.

.PARAMETER Identity
  The user to report on: sAMAccountName (e.g. jdoe), userPrincipalName, or
  distinguished name.

.PARAMETER Server
  Optional Global Catalog / domain controller to query. Defaults to the
  default Global Catalog of the current domain context.

.PARAMETER PasswordMaxAgeDays
  Password staleness threshold. Defaults to 90 days.

.PARAMETER OutputPath
  Optional JSON file to write the report to.

.EXAMPLE
  PS> ./Get-ADUserReport.ps1 -Identity jdoe
  Reports all covered attributes for jdoe.

.EXAMPLE
  PS> ./Get-ADUserReport.ps1 -Identity 'jdoe@company.com' -Server gc01.company.com -OutputPath jdoe.json

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(
    Position = 0,
    Mandatory = $true,
    HelpMessage = 'sAMAccountName, userPrincipalName, or distinguished name of the user.'
  )]
  [string]
  $Identity,

  [Parameter(Mandatory = $false)]
  [string]
  $Server,

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 3650)]
  [int]
  $PasswordMaxAgeDays = 90,

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath
)

Import-Module PSFoundation -Force

function Convert-ADSLargeInteger {
  <#
    Decodes an AD LargeInteger attribute value (returned as a two-element
    Int32 array from the high/low parts) into an Int64. Tolerates values that
    already come through as Int64.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $Value
  )

  if ($Value -is [array] -and $Value.Count -eq 2) {
    return [int64]([uint32]$Value[1] + ([uint32]$Value[0] * 4294967296))
  }

  try {
    return [int64]$Value
  }
  catch {
    return 0
  }
}

function Convert-FileTime {
  <#
    Converts an AD LargeInteger FILETIME value to a DateTime. Returns $null
    for zero/unset values.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $Value
  )

  $ticks = Convert-ADSLargeInteger -Value $Value
  if ($ticks -le 0) {
    return $null
  }

  try {
    return [datetime]::FromFileTime($ticks)
  }
  catch {
    return $null
  }
}

function ConvertTo-LdapFilterString {
  <#
    Escapes LDAP filter special characters in a user-supplied value.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  $Value = $Value.Replace('\', '\5c')
  $Value = $Value.Replace('*', '\2a')
  $Value = $Value.Replace('(', '\28')
  $Value = $Value.Replace(')', '\29')
  $Value = $Value.Replace([string][char]0, '\00')
  $Value
}

function Get-ADUserReport {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$UserIdentity,

    [string]$GlobalCatalogServer,

    [int]$MaxAgeDays
  )

  $escapedIdentity = ConvertTo-LdapFilterString -Value $UserIdentity
  $filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$escapedIdentity)(userPrincipalName=$escapedIdentity)(distinguishedName=$escapedIdentity)))"

  $searchRoot = if ($GlobalCatalogServer) { "GC://$GlobalCatalogServer" } else { 'GC://' }
  $searcher = New-Object System.DirectoryServices.DirectorySearcher([adsi]$searchRoot)
  $searcher.Filter = $filter
  $searcher.PageSize = 100
  $searcher.SizeLimit = 1
  $searcher.PropertiesToLoad.AddRange(@(
      'sAMAccountName',
      'userPrincipalName',
      'displayName',
      'mail',
      'proxyAddresses',
      'department',
      'title',
      'manager',
      'directReports',
      'memberOf',
      'pwdLastSet'
    ))

  $result = $searcher.FindOne()
  if (-not $result) {
    throw "No user found matching '$UserIdentity'."
  }

  $properties = $result.Properties

  $mailAliases = @()
  $proxyAddresses = @()
  foreach ($proxy in @($properties['proxyAddresses'])) {
    $proxyAddresses += [string]$proxy
    if ([string]$proxy -like 'smtp:*') {
      $mailAliases += [string]$proxy.Substring(5)
    }
  }

  $groups = @()
  foreach ($groupDn in @($properties['memberOf'])) {
    $groups += ($groupDn -split ',')[0].Substring(3)
  }

  $directReports = @()
  foreach ($reportDn in @($properties['directReports'])) {
    $directReports += ($reportDn -split ',')[0].Substring(3)
  }

  $managerDn = [string]$properties['manager']
  $managerName = $null
  if ($managerDn) {
    try {
      $managerEntry = [adsi]"LDAP://$managerDn"
      $managerName = [string]$managerEntry.Properties['displayName'].Value
      if (-not $managerName) {
        $managerName = ($managerDn -split ',')[0].Substring(3)
      }
    }
    catch {
      $managerName = $null
    }
  }

  $passwordLastSet = $null
  if ($properties['pwdLastSet']) {
    $passwordLastSet = Convert-FileTime -Value $properties['pwdLastSet']
  }
  $passwordAgeDays = $null
  if ($passwordLastSet) {
    $passwordAgeDays = [math]::Round(((Get-Date) - $passwordLastSet).TotalDays, 1)
  }

  [pscustomobject]@{
    SamAccountName = [string]$properties['sAMAccountName']
    UserPrincipalName = [string]$properties['userPrincipalName']
    DisplayName = [string]$properties['displayName']
    Mail = [string]$properties['mail']
    MailAliases = @($mailAliases)
    ProxyAddresses = @($proxyAddresses)
    Department = [string]$properties['department']
    Title = [string]$properties['title']
    Manager = $managerDn
    ManagerName = $managerName
    DirectReports = @($directReports)
    DirectReportCount = $directReports.Count
    Groups = @($groups)
    GroupCount = $groups.Count
    PasswordLastSet = $passwordLastSet
    PasswordAgeDays = $passwordAgeDays
    PasswordStale = (($null -ne $passwordAgeDays) -and ($passwordAgeDays -gt $MaxAgeDays))
  }
}

try {
  $report = Get-ADUserReport -UserIdentity $Identity -GlobalCatalogServer $Server -MaxAgeDays $PasswordMaxAgeDays

  if ($report.PasswordStale) {
    Write-Log -Message "Password for $($report.SamAccountName) is stale ($($report.PasswordAgeDays) days, threshold $PasswordMaxAgeDays)." -Color Yellow
  }

  $report | Write-Output

  if ($OutputPath) {
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
    Write-Log -Message "Report written to $OutputPath" -Color Gray
  }
}
catch {
  Write-Log -Message "AD user report failed: $($_.Exception.Message)" -Color Red
  exit 1
}

exit 0
