#Requires -Version 5.1
#Requires -Module ActiveDirectory
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.1' }

<#
.SYNOPSIS
  Generates the AD account map consumed by the plead mail tool.

.DESCRIPTION
  Resolves selected AD users to full email addresses and writes an
  ad-account-map.json object of samAccountName to email. The AD mail attribute
  is authoritative; an explicit override CSV can replace it for accounts whose
  mail attribute is unreliable. No address is guessed from a domain name.

  Only exact allowlisted domains are accepted. Disabled users are excluded by
  default; supply -IncludeDisabled to include them. Missing mail fails the run
  unless -IncludeMissing is supplied, in which case affected accounts are
  reported and omitted from the map.

  The map is UTF-8 without BOM with deterministic key order. An optional CSV
  audit report records every selected account and the reason for any omission.
  Existing files are only replaced with -Force.

.PARAMETER OutputPath
  Map file. Defaults to ad-account-map.json in the current directory.

.PARAMETER ReportPath
  Optional CSV audit report.

.PARAMETER SearchBase
  Distinguished name of an OU or container that limits the search.

.PARAMETER Filter
  Active Directory filter. Defaults to *.

.PARAMETER LDAPFilter
  LDAP filter. Mutually exclusive with an explicitly supplied -Filter.

.PARAMETER Group
  Group whose user members are selected.

.PARAMETER Recursive
  Expand nested group membership.

.PARAMETER Identity
  Explicit user identities. Cannot be combined with Group, SearchBase, or a
  filter.

.PARAMETER AllowedDomain
  Exact email-domain allowlist. ASCII/Punycode only; Unicode domain spellings
  are rejected.

.PARAMETER OverridePath
  Optional UTF-8 CSV with samAccountName, email, and optional note headers. An
  empty email marks an intentional exclusion.

.PARAMETER OnDuplicate
  Duplicate samAccountName policy: Fail, First, or Last.

.PARAMETER AllowOverrideOutsideDomain
  Permit an explicit override whose domain is outside AllowedDomain. Every use
  is recorded in the audit report.

.PARAMETER IncludeUnmatchedOverride
  Include override rows that do not match a selected AD account. By default
  they are reported and excluded.

.PARAMETER IncludeMissing
  Omit and report accounts with missing required values instead of failing.

.PARAMETER IncludeDisabled
  Include disabled users. By default only enabled users are selected.

.PARAMETER Server
  Optional domain controller.

.PARAMETER Credential
  Optional credential for the AD queries.

.PARAMETER Force
  Replace existing output or report files.

.PARAMETER DryRun
  Resolve and preview without writing files.

.PARAMETER PassThru
  Return the typed account-map result.

.EXAMPLE
  PS> ./Export-ADAccountMap.ps1 -AllowedDomain adnoctem.co,example.com
  Writes ad-account-map.json in the current directory.

.EXAMPLE
  PS> ./Export-ADAccountMap.ps1 -AllowedDomain adnoctem.co,example.com -OverridePath C:\ADReports\overrides.csv -OutputPath C:\ADReports\ad-account-map.json -ReportPath C:\ADReports\ad-account-map-audit.csv -DryRun
  Previews the map and the audit report without writing either file.

.EXAMPLE
  PS> ./Export-ADAccountMap.ps1 -AllowedDomain adnoctem.co,example.com -SearchBase 'OU=Users,DC=adnoctem,DC=co' -IncludeMissing -Force
  Regenerates the map for one OU, allowing accounts without mail to be omitted and reported.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported when the ActiveDirectory PowerShell module and AD Web Services connectivity are available.
  SYSTEM-account execution: supported when the module and network/AD permissions are available; no user profile is required.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [string]
  $OutputPath = 'ad-account-map.json',

  [string]
  $ReportPath,

  [string]
  $SearchBase,

  [string]
  $Filter = '*',

  [string]
  $LDAPFilter,

  [string]
  $Group,

  [switch]
  $Recursive,

  [string[]]
  $Identity,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string[]]
  $AllowedDomain,

  [string]
  $OverridePath,

  [ValidateSet('Fail', 'First', 'Last')]
  [string]
  $OnDuplicate = 'Fail',

  [switch]
  $AllowOverrideOutsideDomain,

  [switch]
  $IncludeUnmatchedOverride,

  [switch]
  $IncludeMissing,

  [switch]
  $IncludeDisabled,

  [string]
  $Server,

  [pscredential]
  $Credential,

  [switch]
  $Force,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force

. (Join-Path -Path $PSScriptRoot -ChildPath 'lib/ADDS.ps1')

if ($DryRun) {
  $WhatIfPreference = $true
}

if ($PSBoundParameters.ContainsKey('Filter') -and $PSBoundParameters.ContainsKey('LDAPFilter')) {
  Write-Log -Message 'Use either -Filter or -LDAPFilter, not both.' -Color Red
  exit 1
}

if ($Identity -and ($Group -or $LDAPFilter -or $SearchBase -or $PSBoundParameters.ContainsKey('Filter'))) {
  Write-Log -Message 'Identity cannot be combined with Group, SearchBase, Filter, or LDAPFilter.' -Color Red
  exit 1
}

if ($Group -and ($LDAPFilter -or $SearchBase -or $PSBoundParameters.ContainsKey('Filter'))) {
  Write-Log -Message 'Group cannot be combined with SearchBase, Filter, or LDAPFilter.' -Color Red
  exit 1
}

$outputFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputDirectory = Split-Path -Path $outputFullPath -Parent
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
  Write-Log -Message "Output directory not found: $outputDirectory" -Color Red
  exit 1
}

if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
  $reportFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
  $reportDirectory = Split-Path -Path $reportFullPath -Parent
  if (-not (Test-Path -LiteralPath $reportDirectory -PathType Container)) {
    Write-Log -Message "Report directory not found: $reportDirectory" -Color Red
    exit 1
  }
}

$mapParams = @{
  AllowedDomain              = $AllowedDomain
  OnDuplicate                = $OnDuplicate
  AllowOverrideOutsideDomain = $AllowOverrideOutsideDomain.IsPresent
  IncludeUnmatchedOverride   = $IncludeUnmatchedOverride.IsPresent
  IncludeMissing             = $IncludeMissing.IsPresent
  IncludeDisabled            = $IncludeDisabled.IsPresent
}

if (-not [string]::IsNullOrWhiteSpace($OverridePath)) {
  $mapParams.OverridePath = $OverridePath
}
if (-not [string]::IsNullOrWhiteSpace($Server)) {
  $mapParams.Server = $Server
}
if ($null -ne $Credential) {
  $mapParams.Credential = $Credential
}

if ($Identity) {
  $mapParams.Identity = $Identity
}
elseif ($Group) {
  $mapParams.Group = $Group
  if ($Recursive) {
    $mapParams.Recursive = $true
  }
}
elseif (-not [string]::IsNullOrWhiteSpace($LDAPFilter)) {
  $mapParams.LDAPFilter = $LDAPFilter
  if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
    $mapParams.SearchBase = $SearchBase
  }
}
else {
  $mapParams.Filter = $Filter
  if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
    $mapParams.SearchBase = $SearchBase
  }
}

try {
  Write-Log -Message 'Resolving AD account map...' -Color Cyan
  $mapResult = New-ADAccountMap @mapParams

  Write-Log -Message "Selected: $($mapResult.EntryCount) | Resolved: $($mapResult.ResolvedCount) | Unresolved: $($mapResult.UnresolvedCount) | Excluded: $(@($mapResult.Excluded).Count) | Duplicates: $(@($mapResult.Duplicates).Count)" -Color Cyan

  $exportParams = @{
    InputObject = $mapResult
    OutputPath  = $outputFullPath
    Force       = $Force.IsPresent
    DryRun      = $DryRun.IsPresent
    PassThru    = $true
  }
  if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $exportParams.ReportPath = $reportFullPath
  }

  $exportedResult = Export-ADAccountMap @exportParams

  if ($exportedResult.Status -eq 'Failed') {
    $reasons = @($exportedResult.Unresolved | ForEach-Object { "$($_.SamAccountName): $($_.Reason)" })
    if ($reasons.Count -gt 0) {
      Write-Log -Message "Unresolved accounts: $($reasons -join ', ')" -Color Red
    }
    Write-Log -Message "Account map not written: $($exportedResult.Errors -join ' ')" -Color Red
    if ($PassThru -or $DryRun) {
      $exportedResult
    }
    exit 1
  }

  $color = 'Green'
  if ($exportedResult.Status -eq 'DryRun') {
    $color = 'Yellow'
  }
  Write-Log -Message "Account map: $($exportedResult.Status) | Entries: $($exportedResult.ResolvedCount) | File: $($exportedResult.Path)" -Color $color
  if ($exportedResult.ReportPath) {
    Write-Log -Message "Audit report: $($exportedResult.ReportPath)" -Color Cyan
  }
  if ($exportedResult.Sha256) {
    Write-Log -Message "SHA256: $($exportedResult.Sha256)" -Color Gray
  }

  if ($PassThru -or $DryRun) {
    $exportedResult
  }
}
catch {
  Write-Log -Message "Account map generation failed: $($_.Exception.Message)" -Color Red
  exit 1
}

exit 0
