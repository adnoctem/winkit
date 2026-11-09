#Requires -Version 5.1
#Requires -Module ActiveDirectory
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.1' }

<#
.SYNOPSIS
  Exports Active Directory users, computers, and groups to CSV or JSON files.

.DESCRIPTION
  Queries the selected AD resources and writes one file per resource:
  AD-Users.csv, AD-Computers.csv, and AD-Groups.csv, or the JSON equivalents.
  The column order is fixed by the Server/ADDS library so the output is stable
  and reviewable.

  Users and computers include enabled accounts only by default. Supply
  -IncludeDisabled to include disabled accounts. Filters, LDAP filters, an OU
  search base, a domain controller, and alternate credentials are supported.

  Missing requested attributes fail the export by default. Use -Omit to remove
  unwanted attributes from the output or -IncludeMissing to export records with
  empty values. CSV and JSON are first-class output formats.

.PARAMETER OutputPath
  Existing directory that receives the export files.

.PARAMETER Resource
  Users, Computers, Groups, or All. Defaults to All.

.PARAMETER Format
  Csv or Json. Defaults to Csv. JSON is written as an array of objects in
  UTF-8 without BOM.

.PARAMETER Encoding
  CSV encoding. Defaults to UTF8. JSON is always UTF-8 without BOM.

.PARAMETER Filter
  Active Directory filter for the selected resources. Defaults to *.

.PARAMETER LDAPFilter
  LDAP filter. Mutually exclusive with an explicitly supplied -Filter.

.PARAMETER SearchBase
  Distinguished name of an OU or container that limits the search.

.PARAMETER Server
  Optional domain controller.

.PARAMETER Credential
  Optional credential for the AD queries.

.PARAMETER Properties
  Additional AD attributes to include after the standard export columns.

.PARAMETER Omit
  Attribute names to exclude from the output. Omitted attributes are also
  excluded from the missing-value check.

.PARAMETER IncludeMissing
  Export records that have missing requested attributes with empty values
  instead of failing.

.PARAMETER IncludeDisabled
  Include disabled users and computers. Groups are unaffected.

.PARAMETER DryRun
  Query and validate without writing files.

.PARAMETER PassThru
  Return one structured operation result per resource.

.EXAMPLE
  PS> ./Export-AD.ps1 -OutputPath C:\ADReports
  Exports enabled users, computers, and groups to CSV files.

.EXAMPLE
  PS> ./Export-AD.ps1 -OutputPath C:\ADReports -Resource Users -Format Json -SearchBase 'OU=Staff,DC=example,DC=com'
  Writes AD-Users.json for the selected OU.

.EXAMPLE
  PS> ./Export-AD.ps1 -OutputPath C:\ADReports -Omit lastLogon,lastLogonTimestamp -IncludeMissing -PassThru
  Exports without the noisy logon columns and allows other attributes to be empty.

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
  [Parameter(Position = 0, Mandatory = $true)]
  [string]
  $OutputPath,

  [Parameter(Position = 1)]
  [ValidateSet('Users', 'Computers', 'Groups', 'All')]
  [string]
  $Resource = 'All',

  [Parameter(Position = 2)]
  [ValidateSet('Csv', 'Json')]
  [string]
  $Format = 'Csv',

  [ValidateSet('UTF8', 'ASCII', 'Unicode', 'UTF7', 'UTF32', 'BigEndianUnicode')]
  [string]
  $Encoding = 'UTF8',

  [string]
  $Filter = '*',

  [string]
  $LDAPFilter,

  [string]
  $SearchBase,

  [string]
  $Server,

  [pscredential]
  $Credential,

  [string[]]
  $Properties,

  [string[]]
  $Omit,

  [switch]
  $IncludeMissing,

  [switch]
  $IncludeDisabled,

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

if ($Format -eq 'Json' -and $PSBoundParameters.ContainsKey('Encoding')) {
  Write-Log -Message 'Encoding applies to CSV output only; JSON is always UTF-8 without BOM.' -Color Yellow
}

if (-not (Test-Path -LiteralPath $OutputPath -PathType Container)) {
  Write-Log -Message "OutputPath directory not found: $OutputPath" -Color Red
  exit 1
}

$resources = @($Resource)
if ($Resource -eq 'All') {
  $resources = @('Users', 'Computers', 'Groups')
}

$results = New-Object System.Collections.Generic.List[object]

try {
  foreach ($currentResource in $resources) {
    Write-Log -Message "Querying AD $currentResource..." -Color Cyan

    $query = @{
      Resource        = $currentResource
      IncludeDisabled = $IncludeDisabled.IsPresent
    }

    if (-not [string]::IsNullOrWhiteSpace($LDAPFilter)) {
      $query.LDAPFilter = $LDAPFilter
    }
    else {
      $query.Filter = $Filter
    }

    if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
      $query.SearchBase = $SearchBase
    }
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
      $query.Server = $Server
    }
    if ($null -ne $Credential) {
      $query.Credential = $Credential
    }
    if ($Properties) {
      $query.Properties = $Properties
    }

    $writer = @{
      Resource       = $currentResource
      OutputPath     = $OutputPath
      Format         = $Format
      IncludeMissing = $IncludeMissing.IsPresent
      PassThru       = $true
    }
    if ($Omit) {
      $writer.Omit = $Omit
    }
    if ($Format -eq 'Csv') {
      $writer.Encoding = $Encoding
    }

    $exportResult = @(Get-ADDirectoryRecord @query | Export-ADDirectoryRecord @writer)
    foreach ($item in $exportResult) {
      $results.Add($item)
      $color = 'Green'
      if ($item.Status -eq 'Skipped' -or $item.Status -eq 'DryRun') {
        $color = 'Yellow'
      }
      Write-Log -Message "AD $($item.Resource): $($item.Status) | Records: $($item.Count) | File: $($item.Path)" -Color $color
    }
  }
}
catch {
  Write-Log -Message "AD export failed: $($_.Exception.Message)" -Color Red
  exit 1
}

if ($PassThru) {
  foreach ($result in $results) {
    $target = $result.Path
    if ([string]::IsNullOrWhiteSpace($target)) {
      $target = "AD-$($result.Resource)"
    }
    New-OperationResult -Target $target -Source 'ADDS' -Action 'Export' -Status $result.Status -Detail $result.Error -Property @{
      Resource = $result.Resource
      Count    = $result.Count
    }
  }
}

if (@($results | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) {
  exit 1
}

exit 0
