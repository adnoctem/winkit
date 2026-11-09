#Requires -Version 5.1

<#
.SYNOPSIS
  Shared Active Directory directory-record and account-map functions.

.DESCRIPTION
  Dot-source this library from the Server/ADDS entry scripts. It contains no
  workflow, no elevation requirement, and no file writes outside the explicit
  export functions. Callers own logging, exit codes, and preview policy.

  The library returns typed PSCustomObject records named Winkit.ADDS.User,
  Winkit.ADDS.Computer, and Winkit.ADDS.Group. Account-map results use
  Winkit.ADDS.AccountMapEntry and Winkit.ADDS.AccountMapResult.
#>

function Get-ADResourceProperty {
  <#
  .SYNOPSIS
    Returns the stable export property list for one AD resource.

  .DESCRIPTION
    Defines the column order used by the CSV and JSON exports and the AD
    attributes requested by Get-ADDirectoryRecord. The list is authoritative:
    changing it changes the documented export schema.

  .PARAMETER Resource
    Resource whose property list is returned.

  .OUTPUTS
    System.String[]
  #>
  [CmdletBinding()]
  [OutputType([string[]])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('Users', 'Computers', 'Groups')]
    [string]
    $Resource
  )

  switch ($Resource) {
    'Users' {
      return @(
        'DistinguishedName',
        'Enabled',
        'Surname',
        'UserPrincipalName',
        'sAMAccountName',
        'SID',
        'Name',
        'ObjectClass',
        'ObjectGuid',
        'lastLogon',
        'lastLogonTimestamp',
        'whenCreated',
        'whenChanged',
        'accountExpires',
        'homeDirectory',
        'homeDrive',
        'profilePath',
        'cn',
        'sn',
        'givenName',
        'displayName',
        'mail',
        'mailNickname',
        'description',
        'telephoneNumber'
      )
    }

    'Computers' {
      return @(
        'DistinguishedName',
        'Enabled',
        'DNSHostName',
        'UserPrincipalName',
        'sAMAccountName',
        'SID',
        'Name',
        'ObjectClass',
        'ObjectGuid',
        'cn',
        'servicePrincipalName',
        'lastLogon',
        'lastLogonTimestamp',
        'whenCreated',
        'whenChanged',
        'accountExpires',
        'operatingSystem',
        'operatingSystemVersion'
      )
    }

    'Groups' {
      return @(
        'DistinguishedName',
        'GroupScope',
        'GroupCategory',
        'sAMAccountName',
        'SID',
        'Name',
        'ObjectClass',
        'ObjectGuid',
        'cn',
        'description',
        'whenCreated',
        'whenChanged',
        'groupType'
      )
    }
  }
}

function Get-ADAttributeValue {
  <#
  .SYNOPSIS
    Reads one attribute from an AD object or a typed record.

  .DESCRIPTION
    PowerShell property lookup is case-insensitive, but this helper also
    tolerates a null object and an exact-name property stored on a typed
    record. Missing attributes return $null.

  .PARAMETER InputObject
    AD object or typed record to inspect.

  .PARAMETER Name
    Attribute name to read.

  .OUTPUTS
    System.Object
  #>
  [CmdletBinding()]
  [OutputType([object])]
  param (
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [psobject]
    $InputObject,

    [Parameter(Mandatory = $true)]
    [string]
    $Name
  )

  if ($null -eq $InputObject) {
    return $null
  }

  $property = $InputObject.PSObject.Properties[$Name]
  if ($null -ne $property) {
    return $property.Value
  }

  foreach ($candidate in $InputObject.PSObject.Properties) {
    if ($candidate.Name -ieq $Name) {
      return $candidate.Value
    }
  }

  return $null
}

function Get-ADRecordAttribute {
  <#
  .SYNOPSIS
    Reads an attribute from a typed directory record or an AD-like object.

  .PARAMETER Record
    Typed record or AD object.

  .PARAMETER Name
    Attribute name to read.

  .OUTPUTS
    System.Object
  #>
  [CmdletBinding()]
  [OutputType([object])]
  param (
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [psobject]
    $Record,

    [Parameter(Mandatory = $true)]
    [string]
    $Name
  )

  if ($null -eq $Record) {
    return $null
  }

  $propertyBag = $Record.PSObject.Properties['Properties']
  if ($null -ne $propertyBag -and $propertyBag.Value -is [System.Collections.IDictionary]) {
    $value = $propertyBag.Value[$Name]
    if ($null -ne $value) {
      return $value
    }
  }

  return Get-ADAttributeValue -InputObject $Record -Name $Name
}

function New-ADDirectoryRecord {
  <#
  .SYNOPSIS
    Converts one AD object into a typed directory record.

  .PARAMETER InputObject
    AD object returned by Get-ADUser, Get-ADComputer, or Get-ADGroup.

  .PARAMETER Resource
    Resource type of the object.

  .PARAMETER Property
    Ordered attributes to copy into the record property bag.

  .OUTPUTS
    Winkit.ADDS.User, Winkit.ADDS.Computer, or Winkit.ADDS.Group
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory record only; no system state changes.')]
  [CmdletBinding()]
  [OutputType('Winkit.ADDS.User', 'Winkit.ADDS.Computer', 'Winkit.ADDS.Group')]
  param (
    [Parameter(Mandatory = $true)]
    [psobject]
    $InputObject,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Users', 'Computers', 'Groups')]
    [string]
    $Resource,

    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [string[]]
    $Property
  )

  $properties = [ordered]@{}
  foreach ($name in $Property) {
    $properties[$name] = Get-ADAttributeValue -InputObject $InputObject -Name $name
  }

  $typeName = switch ($Resource) {
    'Users' { 'Winkit.ADDS.User' }
    'Computers' { 'Winkit.ADDS.Computer' }
    'Groups' { 'Winkit.ADDS.Group' }
  }

  $objectGuidValue = Get-ADAttributeValue -InputObject $InputObject -Name 'ObjectGuid'
  $objectGuid = $null
  if ($null -ne $objectGuidValue -and -not [string]::IsNullOrWhiteSpace([string]$objectGuidValue)) {
    try {
      $objectGuid = [guid]$objectGuidValue
    }
    catch {
      $objectGuid = $null
    }
  }

  $record = [ordered]@{
    PSTypeName        = $typeName
    Resource          = $Resource
    DistinguishedName = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'DistinguishedName')
    SamAccountName    = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'sAMAccountName')
    Name              = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'Name')
    ObjectGuid        = $objectGuid
    Properties        = $properties
  }

  switch ($Resource) {
    'Users' {
      $record.UserPrincipalName = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'UserPrincipalName')
      $record.Mail = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'mail')
      $record.Enabled = [bool](Get-ADAttributeValue -InputObject $InputObject -Name 'Enabled')
    }

    'Computers' {
      $record.DNSHostName = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'DNSHostName')
      $record.Enabled = [bool](Get-ADAttributeValue -InputObject $InputObject -Name 'Enabled')
    }

    'Groups' {
      $record.GroupScope = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'GroupScope')
      $record.GroupCategory = [string](Get-ADAttributeValue -InputObject $InputObject -Name 'GroupCategory')
    }
  }

  return [PSCustomObject]$record
}

function Get-ADRawDirectoryRecord {
  <#
  .SYNOPSIS
    Runs one AD query for a resource and returns raw AD objects.

  .PARAMETER Resource
    Resource to query.

  .PARAMETER Identity
    Optional explicit identity. When set, Filter and LDAPFilter are ignored.

  .PARAMETER Filter
    Optional AD filter.

  .PARAMETER LDAPFilter
    Optional LDAP filter.

  .PARAMETER SearchBase
    Optional search base.

  .PARAMETER Server
    Optional domain controller.

  .PARAMETER Credential
    Optional credential.

  .PARAMETER Property
    Attributes to request.

  .OUTPUTS
    Microsoft.ActiveDirectory.Management.ADObject
  #>
  [CmdletBinding()]
  [OutputType([object])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('Users', 'Computers', 'Groups')]
    [string]
    $Resource,

    [string]
    $Identity,

    [string]
    $Filter,

    [string]
    $LDAPFilter,

    [string]
    $SearchBase,

    [string]
    $Server,

    [pscredential]
    $Credential,

    [string[]]
    $Property
  )

  $query = @{
    ErrorAction = 'Stop'
  }

  if ($Property) {
    $query.Properties = $Property
  }

  if (-not [string]::IsNullOrWhiteSpace($Identity)) {
    $query.Identity = $Identity
  }
  elseif (-not [string]::IsNullOrWhiteSpace($LDAPFilter)) {
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

  switch ($Resource) {
    'Users' { return Get-ADUser @query }
    'Computers' { return Get-ADComputer @query }
    'Groups' { return Get-ADGroup @query }
  }
}

function Write-ADDirectoryRecordQuery {
  <#
  .SYNOPSIS
    Executes one AD directory query and emits typed records.

  .PARAMETER Resource
    Resource to query.

  .PARAMETER Property
    Ordered properties to copy into each record.

  .PARAMETER Identity
    Optional explicit identities.

  .PARAMETER Filter
    Optional AD filter.

  .PARAMETER LDAPFilter
    Optional LDAP filter.

  .PARAMETER SearchBase
    Optional search base.

  .PARAMETER Group
    Optional group whose members are selected.

  .PARAMETER Recursive
    Expand nested group membership.

  .PARAMETER Server
    Optional domain controller.

  .PARAMETER Credential
    Optional credential.

  .PARAMETER IncludeDisabled
    Include disabled users and computers.

  .OUTPUTS
    Winkit.ADDS.User, Winkit.ADDS.Computer, or Winkit.ADDS.Group
  #>
  [CmdletBinding()]
  [OutputType('Winkit.ADDS.User', 'Winkit.ADDS.Computer', 'Winkit.ADDS.Group')]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('Users', 'Computers', 'Groups')]
    [string]
    $Resource,

    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [string[]]
    $Property,

    [string[]]
    $Identity,

    [string]
    $Filter,

    [string]
    $LDAPFilter,

    [string]
    $SearchBase,

    [string]
    $Group,

    [switch]
    $Recursive,

    [string]
    $Server,

    [pscredential]
    $Credential,

    [switch]
    $IncludeDisabled
  )

  $rawRecords = @()

  if (-not [string]::IsNullOrWhiteSpace($Group)) {
    $memberQuery = @{
      Identity    = $Group
      ErrorAction = 'Stop'
    }
    if ($Recursive) {
      $memberQuery.Recursive = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
      $memberQuery.Server = $Server
    }
    if ($null -ne $Credential) {
      $memberQuery.Credential = $Credential
    }

    $members = @(Get-ADGroupMember @memberQuery)
    foreach ($member in $members) {
      $objectClass = @($member.objectClass)
      $expectedClass = switch ($Resource) {
        'Users' { 'user' }
        'Computers' { 'computer' }
        'Groups' { 'group' }
      }

      if (-not ($objectClass -contains $expectedClass)) {
        Write-Verbose "Skipping group member '$($member.DistinguishedName)' because it is not a $expectedClass."
        continue
      }

      $query = @{
        Resource = $Resource
        Identity = [string]$member.DistinguishedName
        Property = $Property
      }
      if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $query.Server = $Server
      }
      if ($null -ne $Credential) {
        $query.Credential = $Credential
      }
      $rawRecords += Get-ADRawDirectoryRecord @query
    }
  }
  elseif ($null -ne $Identity -and $Identity.Count -gt 0) {
    foreach ($identityValue in $Identity) {
      $query = @{
        Resource = $Resource
        Identity = $identityValue
        Property = $Property
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
      $rawRecords += Get-ADRawDirectoryRecord @query
    }
  }
  else {
    $query = @{
      Resource = $Resource
      Property = $Property
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
    $rawRecords = @(Get-ADRawDirectoryRecord @query)
  }

  foreach ($rawRecord in $rawRecords) {
    if (-not $IncludeDisabled -and $Resource -ne 'Groups') {
      $enabled = Get-ADAttributeValue -InputObject $rawRecord -Name 'Enabled'
      if (-not [bool]$enabled) {
        Write-Verbose "Skipping disabled $Resource record '$([string](Get-ADAttributeValue -InputObject $rawRecord -Name 'sAMAccountName'))'."
        continue
      }
    }

    New-ADDirectoryRecord -InputObject $rawRecord -Resource $Resource -Property $Property
  }
}

function Get-ADDirectoryRecord {
  <#
  .SYNOPSIS
    Queries AD users, computers, or groups and returns typed records.

  .DESCRIPTION
    Supports filtering, a search base, explicit identities from the pipeline,
    and group membership. Disabled users and computers are excluded unless
    -IncludeDisabled is supplied. All AD calls use terminating errors.

  .PARAMETER Resource
    Resource to query. All is valid only with the Filter and LDAPFilter sets.

  .PARAMETER Filter
    AD filter used by the default parameter set.

  .PARAMETER LDAPFilter
    LDAP filter. Mutually exclusive with Filter.

  .PARAMETER SearchBase
    OU or container to search.

  .PARAMETER Identity
    One or more identities from the pipeline.

  .PARAMETER Group
    Group whose members are returned.

  .PARAMETER Recursive
    Expand nested group membership.

  .PARAMETER Server
    Optional domain controller.

  .PARAMETER Credential
    Optional credential.

  .PARAMETER Properties
    Additional attributes to request.

  .PARAMETER IncludeDisabled
    Include disabled users and computers. Groups are unaffected.

  .OUTPUTS
    Winkit.ADDS.User, Winkit.ADDS.Computer, or Winkit.ADDS.Group
  #>
  [CmdletBinding(DefaultParameterSetName = 'Filter')]
  [OutputType('Winkit.ADDS.User', 'Winkit.ADDS.Computer', 'Winkit.ADDS.Group')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('Users', 'Computers', 'Groups', 'All')]
    [string]
    $Resource,

    [Parameter(ParameterSetName = 'Filter')]
    [string]
    $Filter = '*',

    [Parameter(Mandatory = $true, ParameterSetName = 'LDAPFilter')]
    [string]
    $LDAPFilter,

    [Parameter(ParameterSetName = 'Filter')]
    [Parameter(ParameterSetName = 'LDAPFilter')]
    [string]
    $SearchBase,

    [Parameter(
      Mandatory = $true,
      ParameterSetName = 'Identity',
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true
    )]
    [Alias('SamAccountName', 'DistinguishedName', 'ObjectGUID', 'Name')]
    [string[]]
    $Identity,

    [Parameter(Mandatory = $true, ParameterSetName = 'Group')]
    [string]
    $Group,

    [Parameter(ParameterSetName = 'Group')]
    [switch]
    $Recursive,

    [Parameter(ParameterSetName = 'Filter')]
    [Parameter(ParameterSetName = 'LDAPFilter')]
    [Parameter(ParameterSetName = 'Identity')]
    [Parameter(ParameterSetName = 'Group')]
    [string]
    $Server,

    [pscredential]
    $Credential,

    [string[]]
    $Properties,

    [switch]
    $IncludeDisabled
  )

  begin {
    if ($Resource -eq 'All' -and $PSCmdlet.ParameterSetName -ne 'Filter' -and $PSCmdlet.ParameterSetName -ne 'LDAPFilter') {
      throw 'Resource All can only be used with -Filter or -LDAPFilter. Select Users, Computers, or Groups for identity and group queries.'
    }

    $resources = @($Resource)
    if ($Resource -eq 'All') {
      $resources = @('Users', 'Computers', 'Groups')
    }

    $propertyByResource = @{}
    foreach ($currentResource in $resources) {
      $property = @(Get-ADResourceProperty -Resource $currentResource)
      if ($Properties) {
        foreach ($name in $Properties) {
          if (-not ($property -contains $name)) {
            $property += $name
          }
        }
      }
      $propertyByResource[$currentResource] = $property
    }
  }

  process {
    if ($PSCmdlet.ParameterSetName -eq 'Identity') {
      foreach ($currentResource in $resources) {
        Write-ADDirectoryRecordQuery `
          -Resource $currentResource `
          -Property $propertyByResource[$currentResource] `
          -Identity $Identity `
          -Server $Server `
          -Credential $Credential `
          -IncludeDisabled:$IncludeDisabled
      }
    }
  }

  end {
    if ($PSCmdlet.ParameterSetName -eq 'Identity') {
      return
    }

    foreach ($currentResource in $resources) {
      $query = @{
        Resource        = $currentResource
        Property        = $propertyByResource[$currentResource]
        IncludeDisabled = $IncludeDisabled.IsPresent
      }

      if ($PSCmdlet.ParameterSetName -eq 'Group') {
        $query.Group = $Group
        if ($Recursive) {
          $query.Recursive = $true
        }
      }
      elseif ($PSCmdlet.ParameterSetName -eq 'LDAPFilter') {
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

      Write-ADDirectoryRecordQuery @query
    }
  }
}

function Test-ADMissingValue {
  <#
  .SYNOPSIS
    Tests whether an export value counts as missing.

  .PARAMETER Value
    Value to test.

  .OUTPUTS
    System.Boolean
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param (
    [AllowNull()]
    [object]
    $Value
  )

  if ($null -eq $Value) {
    return $true
  }

  if ($Value -is [string]) {
    return [string]::IsNullOrWhiteSpace($Value)
  }

  if ($Value -is [System.Collections.ICollection]) {
    return ($Value.Count -eq 0)
  }

  return $false
}

function Export-ADDirectoryRecord {
  <#
  .SYNOPSIS
    Writes typed directory records to one CSV or JSON file.

  .DESCRIPTION
    Writes AD-Users.csv, AD-Computers.csv, or AD-Groups.csv (or the JSON
    equivalents) with a stable property order. Missing requested attributes
    terminate the export unless -IncludeMissing is supplied. -Omit removes
    attributes from the output and from the missing-value check.

  .PARAMETER InputObject
    Typed records from Get-ADDirectoryRecord.

  .PARAMETER Resource
    Resource contained in the input records.

  .PARAMETER OutputPath
    Existing directory that receives the file.

  .PARAMETER Format
    Csv or Json.

  .PARAMETER Encoding
    CSV encoding. JSON is always UTF-8 without BOM.

  .PARAMETER Omit
    Attribute names to exclude from the output.

  .PARAMETER IncludeMissing
    Export records that have missing requested attributes instead of failing.

  .PARAMETER PassThru
    Return a typed export result.

  .OUTPUTS
    Winkit.ADDS.ExportResult
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  [OutputType('Winkit.ADDS.ExportResult')]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [psobject[]]
    $InputObject,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Users', 'Computers', 'Groups')]
    [string]
    $Resource,

    [Parameter(Mandatory = $true, Position = 0)]
    [string]
    $OutputPath,

    [ValidateSet('Csv', 'Json')]
    [string]
    $Format = 'Csv',

    [ValidateSet('UTF8', 'ASCII', 'Unicode', 'UTF7', 'UTF32', 'BigEndianUnicode')]
    [string]
    $Encoding = 'UTF8',

    [string[]]
    $Omit,

    [switch]
    $IncludeMissing,

    [switch]
    $PassThru
  )

  begin {
    if (-not (Test-Path -LiteralPath $OutputPath -PathType Container)) {
      throw "OutputPath directory not found: $OutputPath"
    }

    $records = New-Object System.Collections.Generic.List[object]
  }

  process {
    foreach ($record in $InputObject) {
      if ($null -ne $record) {
        $records.Add($record)
      }
    }
  }

  end {
    $result = [PSCustomObject]@{
      PSTypeName = 'Winkit.ADDS.ExportResult'
      Resource   = $Resource
      Path       = $null
      Count      = 0
      Status     = 'Skipped'
      Error      = $null
    }

    if ($records.Count -eq 0) {
      Write-Warning "No $Resource records were returned; no file was written."
      if ($PassThru) {
        return $result
      }
      return
    }

    $columnNames = New-Object System.Collections.Generic.List[string]
    foreach ($name in (Get-ADResourceProperty -Resource $Resource)) {
      if (-not ($columnNames -contains $name)) {
        $columnNames.Add($name)
      }
    }
    foreach ($record in $records) {
      if ($null -ne $record.Properties -and $record.Properties -is [System.Collections.IDictionary]) {
        foreach ($name in $record.Properties.Keys) {
          if (-not ($columnNames -contains [string]$name)) {
            $columnNames.Add([string]$name)
          }
        }
      }
    }

    $omitted = New-Object System.Collections.Generic.List[string]
    if ($Omit) {
      foreach ($name in $Omit) {
        $trimmed = $name.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
          continue
        }

        $match = $null
        foreach ($column in $columnNames) {
          if ($column -ieq $trimmed) {
            $match = $column
            break
          }
        }

        if ($null -eq $match) {
          throw "Omitted property '$trimmed' is not part of the $Resource export schema."
        }

        if (-not ($omitted -contains $match)) {
          $omitted.Add($match)
        }
      }
    }

    $columns = @($columnNames | Where-Object { $omitted -notcontains $_ })
    if ($columns.Count -eq 0) {
      throw "All properties for $Resource were omitted; at least one output property is required."
    }

    $missingRecords = New-Object System.Collections.Generic.List[object]
    foreach ($record in $records) {
      $missing = @()
      foreach ($column in $columns) {
        $value = $null
        if ($null -ne $record.Properties -and $record.Properties -is [System.Collections.IDictionary]) {
          $value = $record.Properties[$column]
        }
        if (Test-ADMissingValue -Value $value) {
          $missing += $column
        }
      }

      if ($missing.Count -gt 0) {
        $missingRecords.Add([PSCustomObject]@{
            Record  = $record
            Missing = $missing
          })
      }
    }

    if ($missingRecords.Count -gt 0 -and -not $IncludeMissing) {
      $details = foreach ($entry in $missingRecords) {
        $sam = [string](Get-ADRecordAttribute -Record $entry.Record -Name 'sAMAccountName')
        "$sam ($($entry.Missing -join ', '))"
      }
      throw "Missing requested properties for $Resource. Use -Omit to exclude attributes or -IncludeMissing to export them with empty values: $($details -join '; ')."
    }

    if ($missingRecords.Count -gt 0) {
      foreach ($entry in $missingRecords) {
        $sam = [string](Get-ADRecordAttribute -Record $entry.Record -Name 'sAMAccountName')
        Write-Warning "Record '$sam' is missing: $($entry.Missing -join ', '). Exporting with empty values because -IncludeMissing was supplied."
      }
    }

    $extension = 'csv'
    if ($Format -eq 'Json') {
      $extension = 'json'
    }

    $file = Join-Path -Path $OutputPath -ChildPath "AD-$Resource.$extension"
    $result.Path = $file
    $result.Count = $records.Count

    if (-not $PSCmdlet.ShouldProcess($file, "Export AD $Resource records")) {
      $result.Status = 'DryRun'
      if ($PassThru) {
        return $result
      }
      return
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($record in $records) {
      $row = [ordered]@{}
      foreach ($column in $columns) {
        $value = $null
        if ($null -ne $record.Properties -and $record.Properties -is [System.Collections.IDictionary]) {
          $value = $record.Properties[$column]
        }
        $row[$column] = $value
      }
      $rows.Add([PSCustomObject]$row)
    }

    try {
      if ($Format -eq 'Json') {
        $json = ConvertTo-Json -InputObject $rows.ToArray() -Depth 6
        $json = $json -replace "`r?`n", "`r`n"
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [IO.File]::WriteAllText($file, ($json + "`r`n"), $utf8)
      }
      else {
        $rows | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding $Encoding -Delimiter ';'
      }
    }
    catch {
      $result.Status = 'Failed'
      $result.Error = $_.Exception.Message
      throw
    }

    $result.Status = 'Completed'
    if ($PassThru) {
      return $result
    }
  }
}

function Test-ADMailAddress {
  <#
  .SYNOPSIS
    Tests whether a value is a valid, bare email address.

  .DESCRIPTION
    Rejects empty values, whitespace, display-name forms, multiple addresses,
    overlong values, and Unicode domain spellings. When -AllowedDomain is
    supplied, the domain must match one entry exactly and case-insensitively.

  .PARAMETER Address
    Address to validate.

  .PARAMETER AllowedDomain
    Optional exact domain allowlist.

  .OUTPUTS
    System.Boolean
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [AllowEmptyString()]
    [string]
    $Address,

    [string[]]
    $AllowedDomain
  )

  process {
    if ([string]::IsNullOrWhiteSpace($Address) -or $Address.Length -gt 254 -or $Address -match '\s') {
      return $false
    }

    $value = $Address.Trim()

    try {
      $parsed = New-Object System.Net.Mail.MailAddress($value)
    }
    catch {
      return $false
    }

    if ($parsed.Address -cne $value) {
      return $false
    }

    if ([string]::IsNullOrWhiteSpace($parsed.User) -or [string]::IsNullOrWhiteSpace($parsed.Host)) {
      return $false
    }

    if ($parsed.Host -match '[^\x00-\x7F]') {
      return $false
    }

    if ($null -ne $AllowedDomain -and $AllowedDomain.Count -gt 0) {
      $allowed = $false
      foreach ($domain in $AllowedDomain) {
        if ([string]::Equals($parsed.Host, $domain.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
          $allowed = $true
          break
        }
      }

      if (-not $allowed) {
        return $false
      }
    }

    return $true
  }
}

function Import-ADAccountMapOverride {
  <#
  .SYNOPSIS
    Reads the account-map override CSV.

  .DESCRIPTION
    Reads a UTF-8 CSV with case-insensitive samAccountName, email, and optional
    note headers. An empty email marks an intentional exclusion. Duplicate
    account keys and malformed addresses terminate with the offending row.

  .PARAMETER Path
    Override CSV path.

  .OUTPUTS
    Winkit.ADDS.AccountMapOverride
  #>
  [CmdletBinding()]
  [OutputType('Winkit.ADDS.AccountMapOverride')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]
    $Path
  )

  $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
  if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
    throw "Override file not found: $resolvedPath"
  }

  $text = [System.IO.File]::ReadAllText($resolvedPath)
  $lines = @($text -split "`r?`n")
  if ($lines.Count -eq 0 -or [string]::IsNullOrWhiteSpace($lines[0])) {
    throw "Override file is empty: $resolvedPath"
  }

  $headerNames = @($lines[0] -split ',' | ForEach-Object { $_.Trim().Trim('"') })
  $keyHeader = $null
  $emailHeader = $null
  $noteHeader = $null
  foreach ($headerName in $headerNames) {
    if ($headerName -ieq 'samAccountName') {
      $keyHeader = $headerName
    }
    elseif ($headerName -ieq 'email') {
      $emailHeader = $headerName
    }
    elseif ($headerName -ieq 'note') {
      $noteHeader = $headerName
    }
  }

  if ($null -eq $keyHeader -or $null -eq $emailHeader) {
    throw "Override file must contain samAccountName and email headers: $resolvedPath"
  }

  $rows = @($lines | ConvertFrom-Csv)
  $seen = @{}
  $lineNumber = 1

  foreach ($row in $rows) {
    $lineNumber++
    $key = [string](Get-ADAttributeValue -InputObject $row -Name $keyHeader)
    $email = [string](Get-ADAttributeValue -InputObject $row -Name $emailHeader)
    $note = $null
    if ($null -ne $noteHeader) {
      $note = [string](Get-ADAttributeValue -InputObject $row -Name $noteHeader)
    }

    $key = $key.Trim()
    $email = $email.Trim()

    if ([string]::IsNullOrWhiteSpace($key)) {
      throw "Override row $lineNumber has an empty samAccountName: $resolvedPath"
    }

    if ($seen.ContainsKey($key)) {
      throw "Override row $lineNumber repeats samAccountName '$key': $resolvedPath"
    }
    $seen[$key] = $true

    if (-not [string]::IsNullOrWhiteSpace($email) -and -not (Test-ADMailAddress -Address $email)) {
      throw "Override row $lineNumber has an invalid email address '$email': $resolvedPath"
    }

    [PSCustomObject]@{
      PSTypeName     = 'Winkit.ADDS.AccountMapOverride'
      SamAccountName = $key
      Email          = $email
      Note           = $note
      RowNumber      = $lineNumber
      Path           = $resolvedPath
    }
  }
}

function Get-ADSortedKey {
  <#
  .SYNOPSIS
    Sorts keys with a deterministic ordinal comparison.

  .PARAMETER Key
    Keys to sort.

  .OUTPUTS
    System.String[]
  #>
  [CmdletBinding()]
  [OutputType([string[]])]
  param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [string[]]
    $Key
  )

  $list = New-Object 'System.Collections.Generic.List[string]'
  foreach ($item in $Key) {
    $list.Add([string]$item)
  }

  $comparison = [System.Comparison[string]] {
    param($left, $right)

    $result = [System.StringComparer]::OrdinalIgnoreCase.Compare($left, $right)
    if ($result -ne 0) {
      return $result
    }

    return [System.StringComparer]::Ordinal.Compare($left, $right)
  }

  $list.Sort($comparison)
  return $list.ToArray()
}

function New-ADAccountMap {
  <#
  .SYNOPSIS
    Resolves selected AD users into an account map result.

  .DESCRIPTION
    Selects users by filter, LDAP filter, search base, group, identity, or
    pipeline records. Resolution order is override CSV, then the AD mail
    attribute. Only exact allowlisted domains are accepted. No address is
    guessed. Missing required values fail the result unless -IncludeMissing is
    supplied, in which case the affected accounts are reported and omitted
    from the map.

  .PARAMETER InputObject
    Typed user records from Get-ADDirectoryRecord.

  .PARAMETER Identity
    Explicit user identities.

  .PARAMETER Group
    Group whose user members are selected.

  .PARAMETER Recursive
    Expand nested group membership.

  .PARAMETER Filter
    AD filter used by the default parameter set.

  .PARAMETER LDAPFilter
    LDAP filter. Mutually exclusive with Filter.

  .PARAMETER SearchBase
    OU or container to search.

  .PARAMETER AllowedDomain
    Exact email-domain allowlist. ASCII/Punycode only.

  .PARAMETER OverridePath
    Optional override CSV.

  .PARAMETER OnDuplicate
    Duplicate samAccountName policy: Fail, First, or Last.

  .PARAMETER AllowOverrideOutsideDomain
    Permit an explicit override outside AllowedDomain.

  .PARAMETER IncludeUnmatchedOverride
    Include override rows that did not match a selected account.

  .PARAMETER IncludeMissing
    Report and omit unresolved accounts instead of failing the result.

  .PARAMETER IncludeDisabled
    Include disabled users. By default only enabled users are selected.

  .PARAMETER Server
    Optional domain controller.

  .PARAMETER Credential
    Optional credential.

  .OUTPUTS
    Winkit.ADDS.AccountMapResult
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Builds an in-memory result only; no system state changes.')]
  [CmdletBinding(DefaultParameterSetName = 'Filter')]
  [OutputType('Winkit.ADDS.AccountMapResult')]
  param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Input', ValueFromPipeline = $true)]
    [psobject]
    $InputObject,

    [Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
    [string[]]
    $Identity,

    [Parameter(Mandatory = $true, ParameterSetName = 'Group')]
    [string]
    $Group,

    [Parameter(ParameterSetName = 'Group')]
    [switch]
    $Recursive,

    [Parameter(ParameterSetName = 'Filter')]
    [string]
    $Filter = '*',

    [Parameter(Mandatory = $true, ParameterSetName = 'LDAPFilter')]
    [string]
    $LDAPFilter,

    [Parameter(ParameterSetName = 'Filter')]
    [Parameter(ParameterSetName = 'LDAPFilter')]
    [string]
    $SearchBase,

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
    $Credential
  )

  begin {
    $domains = @()
    foreach ($domain in $AllowedDomain) {
      $trimmed = $domain.Trim()
      if ([string]::IsNullOrWhiteSpace($trimmed) -or
        $trimmed -notmatch '^[A-Za-z0-9]([A-Za-z0-9.-]*[A-Za-z0-9])?$' -or
        $trimmed -match '\.\.') {
        throw "AllowedDomain entry '$domain' is not an ASCII DNS domain."
      }
      $domains += $trimmed
    }

    $overrideEntries = @()
    if (-not [string]::IsNullOrWhiteSpace($OverridePath)) {
      $overrideEntries = @(Import-ADAccountMapOverride -Path $OverridePath)
    }

    $records = New-Object System.Collections.Generic.List[object]
    if ($PSCmdlet.ParameterSetName -ne 'Input') {
      $query = @{
        Resource        = 'Users'
        IncludeDisabled = $IncludeDisabled.IsPresent
      }

      if ($PSCmdlet.ParameterSetName -eq 'Identity') {
        $query.Identity = $Identity
      }
      elseif ($PSCmdlet.ParameterSetName -eq 'Group') {
        $query.Group = $Group
        if ($Recursive) {
          $query.Recursive = $true
        }
      }
      elseif ($PSCmdlet.ParameterSetName -eq 'LDAPFilter') {
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

      foreach ($record in @(Get-ADDirectoryRecord @query)) {
        $records.Add($record)
      }
    }
  }

  process {
    if ($PSCmdlet.ParameterSetName -eq 'Input' -and $null -ne $InputObject) {
      $records.Add($InputObject)
    }
  }

  end {
    $overrideLookup = @{}
    foreach ($override in $overrideEntries) {
      $overrideLookup[$override.SamAccountName] = $override
    }

    $entries = New-Object System.Collections.Generic.List[object]
    $matchedOverrides = @{}

    foreach ($record in $records) {
      $sam = [string](Get-ADRecordAttribute -Record $record -Name 'sAMAccountName')
      $sam = $sam.Trim()
      $enabledValue = Get-ADRecordAttribute -Record $record -Name 'Enabled'
      if (-not $IncludeDisabled -and $null -ne $enabledValue -and -not [bool]$enabledValue) {
        continue
      }

      $rawMail = [string](Get-ADRecordAttribute -Record $record -Name 'mail')
      $displayName = [string](Get-ADRecordAttribute -Record $record -Name 'DisplayName')
      $distinguishedName = [string](Get-ADRecordAttribute -Record $record -Name 'DistinguishedName')

      $entry = [ordered]@{
        PSTypeName        = 'Winkit.ADDS.AccountMapEntry'
        SamAccountName    = $sam
        Email             = $null
        Source            = 'None'
        Status            = 'Unresolved'
        Reason            = $null
        DisplayName       = $displayName
        DistinguishedName = $distinguishedName
        RawMail           = $rawMail
        Note              = $null
        OverridePath      = $null
      }

      if ([string]::IsNullOrWhiteSpace($sam)) {
        $entry.Reason = 'MissingSamAccountName'
        $entries.Add([PSCustomObject]$entry)
        continue
      }

      if ($overrideLookup.ContainsKey($sam)) {
        $override = $overrideLookup[$sam]
        $matchedOverrides[$sam] = $true
        $entry.OverridePath = $override.Path
        $entry.Note = $override.Note
        $entry.Source = 'Override'

        if ([string]::IsNullOrWhiteSpace($override.Email)) {
          $entry.Status = 'Excluded'
          $entry.Reason = 'OverrideExcluded'
        }
        elseif (-not (Test-ADMailAddress -Address $override.Email)) {
          $entry.Reason = 'InvalidOverrideMail'
        }
        else {
          $overrideDomain = ($override.Email -split '@')[-1]
          $overrideAllowed = $false
          foreach ($domain in $domains) {
            if ([string]::Equals($overrideDomain, $domain, [System.StringComparison]::OrdinalIgnoreCase)) {
              $overrideAllowed = $true
              break
            }
          }

          if ($overrideAllowed -or $AllowOverrideOutsideDomain) {
            $entry.Status = 'Resolved'
            $entry.Email = $override.Email
          }
          else {
            $entry.Reason = 'OverrideDomainNotAllowed'
          }
        }
      }
      elseif ([string]::IsNullOrWhiteSpace($rawMail)) {
        $entry.Reason = 'MissingMail'
      }
      elseif (-not (Test-ADMailAddress -Address $rawMail)) {
        $entry.Reason = 'InvalidMail'
      }
      elseif (-not (Test-ADMailAddress -Address $rawMail -AllowedDomain $domains)) {
        $entry.Reason = 'MailDomainNotAllowed'
      }
      else {
        $entry.Status = 'Resolved'
        $entry.Source = 'ADMail'
        $entry.Email = $rawMail
      }

      $entries.Add([PSCustomObject]$entry)
    }

    $unmatchedOverrides = New-Object System.Collections.Generic.List[object]
    foreach ($override in $overrideEntries) {
      if ($matchedOverrides.ContainsKey($override.SamAccountName)) {
        continue
      }

      $entry = [ordered]@{
        PSTypeName        = 'Winkit.ADDS.AccountMapEntry'
        SamAccountName    = $override.SamAccountName
        Email             = $null
        Source            = 'Override'
        Status            = 'Unresolved'
        Reason            = 'UnmatchedOverride'
        DisplayName       = $null
        DistinguishedName = $null
        RawMail           = $null
        Note              = $override.Note
        OverridePath      = $override.Path
      }

      if ([string]::IsNullOrWhiteSpace($override.Email)) {
        $entry.Status = 'Excluded'
      }
      elseif (-not (Test-ADMailAddress -Address $override.Email)) {
        $entry.Reason = 'InvalidOverrideMail'
      }
      else {
        $overrideDomain = ($override.Email -split '@')[-1]
        $overrideAllowed = $false
        foreach ($domain in $domains) {
          if ([string]::Equals($overrideDomain, $domain, [System.StringComparison]::OrdinalIgnoreCase)) {
            $overrideAllowed = $true
            break
          }
        }

        if ($overrideAllowed -or $AllowOverrideOutsideDomain) {
          $entry.Status = 'Resolved'
          $entry.Email = $override.Email
        }
        else {
          $entry.Reason = 'OverrideDomainNotAllowed'
        }
      }

      $unmatchedEntry = [PSCustomObject]$entry
      $unmatchedOverrides.Add($unmatchedEntry)
      if ($IncludeUnmatchedOverride) {
        $entries.Add($unmatchedEntry)
      }
    }

    $duplicates = New-Object System.Collections.Generic.List[object]
    $duplicateFailure = $false
    $grouped = @{}
    foreach ($entry in $entries) {
      if ([string]::IsNullOrWhiteSpace($entry.SamAccountName)) {
        continue
      }

      $key = $entry.SamAccountName.ToLowerInvariant()
      if (-not $grouped.ContainsKey($key)) {
        $grouped[$key] = New-Object System.Collections.Generic.List[object]
      }
      $grouped[$key].Add($entry)
    }

    foreach ($key in $grouped.Keys) {
      $duplicateGroup = $grouped[$key].ToArray()
      if ($duplicateGroup.Count -le 1) {
        continue
      }

      foreach ($entry in $duplicateGroup) {
        $duplicates.Add($entry)
      }

      if ($OnDuplicate -eq 'Fail') {
        $duplicateFailure = $true
        foreach ($entry in $duplicateGroup) {
          $entry.Status = 'Duplicate'
          $entry.Reason = 'DuplicateSamAccountName'
        }
        continue
      }

      $sorted = @($duplicateGroup | Sort-Object -Property DistinguishedName -CaseSensitive)
      if ($OnDuplicate -eq 'Last') {
        $keep = $sorted[$sorted.Count - 1]
      }
      else {
        $keep = $sorted[0]
      }

      foreach ($entry in $duplicateGroup) {
        if (-not [object]::ReferenceEquals($entry, $keep)) {
          $entry.Status = 'Duplicate'
          $entry.Reason = 'DuplicateSamAccountName'
        }
      }
    }

    $entryArray = $entries.ToArray()
    $unresolved = @($entryArray | Where-Object { $_.Status -eq 'Unresolved' })
    $excluded = @($entryArray | Where-Object { $_.Status -eq 'Excluded' })
    $resolved = @($entryArray | Where-Object { $_.Status -eq 'Resolved' })

    $status = 'Completed'
    if ($duplicateFailure -or ($unresolved.Count -gt 0 -and -not $IncludeMissing)) {
      $status = 'Failed'
    }

    $map = [ordered]@{}
    foreach ($entry in $resolved) {
      if (-not [string]::IsNullOrWhiteSpace($entry.Email) -and -not $map.Contains($entry.SamAccountName)) {
        $map[$entry.SamAccountName] = $entry.Email
      }
    }

    $sortedMap = [ordered]@{}
    foreach ($key in (Get-ADSortedKey -Key @($map.Keys))) {
      $sortedMap[$key] = $map[$key]
    }

    $sourceSummary = [ordered]@{
      ADMail   = @($entryArray | Where-Object { $_.Source -eq 'ADMail' }).Count
      Override = @($entryArray | Where-Object { $_.Source -eq 'Override' }).Count
      None     = @($entryArray | Where-Object { $_.Source -eq 'None' }).Count
    }

    $errors = @()
    if ($duplicateFailure) {
      $errors += 'Duplicate samAccountName values were found and OnDuplicate is Fail.'
    }

    $result = [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.AccountMapResult'
      Status            = $status
      Selection         = [ordered]@{
        Resource        = 'Users'
        Filter          = $Filter
        LDAPFilter      = $LDAPFilter
        SearchBase      = $SearchBase
        Group           = $Group
        Recursive       = [bool]$Recursive
        IdentityCount   = @($Identity).Count
        Server          = $Server
        IncludeDisabled = [bool]$IncludeDisabled
      }
      Entries           = $entryArray
      Unresolved        = $unresolved
      Excluded          = $excluded
      Duplicates        = $duplicates.ToArray()
      UnmatchedOverride = $unmatchedOverrides.ToArray()
      Map               = $sortedMap
      SourceSummary     = $sourceSummary
      EntryCount        = $entryArray.Count
      ResolvedCount     = $resolved.Count
      UnresolvedCount   = $unresolved.Count
      Path              = $null
      ReportPath        = $null
      Sha256            = $null
      Errors            = $errors
    }

    return $result
  }
}

function ConvertTo-ADJsonString {
  <#
  .SYNOPSIS
    Escapes one string for the deterministic account-map serializer.

  .PARAMETER Value
    String to escape.

  .OUTPUTS
    System.String
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]
    $Value
  )

  $builder = New-Object System.Text.StringBuilder
  [void]$builder.Append('"')
  foreach ($character in $Value.ToCharArray()) {
    $code = [int]$character
    if ($code -eq 0x22) {
      [void]$builder.Append('\"')
    }
    elseif ($code -eq 0x5C) {
      [void]$builder.Append('\\')
    }
    elseif ($code -lt 0x20) {
      [void]$builder.Append(('\u{0:x4}' -f $code))
    }
    else {
      [void]$builder.Append($character)
    }
  }
  [void]$builder.Append('"')
  return $builder.ToString()
}

function ConvertTo-ADAccountMapJson {
  <#
  .SYNOPSIS
    Serializes an account map deterministically.

  .DESCRIPTION
    Emits one JSON object with ordinal-sorted keys, two-space indentation,
    CRLF line endings, a final CRLF, and RFC 8259 escaping. Duplicate keys
    after case folding terminate. The caller writes the returned string as
    UTF-8 without BOM.

  .PARAMETER Map
    Dictionary of samAccountName to email.

  .OUTPUTS
    System.String
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [System.Collections.IDictionary]
    $Map
  )

  process {
    $seen = @{}
    foreach ($key in $Map.Keys) {
      $textKey = [string]$key
      if ($seen.ContainsKey($textKey)) {
        throw "The account map contains duplicate keys after case folding: '$textKey'."
      }
      $seen[$textKey] = $true
    }

    $keys = @(Get-ADSortedKey -Key @($Map.Keys))
    if ($keys.Count -eq 0) {
      return ('{}' + "`r`n")
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('{')
    for ($index = 0; $index -lt $keys.Count; $index++) {
      $key = $keys[$index]
      $suffix = ''
      if ($index -lt ($keys.Count - 1)) {
        $suffix = ','
      }
      $line = '  ' + (ConvertTo-ADJsonString -Value $key) + ': ' + (ConvertTo-ADJsonString -Value ([string]$Map[$key])) + $suffix
      $lines.Add($line)
    }
    $lines.Add('}')

    return (($lines -join "`r`n") + "`r`n")
  }
}

function Write-ADTextFileAtomic {
  <#
  .SYNOPSIS
    Writes UTF-8 text without BOM through a temporary file.

  .PARAMETER Path
    Destination file.

  .PARAMETER Text
    Text to write.

  .OUTPUTS
    None
  #>
  [CmdletBinding()]
  [OutputType([void])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]
    $Text
  )

  $directory = Split-Path -Path $Path -Parent
  $tempPath = Join-Path -Path $directory -ChildPath ('.' + [System.IO.Path]::GetFileName($Path) + '.' + [guid]::NewGuid().ToString('N') + '.tmp')
  $utf8 = New-Object System.Text.UTF8Encoding($false)

  try {
    [System.IO.File]::WriteAllText($tempPath, $Text, $utf8)
    Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
  }
  finally {
    if (Test-Path -LiteralPath $tempPath) {
      Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
  }
}

function Export-ADAccountMap {
  <#
  .SYNOPSIS
    Writes the account map JSON and optional audit report.

  .DESCRIPTION
    Refuses to overwrite an existing file unless -Force is supplied. A Failed
    result writes only the report. The map is written atomically as UTF-8
    without BOM with deterministic key order. ShouldProcess covers both files.

  .PARAMETER InputObject
    Account map result from New-ADAccountMap.

  .PARAMETER OutputPath
    Map file. Defaults to ad-account-map.json in the current directory.

  .PARAMETER ReportPath
    Optional CSV audit report.

  .PARAMETER Force
    Replace existing output or report files.

  .PARAMETER DryRun
    Preview without writing.

  .PARAMETER PassThru
    Return the updated result.

  .OUTPUTS
    Winkit.ADDS.AccountMapResult
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  [OutputType('Winkit.ADDS.AccountMapResult')]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [psobject]
    $InputObject,

    [string]
    $OutputPath = 'ad-account-map.json',

    [string]
    $ReportPath,

    [switch]
    $Force,

    [switch]
    $DryRun,

    [switch]
    $PassThru
  )

  begin {
    $results = New-Object System.Collections.Generic.List[object]
  }

  process {
    if ($null -ne $InputObject) {
      $results.Add($InputObject)
    }
  }

  end {
    if ($results.Count -ne 1) {
      throw 'Export-ADAccountMap accepts exactly one account-map result.'
    }

    if ($DryRun) {
      $WhatIfPreference = $true
    }

    $result = $results[0]
    if ($null -eq $result.PSObject.Properties['Map'] -or $null -eq $result.PSObject.Properties['Status']) {
      throw 'InputObject is not a Winkit.ADDS.AccountMapResult.'
    }

    $outputFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $outputDirectory = Split-Path -Path $outputFullPath -Parent
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
      throw "Output directory not found: $outputDirectory"
    }

    if (Test-Path -LiteralPath $outputFullPath -PathType Container) {
      throw "OutputPath is a directory: $outputFullPath"
    }

    if ((Test-Path -LiteralPath $outputFullPath -PathType Leaf) -and -not $Force) {
      throw "Output file already exists. Use -Force to replace it: $outputFullPath"
    }

    $reportFullPath = $null
    if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
      $reportFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
      $reportDirectory = Split-Path -Path $reportFullPath -Parent
      if (-not (Test-Path -LiteralPath $reportDirectory -PathType Container)) {
        throw "Report directory not found: $reportDirectory"
      }
      if (Test-Path -LiteralPath $reportFullPath -PathType Container) {
        throw "ReportPath is a directory: $reportFullPath"
      }
      if ((Test-Path -LiteralPath $reportFullPath -PathType Leaf) -and -not $Force) {
        throw "Report file already exists. Use -Force to replace it: $reportFullPath"
      }
    }

    $writeMap = ($result.Status -ne 'Failed')
    if ($writeMap) {
      $json = ConvertTo-ADAccountMapJson -Map $result.Map
      if ($PSCmdlet.ShouldProcess($outputFullPath, 'Write AD account map')) {
        Write-ADTextFileAtomic -Path $outputFullPath -Text $json
        $result.Path = $outputFullPath
        $result.Sha256 = (Get-FileHash -LiteralPath $outputFullPath -Algorithm SHA256).Hash
      }
      else {
        $result.Status = 'DryRun'
      }
    }

    if ($null -ne $reportFullPath) {
      $reportRows = New-Object System.Collections.Generic.List[object]
      foreach ($entry in @($result.Entries) + @($result.UnmatchedOverride)) {
        $reportRows.Add([PSCustomObject][ordered]@{
            samAccountName    = $entry.SamAccountName
            email             = $entry.Email
            source            = $entry.Source
            status            = $entry.Status
            reason            = $entry.Reason
            rawMail           = $entry.RawMail
            distinguishedName = $entry.DistinguishedName
            note              = $entry.Note
          })
      }

      if ($PSCmdlet.ShouldProcess($reportFullPath, 'Write AD account-map audit report')) {
        if ($reportRows.Count -eq 0) {
          $reportText = 'samAccountName;email;source;status;reason;rawMail;distinguishedName;note' + "`r`n"
        }
        else {
          $reportLines = @($reportRows | ConvertTo-Csv -NoTypeInformation -Delimiter ';')
          $reportText = ($reportLines -join "`r`n") + "`r`n"
        }
        Write-ADTextFileAtomic -Path $reportFullPath -Text $reportText
        $result.ReportPath = $reportFullPath
      }
      elseif ($result.Status -ne 'Failed') {
        $result.Status = 'DryRun'
      }
    }

    if ($PassThru -or $WhatIfPreference) {
      return $result
    }
  }
}
