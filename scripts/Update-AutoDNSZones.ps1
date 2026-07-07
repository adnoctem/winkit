#Requires -Version 5.1

<#
.SYNOPSIS
  Manages DNS zone and domain configurations on AutoDNS (InternetX Domainrobot) via the JSON API.

.DESCRIPTION
  Reads a JSON configuration file describing desired states for domains and DNS zones,
  then applies changes to AutoDNS.  Supports two levels of configuration:

  - Domain level: nameserver delegation, DNSSEC key material, and other registry-level settings.
  - Zone level: resource records, main IP address, zone DNSSEC toggle, and wwwInclude.

  Discovers the virtual name server automatically for each zone.  Provides safety checks
  including TTL warnings and explicit confirmation prompts before making changes.

.PARAMETER Credential
  A PSCredential object for HTTP Basic authentication against the AutoDNS JSON API.
  Obtain via Get-Credential or pass a credential object.

.PARAMETER Config
  Path to a JSON configuration file describing the desired states.  See -ExportConfig
  for the template format, or use -ExportCurrentState to generate one from live data.

.PARAMETER Context
  The AutoDNS context number (default: 4 for live system; 1 for demo).

.PARAMETER DryRun
  Preview the changes that would be made without applying them.

.PARAMETER Force
  Skip confirmation prompts for destructive operations.

.PARAMETER ExportConfig
  Prints an example JSON configuration template to the console (or to a file with -ExportPath)
  and exits.  Use -ExportCurrentState to generate a config from your live AutoDNS data.

.PARAMETER ExportCurrentState
  Exports the current state of one or more domains and zones from AutoDNS as a JSON
  configuration file.  When combined with -Config, exports only the entries listed in that
  file.  Without -Config, exports all domains (and their zones where applicable) on the
  account.  Requires -Credential.

.PARAMETER ExportPath
  When used together with -ExportConfig or -ExportCurrentState, writes the JSON to this
  file path instead of printing to the console.

.PARAMETER PassThru
  Output result objects for each processed entry.

.EXAMPLE
  PS> $cred = Get-Credential
  PS> ./Update-AutoDNSZones.ps1 -Config .\config.json -Credential $cred
  Reads configuration from config.json and applies domain and zone changes to AutoDNS.

.EXAMPLE
  PS> ./Update-AutoDNSZones.ps1 -Config .\config.json -Credential $cred -DryRun
  Previews the changes that would be made without applying them.

.EXAMPLE
  PS> ./Update-AutoDNSZones.ps1 -ExportConfig
  Prints an example configuration template to the console.

.EXAMPLE
  PS> ./Update-AutoDNSZones.ps1 -ExportCurrentState -Credential $cred
  Fetches all domains from AutoDNS and prints their current state as JSON.

.EXAMPLE
  PS> ./Update-AutoDNSZones.ps1 -ExportCurrentState -Config .\config.json -Credential $cred -ExportPath .\current.json
  Fetches only the entries listed in config.json and saves their current state.

.LINK
  https://github.com/adnoctem/winkit
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false, HelpMessage = 'PSCredential for HTTP Basic authentication against the AutoDNS API.')]
  [pscredential]$Credential,

  [Parameter(Mandatory = $false, HelpMessage = 'Path to a JSON configuration file describing desired states.')]
  [string]$Config,

  [Parameter(Mandatory = $false, HelpMessage = 'AutoDNS context number (4 = live, 1 = demo).')]
  [int]$Context = 4,

  [Parameter(Mandatory = $false, HelpMessage = 'Preview changes without applying them.')]
  [switch]$DryRun,

  [Parameter(Mandatory = $false, HelpMessage = 'Skip confirmation prompts.')]
  [switch]$Force,

  [Parameter(Mandatory = $false, HelpMessage = 'Print an example configuration template.')]
  [switch]$ExportConfig,

  [Parameter(Mandatory = $false, HelpMessage = 'Export current state from AutoDNS as JSON.')]
  [switch]$ExportCurrentState,

  [Parameter(Mandatory = $false, HelpMessage = 'File path for -ExportConfig or -ExportCurrentState.')]
  [string]$ExportPath,

  [Parameter(Mandatory = $false)]
  [switch]$PassThru
)

# ---- Module import -----------------------------------------------------------
$scriptRoot = Split-Path $PSScriptRoot -Parent
$module = Join-Path $scriptRoot 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

$baseUrl = 'https://api.autodns.com/v1'
$highTtlThreshold = 1800
$apiDelayMs = 350

# ---- Helper functions -------------------------------------------------------

function Invoke-AutoDNSRequest {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $false)]
    [string]$Method = 'GET',
    [Parameter(Mandatory = $false)]
    [object]$Body
  )

  $pair = "$($Credential.UserName):$($Credential.GetNetworkCredential().Password)"
  $encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))

  $params = @{
    Uri = "$baseUrl$Path"
    Method = $Method
    Headers = @{
      Authorization = "Basic $encodedCreds"
      'X-Domainrobot-Context' = $Context
      'User-Agent' = 'winkit/1.0'
    }
    ContentType = 'application/json'
  }

  if ($Body) {
    $params['Body'] = ($Body | ConvertTo-Json -Depth 10)
  }

  try {
    $response = Invoke-RestMethod @params
  }
  catch {
    $statusCode = $_.Exception.Response.StatusCode.value__
    $statusText = $_.Exception.Response.StatusDescription
    throw "AutoDNS API error ($statusCode $statusText) for $Method $Path`: $_"
  }

  if ($response.status.type -ne 'SUCCESS' -and $response.status.type -ne 'N') {
    throw "AutoDNS API returned status '$($response.status.type)': $($response.status.text)"
  }

  Start-Sleep -Milliseconds $apiDelayMs

  return $response
}

# -- Zone helpers -------------------------------------------------------------

function Get-AutoDNSZone {
  [CmdletBinding()]
  param([string]$Origin)

  $response = Invoke-AutoDNSRequest -Path "/zone/$Origin"
  if (-not $response.data -or $response.data.Count -eq 0) {
    throw "Zone '$Origin' not found on AutoDNS."
  }
  return $response.data[0]
}

function Get-AutoDNSZoneList {
  [CmdletBinding()]
  param()

  $response = Invoke-AutoDNSRequest -Path '/zone/_search' -Method POST -Body @{}
  if (-not $response.data) { return @() }
  return $response.data
}

# -- Domain helpers -----------------------------------------------------------

function Get-AutoDNSDomain {
  [CmdletBinding()]
  param([string]$Name)

  $response = Invoke-AutoDNSRequest -Path "/domain/$Name"
  if (-not $response.data -or $response.data.Count -eq 0) {
    throw "Domain '$Name' not found on AutoDNS."
  }
  return $response.data[0]
}

function Get-AutoDNSDomainList {
  [CmdletBinding()]
  param()

  $response = Invoke-AutoDNSRequest -Path '/domain/_search' -Method POST -Body @{}
  if (-not $response.data) { return @() }
  return $response.data
}

function Update-AutoDNSDomain {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [string]$Name,
    [object]$Body
  )

  if (-not $PSCmdlet.ShouldProcess($Name, 'Update domain via AutoDNS API')) {
    return
  }

  $response = Invoke-AutoDNSRequest -Path "/domain/$Name" -Method PUT -Body $Body

  $jobInfo = $null
  if ($response.data -and $response.data.Count -gt 0 -and $response.data[0].id) {
    $jobInfo = @{ JobId = $response.data[0].id }
  }

  return $jobInfo
}

# -- Comparison helpers -------------------------------------------------------

function Compare-ZoneState {
  [CmdletBinding()]
  param(
    [object]$Current,
    [object]$Desired
  )

  $changes = @()

  $currentMain = if ($Current.main) { $Current.main.address } else { $null }
  $desiredMain = if ($Desired.main) { $Desired.main.address } else { $null }
  if ($desiredMain -and $currentMain -ne $desiredMain) {
    $changes += @{
      Field = 'main'
      Current = $currentMain
      Desired = $desiredMain
    }
  }

  $currentWww = if ($Current.PSObject.Properties.Name -contains 'wwwInclude') { [bool]$Current.wwwInclude } else { $null }
  $desiredWww = if ($Desired.PSObject.Properties.Name -contains 'wwwInclude') { $Desired.wwwInclude } else { $null }
  if ($null -ne $desiredWww -and $currentWww -ne $desiredWww) {
    $changes += @{
      Field = 'wwwInclude'
      Current = $currentWww
      Desired = $desiredWww
    }
  }

  $currentDnssec = if ($Current.PSObject.Properties.Name -contains 'dnssec') { [bool]$Current.dnssec } else { $null }
  $desiredDnssec = if ($Desired.PSObject.Properties.Name -contains 'dnssec') { $Desired.dnssec } else { $null }
  if ($null -ne $desiredDnssec -and $currentDnssec -ne $desiredDnssec) {
    $changes += @{
      Field = 'zone dnssec'
      Current = $currentDnssec
      Desired = $desiredDnssec
    }
  }

  $currentRecords = if ($Current.resourceRecords) { $Current.resourceRecords } else { @() }
  $desiredRecords = if ($Desired.resourceRecords) { $Desired.resourceRecords } else { @() }
  if ($desiredRecords.Count -gt 0) {
    $recordsMatch = Compare-RecordArrays -Current $currentRecords -Desired $desiredRecords
    if (-not $recordsMatch) {
      $changes += @{
        Field = 'records'
        Current = "($($currentRecords.Count) existing)"
        Desired = "($($desiredRecords.Count) desired)"
        ReplaceAll = $true
      }
    }
  }

  return $changes
}

function Compare-DomainState {
  [CmdletBinding()]
  param(
    [object]$Current,
    [object]$Desired
  )

  $changes = @()

  if ($Desired.PSObject.Properties.Name -contains 'nameServers' -and $Desired.nameServers.Count -gt 0) {
    $currentNs = if ($Current.nameServers) { $Current.nameServers } else { @() }
    $nsMatch = Compare-NameServerArrays -Current $currentNs -Desired $Desired.nameServers
    if (-not $nsMatch) {
      $currentNsNames = ($currentNs | ForEach-Object { $_.name }) -join ', '
      $desiredNsNames = ($Desired.nameServers | ForEach-Object { $_.name }) -join ', '
      $changes += @{
        Field = 'nameServers'
        Current = $currentNsNames
        Desired = $desiredNsNames
      }
    }
  }

  # Compare DNSSEC settings
  if ($Desired.PSObject.Properties.Name -contains 'dnssec') {
    $dnssecCfg = $Desired.dnssec
    if ($dnssecCfg.PSObject.Properties.Name -contains 'enabled') {
      $currentVal = if ($Current.PSObject.Properties.Name -contains 'dnssec') { [bool]$Current.dnssec } else { $false }
      $desiredVal = [bool]$dnssecCfg.enabled
      if ($currentVal -ne $desiredVal) {
        $changes += @{
          Field = 'dnssec enabled'
          Current = $currentVal
          Desired = $desiredVal
        }
      }
    }
    if ($dnssecCfg.PSObject.Properties.Name -contains 'auto') {
      $currentVal = if ($Current.PSObject.Properties.Name -contains 'autoDnssec') { [bool]$Current.autoDnssec } else { $false }
      $desiredVal = [bool]$dnssecCfg.auto
      if ($currentVal -ne $desiredVal) {
        $changes += @{
          Field = 'dnssec auto'
          Current = $currentVal
          Desired = $desiredVal
        }
      }
    }
    if ($dnssecCfg.PSObject.Properties.Name -contains 'keys') {
      $currentKeys = if ($Current.dnssecData) { $Current.dnssecData } else { @() }
      $desiredKeys = $dnssecCfg.keys
      $keysMatch = $desiredKeys.Count -eq 0
      if ($desiredKeys.Count -gt 0 -and $currentKeys.Count -eq $desiredKeys.Count) {
        $keysMatch = $true
        for ($i = 0; $i -lt $desiredKeys.Count; $i++) {
          $dk = $desiredKeys[$i]
          $ck = $currentKeys[$i]
          $keyMatch = ($ck.algorithm -eq $dk.algorithm) -and ($ck.flags -eq $dk.flags) -and ($ck.protocol -eq $dk.protocol) -and ($ck.publicKey -eq $dk.publicKey)
          if (-not $keyMatch) { $keysMatch = $false; break }
        }
      }
      if (-not $keysMatch) {
        $changes += @{
          Field = 'dnssec keys'
          Current = "($($currentKeys.Count) key(s))"
          Desired = "($($desiredKeys.Count) key(s))"
        }
      }
    }
    if ($dnssecCfg.PSObject.Properties.Name -contains 'keyRollover' -and [bool]$dnssecCfg.keyRollover) {
      $changes += @{
        Field = 'dnssec key rollover'
        Current = 'inactive'
        Desired = 'triggered'
      }
    }
  }

  return $changes
}

function Compare-RecordArrays {
  [CmdletBinding()]
  param(
    [object[]]$Current,
    [object[]]$Desired
  )

  if ($Current.Count -ne $Desired.Count) { return $false }

  foreach ($desired in $Desired) {
    $match = $false
    foreach ($current in $Current) {
      $nameMatch = $current.name -eq $desired.name
      $typeMatch = $current.type -eq $desired.type
      $valueMatch = $current.value -eq $desired.value
      $ttlMatch = (-not $desired.PSObject.Properties.Name.Contains('ttl')) -or ($null -eq $desired.ttl) -or ($current.ttl -eq $desired.ttl)
      $prefMatch = (-not $desired.PSObject.Properties.Name.Contains('pref')) -or ($null -eq $desired.pref) -or ($current.pref -eq $desired.pref)
      if ($nameMatch -and $typeMatch -and $valueMatch -and $ttlMatch -and $prefMatch) {
        $match = $true
        break
      }
    }
    if (-not $match) { return $false }
  }
  return $true
}

function Compare-NameServerArrays {
  [CmdletBinding()]
  param(
    [object[]]$Current,
    [object[]]$Desired
  )

  if ($Current.Count -ne $Desired.Count) { return $false }

  foreach ($desired in $Desired) {
    $match = $false
    foreach ($current in $Current) {
      if ($current.name -eq $desired.name) {
        $match = $true
        break
      }
    }
    if (-not $match) { return $false }
  }
  return $true
}

function Get-HighTtlRecords {
  [CmdletBinding()]
  param([object[]]$Records)

  $high = @()
  foreach ($r in $Records) {
    if ($r.ttl -and $r.ttl -gt $highTtlThreshold) {
      $high += $r
    }
  }
  return $high
}

function Confirm-DestructiveOperation {
  [CmdletBinding()]
  param([string]$Message)

  if ($Force) { return $true }

  $response = Read-Host "$Message`nType 'yes' to confirm"
  return $response -eq 'yes'
}

# ---- Example config for -ExportConfig ---------------------------------------
$exampleConfig = @(
  [PSCustomObject]@{
    origin = 'example.com'
    domain = [PSCustomObject]@{
      nameServers = @(
        [PSCustomObject]@{ name = 'ns1.example.com' }
        [PSCustomObject]@{ name = 'ns2.example.com' }
      )
      dnssec = [PSCustomObject]@{
        enabled = $true
        auto = $false
        keys = @(
          [PSCustomObject]@{
            algorithm = 13
            flags = 257
            protocol = 3
            publicKey = 'base64-ksk-key'
          }
          [PSCustomObject]@{
            algorithm = 13
            flags = 256
            protocol = 3
            publicKey = 'base64-zsk-key'
          }
        )
        keyRollover = $false
      }
    }
    wwwInclude = $true
    records = @(
      [PSCustomObject]@{
        name = '@'
        type = 'A'
        value = '203.0.113.10'
        ttl = 3600
      }
      [PSCustomObject]@{
        name = 'www'
        type = 'A'
        value = '203.0.113.10'
        ttl = 3600
      }
      [PSCustomObject]@{
        name = 'mail'
        type = 'MX'
        value = 'mail.example.com'
        pref = 10
        ttl = 3600
      }
      [PSCustomObject]@{
        name = '@'
        type = 'TXT'
        value = 'v=spf1 mx ~all'
        ttl = 3600
      }
    )
  }
)

# ---- -ExportConfig: print example template ----------------------------------
if ($ExportConfig) {
  $tipMessage = @'

TIP: This is an example template. To generate a real config from your live
     data on AutoDNS, use -ExportCurrentState with -Credential.

     Example: .\Update-AutoDNSZones.ps1 -ExportCurrentState -Credential (Get-Credential)

'@
  Write-Log -Message $tipMessage -Color Cyan

  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $exampleConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Example configuration exported to: $_exportPath" -Color Green
  }
  else {
    $exampleConfig | ConvertTo-Json -Depth 10
  }
  exit 0
}

# ---- Credential validation for API-dependent modes --------------------------
if ($ExportCurrentState -or (-not $ExportConfig -and -not $ExportCurrentState -and $PSBoundParameters.ContainsKey('Config'))) {
  if (-not $Credential) {
    Write-Log -Message '-Credential is required for AutoDNS API access.' -Color Red
    exit 1
  }
}

# ---- -ExportCurrentState: fetch live data from API --------------------------
if ($ExportCurrentState) {
  $originList = @()

  if ($PSBoundParameters.ContainsKey('Config') -and -not [string]::IsNullOrWhiteSpace($Config)) {
    $_cfgPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
    if (-not (Test-Path -LiteralPath $_cfgPath)) {
      Write-Log -Message "Config file not found: '$_cfgPath'" -Color Red
      exit 1
    }
    try {
      $_cfgJson = Get-Content -LiteralPath $_cfgPath -Raw -ErrorAction Stop
      $_cfgEntries = ConvertFrom-Json -InputObject $_cfgJson -ErrorAction Stop
      if ($_cfgEntries -isnot [array]) { $_cfgEntries = @($_cfgEntries) }
      $originList = $_cfgEntries | ForEach-Object { $_.origin } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
    catch {
      Write-Log -Message "Failed to parse config file '$_cfgPath': $_" -Color Red
      exit 1
    }
    Write-Log -Message "Exporting current state for $($originList.Count) entry(ies) from config ..." -Color Cyan
  }
  else {
    Write-Log -Message 'Fetching domain list from AutoDNS ...' -Color Cyan
    try {
      $allDomains = Get-AutoDNSDomainList
      $originList = $allDomains | ForEach-Object { $_.name } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
    catch {
      Write-Log -Message "Failed to fetch domain list: $_" -Color Red
      exit 1
    }
    Write-Log -Message "Found $($originList.Count) domain(s) on AutoDNS." -Color Cyan
  }

  # Also fetch zone list to know which domains have zones
  $zoneOrigins = @()
  try {
    $allZones = Get-AutoDNSZoneList
    $zoneOrigins = $allZones | ForEach-Object { $_.origin } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  }
  catch {
    Write-Log -Message "Warning: Could not fetch zone list: $_" -Color Yellow
  }

  $exported = @()

  foreach ($origin in $originList) {
    Write-Log -Message "  Fetching domain: $origin ..." -Color Gray

    # Fetch domain info
    try {
      $domain = Get-AutoDNSDomain -Name $origin
    }
    catch {
      Write-Log -Message "  -> Failed to fetch domain '$origin': $_" -Color Yellow
      continue
    }

    $entry = [PSCustomObject]@{
      origin = $domain.name
    }

    # Add domain block
    $domainBlock = [PSCustomObject]@{}
    $domainBlock | Add-Member -NotePropertyName nameServers -NotePropertyValue $domain.nameServers

    # Build DNSSEC object if domain has DNSSEC data
    $hasDnssec = $domain.PSObject.Properties.Name -contains 'dnssec'
    $hasAutoDnssec = $domain.PSObject.Properties.Name -contains 'autoDnssec'
    $hasDnssecData = $domain.PSObject.Properties.Name -contains 'dnssecData'
    if ($hasDnssec -or $hasAutoDnssec -or $hasDnssecData) {
      $dnssecObj = [PSCustomObject]@{}
      if ($hasDnssec) { $dnssecObj | Add-Member -NotePropertyName enabled -NotePropertyValue ([bool]$domain.dnssec) }
      if ($hasAutoDnssec) { $dnssecObj | Add-Member -NotePropertyName auto -NotePropertyValue ([bool]$domain.autoDnssec) }
      if ($hasDnssecData) { $dnssecObj | Add-Member -NotePropertyName keys -NotePropertyValue $domain.dnssecData }
      $domainBlock | Add-Member -NotePropertyName dnssec -NotePropertyValue $dnssecObj
    }

    $domainBlockProps = $domainBlock.PSObject.Properties.Name
    if ($domainBlockProps.Count -gt 0) {
      $entry | Add-Member -NotePropertyName domain -NotePropertyValue $domainBlock
    }

    # Fetch zone info if available
    if ($zoneOrigins -contains $origin) {
      try {
        $zone = Get-AutoDNSZone -Origin $origin
        $entry | Add-Member -NotePropertyName main -NotePropertyValue $zone.main
        $entry | Add-Member -NotePropertyName wwwInclude -NotePropertyValue ([bool]$zone.wwwInclude)
        $entry | Add-Member -NotePropertyName dnssec -NotePropertyValue ([bool]$zone.dnssec)
        $entry | Add-Member -NotePropertyName records -NotePropertyValue $zone.resourceRecords
      }
      catch {
        Write-Log -Message "  -> (zone not available: $_)" -Color Gray
      }
    }

    $exported += $entry
  }

  if ($exported.Count -eq 0) {
    Write-Log -Message 'No entries were exported.' -Color Yellow
    exit 1
  }

  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $exported | ConvertTo-Json -Depth 10 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Current state exported to: $_exportPath" -Color Green
  }
  else {
    $exported | ConvertTo-Json -Depth 10
  }

  Write-Log -Message "Exported $($exported.Count) entry(ies)." -Color Green
  exit 0
}

# ---- Validate config --------------------------------------------------------
if (-not $PSBoundParameters.ContainsKey('Config') -or [string]::IsNullOrWhiteSpace($Config)) {
  Write-Log -Message '-Config is required (or use -ExportConfig or -ExportCurrentState).' -Color Red
  exit 1
}

$_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
if (-not (Test-Path -LiteralPath $_configPath)) {
  Write-Log -Message "Config file not found: '$_configPath'" -Color Red
  exit 1
}

# ---- Parse config -----------------------------------------------------------
try {
  $_jsonContent = Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop
  $configEntries = ConvertFrom-Json -InputObject $_jsonContent -ErrorAction Stop
  if ($configEntries -isnot [array]) { $configEntries = @($configEntries) }
}
catch {
  Write-Log -Message "Failed to parse config file '$_configPath': $_" -Color Red
  exit 1
}

Write-Log -Message "Loaded $($configEntries.Count) configuration entry(ies) from '$_configPath'" -Color Cyan
Write-Log -Message "Connecting to AutoDNS at $baseUrl (context $Context) ..." -Color Cyan

# ---- DryRun preamble --------------------------------------------------------
if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

# ---- Phase 1: Collect all changes across entries ----------------------------
$changePlan = @()

foreach ($entry in $configEntries) {
  $origin = $entry.origin
  if ([string]::IsNullOrWhiteSpace($origin)) {
    Write-Log -Message 'Skipping config entry with empty or missing origin.' -Color Yellow
    continue
  }

  Write-Log -Message "`nInspecting: $origin" -Color Cyan

  $hasDomainBlock = $entry.PSObject.Properties.Name -contains 'domain'
  $hasZoneFields = ($entry.PSObject.Properties.Name -contains 'main') -or ($entry.PSObject.Properties.Name -contains 'records') -or ($entry.PSObject.Properties.Name -contains 'wwwInclude') -or ($entry.PSObject.Properties.Name -contains 'dnssec') -or ($entry.PSObject.Properties.Name -contains 'resourceRecords')

  if (-not $hasDomainBlock -and -not $hasZoneFields) {
    Write-Log -Message "  -> No domain or zone fields specified. Skipping." -Color Yellow
    continue
  }

  # Fetch domain info
  try {
    $currentDomain = Get-AutoDNSDomain -Name $origin
  }
  catch {
    Write-Log -Message "  -> Failed to fetch domain '$origin': $_" -Color Red
    continue
  }

  # Fetch zone info if needed
  $currentZone = $null
  if ($hasZoneFields) {
    try {
      $currentZone = Get-AutoDNSZone -Origin $origin
    }
    catch {
      Write-Log -Message "  -> Zone not found for '$origin'. Zone fields will be skipped." -Color Yellow
      $hasZoneFields = $false
    }
  }

  $vns = if ($currentZone) { $currentZone.virtualNameServer } else { $null }

  # -- Compute domain-level changes --
  $domainChanges = @()
  if ($hasDomainBlock) {
    $domainDesired = $entry.domain
    $domainChanges = Compare-DomainState -Current $currentDomain -Desired $domainDesired
  }

  # -- Compute zone-level changes --
  $zoneChanges = @()
  $desiredPayload = @{}
  if ($hasZoneFields -and $currentZone) {
    if ($entry.PSObject.Properties.Name -contains 'main') {
      $desiredPayload.main = $entry.main
    }
    if ($entry.PSObject.Properties.Name -contains 'wwwInclude') {
      $desiredPayload.wwwInclude = [bool]$entry.wwwInclude
    }
    $zoneDnssecProp = if ($entry.PSObject.Properties.Name -contains 'dnssec') { 'dnssec' } else { $null }
    if ($zoneDnssecProp -and $entry.$zoneDnssecProp -is [bool]) {
      $desiredPayload.dnssec = [bool]$entry.$zoneDnssecProp
    }
    if ($entry.PSObject.Properties.Name -contains 'records' -and $entry.records.Count -gt 0) {
      $desiredPayload.resourceRecords = $entry.records
    }

    $zoneChanges = Compare-ZoneState -Current $currentZone -Desired $desiredPayload

    # Check TTL warnings on existing zone records
    $currentRecords = if ($currentZone.resourceRecords) { $currentZone.resourceRecords } else { @() }
    $highTtlRecords = Get-HighTtlRecords -Records $currentRecords
  }
  else {
    $highTtlRecords = @()
  }

  $totalChanges = $domainChanges + $zoneChanges
  if ($totalChanges.Count -eq 0) {
    Write-Log -Message "  -> No changes detected." -Color Green
    continue
  }

  $hasRecordReplace = ($zoneChanges | Where-Object { $_.Field -eq 'records' -and $_.ReplaceAll }).Count -gt 0

  $changePlan += @{
    Origin = $origin
    Vns = $vns
    CurrentDomain = $currentDomain
    CurrentZone = $currentZone
    DesiredPayload = $desiredPayload
    DomainChanges = $domainChanges
    ZoneChanges = $zoneChanges
    HasRecordReplace = $hasRecordReplace
    HighTtlRecords = $highTtlRecords
    HasDomainBlock = $hasDomainBlock
    HasZoneFields = $hasZoneFields
    DomainConfig = if ($hasDomainBlock) { $entry.domain } else { $null }
  }
}

# ---- Phase 2: Display summary table -----------------------------------------
if ($changePlan.Count -eq 0) {
  Write-Log -Message "`nAll entries are already at the desired state. Nothing to do." -Color Green
  if ($PassThru) { @() }
  exit 0
}

Write-Log -Message "`n========== CHANGE SUMMARY ==========" -Color Cyan
foreach ($plan in $changePlan) {
  Write-Log -Message "Entry: $($plan.Origin)" -Color Yellow

  foreach ($c in $plan.DomainChanges) {
    Write-Log -Message "  [domain] $($c.Field): $($c.Current) -> $($c.Desired)" -Color Yellow
  }
  foreach ($c in $plan.ZoneChanges) {
    Write-Log -Message "  [zone] $($c.Field): $($c.Current) -> $($c.Desired)" -Color Yellow
  }

  if ($plan.HasRecordReplace) {
    Write-Log -Message '  *** RECORDS WILL BE FULLY REPLACED ***' -Color Red
  }
  if ($plan.HighTtlRecords.Count -gt 0) {
    Write-Log -Message "  *** $($plan.HighTtlRecords.Count) record(s) with TTL > $highTtlThreshold (DNS propagation risk) ***" -Color Red
  }
}
Write-Log -Message "====================================`n" -Color Cyan

# ---- Determine entries to skip (prompts) ------------------------------------
$confirmedPlans = @()

if ($DryRun) {
  $confirmedPlans = $changePlan
}
else {
  foreach ($plan in $changePlan) {
    Write-Log -Message "Processing entry: $($plan.Origin)" -Color Cyan
    $skip = $false

    # TTL warning prompt (zone-level)
    if ($plan.HighTtlRecords.Count -gt 0) {
      Write-Log -Message "  WARNING: $($plan.HighTtlRecords.Count) existing record(s) have TTL > $highTtlThreshold seconds:" -Color Red
      foreach ($r in $plan.HighTtlRecords) {
        Write-Log -Message "    $($r.name) $($r.type) TTL=$($r.ttl)" -Color Red
      }
      Write-Log -Message '    Outstanding DNS propagation may cause stale values to persist.' -Color Red
      if (-not (Confirm-DestructiveOperation -Message "  Continue with '$($plan.Origin)' despite high TTLs?")) {
        Write-Log -Message "  -> Skipped '$($plan.Origin)' (user cancelled)." -Color Yellow
        $skip = $true
      }
    }

    # Nameserver change warning
    if (-not $skip -and $plan.DomainChanges.Count -gt 0) {
      $hasNsChange = ($plan.DomainChanges | Where-Object { $_.Field -eq 'nameServers' }).Count -gt 0
      if ($hasNsChange) {
        Write-Log -Message '  WARNING: Changing nameservers will delegate DNS authority to new providers.' -Color Red
      }
      if (-not (Confirm-DestructiveOperation -Message "  Apply domain-level changes for '$($plan.Origin)'?")) {
        Write-Log -Message "  -> Domain changes skipped for '$($plan.Origin)'." -Color Yellow
        $plan.DomainChanges = @()
      }
    }

    # Records replacement prompt
    if (-not $skip -and $plan.HasRecordReplace) {
      Write-Log -Message '  WARNING: This will REPLACE all existing resource records for this zone.' -Color Red
      if (-not (Confirm-DestructiveOperation -Message "  Replace all resource records for '$($plan.Origin)'?")) {
        Write-Log -Message "  -> Zone changes skipped for '$($plan.Origin)'." -Color Yellow
        $plan.ZoneChanges = @()
      }
    }

    # Main zone prompt
    if (-not $skip -and $plan.ZoneChanges.Count -gt 0 -and -not $plan.HasRecordReplace) {
      $changeDesc = ($plan.ZoneChanges | ForEach-Object { "$($_.Field): $($_.Current) -> $($_.Desired)" }) -join '; '
      if (-not (Confirm-DestructiveOperation -Message "  Apply zone changes for '$($plan.Origin)'?`n  $changeDesc")) {
        Write-Log -Message "  -> Zone changes skipped for '$($plan.Origin)'." -Color Yellow
        $plan.ZoneChanges = @()
      }
    }

    # Skip if nothing was confirmed
    if ($plan.DomainChanges.Count -eq 0 -and $plan.ZoneChanges.Count -eq 0) {
      Write-Log -Message "  -> No changes confirmed for '$($plan.Origin)'." -Color Yellow
    }
    else {
      $confirmedPlans += $plan
    }
  }
}

# ---- Phase 3: Apply confirmed changes ---------------------------------------
$results = @()
$appliedCount = 0

foreach ($plan in $confirmedPlans) {
  $origin = $plan.Origin

  # -- Apply domain-level changes --
  if ($plan.DomainChanges.Count -gt 0) {
    $nsChange = $plan.DomainChanges | Where-Object { $_.Field -eq 'nameServers' }
    $dnssecEnabledChange = $plan.DomainChanges | Where-Object { $_.Field -eq 'dnssec enabled' }
    $dnssecAutoChange = $plan.DomainChanges | Where-Object { $_.Field -eq 'dnssec auto' }
    $dnssecKeysChange = $plan.DomainChanges | Where-Object { $_.Field -eq 'dnssec keys' }
    $dnssecRolloverChange = $plan.DomainChanges | Where-Object { $_.Field -eq 'dnssec key rollover' }
    $domainEntry = $plan.DomainConfig

    # --- Nameserver changes: PUT /domain/{name} ---
    if ($nsChange) {
      $domainBody = @{ nameServers = $domainEntry.nameServers }

      Write-Log -Message "  Updating nameservers for: $origin ..." -Color Yellow
      if ($DryRun) {
        Write-Log -Message "    [DRY RUN] Would update nameservers for '$origin'" -Color Yellow
      }
      else {
        try {
          $jobInfo = Update-AutoDNSDomain -Name $origin -Body $domainBody
          if ($jobInfo) {
            Write-Log -Message "    -> Nameserver update submitted (Job ID: $($jobInfo.JobId))." -Color Green
          }
          else {
            Write-Log -Message "    -> Nameservers updated successfully." -Color Green
          }
        }
        catch {
          Write-Log -Message "    -> Nameserver update failed: $_" -Color Red
        }
      }
    }

    # --- DNSSEC changes: PUT /domain/{name}/_dnssec ---
    if ($dnssecEnabledChange -or $dnssecAutoChange -or $dnssecKeysChange) {
      $dnssecBody = @{}

      if ($domainEntry.dnssec.PSObject.Properties.Name -contains 'enabled') {
        $dnssecBody.dnssec = [bool]$domainEntry.dnssec.enabled
      }
      if ($domainEntry.dnssec.PSObject.Properties.Name -contains 'auto') {
        $dnssecBody.autoDnssec = [bool]$domainEntry.dnssec.auto
      }
      if ($domainEntry.dnssec.PSObject.Properties.Name -contains 'keys') {
        $dnssecBody.dnssecData = $domainEntry.dnssec.keys
      }

      Write-Log -Message "  Updating DNSSEC configuration for: $origin ..." -Color Yellow
      if ($DryRun) {
        Write-Log -Message "    [DRY RUN] Would update DNSSEC for '$origin'" -Color Yellow
      }
      else {
        try {
          $jobInfo = Invoke-AutoDNSRequest -Path "/domain/$origin/_dnssec" -Method PUT -Body $dnssecBody
          if ($jobInfo.data -and $jobInfo.data.Count -gt 0 -and $jobInfo.data[0].id) {
            Write-Log -Message "    -> DNSSEC update submitted (Job ID: $($jobInfo.data[0].id))." -Color Green
          }
          else {
            Write-Log -Message "    -> DNSSEC updated successfully." -Color Green
          }
        }
        catch {
          Write-Log -Message "    -> DNSSEC update failed: $_" -Color Red
        }
      }
    }

    # --- DNSSEC key rollover: PUT /domain/{name}/_autoDnssecKeyRollover ---
    if ($dnssecRolloverChange) {
      Write-Log -Message "  Triggering DNSSEC key rollover for: $origin ..." -Color Yellow
      $dnssecRolloverNote = @'

NOTE: Key rollover is an asynchronous process. Monitor the AutoDNS Web UI
      or check Job status to verify completion.

'@
      Write-Log -Message $dnssecRolloverNote -Color Cyan

      if ($DryRun) {
        Write-Log -Message "    [DRY RUN] Would trigger key rollover for '$origin'" -Color Yellow
      }
      else {
        try {
          Invoke-AutoDNSRequest -Path "/domain/$origin/_autoDnssecKeyRollover" -Method PUT | Out-Null
          Write-Log -Message "    -> Key rollover triggered for '$origin'." -Color Green
        }
        catch {
          Write-Log -Message "    -> Key rollover failed: $_" -Color Red
        }
      }
    }
  }

  # -- Apply zone-level changes --
  if ($plan.ZoneChanges.Count -gt 0) {
    $vns = $plan.Vns

    $putBody = @{
      origin = $origin
      virtualNameServer = $vns
      main = if ($plan.DesiredPayload.Contains('main')) { $plan.DesiredPayload.main } else { $plan.CurrentZone.main }
      wwwInclude = if ($plan.DesiredPayload.Contains('wwwInclude')) { $plan.DesiredPayload.wwwInclude } else { [bool]$plan.CurrentZone.wwwInclude }
      dnssec = if ($plan.DesiredPayload.Contains('dnssec')) { $plan.DesiredPayload.dnssec } else { [bool]$plan.CurrentZone.dnssec }
      resourceRecords = if ($plan.DesiredPayload.Contains('resourceRecords')) { $plan.DesiredPayload.resourceRecords } else { $plan.CurrentZone.resourceRecords }
      nameServers = $plan.CurrentZone.nameServers
      soa = $plan.CurrentZone.soa
    }

    if ($DryRun) {
      Write-Log -Message "  [DRY RUN] Would update zone '$origin' on $vns" -Color Yellow
      $results += @{
        Origin = $origin
        Status = 'DryRun'
        Changes = $plan.ZoneChanges
      }
      continue
    }

    try {
      Invoke-AutoDNSRequest -Path "/zone/$origin/$vns" -Method PUT -Body $putBody
      Write-Log -Message "  -> Zone '$origin' updated successfully." -Color Green
      $appliedCount++
      $results += @{
        Origin = $origin
        Status = 'Updated'
        Changes = $plan.ZoneChanges
      }
    }
    catch {
      Write-Log -Message "  -> Failed to update zone '$origin': $_" -Color Red
      $results += @{
        Origin = $origin
        Status = 'Failed'
        Error = $_
      }
      continue
    }

    # DNSSEC domain-level update if zone DNSSEC toggled
    if ($plan.DesiredPayload.Contains('dnssec') -and $plan.DesiredPayload.dnssec -ne [bool]$plan.CurrentZone.dnssec) {
      Write-Log -Message '  -> Updating DNSSEC configuration at domain level ...' -Color Yellow
      try {
        $dnssecBody = @{ dnssec = $plan.DesiredPayload.dnssec }
        Invoke-AutoDNSRequest -Path "/domain/$origin/_dnssec" -Method PUT -Body $dnssecBody | Out-Null
        Write-Log -Message '  -> DNSSEC updated at domain level.' -Color Green
      }
      catch {
        Write-Log -Message "  -> Warning: DNSSEC domain update failed: $_" -Color Yellow
      }
    }
  }
}

# ---- Summary table ----------------------------------------------------------
Write-Log -Message "`n========== RESULTS ==========" -Color Cyan
foreach ($r in $results) {
  $statusColor = switch ($r.Status) {
    'Updated' { 'Green' }
    'DryRun' { 'Yellow' }
    'Failed' { 'Red' }
    default { 'Gray' }
  }
  $changeFields = ($r.Changes | ForEach-Object { $_.Field }) -join ', '
  Write-Log -Message "  $($r.Origin) [$($r.Status)] $changeFields" -Color $statusColor
}
Write-Log -Message "==============================" -Color Cyan

if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no changes were made." -Color Yellow
}
else {
  Write-Log -Message "`nDone. $appliedCount zone(s) updated." -Color Cyan
}

if ($PassThru -or $DryRun) {
  $results
}
