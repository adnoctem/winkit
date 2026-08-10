#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Resolves the MAC address of a target machine from its hostname or IP address.

.DESCRIPTION
  Find-MACAddress looks up the hardware address of a target device using a
  tiered resolution strategy:

    1. DHCP lease lookup on a Microsoft DHCP Server (Get-DhcpServerv4Lease).
       This is the only source that can resolve a machine that is currently
       powered off, as long as its lease has not expired. It requires the
       DhcpServer module (RSAT DHCP Tools or the DHCP Server role) and remote
       management rights (typically admin) on the DHCP server.
    2. Local ARP/neighbor cache lookup (Get-NetNeighbor). Works with any DHCP
       server since it only reads the local machine's own neighbor table, but
       the target must have been reachable from this host recently.
    3. ICMP probing. A single lightweight echo request to the resolved address
       by default; with -ProbeSubnet the entire local subnet is swept to
       repopulate the ARP cache, after which the lookup is retried.

  The DHCP server queried by default is the one the local machine currently
  leases from (Win32_NetworkAdapterConfiguration.DHCPServer of the default
  network adapter). Override with -DhcpServerName. Non-Microsoft DHCP servers
  (router/firewall DHCP) cannot be queried through Get-DhcpServerv4Lease; for
  those, tiers 2 and 3 apply.

  The resolved MAC is printed in colon-separated form ready for
  Send-WOLPacket. Use -Wake to chain the result directly into
  Send-WOLPacket.ps1 (located next to this script) and send the magic packet.

.PARAMETER HostName
  Target machine hostname, NetBIOS name, or FQDN. Mutually exclusive with
  -IPAddress.

.PARAMETER IPAddress
  Target machine IPv4 or IPv6 address. Mutually exclusive with -HostName.

.PARAMETER DhcpServerName
  Microsoft DHCP server to query for lease information. Defaults to the DHCP
  server that issued the local machine's lease.

.PARAMETER ScopeId
  Optional DHCP scope network ID (e.g. 192.168.1.0) to restrict the lease
  lookup. When omitted, all scopes on the DHCP server are checked.

.PARAMETER ProbeSubnet
  Sweep the entire local subnet with ICMP echo requests to repopulate the ARP
  cache before retrying the lookup. A lightweight probe of the target address
  alone is performed by default. Sweeps are limited to subnets with at most
  2048 usable host addresses (larger than /22).

.PARAMETER Wake
  Chain the resolved MAC address into Send-WOLPacket.ps1 and send a
  Wake-on-LAN magic packet.

.PARAMETER DryRun
  Print the resolution plan without performing any lookups or sending packets.

.PARAMETER PassThru
  Return the resolution result(s) as structured objects.

.EXAMPLE
  PS> ./Find-MACAddress.ps1 -HostName 'WORKSTATION42'
  Resolves the MAC of a machine by hostname via DHCP lease, then the ARP cache.

.EXAMPLE
  PS> ./Find-MACAddress.ps1 -IPAddress '192.168.1.77'
  Resolves the MAC of a machine by IP address.

.EXAMPLE
  PS> ./Find-MACAddress.ps1 -HostName 'WS42.corp.example' -DhcpServerName 'dhcp01' -Wake
  Queries the specified DHCP server directly and sends a Wake-on-LAN packet.

.EXAMPLE
  PS> ./Find-MACAddress.ps1 -IPAddress '192.168.1.77' -ProbeSubnet
  Probes the target and sweeps the local subnet to populate the ARP cache.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'DhcpServerName', Justification = 'Used by nested DHCP server helper through script scope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ScopeId', Justification = 'Used by nested DHCP lease helper through script scope.')]
[CmdletBinding()]
param (
  [Parameter(ParameterSetName = 'HostName', Position = 0, Mandatory = $true)]
  [string]
  $HostName,

  [Parameter(ParameterSetName = 'IPAddress', Position = 0, Mandatory = $true)]
  [string]
  $IPAddress,

  [Parameter(Mandatory = $false)]
  [string]
  $DhcpServerName,

  [Parameter(Mandatory = $false)]
  [System.Net.IPAddress]
  $ScopeId,

  [Parameter(Mandatory = $false)]
  [switch]
  $ProbeSubnet,

  [Parameter(Mandatory = $false)]
  [switch]
  $Wake,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

$Host.UI.RawUI.WindowTitle = 'winkit - Find-MACAddress'

if ($DryRun) {
  Write-Log -Message "DRY RUN - no lookups will be performed and no packets will be sent`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_resolutionResults = New-Object System.Collections.ArrayList
$_seenMacs = @{}

# ---- Script-local helpers ----------------------------------------------------

function Format-MacAddress {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Value
  )

  $_normalized = $Value -replace '[^0-9A-Fa-f]', ''
  if ($_normalized.Length -ne 12) {
    return $null
  }

  return (($_normalized -replace '(..)', '$1:').TrimEnd(':')).ToUpperInvariant()
}

function Get-EffectiveDhcpServer {
  [OutputType([string])]
  [CmdletBinding()]
  param()

  if (-not [string]::IsNullOrWhiteSpace($DhcpServerName)) {
    return $DhcpServerName
  }

  try {
    $_adapter = Get-DefaultNetworkAdapter -Required
  }
  catch {
    Write-Log -Message "  Could not resolve the default network adapter: $_" -Color Gray
    return $null
  }

  $_dhcpServer = @($_adapter.CimConfig.DHCPServer | Where-Object { $_ -notmatch ':' } | Select-Object -First 1)
  if ($_dhcpServer.Count -eq 0 -or [string]::IsNullOrWhiteSpace($_dhcpServer[0])) {
    Write-Log -Message '  The default adapter uses a static IP - DHCP lease lookup skipped. Use -DhcpServerName to query a specific DHCP server.' -Color Gray
    return $null
  }

  return $_dhcpServer[0].Trim()
}

function Get-DhcpLeaseRecord {
  [OutputType([System.Collections.ArrayList])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$ServerName,

    [string[]]$Addresses,

    [string]$TargetHostName
  )

  if (-not (Get-Module -ListAvailable -Name DhcpServer -ErrorAction SilentlyContinue)) {
    Write-Log -Message '  DhcpServer module not available (RSAT DHCP Tools or the DHCP Server role is required) - DHCP lease lookup skipped.' -Color Gray
    Add-OperationResult -Results $_results -Target $TargetHostName -Source 'DhcpServer' -Action 'QueryLease' -Status 'Skipped' -SkippedReason 'DhcpServerModuleMissing'
    return New-Object System.Collections.ArrayList
  }

  Import-Module DhcpServer -ErrorAction SilentlyContinue

  $_scopes = @()
  if ($ScopeId) {
    $_scopes = @($ScopeId.ToString())
  }
  else {
    try {
      $_scopes = @(Get-DhcpServerv4Scope -ComputerName $ServerName -ErrorAction Stop |
          ForEach-Object { $_.ScopeId.ToString() } |
          Sort-Object -Unique)
    }
    catch {
      $_elevationHint = ''
      try {
        if (-not (Test-Elevation)) {
          $_elevationHint = ' Running elevated or with DCOM permissions on the DHCP server may be required.'
        }
      }
      catch {
        $_elevationHint = ''
      }
      Write-Log -Message "  Could not enumerate DHCP scopes on '$ServerName': $_$_elevationHint" -Color Red
      Add-OperationResult -Results $_results -Target $ServerName -Source 'DhcpServer' -Action 'QueryLease' -Status 'Failed' -Detail $_.Exception.Message
      return New-Object System.Collections.ArrayList
    }
  }

  $_leaseAddresses = @($Addresses | Where-Object { $_ -and $_ -notmatch ':' })

  $_targetNameLeaf = if ($TargetHostName) { $TargetHostName.Split('.')[0] } else { $null }

  $_found = New-Object System.Collections.ArrayList
  foreach ($_scope in $_scopes) {
    try {
      if ($_leaseAddresses.Count -gt 0) {
        $_leases = @(Get-DhcpServerv4Lease -ComputerName $ServerName -ScopeId $_scope -IPAddress $_leaseAddresses -ErrorAction SilentlyContinue)
      }
      else {
        $_leases = @(Get-DhcpServerv4Lease -ComputerName $ServerName -ScopeId $_scope -ErrorAction SilentlyContinue |
            Where-Object {
              $_.HostName -and (
                $_.HostName.TrimEnd('.').Split('.')[0] -eq $_targetNameLeaf -or
                $_.HostName.TrimEnd('.') -eq $TargetHostName.TrimEnd('.')
              )
            })
      }
    }
    catch {
      continue
    }

    foreach ($_lease in $_leases) {
      $_mac = Format-MacAddress -Value $_lease.ClientId
      if ($_mac) {
        [void]$_found.Add([PSCustomObject]@{
            IPAddress = $_lease.IPAddress.ToString()
            MacAddress = $_mac
            HostName = $_lease.HostName
            ScopeId = $_scope
          })
      }
    }
  }

  return $_found
}

function Get-NeighborMacAddress {
  [OutputType([System.Collections.Generic.List[PSCustomObject]])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Addresses
  )

  $_statePriority = @{
    'Reachable' = 0
    'Permanent' = 1
    'Stale' = 2
    'Delay' = 3
    'Probe' = 4
    'Unreachable' = 5
  }

  $_found = New-Object System.Collections.Generic.List[PSCustomObject]
  foreach ($_address in $Addresses) {
    $_family = if ($_address -match ':') { 'IPv6' } else { 'IPv4' }

    $_neighbors = @(Get-NetNeighbor -IPAddress $_address -AddressFamily $_family -ErrorAction SilentlyContinue |
        Where-Object { $_.State -ne 'Incomplete' -and $_.LinkLayerAddress })
    if ($_neighbors.Count -eq 0) {
      continue
    }

    $_best = $_neighbors |
      Sort-Object -Property @{ Expression = { if ($_statePriority.ContainsKey($_.State)) { $_statePriority[$_.State] } else { 99 } } } |
      Select-Object -First 1

    $_mac = Format-MacAddress -Value $_best.LinkLayerAddress
    if ($_mac -and $_mac -ne '00:00:00:00:00:00') {
      [void]$_found.Add([PSCustomObject]@{
          IPAddress = $_address
          MacAddress = $_mac
          State = $_best.State
        })
    }
  }

  return $_found
}

function Test-TargetReachable {
  [OutputType([bool])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Address
  )

  try {
    return [bool](Test-Connection -ComputerName $Address -Count 1 -Quiet -ErrorAction SilentlyContinue)
  }
  catch {
    return $false
  }
}

function Invoke-SubnetProbe {
  [OutputType([System.Collections.ArrayList])]
  [CmdletBinding()]
  param(
    [int]$TimeoutMilliseconds = 300
  )

  try {
    $_adapter = Get-DefaultNetworkAdapter -Required
  }
  catch {
    Write-Log -Message "  Could not resolve the default network adapter: $_" -Color Gray
    return New-Object System.Collections.ArrayList
  }

  $_localIp = Get-IPAddress -Adapter $_adapter
  $_cidr = Get-NetworkPrefixCIDR -Adapter $_adapter
  $_mask = Get-SubnetMask -Adapter $_adapter
  if (-not $_cidr -or -not $_mask) {
    Write-Log -Message '  Could not determine the local subnet - subnet sweep skipped.' -Color Yellow
    return New-Object System.Collections.ArrayList
  }

  $_cidrParts = $_cidr -split '/'
  $_prefixLength = [int]$_cidrParts[1]
  $_hostCountValue = [math]::Pow(2, 32 - $_prefixLength) - 2

  if ($_hostCountValue -lt 1) {
    Write-Log -Message "  Local subnet $_cidr has no usable host addresses - subnet sweep skipped." -Color Yellow
    return New-Object System.Collections.ArrayList
  }

  if ($_hostCountValue -gt 2048) {
    Write-Log -Message "  Local subnet $_cidr exceeds the 2048 host sweep limit - subnet sweep skipped." -Color Yellow
    return New-Object System.Collections.ArrayList
  }

  $_hostCount = [int]$_hostCountValue
  $_netBytes = [System.Net.IPAddress]::Parse($_cidrParts[0]).GetAddressBytes()

  $_hostList = New-Object System.Collections.Generic.List[string]
  for ($_hostOffset = 1; $_hostOffset -le $_hostCount; $_hostOffset++) {
    $_candidate = [byte[]]@($_netBytes[0], $_netBytes[1], $_netBytes[2], $_netBytes[3])
    $_carry = $_hostOffset
    for ($_byteIndex = 3; $_byteIndex -ge 0 -and $_carry -gt 0; $_byteIndex--) {
      $_sum = $_candidate[$_byteIndex] + ($_carry -band 0xFF)
      $_candidate[$_byteIndex] = [byte]($_sum -band 0xFF)
      $_carry = ($_sum -shr 8) + ($_carry -shr 8)
    }

    $_candidateIp = "$($_candidate[0]).$($_candidate[1]).$($_candidate[2]).$($_candidate[3])"
    if ($_candidateIp -eq $_localIp) {
      continue
    }
    $_hostList.Add($_candidateIp)
  }

  Write-Log -Message "  Sweeping $($_hostList.Count) host address(es) in $_cidr..." -Color Yellow

  $_pingTasks = @()
  foreach ($_hostIp in $_hostList) {
    $_pinger = New-Object System.Net.NetworkInformation.Ping
    $_pingTasks += $_pinger.SendPingAsync($_hostIp, $TimeoutMilliseconds)
  }

  $null = [System.Threading.Tasks.Task]::WaitAll($_pingTasks)

  $_responders = New-Object System.Collections.ArrayList
  for ($_taskIndex = 0; $_taskIndex -lt $_pingTasks.Count; $_taskIndex++) {
    try {
      if ($_pingTasks[$_taskIndex].Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion -and
        $_pingTasks[$_taskIndex].Result.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
        [void]$_responders.Add($_hostList[$_taskIndex])
      }
    }
    catch {
      continue
    }
  }

  return $_responders
}

function Invoke-WakeOnLan {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$MacAddress,

    [string]$TargetName
  )

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would invoke Send-WOLPacket.ps1 -MacAddress '$MacAddress'." -Color Yellow
    Add-OperationResult -Results $_results -Target $TargetName -Source 'WakeOnLan' -Action 'Wake' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  $_wolScript = Join-Path -Path $PSScriptRoot -ChildPath 'Send-WOLPacket.ps1'
  if (-not (Test-Path -LiteralPath $_wolScript -PathType Leaf)) {
    Write-Log -Message '  Send-WOLPacket.ps1 not found next to this script - skipping wake.' -Color Yellow
    Add-OperationResult -Results $_results -Target $TargetName -Source 'WakeOnLan' -Action 'Wake' -Status 'Skipped' -SkippedReason 'SendWOLPacketMissing'
    return
  }

  Write-Log -Message "Sending Wake-on-LAN packet for '$TargetName' ($MacAddress) via Send-WOLPacket.ps1..." -Color Yellow
  & $_wolScript -MacAddress $MacAddress

  if ($LASTEXITCODE -ne 0) {
    Write-Log -Message '  Send-WOLPacket.ps1 reported an error.' -Color Red
    Add-OperationResult -Results $_results -Target $TargetName -Source 'WakeOnLan' -Action 'Wake' -Status 'Failed' -Detail "Send-WOLPacket.ps1 exited with code $LASTEXITCODE"
  }
  else {
    Add-OperationResult -Results $_results -Target $TargetName -Source 'WakeOnLan' -Action 'Wake' -Status 'Completed' -Detail $MacAddress
  }
}

function Add-MacResolution {
  [OutputType([bool])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$IPAddress,

    [Parameter(Mandatory = $true)]
    [string]$MacAddress,

    [Parameter(Mandatory = $true)]
    [ValidateSet('DhcpLease', 'ArpCache')]
    [string]$Source,

    [string]$HostName
  )

  if ($_seenMacs.ContainsKey($MacAddress)) {
    return $false
  }

  $_seenMacs[$MacAddress] = $true
  [void]$_resolutionResults.Add([PSCustomObject]@{
      HostName = if ($HostName) { $HostName } else { $null }
      IPAddress = $IPAddress
      MacAddress = $MacAddress
      Source = $Source
    })
  return $true
}

# ---- Target resolution -------------------------------------------------------

$_targetName = if ($HostName) { $HostName } else { $IPAddress }
$_candidateAddresses = @()

if ($IPAddress) {
  $_isValidAddress = (Test-IPv4Address -Address $IPAddress) -or (Test-IPv6Address -Address $IPAddress)
  if (-not $_isValidAddress) {
    Write-Log -Message "'$IPAddress' is not a valid IPv4 or IPv6 address." -Color Red
    exit 1
  }
  $_candidateAddresses = @($IPAddress)
}
else {
  try {
    $_dnsAddresses = [System.Net.Dns]::GetHostAddresses($HostName)
  }
  catch {
    Write-Log -Message "Could not resolve hostname '$HostName': $_" -Color Red
    Add-OperationResult -Results $_results -Target $HostName -Source 'Dns' -Action 'Resolve' -Status 'Failed' -Detail $_.Exception.Message
    exit 1
  }

  $_ipv4Addresses = @($_dnsAddresses | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | Select-Object -ExpandProperty IPAddressToString)
  $_ipv6Addresses = @($_dnsAddresses | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6 } | Select-Object -ExpandProperty IPAddressToString)

  if ($_ipv4Addresses.Count -eq 0 -and $_ipv6Addresses.Count -eq 0) {
    Write-Log -Message "Hostname '$HostName' resolved to no addresses." -Color Red
    Add-OperationResult -Results $_results -Target $HostName -Source 'Dns' -Action 'Resolve' -Status 'Failed' -Detail 'NoAddressFound'
    exit 1
  }

  $_candidateAddresses = @($_ipv4Addresses) + @($_ipv6Addresses)
  Write-Log -Message "Resolved '$HostName' -> $($_candidateAddresses -join ', ')" -Color Gray
}

$_effectiveDhcpServer = Get-EffectiveDhcpServer

# ---- Dry run ----------------------------------------------------------------

if ($DryRun) {
  Write-Log -Message "`nPlan for '$_targetName':" -Color Yellow
  Write-Log -Message "  -> target address(es): $($_candidateAddresses -join ', ')" -Color Gray
  if ($_effectiveDhcpServer) {
    Write-Log -Message "  -> would query DHCP server '$_effectiveDhcpServer' for lease information" -Color Gray
  }
  else {
    Write-Log -Message '  -> no default DHCP server detected - DHCP lease lookup would be skipped' -Color Gray
  }
  Write-Log -Message '  -> would check the local ARP/neighbor cache for each address' -Color Gray
  if ($ProbeSubnet) {
    Write-Log -Message '  -> would probe the target and sweep the local subnet to populate the ARP cache' -Color Gray
  }
  else {
    Write-Log -Message '  -> would probe each target address with one ICMP echo request' -Color Gray
  }
  if ($Wake) {
    Write-Log -Message '  -> would send a Wake-on-LAN packet via Send-WOLPacket.ps1' -Color Gray
  }
  Add-OperationResult -Results $_results -Target $_targetName -Source 'Resolution' -Action 'ResolveMac' -Status 'Skipped' -Detail 'DryRun'
  Write-Log -Message "`nDRY RUN COMPLETE - no lookups were performed and no packets were sent." -Color Yellow

  if ($PassThru) {
    $_resolutionResults
  }
  return
}

# ---- Tier 1: DHCP lease lookup ----------------------------------------------

if ($_effectiveDhcpServer) {
  Write-Log -Message "Querying DHCP server '$_effectiveDhcpServer' for lease information..." -Color Yellow
  $_leases = Get-DhcpLeaseRecord -ServerName $_effectiveDhcpServer -Addresses $_candidateAddresses -TargetHostName $_targetName

  foreach ($_lease in $_leases) {
    $_added = Add-MacResolution -IPAddress $_lease.IPAddress -MacAddress $_lease.MacAddress -Source 'DhcpLease' -HostName $_lease.HostName
    if ($_added) {
      Write-Log -Message "  -> DHCP lease: $($_lease.IPAddress) = $($_lease.MacAddress) ($($_lease.HostName))" -Color Green
    }
  }

  if ($_leases.Count -gt 0) {
    Add-OperationResult -Results $_results -Target $_targetName -Source 'DhcpServer' -Action 'QueryLease' -Status 'Completed' -Detail "$($_leases.Count) lease record(s) matched."
  }
}

# ---- Tier 2: ARP / neighbor cache lookup ------------------------------------

if ($_resolutionResults.Count -eq 0) {
  Write-Log -Message 'No DHCP lease match - checking the local ARP/neighbor cache...' -Color Yellow
  $_neighbors = Get-NeighborMacAddress -Addresses $_candidateAddresses

  foreach ($_neighbor in $_neighbors) {
    $_added = Add-MacResolution -IPAddress $_neighbor.IPAddress -MacAddress $_neighbor.MacAddress -Source 'ArpCache'
    if ($_added) {
      Write-Log -Message "  -> ARP cache: $($_neighbor.IPAddress) = $($_neighbor.MacAddress) (state: $($_neighbor.State))" -Color Green
    }
  }

  Add-OperationResult -Results $_results -Target $_targetName -Source 'NeighborCache' -Action 'LookupMac' -Status $(if ($_neighbors.Count -gt 0) { 'Completed' } else { 'Skipped' }) -Detail "$($_neighbors.Count) neighbor entry(ies) matched."
}

# ---- Tier 3a: lightweight probe of the target -------------------------------

if ($_resolutionResults.Count -eq 0) {
  Write-Log -Message 'No ARP entry - probing the target address(es) with ICMP...' -Color Yellow
  $_probed = $false
  foreach ($_address in $_candidateAddresses) {
    if (Test-TargetReachable -Address $_address) {
      $_probed = $true
      Write-Log -Message "  -> $_address responded." -Color Green
    }
  }

  if ($_probed) {
    Write-Log -Message '  Re-checking the ARP/neighbor cache after the probe...' -Color Gray
    $_neighbors = Get-NeighborMacAddress -Addresses $_candidateAddresses
    foreach ($_neighbor in $_neighbors) {
      $_added = Add-MacResolution -IPAddress $_neighbor.IPAddress -MacAddress $_neighbor.MacAddress -Source 'ArpCache'
      if ($_added) {
        Write-Log -Message "  -> ARP cache: $($_neighbor.IPAddress) = $($_neighbor.MacAddress) (state: $($_neighbor.State))" -Color Green
      }
    }
  }

  Add-OperationResult -Results $_results -Target $_targetName -Source 'IcmpProbe' -Action 'Probe' -Status $(if ($_probed) { 'Completed' } else { 'Skipped' }) -Detail "Probed $($_candidateAddresses.Count) address(es)."
}

# ---- Tier 3b: subnet-wide sweep ---------------------------------------------

if ($ProbeSubnet -and $_resolutionResults.Count -eq 0) {
  Write-Log -Message 'Probing the local subnet to populate the ARP cache...' -Color Yellow
  $_responders = Invoke-SubnetProbe
  Add-OperationResult -Results $_results -Target $_targetName -Source 'IcmpSweep' -Action 'Probe' -Status $(if ($_responders.Count -gt 0) { 'Completed' } else { 'Empty' }) -Detail "$($_responders.Count) host(s) responded."

  if ($_responders.Count -gt 0) {
    Write-Log -Message '  Re-checking the ARP/neighbor cache after the sweep...' -Color Gray
    $_neighbors = Get-NeighborMacAddress -Addresses $_candidateAddresses
    foreach ($_neighbor in $_neighbors) {
      $_added = Add-MacResolution -IPAddress $_neighbor.IPAddress -MacAddress $_neighbor.MacAddress -Source 'ArpCache'
      if ($_added) {
        Write-Log -Message "  -> ARP cache: $($_neighbor.IPAddress) = $($_neighbor.MacAddress) (state: $($_neighbor.State))" -Color Green
      }
    }
  }
}

# ---- Outcome ----------------------------------------------------------------

if ($_resolutionResults.Count -eq 0) {
  Write-Log -Message "`nNo MAC address found for '$_targetName'." -Color Red
  Write-Log -Message 'The machine may currently be powered off with no valid DHCP lease, or it is not reachable from this host.' -Color Yellow
  Write-Log -Message 'Tip: query the lease via a Microsoft DHCP server with -DhcpServerName, or wake the machine once so its ARP entry is re-created.' -Color Gray
  Add-OperationResult -Results $_results -Target $_targetName -Source 'Resolution' -Action 'ResolveMac' -Status 'Failed' -Detail 'NoMacFound'
}
else {
  foreach ($_found in $_resolutionResults) {
    $_displayName = if ($_found.HostName) { $_found.HostName } else { $_targetName }
    Write-Log -Message "`nMAC address for '$_displayName' ($($_found.IPAddress)): $($_found.MacAddress)  [source: $($_found.Source)]" -Color Green
  }

  Write-Log -Message "`nReady to wake: Send-WOLPacket -MacAddress '$($_resolutionResults[0].MacAddress)'" -Color Cyan
  Add-OperationResult -Results $_results -Target $_targetName -Source 'Resolution' -Action 'ResolveMac' -Status 'Completed' -Detail $_resolutionResults[0].MacAddress

  if ($Wake) {
    Invoke-WakeOnLan -MacAddress $_resolutionResults[0].MacAddress -TargetName $_targetName
  }
}

# ---- Summary ----------------------------------------------------------------

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Find-MACAddress'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_resolutionResults
}
