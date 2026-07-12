#Requires -Version 5.1

<#
.SYNOPSIS
  Generates WireGuard peer configuration files based on a template and peer data from pfSense.

.DESCRIPTION
  Reads peer configuration either from a pfSense tunnel export file or directly from the pfSense REST API
  (when the pfSense-pkg-RESTAPI package is installed). Matches peers to local user folders, replacing
  placeholders in a template with actual values such as private keys, public keys, pre-shared keys,
  and allowed IPs.

.PARAMETER ServerPublicKey
  The public key of the WireGuard server. When using -APIExport, this can be omitted and will be
  fetched from the first tunnel returned by the API.

.PARAMETER RootDirectory
  The root directory containing subfolders for each user. Each subfolder should contain a 'privatekey' file.

.PARAMETER TemplatePath
  Path to the template configuration file relative to RootDirectory (default: wg.conf.tpl). Falls back
  to a built-in template if not found.

.PARAMETER PfSensePath
  Path to the pfSense export file relative to RootDirectory (default: pfsense.conf). Only used in
  file-based mode (when -APIExport is not specified).

.PARAMETER ConfigurationInterfaceDNSServers
  DNS servers to be inserted into generated configs (defaults: local network DNS).

.PARAMETER ConfigurationPeerAllowedIPs
  Allowed IP ranges for the [Peer] section (defaults: the local subnets).

.PARAMETER ConfigurationPeerEndpoint
  The WireGuard server endpoint (e.g., hq.delta4x4.net:57173).

.PARAMETER ConfigurationPersistentKeepalive
  Persistent keepalive interval in seconds (default: 25).

.PARAMETER APIExport
  When set, fetches WireGuard peer data from the pfSense REST API instead of reading a local export file.
  Requires -PfSenseHost and -PfSenseCredential.

.PARAMETER PfSenseHost
  The hostname or IP address of the pfSense firewall (required with -APIExport).

.PARAMETER PfSenseCredential
  A PSCredential object for HTTP Basic authentication against the pfSense REST API (required with -APIExport).
  Obtain via Get-Credential or pass a credential object.

.PARAMETER SkipCertificateCheck
  Skips TLS certificate validation (useful for self-signed certificates on pfSense).

.EXAMPLE
  PS> ./New-WGPeerConfig.ps1 -ServerPublicKey "SERVER_PUBLIC_KEY_HERE" -RootDirectory "C:\WGConfigs"
  Generates configs from a local pfsense.conf export file using the built-in template.

.EXAMPLE
  PS> ./New-WGPeerConfig.ps1 -APIExport -PfSenseHost "pfsense.example.com" -PfSenseCredential (Get-Credential) -RootDirectory "C:\WGConfigs"
  Fetches peer data live from the pfSense REST API.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

param (
  [Parameter(Position = 0, Mandatory = $false, HelpMessage = "The public key of the WireGuard server.")]
  [string]$ServerPublicKey,

  [Parameter(Position = 1, Mandatory = $true, HelpMessage = "The root directory containing user subfolders.")]
  [string]$RootDirectory,

  [Parameter(Position = 2, Mandatory = $false, HelpMessage = "The path to the template configuration file (default: wg.conf.tpl).")]
  [string]$TemplatePath = "wg.conf.tpl",

  [Parameter(Position = 3, Mandatory = $false, HelpMessage = "The path to the pfSense export file (default: pfsense.conf).")]
  [string]$PfSensePath = "pfsense.conf",

  [Parameter(Position = 4, Mandatory = $false, HelpMessage = "The DNS servers to be used in the configuration.")]
  [string[]]$ConfigurationInterfaceDNSServers = @("192.168.1.3", "192.168.99.254", "192.168.1.254"),

  [Parameter(Position = 5, Mandatory = $false, HelpMessage = "The peer allowed IP addresses.")]
  [string[]]$ConfigurationPeerAllowedIPs = @("192.168.99.0/24", "192.168.0.0/23"),

  [Parameter(Position = 6, Mandatory = $false, HelpMessage = "The peer endpoint (e.g., hq.delta4x4.net:57173).")]
  [string]$ConfigurationPeerEndpoint = "hq.delta4x4.net:57173",

  [Parameter(Position = 7, Mandatory = $false, HelpMessage = "The persistent keepalive interval in seconds (default: 25).")]
  [int]$ConfigurationPersistentKeepalive = 25,

  [Parameter(Mandatory = $false, HelpMessage = "Fetch peer data from pfSense REST API instead of a local file.")]
  [switch]$APIExport,

  [Parameter(Mandatory = $false, HelpMessage = "The pfSense hostname or IP address (required with -APIExport).")]
  [string]$PfSenseHost,

  [Parameter(Mandatory = $false, HelpMessage = "Credentials for HTTP Basic authentication against the pfSense REST API.")]
  [pscredential]$PfSenseCredential,

  [Parameter(Mandatory = $false, HelpMessage = "Skip TLS certificate validation (for self-signed certificates).")]
  [switch]$SkipCertificateCheck
)

Import-Module PSFoundation -Force

#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

$builtInTemplate = @"
[Interface]
PrivateKey = <privatekey>
Address = <allowed_ips>
DNS = <dns_servers>

[Peer]
PublicKey = <pfsense_publickey>
PresharedKey = <pre-shared-key>
AllowedIPs = <peer_allowed_ips>
Endpoint = <peer_endpoint>
PersistentKeepalive = <persistent_keepalive>
"@

# ---- Helper functions -------------------------------------------------

function Read-PfSenseFileExport {
  <#
  .SYNOPSIS
    Parses a pfSense WireGuard tunnel export file into a peer lookup table.
  .PARAMETER Path
    Full path to the pfsense.conf export file.
  .RETURNS
    A hashtable keyed by peer name, with values @{ PSK = "..."; IP = "..." }.
  #>
  param([string]$Path)

  $peerLookup = @{}
  $content = Get-Content -Path $Path -Raw

  $regex = "(?ms)# Peer:\s*(.*?)\s*\[Peer\](.*?)(?=# Peer:|\z)"
  $rgx_matches = [regex]::Matches($content, $regex)

  foreach ($m in $rgx_matches) {
    $name = $m.Groups[1].Value.Trim()
    $block = $m.Groups[2].Value

    $psk = if ($block -match "PresharedKey\s*=\s*(.*)") { $Matches[1].Trim() } else { "" }
    $ip = if ($block -match "AllowedIPs\s*=\s*(.*)") { $Matches[1].Trim() } else { "" }

    $peerLookup[$name] = @{
      PSK = $psk
      IP = $ip
    }
  }

  return $peerLookup
}

function Read-PfSenseAPIExport {
  <#
  .SYNOPSIS
    Fetches WireGuard peer data from the pfSense REST API.
  .PARAMETER PfSenseHostName
    pfSense hostname or IP address.
  .PARAMETER Credential
    PSCredential object for HTTP Basic authentication.
  .PARAMETER SkipCertCheck
    Whether to skip TLS certificate validation.
  .RETURNS
    A hashtable keyed by peer description (name), with values @{ PSK = "..."; IP = "..." }.
  #>
  param(
    [string]$PfSenseHostName,
    [pscredential]$Credential,
    [switch]$SkipCertCheck
  )

  $pair = "$($Credential.UserName):$($Credential.GetNetworkCredential().Password)"
  $encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
  $headers = @{ Authorization = "Basic $encodedCreds" }

  $restParams = @{
    Uri = "https://$PfSenseHostName/api/v2/vpn/wireguard/peers"
    Method = 'GET'
    Headers = $headers
    ContentType = 'application/json'
  }

  if ($SkipCertCheck) {
    if ($PSVersionTable.PSVersion.Major -ge 6) {
      $restParams['SkipCertificateCheck'] = $true
    }
    else {
      $callback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
      [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
  }

  try {
    $response = Invoke-RestMethod @restParams
  }
  catch {
    throw "Failed to fetch WireGuard peers from pfSense API at $PfSenseHostName`: $_"
  }
  finally {
    if ($SkipCertCheck -and $PSVersionTable.PSVersion.Major -lt 6) {
      [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $callback
    }
  }

  $peers = if ($response.data) { $response.data } else { $response }

  $peerLookup = @{}

  foreach ($peer in $peers) {
    $name = $peer.descr
    if ([string]::IsNullOrWhiteSpace($name)) { continue }

    $psk = ""
    if ($peer.PSObject.Properties.Name -contains 'presharedkey' -and [string]::IsNullOrWhiteSpace($peer.presharedkey)) {
      $psk = $peer.presharedkey
    }

    $cidrList = @()
    if ($peer.allowedips -and $peer.allowedips.Count -gt 0) {
      foreach ($ipObj in $peer.allowedips) {
        $cidrList += "$($ipObj.address)/$($ipObj.mask)"
      }
    }
    $ip = $cidrList -join ','

    $peerLookup[$name] = @{
      PSK = $psk
      IP = $ip
    }
  }

  if ($peerLookup.Count -eq 0) {
    throw "No WireGuard peers were returned by the pfSense REST API. Ensure peers are configured."
  }

  return $peerLookup
}

# ----------------------------------------------------------------------

# Paths
$templateFullPath = Join-Path $RootDirectory $TemplatePath
$pfSenseFullPath = Join-Path $RootDirectory $PfSensePath

# Load template content
if (Test-Path -Path $templateFullPath -PathType Leaf) {
  $templateContent = Get-Content -Path $templateFullPath -Raw
}
else {
  Write-Log "Could not find '$TemplatePath' in root directory. Using built-in template." -Color Red
  $templateContent = $builtInTemplate
}

# ---- Resolve peer data source -----------------------------------------
if ($APIExport) {
  if (-not $PfSenseHost) {
    throw "-PfSenseHost is required when using -APIExport."
  }
  if (-not $PfSenseCredential) {
    throw "-PfSenseCredential is required when using -APIExport."
  }

  Write-Log "Fetching peer data from pfSense REST API at $PfSenseHost ..." -Color Cyan
  $peerLookup = Read-PfSenseAPIExport -PfSenseHostName $PfSenseHost -Credential $PfSenseCredential -SkipCertCheck:$SkipCertificateCheck
  Write-Log "Found $($peerLookup.Count) peer(s) via API." -Color Green

  # If ServerPublicKey was not provided, try to fetch it from the first tunnel
  if ([string]::IsNullOrWhiteSpace($ServerPublicKey)) {
    Write-Log "ServerPublicKey not specified; attempting to fetch from pfSense tunnel configuration ..." -Color Yellow

    $pair = "$($PfSenseCredential.UserName):$($PfSenseCredential.GetNetworkCredential().Password)"
    $encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
    $headers = @{ Authorization = "Basic $encodedCreds" }

    $tunnelParams = @{
      Uri = "https://$PfSenseHost/api/v2/vpn/wireguard/tunnels"
      Method = 'GET'
      Headers = $headers
      ContentType = 'application/json'
    }
    if ($SkipCertificateCheck) {
      if ($PSVersionTable.PSVersion.Major -ge 6) {
        $tunnelParams['SkipCertificateCheck'] = $true
      }
      else {
        $callback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
      }
    }

    try {
      $tunnelResponse = Invoke-RestMethod @tunnelParams
      $tunnels = if ($tunnelResponse.data) { $tunnelResponse.data } else { $tunnelResponse }
      if ($tunnels -and $tunnels.Count -gt 0 -and $tunnels[0].publickey) {
        $ServerPublicKey = $tunnels[0].publickey
        Write-Log "Fetched server public key from tunnel '$($tunnels[0].name)'." -Color Green
      }
      else {
        throw "No tunnels returned or no public key found."
      }
    }
    catch {
      throw "ServerPublicKey was not provided and could not be fetched from the pfSense API: $_"
    }
    finally {
      if ($SkipCertificateCheck -and $PSVersionTable.PSVersion.Major -lt 6) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $callback
      }
    }
  }
}
else {
  if (-not (Test-Path $pfSenseFullPath)) {
    throw "Required file '$PfSensePath' is missing from the root directory. Please download the tunnel export from pfSense and place it in the script's directory, or use -APIExport."
  }

  $peerLookup = Read-PfSenseFileExport -Path $pfSenseFullPath
}

# ---- Generate per-user config files ----------------------------------
$userFolders = Get-ChildItem -Directory $RootDirectory

foreach ($folder in $userFolders) {
  $userName = $folder.Name
  $privateKeyPath = Join-Path $folder.FullName "privatekey"
  $outputPath = Join-Path $folder.FullName "wg.conf"

  if ((Test-Path $privateKeyPath) -and $peerLookup.ContainsKey($userName)) {
    $privKey = (Get-Content -Path $privateKeyPath).Trim()
    $externalData = $peerLookup[$userName]

    $finalConfig = $templateContent `
      -replace "<privatekey>", $privKey `
      -replace "<pfsense_publickey>", $ServerPublicKey `
      -replace "<pre-shared-key>", $externalData.PSK `
      -replace "<allowed_ips>", $externalData.IP `
      -replace "<peer_endpoint>", $ConfigurationPeerEndpoint `
      -replace "<persistent_keepalive>", $ConfigurationPersistentKeepalive `
      -replace "<dns_servers>", ($ConfigurationInterfaceDNSServers -join ", ") `
      -replace "<peer_allowed_ips>", ($ConfigurationPeerAllowedIPs -join ", ")

    $finalConfig | Out-File -FilePath $outputPath -Encoding UTF8
    Write-Log "Generated: $userName -> $($externalData.IP)" -Color Green
  }
  else {
    Write-Log "Skipping ${userName}: Missing privatekey file or name not found in peer data" -Color Yellow
  }
}
