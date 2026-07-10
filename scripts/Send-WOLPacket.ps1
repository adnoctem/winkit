Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Sends a Wake-on-LAN magic packet to a target machine on the local subnet.
.DESCRIPTION
  Builds a WoL magic packet (6 x 0xFF followed by the target MAC repeated 16 times)
  and broadcasts it via UDP. Use -DryRun to preview without sending.
.PARAMETER MacAddress
  Target MAC address. Accepted formats:
    '1A:2B:3C:4D:5E:6F'  (colon-separated)
    '1A-2B-3C-4D-5E-6F'  (dash-separated)
    '1A2B3C4D5E6F'        (no separator)
  Casing is ignored.
.PARAMETER BroadcastAddress
  UDP broadcast destination. Defaults to the subnet broadcast of the active
  adapter. Override for cross-subnet or directed broadcast scenarios.
.PARAMETER Port
  UDP port to broadcast on. Defaults to 9. Port 7 is the legacy alternative;
  both are valid at the NIC level.
.PARAMETER DryRun
  Preview the operation without sending any packets.
.EXAMPLE
  PS> Send-WOLPacket -MacAddress '1A:2B:3C:4D:5E:6F'
.EXAMPLE
  PS> Send-WOLPacket -MacAddress '1A:2B:3C:4D:5E:6F' -BroadcastAddress '192.168.1.255'
.EXAMPLE
  PS> Send-WOLPacket -MacAddress '1A:2B:3C:4D:5E:6F' -DryRun
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

param (
  [Parameter(Position = 0, Mandatory = $true)]
  [ValidatePattern('^([0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}$|^[0-9A-Fa-f]{12}$')]
  [string]$MacAddress,

  # No default value here - resolved in the script body after module import
  [Parameter(Position = 1, Mandatory = $false)]
  [string]$BroadcastAddress,

  [Parameter(Position = 2, Mandatory = $false)]
  [ValidateRange(1, 65535)]
  [int]$Port = 9,

  [Parameter(Mandatory = $false)]
  [switch]$DryRun
)

# -----------------------------------------------------------------------------

$Host.UI.RawUI.WindowTitle = 'winkit - Send-WOLPacket'

if ($DryRun) {
  Write-Log -Message "DRY RUN - no packets will be sent`n" -Color Yellow
}

# Resolve broadcast address after module import - cannot be done in param default
if ([string]::IsNullOrEmpty($BroadcastAddress)) {
  try {
    $adapter = Get-DefaultNetworkAdapter -Required
    $BroadcastAddress = Get-BroadcastAddress -Adapter $adapter -Required
    Write-Log -Message "  -> Using detected broadcast address: $BroadcastAddress" -Color Gray
  }
  catch {
    Write-Log -Message "Could not determine broadcast address: $_" -Color Red
    exit 1
  }
}

# Build the magic packet
Write-Log -Message "Building magic packet for MAC $MacAddress ..." -Color Yellow

$normalized = $MacAddress -replace '[^0-9A-Fa-f]', ''
if ($normalized.Length -ne 12) {
  Write-Log -Message "Invalid MAC address after normalization - expected 12 hex digits, got $($normalized.Length)." -Color Red
  exit 1
}

$macBytes = $normalized -split '(..)' |
  Where-Object { $_ -ne '' } |
  ForEach-Object { [byte]("0x$_") }

[byte[]]$packet = (@([byte]0xFF) * 6) + ($macBytes * 16)
Write-Log -Message '  -> Magic packet assembled.' -Color Gray

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would broadcast to $BroadcastAddress on UDP port $Port" -Color Yellow
  Write-Log -Message "`nDRY RUN COMPLETE - no packets were sent." -Color Yellow
  return
}

# Send
Write-Log -Message "Broadcasting to $BroadcastAddress on UDP port $Port ..." -Color Yellow
try {
  $client = New-Object System.Net.Sockets.UdpClient
  $client.EnableBroadcast = $true
  $endpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse($BroadcastAddress), $Port)
  $client.Send($packet, $packet.Length, $endpoint) | Out-Null
  $client.Close()
  Write-Log -Message "Magic packet sent - target should wake shortly." -Color Green
}
catch {
  Write-Log -Message "Failed to send magic packet: $_" -Color Red
  exit 1
}
