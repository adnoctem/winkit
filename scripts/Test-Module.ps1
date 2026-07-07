<#
.SYNOPSIS
  TBA..
.DESCRIPTION
  Uses Outlook's COM interface to find items by ReceivedTime and relocate them
  to the specified archive folder. Designed for batch clean-up and retention
  workflows where a fixed date range needs to be archived.
.EXAMPLE
  PS> ./archive-outlook.ps1 -StartDate '2024-01-01' -EndDate '2024-03-31' -ArchiveFolder 'Archive/2024 Q1'
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

# ---- Module import ------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -------------------------------------------------------

# Write-Log -Message "This is a test!" -Color Cyan -Timestamps

# Write-Log "This is some bs" -Color Cyan -Timestamps
# Write-Log "`$root would have been: $root"
# Write-Log "`$module would have been: $module"
# Write-Log "Now we're using $PSScriptRoot\..\lib\winkit.psd1"

# ConvertFrom-HTMLtoWord -FileHTML "C:\Users\Admin\tmp\Pressemitteilung\test.html" -OutputFile "C:\Users\Admin\tmp\Pressemitteilung\test.docx" -Show | Out-Null
# Convert-HTMLToPDF -FilePath '/tmp/test.html' -OutputFilePath '/tmp/test-out.pdf'


# ---------- Networking function tests --------------------

# Adapter resolution
# Get-DefaultNetworkAdapter
# Get-DefaultNetworkAdapter -Type Ethernet
# Get-DefaultNetworkAdapter -Type WiFi -Required

# Address retrieval
Get-IPAddress
Get-IPAddress -AddressFamily IPv6
Get-IPAddress -AddressFamily IPv4 -Required

# Subnet / prefix
Get-SubnetMask
Get-SubnetMask -Required

# Default gateway
Get-DefaultGateway
Get-DefaultGateway -AddressFamily IPv6

# DNS servers
Get-DnsServer
Get-DnsServer -AddressFamily IPv6

# MAC
Get-MACAddress

# Network / prefix addresses
Get-NetworkPrefix
Get-NetworkPrefix -AddressFamily IPv6
Get-Network                # alias
Get-Prefix -AddressFamily IPv6
Get-NetworkPrefixCIDR
Get-NetworkPrefixCIDR -AddressFamily IPv6
Get-NetworkCIDR            # alias
Get-PrefixCIDR -AddressFamily IPv6

# Broadcast / multicast
# Get-BroadcastAddress
# Get-MulticastAddress

# Validation
# Test-IPv4Address -Address '192.168.1.1'
# Test-IPv4Address -Address '999.0.0.1'
# Test-IPv6Address -Address '2001:db8::ff00:42:8329'
# Test-IPv6Address -Address '::1'
# Test-IPv6Address -Address '::ffff:192.168.1.1'

# Piped adapter reuse (avoids re-resolving)
# $adapter = Get-DefaultNetworkAdapter -Type Ethernet
# Get-IPAddress -Adapter $adapter
# Get-IPAddress -AddressFamily IPv6 -Adapter $adapter
# Get-SubnetMask -Adapter $adapter
# Get-DefaultGateway -Adapter $adapter
# Get-DefaultGateway -AddressFamily IPv6 -Adapter $adapter
# Get-DNSServer -Adapter $adapter
# Get-MACAddress -Adapter $adapter
# Get-NetworkPrefix -Adapter $adapter
# Get-NetworkPrefix -AddressFamily IPv6 -Adapter $adapter
# Get-NetworkPrefixCIDR -Adapter $adapter
# Get-NetworkPrefixCIDR -AddressFamily IPv6 -Adapter $adapter
# Get-BroadcastAddress -Adapter $adapter
# Get-MulticastAddress -Adapter $adapter

Get-OSBuildNumber
Get-OSEdition
Get-OSVersionInfo
Get-SystemMemory
