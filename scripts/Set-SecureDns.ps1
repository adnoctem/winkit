#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Configures DNS servers and DNS-over-HTTPS on the default network adapter.

.DESCRIPTION
  Sets IPv4/IPv6 DNS server addresses and, on Windows 11 22H2+, enables
  DNS-over-HTTPS for each server. Supports named presets covering major
  public resolvers with known DoH templates.

  The adapter is resolved automatically via the default IPv4 route.
  Use -AdapterName to target a specific adapter by name.

.PARAMETER Preset
  Named DNS provider preset. Defaults to Google.
  Supported: Google, Cloudflare, Quad9, AdGuard, Mullvad, OpenDNS.

.PARAMETER AdapterName
  Target a specific network adapter by name. When omitted the adapter
  with the lowest-metric 0.0.0.0/0 route is used.

.PARAMETER Undo
  Reset the adapter back to DHCP-assigned DNS servers.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Set-SecureDns.ps1
  Configures Google public DNS and DoH on the default adapter.

.EXAMPLE
  PS> ./Set-SecureDns.ps1 -Preset Quad9
  Uses Quad9 (9.9.9.9 / 149.112.112.112) with DoH.

.EXAMPLE
  PS> ./Set-SecureDns.ps1 -AdapterName 'Ethernet' -Preset Cloudflare
  Targets the adapter named 'Ethernet' specifically.

.EXAMPLE
  PS> ./Set-SecureDns.ps1 -Undo
  Resets the adapter back to DHCP DNS.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Named DNS provider preset.'
  )]
  [ValidateSet('Google', 'Cloudflare', 'Quad9', 'AdGuard', 'Mullvad', 'OpenDNS')]
  [string]
  $Preset = 'Google',

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Target a specific adapter by name instead of the default route.'
  )]
  [string]
  $AdapterName,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Reset DNS back to DHCP.'
  )]
  [switch]
  $Undo,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Preview changes without applying them.'
  )]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

$_dnsPresets = @{
  'Google' = @{
    IPv4 = @('8.8.8.8', '8.8.4.4')
    IPv6 = @('2001:4860:4860::8888', '2001:4860:4860::8844')
    DoHTemplate = 'https://dns.google/dns-query'
  }
  'Cloudflare' = @{
    IPv4 = @('1.1.1.1', '1.0.0.1')
    IPv6 = @('2606:4700:4700::1111', '2606:4700:4700::1001')
    DoHTemplate = 'https://cloudflare-dns.com/dns-query'
  }
  'Quad9' = @{
    IPv4 = @('9.9.9.9', '149.112.112.112')
    IPv6 = @('2620:fe::fe', '2620:fe::9')
    DoHTemplate = 'https://dns.quad9.net/dns-query'
  }
  'AdGuard' = @{
    IPv4 = @('94.140.14.14', '94.140.15.15')
    IPv6 = @('2a10:50c0::ad1:ff', '2a10:50c0::ad2:ff')
    DoHTemplate = 'https://dns.adguard-dns.com/dns-query'
  }
  'Mullvad' = @{
    IPv4 = @('194.242.2.2', '194.242.2.3')
    IPv6 = @('2a07:e340::2', '2a07:e340::3')
    DoHTemplate = 'https://dns.mullvad.net/dns-query'
  }
  'OpenDNS' = @{
    IPv4 = @('208.67.222.222', '208.67.220.220')
    IPv6 = @('2620:119:35::35', '2620:119:53::53')
    DoHTemplate = 'https://doh.opendns.com/dns-query'
  }
}

$_presetData = $_dnsPresets[$Preset]

# ---- Undo --------------------------------------------------------------------
if ($Undo) {
  if ($AdapterName) {
    $_adapter = Get-NetAdapter -Name $AdapterName -ErrorAction SilentlyContinue
    if (-not $_adapter) {
      Write-Log -Message "Adapter '$AdapterName' not found." -Color Red
      Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Reset' -Status 'Failed' -Detail 'Adapter not found.'
      if ($PassThru -or $DryRun) { $_results }
      exit 1
    }
    $ifIndex = $_adapter.ifIndex
  }
  else {
    $_defaultAdapter = Get-DefaultNetworkAdapter -Required:$true
    $ifIndex = $_defaultAdapter.ifIndex
    $AdapterName = $_defaultAdapter.Name
  }

  Write-Log -Message "Resetting DNS on adapter '$AdapterName' to DHCP." -Color Yellow

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would reset DNS on interface $ifIndex to DHCP." -Color Yellow
    Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Reset' -Status 'Skipped' -Detail 'DryRun'
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }

  try {
    Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ResetServerAddresses -ErrorAction Stop
    Write-Log -Message '  -> DNS reset to DHCP.' -Color Green
    Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Reset' -Status 'Completed' -Detail 'DNS reset to DHCP.'
  }
  catch {
    Write-Log -Message "  -> FAILED: $_" -Color Red
    Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Reset' -Status 'Failed' -Detail $_.Exception.Message
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }

  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Set-SecureDns'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

# ---- Resolve adapter ---------------------------------------------------------
if ($AdapterName) {
  $_adapter = Get-NetAdapter -Name $AdapterName -ErrorAction SilentlyContinue
  if (-not $_adapter) {
    Write-Log -Message "Adapter '$AdapterName' not found." -Color Red
    Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Set' -Status 'Failed' -Detail 'Adapter not found.'
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }
  $ifIndex = $_adapter.ifIndex
}
else {
  $_defaultAdapter = Get-DefaultNetworkAdapter -Required:$true
  $ifIndex = $_defaultAdapter.ifIndex
  $AdapterName = $_defaultAdapter.Name
}

Write-Log -Message "Configuring DNS on adapter '$AdapterName' (ifIndex: $ifIndex)." -Color Yellow
Write-Log -Message "  Preset: $Preset" -Color Gray

# ---- Apply DNS servers -------------------------------------------------------
$_allServers = @($_presetData.IPv4) + @($_presetData.IPv6)

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would set DNS servers: $($_allServers -join ', ')" -Color Yellow
  Write-Log -Message "[DRY RUN] Would configure DoH: $($_presetData.DoHTemplate)" -Color Yellow
  Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Set' -Status 'Skipped' -Detail "DryRun: $Preset ($($_allServers -join ', '))"
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if (-not $PSCmdlet.ShouldProcess($AdapterName, "Set DNS to $Preset ($($_allServers -join ', '))")) {
  Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Set' -Status 'Skipped' -Detail 'WhatIf'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

try {
  Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ServerAddresses $_allServers -ErrorAction Stop
  Write-Log -Message "  -> DNS servers set: $($_allServers -join ', ')" -Color Green
  Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Set' -Status 'Completed' -Detail "Servers: $($_allServers -join ', ')"
}
catch {
  Write-Log -Message "  -> FAILED: $_" -Color Red
  Add-OperationResult -Results $_results -Target $AdapterName -Source 'DNS' -Action 'Set' -Status 'Failed' -Detail $_.Exception.Message
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

# ---- Configure DoH -----------------------------------------------------------
$hasDohCmdlet = Get-Command -Name 'Set-DnsClientDohServerAddress' -ErrorAction SilentlyContinue

if ($hasDohCmdlet) {
  Write-Log -Message 'Configuring DNS-over-HTTPS.' -Color Yellow
  try {
    foreach ($_server in $_allServers) {
      try {
        Set-DnsClientDohServerAddress -ServerAddress $_server -DohTemplate $_presetData.DoHTemplate -AllowFallbackToUdp $true -AutoUpgrade $true -ErrorAction Stop
        Write-Log -Message "  -> DoH configured for $_server" -Color Green
      }
      catch {
        Write-Log -Message "  -> DoH for $_server: $_" -Color Yellow
      }
    }
    Write-Log -Message '  -> DNS-over-HTTPS configured.' -Color Green
  }
  catch {
    Write-Log -Message "  -> DoH configuration partially failed: $_" -Color Yellow
  }
}
else {
  Write-Log -Message 'DNS-over-HTTPS cmdlet not available on this Windows build. DNS servers are set, but DoH was not configured.' -Color Yellow
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Set-SecureDns'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
