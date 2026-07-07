#Requires -Version 2.0
#Requires -Module ActiveDirectory
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Exports the current DHCP server configuration to a specified XML file.

.DESCRIPTION
  This script connects to the DHCP server and retrieves its configuration, exporting it to an XML file. It allows for filtering options to include specific DHCP scopes and can handle large datasets efficiently.

.PARAMETER OutputPath
  The full path to the XML file where the exported DHCP configuration will be saved.

.EXAMPLE
  PS> ./Export-DHCP.ps1 -OutputPath 'C:\Exports\DHCPConfig.xml'
  This command exports the current DHCP server configuration to the specified XML file.

.EXAMPLE
  PS> ./Export-DHCP.ps1 -OutputPath 'C:\Exports\DHCPConfig.xml' -Scope '192.168.1.0'
  This command exports the DHCP configuration for the specified scope to the XML file.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

# Parameters
param (
  [Parameter(
    Position = 0,
    Mandatory = $true,
    HelpMessage = "The output file path for the exported XML."
  )]
  [string]$OutputPath,

  [Parameter(
    Position = 1,
    Mandatory = $false,
    HelpMessage = "The DHCP scope to export."
  )]
  [string]$Scope = '192.168.1.0',

  [Parameter(
    Position = 2,
    Mandatory = $false,
    HelpMessage = "The hostname of the DHCP server."
  )]
  [string]$ComputerName
)

if (Test-Path -Path $OutputPath -PathType Leaf) {
  throw("OutputPath is a file. Please provide a directory path.")
}

Export-DhcpServer -ComputerName $ComputerName -ScopeId $Scope -File $OutputPath -Force
