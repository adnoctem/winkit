#Requires -Version 2.0
#Requires -Module ActiveDirectory
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Imports the current DHCP server configuration from a specified XML file.

.DESCRIPTION
  This script connects to the DHCP server and imports its configuration from an XML file. It allows for filtering options to include specific DHCP scopes and can handle large datasets efficiently.

.PARAMETER InputPath
  The full path to the XML file from which the DHCP configuration will be imported.

.EXAMPLE
  PS> ./Import-DHCP.ps1 -InputPath 'C:\Exports\DHCPConfig.xml'
  This command imports the DHCP server configuration from the specified XML file.

.EXAMPLE
  PS> ./Import-DHCP.ps1 -InputPath 'C:\Exports\DHCPConfig.xml' -Scope '192.168.1.0'
  This command imports the DHCP configuration for the specified scope from the XML file.

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
    HelpMessage = "The input file path for the XML file."
  )]
  [string]$InputPath,

  [Parameter(
    Position = 1,
    Mandatory = $false,
    HelpMessage = "The hostname of the DHCP server."
  )]
  [string]$ComputerName
)

if (-not (Test-Path -Path $InputPath -PathType Leaf)) {
  throw("InputPath is not a valid file. Please provide a valid XML file path.")
}

Import-DhcpServer -ComputerName $ComputerName -File $InputPath -Force
