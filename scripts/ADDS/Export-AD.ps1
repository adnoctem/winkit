#Requires -Version 2.0
#Requires -Module ActiveDirectory
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Exports Active Directory users to a specified CSV file.

.DESCRIPTION
  This script connects to the Active Directory and retrieves user information, exporting it to a CSV file. It allows for filtering options to include specific user attributes and can handle large datasets efficiently.

.PARAMETER OutputFile
  The full path to the CSV file where the exported user data will be saved.

.PARAMETER Filter
  Optional parameter to filter users based on specific criteria (e.g., department, status).

.EXAMPLE
  PS> ./Export-AD.ps1 -OutputPath 'C:\Exports'
  This command exports all Active Directory users to the specified CSV file.

.EXAMPLE
  PS> ./Export-AD.ps1 -OutputPath 'C:\Exports' -Resource 'Users'
  This command exports Active Directory users from the Sales department to the specified CSV file.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

# Set-Variable -Name resources -Value @('Users', 'Computers', 'Groups') -Option ReadOnly
# Set-Variable -Name resourcesAll -Value 'All' -Option ReadOnly

# Parameters
param (
  [Parameter(
    Position = 0,
    Mandatory = $true,
    HelpMessage = "The output file path for the exported CSV."
  )]
  [string]$OutputPath,

  [Parameter(
    Position = 1,
    Mandatory = $false,
    HelpMessage = "The resource to export from Active Directory."
  )]
  [ValidateSet('Users', 'Computers', 'Groups', 'All')]
  [string]$Resource = 'All',

  [Parameter(
    Position = 2,
    Mandatory = $false,
    HelpMessage = "The encoding for the output CSV file."
  )]
  [ValidateSet("UTF8", "ASCII", "Unicode", "UTF7", "UTF32", "BigEndianUnicode")]
  [string]$Encoding = "UTF8"
)

# ---- Module import ------------------------------------
# $root = Split-Path $PSScriptRoot -Parent
# $module = Join-Path -Path $root 'lib/winkit.psm1'

Import-Module ActiveDirectory -Force
# -------------------------------------------------------

$properties = [PSCustomObject]@{
  Users = @(
    # Windows AD specific
    "sAMAccountName",
    "userPrincipalName",
    "lastLogon",
    "lastLogonTimestamp",
    "whenCreated",
    "whenChanged",
    "accountExpires",

    # Paths
    "homeDirectory",
    "homeDrive",
    "profilePath",

    # General
    "cn",
    "sn",
    "name",
    "givenName",
    "displayName",
    "mail",
    "mailNickname",
    "description",
    "telephoneNumber"
  )

  Computers = @(
    # Windows AD specific
    "cn",
    "name",
    "sAMAccountName",
    "servicePrincipalName",
    "dNSHostName",
    "lastLogon",
    "lastLogonTimestamp",
    "whenCreated",
    "whenChanged",
    "accountExpires",

    # General
    "operatingSystem",
    "operatingSystemVersion"
  )

  Groups = @(
    # Windows AD specific
    "cn",
    "name",
    "sAMAccountName",
    "description",
    "whenCreated",
    "whenChanged",

    # General
    "groupType"
  )
}

$fmt = (Get-Culture).TextInfo

# Initialize an empty object to hold the results
# Distinguish between PowerShell 3 and later versions since [PSCustomObject] is not available in PS2
if ($PSVersionTable.PSVersion.Major -ge 3) {
  $obj = [PSCustomObject]@{
    Users = $null
    Computers = $null
    Groups = $null
  }
}
else {
  $obj = New-Object PSObject -Property @{
    Users = $null
    Computers = $null
    Groups = $null
  }
}

switch ($Resource) {
  'All' {
    $obj.Users = Get-ADUser -Filter * -Properties $properties.Users
    $obj.Computers = Get-ADComputer -Filter * -Properties $properties.Computers
    $obj.Groups = Get-ADGroup -Filter * -Properties $properties.Groups
  }

  'Users' {
    $obj.Users = Get-ADUser -Filter * -Properties $properties.Users
  }

  'Computers' {
    $obj.Computers = Get-ADComputer -Filter * -Properties $properties.Computers
  }

  'Groups' {
    $obj.Groups = Get-ADGroup -Filter * -Properties $properties.Groups
  }

  default {
    throw("Invalid resource specified. Use one of: ('Users', 'Computers', 'Groups', 'All').")
  }
}

# ensure OutputPath is a directory
#
# ref: https://stackoverflow.com/questions/39825440/check-if-a-path-is-a-folder-or-a-file-in-powershell
if (Test-Path -Path $OutputPath -PathType Leaf) {
  throw("OutputPath is a file. Please provide a directory path.")
}

$obj.PSObject.Properties | ForEach-Object {
  $name = $_.Name
  $data = $_.Value

  if ($data -ne $null) {
    $output = Join-Path -Path $OutputPath -ChildPath "AD-$($fmt.ToTitleCase($name)).csv"
    $data | Export-Csv -Path $output -NoTypeInformation -Encoding $Encoding -Delimiter ';'
  }
  else {
    Write-Error "No data found for resource 'AD-$($fmt.ToTitleCase($name)).csv'. Skipping export for this resource."
  }
}
