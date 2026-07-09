#Requires -Version 5.0

<#
.SYNOPSIS
  Move Outlook messages received between two dates into an archive folder.

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

# Parameters
param (
  # Start
  [Parameter(
    Position = 0,
    Mandatory = $true,
    HelpMessage = "The starting date from which to begin email archival."
  )]
  [ValidatePattern("^[0-9]{4}-[0-9]{2}-[0-9]{2}$")]
  [string]$Start,

  # End
  [Parameter(
    Position = 1,
    Mandatory = $true,
    HelpMessage = "The ending date to which we're archiving emails."
  )]
  [ValidatePattern("^[0-9]{4}-[0-9]{2}-[0-9]{2}$")]
  [string]$End
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

# ---- Assembly import ------------------------------------
$_searchPath = 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook'
$_searchFilter = 'Microsoft.Office.Interop.Outlook.dll'
$_assembly = Get-ChildItem -LiteralPath $_searchPath `
  -Filter $_searchFilter `
  -Recurse |
  Select-Object -ExpandProperty FullName -Last 1

Add-Type -AssemblyName $_assembly -ErrorAction Stop
# -------------------------------------------------------


$destination = Get-NewPath -Path 'OutlookArchive'
Write-Output "Archiving Outlook mail received from $Start through $End."

$_outlook = New-Object -com outlook.application

$ns = $_outlook.GetNamespace("MAPI")
$store = $ns.Stores | Select-Object -First 1

$folders = $store.GetRootFolder().Folders

foreach ($folder in $folders) {
  Write-Output "Copying $(folder.Name)"
}

Write-Output $destination
