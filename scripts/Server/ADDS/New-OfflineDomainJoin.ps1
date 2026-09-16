#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.1.0' }

<#
.SYNOPSIS
  Provisions an AD computer account and writes an offline-join blob file.

.DESCRIPTION
  Wrapper around PSFoundation's New-OfflineDomainJoinBlob and New-DjoinFile:
  provisions the computer account in Active Directory and generates the
  join blob via the native NetCreateProvisioningPackage Win32 API - no
  djoin.exe dependency - then writes it in the format Windows Setup and
  djoin.exe /loadfile expect.

  Use this for provisioning/imaging pipelines: the blob file is consumed on
  the target machine during unattended setup or first boot, before it has
  live DC connectivity.

  The account running this script (or -Credential) needs rights to create
  computer accounts in the target domain/OU.

.PARAMETER ComputerName
  Computer account name to provision.

.PARAMETER DomainName
  Domain (DNS or NetBIOS) to join.

.PARAMETER Credential
  Optional credential with rights to create computer accounts. When
  omitted, the current context is used.

.PARAMETER MachineAccountOU
  Optional OU for the computer account (distinguished name).

.PARAMETER DestinationFile
  Full path of the blob file to write. Defaults to C:\temp\djoin.tmp.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./New-OfflineDomainJoin.ps1 -ComputerName SRV001 -DomainName company.com -Credential $domainAdmin -MachineAccountOU 'OU=Servers,DC=company,DC=com' -DestinationFile C:\join\SRV001.djoin
  Provisions SRV001 and writes the join blob to C:\join\SRV001.djoin.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [string]
  $ComputerName,

  [Parameter(Mandatory = $true)]
  [string]
  $DomainName,

  [Parameter(Mandatory = $false)]
  [pscredential]
  $Credential,

  [Parameter(Mandatory = $false)]
  [string]
  $MachineAccountOU,

  [Parameter(Mandatory = $false)]
  [string]
  $DestinationFile = 'C:\temp\djoin.tmp',

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

$_results = New-Object System.Collections.ArrayList

Write-Log -Message "Provisioning computer account $ComputerName in $DomainName..." -Color Cyan

try {
  $blobParams = @{
    ComputerName = $ComputerName
    Domain = $DomainName
  }
  if ($Credential) {
    $blobParams.Credential = $Credential
  }
  if ($MachineAccountOU) {
    $blobParams.MachineAccountOU = $MachineAccountOU
  }

  $blobResult = New-OfflineDomainJoinBlob @blobParams -Confirm:$false -ErrorAction Stop
  if (-not $blobResult -or -not $blobResult.Blob) {
    throw 'New-OfflineDomainJoinBlob returned no blob.'
  }

  $fileResult = New-DjoinFile -Blob $blobResult.Blob -DestinationFile $DestinationFile -Confirm:$false -ErrorAction Stop
  if (-not $fileResult -or -not (Test-Path -LiteralPath $DestinationFile -PathType Leaf)) {
    throw "Blob file was not written to $DestinationFile."
  }

  Add-OperationResult -Results $_results -Target $ComputerName -Source 'OfflineDomainJoin' -Action 'Provision' -Status 'Completed' -Detail "Blob written to $DestinationFile (join during unattended setup or first boot)."
  Write-Log -Message "Blob written to $DestinationFile" -Color Green
}
catch {
  Add-OperationResult -Results $_results -Target $ComputerName -Source 'OfflineDomainJoin' -Action 'Provision' -Status 'Failed' -Detail $_.Exception.Message
  Write-Log -Message "Offline domain join provisioning failed: $($_.Exception.Message)" -Color Red
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'New-OfflineDomainJoin'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_results
}

if ($_results | Where-Object { $_.Status -eq 'Failed' }) {
  exit 1
}
exit 0
