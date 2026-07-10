Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Generates hostnames according to the OFC hostname schema.

.DESCRIPTION
  Produces one or more hostname strings following the pattern
  OFC-{BuildingId}-{MachineType}-{SubType}-{Serial}. The schema is
  designed for campus-wide device inventory where each hostname encodes
  the company code, building, machine type, subtype, and a zero-padded
  serial number.

  Machine types and their expected subtypes:
  - NI (Network Infrastructure): FW, SW, AP, MC, BX, MD, VS, SS; 3-digit serial
  - VM (Virtual Machine): DC, MX, CJ, AS, NS, DH, CA; 3-digit serial
  - CT (Container): DC, MX, CJ, AS, NS, DH, CA; 3-digit serial
  - PC (Personal Computer), NB (Notebook), MP (Mobile Phone):
    W (Windows), L (Linux), I (macOS), A (Android), B (BSD), U (Unix); 4-digit serial

.PARAMETER CompanyCode
  Company code prefix. Defaults to OFC.

.PARAMETER BuildingId
  Building identifier in the format of a letter (H, B, W, S) followed by
  two digits, e.g. H00, B01, W03.

.PARAMETER MachineType
  Machine type category: NI (Network Infrastructure), VM (Virtual Machine),
  CT (Container), PC (Personal Computer), NB (Notebook), MP (Mobile Phone).

.PARAMETER InfrastructureType
  Required when MachineType is NI. One of: FW, SW, AP, MC, BX, MD, VS, SS.

.PARAMETER WorkloadPurpose
  Required when MachineType is VM or CT. One of: DC, MX, CJ, AS, NS, DH, CA.

.PARAMETER OS
  Required when MachineType is PC, NB, or MP. One of: W, L, I, A, B, U.

.PARAMETER StartIndex
  Starting serial number. Defaults to 1.

.PARAMETER Count
  Number of hostnames to generate. Defaults to 1.

.PARAMETER DryRun
  Preview which hostnames would be generated without outputting them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./New-HostnameScheme.ps1 -BuildingId H00 -MachineType NI -InfrastructureType FW

.EXAMPLE
  PS> ./New-HostnameScheme.ps1 -BuildingId H00 -MachineType PC -OS W -StartIndex 1 -Count 5

.EXAMPLE
  PS> ./New-HostnameScheme.ps1 -BuildingId H00 -MachineType VM -WorkloadPurpose DC -StartIndex 10 -Count 3

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string]
  $CompanyCode = 'OFC',

  [Parameter(Mandatory = $true)]
  [ValidatePattern('^[HBWS]\d{2}$')]
  [string]
  $BuildingId,

  [Parameter(Mandatory = $true)]
  [ValidateSet('NI', 'VM', 'CT', 'PC', 'NB', 'MP')]
  [string]
  $MachineType,

  [Parameter(Mandatory = $false)]
  [ValidateSet('FW', 'SW', 'AP', 'MC', 'BX', 'MD', 'VS', 'SS')]
  [string]
  $InfrastructureType,

  [Parameter(Mandatory = $false)]
  [ValidateSet('DC', 'MX', 'CJ', 'AS', 'NS', 'DH', 'CA')]
  [string]
  $WorkloadPurpose,

  [Parameter(Mandatory = $false)]
  [ValidateSet('W', 'L', 'I', 'A', 'B', 'U')]
  [string]
  $OS,

  [Parameter(Mandatory = $false)]
  [int]
  $StartIndex = 1,

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 10000)]
  [int]
  $Count = 1,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no hostnames will be output`n" -Color Yellow
}

# ---- Validate subtype based on machine type ----------------------------------
switch ($MachineType) {
  'NI' {
    if (-not $InfrastructureType) {
      Write-Log -Message "MachineType 'NI' requires -InfrastructureType (FW, SW, AP, MC, BX, MD, VS, SS)." -Color Red
      exit 1
    }
  }
  'VM' {
    if (-not $WorkloadPurpose) {
      Write-Log -Message "MachineType 'VM' requires -WorkloadPurpose (DC, MX, CJ, AS, NS, DH, CA)." -Color Red
      exit 1
    }
  }
  'CT' {
    if (-not $WorkloadPurpose) {
      Write-Log -Message "MachineType 'CT' requires -WorkloadPurpose (DC, MX, CJ, AS, NS, DH, CA)." -Color Red
      exit 1
    }
  }
  { $_ -in @('PC', 'NB', 'MP') } {
    if (-not $OS) {
      Write-Log -Message "MachineType '$MachineType' requires -OS (W, L, I, A, B, U)." -Color Red
      exit 1
    }
  }
}

# ---- Determine serial number padding -----------------------------------------
$padWidth = if ($MachineType -in @('NI', 'VM', 'CT')) { 3 } else { 4 }
$_results = New-Object System.Collections.ArrayList
$_hostnames = New-Object System.Collections.ArrayList

for ($i = 0; $i -lt $Count; $i++) {
  $currentIndex = $StartIndex + $i
  $serial = $currentIndex.ToString("D$padWidth")

  $segments = @($CompanyCode, $BuildingId, $MachineType)

  switch ($MachineType) {
    'NI' { $segments += $InfrastructureType }
    'VM' { $segments += $WorkloadPurpose }
    'CT' { $segments += $WorkloadPurpose }
    'PC' { $segments += $OS }
    'NB' { $segments += $OS }
    'MP' { $segments += $OS }
  }

  $segments += $serial
  $hostname = $segments -join '-'

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would generate: $hostname" -Color Yellow
  }
  else {
    Write-Output $hostname
  }

  [void]$_hostnames.Add($hostname)
  Add-OperationResult -Results $_results -Target $hostname -Source 'HostnameScheme' -Action 'Generate' -Status 'Completed' -Detail "Type=$MachineType, Building=$BuildingId, Serial=$serial"
}

# ---- Summary ----------------------------------------------------------------
if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - $Count hostname(s) would have been generated." -Color Yellow
}
elseif ($_hostnames.Count -gt 0) {
  Write-Log -Message "`nGenerated $($_hostnames.Count) hostname(s) in $CompanyCode-$BuildingId-$MachineType namespace." -Color $(
    if ($_hostnames.Count -gt 0) { 'Cyan' } else { 'Gray' })
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'New-HostnameScheme'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
