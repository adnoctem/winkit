Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Applies TCP latency-reduction registry tweaks per network adapter.

.DESCRIPTION
  Disables Nagle's algorithm by writing TcpAckFrequency=1 and
  TCPNoDelay=1 under each adapter's TCP/IP interface registry key.

  This is a latency-vs-throughput tradeoff popular with gaming and
  real-time workloads. Bulk-transfer throughput may decrease. The
  script is opt-in and not included in any default optimizer profile.

.PARAMETER AdapterName
  Target a specific adapter by name. When omitted all adapters with
  TCP/IP interface registry keys are updated.

.PARAMETER Undo
  Remove TcpAckFrequency and TCPNoDelay values from adapters.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Optimize-TCPParameters.ps1
  Applies TCP tweaks to all adapters.

.EXAMPLE
  PS> ./Optimize-TCPParameters.ps1 -AdapterName 'Ethernet'
  Targets a single adapter.

.EXAMPLE
  PS> ./Optimize-TCPParameters.ps1 -Undo
  Removes the tweaks from all adapters.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Target a specific adapter by name.'
  )]
  [string]
  $AdapterName,

  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Remove TCP tweak values from adapters.'
  )]
  [switch]
  $Undo,

  [Parameter(
    Position = 1,
    Mandatory = $false,
    HelpMessage = 'Preview changes without applying them.'
  )]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

$_tcpipInterfacesKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces'

if (-not (Test-Path -LiteralPath $_tcpipInterfacesKey)) {
  Write-Log -Message "TCP/IP interfaces key not found: $_tcpipInterfacesKey" -Color Red
  Add-OperationResult -Results $_results -Target $_tcpipInterfacesKey -Source 'TCPOptimization' -Action 'Read' -Status 'Failed' -Detail 'Registry key not found.'
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

$_subkeys = Get-ChildItem -LiteralPath $_tcpipInterfacesKey -ErrorAction Stop

if ($AdapterName) {
  $_matchedKey = $null
  foreach ($_sub in $_subkeys) {
    $_netAdapter = Get-NetAdapter -Name $AdapterName -ErrorAction SilentlyContinue
    if ($_netAdapter) {
      $_ifGuid = $_netAdapter.InterfaceGuid.ToString('B')
      if ($_sub.PSChildName -eq $_ifGuid) {
        $_matchedKey = $_sub
        break
      }
    }
  }

  if (-not $_matchedKey) {
    Write-Log -Message "No TCP/IP interface found for adapter '$AdapterName'." -Color Red
    Add-OperationResult -Results $_results -Target $AdapterName -Source 'TCPOptimization' -Action 'Read' -Status 'Failed' -Detail 'No matching interface key found.'
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }

  $_interfaces = @($_matchedKey)
}
else {
  $_interfaces = $_subkeys
}

$_targetLabel = if ($Undo) { 'Removing' } else { 'Applying' }
$_anyChanges = $false

foreach ($_if in $_interfaces) {
  $_ifPath = $_if.PSPath
  $_ifGuid = $_if.PSChildName

  $_adapterName = $AdapterName
  if (-not $_adapterName) {
    try {
      $_netAdapter = Get-NetAdapter -InterfaceDescription '*' -ErrorAction SilentlyContinue |
        Where-Object { $_.InterfaceGuid.ToString('B') -eq $_ifGuid } |
        Select-Object -First 1
      $_adapterName = if ($_netAdapter) { $_netAdapter.Name } else { $_ifGuid }
    }
    catch {
      $_adapterName = $_ifGuid
    }
  }

  Write-Log -Message "$_targetLabel TCP tweaks on adapter '$_adapterName'" -Color Yellow

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would $($_targetLabel.ToLower()) TcpAckFrequency and TCPNoDelay on $_ifGuid" -Color Yellow
    Add-OperationResult -Results $_results -Target $_adapterName -Source 'TCPOptimization' -Action 'Set' -Status 'Skipped' -Detail 'DryRun'
    continue
  }

  if ($Undo) {
    $removedAny = $false
    foreach ($_valueName in @('TcpAckFrequency', 'TCPNoDelay')) {
      if (Test-Path -LiteralPath "$_ifPath\$_valueName") {
        try {
          Remove-ItemProperty -LiteralPath $_ifPath -Name $_valueName -ErrorAction Stop
          Write-Log -Message "  -> Removed $_valueName" -Color Green
          $removedAny = $true
          $_anyChanges = $true
        }
        catch {
          Write-Log -Message "  -> FAILED to remove $_valueName: $_" -Color Red
        }
      }
    }
    if ($removedAny) {
      Add-OperationResult -Results $_results -Target $_adapterName -Source 'TCPOptimization' -Action 'Remove' -Status 'Completed' -Detail 'TcpAckFrequency and TCPNoDelay removed.'
    }
    else {
      Write-Log -Message "  -> Values not present - nothing to remove." -Color Gray
      Add-OperationResult -Results $_results -Target $_adapterName -Source 'TCPOptimization' -Action 'Remove' -Status 'Skipped' -Detail 'Values not present.'
    }
  }
  else {
    try {
      Set-ItemProperty -LiteralPath $_ifPath -Name 'TcpAckFrequency' -Value 1 -Type DWord -ErrorAction Stop
      Set-ItemProperty -LiteralPath $_ifPath -Name 'TCPNoDelay' -Value 1 -Type DWord -ErrorAction Stop
      Write-Log -Message '  -> TcpAckFrequency = 1, TCPNoDelay = 1' -Color Green
      $_anyChanges = $true
      Add-OperationResult -Results $_results -Target $_adapterName -Source 'TCPOptimization' -Action 'Set' -Status 'Completed' -Detail 'TcpAckFrequency=1, TCPNoDelay=1 applied.'
    }
    catch {
      Write-Log -Message "  -> FAILED: $_" -Color Red
      Add-OperationResult -Results $_results -Target $_adapterName -Source 'TCPOptimization' -Action 'Set' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - no changes were made" -Color Yellow
}
elseif ($_anyChanges) {
  if ($Undo) {
    Write-Log -Message "`nTCP tweaks removed. A restart is required for changes to take effect." -Color Green
  }
  else {
    Write-Log -Message "`nTCP tweaks applied. A restart is required for changes to take effect." -Color Green
  }
}
else {
  Write-Log -Message "`nAll adapters were already at the desired target - nothing to do." -Color Green
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Optimize-TCPParameters'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
