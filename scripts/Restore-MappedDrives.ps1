#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Reconnects mapped network drives.

.DESCRIPTION
  Enumerates current SMB mappings (drive-letter - remote share) and
  re-establishes each one. Useful after a network change, VPN reconnect,
  or when drives show as "Unavailable" / "Disconnected" because they
  were mapped while the server was unreachable.

  Uses the SmbMapping cmdlets (Get-SmbMapping / New-SmbMapping) rather
  than parsing 'net use' text. Driveless SMB connections (UNC sessions
  with no letter) are out of scope.

.PARAMETER Credential
  Optional credential for re-establishing mappings that require
  authentication. If omitted, mappings reconnect using the current
  security context.

.PARAMETER Persistent
  Make the re-established mappings persistent (reconnect at sign-in).
  Defaults to true.

.PARAMETER DryRun
  Preview which drives would be reconnected without making changes.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Restore-MappedDrives.ps1
  Reconnects all current mappings using the current security context.

.EXAMPLE
  PS> ./Restore-MappedDrives.ps1 -Credential (Get-Credential)
  Reconnects, supplying explicit credentials for shares that need them.

.EXAMPLE
  PS> ./Restore-MappedDrives.ps1 -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [System.Management.Automation.PSCredential]
  $Credential,

  [Parameter(Mandatory = $false)]
  [bool]
  $Persistent = $true,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no drive mappings will be changed`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

Write-Log -Message 'Scanning for mapped network drives...' -Color Yellow

$mappings = @(Get-SmbMapping -ErrorAction SilentlyContinue |
    Where-Object { $_.LocalPath -match '^[A-Za-z]:$' })

if ($mappings.Count -eq 0) {
  Write-Log -Message 'No mapped drives found. Nothing to do.' -Color Gray
  exit 0
}

Write-Log -Message "  -> $($mappings.Count) mapped drive(s) found`n" -Color Gray

$_reconnected = 0
$_failed = 0
$_skipped = 0

foreach ($m in $mappings) {
  $drive = $m.LocalPath
  $remote = $m.RemotePath
  $status = $m.Status

  if ($status -eq 'OK') {
    Add-OperationResult -Results $_results -Target $drive -Source 'SmbMapping' -Action 'Reconnect' -Status 'Skipped' -Detail 'Already connected.'
    $_skipped++
    continue
  }

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would reconnect $drive -> $remote (currently '$status')" -Color Yellow
    Add-OperationResult -Results $_results -Target $drive -Source 'SmbMapping' -Action 'Reconnect' -Status 'Skipped' -Detail 'DryRun'
    continue
  }

  if (-not $PSCmdlet.ShouldProcess("$drive -> $remote", 'Reconnect mapped drive')) {
    Add-OperationResult -Results $_results -Target $drive -Source 'SmbMapping' -Action 'Reconnect' -Status 'Skipped' -Detail 'WhatIf'
    continue
  }

  Write-Log -Message "  $drive is currently '$status' - reconnecting..." -Color Gray

  try {
    Remove-SmbMapping -LocalPath $drive -Force -UpdateProfile -ErrorAction SilentlyContinue

    $params = @{
      LocalPath = $drive
      RemotePath = $remote
      Persistent = $Persistent
      ErrorAction = 'Stop'
    }
    if ($Credential) {
      $params['UserName'] = $Credential.UserName
      $params['Password'] = $Credential.GetNetworkCredential().Password
    }

    $null = New-SmbMapping @params
    Write-Log -Message "    Reconnected $drive -> $remote." -Color Green
    Add-OperationResult -Results $_results -Target $drive -Source 'SmbMapping' -Action 'Reconnect' -Status 'Completed' -Detail "$drive -> $remote"
    $_reconnected++
  }
  catch {
    Write-Log -Message "    FAILED - $drive -> $remote ($($_.Exception.Message))" -Color Red
    Add-OperationResult -Results $_results -Target $drive -Source 'SmbMapping' -Action 'Reconnect' -Status 'Failed' -Detail $_.Exception.Message
    $_failed++
  }
}

# ---- Summary ----------------------------------------------------------------
Write-Log -Message "`nReconnected: $_reconnected | Skipped (OK): $_skipped | Failed: $_failed | Total: $($mappings.Count)" -Color $(
  if ($_failed -gt 0) { 'Yellow' } else { 'Green' })

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Restore-MappedDrives'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
