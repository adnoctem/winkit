Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Applies or removes a curated hosts-file blocklist targeting telemetry,
  crash-report, and location-data hosts.

.DESCRIPTION
  Writes 0.0.0.0-routed block entries into the system hosts file under a
  winkit marker block. The first apply backs up the original hosts file to
  hosts.winkit.bak. Re-application is idempotent — existing winkit entries
  are replaced rather than duplicated.

  This is defence-in-depth that complements policy-based telemetry
  disabling: even if a telemetry channel is not covered by registry
  settings the DNS lookup for its host will not resolve.

.PARAMETER Undo
  Remove the winkit-managed blocklist from the hosts file and restore
  the original file from the backup when it is the only remaining content.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Set-HostsBlocklist.ps1
  Applies the curated blocklist.

.EXAMPLE
  PS> ./Set-HostsBlocklist.ps1 -Undo
  Removes the block and restores the original hosts file.

.EXAMPLE
  PS> ./Set-HostsBlocklist.ps1 -DryRun
  Shows which hosts would be blocked without writing them.

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
    HelpMessage = 'Remove the winkit-managed blocklist and restore the original hosts file.'
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

$_hostsFile = "$env:SystemRoot\System32\drivers\etc\hosts"
$_backupFile = "$env:SystemRoot\System32\drivers\etc\hosts.winkit.bak"

$_markerStart = '# ---- winkit managed blocklist ----'
$_markerEnd = '# ---- end winkit managed blocklist ----'

$_blocklist = @(
  @{
    Host = 'vortex.data.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'vortex-win.data.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'telecommand.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'telecommand.telemetry.microsoft.com.nsatc.net'
    Category = 'Telemetry'
  }
  @{
    Host = 'oca.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'oca.telemetry.microsoft.com.nsatc.net'
    Category = 'Telemetry'
  }
  @{
    Host = 'sqm.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'sqm.telemetry.microsoft.com.nsatc.net'
    Category = 'Telemetry'
  }
  @{
    Host = 'watson.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'watson.telemetry.microsoft.com.nsatc.net'
    Category = 'Telemetry'
  }
  @{
    Host = 'redir.metaservices.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'choice.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'choice.microsoft.com.nsatc.net'
    Category = 'Telemetry'
  }
  @{
    Host = 'df.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'reports.wes.df.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'services.wes.df.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'sqm.df.telemetry.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'telemetry.urs.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'telemetry.appex.bing.net'
    Category = 'Telemetry'
  }
  @{
    Host = 'settings-sandbox.data.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'settings-win.data.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'settings.data.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'diagnostics.support.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'corp.sts.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'statsfe2.ws.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'statsfe1.ws.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'feedback.windows.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'feedback.microsoft-hohm.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'feedback.search.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'pre.footprintpredict.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'i1.services.social.microsoft.com'
    Category = 'Telemetry'
  }
  @{
    Host = 'i1.services.social.microsoft.com.nsatc.net'
    Category = 'Telemetry'
  }
  @{
    Host = 'watson.live.com'
    Category = 'CrashReport'
  }
  @{
    Host = 'watson.ppe.telemetry.microsoft.com'
    Category = 'CrashReport'
  }
  @{
    Host = 'compatexchange.cloudapp.net'
    Category = 'CrashReport'
  }
  @{
    Host = 'inference.location.live.net'
    Category = 'Location'
  }
  @{
    Host = 'rad.msn.com'
    Category = 'Advertising'
  }
  @{
    Host = 'preview.msn.com'
    Category = 'Advertising'
  }
)

if (-not (Test-Path -LiteralPath $_hostsFile -PathType Leaf)) {
  Write-Log -Message "Hosts file not found at '$_hostsFile'." -Color Red
  Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Read' -Status 'Failed' -Detail 'Hosts file not found.'
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

$_currentContent = Get-Content -LiteralPath $_hostsFile -Raw -Encoding Ascii -ErrorAction Stop
$_lines = $_currentContent -split '\r?\n'

$hasMarker = ($_lines | Where-Object { $_ -eq $_markerStart }).Count -gt 0

# ---- Undo --------------------------------------------------------------------
if ($Undo) {
  if (-not $hasMarker) {
    Write-Log -Message 'No winkit-managed blocklist found in hosts file - nothing to undo.' -Color Yellow
    Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Remove' -Status 'Skipped' -Detail 'No winkit blocklist present.'
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }

  Write-Log -Message 'Removing winkit-managed hosts blocklist.' -Color Yellow

  if ($DryRun) {
    Write-Log -Message '[DRY RUN] Would remove winkit entries and restore from backup.' -Color Yellow
    Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Remove' -Status 'Skipped' -Detail 'DryRun'
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }

  $startIdx = -1
  $endIdx = -1
  for ($i = 0; $i -lt $_lines.Count; $i++) {
    if ($_lines[$i] -eq $_markerStart) { $startIdx = $i }
    if ($_lines[$i] -eq $_markerEnd) { $endIdx = $i; break }
  }

  if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
    $_newLines = $_lines[0..($startIdx - 1)] + $_lines[($endIdx + 1)..($_lines.Count - 1)]
  }
  elseif ($startIdx -ge 0) {
    $_newLines = $_lines[0..($startIdx - 1)]
  }
  else {
    $_newLines = $_lines
  }

  $_newContent = ($_newLines -join "`r`n").TrimEnd("`r`n")
  if (-not $_newContent) {
    if (Test-Path -LiteralPath $_backupFile) {
      Write-Log -Message '  -> Restoring hosts file from backup.' -Color Green
      Copy-Item -LiteralPath $_backupFile -Destination $_hostsFile -Force -ErrorAction Stop
      Remove-Item -LiteralPath $_backupFile -Force -ErrorAction SilentlyContinue
    }
    else {
      Write-Log -Message '  -> Hosts file empty after removal. Restoring default hosts template.' -Color Yellow
      $_defaultTemplate = @"
# Copyright (c) 1993-2009 Microsoft Corp.
#
# This is a sample HOSTS file used by Microsoft TCP/IP for Windows.
#
# This file contains the mappings of IP addresses to host names. Each
# entry should be kept on an individual line. The IP address should
# be placed in the first column followed by the corresponding host name.
# The IP address and the host name should be separated by at least one
# space.
#
# Additionally, comments (such as these) may be inserted on individual
# lines or following the machine name denoted by a '#' symbol.
#
# For example:
#
#      102.54.94.97     rhino.acme.com          # source server
#       38.25.63.10     x.acme.com              # x client host

# localhost name resolution is handled within DNS itself.
#   127.0.0.1       localhost
#   ::1             localhost
"@
      Set-Content -LiteralPath $_hostsFile -Value $_defaultTemplate -Encoding Ascii -ErrorAction Stop
    }
  }
  else {
    $_newContent += "`r`n"
    Set-Content -LiteralPath $_hostsFile -Value $_newContent -Encoding Ascii -ErrorAction Stop
  }

  Write-Log -Message '  -> Winkit blocklist removed.' -Color Green
  Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Remove' -Status 'Completed' -Detail 'Blocklist removed.'
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Set-HostsBlocklist'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

# ---- Apply -------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $_backupFile)) {
  Write-Log -Message 'Backing up original hosts file.' -Color Yellow
  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would create backup: $_backupFile" -Color Yellow
  }
  else {
    Copy-Item -LiteralPath $_hostsFile -Destination $_backupFile -Force -ErrorAction Stop
    Write-Log -Message "  -> Backup created: $_backupFile" -Color Green
    Add-OperationResult -Results $_results -Target $_backupFile -Source 'HostsFile' -Action 'Backup' -Status 'Completed' -Detail 'Original hosts file backed up.'
  }
}

if ($hasMarker) {
  Write-Log -Message 'Winkit blocklist already present - replacing entries.' -Color Yellow

  $startIdx = -1
  $endIdx = -1
  for ($i = 0; $i -lt $_lines.Count; $i++) {
    if ($_lines[$i] -eq $_markerStart) { $startIdx = $i }
    if ($_lines[$i] -eq $_markerEnd) { $endIdx = $i; break }
  }

  if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
    $_lines = $_lines[0..($startIdx - 1)] + $_lines[($endIdx + 1)..($_lines.Count - 1)]
  }
}

$_blockLines = @("", $_markerStart)
foreach ($entry in $_blocklist) {
  $_blockLines += "# $($entry.Category): $($entry.Host)"
  $_blockLines += "0.0.0.0 $($entry.Host)"
}
$_blockLines += $_markerEnd

$_targetContent = foreach ($_line in $_lines + $_blockLines) { $_line }
$_targetContent = ($_targetContent -join "`r`n").TrimEnd("`r`n") + "`r`n"

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would write $($_blocklist.Count) block entries to hosts file." -Color Yellow
  Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Set' -Status 'Skipped' -Detail "DryRun: $($_blocklist.Count) hosts."
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

if (-not $PSCmdlet.ShouldProcess($_hostsFile, "Apply $($_blocklist.Count) hosts-file block entries")) {
  Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Set' -Status 'Skipped' -Detail 'WhatIf'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

Write-Log -Message "Writing $($_blocklist.Count) block entries to hosts file." -Color Yellow
try {
  Set-Content -LiteralPath $_hostsFile -Value $_targetContent -Encoding Ascii -ErrorAction Stop
  Write-Log -Message '  -> Hosts blocklist applied.' -Color Green
  Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Set' -Status 'Completed' -Detail "$($_blocklist.Count) hosts blocked."
}
catch {
  Write-Log -Message "  -> FAILED: $_" -Color Red
  Add-OperationResult -Results $_results -Target $_hostsFile -Source 'HostsFile' -Action 'Set' -Status 'Failed' -Detail $_.Exception.Message
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Set-HostsBlocklist'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
