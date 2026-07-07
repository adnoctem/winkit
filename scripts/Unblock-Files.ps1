#Requires -Version 5.0

<#
.SYNOPSIS
  Unblocks files downloaded from the internet by removing the Mark-of-the-Web
  alternate data stream.

.DESCRIPTION
  Recursively scans a directory and unblocks every file whose extension matches
  the supplied -Filter.  This removes the security warning that Windows attaches
  to files originating from the internet, browsers, or email clients.
  Extensions are matched case-insensitively so '.pdf', '.PDF', and '.Pdf' are
  all treated equally.

.PARAMETER Path
  The root directory to scan recursively.

.PARAMETER Filter
  A comma-separated list of file extensions to target (dot is optional).
  Examples: 'pdf', 'pdf,docx,xlsx', '.jpg,.png,.3mf'.

.PARAMETER DryRun
  Preview which files would be unblocked without making changes.
.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Unblock-Files.ps1 -Path 'C:\Invoices' -Filter 'pdf'
  Recursively unblocks all .pdf files under C:\Invoices.

.EXAMPLE
  PS> ./Unblock-Files.ps1 -Path 'C:\Downloads' -Filter 'pdf,docx,xlsx'
  Recursively unblocks all .pdf, .docx, and .xlsx files.

.EXAMPLE
  PS> ./Unblock-Files.ps1 -Path 'C:\Models' -Filter '3mf,jpg,png'
  Recursively unblocks all .3mf, .jpg, and .png files - useful for 3D printing
  assets or images downloaded from a browser.

.EXAMPLE
  PS> ./Unblock-Files.ps1 -Path 'C:\Downloads' -Filter 'zip,msi,exe' -DryRun
  Previews which .zip, .msi, and .exe files would be unblocked without touching
  them.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
# Parameters
param (
  [Parameter(
    Position = 0,
    Mandatory = $true,
    HelpMessage = "The root directory to scan recursively, e.g. 'C:\Downloads'."
  )]
  [string]
  $Path,

  [Parameter(
    Position = 1,
    Mandatory = $true,
    HelpMessage = "Comma-separated file extensions to target, e.g. 'pdf,docx,xlsx'."
  )]
  [string]
  $Filter,

  [Parameter(
    Position = 2,
    Mandatory = $false,
    HelpMessage = 'Preview which files would be unblocked without making changes.'
  )]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

# When -DryRun is active, enable WhatIf for downstream lib calls and log intent.
if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no files will be modified`n" -Color Yellow
}

# Validate the root path
if (-not (Test-Path -Path $Path -PathType Container)) {
  Write-Log -Message "Path not found or not a directory: '$Path'" -Color Red
  exit 1
}

# Parse the filter into a case-insensitive extension regex.
# Accepts 'pdf', '.pdf', 'pdf,jpg', '.pdf, .jpg' etc.
$extensions = $Filter -split ',' |
  ForEach-Object { $_.Trim().TrimStart('.') } |
  Where-Object { $_ -ne '' }

if ($extensions.Count -eq 0) {
  Write-Log -Message 'No valid extensions found in -Filter - nothing to do.' -Color Red
  exit 1
}

$extensionPattern = '\.(' + ($extensions -join '|') + ')$'
Write-Log -Message "Scanning '$Path' for files matching: $($extensions -join ', ') ..." -Color Yellow

# Collect matching files
$files = @(Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -match $extensionPattern })

if ($files.Count -eq 0) {
  Write-Log -Message '  -> No matching files found.' -Color Gray
  exit 0
}

Write-Log -Message "  -> Found $($files.Count) matching file(s)." -Color Gray

# Unblock each file
$unblocked = 0
$failed = 0
$results = New-Object System.Collections.ArrayList

foreach ($file in $files) {
  $relativePath = $file.FullName.Replace($Path, '').TrimStart('\')

  if ($DryRun) {
    Write-Log -Message "[DRY RUN] Would unblock: $relativePath" -Color Yellow
    Add-OperationResult -Results $results -Target $file.FullName -Source 'FileSystem' -Action 'Unblock' -Status 'Skipped' -Detail 'DryRun'
    continue
  }

  try {
    Unblock-File -Path $file.FullName -ErrorAction Stop
    Write-Log -Message "Unblocked: $relativePath" -Color Green
    Add-OperationResult -Results $results -Target $file.FullName -Source 'FileSystem' -Action 'Unblock' -Status 'Completed' -Detail 'Mark-of-the-Web removed.'
    $unblocked++
  }
  catch {
    Write-Log -Message "FAILED: $relativePath - $_" -Color Red
    Add-OperationResult -Results $results -Target $file.FullName -Source 'FileSystem' -Action 'Unblock' -Status 'Failed' -Detail $_.Exception.Message
    $failed++
  }
}

# Summary
if ($DryRun) {
  Write-Log -Message "`nDRY RUN COMPLETE - $($files.Count) file(s) would have been unblocked" -Color Yellow
}
else {
  Write-Log -Message "`nUnblocked: $unblocked  |  Failed: $failed  |  Total matched: $($files.Count)" -Color $(if ($failed -gt 0) { 'Yellow' } else { 'Green' })
}

$_operationLog = Write-OperationResultLog -Results $results -ScriptName 'Unblock-Files'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $results
}
