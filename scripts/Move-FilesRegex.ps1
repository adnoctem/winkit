#Requires -Version 5.0

<#
.SYNOPSIS
  Moves files whose names match a regular expression.
.DESCRIPTION
  Searches a source directory for files whose file names match the supplied
  regular expression and moves those files into the destination directory.
  Existing destination files are skipped unless -Force is supplied.
.PARAMETER Source
  Directory to search for matching files.
.PARAMETER RegEx
  Regular expression matched against each file name.
.PARAMETER Destination
  Directory that receives matching files. It is created when missing.
.PARAMETER Recurse
  Search child directories below Source.
.PARAMETER Force
  Overwrite existing destination files.
.PARAMETER PassThru
  Return operation result objects for moved and skipped files.
.EXAMPLE
  PS> .\Move-FilesRegex.ps1 -Source C:\Users\Markus\Images -RegEx '^IMG_\d{4}_\d{2}\.(jpg|jpeg|png)$' -Destination D:\Storage\Images
.LINK
  https://github.com/adnoctem/winkit/scripts/Move-FilesRegex.ps1
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
  [string]
  $Source,

  [Parameter(Mandatory = $true)]
  [regex]
  $RegEx,

  [Parameter(Mandatory = $true)]
  [string]
  $Destination,

  [switch]
  $Recurse,

  [switch]
  $Force,

  [switch]
  $PassThru
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path $PSScriptRoot -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

$_results = New-Object System.Collections.ArrayList
$_sourcePath = (Resolve-Path -LiteralPath $Source).ProviderPath

if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
  if ($PSCmdlet.ShouldProcess($Destination, 'Create destination directory')) {
    $null = New-Item -Path $Destination -ItemType Directory -Force
  }
}

$_destinationPath = if (Test-Path -LiteralPath $Destination -PathType Container) {
  (Resolve-Path -LiteralPath $Destination).ProviderPath
}
else {
  $Destination
}

$_files = @(Get-ChildItem -LiteralPath $_sourcePath -File -Recurse:$Recurse | Where-Object { $_.Name -match $RegEx })

if ($_files.Count -eq 0) {
  Write-Log -Message "No files in '$_sourcePath' matched '$RegEx'." -Color Yellow
}

foreach ($_file in $_files) {
  $_targetPath = Join-Path -Path $_destinationPath -ChildPath $_file.Name

  if ((Test-Path -LiteralPath $_targetPath) -and -not $Force) {
    Write-Log -Message "Skipping '$($_file.FullName)' because '$($_targetPath)' already exists." -Color Yellow
    Add-OperationResult -Results $_results -Target $_file.FullName -Source 'FileSystem' -Action 'Move' -Status 'Skipped' -Detail 'DestinationExists'
    continue
  }

  if ($PSCmdlet.ShouldProcess($_file.FullName, "Move to $_targetPath")) {
    Move-Item -LiteralPath $_file.FullName -Destination $_targetPath -Force:$Force
    Write-Log -Message "Moved '$($_file.FullName)' -> '$_targetPath'." -Color Green
    Add-OperationResult -Results $_results -Target $_file.FullName -Source 'FileSystem' -Action 'Move' -Status 'Moved' -Detail $_targetPath
  }
}

if ($PassThru) {
  $_results
}
