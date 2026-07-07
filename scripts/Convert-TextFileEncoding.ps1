#Requires -Version 5.0

<#
.SYNOPSIS
  Converts text files in a repository or directory tree from one encoding to another.

.DESCRIPTION
  Recursively scans a target directory and rewrites matching text files using the
  requested output encoding. This is useful when scripts authored as UTF-8 without
  BOM need to run reliably under Windows PowerShell 5.1, which may otherwise read
  UTF-8 files as the system ANSI code page and produce parsing errors.

  By default, the script assumes it is located in a repository's "scripts" folder
  and processes the repository root by resolving the parent directory of
  $PSScriptRoot.

  The default conversion is from UTF-8 to UTF-8 with BOM.

.PARAMETER Path
  Root directory to process. Defaults to the parent directory of $PSScriptRoot.

.PARAMETER Include
  File patterns to include. Defaults to PowerShell script/module/manifest files:
  *.ps1, *.psm1, *.psd1.

.PARAMETER InputFormat
  Encoding used to read the source files. Defaults to UTF8.

  Supported values:
    UTF8
    UTF8BOM
    Windows1252
    ANSI
    ASCII
    UTF16LE
    UTF16BE

.PARAMETER OutputFormat
  Encoding used to write the converted files. Defaults to UTF8BOM.

  Supported values:
    UTF8
    UTF8BOM
    Windows1252
    ANSI
    ASCII
    UTF16LE
    UTF16BE

.PARAMETER Recurse
  Recursively process child directories. Enabled by default.

.PARAMETER NoRecurse
  Only process files directly in the target directory.

.PARAMETER WhatIf
  Preview which files would be converted without modifying them.

.EXAMPLE
  PS> .\Convert-TextFileEncoding.ps1

  Converts all *.ps1, *.psm1 and *.psd1 files below the repository root from UTF-8
  to UTF-8 with BOM.

.EXAMPLE
  PS> .\Convert-TextFileEncoding.ps1 -OutputFormat UTF8BOM

  Explicitly converts matching files to UTF-8 with BOM.

.EXAMPLE
  PS> .\Convert-TextFileEncoding.ps1 -Path "C:\Repo" -Include *.ps1,*.psm1 -OutputFormat UTF8BOM

  Converts PowerShell files in C:\Repo to UTF-8 with BOM.

.EXAMPLE
  PS> .\Convert-TextFileEncoding.ps1 -OutputFormat Windows1252

  Converts matching files to Windows-1252.

.EXAMPLE
  PS> .\Convert-TextFileEncoding.ps1 -WhatIf

  Shows which files would be converted without writing changes.

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string]$Path = (Split-Path $PSScriptRoot -Parent),

  [Parameter(Mandatory = $false)]
  [string[]]$Include = @('*.ps1', '*.psm1', '*.psd1'),

  [Parameter(Mandatory = $false)]
  [ValidateSet('UTF8', 'UTF8BOM', 'Windows1252', 'ANSI', 'ASCII', 'UTF16LE', 'UTF16BE')]
  [string]$InputFormat = 'UTF8',

  [Parameter(Mandatory = $false)]
  [ValidateSet('UTF8', 'UTF8BOM', 'Windows1252', 'ANSI', 'ASCII', 'UTF16LE', 'UTF16BE')]
  [string]$OutputFormat = 'UTF8BOM',

  [Parameter(Mandatory = $false)]
  [switch]$NoRecurse
)

function Get-TextEncoding {
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('UTF8', 'UTF8BOM', 'Windows1252', 'ANSI', 'ASCII', 'UTF16LE', 'UTF16BE')]
    [string]$Name
  )

  # Required on PowerShell 7+ / .NET Core for legacy code pages such as Windows-1252.
  try {
    Add-Type -AssemblyName System.Text.Encoding.CodePages -ErrorAction Stop
    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
  }
  catch {
    # Windows PowerShell 5.1 / .NET Framework usually does not need this.
    Write-Verbose 'Code page provider registration was skipped or unavailable.'
  }

  switch ($Name) {
    'UTF8' {
      return New-Object System.Text.UTF8Encoding($false)
    }

    'UTF8BOM' {
      return New-Object System.Text.UTF8Encoding($true)
    }

    'Windows1252' {
      return [System.Text.Encoding]::GetEncoding(1252)
    }

    'ANSI' {
      $ansiCodePage = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
      return [System.Text.Encoding]::GetEncoding($ansiCodePage)
    }

    'ASCII' {
      return [System.Text.Encoding]::ASCII
    }

    'UTF16LE' {
      return [System.Text.Encoding]::Unicode
    }

    'UTF16BE' {
      return [System.Text.Encoding]::BigEndianUnicode
    }
  }
}

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
  throw "Path does not exist or is not a directory: $Path"
}

$inputEncoding = Get-TextEncoding -Name $InputFormat
$outputEncoding = Get-TextEncoding -Name $OutputFormat

$searchParams = @{
  LiteralPath = $Path
  File = $true
  Include = $Include
}

if (-not $NoRecurse) {
  $searchParams.Recurse = $true
}

$files = Get-ChildItem @searchParams

if (-not $files) {
  Write-Warning "No matching files found below: $Path"
  return
}

foreach ($file in $files) {
  $filePath = $file.FullName

  if ($PSCmdlet.ShouldProcess($filePath, "Convert from $InputFormat to $OutputFormat")) {
    try {
      $text = [System.IO.File]::ReadAllText($filePath, $inputEncoding)
      [System.IO.File]::WriteAllText($filePath, $text, $outputEncoding)
      Write-Output "Converted: $filePath"
    }
    catch {
      Write-Error "Failed to convert '$filePath': $_"
    }
  }
}
