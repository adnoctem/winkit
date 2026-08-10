#Requires -Version 5.0

<#
.SYNOPSIS
  Builds deployable winkit source archives.

.DESCRIPTION
  Creates a clean bundle containing the repository's scripts, bin, and
  resources directories, preserving their relative layout so scripts and
  launchers can be distributed and unpacked as a unit.

  Archives are written to the dist directory, which is created when it does not
  already exist. By default, the script builds both:

    dist/winkit.zip
    dist/winkit.tar.gz
    dist/CHECKSUMS_SHA256.txt

  Existing archives with the same names are overwritten. Archives are produced
  directly from the source directories — no staging copy is made, which avoids
  the file-handle contention that can occur when cleaning up a staging folder
  immediately after an archiver has read from it. The checksum file lists the
  SHA-256 hash of every archive produced in the same run; it is written after
  the archives so only existing files are referenced.

.PARAMETER OutputDirectory
  Directory where archives are written. Defaults to the repository dist folder.

.PARAMETER Name
  Base archive name without extension. Defaults to winkit.

.PARAMETER Format
  Archive format to build. Use Zip, TarGz, or Both. Defaults to Both.

.EXAMPLE
  PS> ./build.ps1
  Creates dist/winkit.zip, dist/winkit.tar.gz, and dist/CHECKSUMS_SHA256.txt.

.EXAMPLE
  PS> ./build.ps1 -Format Zip
  Creates only dist/winkit.zip.

.EXAMPLE
  PS> ./build.ps1 -OutputDirectory C:\Temp -Name winkit-vm-test
  Creates C:\Temp\winkit-vm-test.zip and C:\Temp\winkit-vm-test.tar.gz.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [string]$OutputDirectory = (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'dist'),

  [ValidateNotNullOrEmpty()]
  [string]$Name = 'winkit',

  [ValidateSet('Both', 'Zip', 'TarGz')]
  [string]$Format = 'Both'
)

$ErrorActionPreference = 'Stop'

$repositoryRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Split-Path -Path $PSScriptRoot -Parent))
$outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$zipPath = Join-Path -Path $outputPath -ChildPath "$Name.zip"
$tarGzPath = Join-Path -Path $outputPath -ChildPath "$Name.tar.gz"
$checksumPath = Join-Path -Path $outputPath -ChildPath 'CHECKSUMS_SHA256.txt'

# The source directories that make up a bundle. Validated up front so a
# missing directory fails before any archive work begins.
$sourceDirs = 'scripts', 'bin', 'resources'
$sourcePaths = foreach ($dir in $sourceDirs) {
  $path = Join-Path -Path $repositoryRoot -ChildPath $dir
  if (-not (Test-Path -LiteralPath $path -PathType Container)) {
    throw "Required build source directory not found: $path"
  }
  $path
}

function Clear-BuildPath {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  if (Test-Path -LiteralPath $Path) {
    Remove-Item -LiteralPath $Path -Recurse -Force
  }
}

if (-not (Test-Path -LiteralPath $outputPath -PathType Container)) {
  New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
}

if ($Format -eq 'Both' -or $Format -eq 'Zip') {
  if ($PSCmdlet.ShouldProcess($zipPath, 'Create ZIP archive')) {
    Clear-BuildPath -Path $zipPath
    # Compress-Archive accepts multiple -Path roots and preserves each top-level
    # directory name, so 'lib' and 'scripts' land in the archive exactly as they
    # are on disk — no staging copy needed.
    Compress-Archive -Path $sourcePaths -DestinationPath $zipPath -Force
    Write-Output "Built: $zipPath"
  }
}

if ($Format -eq 'Both' -or $Format -eq 'TarGz') {
  $tarCommand = Get-Command -Name tar -ErrorAction SilentlyContinue
  if (-not $tarCommand) {
    throw 'tar was not found on PATH. Build the ZIP archive instead or install a tar-compatible tool.'
  }

  if ($PSCmdlet.ShouldProcess($tarGzPath, 'Create tar.gz archive')) {
    Clear-BuildPath -Path $tarGzPath
    # -C changes tar's working directory to the repo root before archiving, so
    # the stored paths are 'scripts/...' and 'bin/...' rather than absolute or
    # deeply-nested. No staging copy and no Push-Location needed.
    & $tarCommand.Source -C $repositoryRoot -czf $tarGzPath $sourceDirs
    if ($LASTEXITCODE -ne 0) {
      throw "tar exited with code $LASTEXITCODE while creating $tarGzPath."
    }
    Write-Output "Built: $tarGzPath"
  }
}

# Write SHA-256 checksums for every archive produced in this run.
if ($PSCmdlet.ShouldProcess($checksumPath, 'Write SHA-256 checksums')) {
  $_archives = @()
  if (Test-Path -LiteralPath $zipPath -PathType Leaf) { $_archives += $zipPath }
  if (Test-Path -LiteralPath $tarGzPath -PathType Leaf) { $_archives += $tarGzPath }

  if ($_archives.Count -gt 0) {
    $_lines = foreach ($_archive in $_archives) {
      $_hash = (Get-FileHash -LiteralPath $_archive -Algorithm SHA256).Hash
      "{0}  {1}" -f $_hash, (Split-Path -Path $_archive -Leaf)
    }
    # UTF-8 without BOM so the file is script-friendly on any platform.
    [System.IO.File]::WriteAllLines($checksumPath, [string[]]$_lines, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output "Built: $checksumPath"
  }
}
