#Requires -Version 5.0

<#
.SYNOPSIS
  Installs or updates a managed winkit release.

.DESCRIPTION
  Downloads a winkit release from GitHub, verifies the release ZIP against its
  published SHA-256 checksum, validates the archive layout, installs the pinned
  PSFoundation runtime dependency, and activates the release at a stable path.

  Re-running the script updates an existing managed installation. Updates are
  staged beside the installation and rolled back if activation fails. Existing
  directories without winkit's ownership manifest are never overwritten.

  The installer is intentionally standalone and does not import PSFoundation.

.PARAMETER Scope
  Installation scope. CurrentUser installs below LocalAppData and updates the
  user PATH. AllUsers installs below Program Files, updates the machine PATH,
  and requires an elevated PowerShell session. Defaults to CurrentUser.

.PARAMETER InstallPath
  Destination directory. Defaults to %LOCALAPPDATA%\Programs\winkit for
  CurrentUser or %ProgramFiles%\winkit for AllUsers.

.PARAMETER Repository
  GitHub release repository in OWNER/REPOSITORY form. Defaults to
  adnoctem/winkit. Managed installations remain bound to their original
  repository.

.PARAMETER Version
  Semantic release version to install, with or without a leading v. The latest
  stable GitHub release is selected when this parameter is omitted.

.PARAMETER NoPath
  Do not add the installed bin directory to the persistent PATH.

.PARAMETER Force
  Download and reinstall the selected release even when that version is already
  installed. Force never permits replacement of an unrecognized directory.

.PARAMETER NonInteractive
  Suppress installer confirmation prompts and use safe defaults. The
  WINKIT_NON_INTERACTIVE environment variable accepts the same behavior when
  set to 1 or true. Unsafe conflicts still fail instead of being accepted.

.PARAMETER DryRun
  Resolve the selected release and show the intended operation without
  downloading or changing the system.

.PARAMETER PassThru
  Return a result object describing the installation operation.

.EXAMPLE
  PS> irm https://raw.githubusercontent.com/adnoctem/winkit/main/install.ps1 | iex
  Installs or updates the latest stable winkit release for the current user.

.EXAMPLE
  PS> & ([scriptblock]::Create((irm 'https://raw.githubusercontent.com/adnoctem/winkit/main/install.ps1'))) -InstallPath 'D:\Tools\winkit'
  Installs winkit at a custom location.

.EXAMPLE
  PS> $env:WINKIT_NON_INTERACTIVE = '1'; irm https://raw.githubusercontent.com/adnoctem/winkit/main/install.ps1 | iex
  Installs or updates winkit without installer prompts.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core support: Yes.
  SYSTEM-account suitability: Use -Scope AllUsers and an explicit InstallPath;
  CurrentUser installs target the invoking account's profile.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [ValidateSet('CurrentUser', 'AllUsers')]
  [string]$Scope,

  [ValidateNotNullOrEmpty()]
  [string]$InstallPath,

  [ValidatePattern('^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$')]
  [string]$Repository = 'adnoctem/winkit',

  [ValidatePattern('^v?(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-[0-9A-Za-z.-]+)?(\+[0-9A-Za-z.-]+)?$')]
  [string]$Version,

  [switch]$NoPath,

  [switch]$Force,

  [switch]$NonInteractive,

  [switch]$DryRun,

  [switch]$PassThru
)

& {
  param (
    [Parameter(Mandatory = $true)]
    [pscustomobject]$Context
  )

  Set-StrictMode -Version 2.0
  $ErrorActionPreference = 'Stop'

  function Write-InstallerMessage {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Message
    )

    Write-Information -MessageData $Message -InformationAction Continue
  }

  function Get-EnvironmentBoolean {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Name
    )

    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if ([string]::IsNullOrWhiteSpace($value)) {
      return $false
    }

    switch ($value.Trim().ToLowerInvariant()) {
      { $_ -in @('1', 'true', 'yes', 'on') } { return $true }
      { $_ -in @('0', 'false', 'no', 'off') } { return $false }
      default { throw "$Name must be 1/true/yes/on or 0/false/no/off." }
    }
  }

  function Get-NormalizedPath {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Path
    )

    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    $fullPath = [System.IO.Path]::GetFullPath($expanded)
    $root = [System.IO.Path]::GetPathRoot($fullPath)
    $trimmed = $fullPath.TrimEnd([char[]]@('\', '/'))

    if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.Equals($root.TrimEnd([char[]]@('\', '/')), [System.StringComparison]::OrdinalIgnoreCase)) {
      throw 'Choose a dedicated installation directory rather than a drive root.'
    }

    return $trimmed
  }

  function Test-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  }

  function Get-InstallState {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
      return $null
    }

    $item = Get-Item -LiteralPath $Path -Force
    if (-not $item.PSIsContainer) {
      throw "InstallPath exists and is not a directory: $Path"
    }
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
      throw "Refusing an installation directory that is a reparse point: $Path"
    }

    $statePath = Join-Path -Path $Path -ChildPath '.winkit-install.json'
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) {
      throw "Refusing to overwrite an unrecognized directory without .winkit-install.json: $Path"
    }

    try {
      $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
      throw "Could not read the winkit ownership manifest at '$statePath': $($_.Exception.Message)"
    }

    $requiredProperties = @('schemaVersion', 'repository', 'version', 'scope', 'installPath')
    foreach ($property in $requiredProperties) {
      if ($state.PSObject.Properties.Name -notcontains $property) {
        throw "The winkit ownership manifest is missing '$property'."
      }
    }

    if ([int]$state.schemaVersion -ne 1) {
      throw "Unsupported winkit ownership manifest schema: $($state.schemaVersion)"
    }
    if (-not ([string]$state.installPath).Equals($Path, [System.StringComparison]::OrdinalIgnoreCase)) {
      throw 'The winkit ownership manifest does not match InstallPath.'
    }
    if ([string]$state.scope -notin @('CurrentUser', 'AllUsers')) {
      throw "Invalid installation scope in the ownership manifest: $($state.scope)"
    }
    if ([string]$state.repository -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') {
      throw 'Invalid repository in the winkit ownership manifest.'
    }

    return $state
  }

  function Get-WinkitRelease {
    param (
      [Parameter(Mandatory = $true)]
      [string]$ReleaseRepository,

      [string]$ReleaseVersion
    )

    $headers = @{
      Accept                 = 'application/vnd.github+json'
      'User-Agent'           = 'winkit-installer'
      'X-GitHub-Api-Version' = '2022-11-28'
    }

    if ($ReleaseVersion) {
      $tag = $ReleaseVersion
      if (-not $tag.StartsWith('v', [System.StringComparison]::OrdinalIgnoreCase)) {
        $tag = "v$tag"
      }
      $escapedTag = [Uri]::EscapeDataString($tag)
      $uri = "https://api.github.com/repos/$ReleaseRepository/releases/tags/$escapedTag"
    }
    else {
      $uri = "https://api.github.com/repos/$ReleaseRepository/releases/latest"
    }

    Write-InstallerMessage -Message "Resolving winkit release from $ReleaseRepository..."
    $release = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop

    if (-not $release.tag_name -or [string]$release.tag_name -notmatch '^v?(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-[0-9A-Za-z.-]+)?(\+[0-9A-Za-z.-]+)?$') {
      throw 'GitHub returned a release without a valid semantic version tag.'
    }

    $zipAssets = @($release.assets | Where-Object { $_.name -eq 'winkit.zip' })
    $checksumAssets = @($release.assets | Where-Object { $_.name -eq 'CHECKSUMS_SHA256.txt' })
    if ($zipAssets.Count -ne 1 -or $checksumAssets.Count -ne 1) {
      throw "Release '$($release.tag_name)' must contain exactly one winkit.zip and CHECKSUMS_SHA256.txt asset."
    }

    return [pscustomobject]@{
      Version      = ([string]$release.tag_name -replace '^v', '')
      Tag          = [string]$release.tag_name
      ZipUri       = [string]$zipAssets[0].browser_download_url
      ChecksumUri  = [string]$checksumAssets[0].browser_download_url
      PublishedUtc = [string]$release.published_at
    }
  }

  function Invoke-FileDownload {
    param (
      [Parameter(Mandatory = $true)]
      [ValidatePattern('^https://')]
      [string]$Uri,

      [Parameter(Mandatory = $true)]
      [string]$Destination
    )

    $parameters = @{
      Uri             = $Uri
      OutFile         = $Destination
      UseBasicParsing = $true
      ErrorAction     = 'Stop'
    }
    Invoke-WebRequest @parameters | Out-Null
  }

  function Test-ArchiveChecksum {
    param (
      [Parameter(Mandatory = $true)]
      [string]$ArchivePath,

      [Parameter(Mandatory = $true)]
      [string]$ChecksumPath
    )

    $expectedHashes = @()
    foreach ($line in [System.IO.File]::ReadAllLines($ChecksumPath)) {
      $match = [regex]::Match($line, '^([0-9A-Fa-f]{64})\s+\*?winkit\.zip$')
      if ($match.Success) {
        $expectedHashes += $match.Groups[1].Value.ToUpperInvariant()
      }
    }

    if ($expectedHashes.Count -ne 1) {
      throw 'CHECKSUMS_SHA256.txt must contain exactly one valid winkit.zip entry.'
    }

    $actualHash = (Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actualHash -ne $expectedHashes[0]) {
      throw 'The winkit.zip SHA-256 checksum does not match the published checksum.'
    }
  }

  function Test-ReleaseArchive {
    param (
      [Parameter(Mandatory = $true)]
      [string]$ArchivePath,

      [Parameter(Mandatory = $true)]
      [string]$DestinationPath
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    $destinationRoot = [System.IO.Path]::GetFullPath($DestinationPath).TrimEnd([char[]]@('\', '/'))
    $destinationPrefix = "$destinationRoot$([System.IO.Path]::DirectorySeparatorChar)"
    $seenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $requiredPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $allowedRoots = @('scripts', 'bin', 'resources', 'requirements.psd1', 'LICENSE')
    [long]$expandedSize = 0

    try {
      foreach ($entry in $archive.Entries) {
        $relativePath = $entry.FullName.Replace('\', '/')
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
          throw 'The release archive contains an empty path.'
        }
        if ($relativePath.StartsWith('/') -or $relativePath.Contains(':') -or $relativePath -match '[\x00-\x1F]') {
          throw "The release archive contains an unsafe path: $relativePath"
        }

        $segments = @($relativePath.TrimEnd('/').Split('/'))
        if ($segments.Count -eq 0 -or $allowedRoots -notcontains $segments[0]) {
          throw "The release archive contains an unexpected top-level path: $relativePath"
        }
        foreach ($segment in $segments) {
          if ([string]::IsNullOrWhiteSpace($segment) -or $segment -in @('.', '..') -or $segment.EndsWith('.') -or $segment.EndsWith(' ')) {
            throw "The release archive contains an unsafe path segment: $relativePath"
          }
        }

        $candidate = [System.IO.Path]::GetFullPath((Join-Path -Path $destinationRoot -ChildPath ($relativePath.Replace('/', '\'))))
        if (-not $candidate.StartsWith($destinationPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
          throw "The release archive path escapes the staging directory: $relativePath"
        }
        if (-not $seenPaths.Add($candidate)) {
          throw "The release archive contains duplicate Windows paths: $relativePath"
        }

        $unixType = ($entry.ExternalAttributes -shr 16) -band 0xF000
        if ($unixType -eq 0xA000) {
          throw "The release archive contains an unsupported symbolic link: $relativePath"
        }
        if ($unixType -notin @(0, 0x4000, 0x8000)) {
          throw "The release archive contains an unsupported special file: $relativePath"
        }

        $expandedSize += $entry.Length
        if ($entry.Length -gt 268435456 -or $expandedSize -gt 536870912) {
          throw 'The release archive exceeds the installer extraction limit.'
        }

        [void]$requiredPaths.Add($relativePath.TrimEnd('/'))
      }
    }
    finally {
      $archive.Dispose()
    }

    foreach ($requiredPath in @('scripts/Invoke-Bootstrap.ps1', 'bin/bootstrap.cmd', 'requirements.psd1', 'LICENSE')) {
      if (-not $requiredPaths.Contains($requiredPath)) {
        throw "The release archive is missing required path '$requiredPath'."
      }
    }
  }

  function Get-RuntimeRequirement {
    param (
      [Parameter(Mandatory = $true)]
      [string]$RequirementsPath
    )

    $requirements = Import-PowerShellDataFile -LiteralPath $RequirementsPath -ErrorAction Stop
    if (-not $requirements.ContainsKey('PSFoundation')) {
      throw 'requirements.psd1 does not define PSFoundation.'
    }

    $requiredVersion = [string]$requirements.PSFoundation
    if ($requiredVersion -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-[0-9A-Za-z.-]+)?$') {
      throw "requirements.psd1 contains an invalid PSFoundation version: $requiredVersion"
    }

    return $requiredVersion
  }

  function Test-ModuleScope {
    param (
      [Parameter(Mandatory = $true)]
      [psobject]$Module,

      [Parameter(Mandatory = $true)]
      [ValidateSet('CurrentUser', 'AllUsers')]
      [string]$InstallScope
    )

    if ($InstallScope -eq 'CurrentUser') {
      return $true
    }

    $userProfile = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile).TrimEnd([char[]]@('\', '/'))
    $userPrefix = "$userProfile$([System.IO.Path]::DirectorySeparatorChar)"
    return -not ([string]$Module.ModuleBase).StartsWith($userPrefix, [System.StringComparison]::OrdinalIgnoreCase)
  }

  function Invoke-DependencyInstall {
    param (
      [Parameter(Mandatory = $true)]
      [string]$RequiredVersion,

      [Parameter(Mandatory = $true)]
      [ValidateSet('CurrentUser', 'AllUsers')]
      [string]$InstallScope,

      [switch]$Reinstall
    )

    $installed = @(Get-Module -ListAvailable -Name PSFoundation |
        Where-Object { $_.Version -eq [version]$RequiredVersion } |
        Where-Object { Test-ModuleScope -Module $_ -InstallScope $InstallScope })
    if ($installed.Count -gt 0 -and -not $Reinstall) {
      Write-InstallerMessage -Message "PSFoundation $RequiredVersion is already installed."
      return $false
    }

    if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) {
      throw 'Install-Module is unavailable. Install PowerShellGet, then rerun the winkit installer.'
    }

    if (Get-Command -Name Get-PackageProvider -ErrorAction SilentlyContinue) {
      $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
      if (-not $nuget) {
        Write-InstallerMessage -Message 'Installing the NuGet package provider...'
        $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $InstallScope -Force -Confirm:$false -ErrorAction Stop
      }
    }

    Write-InstallerMessage -Message "Installing PSFoundation $RequiredVersion for $InstallScope..."
    $installParameters = @{
      Name            = 'PSFoundation'
      RequiredVersion = $RequiredVersion
      Repository      = 'PSGallery'
      Scope           = $InstallScope
      Force           = $true
      AllowClobber    = $true
      Confirm         = $false
      ErrorAction     = 'Stop'
    }
    Install-Module @installParameters

    $verified = @(Get-Module -ListAvailable -Name PSFoundation |
        Where-Object { $_.Version -eq [version]$RequiredVersion } |
        Where-Object { Test-ModuleScope -Module $_ -InstallScope $InstallScope })
    if ($verified.Count -eq 0) {
      throw "PSFoundation $RequiredVersion was not found after installation."
    }

    return $true
  }

  function Invoke-PathUpdate {
    param (
      [Parameter(Mandatory = $true)]
      [string]$BinPath,

      [Parameter(Mandatory = $true)]
      [ValidateSet('CurrentUser', 'AllUsers')]
      [string]$InstallScope
    )

    $environmentTarget = if ($InstallScope -eq 'AllUsers') { 'Machine' } else { 'User' }
    $persistentPath = [Environment]::GetEnvironmentVariable('Path', $environmentTarget)
    $parts = @($persistentPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $normalizedBin = $BinPath.TrimEnd([char[]]@('\', '/'))
    $alreadyPresent = $false

    foreach ($part in $parts) {
      if ($part.Trim().TrimEnd([char[]]@('\', '/')).Equals($normalizedBin, [System.StringComparison]::OrdinalIgnoreCase)) {
        $alreadyPresent = $true
        break
      }
    }

    $persistentChanged = $false
    if (-not $alreadyPresent) {
      $newPath = (@($parts) + $BinPath) -join ';'
      [Environment]::SetEnvironmentVariable('Path', $newPath, $environmentTarget)
      $persistentChanged = $true
      Write-InstallerMessage -Message "Added '$BinPath' to the $environmentTarget PATH."
    }
    else {
      Write-InstallerMessage -Message "The $environmentTarget PATH already contains '$BinPath'."
    }

    $processParts = @($env:Path -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $processContainsPath = $false
    foreach ($part in $processParts) {
      if ($part.Trim().TrimEnd([char[]]@('\', '/')).Equals($normalizedBin, [System.StringComparison]::OrdinalIgnoreCase)) {
        $processContainsPath = $true
        break
      }
    }
    if (-not $processContainsPath) {
      $env:Path = (@($processParts) + $BinPath) -join ';'
    }

    if ($persistentChanged) {
      try {
        if (-not ('WinkitInstaller.NativeMethods' -as [type])) {
          Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace WinkitInstaller {
  public static class NativeMethods {
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessageTimeout(
      IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
      uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
  }
}
'@
        }
        $result = [UIntPtr]::Zero
        $null = [WinkitInstaller.NativeMethods]::SendMessageTimeout(
          [IntPtr]0xffff, 0x001A, [UIntPtr]::Zero, 'Environment', 0x0002, 5000, [ref]$result
        )
      }
      catch {
        Write-Warning "PATH was updated, but the environment-change broadcast failed: $($_.Exception.Message)"
      }
    }

    return $persistentChanged
  }

  function Invoke-StateWrite {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Destination,

      [Parameter(Mandatory = $true)]
      [pscustomobject]$State
    )

    $statePath = Join-Path -Path $Destination -ChildPath '.winkit-install.json'
    $json = $State | ConvertTo-Json -Depth 4
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($statePath, $json + [Environment]::NewLine, $encoding)
  }

  function Test-SafeChildPath {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Path,

      [Parameter(Mandatory = $true)]
      [string]$Parent
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullParent = [System.IO.Path]::GetFullPath($Parent).TrimEnd([char[]]@('\', '/'))
    $prefix = "$fullParent$([System.IO.Path]::DirectorySeparatorChar)"
    return $fullPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
  }

  function Test-ReparsePointInTree {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Path
    )

    $pending = New-Object 'System.Collections.Generic.Stack[string]'
    $pending.Push($Path)

    while ($pending.Count -gt 0) {
      $directory = $pending.Pop()
      foreach ($entry in [System.IO.Directory]::EnumerateFileSystemEntries($directory)) {
        $attributes = [System.IO.File]::GetAttributes($entry)
        if (($attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
          return $true
        }
        if (($attributes -band [System.IO.FileAttributes]::Directory) -ne 0) {
          $pending.Push($entry)
        }
      }
    }

    return $false
  }

  function Invoke-OwnedDirectoryCleanup {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Path,

      [Parameter(Mandatory = $true)]
      [string]$Parent
    )

    if (-not (Test-SafeChildPath -Path $Path -Parent $Parent)) {
      throw "Refusing to remove a path outside the installation parent: $Path"
    }
    if (Test-Path -LiteralPath $Path) {
      $item = Get-Item -LiteralPath $Path -Force
      if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0 -or (Test-ReparsePointInTree -Path $Path)) {
        throw "Refusing recursive cleanup because a reparse point exists below: $Path"
      }
      Remove-Item -LiteralPath $Path -Recurse -Force
    }
  }

  function Get-InstallResult {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Status,

      [Parameter(Mandatory = $true)]
      [string]$ResultVersion,

      [string]$PreviousVersion,

      [Parameter(Mandatory = $true)]
      [string]$ResultPath,

      [Parameter(Mandatory = $true)]
      [string]$ResultScope,

      [bool]$PathUpdated = $false,

      [bool]$DependencyInstalled = $false
    )

    return [pscustomobject]@{
      PSTypeName          = 'winkit.InstallationResult'
      Status              = $Status
      Version             = $ResultVersion
      PreviousVersion     = $PreviousVersion
      InstallPath         = $ResultPath
      Scope               = $ResultScope
      PathUpdated         = $PathUpdated
      DependencyInstalled = $DependencyInstalled
    }
  }

  function Invoke-WinkitInstaller {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
      [Parameter(Mandatory = $true)]
      [pscustomobject]$Context
    )

    $scopeSpecified = $Context.BoundParameters.ContainsKey('Scope')
    $pathSpecified = $Context.BoundParameters.ContainsKey('InstallPath')
    $repositorySpecified = $Context.BoundParameters.ContainsKey('Repository')
    $nonInteractiveSpecified = $Context.BoundParameters.ContainsKey('NonInteractive')

    $effectiveScope = if ($scopeSpecified) { $Context.Scope } else { 'CurrentUser' }
    if ($pathSpecified) {
      $effectiveInstallPath = Get-NormalizedPath -Path $Context.InstallPath
    }
    elseif ($effectiveScope -eq 'AllUsers') {
      $programFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
      $effectiveInstallPath = Get-NormalizedPath -Path (Join-Path -Path $programFiles -ChildPath 'winkit')
    }
    else {
      $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
      $effectiveInstallPath = Get-NormalizedPath -Path (Join-Path -Path $localAppData -ChildPath 'Programs\winkit')
    }

    $state = Get-InstallState -Path $effectiveInstallPath
    if ($state) {
      if (-not $scopeSpecified) {
        $effectiveScope = [string]$state.scope
      }
      elseif (-not ([string]$state.scope).Equals($effectiveScope, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "The managed installation uses scope '$($state.scope)', not '$effectiveScope'."
      }
    }

    if ($effectiveScope -eq 'AllUsers' -and -not (Test-Administrator)) {
      throw 'AllUsers installation requires an elevated PowerShell session.'
    }

    $effectiveRepository = $Context.Repository
    if ($state) {
      if (-not $repositorySpecified) {
        $effectiveRepository = [string]$state.repository
      }
      elseif (-not ([string]$state.repository).Equals($effectiveRepository, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "The managed installation belongs to '$($state.repository)'. Choose another InstallPath for '$effectiveRepository'."
      }
    }

    $nonInteractiveEnabled = if ($nonInteractiveSpecified) {
      [bool]$Context.NonInteractive
    }
    else {
      Get-EnvironmentBoolean -Name 'WINKIT_NON_INTERACTIVE'
    }

    $release = Get-WinkitRelease -ReleaseRepository $effectiveRepository -ReleaseVersion $Context.Version
    $previousVersion = if ($state) { [string]$state.version } else { $null }
    $layoutValid = $state -and
    (Test-Path -LiteralPath (Join-Path -Path $effectiveInstallPath -ChildPath 'scripts\Invoke-Bootstrap.ps1') -PathType Leaf) -and
    (Test-Path -LiteralPath (Join-Path -Path $effectiveInstallPath -ChildPath 'bin\bootstrap.cmd') -PathType Leaf) -and
    (Test-Path -LiteralPath (Join-Path -Path $effectiveInstallPath -ChildPath 'requirements.psd1') -PathType Leaf)
    $payloadRequired = -not $state -or $Context.Force -or $previousVersion -ne $release.Version -or -not $layoutValid

    if (-not $state) {
      $operation = 'Install'
    }
    elseif ($previousVersion -ne $release.Version) {
      $operation = 'Update'
    }
    elseif ($Context.Force) {
      $operation = 'Reinstall'
    }
    elseif (-not $layoutValid) {
      $operation = 'Repair'
    }
    else {
      $operation = 'Verify'
    }

    $description = "$operation winkit $($release.Version) at '$effectiveInstallPath'"
    if ($Context.DryRun) {
      $WhatIfPreference = $true
    }

    if ($Context.DryRun -or $Context.WhatIfPreference) {
      $null = $Context.Cmdlet.ShouldProcess($effectiveInstallPath, $description)
      $result = Get-InstallResult -Status 'Planned' -ResultVersion $release.Version -PreviousVersion $previousVersion -ResultPath $effectiveInstallPath -ResultScope $effectiveScope
      if ($Context.PassThru) { return $result }
      return
    }

    $approved = if ($nonInteractiveEnabled) {
      $true
    }
    else {
      $Context.Cmdlet.ShouldProcess($effectiveInstallPath, $description)
    }
    if (-not $approved) {
      $result = Get-InstallResult -Status 'Skipped' -ResultVersion $release.Version -PreviousVersion $previousVersion -ResultPath $effectiveInstallPath -ResultScope $effectiveScope
      if ($Context.PassThru) { return $result }
      return
    }

    $parentPath = Split-Path -Path $effectiveInstallPath -Parent
    if (-not (Test-Path -LiteralPath $parentPath -PathType Container)) {
      New-Item -Path $parentPath -ItemType Directory -Force | Out-Null
    }

    $lockPath = "$effectiveInstallPath.install-lock"
    $lockAcquired = $false
    $workPath = $null
    $stagePath = $null
    $backupPath = $null
    $activated = $false
    $committed = $false
    $dependencyInstalled = $false
    $pathUpdated = $false

    try {
      try {
        New-Item -Path $lockPath -ItemType Directory -ErrorAction Stop | Out-Null
        $lockAcquired = $true
      }
      catch {
        throw "Another installer may be using '$effectiveInstallPath' (lock: $lockPath)."
      }

      if ($payloadRequired) {
        $workPath = Join-Path -Path $parentPath -ChildPath ('.winkit-install-' + [guid]::NewGuid().ToString('N'))
        $stagePath = Join-Path -Path $workPath -ChildPath 'stage'
        New-Item -Path $stagePath -ItemType Directory -Force | Out-Null

        $archivePath = Join-Path -Path $workPath -ChildPath 'winkit.zip'
        $checksumPath = Join-Path -Path $workPath -ChildPath 'CHECKSUMS_SHA256.txt'
        Write-InstallerMessage -Message "Downloading winkit $($release.Version)..."
        Invoke-FileDownload -Uri $release.ChecksumUri -Destination $checksumPath
        Invoke-FileDownload -Uri $release.ZipUri -Destination $archivePath
        Test-ArchiveChecksum -ArchivePath $archivePath -ChecksumPath $checksumPath
        Test-ReleaseArchive -ArchivePath $archivePath -DestinationPath $stagePath
        Expand-Archive -LiteralPath $archivePath -DestinationPath $stagePath -Force

        $requirementsPath = Join-Path -Path $stagePath -ChildPath 'requirements.psd1'
      }
      else {
        $requirementsPath = Join-Path -Path $effectiveInstallPath -ChildPath 'requirements.psd1'
      }

      $requiredPSFoundation = Get-RuntimeRequirement -RequirementsPath $requirementsPath
      $dependencyInstalled = Invoke-DependencyInstall -RequiredVersion $requiredPSFoundation -InstallScope $effectiveScope -Reinstall:$Context.Force

      if ($payloadRequired) {
        $newState = [pscustomobject][ordered]@{
          schemaVersion  = 1
          repository     = $effectiveRepository
          version        = $release.Version
          scope          = $effectiveScope
          installPath    = $effectiveInstallPath
          installedAtUtc = [DateTime]::UtcNow.ToString('o')
        }
        Invoke-StateWrite -Destination $stagePath -State $newState

        if (Test-Path -LiteralPath $effectiveInstallPath) {
          $backupPath = Join-Path -Path $parentPath -ChildPath ((Split-Path -Path $effectiveInstallPath -Leaf) + '.backup-' + [guid]::NewGuid().ToString('N'))
          Move-Item -LiteralPath $effectiveInstallPath -Destination $backupPath -ErrorAction Stop
        }

        try {
          Move-Item -LiteralPath $stagePath -Destination $effectiveInstallPath -ErrorAction Stop
          $activated = $true
        }
        catch {
          if ($backupPath -and (Test-Path -LiteralPath $backupPath) -and -not (Test-Path -LiteralPath $effectiveInstallPath)) {
            Move-Item -LiteralPath $backupPath -Destination $effectiveInstallPath -ErrorAction SilentlyContinue
          }
          throw
        }
      }

      if (-not $Context.NoPath) {
        $binPath = Join-Path -Path $effectiveInstallPath -ChildPath 'bin'
        $pathUpdated = Invoke-PathUpdate -BinPath $binPath -InstallScope $effectiveScope
      }

      $committed = $true

      if ($backupPath -and (Test-Path -LiteralPath $backupPath)) {
        try {
          Invoke-OwnedDirectoryCleanup -Path $backupPath -Parent $parentPath
        }
        catch {
          Write-Warning "The update succeeded, but the previous release could not be removed from '$backupPath': $($_.Exception.Message)"
        }
      }

      $status = if (-not $state) {
        'Installed'
      }
      elseif ($previousVersion -ne $release.Version) {
        'Updated'
      }
      elseif ($Context.Force) {
        'Reinstalled'
      }
      elseif (-not $layoutValid -or $dependencyInstalled -or $pathUpdated) {
        'Repaired'
      }
      else {
        'Current'
      }

      Write-InstallerMessage -Message "winkit $($release.Version) is ready at '$effectiveInstallPath'."
      if (-not $Context.NoPath) {
        Write-InstallerMessage -Message 'Open a new terminal if the winkit commands are not immediately available.'
      }

      $result = Get-InstallResult -Status $status -ResultVersion $release.Version -PreviousVersion $previousVersion -ResultPath $effectiveInstallPath -ResultScope $effectiveScope -PathUpdated $pathUpdated -DependencyInstalled $dependencyInstalled
      if ($Context.PassThru) { return $result }
    }
    catch {
      if (-not $committed -and $activated -and (Test-Path -LiteralPath $effectiveInstallPath)) {
        Invoke-OwnedDirectoryCleanup -Path $effectiveInstallPath -Parent $parentPath
      }
      if (-not $committed -and $backupPath -and (Test-Path -LiteralPath $backupPath) -and -not (Test-Path -LiteralPath $effectiveInstallPath)) {
        Move-Item -LiteralPath $backupPath -Destination $effectiveInstallPath -ErrorAction SilentlyContinue
      }
      throw
    }
    finally {
      if ($workPath -and (Test-Path -LiteralPath $workPath)) {
        try {
          Invoke-OwnedDirectoryCleanup -Path $workPath -Parent $parentPath
        }
        catch {
          Write-Warning "Could not remove installer staging directory '$workPath': $($_.Exception.Message)"
        }
      }
      if ($lockAcquired -and (Test-Path -LiteralPath $lockPath -PathType Container)) {
        Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
      }
    }
  }

  $originalSecurityProtocol = [Net.ServicePointManager]::SecurityProtocol
  try {
    [Net.ServicePointManager]::SecurityProtocol = $originalSecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WinkitInstaller -Context $Context -Confirm:$false
  }
  finally {
    [Net.ServicePointManager]::SecurityProtocol = $originalSecurityProtocol
  }
} ([pscustomobject]@{
    BoundParameters  = $PSBoundParameters
    Cmdlet           = $PSCmdlet
    Scope            = $Scope
    InstallPath      = $InstallPath
    Repository       = $Repository
    Version          = $Version
    NoPath           = [bool]$NoPath
    Force            = [bool]$Force
    NonInteractive   = [bool]$NonInteractive
    DryRun           = [bool]$DryRun
    PassThru         = [bool]$PassThru
    WhatIfPreference = [bool]$WhatIfPreference
  })
