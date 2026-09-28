#Requires -Version 5.0

<#
.SYNOPSIS
  Installs, updates, removes, or changes the scope of a registered winkit release.

.DESCRIPTION
  Downloads a winkit release from GitHub, verifies the release ZIP against its
  published SHA-256 checksum, validates the archive layout, installs the pinned
  PSFoundation runtime dependency, and activates the release at a stable path.

  Re-running the script updates an existing managed installation. Updates are
  staged beside the installation and rolled back if activation fails. Existing
  directories without winkit's ownership manifest are never overwritten.

  The installer is intentionally standalone and does not import PSFoundation.

  Configure the installer through process environment variables, not script
  parameters. Unset or blank values reuse stored choices or defaults. Booleans accept
  1/true/yes/on or 0/false/no/off, case-insensitively; invalid values fail.

  WINKIT_SCOPE: CurrentUser or AllUsers; defaults to AllUsers when elevated.
  AllUsers installation requires elevation. Existing registrations retain scope.
  WINKIT_INSTALL_PATH: destination; defaults to %LOCALAPPDATA%\Programs\winkit
  for CurrentUser or %ProgramFiles%\winkit for AllUsers.
  WINKIT_REPOSITORY: OWNER/REPOSITORY; defaults to adnoctem/winkit. Existing
  managed installations retain their repository unless specified.
  WINKIT_VERSION: semantic release version, optionally prefixed by v; defaults
  to the latest stable GitHub release. Use latest to clear a stored version pin.
  WINKIT_NO_PATH: skip persistent and current-session PATH updates.
  WINKIT_FORCE: reinstall the release and pinned dependency. Never replaces an
  unrecognized directory or bypasses checksum and ownership validation.
  WINKIT_NON_INTERACTIVE: suppress installer confirmations; conflicts still fail.
  WINKIT_DRY_RUN: read release metadata and print sources, destinations,
  dependency installation, and PATH changes without downloading release assets
  or changing the system. Takes precedence over force and non-interactive mode.
  WINKIT_PASS_THRU: return a structured installation result.
  WINKIT_UNINSTALL: remove the selected registered installation offline.
  WINKIT_CHANGE_SCOPE: move the registered source to explicit WINKIT_SCOPE.
  Installation choices are stored in the registry; invocation flags are not.
  See dist/README.md for the registry schema, ownership checks, and recovery limits.
  All Boolean settings default to false.

.EXAMPLE
  PS> irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
  Installs or updates the latest stable winkit release in the selected scope.

.EXAMPLE
  PS> $env:WINKIT_INSTALL_PATH = 'D:\Tools\winkit'
  PS> irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
  Installs winkit at a custom location.

.EXAMPLE
  PS> $env:WINKIT_NON_INTERACTIVE = '1'; irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
  Installs or updates winkit without installer prompts.

.EXAMPLE
  PS> $env:WINKIT_DRY_RUN = '1'
  PS> irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
  Prints the plan. Remove Env:WINKIT_DRY_RUN before an actual installation.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core support: Yes.
  SYSTEM-account suitability: Set WINKIT_SCOPE=AllUsers and WINKIT_INSTALL_PATH;
  CurrentUser installs target the invoking account's profile.
#>

# Keep helper functions and installer state out of the invoking session.
& {
  Set-StrictMode -Version 2.0
  $ErrorActionPreference = 'Stop'

  if ($args.Count) {
    throw 'The installer accepts WINKIT_* environment variables, not command-line parameters.'
  }

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

  function Get-InstallerConfiguration {
    $configuration = [ordered]@{}
    $stringSettings = [ordered]@{
      Scope       = 'WINKIT_SCOPE'
      InstallPath = 'WINKIT_INSTALL_PATH'
      Repository  = 'WINKIT_REPOSITORY'
      Version     = 'WINKIT_VERSION'
    }

    foreach ($name in $stringSettings.Keys) {
      $value = [Environment]::GetEnvironmentVariable($stringSettings[$name], 'Process')
      $configuration[$name] = if ([string]::IsNullOrWhiteSpace($value)) { $null } else { $value.Trim() }
    }

    if ($configuration.Scope -and $configuration.Scope -notin @('CurrentUser', 'AllUsers')) {
      throw 'WINKIT_SCOPE must be CurrentUser or AllUsers.'
    }
    if ($configuration.Repository -and $configuration.Repository -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') {
      throw 'WINKIT_REPOSITORY must use OWNER/REPOSITORY format.'
    }
    if ($configuration.Version -and $configuration.Version -ne 'latest' -and $configuration.Version -notmatch '^v?(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(-[0-9A-Za-z.-]+)?(\+[0-9A-Za-z.-]+)?$') {
      throw 'WINKIT_VERSION must be a semantic release version, optionally prefixed by v.'
    }

    $configuration.NoPath = Get-EnvironmentBoolean -Name 'WINKIT_NO_PATH'
    $configuration.Force = Get-EnvironmentBoolean -Name 'WINKIT_FORCE'
    $configuration.NonInteractive = Get-EnvironmentBoolean -Name 'WINKIT_NON_INTERACTIVE'
    $configuration.DryRun = Get-EnvironmentBoolean -Name 'WINKIT_DRY_RUN'
    $configuration.PassThru = Get-EnvironmentBoolean -Name 'WINKIT_PASS_THRU'
    $configuration.Uninstall = Get-EnvironmentBoolean -Name 'WINKIT_UNINSTALL'
    $configuration.ChangeScope = Get-EnvironmentBoolean -Name 'WINKIT_CHANGE_SCOPE'
    $configuration.NoPathSpecified = -not [string]::IsNullOrWhiteSpace($env:WINKIT_NO_PATH)
    $configuration.VersionSpecified = -not [string]::IsNullOrWhiteSpace($env:WINKIT_VERSION)

    if ($configuration.Uninstall -and $configuration.ChangeScope) {
      throw 'WINKIT_UNINSTALL and WINKIT_CHANGE_SCOPE cannot be combined.'
    }
    if ($configuration.Uninstall -and ($configuration.Force -or $configuration.Version -or $configuration.NoPathSpecified)) {
      throw 'Uninstall does not accept WINKIT_FORCE, WINKIT_VERSION, or WINKIT_NO_PATH.'
    }
    if ($configuration.ChangeScope -and (-not $configuration.Scope -or $configuration.Version -or $configuration.Force)) {
      throw 'WINKIT_CHANGE_SCOPE requires WINKIT_SCOPE and does not accept WINKIT_VERSION or WINKIT_FORCE.'
    }

    return [pscustomobject]$configuration
  }

  function Get-NormalizedPath {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Path
    )

    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if ($expanded.StartsWith('\\') -or -not [IO.Path]::IsPathRooted($expanded)) {
      throw 'WINKIT_INSTALL_PATH must be an absolute local filesystem path.'
    }
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
      $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
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

    if ([int]$state.schemaVersion -ne 2) {
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
    $allowedRoots = @('scripts', 'bin', 'resources', 'dist', 'requirements.psd1', 'LICENSE')
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

    foreach ($requiredPath in @('scripts/Invoke-Bootstrap.ps1', 'bin/bootstrap.cmd', 'dist/install.ps1', 'dist/uninstall.ps1', 'dist/README.md', 'requirements.psd1', 'LICENSE')) {
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
      Assert-PlainInstallPath $Path
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

  function Get-RegistrationPath {
    param (
      [string]
      $Scope
    )

    $hive = if ($Scope -eq 'AllUsers') { 'HKLM:' } else { 'HKCU:' }
    [pscustomobject]@{
      Application = "$hive\Software\AdNoctem\winkit"
      Options     = "$hive\Software\AdNoctem\winkit\InstallOptions"
      Uninstall   = "$hive\Software\Microsoft\Windows\CurrentVersion\Uninstall\winkit"
    }
  }

  function Get-InstallRegistration {
    param (
      [string]
      $Scope
    )

    $paths = Get-RegistrationPath $Scope
    if (-not (Test-Path -LiteralPath $paths.Options)) {
      if (Test-Path -LiteralPath $paths.Uninstall) {
        throw "Incomplete or conflicting winkit registration in $Scope."
      }
      if (Test-Path -LiteralPath $paths.Application) {
        $root = Get-ItemProperty -LiteralPath $paths.Application -ErrorAction Stop
        if (@($root.PSObject.Properties.Name | Where-Object { $_ -in @('InstallId', 'InstalledVersion', 'SchemaVersion') }).Count) {
          throw "Incomplete winkit installation metadata in $Scope."
        }
      }
      return $null
    }

    $root = Get-ItemProperty -LiteralPath $paths.Application -ErrorAction Stop
    $options = Get-ItemProperty -LiteralPath $paths.Options -ErrorAction Stop
    $uninstall = Get-ItemProperty -LiteralPath $paths.Uninstall -ErrorAction Stop
    if ($root.SchemaVersion -ne 2 -or $root.InstallId -notmatch '^[0-9a-f-]{36}$' -or
      $options.Scope -ne $Scope -or $options.NoPath -notin @(0, 1) -or
      $uninstall.InstallId -ne $root.InstallId -or $uninstall.InstallLocation -ne $options.InstallPath) {
      throw "Invalid winkit registration in $Scope."
    }

    $path = Get-NormalizedPath $options.InstallPath
    $state = Get-InstallState $path
    if (-not $state -or $state.installId -ne $root.InstallId -or $state.scope -ne $Scope -or
      $state.repository -ne $options.Repository -or $state.version -ne $root.InstalledVersion) {
      throw "Registry and ownership manifest disagree for '$path'."
    }

    [pscustomobject]@{
      Scope            = $Scope
      InstallPath      = $path
      Repository       = [string]$options.Repository
      Version          = [string]$options.Version
      NoPath           = [bool]$options.NoPath
      InstallId        = [string]$root.InstallId
      InstalledVersion = [string]$root.InstalledVersion
      State            = $state
    }
  }

  function Write-RegistryValueSet {
    param (
      [string]
      $Path,

      [System.Collections.IDictionary]
      $Values
    )

    if (-not (Test-Path -LiteralPath $Path)) {
      $null = New-Item -Path $Path -Force -ErrorAction Stop
    }
    foreach ($name in $Values.Keys) {
      $type = if ($Values[$name] -is [int]) { 'DWord' } else { 'String' }
      $null = New-ItemProperty -LiteralPath $Path -Name $name -Value $Values[$name] -PropertyType $type -Force -ErrorAction Stop
    }
  }

  function Set-InstallRegistration {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private helper; the maintenance workflow owns dry-run and explicit confirmation, including rollback.')]
    [CmdletBinding()]
    param (
      [pscustomobject]
      $Record
    )

    $paths = Get-RegistrationPath $Record.Scope
    Write-RegistryValueSet $paths.Application ([ordered]@{
        SchemaVersion    = 2
        InstallId        = $Record.InstallId
        InstalledVersion = $Record.InstalledVersion
      })
    Write-RegistryValueSet $paths.Options ([ordered]@{
        Scope       = $Record.Scope
        InstallPath = $Record.InstallPath
        Repository  = $Record.Repository
        Version     = [string]$Record.Version
        NoPath      = [int]$Record.NoPath
      })

    $hostPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $uninstaller = Join-Path $Record.InstallPath 'dist\uninstall.ps1'
    Write-RegistryValueSet $paths.Uninstall ([ordered]@{
        DisplayName     = 'winkit'
        DisplayVersion  = $Record.InstalledVersion
        Publisher       = 'AdNoctem'
        InstallLocation = $Record.InstallPath
        InstallDate     = ([datetime]$Record.State.installedAtUtc).ToString('yyyyMMdd')
        UninstallString = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $hostPath, $uninstaller
        URLInfoAbout    = "https://github.com/$($Record.Repository)"
        NoModify        = 1
        NoRepair        = 1
        InstallId       = $Record.InstallId
      })
  }

  function Remove-InstallRegistration {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private helper; the maintenance workflow owns dry-run and explicit confirmation, including rollback.')]
    [CmdletBinding()]
    param (
      [pscustomobject]
      $Record
    )

    $paths = Get-RegistrationPath $Record.Scope
    # Paths come exclusively from the two fixed registry locations. Never use a
    # filesystem path or a registry command supplied by the ownership manifest.
    foreach ($path in @($paths.Application, $paths.Uninstall)) {
      if (Test-Path -LiteralPath $path) {
        $values = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
        if ($values.PSObject.Properties.Name -contains 'InstallId' -and $values.InstallId -ne $Record.InstallId) {
          throw "Registration ownership changed at '$path'."
        }
      }
    }

    foreach ($path in @($paths.Uninstall, $paths.Options)) {
      if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
      }
    }
    if (Test-Path -LiteralPath $paths.Application) {
      $values = Get-ItemProperty -LiteralPath $paths.Application -ErrorAction Stop
      foreach ($name in @('SchemaVersion', 'InstallId', 'InstalledVersion')) {
        if ($values.PSObject.Properties.Name -contains $name) {
          Remove-ItemProperty -LiteralPath $paths.Application -Name $name -ErrorAction Stop
        }
      }
      # Leave unrelated/future application settings intact.
      $key = Get-Item -LiteralPath $paths.Application
      if (-not $key.SubKeyCount -and -not $key.ValueCount) {
        Remove-Item -LiteralPath $paths.Application -Force -ErrorAction Stop
      }
    }
  }

  function Assert-PlainInstallPath {
    param (
      [string]
      $Path
    )

    $current = Get-NormalizedPath $Path
    while ($current) {
      if (Test-Path -LiteralPath $current) {
        $item = Get-Item -LiteralPath $current -Force
        if (-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
          throw "Installation path contains a non-directory or reparse point: $current"
        }
      }
      $current = Split-Path $current -Parent
    }
  }

  function Get-OwnedFileInventory {
    param (
      [string]
      $Path
    )

    Assert-PlainInstallPath $Path
    if (Test-ReparsePointInTree $Path) {
      throw "Refusing a reparse point within '$Path'."
    }
    $prefix = $Path.TrimEnd('\') + '\'
    foreach ($file in (Get-ChildItem -LiteralPath $Path -Recurse -Force -File | Sort-Object FullName)) {
      $relative = $file.FullName.Substring($prefix.Length)
      if ($relative -ne '.winkit-install.json') {
        [pscustomobject]@{
          Path   = $relative
          Sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
        }
      }
    }
  }

  function Assert-OwnedInstallation {
    param (
      [pscustomobject]
      $Record
    )

    $state = Get-InstallState $Record.InstallPath
    if (-not $state -or $state.installId -ne $Record.InstallId -or $state.scope -ne $Record.Scope) {
      throw 'Installation ownership changed.'
    }
    $actual = @(Get-OwnedFileInventory $Record.InstallPath)
    $expected = @($state.files)
    if (-not $expected.Count -or $expected.Count -ne $actual.Count) {
      throw 'Installation has missing or additional files. Preserve your files and restore the managed layout before proceeding.'
    }
    $index = @{}
    foreach ($file in $expected) {
      if ($index.ContainsKey($file.Path)) {
        throw 'Ownership manifest contains duplicate file paths.'
      }
      $index[$file.Path] = $file.Sha256
    }
    foreach ($file in $actual) {
      if (-not $index.ContainsKey($file.Path) -or $index[$file.Path] -ne $file.Sha256) {
        throw "Added or modified file blocks replacement/removal: $($file.Path)"
      }
    }
  }

  function Get-InstallerPathValue {
    param (
      [string]
      $Target
    )

    [Environment]::GetEnvironmentVariable('Path', $Target)
  }

  function Set-InstallerPathValue {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private helper; the maintenance workflow owns dry-run and explicit confirmation, including rollback.')]
    [CmdletBinding()]
    param (
      [string]
      $Target,

      [AllowNull()]
      [string]
      $Value
    )

    [Environment]::SetEnvironmentVariable('Path', $Value, $Target)
  }

  function Update-InstalledPath {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private helper; the maintenance workflow owns dry-run and explicit confirmation, including rollback.')]
    [CmdletBinding()]
    param (
      [pscustomobject]
      $Record,

      [switch]
      $Remove
    )

    if ($Record.NoPath -and -not $Remove) {
      return
    }
    $bin = Join-Path $Record.InstallPath 'bin'
    $target = if ($Record.Scope -eq 'AllUsers') { 'Machine' } else { 'User' }
    foreach ($environmentTarget in @($target, 'Process')) {
      $old = Get-InstallerPathValue $environmentTarget
      $parts = @($old -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
      $matching = @($parts | Where-Object { $_.Trim().TrimEnd('\', '/') -eq $bin })
      if (($Remove -and -not $matching.Count) -or (-not $Remove -and $matching.Count)) {
        continue
      }
      if ($Remove) {
        $parts = @($parts | Where-Object { $_.Trim().TrimEnd('\', '/') -ne $bin })
      }
      elseif (-not $matching.Count -and -not $Record.NoPath) {
        $parts += $bin
      }
      $new = $parts -join ';'
      if ($new -ne $old) {
        Set-InstallerPathValue -Target $environmentTarget -Value $new
      }
    }
  }

  function Enter-MaintenanceLock {
    $mutex = New-Object Threading.Mutex($false, 'Global\AdNoctem.winkit.Maintenance')
    try {
      $acquired = $false
      try {
        $acquired = $mutex.WaitOne(0)
      }
      catch [Threading.AbandonedMutexException] {
        $acquired = $true
      }
      if (-not $acquired) {
        throw 'Another winkit maintenance operation is running.'
      }
      return $mutex
    }
    catch {
      $mutex.Dispose()
      throw
    }
  }

  function Confirm-Maintenance {
    param (
      [string]
      $Description,

      [bool]
      $NonInteractive
    )

    if ($NonInteractive) {
      return $true
    }
    Write-InstallerMessage $Description
    $answer = Read-Host 'Type YES to continue (default: No)'
    return $answer -ceq 'YES'
  }

  function Resolve-InstallPlan {
    param (
      [pscustomobject]
      $Context
    )

    $records = @()
    foreach ($scope in @('CurrentUser', 'AllUsers')) {
      $record = Get-InstallRegistration $scope
      if ($record) {
        $records += $record
      }
    }
    $selected = $null
    $source = $null
    if ($Context.ChangeScope) {
      $source = @($records | Where-Object Scope -NE $Context.Scope)
      if ($source.Count -ne 1 -or @($records | Where-Object Scope -EQ $Context.Scope).Count) {
        throw 'Scope change requires one source registration and an unused destination scope.'
      }
      $source = $source[0]
    }
    else {
      $candidates = @($records)
      if ($Context.InstallPath) {
        $requested = Get-NormalizedPath $Context.InstallPath
        $candidates = @($records | Where-Object InstallPath -EQ $requested)
        if (-not $candidates.Count -and ($Context.Uninstall -or $records.Count)) {
          throw 'WINKIT_INSTALL_PATH does not match a registered installation. No fallback target was selected.'
        }
      }
      if ($Context.Scope) {
        $candidates = @($candidates | Where-Object Scope -EQ $Context.Scope)
      }
      if ($candidates.Count -gt 1) {
        throw 'Multiple installations match. Set WINKIT_SCOPE or WINKIT_INSTALL_PATH.'
      }
      if ($candidates.Count) {
        $selected = $candidates[0]
      }
      elseif ($records.Count -or $Context.Uninstall) {
        throw 'No registered installation matches. A scope mismatch requires WINKIT_CHANGE_SCOPE=1.'
      }
    }

    $scope = $Context.Scope
    if ($selected) {
      $scope = $selected.Scope
    }
    elseif (-not $scope) {
      $scope = if (Test-Administrator) { 'AllUsers' } else { 'CurrentUser' }
      if ($scope -eq 'AllUsers') {
        Write-InstallerMessage 'Elevated session detected; defaulting to AllUsers. Set WINKIT_SCOPE=CurrentUser to install for this account.'
      }
    }

    $path = $Context.InstallPath
    if ($selected) {
      $path = $selected.InstallPath
    }
    elseif (-not $path) {
      $path = if ($scope -eq 'AllUsers') {
        Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'winkit'
      }
      else {
        Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\winkit'
      }
    }
    $path = Get-NormalizedPath $path
    Assert-PlainInstallPath $path
    if (-not $selected -and (Test-Path -LiteralPath $path)) {
      throw "Refusing to overwrite an unrecognized directory or occupied destination: $path"
    }

    $baseline = if ($source) { $source } else { $selected }
    if ($source -and ((Test-SafeChildPath $path $source.InstallPath) -or (Test-SafeChildPath $source.InstallPath $path))) {
      throw 'Source and destination installation paths must not overlap.'
    }
    $repository = if ($Context.Repository) { $Context.Repository } elseif ($baseline) { $baseline.Repository } else { 'adnoctem/winkit' }
    if ($baseline -and $repository -ne $baseline.Repository) {
      throw 'The registered installation belongs to a different repository.'
    }
    $version = if ($Context.VersionSpecified) { $Context.Version } elseif ($baseline) { $baseline.Version } else { '' }
    if ($version -eq 'latest') {
      $version = ''
    }
    $noPath = if ($Context.NoPathSpecified) { $Context.NoPath } elseif ($baseline) { $baseline.NoPath } else { $false }
    $dryRun = $Context.DryRun -or $WhatIfPreference
    if (($scope -eq 'AllUsers' -or $source) -and -not $dryRun -and -not (Test-Administrator)) {
      throw 'AllUsers installation and scope changes require an elevated 64-bit PowerShell session.'
    }
    if ($baseline) {
      Assert-OwnedInstallation $baseline
    }

    [pscustomobject]@{
      Scope       = $scope
      InstallPath = $path
      Repository  = $repository
      Version     = $version
      NoPath      = [bool]$noPath
      Existing    = $selected
      Source      = $source
      DryRun      = [bool]$dryRun
    }
  }

  function Invoke-WinkitInstaller {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param (
      [Parameter(Mandatory = $true)]
      [pscustomobject]
      $Context
    )

    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
      throw 'Use 64-bit PowerShell so registry registration uses the native registry view.'
    }
    $plan = Resolve-InstallPlan $Context
    $previous = if ($plan.Source) { $plan.Source } else { $plan.Existing }
    $release = $null
    if ($Context.Uninstall) {
      $operation = 'Uninstall'
      $version = $previous.InstalledVersion
    }
    elseif ($plan.Source) {
      $operation = 'ChangeScope'
      $version = $previous.InstalledVersion
    }
    else {
      $release = Get-WinkitRelease -ReleaseRepository $plan.Repository -ReleaseVersion $plan.Version
      $version = $release.Version
      $operation = if (-not $previous) { 'Install' } elseif ($previous.InstalledVersion -ne $version) { 'Update' } elseif ($Context.Force) { 'Reinstall' } else { 'Verify' }
    }
    $payloadRequired = $operation -in @('Install', 'Update', 'Reinstall', 'ChangeScope')
    $description = "$operation winkit $version at '$($plan.InstallPath)' ($($plan.Scope))"
    $previousVersion = if ($previous) { $previous.InstalledVersion } else { $null }
    $resultArgs = @{
      ResultVersion   = $version
      PreviousVersion = $previousVersion
      ResultPath      = $plan.InstallPath
      ResultScope     = $plan.Scope
    }
    $parent = Split-Path $plan.InstallPath -Parent

    if ($plan.DryRun) {
      Write-InstallerMessage "DRY RUN: $description"
      if ($Context.Uninstall) {
        Write-InstallerMessage 'Would remove the verified owned files, matching PATH entry, and winkit installation/uninstall registration. Shared dependencies remain installed.'
      }
      else {
        Write-InstallerMessage "Scripts: $(Join-Path $plan.InstallPath 'scripts')"
        Write-InstallerMessage "Launchers: $(Join-Path $plan.InstallPath 'bin')"
        Write-InstallerMessage "Resources: $(Join-Path $plan.InstallPath 'resources')"
        if ($plan.Source) {
          Write-InstallerMessage "Would copy verified files from '$($plan.Source.InstallPath)', then retire that $($plan.Source.Scope) installation after activation."
        }
        elseif ($payloadRequired) {
          Write-InstallerMessage "Would download archive: $($release.ZipUri)"
          Write-InstallerMessage "Would download checksums: $($release.ChecksumUri)"
        }
        Write-InstallerMessage "Would use temporary staging below '$parent', validate files, and register the installation."
        Write-InstallerMessage "Would ensure the pinned PSFoundation dependency from PSGallery for $($plan.Scope); NuGet is installed only if needed."
        if ($plan.NoPath) {
          Write-InstallerMessage 'PATH update disabled by WINKIT_NO_PATH.'
        }
        else {
          $pathTarget = if ($plan.Scope -eq 'AllUsers') { 'machine' } else { 'user' }
          Write-InstallerMessage "Would ensure '$(Join-Path $plan.InstallPath 'bin')' is on the $pathTarget PATH and current session PATH."
        }
      }
      $registryPaths = Get-RegistrationPath $plan.Scope
      Write-InstallerMessage "Registration: $($registryPaths.Application); $($registryPaths.Uninstall)"
      if ($plan.Scope -eq 'AllUsers' -or $plan.Source) {
        Write-InstallerMessage 'Actual AllUsers installation or scope change requires an elevated PowerShell session.'
      }
      Write-InstallerMessage 'Preview only: no files, registry values, dependencies, or PATH settings were changed.'
      if ($Context.PassThru) {
        return Get-InstallResult -Status Planned @resultArgs
      }
      return
    }

    # Removal and scope changes require explicit consent independently of the
    # host's ConfirmPreference. Ordinary installation retains ShouldProcess.
    $approved = if ($operation -in @('Uninstall', 'ChangeScope')) {
      Confirm-Maintenance -Description $description -NonInteractive $Context.NonInteractive
    }
    elseif ($Context.NonInteractive) {
      $true
    }
    else {
      $PSCmdlet.ShouldProcess($plan.InstallPath, $description)
    }
    if (-not $approved) {
      if ($Context.PassThru) {
        return Get-InstallResult -Status Skipped @resultArgs
      }
      return
    }

    # Confirmation belongs to this workflow, not individual filesystem/registry
    # cmdlets or rollback steps. Non-interactive mode must not inherit prompts.
    $ConfirmPreference = 'None'
    $lock = Enter-MaintenanceLock
    $work = $null
    $backup = $null
    $sourceBackup = $null
    $activated = $false
    $metadataTouched = $false
    $committed = $false
    $dependencyInstalled = $false
    $newRecord = $null
    $pathSnapshot = @{}
    $cleanupPending = @()
    try {
      $fresh = Resolve-InstallPlan $Context
      if (($fresh | ConvertTo-Json -Depth 8 -Compress) -ne ($plan | ConvertTo-Json -Depth 8 -Compress)) {
        throw 'Installation state changed after planning. Rerun to review the new plan.'
      }
      $pathTargets = @('Process')
      $pathTargets += if ($plan.Scope -eq 'AllUsers') { 'Machine' } else { 'User' }
      if ($plan.Source) {
        $pathTargets += if ($plan.Source.Scope -eq 'AllUsers') { 'Machine' } else { 'User' }
      }
      foreach ($target in ($pathTargets | Select-Object -Unique)) {
        $pathSnapshot[$target] = Get-InstallerPathValue $target
      }

      if ($Context.Uninstall) {
        $backup = Join-Path $parent ('.winkit-remove-' + [guid]::NewGuid().ToString('N'))
        Assert-PlainInstallPath $parent
        Move-Item -LiteralPath $plan.InstallPath -Destination $backup -ErrorAction Stop
        $metadataTouched = $true
        Update-InstalledPath -Record $previous -Remove
        Remove-InstallRegistration $previous
      }
      else {
        if ($payloadRequired) {
          $null = New-Item -Path $parent -ItemType Directory -Force
          Assert-PlainInstallPath $parent
          $work = Join-Path $parent ('.winkit-install-' + [guid]::NewGuid().ToString('N'))
          $stage = Join-Path $work 'stage'
          $null = New-Item -Path $stage -ItemType Directory
          if ($plan.Source) {
            foreach ($item in (Get-ChildItem -LiteralPath $plan.Source.InstallPath -Force)) {
              Copy-Item -LiteralPath $item.FullName -Destination $stage -Recurse -Force -ErrorAction Stop
            }
            $copied = @(Get-OwnedFileInventory $stage)
            if (($copied | ConvertTo-Json -Compress) -ne ($plan.Source.State.files | ConvertTo-Json -Compress)) {
              throw 'Staged files do not match the source ownership manifest.'
            }
          }
          else {
            $archive = Join-Path $work 'winkit.zip'
            $checksum = Join-Path $work 'CHECKSUMS_SHA256.txt'
            Invoke-FileDownload -Uri $release.ChecksumUri -Destination $checksum
            Invoke-FileDownload -Uri $release.ZipUri -Destination $archive
            Test-ArchiveChecksum -ArchivePath $archive -ChecksumPath $checksum
            Test-ReleaseArchive -ArchivePath $archive -DestinationPath $stage
            Expand-Archive -LiteralPath $archive -DestinationPath $stage -Force
          }
          $requirementsPath = Join-Path $stage 'requirements.psd1'
        }
        else {
          $requirementsPath = Join-Path $plan.InstallPath 'requirements.psd1'
        }

        $requiredVersion = Get-RuntimeRequirement $requirementsPath
        $dependencyInstalled = Invoke-DependencyInstall -RequiredVersion $requiredVersion -InstallScope $plan.Scope -Reinstall:$Context.Force
        $installId = if ($previous) { $previous.InstallId } else { [guid]::NewGuid().ToString() }
        $state = if ($payloadRequired) {
          [pscustomobject][ordered]@{
            schemaVersion  = 2
            installId      = $installId
            repository     = $plan.Repository
            version        = $version
            scope          = $plan.Scope
            installPath    = $plan.InstallPath
            installedAtUtc = [DateTime]::UtcNow.ToString('o')
            files          = @(Get-OwnedFileInventory $stage)
          }
        }
        else {
          $previous.State
        }
        $newRecord = [pscustomobject]@{
          Scope            = $plan.Scope
          InstallPath      = $plan.InstallPath
          Repository       = $plan.Repository
          Version          = [string]$plan.Version
          NoPath           = $plan.NoPath
          InstallId        = $installId
          InstalledVersion = $version
          State            = $state
        }

        if ($payloadRequired) {
          Invoke-StateWrite -Destination $stage -State $state
          if ($plan.Existing) {
            Assert-OwnedInstallation $plan.Existing
            $backup = Join-Path $parent ('.winkit-backup-' + [guid]::NewGuid().ToString('N'))
            Move-Item -LiteralPath $plan.InstallPath -Destination $backup -ErrorAction Stop
          }
          Move-Item -LiteralPath $stage -Destination $plan.InstallPath -ErrorAction Stop
          $activated = $true
          Assert-OwnedInstallation $newRecord
        }
        if ($plan.Source) {
          Assert-OwnedInstallation $plan.Source
          $sourceParent = Split-Path $plan.Source.InstallPath -Parent
          $sourceBackup = Join-Path $sourceParent ('.winkit-scope-' + [guid]::NewGuid().ToString('N'))
          Move-Item -LiteralPath $plan.Source.InstallPath -Destination $sourceBackup -ErrorAction Stop
        }

        $metadataTouched = $true
        Set-InstallRegistration $newRecord
        Update-InstalledPath $newRecord
        if ($plan.Source) {
          Update-InstalledPath -Record $plan.Source -Remove
          Remove-InstallRegistration $plan.Source
        }
      }
      $committed = $true
    }
    catch {
      $failure = $_
      $rollbackErrors = @()
      if ($activated) {
        try {
          Assert-OwnedInstallation $newRecord
          Invoke-OwnedDirectoryCleanup -Path $plan.InstallPath -Parent $parent
        }
        catch {
          $rollbackErrors += $_.Exception.Message
        }
      }
      foreach ($restore in @(
          @{ Backup = $backup; Destination = $plan.InstallPath }
          @{ Backup = $sourceBackup; Destination = $(if ($plan.Source) { $plan.Source.InstallPath } else { $null }) }
        )) {
        if ($restore.Backup -and (Test-Path -LiteralPath $restore.Backup)) {
          try {
            if (Test-Path -LiteralPath $restore.Destination) {
              throw "Rollback destination is occupied; backup retained at '$($restore.Backup)'."
            }
            Move-Item -LiteralPath $restore.Backup -Destination $restore.Destination -ErrorAction Stop
          }
          catch {
            $rollbackErrors += $_.Exception.Message
          }
        }
      }
      if ($metadataTouched) {
        if ($newRecord -and -not $plan.Existing) {
          try {
            Remove-InstallRegistration $newRecord
          }
          catch {
            $rollbackErrors += $_.Exception.Message
          }
        }
        foreach ($record in @($plan.Existing, $plan.Source)) {
          if ($record) {
            try {
              Set-InstallRegistration $record
            }
            catch {
              $rollbackErrors += $_.Exception.Message
            }
          }
        }
        foreach ($target in $pathSnapshot.Keys) {
          try {
            Set-InstallerPathValue -Target $target -Value $pathSnapshot[$target]
          }
          catch {
            $rollbackErrors += $_.Exception.Message
          }
        }
      }
      if ($rollbackErrors.Count) {
        throw "Maintenance failed: $($failure.Exception.Message) Rollback needs attention: $($rollbackErrors -join '; '). Retained paths: $backup $sourceBackup"
      }
      throw $failure
    }
    finally {
      $cleanup = @($work)
      if ($committed) {
        $cleanup += @($backup, $sourceBackup)
      }
      foreach ($path in ($cleanup | Where-Object { $_ })) {
        if (Test-Path -LiteralPath $path) {
          try {
            Invoke-OwnedDirectoryCleanup -Path $path -Parent (Split-Path $path -Parent)
          }
          catch {
            $cleanupPending += $path
            Write-Warning "Cleanup remains at '$path': $($_.Exception.Message)"
          }
        }
      }
      $lock.ReleaseMutex()
      $lock.Dispose()
    }

    $status = switch ($operation) {
      'Uninstall' { 'Uninstalled' }
      'ChangeScope' { 'ScopeChanged' }
      'Install' { 'Installed' }
      'Update' { 'Updated' }
      'Reinstall' { 'Reinstalled' }
      default { 'Current' }
    }
    if ($cleanupPending.Count) {
      $status = 'CleanupPending'
    }
    $pathChanged = $false
    foreach ($target in $pathSnapshot.Keys) {
      if ((Get-InstallerPathValue $target) -ne $pathSnapshot[$target]) { $pathChanged = $true }
    }
    $result = Get-InstallResult -Status $status -DependencyInstalled $dependencyInstalled -PathUpdated $pathChanged @resultArgs
    $result | Add-Member -NotePropertyName CleanupPaths -NotePropertyValue @($cleanupPending)
    Write-InstallerMessage "$status`: $description"
    if ($pathChanged) {
      Write-InstallerMessage 'PATH changes apply to this session and newly opened terminals.'
    }
    if ($Context.PassThru) {
      return $result
    }
  }

  $originalSecurityProtocol = [Net.ServicePointManager]::SecurityProtocol
  try {
    [Net.ServicePointManager]::SecurityProtocol = $originalSecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WinkitInstaller -Context (Get-InstallerConfiguration)
  }
  finally {
    [Net.ServicePointManager]::SecurityProtocol = $originalSecurityProtocol
  }
} @args
