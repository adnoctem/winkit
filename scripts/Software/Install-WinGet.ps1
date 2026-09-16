#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Installs the Windows Package Manager (winget) and its dependencies, or
  repairs a broken winget installation.

.DESCRIPTION
  Two modes:

  Install mode (default):
   1. Repair-WinGetPackageManager (preferred). Installs the
      Microsoft.WinGet.Client module from PSGallery and uses the
      official repair cmdlet to bootstrap winget and all current
      dependencies natively.
   2. Manual asset download (fallback). Pulls the App Installer
      msixbundle, framework dependency packages, and license from a
      GitHub release, then provisions them with DISM. Used when
      PSGallery is unreachable or the module path fails.

  Repair mode (-Repair):
   Runs only the repair workflow instead of an installation:
    1. Diagnose the installation: winget binary presence and version,
       winget source health (winget source list), and registration state
       of the Microsoft.Winget.Source package.
    2. If the installation is healthy, report that and exit without
       making any changes.
    3. If it is broken, present the findings and ask for explicit
       confirmation (skip the prompt with -Force) before running the
       auto-repair, which consists of:
       a. Repair-WinGetPackageManager -AllUsers -Latest -Force (the
          official Microsoft repair cmdlet; reinstalls winget and all of
          its dependencies),
       b. a reinstall of the Microsoft.Winget.Source package from
          https://cdn.winget.microsoft.com/cache/source.msix - the fix
          for error 0x8a15000f "Data required by the source is missing"
          documented in winget-cli issue #4799 (see .LINK). A corrupt or
          volume-lost source package makes winget fail with exactly this
          error on search/upgrade, and simply re-adding the source
          package resolves it,
       c. winget source reset --force to restore the default sources and
          clear any remaining broken source state.
    4. Re-verify the installation and report the final health.

  The script gates on OS compatibility up front: Windows 10 1809 (build
  17763) or Windows Server 2019+ is required, with a clear message instead
  of a deep install-path failure on unsupported builds.

  On Windows Server Core (no AppX subsystem) the manual fallback extracts
  winget as a portable executable instead of provisioning AppX packages.
  This path is in beta - it installs successfully but may not function
  properly due to missing dependencies (per the asheroto/winget-install
  reference, see .LINK).

  When run as the SYSTEM account (scheduled task / SCCM / RMM), detection
  and installation adapt accordingly (provisioned-package registration
  instead of per-user registration). The SYSTEM account is not officially
  supported by winget and may not work - carried from the reference's
  documented caveat.

  Known winget/DISM/AppX error codes are translated to actionable guidance
  instead of raw exception text; codes that indicate "already installed"
  are treated as benign.

  Requires administrator elevation. Skips if winget is already
  installed (install mode only).

.PARAMETER Version
  Specific winget release version for the manual install fallback. The
  preferred module path always installs the current stable regardless of
  this value. Not used in repair mode.

.PARAMETER ForceManual
  Skip the Repair-WinGetPackageManager path and use manual asset download
  directly. Not valid together with -Repair.

.PARAMETER Repair
  Run only the repair workflow (diagnose, confirm, auto-repair, re-verify)
  instead of attempting an installation.

.PARAMETER Force
  In repair mode, skip the explicit confirmation prompt for the auto-repair.

.PARAMETER DryRun
  Preview the steps without executing them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Install-WinGet.ps1 -Repair
  Diagnoses the winget installation and, if broken, prompts for
  confirmation before running the auto-repair.

.EXAMPLE
  PS> ./Install-WinGet.ps1 -Repair -Force
  Runs the auto-repair without prompting.

.EXAMPLE
  PS> ./Install-WinGet.ps1 -Repair -DryRun
  Previews what the auto-repair would do.

.EXAMPLE
  PS> ./Install-WinGet.ps1 -ForceManual -DryRun

.LINK
  https://github.com/adnoctem/winkit
  https://github.com/microsoft/winget-cli/issues/4799
  https://github.com/asheroto/winget-install

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT

  Caveats carried from the asheroto/winget-install reference (feature-diff in
  secrets/asheroto-winget-install.md):
  - Server Core portable installation is in beta: it installs successfully
    but may not function properly due to missing dependencies.
  - The SYSTEM account is not officially supported by winget and may not
    work; the script adapts (provisioned registration) but carries this as a
    documented risk.
  Deliberately not implemented (rejected in the reference evaluation):
  - conhost relaunch for the "resources in use" problem (error translation
    gives an actionable message instead),
  - self-update plumbing (winkit scripts are repo-versioned),
  - automatic system PATH/ACL fixing (the script reports what is needed
    instead of mutating system state).
  -GHtoken support is deferred until GitHub API rate-limiting is observed.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string]
  $Version,

  [Parameter(Mandatory = $false)]
  [switch]
  $ForceManual,

  [Parameter(Mandatory = $false)]
  [switch]
  $Repair,

  [Parameter(Mandatory = $false)]
  [switch]
  $Force,

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
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$ProgressPreference = 'SilentlyContinue'

$RunAsSystem = ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18')

# ---- OS compatibility gate ---------------------------------------------------
# winget requires Windows 10 1809 / Windows Server 2019 (build 17763) or newer.
$minSupportedBuild = 17763
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
$isServerOs = ($null -ne $osInfo -and [int]$osInfo.ProductType -ge 2)
if ($osInfo -and [int]$osInfo.OperatingSystemSKU -ge 112 -and [int]$osInfo.OperatingSystemSKU -le 115) {
  # Server multi-session editions (Azure Virtual Desktop) report as Server but
  # behave like a workstation for compatibility purposes.
  $isServerOs = $false
}
if ((Get-OSBuildNumber) -lt $minSupportedBuild) {
  $osVersion = Get-OSVersionInfo
  $osLabel = if ($isServerOs) { 'Windows Server' } else { 'Windows' }
  Write-Log -Message "Unsupported OS: this script requires $osLabel build $minSupportedBuild or newer (current: $($osVersion.DisplayVersion), build $($osVersion.CurrentBuild))." -Color Red
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Validate' -Status 'Failed' -Detail "Unsupported OS - build $($osVersion.CurrentBuild), minimum $minSupportedBuild."
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

# ---- Helpers ----------------------------------------------------------------

function Get-WingetInstallError {
  <#
    Translates known winget/DISM/AppX error codes to actionable guidance.
    Codes marked Benign indicate "already installed" states that should not
    be reported as failures.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  $translationTable = @(
    @{ Code = '0x80073D06'; Benign = $true; Detail = 'A newer version is already installed.' }
    @{ Code = '0x80073CF0'; Benign = $true; Detail = 'The same version is already installed.' }
    @{ Code = '0x80073D02'; Benign = $false; Detail = 'Resources are in use (commonly Windows Terminal holding a lock). Close Windows Terminal and retry.' }
    @{ Code = '0x80073CF3'; Benign = $false; Detail = 'A prerequisite was not detected. Retry - this is usually transient.' }
    @{ Code = '0x80073CF9'; Benign = $false; Detail = 'Registration failed under the SYSTEM account. Use an Administrator account instead.' }
  )

  foreach ($entry in $translationTable) {
    if ($Message -match $entry.Code) {
      return [pscustomobject]@{ Known = $true; Benign = $entry.Benign; Detail = $entry.Detail }
    }
  }

  [pscustomobject]@{ Known = $false; Benign = $false; Detail = $null }
}

function Find-WinGet {
  <#
    Resolves winget.exe under Program Files\WindowsApps (used under the SYSTEM
    account, where Get-Command does not see the execution alias).
  #>
  [CmdletBinding()]
  param()

  try {
    $searchPath = Join-Path -Path $env:ProgramFiles -ChildPath 'WindowsApps\Microsoft.DesktopAppInstaller_*_*__8wekyb3d8bbwe'
    $resolved = @(Resolve-Path -Path $searchPath -ErrorAction Stop |
        Sort-Object -Property { [version](($_.Path -replace '^.*?Microsoft\.DesktopAppInstaller_([\d\.]+)_.*$', '$1')) } |
        Select-Object -Last 1)
    if ($resolved.Count -gt 0) {
      $winget = Join-Path -Path $resolved[0].Path -ChildPath 'winget.exe'
      if (Test-Path -LiteralPath $winget) {
        return $winget
      }
    }
  }
  catch {
    return $null
  }

  $null
}

function Test-WinGetDetected {
  <#
    winget detection that adapts to the SYSTEM account (execution alias is not
    visible there; resolve the WindowsApps path instead).
  #>
  [CmdletBinding()]
  param()

  if ($RunAsSystem) {
    return ($null -ne (Find-WinGet))
  }

  ($null -ne (Get-Command winget -ErrorAction SilentlyContinue))
}

function Test-ServerCore {
  <#
    Detects Windows Server Core via the InstallationType registry value.
  #>
  [CmdletBinding()]
  param()

  $installationType = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name InstallationType -ErrorAction SilentlyContinue).InstallationType
  ($installationType -eq 'Server Core')
}

function Get-AppxManifestInfo {
  <#
    Reads the Identity Name/Version from AppxManifest.xml inside an appx/msix
    package without extracting it.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$PackagePath
  )

  try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
      $entry = $archive.Entries | Where-Object { $_.FullName -eq 'AppxManifest.xml' } | Select-Object -First 1
      if (-not $entry) {
        return $null
      }
      $reader = New-Object System.IO.StreamReader($entry.Open())
      try {
        $xml = $reader.ReadToEnd()
      }
      finally {
        $reader.Dispose()
      }
      $manifest = [xml]$xml
      $identity = $manifest.Package.Identity
      if (-not $identity) {
        return $null
      }
      return [pscustomobject]@{ Name = $identity.Name; Version = $identity.Version }
    }
    finally {
      $archive.Dispose()
    }
  }
  catch {
    return $null
  }
}

function Expand-WingetPortable {
  <#
    Extracts winget as a portable executable from the App Installer msixbundle
    (the Server Core path, which lacks the AppX subsystem for normal
    registration). Beta: may not function fully due to missing dependencies.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    [string]$BundlePath,

    [Parameter(Mandatory = $true)]
    [string]$Architecture,

    [Parameter(Mandatory = $true)]
    [string]$Destination
  )

  if (-not $PSCmdlet.ShouldProcess($Destination, 'Extract portable winget from the msixbundle')) {
    return $false
  }

  try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $bundle = [System.IO.Compression.ZipFile]::OpenRead($BundlePath)
    try {
      $archTag = switch ($Architecture) {
        'arm64' { 'arm64' }
        'x86' { 'x86' }
        default { 'x64' }
      }
      $inner = $bundle.Entries | Where-Object { $_.FullName -match '\.msix$' -and $_.FullName -match $archTag } | Select-Object -First 1
      if (-not $inner) {
        $inner = $bundle.Entries | Where-Object { $_.FullName -match '\.msix$' } | Select-Object -First 1
      }
      if (-not $inner) {
        throw 'No inner MSIX found in the bundle.'
      }

      $stream = $inner.Open()
      $innerZip = New-Object System.IO.Compression.ZipArchive($stream)
      try {
        foreach ($entry in $innerZip.Entries) {
          if ($entry.FullName.EndsWith('/')) {
            continue
          }
          $target = Join-Path -Path $Destination -ChildPath $entry.FullName
          $folder = Split-Path -Path $target -Parent
          $null = New-Item -ItemType Directory -Force -Path $folder -ErrorAction SilentlyContinue
          $in = $entry.Open()
          $out = [System.IO.File]::Create($target)
          $in.CopyTo($out)
          $out.Dispose()
          $in.Dispose()
        }
      }
      finally {
        $innerZip.Dispose()
        $stream.Dispose()
      }
    }
    finally {
      $bundle.Dispose()
    }
    return $true
  }
  catch {
    return $false
  }
}

function Get-WingetHealth {
  <#
    Diagnoses the winget installation: binary presence and version, source
    health (winget source list), and registration state of the
    Microsoft.Winget.Source package. Healthy means all three pass.
  #>
  [CmdletBinding()]
  param()

  $wingetCommand = Get-Command winget -ErrorAction SilentlyContinue
  $version = $null
  $sourcesOk = $null
  $sourceError = $null
  $sourcePackagePresent = $false

  if ($wingetCommand) {
    try {
      $version = (& winget --version 2>$null | Select-Object -First 1)
    }
    catch {
      $version = $null
    }

    try {
      $sourceOutput = @(& winget source list 2>&1)
      $sourcesOk = ($LASTEXITCODE -eq 0)
      if (-not $sourcesOk) {
        $sourceError = (($sourceOutput | Select-Object -Last 3) -join ' ').Trim()
      }
    }
    catch {
      $sourcesOk = $false
      $sourceError = $_.Exception.Message
    }
  }

  $sourcePackage = Get-AppxPackage -Name 'Microsoft.Winget.Source' -ErrorAction SilentlyContinue
  $sourcePackagePresent = ($null -ne $sourcePackage)

  [pscustomobject]@{
    WingetPresent = ($null -ne $wingetCommand)
    Version = $version
    SourcesOk = $sourcesOk
    SourceError = $sourceError
    SourcePackagePresent = $sourcePackagePresent
    Healthy = (($null -ne $wingetCommand) -and $version -and ($sourcesOk -eq $true) -and $sourcePackagePresent)
  }
}

function Invoke-WingetRepair {
  <#
    Runs the auto-repair sequence:
      1. Repair-WinGetPackageManager (official Microsoft repair cmdlet).
      2. Reinstall of the Microsoft.Winget.Source package from the CDN
         (winget-cli issue #4799 fix for 0x8a15000f).
      3. winget source reset --force to restore default sources.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  if (-not $PSCmdlet.ShouldProcess('WinGet', 'Run the auto-repair sequence')) {
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Repair' -Status 'Skipped' -Detail 'WhatIf - auto-repair skipped.'
    return
  }

  # 1. Official repair cmdlet (reinstalls winget and its dependencies).
  try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client -ErrorAction SilentlyContinue)) {
      Write-Log -Message '  Installing Microsoft.WinGet.Client from PSGallery (CurrentUser scope)...' -Color Gray
      $null = Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue
      $null = Install-Module -Name Microsoft.WinGet.Client -Force -Scope CurrentUser -Repository PSGallery
    }
    Import-Module Microsoft.WinGet.Client -ErrorAction Stop

    try {
      Repair-WinGetPackageManager -AllUsers -Latest -Force -ErrorAction Stop
    }
    catch {
      Write-Log -Message '  Repair-WinGetPackageManager -Latest failed; retrying without -Latest (older module?).' -Color Yellow
      Repair-WinGetPackageManager -AllUsers -ErrorAction Stop
    }
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Repair' -Status 'Completed' -Detail 'Repair-WinGetPackageManager completed.'
  }
  catch {
    $translated = Get-WingetInstallError -Message $_.Exception.Message
    $detail = if ($translated.Known) { $translated.Detail } else { $_.Exception.Message }
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Repair' -Status 'Failed' -Detail "Repair-WinGetPackageManager failed: $detail"
  }

  # 2. winget-cli issue #4799: reinstall the Microsoft.Winget.Source package.
  #    Error 0x8a15000f "Data required by the source is missing" is caused by
  #    a corrupt or volume-lost source package; re-adding it from the CDN is
  #    the documented fix.
  $sourceMsixUrl = 'https://cdn.winget.microsoft.com/cache/source.msix'
  try {
    $existingSource = Get-AppxPackage -Name 'Microsoft.Winget.Source' -ErrorAction SilentlyContinue
    if ($existingSource) {
      try {
        Remove-AppxPackage -Package $existingSource -ErrorAction Stop
        Write-Log -Message '  Removed the broken Microsoft.Winget.Source registration.' -Color Gray
      }
      catch {
        Write-Log -Message "  Could not remove the existing source package ($($_.Exception.Message)); attempting re-add anyway." -Color Yellow
      }
    }
    Add-AppxPackage -Path $sourceMsixUrl -ErrorAction Stop
    Add-OperationResult -Results $_results -Target 'Microsoft.Winget.Source' -Source 'Winget' -Action 'Repair' -Status 'Completed' -Detail 'Source package reinstalled from cdn.winget.microsoft.com (issue #4799 fix).'
  }
  catch {
    $translated = Get-WingetInstallError -Message $_.Exception.Message
    $detail = if ($translated.Known) { $translated.Detail } else { $_.Exception.Message }
    Add-OperationResult -Results $_results -Target 'Microsoft.Winget.Source' -Source 'Winget' -Action 'Repair' -Status 'Failed' -Detail "Source package reinstall failed: $detail"
  }

  # 3. Restore default sources and clear any remaining broken source state.
  if (Get-Command winget -ErrorAction SilentlyContinue) {
    try {
      (& winget source reset --force 2>&1) | Out-Null
      $resetOk = ($LASTEXITCODE -eq 0)
      Add-OperationResult -Results $_results -Target 'WinGetSources' -Source 'Winget' -Action 'Repair' -Status $(if ($resetOk) { 'Completed' } else { 'Warn' }) -Detail $(if ($resetOk) { 'winget source reset --force completed.' } else { "winget source reset exited with code $LASTEXITCODE." })
    }
    catch {
      Add-OperationResult -Results $_results -Target 'WinGetSources' -Source 'Winget' -Action 'Repair' -Status 'Warn' -Detail "winget source reset failed: $($_.Exception.Message)"
    }
  }
}

# ---- Repair mode ------------------------------------------------------------

if ($Repair) {
  if ($ForceManual) {
    Write-Error '-ForceManual is not applicable with -Repair (the auto-repair uses the official Repair-WinGetPackageManager cmdlet).'
    exit 1
  }

  Write-Log -Message 'Repair mode: diagnosing the winget installation...' -Color Cyan
  $health = Get-WingetHealth

  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Check' -Status $(if ($health.Healthy) { 'Completed' } else { 'Failed' }) -Detail "winget: $($health.Version -or 'not found'), sources ok: $($health.SourcesOk), source package present: $($health.SourcePackagePresent)"

  if ($health.Healthy) {
    Write-Log -Message "winget is installed and healthy ($($health.Version)). No repair needed." -Color Green
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Repair' -Status 'Skipped' -Detail 'AlreadyHealthy'
    $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-WinGet'
    if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }

  Write-Log -Message 'The winget installation is broken:' -Color Red
  if (-not $health.WingetPresent) {
    Write-Log -Message '  - winget command not found.' -Color Red
  }
  if ($health.WingetPresent -and -not $health.Version) {
    Write-Log -Message '  - winget is present but does not report a version.' -Color Red
  }
  if ($health.WingetPresent -and $health.SourcesOk -eq $false) {
    Write-Log -Message "  - winget sources are broken: $($health.SourceError)" -Color Red
  }
  if (-not $health.SourcePackagePresent) {
    Write-Log -Message '  - Microsoft.Winget.Source package is not registered.' -Color Red
  }

  if (-not $WhatIfPreference -and -not $Force) {
    $repairPlan = 'The auto-repair will: 1) run Repair-WinGetPackageManager (official repair cmdlet), 2) reinstall the Microsoft.Winget.Source package from cdn.winget.microsoft.com (winget-cli issue #4799 fix), 3) reset winget sources to defaults.'
    $answer = Read-Host "$repairPlan`nType 'yes' to run the auto-repair"
    if ($answer -ne 'yes') {
      Write-Log -Message 'Auto-repair declined; no changes were made.' -Color Yellow
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Repair' -Status 'Skipped' -Detail 'Declined by user.'
      if ($PassThru -or $DryRun) { $_results }
      exit 1
    }
  }

  Invoke-WingetRepair

  Write-Log -Message 'Re-verifying winget...' -Color Cyan
  $afterHealth = Get-WingetHealth
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Verify' -Status $(if ($afterHealth.Healthy) { 'Completed' } else { 'Failed' }) -Detail "winget: $($afterHealth.Version -or 'not found'), sources ok: $($afterHealth.SourcesOk), source package present: $($afterHealth.SourcePackagePresent)"

  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-WinGet'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }

  if ($afterHealth.Healthy) {
    Write-Log -Message 'winget repaired successfully.' -Color Green
    exit 0
  }
  Write-Log -Message 'winget is still broken after the auto-repair. See the results above; manual intervention may be required (check the winget-cli issue #4799 thread for further steps).' -Color Red
  exit 1
}

# ---- Already installed? -----------------------------------------------------
if (Test-WinGetDetected) {
  Write-Log -Message "winget is already installed ($((winget --version) 2>$null))." -Color Green
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'AlreadyInstalled'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

# ---- Preferred path - Repair-WinGetPackageManager ---------------------------
if (-not $ForceManual) {
  if ($DryRun) {
    Write-Log -Message '[DRY RUN] Would install winget via Microsoft.WinGet.Client module + Repair-WinGetPackageManager.' -Color Yellow
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'DryRun - preferred method.'
  }
  else {
    Write-Log -Message 'Attempting install via Microsoft.WinGet.Client module.' -Color Yellow
    try {
      if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client -ErrorAction SilentlyContinue)) {
        Write-Log -Message '  Module not present; installing from PSGallery (CurrentUser scope).' -Color Gray
        $null = Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue
        $null = Install-Module -Name Microsoft.WinGet.Client -Force -Scope CurrentUser -Repository PSGallery
      }
      Import-Module Microsoft.WinGet.Client -ErrorAction Stop

      if ($PSCmdlet.ShouldProcess('winget', 'Repair-WinGetPackageManager')) {
        Repair-WinGetPackageManager -AllUsers -ErrorAction Stop
      }

      if (Test-WinGetDetected) {
        Write-Log -Message '  -> winget installed successfully via Repair-WinGetPackageManager.' -Color Green
        Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Installed via Repair-WinGetPackageManager.'
        $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-WinGet'
        if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
        if ($PassThru -or $DryRun) { $_results }
        exit 0
      }
      Write-Log -Message '  -> Repair path completed but winget still not resolvable; falling back to manual.' -Color Yellow
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'Repair path completed but winget not on PATH; falling back to manual.'
    }
    catch {
      $translated = Get-WingetInstallError -Message $_.Exception.Message
      $detail = if ($translated.Known) { $translated.Detail } else { $_.Exception.Message }
      Write-Log -Message "  -> Module path failed ($detail); falling back to manual download." -Color Yellow
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail $detail
    }
  }
}

# ---- DryRun - manual path ---------------------------------------------------
if ($DryRun) {
  Write-Log -Message '[DRY RUN] Would download winget assets from GitHub and provision via DISM.' -Color Yellow
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'DryRun - manual fallback.'
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-WinGet'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

# ---- Fallback path - manual asset download + DISM/portable provisioning ------
$arch = switch ($env:PROCESSOR_ARCHITECTURE) {
  'AMD64' { 'x64' }
  'ARM64' { 'arm64' }
  'x86' { 'x86' }
  default { 'x64' }
}
Write-Log -Message "Manual install for architecture: $arch" -Color Yellow

$apiBase = 'https://api.github.com/repos/microsoft/winget-cli/releases'
try {
  $release = if ($Version) {
    Invoke-RestMethod -Uri "$apiBase/tags/v$Version" -Headers @{ 'User-Agent' = 'winkit' }
  }
  else {
    Invoke-RestMethod -Uri "$apiBase/latest" -Headers @{ 'User-Agent' = 'winkit' }
  }
}
catch {
  Write-Log -Message "Could not query winget release metadata from GitHub: $($_.Exception.Message)" -Color Red
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail "GitHub API query failed: $($_.Exception.Message)"
  exit 1
}

$msix = $release.assets | Where-Object { $_.name -eq 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle' } | Select-Object -First 1
$depsZip = $release.assets | Where-Object { $_.name -eq 'DesktopAppInstaller_Dependencies.zip' } | Select-Object -First 1
$license = $release.assets | Where-Object { $_.name -like '*_License1.xml' } | Select-Object -First 1

if (-not $msix -or -not $license) {
  Write-Log -Message "Required winget release assets not found in release '$($release.tag_name)'." -Color Red
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail "Missing assets in release $($release.tag_name)"
  exit 1
}

$work = Join-Path $env:TEMP "winget-install-$(New-Guid)"
$null = New-Item -Path $work -ItemType Directory -Force -ErrorAction SilentlyContinue
try {
  $msixFile = Join-Path $work 'winget.msixbundle'
  $licenseFile = Join-Path $work 'license.xml'

  Write-Log -Message "Downloading App Installer bundle ($($release.tag_name))..." -Color Gray
  Invoke-WebRequest -Uri $msix.browser_download_url -OutFile $msixFile

  Write-Log -Message 'Downloading license...' -Color Gray
  Invoke-WebRequest -Uri $license.browser_download_url -OutFile $licenseFile

  $depPackages = @()
  if ($depsZip) {
    Write-Log -Message 'Downloading dependency bundle...' -Color Gray
    $depsZipFile = Join-Path $work 'deps.zip'
    Invoke-WebRequest -Uri $depsZip.browser_download_url -OutFile $depsZipFile
    $depsExtract = Join-Path $work 'deps'
    Expand-Archive -Path $depsZipFile -DestinationPath $depsExtract -Force
    $archDir = Join-Path $depsExtract $arch
    if (Test-Path -LiteralPath $archDir) {
      $depPackages = Get-ChildItem -LiteralPath $archDir -File |
        Where-Object { $_.Extension -in @('.appx', '.msix') } |
        Select-Object -ExpandProperty FullName
    }
  }
  if (-not $depPackages) {
    Write-Log -Message 'Dependency zip unavailable; fetching VCLibs directly.' -Color Yellow
    $vclibsUrl = if ($arch -eq 'arm64') {
      'https://aka.ms/Microsoft.VCLibs.arm64.14.00.Desktop.appx'
    }
    else {
      'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
    }
    $vclibsFile = Join-Path $work 'VCLibs.appx'
    Invoke-WebRequest -Uri $vclibsUrl -OutFile $vclibsFile
    $depPackages = @($vclibsFile)
  }

  $isServerCore = Test-ServerCore

  if ($isServerCore) {
    Write-Log -Message 'Server Core detected - performing portable winget extraction (beta: may not function properly due to missing dependencies).' -Color Yellow
    if ($PSCmdlet.ShouldProcess('winget + dependencies', 'Extract portable winget')) {
      $portableDirectory = Join-Path -Path $env:ProgramFiles -ChildPath 'Microsoft\winget'
      $extracted = Expand-WingetPortable -BundlePath $msixFile -Architecture $arch -Destination $portableDirectory
      if ($extracted) {
        Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail "Portable winget extracted to $portableDirectory (beta)."
      }
      else {
        Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail 'Portable winget extraction failed.'
      }
    }
    else {
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf'
    }

    if (-not (Test-WinGetDetected)) {
      Write-Log -Message '  -> winget extracted but the portable directory is not on PATH. Add "%ProgramFiles%\Microsoft\winget" to the system PATH or invoke winget.exe by full path.' -Color Yellow
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Extracted; portable directory needs to be added to PATH.'
    }
    else {
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Portable winget detected after extraction.'
    }
  }
  else {
    if ($PSCmdlet.ShouldProcess('winget + dependencies', 'Provision via DISM')) {
      foreach ($dep in $depPackages) {
        # Skip provisioning when an equal or newer version is already installed.
        $manifestInfo = Get-AppxManifestInfo -PackagePath $dep
        if ($manifestInfo) {
          $installed = Get-AppxPackage -Name $manifestInfo.Name -ErrorAction SilentlyContinue | Select-Object -First 1
          if ($installed -and ([version]$installed.Version -ge [version]$manifestInfo.Version)) {
            Write-Log -Message "Dependency $(Split-Path $dep -Leaf) already at version $($installed.Version); skipping." -Color Gray
            Add-OperationResult -Results $_results -Target (Split-Path $dep -Leaf) -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail "Already installed ($($installed.Version))."
            continue
          }
        }
        Write-Log -Message "Provisioning dependency: $(Split-Path $dep -Leaf)" -Color Gray
        if ($RunAsSystem) {
          $null = Add-ProvisionedAppxPackage -Online -SkipLicense -PackagePath $dep -ErrorAction Stop
        }
        else {
          $null = dism.exe /Online /Add-ProvisionedAppxPackage /PackagePath:"$dep" /SkipLicense 2>&1
        }
      }
      Write-Log -Message 'Provisioning App Installer...' -Color Gray
      $null = dism.exe /Online /Add-ProvisionedAppxPackage /PackagePath:"$msixFile" /LicensePath:"$licenseFile" 2>&1

      if (-not $RunAsSystem) {
        try {
          Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
        }
        catch {
          $translated = Get-WingetInstallError -Message $_.Exception.Message
          $detail = if ($translated.Known) { $translated.Detail } else { $_.Exception.Message }
          Write-Log -Message "Per-user registration step skipped: $detail" -Color Yellow
        }
      }
      else {
        Write-Log -Message 'SYSTEM account: per-user registration skipped (provisioned registration used instead).' -Color Gray
      }
    }

    if (Test-WinGetDetected) {
      Write-Log -Message '  -> winget installed successfully.' -Color Green
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Installed via manual asset download + DISM.'
    }
    else {
      Write-Log -Message '  -> winget provisioned but not yet on PATH - a sign-out/in may be required.' -Color Yellow
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Provisioned; sign-out required for PATH.'
    }
  }
}
catch {
  $translated = Get-WingetInstallError -Message $_.Exception.Message
  if ($translated.Benign) {
    Write-Log -Message "  -> $($translated.Detail)" -Color Green
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail $translated.Detail
  }
  else {
    $detail = if ($translated.Known) { $translated.Detail } else { $_.Exception.Message }
    Write-Log -Message "  -> FAILED - manual install: $detail" -Color Red
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail $detail
  }
}
finally {
  Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-WinGet'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

exit 0
