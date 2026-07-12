#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Installs the Windows Package Manager (winget) and its dependencies.

.DESCRIPTION
  Two install paths, tried in order:

  1. Repair-WinGetPackageManager (preferred). Installs the
     Microsoft.WinGet.Client module from PSGallery and uses the
     official repair cmdlet to bootstrap winget and all current
     dependencies natively.

  2. Manual asset download (fallback). Pulls the App Installer
     msixbundle, framework dependency packages, and license from a
     GitHub release, then provisions them with DISM. Used when
     PSGallery is unreachable or the module path fails.

  Requires administrator elevation. Skips if winget is already
  installed.

.PARAMETER Version
  Specific winget release version for the manual fallback. The
  preferred module path always installs the current stable regardless
  of this value.

.PARAMETER ForceManual
  Skip the Repair-WinGetPackageManager path and use manual asset
  download directly.

.PARAMETER DryRun
  Preview the install steps without executing them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Install-WinGet.ps1
  Installs the latest winget via the best available method.

.EXAMPLE
  PS> ./Install-WinGet.ps1 -ForceManual -DryRun

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
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

# ---- Already installed? -----------------------------------------------------
if (Get-Command winget -ErrorAction SilentlyContinue) {
  Write-Log -Message "winget is already installed ($((winget --version) 2>$null))." -Color Green
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'AlreadyInstalled'
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

# ---- Preferred path â€" Repair-WinGetPackageManager ---------------------------
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

      if (Get-Command winget -ErrorAction SilentlyContinue) {
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
      Write-Log -Message "  -> Module path failed ($($_.Exception.Message)); falling back to manual download." -Color Yellow
      Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

# ---- DryRun â€" manual path ---------------------------------------------------
if ($DryRun) {
  Write-Log -Message '[DRY RUN] Would download winget assets from GitHub and provision via DISM.' -Color Yellow
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Skipped' -Detail 'DryRun - manual fallback.'
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-WinGet'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

# ---- Fallback path â€" manual asset download + DISM provisioning ---------------
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

  if ($PSCmdlet.ShouldProcess('winget + dependencies', 'Provision via DISM')) {
    foreach ($dep in $depPackages) {
      Write-Log -Message "Provisioning dependency: $(Split-Path $dep -Leaf)" -Color Gray
      $null = dism.exe /Online /Add-ProvisionedAppxPackage /PackagePath:"$dep" /SkipLicense 2>&1
    }
    Write-Log -Message 'Provisioning App Installer...' -Color Gray
    $null = dism.exe /Online /Add-ProvisionedAppxPackage /PackagePath:"$msixFile" /LicensePath:"$licenseFile" 2>&1

    try {
      Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
    }
    catch {
      Write-Log -Message "Per-user registration step skipped: $($_.Exception.Message)" -Color Yellow
    }
  }

  if (Get-Command winget -ErrorAction SilentlyContinue) {
    Write-Log -Message '  -> winget installed successfully.' -Color Green
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Installed via manual asset download + DISM.'
  }
  else {
    Write-Log -Message '  -> winget provisioned but not yet on PATH â€" a sign-out/in may be required.' -Color Yellow
    Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Completed' -Detail 'Provisioned; sign-out required for PATH.'
  }
}
catch {
  Write-Log -Message "  -> FAILED - manual install: $_" -Color Red
  Add-OperationResult -Results $_results -Target 'WinGet' -Source 'Winget' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
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
