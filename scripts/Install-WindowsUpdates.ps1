Import-Module PSFoundation -Force

#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Installs Windows Updates, Microsoft Store updates, or both.

.DESCRIPTION
  Profile-based update orchestrator. Searches for, lists, and installs
  updates from Windows Update and the Microsoft Store.

  Requires the PSWindowsUpdate module for Windows Update management.
  Supply -EnsureInstalled to install it from PSGallery if missing.
  Requires administrator elevation for Windows Update operations.
  Store updates do not require elevation.

.PARAMETER Profile
  Update scope. Recommended installs Security + Critical + UpdateRollup
  updates only. All installs everything. StoreOnly installs only Microsoft
  Store app updates. Defaults to Recommended.

.PARAMETER IncludeStore
  Also check for and install Store app updates after Windows Updates.
  Ignored when Profile is StoreOnly.

.PARAMETER KBArticleID
  Install a specific Windows Update KB only. Overrides the profile for
  the Windows Update pass.

.PARAMETER AutoReboot
  Allow automatic system restart if a Windows Update requires it.

.PARAMETER AcceptAll
  Skip per-update confirmation. All eligible updates are installed
  without prompting.

.PARAMETER EnsureInstalled
  Install the PSWindowsUpdate module from PSGallery if it is not
  already available.

.PARAMETER DryRun
  List available updates without installing anything.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Install-WindowsUpdates.ps1
  Installs Recommended security and critical updates.

.EXAMPLE
  PS> ./Install-WindowsUpdates.ps1 -Profile All -AutoReboot

.EXAMPLE
  PS> ./Install-WindowsUpdates.ps1 -Profile StoreOnly -DryRun

.EXAMPLE
  PS> ./Install-WindowsUpdates.ps1 -Profile Recommended -IncludeStore -EnsureInstalled

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [ValidateSet('Recommended', 'All', 'StoreOnly')]
  [string]
  $Profile = 'Recommended',

  [Parameter(Mandatory = $false)]
  [switch]
  $IncludeStore,

  [Parameter(Mandatory = $false)]
  [string]
  $KBArticleID,

  [Parameter(Mandatory = $false)]
  [switch]
  $AutoReboot,

  [Parameter(Mandatory = $false)]
  [switch]
  $AcceptAll,

  [Parameter(Mandatory = $false)]
  [switch]
  $EnsureInstalled,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

# -----------------------------------------------------------------------------

# Elevation guard: Windows Update operations require admin (StoreOnly is the exception)
if ($Profile -ne 'StoreOnly' -and -not (Test-Elevation)) {
  Write-Error 'Windows Update operations require administrator privileges. Use -Profile StoreOnly for app-only updates without elevation.'
  exit 1
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no updates will be installed`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

# ---- Ensure PSWindowsUpdate if requested ------------------------------------
if ($Profile -ne 'StoreOnly' -and -not (Test-PSWindowsUpdateAvailable)) {
  if ($EnsureInstalled) {
    if ($DryRun) {
      Write-Log -Message '[DRY RUN] Would install PSWindowsUpdate from PSGallery.' -Color Yellow
      Add-OperationResult -Results $_results -Target 'PSWindowsUpdate' -Source 'PowerShellGallery' -Action 'Install' -Status 'Skipped' -Detail 'DryRun'
    }
    else {
      Write-Log -Message 'PSWindowsUpdate not found - installing from PSGallery...' -Color Yellow
      try {
        $null = Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -Repository PSGallery -ErrorAction Stop
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        Write-Log -Message '  -> PSWindowsUpdate installed.' -Color Green
        Add-OperationResult -Results $_results -Target 'PSWindowsUpdate' -Source 'PowerShellGallery' -Action 'Install' -Status 'Completed' -Detail 'Installed from PSGallery.'
      }
      catch {
        Write-Log -Message "  -> FAILED: $_" -Color Red
        Add-OperationResult -Results $_results -Target 'PSWindowsUpdate' -Source 'PowerShellGallery' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
        if ($PassThru -or $DryRun) { $_results }
        exit 1
      }
    }
  }
  else {
    Write-Log -Message 'PSWindowsUpdate module not found. Re-run with -EnsureInstalled to install it from PSGallery.' -Color Red
    exit 1
  }
}

# ---- Windows Update pass ----------------------------------------------------
if ($Profile -ne 'StoreOnly') {
  Write-Log -Message "=== Windows Update (profile: $Profile) ===" -Color Cyan

  $_categoryFilter = if ($Profile -eq 'Recommended') {
    @('Critical Updates', 'Security Updates', 'Update Rollups')
  }
  else { $null }

  $_searchParams = @{}
  if ($KBArticleID) { $_searchParams['KBArticleID'] = $KBArticleID }
  if ($_categoryFilter) { $_searchParams['Category'] = $_categoryFilter }

  $_updates = Get-WindowsUpdate @_searchParams
  if (-not $_updates) {
    Write-Log -Message 'No applicable updates found.' -Color Gray
    Add-OperationResult -Results $_results -Target 'WindowsUpdate' -Source 'WindowsUpdate' -Action 'Search' -Status 'Completed' -Detail 'NoUpdatesAvailable'
  }
  else {
    Write-Log -Message "Available updates: $($_updates.Count)" -Color Yellow
    foreach ($_u in $_updates) {
      Write-Log -Message "  - $($_u.KB): $($_u.Title)" -Color Gray
    }

    if ($DryRun) {
      Write-Log -Message "`n[DRY RUN] Would install $($_updates.Count) update(s)." -Color Yellow
      Add-OperationResult -Results $_results -Target 'WindowsUpdate' -Source 'WindowsUpdate' -Action 'Search' -Status 'Skipped' -Detail "DryRun - $($_updates.Count) update(s) available"
    }
    elseif ($AcceptAll -or $PSCmdlet.ShouldProcess("$($_updates.Count) update(s)", 'Install')) {
      Write-Log -Message "`nInstalling $($_updates.Count) update(s)..." -Color Yellow
      try {
        $_installResult = $_updates | Install-WindowsUpdate -AutoReboot:$AutoReboot -IgnoreReboot:( -not $AutoReboot)
        $_installedCount = @($_installResult | Where-Object { $_.HResult -eq 0 }).Count
        $_failedCount = @($_installResult | Where-Object { $_.HResult -ne 0 }).Count
        Write-Log -Message "Installed: $_installedCount | Failed: $_failedCount" -Color $(if ($_failedCount -gt 0) { 'Yellow' } else { 'Green' })
        Add-OperationResult -Results $_results -Target 'WindowsUpdate' -Source 'WindowsUpdate' -Action 'Install' -Status 'Completed' -Detail "Installed $_installedCount, failed $_failedCount"
      }
      catch {
        Write-Log -Message "  -> FAILED: $_" -Color Red
        Add-OperationResult -Results $_results -Target 'WindowsUpdate' -Source 'WindowsUpdate' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
      }
    }

    if (Test-WindowsUpdateRebootRequired) {
      Write-Log -Message "`nA system reboot is required to complete the updates." -Color Yellow
    }
  }
}

# ---- Store update pass ------------------------------------------------------
if ($Profile -eq 'StoreOnly' -or $IncludeStore) {
  Write-Log -Message "`n=== Microsoft Store Updates ===" -Color Cyan

  $_storeUpdates = Get-MSStoreUpdate
  if (-not $_storeUpdates -or $_storeUpdates.Count -eq 0) {
    Write-Log -Message 'No Store updates available.' -Color Gray
    Add-OperationResult -Results $_results -Target 'MSStoreUpdate' -Source 'MSStoreUpdate' -Action 'Search' -Status 'Completed' -Detail 'NoUpdatesAvailable'
  }
  else {
    Write-Log -Message "Available Store updates: $($_storeUpdates.Count)" -Color Yellow
    foreach ($_s in $_storeUpdates) {
      Write-Log -Message "  - $($_s.PackageFamilyName)" -Color Gray
    }

    if ($DryRun) {
      Write-Log -Message "`n[DRY RUN] Would install $($_storeUpdates.Count) Store update(s)." -Color Yellow
      Add-OperationResult -Results $_results -Target 'MSStoreUpdate' -Source 'MSStoreUpdate' -Action 'Search' -Status 'Skipped' -Detail "DryRun - $($_storeUpdates.Count) update(s) available"
    }
    elseif ($AcceptAll -or $PSCmdlet.ShouldProcess("$($_storeUpdates.Count) Store update(s)", 'Install')) {
      Write-Log -Message "`nInstalling $($_storeUpdates.Count) Store update(s)..." -Color Yellow
      $_storeResult = $_storeUpdates | Install-MSStoreUpdate
      $_storeCompleted = @($_storeResult | Where-Object { $_.Status -eq 'Completed' }).Count
      $_storeSkipped = @($_storeResult | Where-Object { $_.Status -eq 'Skipped' }).Count
      $_storeFailed = @($_storeResult | Where-Object { $_.Status -eq 'Failed' }).Count
      Write-Log -Message "Store: $_storeCompleted completed | $_storeSkipped skipped | $_storeFailed failed" -Color $(if ($_storeFailed -gt 0) { 'Yellow' } else { 'Green' })
      Add-OperationResult -Results $_results -Target 'MSStoreUpdate' -Source 'MSStoreUpdate' -Action 'Install' -Status 'Completed' -Detail "Completed $_storeCompleted, skipped $_storeSkipped, failed $_storeFailed"
    }
  }
}

# ---- Summary ----------------------------------------------------------------
$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Install-WindowsUpdates'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

exit 0
