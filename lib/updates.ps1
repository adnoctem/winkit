#Requires -Version 5.0

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'AcceptAll', Justification = 'Used as a parameter set selector by Install-WindowsUpdate.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'All', Justification = 'Used as a parameter set selector by Install-MSStoreUpdate.')]

# ---- WinRT interop cache (both PS5.1 and PS7+) ---------------------------------
$script:AppInstallManager = $null
$_asTaskGeneric = $null

function _getAppInstallManager {
  if ($null -ne $script:AppInstallManager) { return $script:AppInstallManager }
  Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
  [Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallManager, Windows.ApplicationModel.Store.Preview, ContentType = WindowsRuntime] | Out-Null
  $script:AppInstallManager = New-Object Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallManager
  return $script:AppInstallManager
}

function _awaitWinRt($WinRtTask, $ResultType) {
  if ($null -eq $script:_asTaskGeneric) {
    $script:_asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  }
  $asTask = $script:_asTaskGeneric.MakeGenericMethod($ResultType)
  $netTask = $asTask.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  return $netTask.Result
}

# ---- Test-PSWindowsUpdateAvailable ---------------------------------------------
function Test-PSWindowsUpdateAvailable {
  <#
    .SYNOPSIS
      Returns $true if the PSWindowsUpdate module is installed.
    .DESCRIPTION
      Checks Get-Module -ListAvailable. Does not import the module.
    .EXAMPLE
      PS> Test-PSWindowsUpdateAvailable
    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([bool])]
  param()

  return $null -ne (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)
}

# ---- Get-WindowsUpdate ---------------------------------------------------------
function Get-WindowsUpdate {
  <#
    .SYNOPSIS
      Lists available Windows updates.

    .DESCRIPTION
      Wraps Get-WUList from the PSWindowsUpdate module. Filters by category,
      KB article ID, or title wildcard. Requires admin elevation and the
      PSWindowsUpdate module.

    .PARAMETER Category
      One or more update categories to include. Common values: Security,
      Critical, Definition, Update, UpdateRollup, Drivers.

    .PARAMETER KBArticleID
      Filter to a specific KB article.

    .PARAMETER Title
      Wildcard title filter forwarded to -Title.

    .EXAMPLE
      PS> Get-WindowsUpdate -Category Security, Critical

    .EXAMPLE
      PS> Get-WindowsUpdate -KBArticleID 'KB5021233'

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject[]])]
  param (
    [Parameter(Mandatory = $false)]
    [string[]]
    $Category,

    [Parameter(Mandatory = $false)]
    [string]
    $KBArticleID,

    [Parameter(Mandatory = $false)]
    [string]
    $Title
  )

  if (-not (Test-PSWindowsUpdateAvailable)) {
    Write-Error 'PSWindowsUpdate module is not installed. Install it with: Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser'
    return
  }

  Import-Module PSWindowsUpdate -Force -ErrorAction Stop

  $_params = @{ ErrorAction = 'Stop' }
  if ($PSBoundParameters.ContainsKey('KBArticleID')) { $_params['KBArticleID'] = $KBArticleID }
  if ($PSBoundParameters.ContainsKey('Title')) { $_params['Title'] = $Title }

  try {
    $_updates = Get-WUList @_params

    if ($Category) {
      $_updates = $_updates | Where-Object { $_.Categories -and ($Category | Where-Object { $_ -in $_.Categories.Name }) }
    }

    return $_updates
  }
  catch {
    Write-Error "Failed to query available updates: $_"
    return
  }
}

# ---- Install-WindowsUpdate -----------------------------------------------------
function Install-WindowsUpdate {
  <#
    .SYNOPSIS
      Installs one or more Windows updates.

    .DESCRIPTION
      Accepts pipeline input from Get-WindowsUpdate, a specific KBArticleID,
      or the -AcceptAll switch to install everything available.

      Returns an object per update with KB, Title, HResult, and RebootRequired.
      When -AutoReboot is used and an update requires a restart, the system
      reboots after completion.

    .PARAMETER InputObject
      Update object from Get-WindowsUpdate (pipeline).

    .PARAMETER KBArticleID
      Install a specific KB by article ID.

    .PARAMETER AcceptAll
      Install all available updates without filtering.

    .PARAMETER AutoReboot
      Automatically restart the system if an update requires it.

    .PARAMETER IgnoreReboot
      Suppress the reboot-required check after install.

    .EXAMPLE
      PS> Get-WindowsUpdate -Category Security | Install-WindowsUpdate

    .EXAMPLE
      PS> Install-WindowsUpdate -KBArticleID 'KB5021233'

    .EXAMPLE
      PS> Install-WindowsUpdate -AcceptAll -AutoReboot

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(DefaultParameterSetName = 'Pipeline')]
  [OutputType([PSCustomObject[]])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Pipeline')]
    [psobject[]]
    $InputObject,

    [Parameter(Mandatory = $true, ParameterSetName = 'KB')]
    [string]
    $KBArticleID,

    [Parameter(Mandatory = $true, ParameterSetName = 'All')]
    [switch]
    $AcceptAll,

    [Parameter(Mandatory = $false)]
    [switch]
    $AutoReboot,

    [Parameter(Mandatory = $false)]
    [switch]
    $IgnoreReboot
  )

  begin {
    if (-not (Test-PSWindowsUpdateAvailable)) {
      Write-Error 'PSWindowsUpdate module is not installed.'
      return
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop
    $_results = New-Object System.Collections.ArrayList
    $_updateObjects = New-Object System.Collections.ArrayList
    $null = $PSBoundParameters['AcceptAll']
  }

  process {
    if ($PSCmdlet.ParameterSetName -eq 'Pipeline') {
      foreach ($_u in $InputObject) {
        [void]$_updateObjects.Add($_u)
      }
    }
  }

  end {
    if ($PSCmdlet.ParameterSetName -eq 'KB') {
      $_found = Get-WindowsUpdate -KBArticleID $KBArticleID
      if (-not $_found) {
        Write-Error "No update found for KB: $KBArticleID"
        return
      }
      [void]$_updateObjects.Add($_found)
    }

    if ($PSCmdlet.ParameterSetName -eq 'All') {
      $_all = Get-WindowsUpdate
      if (-not $_all) {
        Write-Error 'No updates available.'
        return
      }
      foreach ($_u in $_all) { [void]$_updateObjects.Add($_u) }
    }

    if ($_updateObjects.Count -eq 0) {
      Write-Error 'No updates to install.'
      return
    }

    try {
      $installResult = Install-WUUpdates -Updates $_updateObjects -AcceptAll -AutoReboot:$AutoReboot -IgnoreReboot:$IgnoreReboot -ErrorAction Stop

      foreach ($_item in $installResult) {
        $obj = [PSCustomObject]@{
          KB = $_item.KB
          Title = $_item.Title
          HResult = $_item.HResult
          Result = $_item.Result
          RebootRequired = $_item.RebootRequired
        }
        [void]$_results.Add($obj)
      }

      return $_results
    }
    catch {
      Write-Error "Failed to install updates: $_"
      return
    }
  }
}

# ---- Hide-WindowsUpdate --------------------------------------------------------
function Hide-WindowsUpdate {
  <#
    .SYNOPSIS
      Hides an available update so it does not appear in future scans.

    .DESCRIPTION
      Wraps Hide-WUUpdate. Useful for suppressing known-problematic updates
      or driver packages. Accepts pipeline input from Get-WindowsUpdate.

    .PARAMETER InputObject
      Update object from Get-WindowsUpdate (pipeline).

    .PARAMETER KBArticleID
      Hide a specific KB.

    .EXAMPLE
      PS> Get-WindowsUpdate -Title '*Driver*' | Hide-WindowsUpdate

    .EXAMPLE
      PS> Hide-WindowsUpdate -KBArticleID 'KB5021233'

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Pipeline')]
    [psobject[]]
    $InputObject,

    [Parameter(Mandatory = $true, ParameterSetName = 'KB')]
    [string]
    $KBArticleID
  )

  begin {
    if (-not (Test-PSWindowsUpdateAvailable)) {
      Write-Error 'PSWindowsUpdate module is not installed.'
      return
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop
    $_updates = New-Object System.Collections.ArrayList
  }

  process {
    if ($PSCmdlet.ParameterSetName -eq 'Pipeline') {
      foreach ($_u in $InputObject) {
        [void]$_updates.Add($_u)
      }
    }
  }

  end {
    if ($PSCmdlet.ParameterSetName -eq 'KB') {
      $_found = Get-WindowsUpdate -KBArticleID $KBArticleID
      if (-not $_found) { return }
      [void]$_updates.Add($_found)
    }

    if ($_updates.Count -eq 0) { return }

    foreach ($_u in $_updates) {
      $_label = "$($_u.KB) - $($_u.Title)"
      if (-not $PSCmdlet.ShouldProcess($_label, 'Hide update')) { continue }

      try {
        $null = Hide-WUUpdate -Update $_u -Confirm:$false -ErrorAction Stop
        Write-Verbose "Hidden: $_label"
      }
      catch {
        Write-Error "Failed to hide update '$_label': $_"
      }
    }
  }
}

# ---- Get-WindowsUpdateHistory --------------------------------------------------
function Get-WindowsUpdateHistory {
  <#
    .SYNOPSIS
      Returns installed Windows update history.

    .DESCRIPTION
      Wraps Get-WUHistory. Optionally filtered by -Last N recent entries
      or a specific -KBArticleID.

    .PARAMETER Last
      Number of most recent history entries to return.

    .PARAMETER KBArticleID
      Filter to a specific KB article.

    .EXAMPLE
      PS> Get-WindowsUpdateHistory -Last 10

    .EXAMPLE
      PS> Get-WindowsUpdateHistory -KBArticleID 'KB5021233'

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject[]])]
  param (
    [Parameter(Mandatory = $false)]
    [int]
    $Last,

    [Parameter(Mandatory = $false)]
    [string]
    $KBArticleID
  )

  if (-not (Test-PSWindowsUpdateAvailable)) {
    Write-Error 'PSWindowsUpdate module is not installed.'
    return
  }

  Import-Module PSWindowsUpdate -Force -ErrorAction Stop

  $_params = @{ ErrorAction = 'Stop' }
  if ($PSBoundParameters.ContainsKey('Last')) { $_params['Last'] = $Last }
  if ($PSBoundParameters.ContainsKey('KBArticleID')) { $_params['KBArticleID'] = $KBArticleID }

  try {
    return Get-WUHistory @_params
  }
  catch {
    Write-Error "Failed to get update history: $_"
    return
  }
}

# ---- Uninstall-WindowsUpdate ---------------------------------------------------
function Uninstall-WindowsUpdate {
  <#
    .SYNOPSIS
      Uninstalls a previously installed Windows update.

    .DESCRIPTION
      Wraps Remove-WUUpdate. Requires a KBArticleID. Asks for confirmation
      by default since removing security updates can leave the system
      vulnerable.

    .PARAMETER KBArticleID
      KB article ID of the update to uninstall.

    .EXAMPLE
      PS> Uninstall-WindowsUpdate -KBArticleID 'KB5021233'

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $KBArticleID
  )

  if (-not (Test-PSWindowsUpdateAvailable)) {
    Write-Error 'PSWindowsUpdate module is not installed.'
    return
  }

  Import-Module PSWindowsUpdate -Force -ErrorAction Stop

  $_history = Get-WindowsUpdateHistory -KBArticleID $KBArticleID
  if (-not $_history) {
    Write-Error "No installed update found for KB: $KBArticleID"
    return
  }

  if (-not $PSCmdlet.ShouldProcess($KBArticleID, 'Uninstall Windows update')) {
    return
  }

  try {
    Remove-WUUpdate -KBArticleID $KBArticleID -Confirm:$false -ErrorAction Stop
    Write-Verbose "Uninstalled: $KBArticleID"
  }
  catch {
    Write-Error "Failed to uninstall ${KBArticleID}: $_"
  }
}

# ---- Test-WindowsUpdateRebootRequired ------------------------------------------
function Test-WindowsUpdateRebootRequired {
  <#
    .SYNOPSIS
      Returns $true if a reboot is pending after a Windows update install.

    .EXAMPLE
      PS> if (Test-WindowsUpdateRebootRequired) { Restart-Computer }

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([bool])]
  param()

  if (-not (Test-PSWindowsUpdateAvailable)) {
    Write-Error 'PSWindowsUpdate module is not installed.'
    return $false
  }

  Import-Module PSWindowsUpdate -Force -ErrorAction Stop

  try {
    $status = Get-WURebootStatus -ErrorAction Stop
    return $status.RebootRequired
  }
  catch {
    Write-Error "Failed to get reboot status: $_"
    return $false
  }
}

# ---- Get-WindowsUpdateConfiguration --------------------------------------------
function Get-WindowsUpdateConfiguration {
  <#
    .SYNOPSIS
      Returns current Windows Update service configuration.

    .DESCRIPTION
      Wraps Get-WUSettings. Returns an object with ServerSelection,
      ServiceID, TargetGroup, and NotificationLevel properties.

    .EXAMPLE
      PS> Get-WindowsUpdateConfiguration

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param()

  if (-not (Test-PSWindowsUpdateAvailable)) {
    Write-Error 'PSWindowsUpdate module is not installed.'
    return
  }

  Import-Module PSWindowsUpdate -Force -ErrorAction Stop

  try {
    return Get-WUSettings -ErrorAction Stop
  }
  catch {
    Write-Error "Failed to get Windows Update configuration: $_"
    return
  }
}

# ---- Get-MSStoreUpdate ---------------------------------------------------------
function Get-MSStoreUpdate {
  <#
    .SYNOPSIS
      Lists available Microsoft Store app updates.

    .DESCRIPTION
      Uses the WinRT AppInstallManager API to search for available Store
      app updates. When no -PackageFamilyName is supplied, returns updates
      for all installed Store-packaged apps.

      Does not require administrator elevation for user-scope updates.

    .PARAMETER PackageFamilyName
      Specific package family name to check, e.g. Microsoft.WindowsStore_8wekyb3d8bbwe.
      Supports wildcard matching.

    .EXAMPLE
      PS> Get-MSStoreUpdate

    .EXAMPLE
      PS> Get-MSStoreUpdate -PackageFamilyName '*WindowsTerminal*'

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject[]])]
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $PackageFamilyName
  )

  try {
    $mgr = _getAppInstallManager
    $searchOp = $mgr.SearchForUpdatesAsync()
    $appUpdates = _awaitWinRt $searchOp ([System.Collections.Generic.IReadOnlyList[Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallItem]])

    if (-not $appUpdates -or $appUpdates.Count -eq 0) {
      return @()
    }

    $results = foreach ($item in $appUpdates) {
      $obj = [PSCustomObject]@{
        PackageFamilyName = $item.PackageFamilyName
        ProductId = $item.ProductId
        ItemKind = $item.ItemKind
        ErrorCode = $item.ErrorCode
        InstallType = $item.InstallType
        CompletedInstallCount = $item.CompletedInstallCount
        TotalInstallCount = $item.TotalInstallCount
      }
      if ($PackageFamilyName) {
        if ($obj.PackageFamilyName -like $PackageFamilyName) { $obj }
      }
      else { $obj }
    }

    return $results
  }
  catch {
    Write-Error "Failed to search for Store updates: $_"
    return
  }
}

# ---- Install-MSStoreUpdate -----------------------------------------------------
function Install-MSStoreUpdate {
  <#
    .SYNOPSIS
      Installs available Microsoft Store app updates.

    .DESCRIPTION
      Uses the WinRT AppInstallManager API to trigger Store app updates.
      Accepts pipeline input from Get-MSStoreUpdate or a specific
      -PackageFamilyName. Progress is reported via Write-Log.

      Does not require administrator elevation for user-scope updates.
      Non-installed apps that match -PackageFamilyName are skipped silently.

    .PARAMETER InputObject
      Store update object from Get-MSStoreUpdate (pipeline).

    .PARAMETER PackageFamilyName
      Specific package family name to update.

    .PARAMETER All
      Install all available Store updates.

    .EXAMPLE
      PS> Get-MSStoreUpdate | Install-MSStoreUpdate

    .EXAMPLE
      PS> Install-MSStoreUpdate -PackageFamilyName '*WindowsStore*'

    .EXAMPLE
      PS> Install-MSStoreUpdate -All

    .LINK
      https://github.com/adnoctem/winkit/lib/updates.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(DefaultParameterSetName = 'Pipeline')]
  [OutputType([PSCustomObject[]])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Pipeline')]
    [psobject[]]
    $InputObject,

    [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
    [string]
    $PackageFamilyName,

    [Parameter(Mandatory = $true, ParameterSetName = 'All')]
    [switch]
    $All
  )

  begin {
    $_items = New-Object System.Collections.ArrayList
    $_results = New-Object System.Collections.ArrayList
    $null = $PSBoundParameters['All']
  }

  process {
    if ($PSCmdlet.ParameterSetName -eq 'Pipeline') {
      foreach ($_i in $InputObject) {
        [void]$_items.Add($_i)
      }
    }
  }

  end {
    if ($PSCmdlet.ParameterSetName -eq 'Name') {
      $_found = Get-MSStoreUpdate -PackageFamilyName $PackageFamilyName
      if (-not $_found) {
        Write-Error "No Store updates found for: $PackageFamilyName"
        return
      }
      foreach ($_f in $_found) { [void]$_items.Add($_f) }
    }

    if ($PSCmdlet.ParameterSetName -eq 'All') {
      $_found = Get-MSStoreUpdate
      if (-not $_found) {
        Write-Error 'No Store updates available.'
        return
      }
      foreach ($_f in $_found) { [void]$_items.Add($_f) }
    }

    if ($_items.Count -eq 0) {
      Write-Error 'No Store updates to install.'
      return
    }

    $mgr = _getAppInstallManager

    foreach ($_item in $_items) {
      $_pfn = $_item.PackageFamilyName
      Write-Log -Message "  Installing: $_pfn" -Color Yellow

      try {
        $updateOp = $mgr.UpdateAppByPackageFamilyNameAsync($_pfn)
        $updateResult = _awaitWinRt $updateOp ([Windows.ApplicationModel.Store.Preview.InstallControl.AppInstallItem])

        if ($null -eq $updateResult) {
          Write-Log -Message "    No update available or already up to date." -Color Gray
          $_item = [PSCustomObject]@{
            PackageFamilyName = $_pfn
            Status = 'NoUpdate'
          }
          [void]$_results.Add($_item)
          continue
        }

        while ($true) {
          $currentStatus = $updateResult.GetCurrentStatus()
          if ($null -eq $currentStatus) { break }

          $pct = $currentStatus.PercentComplete
          if ($pct -eq 100) {
            Write-Log -Message "    Install completed ($_pfn)" -Color Green
            $_item = [PSCustomObject]@{
              PackageFamilyName = $_pfn
              Status = 'Completed'
              ErrorCode = $currentStatus.ErrorCode
            }
            [void]$_results.Add($_item)
            break
          }

          if ($currentStatus.ErrorCode -and $currentStatus.ErrorCode -ne 0) {
            Write-Log -Message "    Failed: $_pfn (ErrorCode $($currentStatus.ErrorCode))" -Color Red
            $_item = [PSCustomObject]@{
              PackageFamilyName = $_pfn
              Status = 'Failed'
              ErrorCode = $currentStatus.ErrorCode
            }
            [void]$_results.Add($_item)
            break
          }

          if ($pct % 25 -eq 0 -or $pct -eq 15) {
            Write-Log -Message "    ${_pfn}: $pct%" -Color Gray
          }
          Start-Sleep -Seconds 3
        }
      }
      catch [System.AggregateException] {
        $inner = $_.Exception.InnerException
        Write-Log -Message "    Skipped: $_pfn (not installed or not applicable: $inner)" -Color Gray
        $_item = [PSCustomObject]@{
          PackageFamilyName = $_pfn
          Status = 'Skipped'
        }
        [void]$_results.Add($_item)
      }
      catch {
        Write-Log -Message "    Unexpected error for ${_pfn}: $_" -Color Red
        $_item = [PSCustomObject]@{
          PackageFamilyName = $_pfn
          Status = 'Failed'
        }
        [void]$_results.Add($_item)
      }
    }

    return $_results
  }
}
