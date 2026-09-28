#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.3.0' }

<#
.SYNOPSIS
  Deduplicates Outlook mail items per folder using the transport Message-ID.
.DESCRIPTION
  Walks every mail folder in the target Outlook store, keys each received mail
  item by its RFC Message-ID, and moves every occurrence after the first into a
  review folder. It never hard-deletes messages.

  Deduplication scope is per folder. The same message legitimately living in
  two different folders is preserved.
  Shows folder traversal and throttled item progress while inspecting headers
  and processing duplicates. ProgressPreference controls the progress display.
.PARAMETER StoreName
  Display name of the Outlook store to process. If omitted, the default
  delivery store is used. Run with -Verbose to list detected stores.
.PARAMETER ReviewFolderName
  Top-level folder created under the store root for duplicate review.
.PARAMETER ExcludeFolders
  Folder display names to skip entirely.
.PARAMETER ReportPath
  Optional CSV path containing Keep, MoveDuplicate, and SkipNoMessageId results.
.PARAMETER DryRun
  Preview changes without moving duplicate messages.
.PARAMETER PassThru
  Return structured operation result objects.
.PARAMETER QuitOutlook
  Quit the Outlook application object on exit. Leave off if Outlook was already
  open interactively.
.PARAMETER IgnoreAdministrator
  Allow an elevated PowerShell session. Use only when Outlook intentionally
  runs elevated under the same Windows user. Does not switch users or profiles.
  Elevated execution is otherwise refused, including during previews.
.EXAMPLE
  PS> .\Optimize-Outlook.ps1 -StoreName 'user@example.com' -ReportPath .\dedup-preview.csv -DryRun
.EXAMPLE
  PS> .\Optimize-Outlook.ps1 -StoreName 'user@example.com' -ReportPath .\dedup-run.csv
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not applicable - Outlook is a desktop client.
  SYSTEM-account execution: not applicable - requires an interactive Outlook profile.
  Outlook version: 2007 (version 12) or later - PropertyAccessor is required.
  Bitness: Outlook 2007 is 32-bit only - run under 32-bit PowerShell (x86).
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [string]
  $StoreName,

  [string]
  $ReviewFolderName = '_Duplicates_Review',

  [string[]]
  $ExcludeFolders = @(
    'Deleted Items',
    'Junk Email',
    'Junk E-mail',
    'Outbox',
    'Sync Issues',
    'Conflicts',
    'Local Failures',
    'Server Failures'
  ),

  [string]
  $ReportPath,

  [switch]
  $DryRun,

  [switch]
  $PassThru,

  [switch]
  $IgnoreAdministrator,

  [switch]
  $QuitOutlook
)

Import-Module PSFoundation -Force

# Outlook automation must use the mailbox user's interactive security context.
$_outlookUser = Get-UserInfo
if ($_outlookUser.IsAdministrator) {
  if (-not $IgnoreAdministrator) {
    throw "PowerShell is elevated as '$($_outlookUser.UserName)'. Run this script from a non-elevated PowerShell window as the Windows user who runs Outlook. Use -IgnoreAdministrator only when Outlook intentionally runs elevated under that same user."
  }

  Write-Warning "IgnoreAdministrator permits elevated execution as '$($_outlookUser.UserName)'. Outlook must run under the same Windows user and elevation. This override does not switch users or profiles."
}

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no Outlook messages will be moved`n" -Color Yellow
}

# olMail object class; PR_TRANSPORT_MESSAGE_HEADERS in Unicode then ANSI form.
$script:OL_MAIL = 43
$script:HDR_TAGS = @(
  'http://schemas.microsoft.com/mapi/proptag/0x007D001F',
  'http://schemas.microsoft.com/mapi/proptag/0x007D001E'
)

$_results = New-Object System.Collections.ArrayList

function Get-MessageId {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'Outlook exposes either Unicode or ANSI headers depending on the item.')]
  param (
    [object]
    $Item
  )

  $_propertyAccessor = $null
  try {
    $_propertyAccessor = $Item.PropertyAccessor
    foreach ($_tag in $script:HDR_TAGS) {
      try {
        $_headers = $_propertyAccessor.GetProperty($_tag)
        if (-not [string]::IsNullOrWhiteSpace($_headers)) {
          $_messageId = Get-TransportMessageId -HeaderText ([string]$_headers)
          if (-not [string]::IsNullOrWhiteSpace($_messageId)) {
            return $_messageId
          }
        }
      }
      catch { }
    }
  }
  catch {
    return $null
  }
  finally {
    Remove-ComObject $_propertyAccessor
  }

  return $null
}

function Add-OutlookItemResult {
  param (
    [System.Collections.IList]
    $Results,

    [string]
    $Action,

    [string]
    $Status,

    [string]
    $Folder,

    [object]
    $Item,

    [string]
    $MessageId,

    [string]
    $Detail
  )

  $_subject = ''
  $_received = $null
  try { $_subject = [string]$Item.Subject } catch { $_subject = '' }
  try { $_received = $Item.ReceivedTime } catch { $_received = $null }

  $_property = @{
    Received  = $_received
    MessageId = $MessageId
  }

  Add-OperationResult `
    -Results $Results `
    -Target $_subject `
    -Source 'Outlook' `
    -Scope $Folder `
    -Action $Action `
    -Status $Status `
    -Detail $Detail `
    -Property $_property
}

function Optimize-OutlookFolder {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $Folder,

    [object]
    $ReviewFolder,

    [System.Collections.IList]
    $Results
  )

  $_folderPath = $Folder.FolderPath
  Write-Progress -Id 21 -ParentId 20 -Activity 'Inspecting messages and processing duplicates' -Status $_folderPath -PercentComplete -1
  $_items = $Folder.Items
  $_seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
  $_progressTimer = [Diagnostics.Stopwatch]::StartNew()

  try {
    $_itemCount = $_items.Count
    for ($_index = $_itemCount; $_index -ge 1; $_index--) {
      if ($_index -eq $_itemCount -or $_progressTimer.ElapsedMilliseconds -ge 200) {
        $_processed = $_itemCount - $_index
        $_progress = @{
          Id               = 21
          ParentId         = 20
          Activity         = 'Inspecting messages and processing duplicates'
          Status           = $_folderPath
          CurrentOperation = "Read headers and compare Message-IDs | $_processed / $_itemCount inspected"
          PercentComplete  = [int](100.0 * $_processed / $_itemCount)
        }

        Write-Progress @_progress
        $_progressTimer.Restart()
      }

      $_item = $null
      $_movedItem = $null
      try {
        $_item = $_items.Item($_index)
        if ($_item.Class -ne $script:OL_MAIL) { continue }

        $_messageId = Get-MessageId -Item $_item
        if (-not $_messageId) {
          Add-OutlookItemResult -Results $Results -Action 'Deduplicate' -Status 'Skipped' -Folder $_folderPath -Item $_item -MessageId '' -Detail 'NoMessageId'
          continue
        }

        if ($_seen.Contains($_messageId)) {
          $_metadata = [PSCustomObject]@{ Subject = $_item.Subject; ReceivedTime = $_item.ReceivedTime }
          $_reviewPath = if ($ReviewFolder) { $ReviewFolder.FolderPath } else { $ReviewFolderName }
          if ($PSCmdlet.ShouldProcess("$_folderPath | $([string]$_item.Subject)", "Move duplicate to $_reviewPath")) {
            $_movedItem = $_item.Move($ReviewFolder)
            Add-OutlookItemResult -Results $Results -Action 'MoveDuplicate' -Status 'Moved' -Folder $_folderPath -Item $_metadata -MessageId $_messageId -Detail "Duplicate moved to $_reviewPath"
          }
          else {
            $_detail = if ($WhatIfPreference) { 'DryRun' } else { 'Declined' }
            Add-OutlookItemResult -Results $Results -Action 'MoveDuplicate' -Status 'Skipped' -Folder $_folderPath -Item $_metadata -MessageId $_messageId -Detail $_detail
          }
        }
        else {
          $null = $_seen.Add($_messageId)
          Add-OutlookItemResult -Results $Results -Action 'Deduplicate' -Status 'Kept' -Folder $_folderPath -Item $_item -MessageId $_messageId -Detail 'First item with Message-ID in folder.'
        }
      }
      catch {
        Add-OperationResult -Results $Results -Target $_folderPath -Source 'Outlook' -Action 'Deduplicate' -Status 'Failed' -Detail $_.Exception.Message
        throw "Deduplication stopped at item $_index in '$_folderPath': $($_.Exception.Message)"
      }
      finally {
        Remove-ComObject $_movedItem $_item
      }
    }
  }
  finally {
    Write-Progress -Id 21 -Activity 'Inspecting messages and processing duplicates' -Completed
    Remove-ComObject $_items
  }
}

function Invoke-OutlookFolderTree {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $Folder,

    [object]
    $ReviewFolder,

    [string]
    $ReviewName,

    [string[]]
    $Exclude,

    [System.Collections.IList]
    $Results
  )

  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Inspecting folder and applying exclusions' -CurrentOperation $Folder.FolderPath -PercentComplete -1

  if ($Folder.Name -eq $ReviewName) { return }

  if ($Exclude -contains $Folder.Name) {
    Write-Verbose "Skipping excluded folder: $($Folder.Name)"
    return
  }
  else {
    Write-Verbose "Processing: $($Folder.FolderPath)"
    $_accessor = $Folder.PropertyAccessor
    try {
      if ($_accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003') -eq 2) { return }
    }
    finally { Remove-ComObject $_accessor }
    if ($Folder.DefaultItemType -ne 0) { return }
    Optimize-OutlookFolder -Folder $Folder -ReviewFolder $ReviewFolder -Results $Results -WhatIf:$WhatIfPreference -Confirm:$false
  }

  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Enumerating subfolders' -CurrentOperation $Folder.FolderPath -PercentComplete -1
  $_folders = $Folder.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_child = $_folders.Item($_index)
      try {
        Invoke-OutlookFolderTree -Folder $_child -ReviewFolder $ReviewFolder -ReviewName $ReviewName -Exclude $Exclude -Results $Results -WhatIf:$WhatIfPreference -Confirm:$false
      }
      finally {
        Remove-ComObject $_child
      }
    }
  }
  finally {
    Remove-ComObject $_folders
  }
}

$_context = $null
$_storeRoot = $null
$_reviewFolder = $null

try {
  Write-Log -Message 'Connecting to Outlook and locating the store for duplicate review...' -Color Cyan
  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Connecting to Outlook' -PercentComplete -1
  $_context = Connect-Outlook
  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Reading Outlook stores' -PercentComplete -1

  $_outlookMajor = [int](($_context.App.Version -split '\.')[0])
  if ($_outlookMajor -lt 12) {
    throw "Outlook 2007 (version 12) or later is required. Detected Outlook version: $($_context.App.Version)"
  }

  # PSFoundation 1.3.0 checks Store.IsDefault, which Outlook does not expose.
  # Resolve the default via Namespace.DefaultStore and reject ambiguous names.
  if ([string]::IsNullOrWhiteSpace($StoreName)) {
    $_selectedStore = $_context.Namespace.DefaultStore
    try { $_storeRoot = $_selectedStore.GetRootFolder() }
    finally { Remove-ComObject $_selectedStore }
  }
  else {
    $_stores = $_context.Namespace.Stores
    $_matches = 0
    try {
      for ($_storeIndex = 1; $_storeIndex -le $_stores.Count; $_storeIndex++) {
        $_store = $_stores.Item($_storeIndex)
        try {
          Write-Verbose "Store: $($_store.DisplayName) | $($_store.FilePath)"
          if ($_store.DisplayName -eq $StoreName) { $_matches++ }
        }
        finally { Remove-ComObject $_store }
      }
    }
    finally { Remove-ComObject $_stores }
    if ($_matches -ne 1) { throw "StoreName '$StoreName' matches $_matches stores. Use a unique display name." }
    $_storeRoot = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $StoreName
  }
  Write-Verbose "Store root: $($_storeRoot.FolderPath)"

  if (-not $WhatIfPreference) {
    if (-not $PSCmdlet.ShouldProcess($_storeRoot.FolderPath, "Move suspected duplicates to '$ReviewFolderName' for review")) { return }
    Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Preparing duplicate review folder' -CurrentOperation $ReviewFolderName -PercentComplete -1
    $_reviewFolder = Get-OutlookSubFolder -ParentFolder $_storeRoot -Name $ReviewFolderName -Create
  }

  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Enumerating source folders' -CurrentOperation $_storeRoot.FolderPath -PercentComplete -1
  $_folders = $_storeRoot.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_child = $_folders.Item($_index)
      try {
        Invoke-OutlookFolderTree -Folder $_child -ReviewFolder $_reviewFolder -ReviewName $ReviewFolderName -Exclude $ExcludeFolders -Results $_results -WhatIf:$WhatIfPreference -Confirm:$false
      }
      finally {
        Remove-ComObject $_child
      }
    }
  }
  finally {
    Remove-ComObject $_folders
  }
}
catch {
  Add-OperationResult -Results $_results -Target $StoreName -Source 'Outlook' -Action 'Deduplicate' -Status 'Failed' -Detail $_.Exception.Message
  Write-Warning $_.Exception.Message
}
finally {
  Write-Progress -Id 21 -Activity 'Inspecting messages and processing duplicates' -Completed
  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Completed
  Remove-ComObject $_reviewFolder
  Remove-ComObject $_storeRoot

  if ($_context) {
    try {
      if ($QuitOutlook -and -not $WhatIfPreference) {
        $_context.App.Quit()
      }
    }
    catch {
      Write-Verbose "Could not quit Outlook: $($_.Exception.Message)"
    }

    Remove-ComObject $_context.Namespace $_context.App
  }

  Invoke-ComGarbageCollection
}

$_moved = @($_results | Where-Object { $_.Action -eq 'MoveDuplicate' -and $_.Status -eq 'Moved' }).Count
$_planned = @($_results | Where-Object { $_.Action -eq 'MoveDuplicate' -and $_.Detail -eq 'DryRun' }).Count
$_kept = @($_results | Where-Object { $_.Status -eq 'Kept' }).Count
$_skipped = @($_results | Where-Object { $_.Detail -eq 'NoMessageId' }).Count
$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count

if ($WhatIfPreference) {
  Write-Log -Message "Outlook deduplication preview complete. Duplicates to move: $_planned | Kept: $_kept | Skipped: $_skipped | Failed: $_failed" -Color Yellow
}
else {
  Write-Log -Message "Outlook deduplication complete. Moved: $_moved | Kept: $_kept | Skipped: $_skipped | Failed: $_failed" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
}

try {
  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Writing reports and operation log' -PercentComplete -1
  if ($ReportPath) {
    $_reportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
    $_reportRoot = Split-Path -Path $_reportPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($_reportRoot) -and -not (Test-Path -LiteralPath $_reportRoot)) {
      $null = New-Item -Path $_reportRoot -ItemType Directory -Force
    }

    $_results | Export-Csv -Path $_reportPath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Report: $_reportPath" -Color Gray
  }

  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Optimize-Outlook'
  if ($_operationLog) {
    Write-Log -Message "Operation log: $_operationLog" -Color Gray
  }
}
finally {
  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Completed
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failed -gt 0) {
  exit 1
}
