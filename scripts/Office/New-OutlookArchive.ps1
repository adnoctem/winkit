#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Archives Outlook mail into a standalone Unicode PST.
.DESCRIPTION
  Adds a new PST store to the current Outlook profile, mirrors the source
  store's mail-folder hierarchy, and copies or moves mail items into it. Optional
  received-date bounds limit which mail items are archived. When finished, the
  PST can be detached. Close Outlook before copying the PST file elsewhere.
  Requires a NEW local PST path; reruns must use a different path to avoid
  duplicate archives. Contacts, calendars, tasks and search folders are skipped.
  Copy temporarily duplicates each message in its SOURCE store before moving
  the duplicate to the archive. Keep a closed-file backup and adequate headroom.
.PARAMETER ArchivePath
  Full path of a new local .pst file. Existing files are refused.
.PARAMETER StoreName
  Display name of the source Outlook store. If omitted, the default delivery
  store is used.
.PARAMETER StartDate
  Optional inclusive received-date lower bound.
.PARAMETER EndDate
  Optional inclusive received-date upper bound, including the exact time.
  A date alone means midnight, not the end of that day. Prefer EndBefore.
.PARAMETER EndBefore
  Optional exclusive received-date upper bound, for example 2025-01-01 to
  include all of 2024. Cannot be combined with EndDate.
.PARAMETER Mode
  Copy leaves source mail intact. Move removes archived items from the source.
.PARAMETER DisplayName
  Display name for the mounted PST store while it is attached.
.PARAMETER DetachWhenDone
  Remove the PST store from the profile at the end.
.PARAMETER DryRun
  Preview changes without copying or moving messages.
.PARAMETER PassThru
  Return structured operation result objects.
.PARAMETER QuitOutlook
  Quit the Outlook application object on exit. Leave off if Outlook was already
  open interactively.
.EXAMPLE
  PS> .\New-OutlookArchive.ps1 -ArchivePath D:\Backups\user-snapshot.pst -StoreName 'user@example.com' -Mode Copy
.EXAMPLE
  PS> .\New-OutlookArchive.ps1 -ArchivePath D:\Archive\user-2025.pst -StoreName 'user@example.com' -StartDate '2025-01-01' -EndBefore '2026-01-01' -Mode Move
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not applicable - Outlook is a desktop client.
  SYSTEM-account execution: not applicable - requires an interactive Outlook profile.
  Outlook version: 2007 (version 12) or later - AddStoreEx creates Unicode PSTs.
  Bitness: Outlook 2007 is 32-bit only - run under 32-bit PowerShell (x86).
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Parameter(Mandatory = $true)]
  [string]
  $ArchivePath,

  [string]
  $StoreName,

  [datetime]
  $StartDate,

  [datetime]
  $EndDate,

  [datetime]
  $EndBefore,

  [ValidateSet('Copy', 'Move')]
  [string]
  $Mode = 'Copy',

  [string]
  $DisplayName = 'Archive',

  [bool]
  $DetachWhenDone = $true,

  [switch]
  $DryRun,

  [switch]
  $PassThru,

  [switch]
  $QuitOutlook
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no Outlook messages will be archived`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$script:OutlookArchiveCopied = 0
$script:OutlookArchiveMoved = 0

function Test-OutlookItemInRange {
  param (
    [object]
    $Item
  )

  if ($Item.Class -ne 43) { return $false }
  if (-not $StartDate -and -not $EndDate -and -not $EndBefore) { return $true }

  $_receivedTime = $null
  try {
    $_receivedTime = $Item.ReceivedTime
  }
  catch {
    throw 'Cannot read ReceivedTime for a mail item; archive stopped.'
  }

  if (-not $_receivedTime) { throw 'Mail item has no ReceivedTime; archive stopped.' }

  if ($StartDate -and $_receivedTime -lt $StartDate) { return $false }
  if ($EndDate -and $_receivedTime -gt $EndDate) { return $false }
  if ($EndBefore -and $_receivedTime -ge $EndBefore) { return $false }

  return $true
}

function Add-OutlookArchiveResult {
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
    $Detail
  )

  $_subject = ''
  $_received = $null
  try { $_subject = [string]$Item.Subject } catch { $_subject = '' }
  try { $_received = $Item.ReceivedTime } catch { $_received = $null }

  $_property = @{
    Received = $_received
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

function Copy-OutlookFolderItem {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $SourceFolder,

    [object]
    $DestinationFolder,

    [string]
    $ArchiveMode,

    [System.Collections.IList]
    $Results
  )

  # Snapshot identifiers before Copy() changes the source Items collection.
  # Never follow a growing/reordered live collection or retain every COM item.
  $_ids = New-Object 'System.Collections.Generic.List[string]'
  $_items = $SourceFolder.Items
  try {
    for ($_index = 1; $_index -le $_items.Count; $_index++) {
      $_item = $null
      try {
        $_item = $_items.Item($_index)
        if (Test-OutlookItemInRange -Item $_item) {
          if (-not $_item.EntryID) { throw 'Mail item has no EntryID.' }
          $_ids.Add([string]$_item.EntryID)
        }
      }
      finally { Remove-ComObject $_item }
    }
  }
  finally { Remove-ComObject $_items }

  foreach ($_id in $_ids) {
    $_item = $null
    $_copy = $null
    $_moved = $null
    try {
      $_item = $_context.Namespace.GetItemFromID($_id, $SourceFolder.StoreID)
      # Move invalidates the original COM item. Capture audit fields first.
      $_metadata = [PSCustomObject]@{ Subject = $_item.Subject; ReceivedTime = $_item.ReceivedTime }
      $_destinationPath = if ($DestinationFolder) { $DestinationFolder.FolderPath } else { $_archivePath }
      if ($PSCmdlet.ShouldProcess("$($SourceFolder.FolderPath) | $($_metadata.Subject)", "$ArchiveMode to $_destinationPath")) {
        if ($ArchiveMode -eq 'Move') {
          $_moved = $_item.Move($DestinationFolder)
          $script:OutlookArchiveMoved++
          $_status = 'Moved'
        }
        else {
          $_copy = $_item.Copy()
          $_moved = $_copy.Move($DestinationFolder)
          $script:OutlookArchiveCopied++
          $_status = 'Copied'
        }
        Add-OutlookArchiveResult -Results $Results -Action $ArchiveMode -Status $_status -Folder $SourceFolder.FolderPath -Item $_metadata -Detail "$_status to $_destinationPath"
      }
      else {
        $_detail = if ($WhatIfPreference) { 'DryRun' } else { 'Declined' }
        Add-OutlookArchiveResult -Results $Results -Action $ArchiveMode -Status 'Skipped' -Folder $SourceFolder.FolderPath -Item $_metadata -Detail $_detail
      }
    }
    catch {
      # A failed Copy().Move() can leave a duplicate in the source. Stop instead
      # of repeatedly filling a near-limit PST. Do not delete uncertain items.
      throw "Archive stopped in '$($SourceFolder.FolderPath)' at EntryID '$_id': $($_.Exception.Message). A failed copy may remain in the source; inspect before retrying."
    }
    finally { Remove-ComObject $_moved $_copy $_item }
  }
}
function Copy-OutlookFolderTree {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $SourceFolder,

    [object]
    $DestinationParent,

    [string]
    $ArchiveMode,

    [System.Collections.IList]
    $Results
  )

  # Search folders are virtual views; processing them would archive mail twice.
  $_accessor = $SourceFolder.PropertyAccessor
  try {
    if ($_accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003') -eq 2) { return }
  }
  finally { Remove-ComObject $_accessor }
  if ($SourceFolder.DefaultItemType -ne 0) {
    Write-Verbose "Skipping non-mail folder: $($SourceFolder.FolderPath)"
    return
  }

  Write-Verbose "Archiving: $($SourceFolder.FolderPath)"
  $_destinationFolder = if ($WhatIfPreference) {
    $null
  }
  else {
    Get-OutlookSubFolder -ParentFolder $DestinationParent -Name $SourceFolder.Name -Create
  }

  try {
    Copy-OutlookFolderItem -SourceFolder $SourceFolder -DestinationFolder $_destinationFolder -ArchiveMode $ArchiveMode -Results $Results -WhatIf:$WhatIfPreference -Confirm:$false

    $_folders = $SourceFolder.Folders
    try {
      for ($_index = 1; $_index -le $_folders.Count; $_index++) {
        $_child = $_folders.Item($_index)
        try {
          Copy-OutlookFolderTree -SourceFolder $_child -DestinationParent $_destinationFolder -ArchiveMode $ArchiveMode -Results $Results -WhatIf:$WhatIfPreference -Confirm:$false
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
  finally {
    Remove-ComObject $_destinationFolder
  }
}

$_context = $null
$_sourceRoot = $null
$_archiveRoot = $null
$_archiveOwned = $false
$_archivePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArchivePath)

try {
  if ([IO.Path]::GetExtension($_archivePath) -ne '.pst') { throw 'ArchivePath must end in .pst.' }
  if (Test-Path -LiteralPath $_archivePath) { throw 'ArchivePath already exists. Use a new PST path for every run.' }
  $_drive = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($_archivePath))
  if ($_archivePath.StartsWith('\\') -or $_drive.DriveType -eq 'Network') { throw 'ArchivePath must be on a local drive.' }
  if (-not (Test-Path -LiteralPath (Split-Path -Parent $_archivePath) -PathType Container)) { throw 'Archive parent directory must already exist.' }
  if ($EndDate -and $EndBefore) { throw 'Use EndDate or EndBefore, not both.' }
  if ($StartDate -and $EndDate -and $StartDate -gt $EndDate) {
    throw 'StartDate is after EndDate.'
  }
  if ($StartDate -and $EndBefore -and $StartDate -ge $EndBefore) { throw 'StartDate must be before EndBefore.' }

  $_context = Connect-Outlook

  $_outlookMajor = [int](($_context.App.Version -split '\.')[0])
  if ($_outlookMajor -lt 12) {
    throw "Outlook 2007 (version 12) or later is required. Detected Outlook version: $($_context.App.Version)"
  }

  # PSFoundation 1.3.0 checks Store.IsDefault, which Outlook does not expose.
  # Resolve the default via Namespace.DefaultStore and reject ambiguous names.
  if ([string]::IsNullOrWhiteSpace($StoreName)) {
    $_selectedStore = $_context.Namespace.DefaultStore
    try { $_sourceRoot = $_selectedStore.GetRootFolder() }
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
    $_sourceRoot = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $StoreName
  }
  Write-Verbose "Source: $($_sourceRoot.FolderPath)"

  if (-not $WhatIfPreference) {
    if (-not $PSCmdlet.ShouldProcess("$($_sourceRoot.FolderPath) -> $_archivePath", "Archive mail ($Mode) to a NEW PST")) { return }
    $_archiveRoot = Add-OutlookStoreRoot -Namespace $_context.Namespace -Path $_archivePath
    if ($_archiveRoot.StoreID -eq $_sourceRoot.StoreID) { throw 'Source and archive must be different stores.' }
    $_archiveOwned = $true
    try {
      $_archiveRoot.Name = $DisplayName
    }
    catch {
      Write-Verbose "Could not rename archive PST store to '$DisplayName': $($_.Exception.Message)"
    }
  }

  Write-Verbose "Archive PST: $_archivePath"

  Copy-OutlookFolderItem -SourceFolder $_sourceRoot -DestinationFolder $_archiveRoot -ArchiveMode $Mode -Results $_results -WhatIf:$WhatIfPreference -Confirm:$false

  $_folders = $_sourceRoot.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_child = $_folders.Item($_index)
      try {
        Copy-OutlookFolderTree -SourceFolder $_child -DestinationParent $_archiveRoot -ArchiveMode $Mode -Results $_results -WhatIf:$WhatIfPreference -Confirm:$false
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
  Add-OperationResult -Results $_results -Target $_archivePath -Source 'Outlook' -Action 'Archive' -Status 'Failed' -Detail $_.Exception.Message
  Write-Warning $_.Exception.Message
}
finally {
  if ($DetachWhenDone -and $_archiveOwned -and $_context -and $_archiveRoot) {
    try {
      $_context.Namespace.RemoveStore($_archiveRoot)
    }
    catch {
      Add-OperationResult -Results $_results -Target $_archivePath -Source 'Outlook' -Action 'DetachStore' -Status 'Failed' -Detail $_.Exception.Message
      Write-Warning "Could not detach PST: $($_.Exception.Message)"
    }
  }

  Remove-ComObject $_archiveRoot $_sourceRoot

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

$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
if ($WhatIfPreference) {
  $_planned = @($_results | Where-Object { $_.Detail -eq 'DryRun' }).Count
  Write-Log -Message "Outlook archive preview complete. Items to process: $_planned | Failed: $_failed | PST: $_archivePath" -Color Yellow
}
else {
  Write-Log -Message "Outlook archive complete. Copied: $script:OutlookArchiveCopied | Moved: $script:OutlookArchiveMoved | Failed: $_failed | PST: $_archivePath" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'New-OutlookArchive'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failed -gt 0) {
  exit 1
}
