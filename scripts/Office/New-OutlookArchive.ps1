#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.7.1' }

<#
.SYNOPSIS
  Archives Outlook mail into a standalone Unicode PST.
.DESCRIPTION
  Creates or opens a PST store in the current Outlook profile, mirrors selected source
  folder paths, and copies or moves mail items into it. Defaults to Inbox only;
  Recurse explicitly includes descendants. Standard folders other than Inbox
  require their Include switch. Custom Exclusions always take precedence. Optional
  received-date bounds limit which mail items are archived. When finished, the
  PST can be detached. Close Outlook before copying the PST file elsewhere.
  Creates a new local PST unless Append explicitly selects an existing file.
  Append does not deduplicate messages; repeating Copy can create duplicates.
  SkipPathPreservation flattens selected mail into the archive root.
  Non-mail items and search folders are always skipped.
  Included non-mail containers allow traversal to their mail subfolders.
  Copy temporarily duplicates each message in its SOURCE store before moving
  the duplicate to the archive. Keep a closed-file backup and adequate headroom.
  Attached IMAP/Exchange stores are supported with classic Outlook 2010 or later.
  Move can remove source mail from the server and other synchronized clients;
  Copy also performs temporary source writes that can synchronize.
  Matching OST messages must be fully downloaded, including during previews.
  Only mail exposed by Outlook is considered; server completeness is not verified.
  Shows folder and item progress and writes a JSON report, including previews.
  Per-message results are stored in the report rather than printed to the console.
.PARAMETER ArchivePath
  Full path of a local .pst file. Existing files require Append.
.PARAMETER Append
  Add mail to an existing PST. A missing file is refused. Existing attachments
  and display names are preserved unless a new display name is supplied.
  This is not an idempotent retry mode; overlapping Copy passes add duplicates.
.PARAMETER SkipPathPreservation
  Put all selected mail directly in the archive root instead of recreating
  source folder paths. With Recurse, mail from multiple folders is combined.
.PARAMETER IncludeSentItems
  Permit the standard SentItems folder within the selected scope. Excluded unless explicitly included.
.PARAMETER IncludeDeletedItems
  Permit the standard DeletedItems folder within the selected scope. Excluded unless explicitly included.
.PARAMETER IncludeJunk
  Permit the standard Junk folder within the selected scope. Excluded unless explicitly included.
.PARAMETER IncludeOutbox
  Permit the standard Outbox folder within the selected scope. Excluded unless explicitly included.
.PARAMETER IncludeDrafts
  Permit the standard Drafts folder within the selected scope. Excluded unless explicitly included.
.PARAMETER IncludeMedia
  Permit Calendar, Contacts, Journal, Notes, Tasks, AllPublicFolders, RssFeeds,
  ToDo, ManagedEmail, and SuggestedContacts within the selected scope. Only mail
  is processed; Recurse permits traversal through included non-mail containers.
.PARAMETER IncludeFailures
  Permit SyncIssues, Conflicts, LocalFailures, and ServerFailures within the
  selected scope. Search folders remain excluded.
.PARAMETER Exclusions
  Exact store-relative paths excluded with their descendants. Exclusions win over Include switches.
.PARAMETER StoreName
  Display name of the source Outlook store. If omitted, the default delivery
  store is used. Select an attached PST, IMAP, or Exchange store, not an OST path.
.PARAMETER FolderName
  Exact folder path relative to the selected store, or empty for the store root.
  When omitted, selects Inbox by identity, including localized or renamed Inboxes.
  Explicit Inbox or Posteingang selects that literal path; no name fallback is used.
.PARAMETER Recurse
  Include descendants of the selected folder. Otherwise only its own mail is processed.
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
  For IMAP/Exchange, removal can synchronize to the server and other clients.
  Copy temporarily creates duplicates in the source before transferring them.
.PARAMETER DisplayName
  Display name for the PST. New files default to the filename without .pst.
  Append preserves the existing name unless this parameter is supplied.
  DataFileName is an alias.
.PARAMETER AddDataFile
  Keep a PST attached when this run added it to the current Outlook profile.
  Cannot be combined with DetachWhenDone set to true.
.PARAMETER DetachWhenDone
  Remove an attachment created by this run at the end. An already attached
  archive is left open; explicitly requesting its detachment is rejected.
.PARAMETER DryRun
  Preview changes without copying or moving messages.
.PARAMETER Sort
  Order report results by received date: NewToOld (default) is newest first;
  OldToNew is oldest first. Undated records follow dated messages. Equal dates
  retain processing order. This affects reports only, not message processing.
.PARAMETER PassThru
  Return one summary with counts and ReportPath. Per-message results are in
  the JSON report's Results array. DryRun and WhatIf also return this summary.
.PARAMETER ReportDirectory
  Directory for uniquely named JSON reports. Defaults to
  %LOCALAPPDATA%\winkit\reports\Outlook for the user running the script.
  Created when needed, including during previews. Cannot be combined with ReportPath.
.PARAMETER ReportPath
  Exact filename for the JSON report. Relative paths resolve from the current
  PowerShell location. Parent directories are created when needed, including
  during previews. Existing files are refused. Cannot be combined with ReportDirectory.
.PARAMETER QuitOutlook
  Quit the Outlook application object on exit. Leave off if Outlook was already
  open interactively.
.PARAMETER IgnoreAdministrator
  Allow an elevated PowerShell session. Use only when Outlook intentionally
  runs elevated under the same Windows user. Does not switch users or profiles.
  Elevated execution is otherwise refused, including during previews.
.EXAMPLE
  PS> .\New-OutlookArchive.ps1 -ArchivePath D:\Backups\user-snapshot.pst -StoreName 'user@example.com' -Mode Copy
.EXAMPLE
  PS> .\New-OutlookArchive.ps1 -ArchivePath D:\Archive\user-2025.pst -StoreName 'user@example.com' -StartDate '2025-01-01' -EndBefore '2026-01-01' -Mode Move
.EXAMPLE
  PS> .\New-OutlookArchive.ps1 -ArchivePath D:\Archive\mail.pst -DryRun -ReportPath .\archive-report.json
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not applicable - Outlook is a desktop client.
  SYSTEM-account execution: not applicable - requires an interactive Outlook profile.
  Outlook version: 2007 for PST sources; 2010 or later for OST/Exchange sources.
  Bitness: Outlook COM automation supports cross-architecture PowerShell clients.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Parameter(Mandatory = $true)]
  [string]
  $ArchivePath,

  [switch]
  $Append,

  [switch]
  $SkipPathPreservation,

  [string]
  $StoreName,

  [Alias('IncludeSentMail')]
  [switch]
  $IncludeSentItems,

  [switch]
  $IncludeDeletedItems,

  [switch]
  $IncludeJunk,

  [switch]
  $IncludeOutbox,

  [switch]
  $IncludeDrafts,

  [switch]
  $IncludeMedia,

  [switch]
  $IncludeFailures,

  [Alias('ExcludeFolders')]
  [string[]]
  $Exclusions = @(),

  [AllowEmptyString()]
  [string]
  $FolderName = 'Inbox',

  [switch]
  $Recurse,

  [datetime]
  $StartDate,

  [datetime]
  $EndDate,

  [datetime]
  $EndBefore,

  [ValidateSet('Copy', 'Move')]
  [string]
  $Mode = 'Copy',

  [Alias('DataFileName')]
  [ValidateNotNullOrEmpty()]
  [string]
  $DisplayName,

  [switch]
  $AddDataFile,

  [bool]
  $DetachWhenDone = $true,

  [switch]
  $DryRun,

  [ValidateSet('OldToNew', 'NewToOld')]
  [string]
  $Sort = 'NewToOld',

  [switch]
  $PassThru,

  [string]
  $ReportDirectory,

  [ValidateNotNullOrEmpty()]
  [string]
  $ReportPath,

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
  Write-Log -Message "DRY RUN - no Outlook messages will be archived`n" -Color Yellow
}

# Script options expand to the identity-based kinds accepted by PSFoundation.
# Inbox is always permitted; Exclusions can still exclude its exact path.
$_inclusionGroups = [ordered]@{
  IncludeSentItems    = @('SentItems')
  IncludeDeletedItems = @('DeletedItems')
  IncludeJunk         = @('Junk')
  IncludeDrafts       = @('Drafts')
  IncludeOutbox       = @('Outbox')
  IncludeMedia        = @(
    'Calendar',
    'Contacts',
    'Journal',
    'Notes',
    'Tasks',
    'AllPublicFolders',
    'RssFeeds',
    'ToDo',
    'ManagedEmail',
    'SuggestedContacts'
  )
  IncludeFailures     = @(
    'SyncIssues',
    'Conflicts',
    'LocalFailures',
    'ServerFailures'
  )
}

$_results = New-Object System.Collections.ArrayList
$script:OutlookArchiveCopied = 0
$script:OutlookArchiveMoved = 0
$script:OutlookArchiveFoldersRead = 0
$script:OutlookArchiveFoldersSkipped = 0
$script:OutlookArchiveItemsRead = 0
$script:OutlookArchiveItemsMatched = 0
$script:OutlookArchiveProgressTimer = [Diagnostics.Stopwatch]::StartNew()
$_startedAt = Get-Date

function Write-OutlookArchiveProgress {
  param (
    [string]
    $Phase,

    [string]
    $Folder,

    [int]
    $Current,

    [int]
    $Total,

    [switch]
    $Force
  )

  # Avoid repainting conhost for every COM item in a large folder.
  if (-not $Force -and $script:OutlookArchiveProgressTimer.ElapsedMilliseconds -lt 200) {
    return
  }

  $script:OutlookArchiveProgressTimer.Restart()
  $_percent = if ($Total -gt 0) {
    [int][math]::Min(100, (100.0 * $Current / $Total))
  }
  else {
    -1
  }

  $_progress = @{
    Id               = 0
    Activity         = 'Outlook archive'
    Status           = $Phase
    CurrentOperation = "$Folder | $Current / $Total | Read: $script:OutlookArchiveItemsRead | Matched: $script:OutlookArchiveItemsMatched"
    PercentComplete  = $_percent
  }

  Write-Progress @_progress
}

function Test-OutlookItemInRange {
  param (
    [object]
    $Item
  )

  if ($Item.Class -ne 43) {
    return $false
  }

  if (-not $StartDate -and -not $EndDate -and -not $EndBefore) {
    return $true
  }

  $_receivedTime = $null
  try {
    $_receivedTime = $Item.ReceivedTime
  }
  catch {
    throw 'Cannot read ReceivedTime for a mail item; archive stopped.'
  }

  if (-not $_receivedTime) {
    throw 'Mail item has no ReceivedTime; archive stopped.'
  }

  if ($StartDate -and $_receivedTime -lt $StartDate) {
    return $false
  }

  if ($EndDate -and $_receivedTime -gt $EndDate) {
    return $false
  }

  if ($EndBefore -and $_receivedTime -ge $EndBefore) {
    return $false
  }

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
    $Detail,

    [string]
    $DestinationFolderPath,

    [string]
    $SourceEntryID,

    [string]
    $SourceStoreID
  )

  $_subject = ''
  $_received = $null
  try {
    $_subject = [string]$Item.Subject
  }
  catch {
    $_subject = ''
  }

  if ([string]::IsNullOrWhiteSpace($_subject)) {
    $_subject = '<No Subject>'
  }

  try {
    if ($Item.ReceivedTime) {
      $_received = ([datetime]$Item.ReceivedTime).ToString('o')
    }
  }
  catch {
    $_received = $null
  }

  $_property = @{
    Received              = $_received
    SourceFolderPath      = $Folder
    DestinationFolderPath = $DestinationFolderPath
    SourceEntryID         = $SourceEntryID
    SourceStoreID         = $SourceStoreID
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

function Get-OutlookArchiveSourceInfo {
  param (
    [object]
    $Store
  )

  $_path = [string]$Store.FilePath
  $_extension = [IO.Path]::GetExtension($_path)
  $_format = switch ($_extension) {
    '.pst' {
      'PST'
    }
    '.ost' {
      'OST'
    }
    '' {
      'None'
    }
    default {
      'Other'
    }
  }
  $_exchangeType = $Store.ExchangeStoreType
  if ($null -eq $_exchangeType) {
    throw 'Cannot determine the source ExchangeStoreType. No archive work was started.'
  }

  # olNotExchange (3) also describes non-Exchange OST providers. An OST
  # extension alone cannot distinguish IMAP from another synchronized provider.
  [PSCustomObject]@{
    DisplayName       = [string]$Store.DisplayName
    StoreID           = [string]$Store.StoreID
    FilePath          = $_path
    DataFileFormat    = $_format
    ExchangeStoreType = [int]$_exchangeType
    MaySynchronize    = $_format -eq 'OST' -or [int]$_exchangeType -ne 3
  }
}

function Test-OutlookArchiveDownload {
  [OutputType([bool])]
  param (
    [object]
    $Item,

    [string]
    $Folder
  )

  $_identity = "folder '$Folder', EntryID '$($Item.EntryID)'"
  try {
    $_state = $Item.DownloadState
  }
  catch {
    throw "Cannot read DownloadState for $_identity. Download full messages in Outlook and retry after synchronization: $($_.Exception.Message)"
  }

  # OlDownloadState: 0 = headers only, 1 = full item. Missing or unknown values
  # must not be cast to zero or silently treated as a complete message.
  if ($null -eq $_state -or $_state -ne 1) {
    throw "Message is not confirmed fully downloaded in $_identity (DownloadState='$($_state)'). Download full messages, including bodies and attachments, in Outlook before retrying."
  }

  return $true
}

function Copy-OutlookFolderItem {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $SourceFolder,

    [object]
    $DestinationFolder,

    [string]
    $DestinationRelativePath,

    [string]
    $ArchiveMode,

    [switch]
    $RequireFullDownload,

    [System.Collections.IList]
    $Results
  )

  # Snapshot identifiers before Copy() changes the source Items collection.
  # Never follow a growing/reordered live collection or retain every COM item.
  $_ids = New-Object 'System.Collections.Generic.List[string]'
  $_folderPath = [string]$SourceFolder.FolderPath
  $script:OutlookArchiveFoldersRead++
  Write-OutlookArchiveProgress -Phase 'Reading folder contents and applying date filters' -Folder $_folderPath -Force
  $_items = $SourceFolder.Items
  try {
    $_itemCount = $_items.Count
    for ($_index = 1; $_index -le $_itemCount; $_index++) {
      $_item = $null
      try {
        $_item = $_items.Item($_index)
        $script:OutlookArchiveItemsRead++
        if (Test-OutlookItemInRange -Item $_item) {
          if (-not $_item.EntryID) {
            throw 'Mail item has no EntryID.'
          }

          if ($RequireFullDownload) {
            $null = Test-OutlookArchiveDownload -Item $_item -Folder $_folderPath
          }

          $_ids.Add([string]$_item.EntryID)
          $script:OutlookArchiveItemsMatched++
        }

        Write-OutlookArchiveProgress -Phase 'Reading folder contents and applying date filters' -Folder $_folderPath -Current $_index -Total $_itemCount
      }
      finally {
        Remove-ComObject $_item
      }
    }
  }
  finally {
    Remove-ComObject $_items
  }

  $_phase = if ($WhatIfPreference) { 'Recording preview results' } else { "$ArchiveMode matching messages" }
  $_processed = 0
  Write-OutlookArchiveProgress -Phase $_phase -Folder $_folderPath -Total $_ids.Count -Force

  foreach ($_id in $_ids) {
    $_item = $null
    $_copy = $null
    $_moved = $null
    try {
      $_item = $_context.Namespace.GetItemFromID($_id, $SourceFolder.StoreID)
      if ($RequireFullDownload) {
        $null = Test-OutlookArchiveDownload -Item $_item -Folder $_folderPath
      }

      # Move invalidates the original COM item. Capture audit fields first.
      $_metadata = [PSCustomObject]@{
        Subject      = $_item.Subject
        ReceivedTime = $_item.ReceivedTime
      }
      # Use a stable file + store-relative path, including when previewing an
      # unattached archive whose Outlook display name is not yet available.
      $_destinationPath = $_archivePath + '::\' + $DestinationRelativePath
      $_resultArguments = @{
        Results               = $Results
        Action                = $ArchiveMode
        Folder                = $_folderPath
        Item                  = $_metadata
        DestinationFolderPath = $_destinationPath
        SourceEntryID         = $_id
        SourceStoreID         = [string]$SourceFolder.StoreID
      }

      # Record previews directly: ShouldProcess would print one WhatIf line
      # per message. Real transfers still require ShouldProcess approval.
      if ($WhatIfPreference) {
        Add-OutlookArchiveResult @_resultArguments -Status 'Skipped' -Detail 'DryRun'
      }
      elseif ($PSCmdlet.ShouldProcess("$_folderPath | $($_metadata.Subject)", "$ArchiveMode to $_destinationPath")) {
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
        Add-OutlookArchiveResult @_resultArguments -Status $_status -Detail "$_status to $_destinationPath"
      }
      else {
        Add-OutlookArchiveResult @_resultArguments -Status 'Skipped' -Detail 'Declined'
      }

      $_processed++
      Write-OutlookArchiveProgress -Phase $_phase -Folder $_folderPath -Current $_processed -Total $_ids.Count
    }
    catch {
      # A failed Copy().Move() can leave a duplicate in the source. Stop instead
      # of repeatedly filling a near-limit store. Do not delete uncertain items.
      throw "Archive stopped in '$($SourceFolder.FolderPath)' at EntryID '$_id': $($_.Exception.Message). A failed copy may remain in the source; inspect before retrying."
    }
    finally {
      Remove-ComObject $_moved $_copy $_item
    }
  }

  Write-OutlookArchiveProgress -Phase $_phase -Folder $_folderPath -Current $_processed -Total $_ids.Count -Force
}

$_context = $null
$_sourceRoot = $null
$_folderPlan = @()
$_destinationFolder = $null
$_archiveRoot = $null
$_archiveOwned = $false
$_archiveAttachedInitially = $false
$_initialStoreIDs = @()
$_destinationValidation = 'NotValidated'
$_archivePath = $ArchivePath
$_sourcePath = $null
$_sourceInfo = $null
$_sourceWarnings = @()
$_reportPath = $null
$_reportStream = $null
$_declined = $false

try {
  if ($AddDataFile -and $PSBoundParameters.ContainsKey('DetachWhenDone') -and $DetachWhenDone) {
    throw 'AddDataFile conflicts with DetachWhenDone set to true.'
  }

  if ($AddDataFile) {
    $DetachWhenDone = $false
  }

  # Reports are intentional local writes even during WhatIf. Reserve a new
  # file before opening Outlook so an unwritable destination stops real work.
  if ($PSBoundParameters.ContainsKey('ReportPath') -and $PSBoundParameters.ContainsKey('ReportDirectory')) {
    throw 'Use ReportPath or ReportDirectory, not both.'
  }

  $_archivePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArchivePath)
  if (-not $Append -and -not $PSBoundParameters.ContainsKey('DisplayName')) {
    $DisplayName = [IO.Path]::GetFileNameWithoutExtension($_archivePath)
  }
  if ((-not $Append -or $PSBoundParameters.ContainsKey('DisplayName')) -and [string]::IsNullOrWhiteSpace($DisplayName)) {
    throw 'DisplayName must not be blank.'
  }
  if ($PSBoundParameters.ContainsKey('ReportPath')) {
    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
      throw 'ReportPath must specify a report filename.'
    }

    $_reportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
    $_reportDirectory = Split-Path -Path $_reportPath -Parent
  }
  else {
    if ([string]::IsNullOrWhiteSpace($ReportDirectory)) {
      $_localAppData = [Environment]::GetFolderPath('LocalApplicationData')
      if ([string]::IsNullOrWhiteSpace($_localAppData)) {
        throw 'LocalApplicationData is unavailable. Supply ReportDirectory or ReportPath.'
      }

      $ReportDirectory = Join-Path $_localAppData 'winkit\reports\Outlook'
    }

    $_reportDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportDirectory)
    $_reportName = 'New-OutlookArchive-{0}-{1}.json' -f $_startedAt.ToString('yyyyMMdd-HHmmss-fff'), [guid]::NewGuid().ToString('N')
    $_reportPath = Join-Path $_reportDirectory $_reportName
  }

  if ($_reportPath -eq $_archivePath) {
    throw 'ReportPath and ArchivePath must be different files.'
  }

  if (Test-Path -LiteralPath $_reportPath) {
    throw "Report path already exists: $_reportPath. Choose a new report filename."
  }

  $null = [IO.Directory]::CreateDirectory($_reportDirectory)
  $_reportStream = [IO.File]::Open($_reportPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)

  if ([IO.Path]::GetExtension($_archivePath) -ne '.pst') {
    throw 'ArchivePath must end in .pst.'
  }

  if ($Append) {
    if (-not (Test-Path -LiteralPath $_archivePath -PathType Leaf)) {
      throw 'Append requires an existing PST file. Omit Append to create a new archive.'
    }
    $_archivePath = Resolve-LongPath -LiteralPath $_archivePath
    if ((Get-Item -LiteralPath $_archivePath -ErrorAction Stop).IsReadOnly) {
      throw 'ArchivePath is read-only.'
    }
  }
  elseif (Test-Path -LiteralPath $_archivePath) {
    throw 'ArchivePath already exists. Supply Append to add mail to this PST, or choose a new path.'
  }

  $_drive = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($_archivePath))
  if ($_archivePath.StartsWith('\\') -or $_drive.DriveType -eq 'Network') {
    throw 'ArchivePath must be on a local drive.'
  }

  if (-not (Test-Path -LiteralPath (Split-Path -Parent $_archivePath) -PathType Container)) {
    throw 'Archive parent directory must already exist.'
  }

  if ($EndDate -and $EndBefore) {
    throw 'Use EndDate or EndBefore, not both.'
  }

  if ($StartDate -and $EndDate -and $StartDate -gt $EndDate) {
    throw 'StartDate is after EndDate.'
  }
  if ($StartDate -and $EndBefore -and $StartDate -ge $EndBefore) {
    throw 'StartDate must be before EndBefore.'
  }

  Write-Log -Message 'Connecting to Outlook and locating the source store...' -Color Cyan
  Write-OutlookArchiveProgress -Phase 'Connecting to Outlook' -Force
  $_context = Connect-Outlook
  Write-OutlookArchiveProgress -Phase 'Reading Outlook stores' -Force

  $_outlookMajor = [int](($_context.App.Version -split '\.')[0])
  if ($_outlookMajor -lt 12) {
    throw "Outlook 2007 (version 12) or later is required. Detected Outlook version: $($_context.App.Version)"
  }

  # PSFoundation 1.3.0 checks Store.IsDefault, which Outlook does not expose.
  # Resolve the default via Namespace.DefaultStore and reject ambiguous names.
  if ([string]::IsNullOrWhiteSpace($StoreName)) {
    $_selectedStore = $_context.Namespace.DefaultStore
    try {
      $_sourceInfo = Get-OutlookArchiveSourceInfo -Store $_selectedStore
      $_sourceRoot = $_selectedStore.GetRootFolder()
    }
    finally {
      Remove-ComObject $_selectedStore
    }
  }
  else {
    $_stores = $_context.Namespace.Stores
    $_matches = 0
    try {
      for ($_storeIndex = 1; $_storeIndex -le $_stores.Count; $_storeIndex++) {
        $_store = $_stores.Item($_storeIndex)
        try {
          Write-Verbose "Store: $($_store.DisplayName) | $($_store.FilePath)"
          if ($_store.DisplayName -eq $StoreName) {
            $_matches++
            $_sourceInfo = Get-OutlookArchiveSourceInfo -Store $_store
          }
        }
        finally {
          Remove-ComObject $_store
        }
      }
    }
    finally {
      Remove-ComObject $_stores
    }

    if ($_matches -ne 1) {
      throw "StoreName '$StoreName' matches $_matches stores. Use a unique display name."
    }

    $_sourceRoot = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $StoreName
  }
  Write-Verbose "Source: $($_sourceRoot.FolderPath)"
  Write-Log -Message ("Archive source: {0} | Data file: {1} | ExchangeStoreType: {2}" -f $_sourceInfo.DisplayName, $_sourceInfo.DataFileFormat, $_sourceInfo.ExchangeStoreType) -Color Cyan
  if ($_sourceInfo.MaySynchronize) {
    $_sourceWarnings += 'Only selected mail exposed by Outlook is considered. Server-mailbox completeness is not verified; check cache history, folder subscriptions, and synchronization before archiving.'
    if ($Mode -eq 'Move') {
      $_sourceWarnings += 'Move removes archived messages from the source. Removal can synchronize to the server and other clients, including when Outlook reconnects later.'
    }
    else {
      $_sourceWarnings += 'Copy preserves originals but temporarily creates duplicates in the source before moving them to the PST. These writes can synchronize and require source write access and quota headroom.'
    }

    foreach ($_warning in $_sourceWarnings) {
      Write-Warning $_warning
    }
  }

  if ($Append) {
    # Reuse an existing attachment by file path, never by its display label.
    $_stores = $_context.Namespace.Stores
    $_archiveMatches = 0
    try {
      for ($_storeIndex = 1; $_storeIndex -le $_stores.Count; $_storeIndex++) {
        $_store = $_stores.Item($_storeIndex)
        try {
          $_initialStoreIDs += [string]$_store.StoreID
          $_storePath = [string]$_store.FilePath
          if ([string]::IsNullOrWhiteSpace($_storePath)) {
            continue
          }
          $_storePath = Resolve-LongPath -LiteralPath $_storePath
          if ($_storePath -eq $_archivePath) {
            $_archiveMatches++
            if ($_archiveMatches -gt 1) {
              throw 'ArchivePath matches more than one attached store.'
            }
            $_archiveRoot = $_store.GetRootFolder()
            $_archiveAttachedInitially = $true
            if ($_archiveRoot.StoreID -eq $_sourceRoot.StoreID) {
              throw 'Source and archive must be different stores.'
            }
          }
        }
        finally {
          Remove-ComObject $_store
        }
      }
    }
    finally {
      Remove-ComObject $_stores
    }

    if ($_archiveAttachedInitially -and $PSBoundParameters.ContainsKey('DetachWhenDone') -and $DetachWhenDone) {
      throw 'The archive was already attached to Outlook. Omit DetachWhenDone to preserve that attachment.'
    }
    $_destinationValidation = if ($_archiveAttachedInitially) { 'AttachedStore' } else { 'FileOnly' }
    if ($WhatIfPreference -and -not $_archiveAttachedInitially) {
      Write-Log -Message 'Append preview: the PST will remain detached; Outlook store and destination-folder validation is deferred until execution.' -Color Yellow
    }
    if ($Mode -eq 'Copy') {
      Write-Warning 'Append with Copy can duplicate previously archived messages. Use disjoint source selections; Append does not deduplicate or resume interrupted runs.'
    }
  }
  $_includedKinds = @('Inbox')
  foreach ($_option in $_inclusionGroups.Keys) {
    if ($PSBoundParameters.ContainsKey($_option) -and $PSBoundParameters[$_option]) {
      $_includedKinds += $_inclusionGroups[$_option]
    }
  }

  $_planArguments = @{
    Namespace  = $_context.Namespace
    StoreRoot  = $_sourceRoot
    Recurse    = [bool]$Recurse
    Include    = $_includedKinds
    Exclusions = $Exclusions
    ProgressId = 0
  }

  if ($PSBoundParameters.ContainsKey('FolderName')) {
    $_planArguments.FolderName = $FolderName
  }

  try {
    $_folderPlan = @(Get-OutlookFolderPlan @_planArguments)
  }
  catch {
    $_selectionError = $_.Exception.Message
    foreach ($_option in $_inclusionGroups.Keys) {
      foreach ($_kind in $_inclusionGroups[$_option]) {
        $_selectionError = $_selectionError.Replace("Include$_kind required", "$_option required")
      }
    }
    throw $_selectionError
  }

  foreach ($_entry in $_folderPlan) {
    foreach ($_option in $_inclusionGroups.Keys) {
      foreach ($_kind in $_inclusionGroups[$_option]) {
        $_entry.Reason = $_entry.Reason.Replace("Include$_kind required", "$_option required")
      }
    }
  }
  $_sourcePath = [string]$_folderPlan[0].FolderPath
  $script:OutlookArchiveFoldersSkipped = @($_folderPlan | Where-Object { -not $_.Process }).Count
  if (-not $WhatIfPreference) {
    $_destinationAction = if ($Append) { 'an EXISTING PST' } else { 'a NEW PST' }
    $_layout = if ($SkipPathPreservation) { 'flatten into archive root' } else { 'preserve source folder paths' }
    $_declined = -not $PSCmdlet.ShouldProcess("$_sourcePath -> $_archivePath", "Archive mail ($Mode), Recurse=$Recurse, to $_destinationAction; $_layout")
  }

  if (-not $WhatIfPreference -and -not $_declined) {
    if (-not $_archiveRoot) {
      # Recheck immediately before AddStoreEx, which creates missing files.
      if ($Append -and -not (Test-Path -LiteralPath $_archivePath -PathType Leaf)) {
        throw 'The archive disappeared before attachment. Append will not create a replacement.'
      }
      if (-not $Append -and (Test-Path -LiteralPath $_archivePath)) {
        throw 'The archive path appeared after validation. Refusing to reuse it without Append.'
      }
      $_phase = if ($Append) { 'Opening existing archive PST' } else { 'Creating archive PST' }
      Write-OutlookArchiveProgress -Phase $_phase -Folder $_archivePath -Force
      $_archiveRoot = Add-OutlookStoreRoot -Namespace $_context.Namespace -Path $_archivePath
      # Store identity also catches path aliases that long-name expansion
      # does not resolve, such as directory junctions.
      $_archiveAttachedInitially = $_archiveRoot.StoreID -in $_initialStoreIDs
      $_archiveOwned = -not $_archiveAttachedInitially -and $_archiveRoot.StoreID -ne $_sourceRoot.StoreID
    }
    if ($_archiveRoot.StoreID -eq $_sourceRoot.StoreID) {
      throw 'Source and archive must be different stores.'
    }
    if ($_archiveAttachedInitially -and $PSBoundParameters.ContainsKey('DetachWhenDone') -and $DetachWhenDone) {
      throw 'The archive was already attached to Outlook. Omit DetachWhenDone to preserve that attachment.'
    }

    $_destinationValidation = 'AttachedStore'
    if (-not $Append -or $PSBoundParameters.ContainsKey('DisplayName')) {
      try {
        $_archiveRoot.Name = $DisplayName
      }
      catch {
        Write-Warning "Could not rename archive PST store to '$DisplayName': $($_.Exception.Message)"
        Add-OperationResult -Results $_results -Target $_archivePath -Source 'Outlook' -Action 'RenameStore' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
    else {
      $DisplayName = [string]$_archiveRoot.Name
    }
  }

  if (-not $_declined) {
    Write-Verbose "Archive PST: $_archivePath"

    foreach ($_entry in $_folderPlan) {
      if (-not $_entry.Process) {
        Write-Verbose "Skipping $($_entry.FolderPath): $($_entry.Reason)"
        continue
      }

      $_folder = $null
      $_destinationFolder = $_archiveRoot
      try {
        $_folder = $_context.Namespace.GetFolderFromID($_entry.EntryID, $_entry.StoreID)
        if ($_folder.StoreID -ne $_entry.StoreID -or $_folder.EntryID -ne $_entry.EntryID) {
          throw 'Resolved source folder no longer matches the reviewed folder plan.'
        }
        $_destinationRelativePath = if ($SkipPathPreservation) { '' } else { [string]$_entry.RelativePath }
        if (-not $WhatIfPreference -and $_destinationRelativePath) {
          foreach ($_segment in $_destinationRelativePath.Split('\')) {
            $_children = $_destinationFolder.Folders
            $_matchingChildren = 0
            try {
              for ($_childIndex = 1; $_childIndex -le $_children.Count; $_childIndex++) {
                $_child = $_children.Item($_childIndex)
                try {
                  if ($_child.Name -eq $_segment) {
                    $_matchingChildren++
                  }
                }
                finally {
                  Remove-ComObject $_child
                }
              }
            }
            finally {
              Remove-ComObject $_children
            }
            if ($_matchingChildren -gt 1) {
              throw "Ambiguous destination folder '$($_destinationFolder.FolderPath)\$_segment'."
            }
            $_next = Get-OutlookSubFolder -ParentFolder $_destinationFolder -Name $_segment -Create
            if ($_destinationFolder -ne $_archiveRoot) {
              Remove-ComObject $_destinationFolder
            }
            $_destinationFolder = $_next
            $_accessor = $_destinationFolder.PropertyAccessor
            try {
              if ($_destinationFolder.DefaultItemType -ne 0 -or
                $_accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003') -eq 2) {
                throw "Destination folder '$($_destinationFolder.FolderPath)' is not a writable mail folder."
              }
            }
            finally {
              Remove-ComObject $_accessor
            }
          }
        }

        $_copyArguments = @{
          SourceFolder            = $_folder
          DestinationFolder       = $_destinationFolder
          DestinationRelativePath = $_destinationRelativePath
          ArchiveMode             = $Mode
          RequireFullDownload     = $_sourceInfo.DataFileFormat -eq 'OST'
          Results                 = $_results
          WhatIf                  = [bool]$WhatIfPreference
          Confirm                 = $false
        }
        Copy-OutlookFolderItem @_copyArguments
      }
      finally {
        if ($_destinationFolder -and $_destinationFolder -ne $_archiveRoot) {
          Remove-ComObject $_destinationFolder
        }
        $_destinationFolder = $null
        Remove-ComObject $_folder
      }
    }
  }
}
catch {
  Add-OperationResult -Results $_results -Target $_archivePath -Source 'Outlook' -Action 'Archive' -Status 'Failed' -Detail $_.Exception.Message
  Write-Warning $_.Exception.Message
}
finally {
  Write-Progress -Id 0 -Activity 'Outlook archive' -Completed
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
$_planned = @($_results | Where-Object { $_.Detail -eq 'DryRun' }).Count
$_status = if ($_failed -gt 0) {
  'Failed'
}
elseif ($_declined) {
  'Skipped'
}
elseif ($WhatIfPreference) {
  'Preview'
}
else {
  'Completed'
}

$_detail = if ($_failed -gt 0) {
  [string](@($_results | Where-Object { $_.Status -eq 'Failed' })[0].Detail)
}
else {
  'Per-message results are available in the JSON report.'
}

if ($_declined) {
  $_detail = 'Archive operation was declined.'
}

$_summaryProperty = @{
  Preview        = [bool]$WhatIfPreference
  FoldersRead    = $script:OutlookArchiveFoldersRead
  FoldersSkipped = $script:OutlookArchiveFoldersSkipped
  ItemsRead      = $script:OutlookArchiveItemsRead
  ItemsMatched   = $script:OutlookArchiveItemsMatched
  Planned        = $_planned
  Copied         = $script:OutlookArchiveCopied
  Moved          = $script:OutlookArchiveMoved
  Failed         = $_failed
  ReportPath     = $_reportPath
}

$_summary = New-OperationResult -Target $_archivePath -Source 'Outlook' -Action 'Archive' -Status $_status -Detail $_detail -Property $_summaryProperty

try {
  if (-not $_reportStream) {
    throw 'The report file could not be opened. No archive work was started.'
  }

  Write-OutlookArchiveProgress -Phase 'Writing JSON report' -Folder $_reportPath -Force
  # Sort a presentation copy only. Ordinal keeps ties stable on PowerShell 5.1.
  $_reportRows = for ($_ordinal = 0; $_ordinal -lt $_results.Count; $_ordinal++) {
    $_result = $_results[$_ordinal]
    $_receivedProperty = $_result.PSObject.Properties['Received']
    $_receivedDate = $null
    if ($_receivedProperty -and $null -ne $_receivedProperty.Value) {
      try {
        $_receivedDate = ([datetimeoffset]$_receivedProperty.Value).UtcDateTime
      }
      catch {
        # Unreadable dates remain visible with the other undated records.
        $_receivedDate = $null
      }
    }

    [PSCustomObject]@{
      Result   = $_result
      Undated  = $null -eq $_receivedDate
      Received = $_receivedDate
      Ordinal  = $_ordinal
    }
  }

  $_sortProperties = @(
    @{
      Expression = 'Undated'
      Descending = $false
    }
    @{
      Expression = 'Received'
      Descending = $Sort -eq 'NewToOld'
    }
    'Ordinal'
  )
  $_reportResults = @($_reportRows | Sort-Object -Property $_sortProperties | ForEach-Object { $_.Result })

  $_report = [ordered]@{
    SchemaVersion  = 1
    Script         = 'New-OutlookArchive'
    StartedAt      = $_startedAt.ToString('o')
    FinishedAt     = (Get-Date).ToString('o')
    Settings       = [ordered]@{
      ArchivePath              = $_archivePath
      Append                   = [bool]$Append
      SkipPathPreservation     = [bool]$SkipPathPreservation
      ArchiveAttachedInitially = $_archiveAttachedInitially
      AttachmentCreated        = $_archiveOwned
      DestinationValidation    = $_destinationValidation
      StoreName                = $StoreName
      SourceStore              = $_sourceInfo
      SourceFolder             = $_sourcePath
      FolderName               = if ($PSBoundParameters.ContainsKey('FolderName')) { $FolderName } else { $null }
      Sort                     = $Sort
      FolderSelection          = if ($PSBoundParameters.ContainsKey('FolderName')) { 'ExplicitPath' } else { 'DefaultInbox' }
      Recurse                  = [bool]$Recurse
      Include                  = @($_includedKinds)
      Exclusions               = @($Exclusions)
      Mode                     = $Mode
      StartDate                = if ($StartDate) { $StartDate.ToString('o') } else { $null }
      EndDate                  = if ($EndDate) { $EndDate.ToString('o') } else { $null }
      EndBefore                = if ($EndBefore) { $EndBefore.ToString('o') } else { $null }
      DetachWhenDone           = $DetachWhenDone
      DisplayName              = $DisplayName
      AddDataFile              = [bool]$AddDataFile
    }
    Summary        = $_summary
    SourceWarnings = @($_sourceWarnings)
    FolderPlan     = @($_folderPlan)
    Results        = $_reportResults
  }

  $_json = ConvertTo-Json -InputObject $_report -Depth 8 -ErrorAction Stop
  $_encoding = New-Object Text.UTF8Encoding($true)
  $_writer = New-Object IO.StreamWriter($_reportStream, $_encoding)
  try {
    $_writer.WriteLine($_json)
  }
  finally {
    $_writer.Dispose()
  }
}
catch {
  $_summary.Status = 'Failed'
  $_summary.Failed++
  $_summary.ReportPath = $null
  $_reportError = "JSON report could not be written: $($_.Exception.Message)"
  $_summary.Detail = if ($_failed -gt 0) { "$_detail $_reportError" } else { $_reportError }
  Write-Warning "$_reportError Archive counts below still describe work already performed; do not rerun blindly."
}
finally {
  if ($_reportStream) {
    $_reportStream.Dispose()
  }

  Write-Progress -Id 0 -Activity 'Outlook archive' -Completed
}

$_color = if ($_summary.Status -eq 'Failed' -or $WhatIfPreference -or $_declined) { 'Yellow' } else { 'Green' }
Write-Log -Message "Outlook archive: $($_summary.Status) | Read: $($_summary.ItemsRead) | Matched: $($_summary.ItemsMatched) | Planned: $_planned | Copied: $script:OutlookArchiveCopied | Moved: $script:OutlookArchiveMoved | Failed: $($_summary.Failed)" -Color $_color
Write-Log -Message "PST: $_archivePath" -Color Gray
if ($_summary.ReportPath) {
  Write-Log -Message "JSON report: $($_summary.ReportPath)" -Color Cyan
}

if ($PassThru -or $WhatIfPreference) {
  $_summary
}

if ($_summary.Status -eq 'Failed') {
  exit 1
}
