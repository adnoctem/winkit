#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.8.2' }

<#
.SYNOPSIS
  Deduplicates Outlook mail items per folder using the transport Message-ID.
.DESCRIPTION
  Processes Inbox by default, with optional FolderName and Recurse selection.
  Standard folders other than Inbox require their Include switch; custom
  Exclusions always take precedence. Keys each received mail
  item by its RFC Message-ID, and moves every occurrence after the first into a
  review folder. It never hard-deletes messages.

  Deduplication scope is per folder. The same message legitimately living in
  two different folders is preserved.
  Shows folder traversal and throttled item progress while inspecting headers
  and processing duplicates. ProgressPreference controls the progress display.
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
  Display name of the Outlook store to process. If omitted, the default
  delivery store is used. Run with -Verbose to list detected stores.
.PARAMETER PSTPath
  Existing local PST to process instead of StoreName or the default mailbox.
  Requires classic Outlook and PSFoundation's Open-OutlookPstStore and
  Close-OutlookPstStore commands. A detached source is temporarily attached,
  even for previews; an existing attachment is preserved. Opening the file
  can update its metadata. Duplicate review remains inside this same PST.
.PARAMETER ReviewFolderName
  Top-level folder created under the store root for duplicate review.
.PARAMETER FolderName
  Exact store-relative folder path, or empty for the store root.
  When omitted, selects Inbox by identity, including localized or renamed Inboxes.
  Explicit Inbox or Posteingang selects that literal path; no name fallback is used.
.PARAMETER Recurse
  Visit descendants of the selected folder. Otherwise process its direct mail only.
.PARAMETER ReportPath
  Optional CSV path containing Keep, MoveDuplicate, and SkipNoMessageId results.
.PARAMETER DryRun
  Preview changes without moving duplicate messages.
.PARAMETER Sort
  Order report results by received date: NewToOld (default) is newest first;
  OldToNew is oldest first. Undated records follow dated messages. Equal dates
  retain processing order. This affects reports only, not message processing.
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
.EXAMPLE
  PS> .\Optimize-Outlook.ps1 -PSTPath D:\Archives\2018.pst -FolderName '' -Recurse -DryRun -ReportPath .\archive-duplicates.csv
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not applicable - Outlook is a desktop client.
  SYSTEM-account execution: not applicable - requires an interactive Outlook profile.
  Outlook version: 2007 (version 12) or later - PropertyAccessor is required.
  Bitness: Outlook COM automation supports cross-architecture PowerShell clients.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Store')]
param (
  [Parameter(ParameterSetName = 'Store')]
  [string]
  $StoreName,

  [Parameter(Mandatory = $true, ParameterSetName = 'Archive')]
  [ValidateNotNullOrEmpty()]
  [string]
  $PSTPath,

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

  [string]
  $ReviewFolderName = '_Duplicates_Review',

  [string]
  $ReportPath,

  [switch]
  $DryRun,

  [ValidateSet('OldToNew', 'NewToOld')]
  [string]
  $Sort = 'NewToOld',

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
    $_received = $Item.ReceivedTime
  }
  catch {
    $_received = $null
  }

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

$_context = $null
$_storeRoot = $null
$_reviewFolder = $null
$_sourceStoreContext = $null
$_target = $StoreName
$_declined = $false

try {
  if ($PSBoundParameters.ContainsKey('PSTPath')) {
    $_target = $PSTPath
    foreach ($_command in @('Open-OutlookPstStore', 'Close-OutlookPstStore')) {
      if (-not (Get-Command -Name $_command -ErrorAction SilentlyContinue)) {
        throw 'PST source selection requires a PSFoundation version providing Open-OutlookPstStore and Close-OutlookPstStore. Update PSFoundation before retrying.'
      }
    }
  }

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
  if ($PSBoundParameters.ContainsKey('PSTPath')) {
    Write-Log -Message "Reading source PST: $PSTPath. A temporary source attachment may be needed, including during preview." -Color Cyan
    $_sourceStoreContext = Open-OutlookPstStore -Namespace $_context.Namespace -LiteralPath $PSTPath -WhatIf:$false -Confirm:$false -ErrorAction Stop
    if ($null -eq $_sourceStoreContext) {
      throw 'The source PST could not be opened.'
    }

    $_storeRoot = $_sourceStoreContext.Root
    $_target = $_sourceStoreContext.Path
  }
  elseif ([string]::IsNullOrWhiteSpace($StoreName)) {
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

  if ([string]::IsNullOrWhiteSpace($ReviewFolderName) -or $ReviewFolderName.Contains('\')) {
    throw 'ReviewFolderName must name one top-level folder.'
  }
  $Exclusions = @($Exclusions) + @($ReviewFolderName)
  $_includedKinds = @('Inbox')
  foreach ($_option in $_inclusionGroups.Keys) {
    if ($PSBoundParameters.ContainsKey($_option) -and $PSBoundParameters[$_option]) {
      $_includedKinds += $_inclusionGroups[$_option]
    }
  }

  $_planArguments = @{
    Namespace  = $_context.Namespace
    StoreRoot  = $_storeRoot
    Recurse    = [bool]$Recurse
    Include    = $_includedKinds
    Exclusions = $Exclusions
    ProgressId = 20
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

  if (-not $WhatIfPreference) {
    $_declined = -not $PSCmdlet.ShouldProcess($_sourcePath, "Move suspected duplicates to '$ReviewFolderName' for review")

    if ($_declined) {
      # Finish cleanup and reporting even after declined approval. A temporary
      # source attachment may still need cleanup, and its failure must surface.
      Add-OperationResult -Results $_results -Target $_target -Source 'Outlook' -Action 'Deduplicate' -Status 'Skipped' -Detail 'Duplicate review was declined.'
    }
    else {
      Write-Progress -Id 20 -Activity 'Outlook deduplication' -Status 'Preparing duplicate review folder' -CurrentOperation $ReviewFolderName -PercentComplete -1
      $_reviewFolder = Get-OutlookSubFolder -ParentFolder $_storeRoot -Name $ReviewFolderName -Create
    }
  }

  foreach ($_entry in $_folderPlan) {
    if ($_declined) {
      break
    }

    if (-not $_entry.Process) {
      Add-OperationResult -Results $_results -Target $_entry.FolderPath -Source 'Outlook' -Action 'SelectFolder' -Status 'Skipped' -Detail $_entry.Reason
      continue
    }

    $_folder = $_context.Namespace.GetFolderFromID($_entry.EntryID, $_entry.StoreID)
    try {
      if ($_folder.StoreID -ne $_entry.StoreID -or $_folder.EntryID -ne $_entry.EntryID) {
        throw 'Resolved source folder no longer matches the reviewed folder plan.'
      }
      Optimize-OutlookFolder -Folder $_folder -ReviewFolder $_reviewFolder -Results $_results -WhatIf:$WhatIfPreference -Confirm:$false
    }
    finally {
      Remove-ComObject $_folder
    }
  }
}
catch {
  $_failureDetail = $_.Exception.Message
  $_sourceCleanupError = $_.Exception.Data['OutlookPstCleanupError']
  if ($_sourceCleanupError) {
    $_failureDetail = "$_failureDetail $_sourceCleanupError"
  }
  Add-OperationResult -Results $_results -Target $_target -Source 'Outlook' -Action 'Deduplicate' -Status 'Failed' -Detail $_failureDetail
  Write-Warning $_failureDetail
}
finally {
  Write-Progress -Id 21 -Activity 'Inspecting messages and processing duplicates' -Completed
  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Completed
  Remove-ComObject $_reviewFolder
  if ($_sourceStoreContext) {
    try {
      Close-OutlookPstStore -Context $_sourceStoreContext -ErrorAction Stop
    }
    catch {
      Add-OperationResult -Results $_results -Target $_target -Source 'Outlook' -Action 'DetachSourceStore' -Status 'Failed' -Detail $_.Exception.Message
      Write-Warning "Could not clean up the source PST attachment: $($_.Exception.Message)"
    }
    $_storeRoot = $null
  }
  else {
    Remove-ComObject $_storeRoot
  }

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

  if ($ReportPath) {
    $_reportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
    $_reportRoot = Split-Path -Path $_reportPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($_reportRoot) -and -not (Test-Path -LiteralPath $_reportRoot)) {
      $null = [IO.Directory]::CreateDirectory($_reportRoot)
    }

    # Folder-selection records have fewer properties than message records.
    # Preserve every column even when an excluded folder is the first result.
    $_columns = @($_results | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
    $_reportResults | Select-Object -Property $_columns |
      Export-Csv -LiteralPath $_reportPath -NoTypeInformation -Encoding UTF8 -NoClobber -WhatIf:$false -Confirm:$false -ErrorAction Stop
    Write-Log -Message "Report: $_reportPath" -Color Gray
  }

  $_operationLog = Write-OperationResultLog -Results $_reportResults -ScriptName 'Optimize-Outlook' -Name 'winkit'
  if ($_operationLog) {
    Write-Log -Message "Operation log: $_operationLog" -Color Gray
  }
}
finally {
  Write-Progress -Id 20 -Activity 'Outlook deduplication' -Completed
}

if ($PassThru -or $DryRun) {
  $_reportResults
}

if ($_failed -gt 0) {
  exit 1
}
