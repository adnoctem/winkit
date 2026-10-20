#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.1.1' }

<#
.SYNOPSIS
  Surveys Outlook mail before archiving it.
.DESCRIPTION
  Reports what an archive run would find without transferring anything. Lists
  every data file in the current Outlook profile, then surveys the selected
  store: how many mail items each folder holds, how far back they reach, how
  they distribute by year, and which folders the current Include switches leave
  out. Nothing is created, moved, copied, renamed, or detached.

  Store and folder selection mirrors New-OutlookArchive.ps1 exactly. The same
  StoreName, SourceArchivePath, FolderName, Recurse, Exclusions, and Include
  switches select the same folders through the same folder planner, so the
  folder and skip counts match what an archive would act on.

  Received-date bounds behave differently here. StartDate, EndDate, and
  EndBefore do not restrict the survey; the year distribution is always
  reported in full. When a bound is supplied, the summary adds how much of the
  surveyed mail that window would select, which is the figure to confirm
  before archiving. Omit the bounds to discover the range that mail occupies.

  Only Class and ReceivedTime are read from each item, and no item is retained.
  Mail items with no readable received date are counted separately, because
  New-OutlookArchive.ps1 stops on them rather than skipping them.

  A JSON report is written only when ReportPath or ReportDirectory is supplied.
.PARAMETER StoreName
  Display name of the Outlook store to survey. If omitted, the default delivery
  store is used. Select an attached PST, IMAP, or Exchange store, not an OST path.
.PARAMETER SourceArchivePath
  Full path of an existing .pst file to survey directly. A temporary source
  attachment may be created and is removed again before the script returns.
.PARAMETER AllStores
  Survey every data file in the profile instead of one store. A store whose
  folder selection cannot be resolved is reported with its error and skipped
  rather than ending the survey.
.PARAMETER IncludeSentItems
  Permit the standard SentItems folder within the surveyed scope. Excluded unless explicitly included.
.PARAMETER IncludeDeletedItems
  Permit the standard DeletedItems folder within the surveyed scope. Excluded unless explicitly included.
.PARAMETER IncludeJunk
  Permit the standard Junk folder within the surveyed scope. Excluded unless explicitly included.
.PARAMETER IncludeOutbox
  Permit the standard Outbox folder within the surveyed scope. Excluded unless explicitly included.
.PARAMETER IncludeDrafts
  Permit the standard Drafts folder within the surveyed scope. Excluded unless explicitly included.
.PARAMETER IncludeMedia
  Permit Calendar, Contacts, Journal, Notes, Tasks, AllPublicFolders, RssFeeds,
  ToDo, ManagedEmail, and SuggestedContacts within the surveyed scope.
.PARAMETER IncludeFailures
  Permit SyncIssues, Conflicts, LocalFailures, and ServerFailures within the
  surveyed scope. Search folders remain excluded.
.PARAMETER Exclusions
  Exact store-relative paths excluded with their descendants. Exclusions win over Include switches.
.PARAMETER FolderName
  Store-relative folder path to survey. Omit for the identity-based default
  Inbox. An empty string selects the store root, which an archive PST needs
  because it has no Inbox.
.PARAMETER Recurse
  Survey descendants of the selected folder as well.
.PARAMETER StartDate
  Inclusive lower received-date bound of the reported window. Does not limit
  the survey itself.
.PARAMETER EndDate
  Inclusive upper received-date bound of the reported window. A date alone
  means midnight, not the end of that day. Prefer EndBefore.
.PARAMETER EndBefore
  Exclusive upper received-date bound of the reported window. Use the day after
  the last day to include, which covers that whole final day.
.PARAMETER Detailed
  Add a per-folder table of item counts and received-date ranges.
.PARAMETER ReportDirectory
  Directory to write a timestamped JSON report into. No report is written
  unless this or ReportPath is supplied.
.PARAMETER ReportPath
  Full path of a new JSON report file. A existing path is refused.
.PARAMETER IgnoreAdministrator
  Permit elevated execution. Outlook must run under the same Windows user and
  elevation. This override does not switch users or profiles.
.PARAMETER PassThru
  Emit the survey as an object in addition to printing it.
.EXAMPLE
  PS> .\Show-OutlookState.ps1 -AllStores
.EXAMPLE
  PS> .\Show-OutlookState.ps1 -StoreName 'user@example.com' -FolderName '' -Recurse
.EXAMPLE
  PS> .\Show-OutlookState.ps1 -StoreName 'user@example.com' -FolderName '' -Recurse -IncludeSentItems -StartDate '2025-01-01' -EndBefore '2026-01-01'
.EXAMPLE
  PS> $state = .\Show-OutlookState.ps1 -SourceArchivePath D:\Archive\Combined.pst -Detailed -PassThru
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

[CmdletBinding(DefaultParameterSetName = 'Store')]
[OutputType([PSCustomObject])]
param (
  [Parameter(ParameterSetName = 'Store')]
  [string]
  $StoreName,

  [Parameter(Mandatory = $true, ParameterSetName = 'Archive')]
  [ValidateNotNullOrEmpty()]
  [string]
  $SourceArchivePath,

  [Parameter(Mandatory = $true, ParameterSetName = 'All')]
  [switch]
  $AllStores,

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

  [switch]
  $Detailed,

  [string]
  $ReportDirectory,

  [ValidateNotNullOrEmpty()]
  [string]
  $ReportPath,

  [switch]
  $IgnoreAdministrator,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.1.1 -Force

# Outlook automation must use the mailbox user's interactive security context.
$_outlookUser = Get-UserInfo
if ($_outlookUser.IsAdministrator) {
  if (-not $IgnoreAdministrator) {
    throw "PowerShell is elevated as '$($_outlookUser.UserName)'. Run this script from a non-elevated PowerShell window as the Windows user who runs Outlook. Use -IgnoreAdministrator only when Outlook intentionally runs elevated under that same user."
  }

  Write-Warning "IgnoreAdministrator permits elevated execution as '$($_outlookUser.UserName)'. Outlook must run under the same Windows user and elevation. This override does not switch users or profiles."
}

# -----------------------------------------------------------------------------

# Script options expand to the identity-based kinds accepted by Substrate.
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

# Only the unambiguous OlObjectClass values are named. Meeting responses share
# one label because their numeric polarity is not worth asserting in a report.
$_itemClassNames = @{
  26 = 'Appointment'
  40 = 'Contact'
  41 = 'Document'
  42 = 'Journal'
  43 = 'Mail'
  44 = 'Note'
  45 = 'Post'
  46 = 'Report'
  48 = 'Task'
  49 = 'TaskRequest'
  50 = 'TaskRequest'
  51 = 'TaskRequest'
  52 = 'TaskRequest'
  53 = 'MeetingRequest'
  54 = 'MeetingCancellation'
  55 = 'MeetingResponse'
  56 = 'MeetingResponse'
  57 = 'MeetingResponse'
  69 = 'DistributionList'
}

$_mailItemClass = 43
$_startedAt = Get-Date
$script:OutlookStateFoldersRead = 0
$script:OutlookStateItemsRead = 0
$script:OutlookStateProgressTimer = [Diagnostics.Stopwatch]::StartNew()

function Write-OutlookStateProgress {
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

  if (-not $Force -and $script:OutlookStateProgressTimer.ElapsedMilliseconds -lt 250) {
    return
  }

  $script:OutlookStateProgressTimer.Restart()

  $_percent = if ($Total -gt 0) {
    [int]([math]::Min(100, 100 * $Current / $Total))
  }
  else {
    -1
  }

  $_progress = @{
    Id               = 0
    Activity         = 'Outlook state'
    Status           = $Phase
    CurrentOperation = "$Folder | $Current / $Total | Read: $script:OutlookStateItemsRead"
    PercentComplete  = $_percent
  }

  Write-Progress @_progress
}

function Get-OutlookStateSourceInfo {
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
    throw 'Cannot determine the source ExchangeStoreType. No survey was started.'
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

function Test-OutlookStateInWindow {
  param (
    [datetime]
    $Received
  )

  # Mirrors Test-OutlookItemInRange in New-OutlookArchive.ps1. The reported
  # window count is worthless if these comparisons drift apart.
  if ($StartDate -and $Received -lt $StartDate) {
    return $false
  }

  if ($EndDate -and $Received -gt $EndDate) {
    return $false
  }

  if ($EndBefore -and $Received -ge $EndBefore) {
    return $false
  }

  return $true
}

function Measure-OutlookStateFolder {
  param (
    [object]
    $Folder,

    [string]
    $RelativePath,

    [hashtable]
    $Totals
  )

  $_folderPath = [string]$Folder.FolderPath
  $script:OutlookStateFoldersRead++
  Write-OutlookStateProgress -Phase 'Reading folder contents' -Folder $_folderPath -Force

  $_folderMail = 0
  $_folderOther = 0
  $_folderUndated = 0
  $_folderOldest = $null
  $_folderNewest = $null

  $_items = $Folder.Items
  try {
    $_itemCount = $_items.Count
    for ($_index = 1; $_index -le $_itemCount; $_index++) {
      $_item = $null
      try {
        $_item = $_items.Item($_index)
        $script:OutlookStateItemsRead++

        $_class = $null
        try {
          $_class = $_item.Class
        }
        catch {
          $_class = $null
        }

        if ($null -eq $_class) {
          $Totals.Unreadable++
          continue
        }

        $_class = [int]$_class
        if (-not $Totals.Classes.ContainsKey($_class)) {
          $Totals.Classes[$_class] = 0
        }

        $Totals.Classes[$_class] = $Totals.Classes[$_class] + 1

        if ($_class -ne $_mailItemClass) {
          $_folderOther++
          continue
        }

        $_folderMail++

        $_received = $null
        try {
          $_received = $_item.ReceivedTime
        }
        catch {
          $_received = $null
        }

        if (-not $_received) {
          $_folderUndated++
          $Totals.UndatedMail++
          continue
        }

        $_received = [datetime]$_received
        $_year = $_received.Year
        if (-not $Totals.Years.ContainsKey($_year)) {
          $Totals.Years[$_year] = 0
        }

        $Totals.Years[$_year] = $Totals.Years[$_year] + 1

        if ($null -eq $Totals.Oldest -or $_received -lt $Totals.Oldest) {
          $Totals.Oldest = $_received
        }

        if ($null -eq $Totals.Newest -or $_received -gt $Totals.Newest) {
          $Totals.Newest = $_received
        }

        if ($null -eq $_folderOldest -or $_received -lt $_folderOldest) {
          $_folderOldest = $_received
        }

        if ($null -eq $_folderNewest -or $_received -gt $_folderNewest) {
          $_folderNewest = $_received
        }

        if ($Totals.HasWindow -and (Test-OutlookStateInWindow -Received $_received)) {
          $Totals.WindowMail++
        }

        Write-OutlookStateProgress -Phase 'Reading folder contents' -Folder $_folderPath -Current $_index -Total $_itemCount
      }
      finally {
        Remove-ComObject $_item
      }
    }
  }
  finally {
    Remove-ComObject $_items
  }

  $Totals.MailItems += $_folderMail
  $Totals.OtherItems += $_folderOther

  [PSCustomObject]@{
    FolderPath   = $_folderPath
    RelativePath = $RelativePath
    MailItems    = $_folderMail
    OtherItems   = $_folderOther
    UndatedMail  = $_folderUndated
    OldestMail   = $_folderOldest
    NewestMail   = $_folderNewest
  }
}

function Get-OutlookStateStoreSurvey {
  param (
    [object]
    $Context,

    [object]
    $StoreRoot,

    [object]
    $SourceInfo,

    [string[]]
    $IncludedKinds,

    [bool]
    $HasWindow
  )

  $_planArguments = @{
    Namespace  = $Context.Namespace
    StoreRoot  = $StoreRoot
    Recurse    = [bool]$Recurse
    Include    = $IncludedKinds
    Exclusions = $Exclusions
    ProgressId = 0
  }

  # $PSBoundParameters here belongs to this function, so the script-level flag
  # is what decides whether FolderName was supplied at all.
  if ($_folderNameBound) {
    $_planArguments.FolderName = $FolderName
  }

  Write-OutlookStateProgress -Phase 'Planning folders' -Folder ([string]$StoreRoot.FolderPath) -Force

  # Capture the planner's own diagnostics. Incomplete folder classification is
  # only visible here, and the console is not evidence.
  $_folderPlan = @(Get-OutlookFolderPlan @_planArguments -WarningVariable _planDiagnostics)
  $_planWarnings = @($_planDiagnostics | ForEach-Object { [string]$_ })

  foreach ($_entry in $_folderPlan) {
    foreach ($_option in $_inclusionGroups.Keys) {
      foreach ($_kind in $_inclusionGroups[$_option]) {
        $_entry.Reason = $_entry.Reason.Replace("Include$_kind required", "$_option required")
      }
    }
  }

  $_totals = @{
    Classes     = @{}
    Years       = @{}
    MailItems   = 0
    OtherItems  = 0
    UndatedMail = 0
    Unreadable  = 0
    WindowMail  = 0
    HasWindow   = $HasWindow
    Oldest      = $null
    Newest      = $null
  }

  $_folders = New-Object System.Collections.ArrayList
  $_skipped = New-Object System.Collections.ArrayList

  foreach ($_entry in $_folderPlan) {
    if (-not $_entry.Process) {
      $null = $_skipped.Add([PSCustomObject]@{
          FolderPath   = [string]$_entry.FolderPath
          RelativePath = [string]$_entry.RelativePath
          StandardKind = [string]$_entry.StandardKind
          Reason       = [string]$_entry.Reason
        })
      continue
    }

    $_folder = $null
    try {
      $_folder = $Context.Namespace.GetFolderFromID($_entry.EntryID, $_entry.StoreID)
      if ($_folder.StoreID -ne $_entry.StoreID -or $_folder.EntryID -ne $_entry.EntryID) {
        throw 'Resolved source folder no longer matches the reviewed folder plan.'
      }

      $null = $_folders.Add((Measure-OutlookStateFolder -Folder $_folder -RelativePath ([string]$_entry.RelativePath) -Totals $_totals))
    }
    finally {
      Remove-ComObject $_folder
    }
  }

  $_years = @($_totals.Years.Keys | Sort-Object | ForEach-Object {
      [PSCustomObject]@{
        Year = [int]$_
        Mail = [int]$_totals.Years[$_]
      }
    })

  $_cumulative = 0
  foreach ($_yearRow in $_years) {
    $_cumulative += $_yearRow.Mail
    $_yearRow | Add-Member -NotePropertyName Cumulative -NotePropertyValue $_cumulative
  }

  $_classes = @($_totals.Classes.Keys | Sort-Object -Property { $_totals.Classes[$_] } -Descending | ForEach-Object {
      $_name = $_itemClassNames[[int]$_]
      if (-not $_name) {
        $_name = "Class $_"
      }

      [PSCustomObject]@{
        Class = [int]$_
        Name  = [string]$_name
        Count = [int]$_totals.Classes[$_]
      }
    })

  $_window = $null
  if ($HasWindow) {
    $_window = [PSCustomObject]@{
      StartDate = if ($StartDate) { $StartDate.ToString('o') } else { $null }
      EndDate   = if ($EndDate) { $EndDate.ToString('o') } else { $null }
      EndBefore = if ($EndBefore) { $EndBefore.ToString('o') } else { $null }
      MailItems = [int]$_totals.WindowMail
    }
  }

  $_selection = if ($_folderPlan.Count -gt 0) { [string]$_folderPlan[0].FolderPath } else { [string]$StoreRoot.FolderPath }

  [PSCustomObject]@{
    DisplayName       = $SourceInfo.DisplayName
    FilePath          = $SourceInfo.FilePath
    DataFileFormat    = $SourceInfo.DataFileFormat
    ExchangeStoreType = $SourceInfo.ExchangeStoreType
    MaySynchronize    = $SourceInfo.MaySynchronize
    SourceFolder      = $_selection
    Include           = @($IncludedKinds)
    Exclusions        = @($Exclusions)
    Recurse           = [bool]$Recurse
    FoldersRead       = @($_folders).Count
    FoldersSkipped    = @($_skipped).Count
    ItemsRead         = [int]($_totals.MailItems + $_totals.OtherItems + $_totals.Unreadable)
    MailItems         = [int]$_totals.MailItems
    OtherItems        = [int]$_totals.OtherItems
    UndatedMail       = [int]$_totals.UndatedMail
    UnreadableItems   = [int]$_totals.Unreadable
    OldestMail        = $_totals.Oldest
    NewestMail        = $_totals.Newest
    Years             = $_years
    Classes           = $_classes
    Window            = $_window
    Folders           = @($_folders)
    SkippedFolders    = @($_skipped)
    PlanWarnings      = @($_planWarnings)
    Error             = $null
  }
}

# -----------------------------------------------------------------------------

if ($PSBoundParameters.ContainsKey('ReportPath') -and $PSBoundParameters.ContainsKey('ReportDirectory')) {
  throw 'Use ReportPath or ReportDirectory, not both.'
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

$_folderNameBound = $PSBoundParameters.ContainsKey('FolderName')
$_hasWindow = [bool]($StartDate -or $EndDate -or $EndBefore)
$_sourceArchivePath = $null
$_reportPath = $null

if ($PSBoundParameters.ContainsKey('SourceArchivePath')) {
  if (-not (Test-Path -LiteralPath $SourceArchivePath -PathType Leaf)) {
    throw 'SourceArchivePath must select an existing PST file.'
  }

  $_sourceArchivePath = Resolve-LongPath -LiteralPath $SourceArchivePath
  if ([IO.Path]::GetExtension($_sourceArchivePath) -ne '.pst') {
    throw 'SourceArchivePath must end in .pst.'
  }
}

# Reserve the report before opening Outlook so an unwritable destination stops
# the survey early. No report is written unless one was explicitly requested.
if ($PSBoundParameters.ContainsKey('ReportPath')) {
  if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    throw 'ReportPath must specify a report filename.'
  }

  $_reportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
  $_reportDirectory = Split-Path -Path $_reportPath -Parent
}
elseif ($PSBoundParameters.ContainsKey('ReportDirectory')) {
  if ([string]::IsNullOrWhiteSpace($ReportDirectory)) {
    throw 'ReportDirectory must specify a directory.'
  }

  $_reportDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportDirectory)
  $_reportName = 'Show-OutlookState-{0}-{1}.json' -f $_startedAt.ToString('yyyyMMdd-HHmmss-fff'), [guid]::NewGuid().ToString('N')
  $_reportPath = Join-Path $_reportDirectory $_reportName
}

if ($_reportPath) {
  if ($_sourceArchivePath -and $_reportPath -eq $_sourceArchivePath) {
    throw 'ReportPath and SourceArchivePath must be different files.'
  }

  if (Test-Path -LiteralPath $_reportPath) {
    throw "Report path already exists: $_reportPath. Choose a new report filename."
  }

  $null = [IO.Directory]::CreateDirectory($_reportDirectory)
}

$_includedKinds = @('Inbox')
foreach ($_option in $_inclusionGroups.Keys) {
  if ($PSBoundParameters.ContainsKey($_option) -and $PSBoundParameters[$_option]) {
    $_includedKinds += $_inclusionGroups[$_option]
  }
}

$_context = $null
$_sourceStoreContext = $null
$_profile = New-Object System.Collections.ArrayList
$_surveys = New-Object System.Collections.ArrayList

try {
  Write-Log -Message 'Connecting to Outlook and reading the profile...' -Color Cyan
  $_context = Connect-Outlook

  $_outlookMajor = [int](($_context.App.Version -split '\.')[0])
  if ($_outlookMajor -lt 12) {
    throw "Outlook 2007 (version 12) or later is required. Detected Outlook version: $($_context.App.Version)"
  }

  # Inventory every data file first. This alone answers which store holds the
  # mail, which is the question a wrong store selection hides.
  $_stores = $_context.Namespace.Stores
  try {
    for ($_storeIndex = 1; $_storeIndex -le $_stores.Count; $_storeIndex++) {
      $_store = $_stores.Item($_storeIndex)
      try {
        $null = $_profile.Add((Get-OutlookStateSourceInfo -Store $_store))
      }
      finally {
        Remove-ComObject $_store
      }
    }
  }
  finally {
    Remove-ComObject $_stores
  }

  # Resolve which stores to survey, then survey each with the shared selection.
  $_targets = New-Object System.Collections.ArrayList

  if ($_sourceArchivePath) {
    Write-Log -Message "Reading source PST: $_sourceArchivePath. A temporary source attachment may be needed." -Color Cyan
    $_sourceStoreContext = Open-OutlookPstStore -Namespace $_context.Namespace -LiteralPath $_sourceArchivePath -WhatIf:$false -Confirm:$false -ErrorAction Stop
    if ($null -eq $_sourceStoreContext) {
      throw 'The source PST could not be opened.'
    }

    $_sourceRoot = $_sourceStoreContext.Root
    $_selectedStore = $_sourceRoot.Store
    try {
      $null = $_targets.Add([PSCustomObject]@{
          Root  = $_sourceRoot
          Info  = Get-OutlookStateSourceInfo -Store $_selectedStore
          Owned = $false
        })
    }
    finally {
      Remove-ComObject $_selectedStore
    }
  }
  elseif ($AllStores) {
    foreach ($_info in $_profile) {
      $null = $_targets.Add([PSCustomObject]@{
          Root  = $null
          Info  = $_info
          Owned = $true
        })
    }
  }
  elseif ([string]::IsNullOrWhiteSpace($StoreName)) {
    $_selectedStore = $_context.Namespace.DefaultStore
    try {
      $null = $_targets.Add([PSCustomObject]@{
          Root  = $_selectedStore.GetRootFolder()
          Info  = Get-OutlookStateSourceInfo -Store $_selectedStore
          Owned = $true
        })
    }
    finally {
      Remove-ComObject $_selectedStore
    }
  }
  else {
    $_matches = @($_profile | Where-Object { $_.DisplayName -eq $StoreName })
    if ($_matches.Count -ne 1) {
      throw "StoreName '$StoreName' matches $($_matches.Count) stores. Use a unique display name."
    }

    $null = $_targets.Add([PSCustomObject]@{
        Root  = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $StoreName
        Info  = $_matches[0]
        Owned = $true
      })
  }

  foreach ($_target in $_targets) {
    $_root = $_target.Root
    try {
      if ($null -eq $_root) {
        $_root = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $_target.Info.DisplayName
      }

      $_surveyArguments = @{
        Context       = $_context
        StoreRoot     = $_root
        SourceInfo    = $_target.Info
        IncludedKinds = $_includedKinds
        HasWindow     = $_hasWindow
      }

      $null = $_surveys.Add((Get-OutlookStateStoreSurvey @_surveyArguments))
    }
    catch {
      # A store without a resolvable selection must not end a profile survey.
      Write-Warning "Could not survey '$($_target.Info.DisplayName)': $($_.Exception.Message)"
      $null = $_surveys.Add([PSCustomObject]@{
          DisplayName       = $_target.Info.DisplayName
          FilePath          = $_target.Info.FilePath
          DataFileFormat    = $_target.Info.DataFileFormat
          ExchangeStoreType = $_target.Info.ExchangeStoreType
          MaySynchronize    = $_target.Info.MaySynchronize
          SourceFolder      = $null
          Include           = @($_includedKinds)
          Exclusions        = @($Exclusions)
          Recurse           = [bool]$Recurse
          FoldersRead       = 0
          FoldersSkipped    = 0
          ItemsRead         = 0
          MailItems         = 0
          OtherItems        = 0
          UndatedMail       = 0
          UnreadableItems   = 0
          OldestMail        = $null
          NewestMail        = $null
          Years             = @()
          Classes           = @()
          Window            = $null
          Folders           = @()
          SkippedFolders    = @()
          PlanWarnings      = @()
          Error             = $_.Exception.Message
        })
    }
    finally {
      if ($_target.Owned) {
        Remove-ComObject $_root
      }
    }
  }
}
finally {
  Write-Progress -Id 0 -Activity 'Outlook state' -Completed

  if ($_sourceStoreContext) {
    try {
      Close-OutlookPstStore -Context $_sourceStoreContext -ErrorAction Stop
    }
    catch {
      Write-Warning "Could not clean up the source PST attachment: $($_.Exception.Message)"
    }
  }

  if ($_context) {
    Remove-ComObject $_context.Namespace $_context.App
  }

  Invoke-ComGarbageCollection
}

# -----------------------------------------------------------------------------

Write-Log -Message "`nOutlook profile: $(@($_profile).Count) data file(s)" -Color Cyan
foreach ($_info in $_profile) {
  Write-Log -Message ("  {0,-34} {1,-5} {2}" -f $_info.DisplayName, $_info.DataFileFormat, $_info.FilePath) -Color Gray
}

foreach ($_survey in $_surveys) {
  Write-Log -Message "`nStore: $($_survey.DisplayName)" -Color Cyan

  if ($_survey.Error) {
    Write-Log -Message "  Not surveyed: $($_survey.Error)" -Color Yellow
    continue
  }

  $_synchronizes = if ($_survey.MaySynchronize) { 'may synchronize' } else { 'local only' }
  Write-Log -Message ("  Data file: {0} | ExchangeStoreType: {1} ({2})" -f $_survey.DataFileFormat, $_survey.ExchangeStoreType, $_synchronizes) -Color Gray
  Write-Log -Message ("  Selection: {0} | Recurse: {1} | Include: {2}" -f $_survey.SourceFolder, $_survey.Recurse, ($_survey.Include -join ', ')) -Color Gray
  Write-Log -Message ("  Folders:   {0} surveyed, {1} skipped" -f $_survey.FoldersRead, $_survey.FoldersSkipped) -Color Gray

  $_range = if ($_survey.OldestMail -and $_survey.NewestMail) {
    '{0:yyyy-MM-dd} to {1:yyyy-MM-dd}' -f $_survey.OldestMail, $_survey.NewestMail
  }
  else {
    'no dated mail'
  }

  Write-Log -Message ("  Mail:      {0} item(s) | {1}" -f $_survey.MailItems, $_range) -Color Gray

  if ($_survey.OtherItems -gt 0) {
    $_otherClasses = @($_survey.Classes | Where-Object { $_.Class -ne $_mailItemClass } | ForEach-Object { "$($_.Name) $($_.Count)" })
    Write-Log -Message ("  Other:     {0} item(s) | {1}" -f $_survey.OtherItems, ($_otherClasses -join ', ')) -Color Gray
  }

  if ($_survey.UndatedMail -gt 0) {
    Write-Warning "$($_survey.UndatedMail) mail item(s) in '$($_survey.DisplayName)' have no readable received date. New-OutlookArchive.ps1 stops on these rather than skipping them."
  }

  if ($_survey.UnreadableItems -gt 0) {
    Write-Warning "$($_survey.UnreadableItems) item(s) in '$($_survey.DisplayName)' did not report a class and were not counted. Verify that Outlook is licensed and the store is healthy."
  }

  if (@($_survey.Years).Count -gt 0) {
    Write-Log -Message "`n  Year      Mail   Cumulative" -Color Gray
    foreach ($_yearRow in $_survey.Years) {
      Write-Log -Message ("  {0,-6} {1,7} {2,12}" -f $_yearRow.Year, $_yearRow.Mail, $_yearRow.Cumulative) -Color Gray
    }
  }

  if ($_survey.Window) {
    $_bounds = @()
    if ($_survey.Window.StartDate) {
      $_bounds += 'from {0:yyyy-MM-dd}' -f [datetime]$_survey.Window.StartDate
    }

    if ($_survey.Window.EndDate) {
      $_bounds += 'through {0:yyyy-MM-dd}' -f [datetime]$_survey.Window.EndDate
    }

    if ($_survey.Window.EndBefore) {
      $_bounds += 'before {0:yyyy-MM-dd}' -f [datetime]$_survey.Window.EndBefore
    }

    $_windowColor = if ($_survey.Window.MailItems -gt 0) { 'Green' } else { 'Yellow' }
    Write-Log -Message ("`n  Window {0}: {1} of {2} mail item(s) would be archived" -f ($_bounds -join ' '), $_survey.Window.MailItems, $_survey.MailItems) -Color $_windowColor

    if ($_survey.Window.MailItems -eq 0 -and $_survey.MailItems -gt 0) {
      Write-Warning "The requested window selects none of the $($_survey.MailItems) mail item(s) in '$($_survey.DisplayName)'. Compare it against the year distribution above before archiving."
    }
  }

  if ($Detailed -and @($_survey.Folders).Count -gt 0) {
    Write-Log -Message "`n  Folder                                 Mail   Other  Oldest       Newest" -Color Gray
    foreach ($_folderRow in $_survey.Folders) {
      $_label = if ($_folderRow.RelativePath) { $_folderRow.RelativePath } else { '<root>' }
      $_oldest = if ($_folderRow.OldestMail) { '{0:yyyy-MM-dd}' -f $_folderRow.OldestMail } else { '-' }
      $_newest = if ($_folderRow.NewestMail) { '{0:yyyy-MM-dd}' -f $_folderRow.NewestMail } else { '-' }
      Write-Log -Message ("  {0,-38} {1,5} {2,7}  {3,-12} {4}" -f $_label, $_folderRow.MailItems, $_folderRow.OtherItems, $_oldest, $_newest) -Color Gray
    }
  }

  if ($Detailed -and @($_survey.SkippedFolders).Count -gt 0) {
    Write-Log -Message "`n  Skipped folders" -Color Gray
    foreach ($_skippedRow in $_survey.SkippedFolders) {
      Write-Log -Message ("  {0,-38} {1}" -f $_skippedRow.RelativePath, $_skippedRow.Reason) -Color Gray
    }
  }
}

$_state = [PSCustomObject]@{
  SchemaVersion = 1
  Script        = 'Show-OutlookState'
  StartedAt     = $_startedAt.ToString('o')
  FinishedAt    = (Get-Date).ToString('o')
  ItemsRead     = [int]$script:OutlookStateItemsRead
  FoldersRead   = [int]$script:OutlookStateFoldersRead
  Profile       = @($_profile)
  Stores        = @($_surveys)
  ReportPath    = $_reportPath
}

if ($_reportPath) {
  try {
    $_json = $_state | ConvertTo-Json -Depth 8
    $_encoding = New-Object Text.UTF8Encoding($true)
    [IO.File]::WriteAllText($_reportPath, $_json, $_encoding)
    Write-Log -Message "`nJSON report: $_reportPath" -Color Cyan
  }
  catch {
    $_state.ReportPath = $null
    Write-Warning "JSON report could not be written: $($_.Exception.Message)"
  }
}

$_surveyResult = New-OperationResult -Target 'Outlook' -Source 'Outlook' -Action 'Survey' -Status 'Completed' -Detail ("{0} store(s) surveyed, {1} item(s) read" -f @($_surveys).Count, $script:OutlookStateItemsRead)
$_operationLog = Write-OperationResultLog -Results @($_surveyResult) -ScriptName 'Show-OutlookState' -Name 'winkit'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_state
}
