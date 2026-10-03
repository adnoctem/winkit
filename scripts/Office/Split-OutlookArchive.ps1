#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.7.1' }

<#
.SYNOPSIS
  Copies or moves one received-date range from an existing PST to another PST.
.DESCRIPTION
  Selects an existing source archive by ArchivePath and a destination by PSTPath.
  Requires an inclusive StartDate and either exclusive EndBefore or inclusive
  EndDate. One invocation transfers one range; there is no automatic yearly loop.
  Delegates filtering, folder preservation, transfer, progress, and JSON reporting
  to New-OutlookArchive. Its report identifies both source and destination files.
  The report's Script and operation names remain New-OutlookArchive and Archive.
  Copy is the default. Move explicitly removes transferred mail from the source.
  Inbox is the default; use FolderName '' with Recurse to select the whole store.
  Standard folder exclusions still apply. Non-mail and search folders are skipped.
  Requires classic Outlook, running under the same interactive Windows user.
  Source attachment can be temporary, including during previews; opening a PST
  can update file metadata. Pre-existing source attachments are preserved.
  Keep a closed-file backup before moving messages. Neither Copy nor Append is
  idempotent: overlapping Copy passes can produce duplicates. Source files are
  never deleted or compacted. Failures can leave a partially completed transfer;
  inspect the report before retrying.
.PARAMETER ArchivePath
  Existing local source PST. It must differ from PSTPath. Requires PSFoundation's
  Open-OutlookPstStore and Close-OutlookPstStore commands.
.PARAMETER PSTPath
  Local destination PST. Must be new unless Append is supplied. Its parent
  directory must already exist. This is New-OutlookArchive's ArchivePath.
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
.PARAMETER FolderName
  Exact folder path relative to the selected store, or empty for the store root.
  When omitted, selects Inbox by identity, including localized or renamed Inboxes.
  Explicit Inbox or Posteingang selects that literal path; no name fallback is used.
.PARAMETER Recurse
  Include descendants of the selected folder. Otherwise only its own mail is processed.
.PARAMETER StartDate
  Required inclusive received-date lower bound.
.PARAMETER EndDate
  Inclusive received-date upper bound, including the exact time.
  A date alone means midnight, not the end of that day. Prefer EndBefore.
.PARAMETER EndBefore
  Exclusive received-date upper bound, for example 2025-01-01 to
  include all of 2024. Cannot be combined with EndDate.
.PARAMETER Mode
  Copy leaves source mail intact. Move removes archived items from the source.
  Copy temporarily creates duplicates in the source before transferring them.
.PARAMETER DisplayName
  Display name for the destination PST. New files default to the filename without .pst.
  Append preserves the existing name unless this parameter is supplied.
  DataFileName is an alias.
.PARAMETER AddDataFile
  Keep the destination PST attached when this run added it to the current Outlook profile.
  Cannot be combined with DetachWhenDone set to true.
.PARAMETER DetachWhenDone
  Remove the destination attachment created by this run at the end. An already
  attached destination is left open; explicitly requesting its detachment is rejected.
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
  PS> .\Split-OutlookArchive.ps1 -ArchivePath D:\Archives\Combined.pst -PSTPath D:\Archives\2018.pst -StartDate '2018-01-01' -EndBefore '2019-01-01' -DryRun
.EXAMPLE
  PS> .\Split-OutlookArchive.ps1 -ArchivePath D:\Archives\Combined.pst -PSTPath D:\Archives\2018.pst -StartDate '2018-01-01' -EndBefore '2019-01-01' -FolderName '' -Recurse -IncludeSentItems -Mode Move -Append -Confirm
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not applicable - Outlook is a desktop client.
  SYSTEM-account execution: not applicable - requires an interactive Outlook profile.
  Outlook version: 2007 (version 12) or later for PST sources.
  Bitness: Outlook COM automation supports cross-architecture PowerShell clients.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The archive script owns ShouldProcess approval; this wrapper forwards WhatIf and Confirm without a second prompt.')]
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium', DefaultParameterSetName = 'EndBefore')]
param (
  [Parameter(Mandatory = $true)]
  [string]
  $ArchivePath,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $PSTPath,

  [switch]
  $Append,

  [switch]
  $SkipPathPreservation,

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

  [Parameter(Mandatory = $true)]
  [datetime]
  $StartDate,

  [Parameter(Mandatory = $true, ParameterSetName = 'EndDate')]
  [datetime]
  $EndDate,

  [Parameter(Mandatory = $true, ParameterSetName = 'EndBefore')]
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

if ($DryRun) {
  $WhatIfPreference = $true
}

# Preserve only explicitly bound options. In particular, an omitted FolderName
# must resolve the localized Inbox by identity in the archive script.
$_archiveArguments = @{}
foreach ($_name in $PSBoundParameters.Keys) {
  if ($_name -notin @('ArchivePath', 'PSTPath')) {
    $_archiveArguments[$_name] = $PSBoundParameters[$_name]
  }
}
$_archiveArguments.SourceArchivePath = $ArchivePath
$_archiveArguments.ArchivePath = $PSTPath

# One workflow owns the report, confirmation, and temporary source attachment.
$LASTEXITCODE = 0
& (Join-Path $PSScriptRoot 'New-OutlookArchive.ps1') @_archiveArguments
if ($LASTEXITCODE -ne 0) {
  exit $LASTEXITCODE
}
