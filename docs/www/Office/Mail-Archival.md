# Mail archival

[Office overview](README.md) | [Requirements](Requirements.md)

Run from your **installed winkit directory**, as the Outlook user in a normal, non-elevated PowerShell session. Use classic Outlook 2007+
for PST sources or 2010+ for IMAP/Exchange sources. Take a [backup or checkpoint](Outlook-Backup-and-Repair.md) before changing real mail.

## Preview and archive the Inbox

By default, `New-OutlookArchive` copies mail from the selected store's Inbox into a new Unicode PST. Inbox selection uses Outlook's
identity, so it also finds `Posteingang` or a renamed Inbox. Subfolders require `-Recurse`; non-mail items and search folders are skipped.

```powershell
$archiveDirectory = Join-Path $env:LOCALAPPDATA 'winkit\archives'
New-Item -ItemType Directory -Path $archiveDirectory -Force | Out-Null

$archive = @{
  StoreName   = 'user@example.com'
  ArchivePath = Join-Path $archiveDirectory 'mail-2024.pst'
  StartDate   = [datetime]'2024-01-01'
  EndBefore   = [datetime]'2025-01-01'
  Mode        = 'Copy'
}

$preview = .\scripts\Office\New-OutlookArchive.ps1 @archive -DryRun -PassThru
notepad.exe $preview.ReportPath

# After reviewing the preview:
$result = .\scripts\Office\New-OutlookArchive.ps1 @archive -Confirm -PassThru
```

Omit `StoreName` to use the default delivery store. Otherwise select a unique display name; duplicate names are rejected. The destination's
parent must exist and the source/destination must be distinct stores.

| Option                | Effect                                                                                                                                              |
| --------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Mode Copy` (default) | Leaves originals intact, but temporarily duplicates each message in the source before moving the copy. Allow write access and space in both stores. |
| `Mode Move`           | Removes successfully archived messages from the source. Server-backed stores can synchronize that removal.                                          |
| `StartDate`           | Inclusive lower bound.                                                                                                                              |
| `EndBefore`           | Exclusive upper bound; convenient for whole days/years.                                                                                             |
| `EndDate`             | Inclusive exact time; a date alone means midnight. Cannot be combined with `EndBefore`.                                                             |

## Split an existing archive

`Split-OutlookArchive` transfers one explicit date range from an existing source PST (`ArchivePath`) to a destination PST (`PSTPath`). Run
it separately for each range. It uses the same archival workflow, folder rules and JSON report as `New-OutlookArchive`.

```powershell
$archiveDirectory = Join-Path $env:LOCALAPPDATA 'winkit\archives'
New-Item -ItemType Directory -Path $archiveDirectory -Force | Out-Null

$split = @{
  ArchivePath      = 'D:\Archives\Combined archive.pst'
  PSTPath          = Join-Path $archiveDirectory 'Archive - 2018.pst'
  StartDate        = [datetime]'2018-01-01'
  EndBefore        = [datetime]'2019-01-01'
  FolderName       = ''
  Recurse          = $true
  IncludeSentItems = $true
  Mode             = 'Move'
  AddDataFile      = $true
}

$preview = .\scripts\Office\Split-OutlookArchive.ps1 @split -DryRun -PassThru
notepad.exe $preview.ReportPath

# After reviewing the preview:
$result = .\scripts\Office\Split-OutlookArchive.ps1 @split -Confirm -PassThru
```

Both paths must differ. `StartDate` and one upper bound (`EndBefore` or `EndDate`) are required. Copy is the default; the example explicitly
chooses Move. Without `FolderName ''` and `Recurse`, only Inbox is selected. Other standard folders still require their Include switch.

To add another selection to an existing destination, supply `Append`. It requires an existing file and does not deduplicate overlapping Copy
passes. Folder paths are preserved unless `SkipPathPreservation` is supplied. The source PST is not deleted or compacted.

```powershell
# After creating the 2018 destination, preview a separately approved folder.
$additional = $split.Clone()
$additional.FolderName = 'Offene Themen'
.\scripts\Office\Split-OutlookArchive.ps1 @additional -Append -DryRun

# Configure the next range separately, preview it, and review its report.
$split.PSTPath = Join-Path $archiveDirectory 'Archive - 2019.pst'
$split.StartDate = [datetime]'2019-01-01'
$split.EndBefore = [datetime]'2020-01-01'
.\scripts\Office\Split-OutlookArchive.ps1 @split -DryRun
```

New can also select an existing source PST directly. Here, `ArchivePath` retains its usual meaning as the **destination**:

```powershell
.\scripts\Office\New-OutlookArchive.ps1 `
  -SourceArchivePath 'D:\Archives\Combined archive.pst' `
  -ArchivePath (Join-Path $archiveDirectory 'Archive - 2018.pst') `
  -StartDate '2018-01-01' -EndBefore '2019-01-01' `
  -FolderName '' -Recurse -IncludeSentItems -Mode Copy -DryRun
```

`SourceArchivePath` replaces `StoreName`; it cannot be combined with it. These operations require classic Outlook and PSFoundation's
`Open-OutlookPstStore`/`Close-OutlookPstStore` commands. A source PST already attached to Outlook stays attached. Otherwise the script
temporarily attaches it and cleans up afterward, including during previews. Opening a PST can update its metadata. Source files must be
writable and local; network paths and paths traversing reparse points are refused. Do not replace the source or change its profile
attachment during a run. Other attached data-file paths must also be inspectable and local; an inaccessible, network or reparse-point path
elsewhere in the profile can block source inspection. `AddDataFile` and `DetachWhenDone` control the destination, not the temporary source.

Split returns the same single summary and report as New, with `SourceFilePath` and `DestinationFilePath`. Its report retains
`Script: New-OutlookArchive` and the `Archive` operation name. The settings also record `SourceArchivePath` and whether the source
attachment was created by this run. A failure can leave completed transfers in place; review the report before retrying.

## Select folders and exclusions

The same rules apply to archiving and duplicate review. Paths are exact, case-insensitive, and relative to the store; they are not wildcard
patterns or searches across every folder. A missing folder or unresolved Inbox identity fails without widening scope.

```powershell
# Use a separate destination for these selection previews; DryRun does not create it.
$selection = $archive.Clone()
$selection.ArchivePath = Join-Path $archiveDirectory ('selection-' + [guid]::NewGuid().ToString('N') + '.pst')

# Preview just a top-level sibling of Inbox.
.\scripts\Office\New-OutlookArchive.ps1 @selection -FolderName 'Offene Themen' -DryRun

# Preview a selected subtree, protecting a nested reference folder.
.\scripts\Office\New-OutlookArchive.ps1 @selection `
  -FolderName 'Posteingang\Hub' -Recurse -Exclusions 'Posteingang\Hub\Referenz' -DryRun

# Preview eligible folders throughout the store, including sent mail.
.\scripts\Office\New-OutlookArchive.ps1 @selection `
  -FolderName '' -Recurse -IncludeSentItems -Exclusions 'Offene Themen' -DryRun
```

| Selection                                            | Scope                                                                                                             |
| ---------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| Omit `FolderName`                                    | Direct Inbox mail, regardless of display language/name.                                                           |
| `-FolderName 'Inbox'` or `'Posteingang'`             | The folder literally named by that path.                                                                          |
| `-FolderName 'Posteingang\Hub\Filters'`              | Direct mail in that nested folder; add `-Recurse` for eligible descendants.                                       |
| `-FolderName ''`                                     | Direct mail at the store root; add `-Recurse` for eligible folders throughout the store.                          |
| `-Exclusions 'Offene Themen','Posteingang\Referenz'` | Excludes whole subtrees, even if a descendant is selected directly. `ExcludeFolders` is an alias for these paths. |

Inbox is permitted by default; other identified standard folders require inclusion. Inclusion only permits folders **within the selected
scope**; it does not enable recursion or add unrelated folders. Exclusions win, and an excluded parent excludes its children.

| Standard folders                                                                                             | Inclusion switch                                                           |
| ------------------------------------------------------------------------------------------------------------ | -------------------------------------------------------------------------- |
| Inbox                                                                                                        | Already permitted; exclude its displayed path with `Exclusions` if needed. |
| Sent mail                                                                                                    | `IncludeSentItems` (alias `IncludeSentMail`)                               |
| Deleted mail                                                                                                 | `IncludeDeletedItems`                                                      |
| Junk mail                                                                                                    | `IncludeJunk`                                                              |
| Drafts                                                                                                       | `IncludeDrafts`                                                            |
| Outbox                                                                                                       | `IncludeOutbox`                                                            |
| Calendar, Contacts, Journal, Notes, Tasks, AllPublicFolders, RssFeeds, ToDo, ManagedEmail, SuggestedContacts | `IncludeMedia`                                                             |
| SyncIssues, Conflicts, LocalFailures, ServerFailures                                                         | `IncludeFailures`                                                          |

Identities follow [Outlook's standard folder enumeration](https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders), not
translated-name guesses. `IncludeMedia` enables traversal to mail subfolders; it never archives contacts, appointments, or other non-mail
items and does not filter by attachment type. Media/Failures groups are off by default. Search folders, including virtual To-Do views, are
always skipped to avoid duplicate processing.

For non-default Outlook 2007 PSTs, unresolved Inbox identity may require an explicit `FolderName`. Outlook 2007 uses PST MAPI properties;
Exchange/OST identity discovery requires a newer classic client. Provider-specific Spam/Trash folders without a standard identity need
explicit exclusions. Unexpected provider errors stop planning before mail transfer or destination-folder creation.

## Append and preserve folder paths

Source paths are preserved from the store root: `Posteingang\Hub\Filters\Amazon` becomes that same path inside the archive. Existing
destination folders are reused; missing folders are created. Ambiguous names or incompatible folders stop processing.

After the first archive run, append a separately approved folder:

```powershell
.\scripts\Office\New-OutlookArchive.ps1 @archive `
  -Append -FolderName 'Posteingang\Hub\Filters\Amazon' -DryRun

.\scripts\Office\New-OutlookArchive.ps1 @archive `
  -Append -FolderName 'Posteingang\Hub\Filters\Amazon' -Confirm
```

To flatten selected mail into the destination root instead:

```powershell
$flat = $archive.Clone()
$flat.ArchivePath = Join-Path $archiveDirectory 'mail-2024-flat.pst'
.\scripts\Office\New-OutlookArchive.ps1 @flat -FolderName '' -Recurse -SkipPathPreservation -DryRun
```

| Destination rule | Behavior                                                                                                                                                     |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| New archive      | Existing files are refused unless `Append` is supplied.                                                                                                      |
| Append           | Requires an existing local PST. It is not deduplication, overwrite, or a resume ledger; overlapping Copy passes create duplicates.                           |
| Failure          | Successful transfers and the archive remain. A failed Copy transfer can leave a duplicate in the source; inspect both stores and the report before retrying. |
| Flattening       | `SkipPathPreservation` combines messages at the archive root; source paths remain in reports.                                                                |
| Store comparison | Long filesystem names are resolved before comparing source/destination stores.                                                                               |

## Keep the archive in Outlook

```powershell
$attached = $archive.Clone()
$attached.ArchivePath = Join-Path $archiveDirectory 'Archive - 2024.pst'
.\scripts\Office\New-OutlookArchive.ps1 @attached -AddDataFile -DisplayName 'Archive - 2024' -DryRun
.\scripts\Office\New-OutlookArchive.ps1 @attached -AddDataFile -DisplayName 'Archive - 2024' -Confirm
```

Use this as an alternative destination, not another overlapping Copy pass without a reason. `StoreName` always identifies the source;
`ArchivePath` identifies the destination; `DisplayName` (alias `DataFileName`) only changes the destination label.

| Attachment option           | Behavior                                                                                                                                  |
| --------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| Default                     | A store attached by the run is detached afterward. New PST display names use the filename without `.pst`.                                 |
| `AddDataFile`               | Keep the archive attached in the current profile/data-file list. `DetachWhenDone:$false` is also supported.                               |
| Conflicting options         | `AddDataFile` with explicit `DetachWhenDone:$true` is rejected.                                                                           |
| Previously attached archive | Append leaves it open and retains its label unless `DisplayName` is explicit. Explicit detachment is refused.                             |
| Preview                     | Never attaches a destination PST. A detached archive has `DestinationValidation = FileOnly`; store/folder validation waits for execution. |
| Rename failure              | Warning and failed report entry.                                                                                                          |

Close Outlook before copying the archive elsewhere; detaching a store does not prove all processes released its file.

## IMAP and Exchange sources

Attached IMAP/Exchange stores use the same selections and destination PST options. There is no standalone OST conversion or OST-path input.
Synchronize complete messages first; verify Exchange cache history and IMAP folder subscriptions cover the requested date range.

```powershell
$serverArchive = @{
  StoreName   = 'user@example.com'
  ArchivePath = Join-Path $archiveDirectory 'server-mail-before-2025.pst'
  EndBefore   = [datetime]'2025-01-01'
  Mode        = 'Copy'
}
.\scripts\Office\New-OutlookArchive.ps1 @serverArchive -DryRun -PassThru
.\scripts\Office\New-OutlookArchive.ps1 @serverArchive -Confirm -PassThru
```

| Server-store condition            | Consequence                                                                                                                                                                                                                                                                                          |
| --------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Copy                              | Temporary source duplicates can synchronize; source write access and quota headroom are needed.                                                                                                                                                                                                      |
| Move                              | Removal can synchronize to the server and other clients, including after reconnecting. Both preview and execution warn; no extra opt-in flag is required.                                                                                                                                            |
| Incomplete/unknown download state | OST matching messages must report full body/attachment download during enumeration and before transfer. Missing, header-only, unreadable, or unknown state stops the run, including previews; failure reports identify folder/EntryID.                                                               |
| Cache coverage                    | The script does not synchronize, change cache settings, or detect older messages absent from the cache. Success covers only selected mail Outlook exposes, not the complete server mailbox.                                                                                                          |
| Result evidence                   | `Settings.SourceStore` records display name, StoreID, path, DataFileFormat, numeric ExchangeStoreType, and MaySynchronize. Extension alone does not identify IMAP; ExchangeStoreType 3 means non-Exchange. SourceWarnings retains notices. Transfer success does not confirm server synchronization. |

Test disposable mail in a dedicated server folder first. Verify PST bodies, attachments, dates, and layout; confirm originals remain for
Copy and the server view matches intended removals for Move. A raw OST checkpoint is supplementary, not a portable PST backup. See
[Microsoft's PST archiving guidance](https://support.microsoft.com/en-us/outlook/mail/archive-in-outlook-for-windows),
[cached-export limitations](https://support.microsoft.com/en-us/outlook/export-emails-contacts-and-calendar-items-to-outlook-using-a-pst-file),
and [MailItem.DownloadState](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.downloadstate).

## Read and sort reports

Archive runs print counts and the JSON path rather than every message. The default is a unique UTF-8 report under
`%LOCALAPPDATA%\winkit\reports\Outlook`. Reports contain metadata, not message bodies or attachments.

```powershell
$reportDirectory = Join-Path $env:LOCALAPPDATA 'winkit\reports\Outlook'
$reportPreview = $archive.Clone()
$reportPreview.ArchivePath = Join-Path $archiveDirectory ('report-preview-' + [guid]::NewGuid().ToString('N') + '.pst')
$preview = .\scripts\Office\New-OutlookArchive.ps1 @reportPreview -DryRun -PassThru `
  -ReportDirectory $reportDirectory -Sort NewToOld
notepad.exe $preview.ReportPath

$report = Get-Content -LiteralPath $preview.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
$messages = $report.Results
@($messages).Count
```

For an exact filename, use `ReportPath` instead of `ReportDirectory`:

```powershell
$reportPath = Join-Path $reportDirectory ('archive-' + [guid]::NewGuid().ToString('N') + '.json')
.\scripts\Office\New-OutlookArchive.ps1 @reportPreview -DryRun -ReportPath $reportPath -Sort OldToNew
```

Absolute and current-directory-relative report paths are accepted, but use a writable directory outside winkit. Existing reports are never
overwritten. The parent and reserved report file are created before Outlook opens; an unwritable destination stops the run. Reports are
finalized on completion, preview, or caught failure. Interruptions can leave incomplete files; report-write failures warn and exit 1 without
undoing transferred mail.

| Report detail            | Meaning                                                                                                                                                                                                       |
| ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Returned summary         | `ReportPath`, `Status`, `Preview`, `FoldersRead`, `FoldersSkipped`, `ItemsRead`, `ItemsMatched`, `Planned`, `Copied`, `Moved`, `Failed`. Returned by PassThru/DryRun/WhatIf.                                  |
| Counts                   | Skipped folders count encountered exclusions/non-mail containers, not all descendants. ItemsRead includes inspected non-mail items; ItemsMatched counts date-filtered mail. Failed-run counts may be partial. |
| `FolderPlan`, `Settings` | Selection reasons, filters, inclusion/exclusion flags, Sort, Append, SkipPathPreservation, attachment ownership, and destination validation. Run timestamps and summary are retained.                         |
| Folder selection         | Settings.SourceFolder is effective path; FolderSelection is DefaultInbox or ExplicitPath. Requested FolderName is null for implicit selection, empty for explicit store root.                                 |
| Message results          | Subject in Target (`<No Subject>` when blank), folder Scope, ISO 8601 Received, action/status/detail, SourceFolderPath, DestinationFolderPath, SourceEntryID, SourceStoreID. Subjects are not modified.       |
| Destination path         | `archive.pst::\store-relative\folder`; `archive.pst::\` is the archive root. Identifiers aid auditing, not crash-safe resumption.                                                                             |
| Preview                  | Status Skipped, Detail DryRun; no per-message console result/WhatIf lines.                                                                                                                                    |
| Sort                     | NewToOld (default) or OldToNew, across all selected folders. Equal dates keep processing order; undated records follow dated mail in original order. Scope identifies the folder.                             |

Sorting affects archive JSON, duplicate-review CSV/logs and per-message output, not processing order or the duplicate kept. The archive
returns one summary. Progress shows connection, store/folder planning, reading/filtering, preview recording, and transfer; item percentages
describe the current folder/phase, not the entire job.

## Review duplicates

`Optimize-Outlook` compares ordinal, case-sensitive transport Message-IDs **within each folder**. The first encountered item stays;
subsequent matches move to `_Duplicates_Review`. It does not compare bodies/attachments, deduplicate across folders, or hard-delete mail.
Messages without usable IDs are skipped; the retained item is not necessarily oldest/newest.

Use `PSTPath` instead of `StoreName` to review an existing local archive. It uses classic Outlook and the same temporary-source attachment
rules as splitting; this is not offline file processing. The review folder is inside the selected PST.

```powershell
$pstReview = @{
  PSTPath          = 'D:\Archives\Archive - 2018.pst'
  FolderName       = ''
  Recurse          = $true
  IncludeSentItems = $true
  Sort             = 'NewToOld'
}

.\scripts\Office\Optimize-Outlook.ps1 @pstReview -DryRun `
  -ReportPath (Join-Path $env:LOCALAPPDATA 'winkit\reports\duplicates-preview.csv')

# After reviewing the CSV, choose a fresh report filename for execution.
.\scripts\Office\Optimize-Outlook.ps1 @pstReview -Confirm `
  -ReportPath (Join-Path $env:LOCALAPPDATA 'winkit\reports\duplicates-run.csv')
```

```powershell
$reportDirectory = Join-Path $env:LOCALAPPDATA 'winkit\reports\Outlook'
New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
$review = @{
  StoreName  = 'user@example.com'
  ReportPath = Join-Path $reportDirectory ('duplicates-' + [guid]::NewGuid().ToString('N') + '.csv')
  Sort       = 'OldToNew'
}
.\scripts\Office\Optimize-Outlook.ps1 @review -DryRun

# Review the preview; execution needs a new CSV filename.
$review.ReportPath = Join-Path $reportDirectory ('duplicates-' + [guid]::NewGuid().ToString('N') + '.csv')
.\scripts\Office\Optimize-Outlook.ps1 @review -Confirm -PassThru
```

Inbox-only is the default; use the [same folder controls](#select-folders-and-exclusions) for other scopes. `ReviewFolderName` changes the
top-level review folder, which is always excluded from scanning. Review candidates before manual deletion. CSV is written only when
requested, including in previews; parents are created and existing files refused. Folder records retain exclusion reasons; message records
retain received dates and Message-IDs even when the first row is a folder record. Blank subjects use `<No Subject>`.

Duplicate review does not compact a PST and is optional, not a migration prerequisite. Folder controls do not apply to whole-file
backup/repair or Office deployment.

## Rehearse and reduce a production PST

1. Record PST paths/names, Outlook version/architecture, account delivery locations, folder counts, and local-only contacts, calendars,
   tasks, rules, signatures, and settings. Mail archiving does not migrate non-mail data or settings.
2. Close Outlook, verify `OUTLOOK.EXE` has exited, and make a hash-verified [backup](Outlook-Backup-and-Repair.md#back-up-pst-files). Keep
   one backup unopened and unchanged. Make a separate working copy with a unique display name.
3. Keep active PSTs on local disk outside synchronization folders. Allow space for backup, working copy, and archives. There is no automatic
   splitting, limit detection, or reliable destination-growth estimate; use small date batches and check capacity between them.

For a local rehearsal, select **Work Offline** in Outlook and dismiss credential dialogs manually. In Outlook 2007 this is on the File menu;
see the
[offline-status guidance](https://support.microsoft.com/en-au/topic/outlook-2007-2010-status-is-always-offline-and-can-t-receive-or-send-mail-normally-easy-fix-articles-3e977242-7f66-2f0a-5956-51007986b38f).
Disconnecting networking after dependencies are installed can avoid obsolete account connections; it does not repair account settings or
close an existing dialog. Local PST archiving does not require authenticating an old POP account. Preserve paths/backups before removing a
profile; configure replacement IMAP separately and verify local POP messages exist on the server before discarding their PST.

```powershell
$rehearsal = @{
  StoreName        = 'PST - WORKING COPY'
  FolderName       = ''
  Recurse          = $true
  IncludeSentItems = $true
  ArchivePath      = Join-Path $archiveDirectory 'before-2024-rehearsal.pst'
  EndBefore        = [datetime]'2024-01-01'
  Mode             = 'Copy'
}
.\scripts\Office\New-OutlookArchive.ps1 @rehearsal -DryRun -PassThru -Verbose
.\scripts\Office\New-OutlookArchive.ps1 @rehearsal -Confirm -PassThru
```

Reopen the archive and compare eligible per-folder counts, first/last dates, selected nested folders, included Sent Items, bodies, and
representative attachments. Inspect the source too: counters alone do not establish integrity, and declined transfers are not completed
transfers. Calendar/contact/task data need separate migration. Rehearse on disposable data before a real-PST working copy.

Outlook 2007's Unicode defaults are a 20 GB maximum and 19 GB data threshold; configured limits may differ. Copy needs source write headroom
and can leave a temporary duplicate on transfer failure. Near a limit, rehearse on a working copy under the target classic Outlook with
sufficient capacity; do not raise limits or choose destructive Move just to bypass a failed Copy rehearsal. See
[PST/OST limits](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/data-files/configure-size-limit-outlook-data-files) and
[network PST limitations](https://learn.microsoft.com/en-us/troubleshoot/outlook/data-files/limits-using-pst-files-over-lan-wan).

After validation, choose the real source and a **different** production destination, then preview before Move:

```powershell
$production = $rehearsal.Clone()
$production.StoreName = 'user@example.com'
$production.ArchivePath = Join-Path $archiveDirectory 'before-2024-production.pst'
$production.Mode = 'Move'
.\scripts\Office\New-OutlookArchive.ps1 @production -DryRun -PassThru
.\scripts\Office\New-OutlookArchive.ps1 @production -Confirm -PassThru
```

Append only later disjoint batches; do not append the same mail to a Copy rehearsal archive. Reconcile partial results before retrying:
there is no resume ledger, transactional rollback, or deduplication of previous archives. Moving mail may not shrink the file immediately.
After validation and another recoverable backup, compact using Outlook's data-file settings; see
[Microsoft's compaction guidance](https://support.microsoft.com/en-us/outlook/reduce-the-size-of-your-mailbox-and-outlook-data-files-pst-and-ost).
Retain the original backup until mail, contacts, calendar, attachments, and send/receive work on the destination installation. Repair only
when needed, with Outlook closed and a preserved backup.
