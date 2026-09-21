# Outlook PST migration and archive rehearsal

Review date: 2026-09-21. These scripts use the classic Outlook COM object model. Use classic desktop Outlook for this workflow. Microsoft's
[automation guidance](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/vba-alternatives) describes the automation
differences in new Outlook; opening a PST in a client does not imply that these scripts can automate it.

## Decision before production use

The changes in this review address concrete script defects, but simulated tests do not certify Outlook 2007 COM behavior or the health of an
existing PST. Rehearse on disposable data and then a working copy of the real PST. Do not start a production Move run unless the scratch
suite and your manual archive checks pass on the actual Outlook installation.

Outlook 2007 defaults to a 20 GB Unicode PST maximum, with a 19 GB data threshold. A source approaching that size needs attention before
attempting more writes. Registry settings can change these limits; the scripts do not inspect or change them. See
[Microsoft's size-limit documentation](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/data-files/configure-size-limit-outlook-data-files).

## Review findings and changes

| Area                | Original problem                                                                                                           | Change / remaining boundary                                                                                           |
| ------------------- | -------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| Archive destination | Existing/source PST could be reused, renamed, detached or recursively archived into itself                                 | Require a new local `.pst` path; check distinct store identity; reject existing files                                 |
| Copy iteration      | `Copy()` changes the source Items collection; a failed transfer could leave copies that the loop subsequently copied again | Snapshot EntryIDs before mutation; stop on first transfer failure                                                     |
| Result accuracy     | Success was logged before Move, including failed or declined operations                                                    | Log success after transfer; capture metadata before Move invalidates the original COM object                          |
| Dates               | `-EndDate '2024-12-31'` meant midnight, excluding most of December 31                                                      | Document exact-time semantics; add exclusive `-EndBefore '2025-01-01'`                                                |
| Archive scope       | Unfiltered runs could move contacts/calendar/tasks into generic mail folders; root mail was omitted                        | Archive mail only, include root mail, skip non-mail folder subtrees and virtual search folders                        |
| Store selection     | Installed PSFoundation 1.3.0 checks nonexistent `Store.IsDefault`; names could be ambiguous                                | Resolve the default via `Namespace.DefaultStore`; reject duplicate display names                                      |
| Deduplication       | Excluded folders' children were still processed; Message-IDs compared case-insensitively                                   | Exclude complete subtrees; compare IDs ordinally; skip search and non-mail folders                                    |
| Confirmation        | Creating/mounting folders bypassed the operation's confirmation                                                            | Confirm the operation before destination creation; previews do not quit Outlook                                       |
| COM lifetime        | Returned Move objects and some child COM objects were discarded without release                                            | Retain and release those references                                                                                   |
| Repair              | Native path arguments with spaces were not quoted; nonzero exits were not failures                                         | Quote the PST path; refuse running Outlook/locked files; report nonzero exit as failure                               |
| Test generator      | ANSI/Unicode tags were reversed; partial header/date injection could report success                                        | Correct tags; require both writes to report successful fixture injection; share Outlook's MAPI session for Redemption |
| Integration suite   | Fixed temporary paths were reused/deleted; archive checks largely trusted script reports                                   | Unique retained scratch directory/store name; reopen archives and compare actual item counts                          |

The default dedup exclusions are English display names. Supply the corresponding names on a localized installation. Deduplication uses
Message-ID alone, not body/attachment equality, and consolidates candidates into one review folder in the same PST. It neither compacts the
file nor frees mail storage by itself. Keep this as a separate, reviewed task; it is not a prerequisite for migration.

## Prepare a recoverable working copy

1. Record all current PST paths and display names, Outlook version/bitness, account delivery locations, and folder counts. Note local-only
   contacts, calendars, tasks, rules, signatures and other settings. This mail archiver does not migrate those settings or non-mail data.
2. Close Outlook and confirm `OUTLOOK.EXE` has exited. Copy every relevant PST to a separate backup location. Verify the copies, for example
   with `Get-FileHash -Algorithm SHA256` on source and backup while Outlook remains closed. Preserve one backup unopened and unchanged.
3. Make a separate working copy for rehearsal. Open that copy in classic Outlook under a distinct, easily recognized store name. Avoid
   simultaneously attaching multiple copies with indistinguishable display names. Keep active PSTs on local disk, outside synchronization
   folders; do not run these scripts against a network PST. Microsoft documents
   [network PST limitations](https://learn.microsoft.com/en-US/outlook/troubleshoot/data-files/limits-using-pst-files-over-lan-wan).
4. Budget disk space for the untouched backup, working copy, and all archive PSTs. The script has no automatic archive splitting, source
   size-limit detection, or reliable estimate of destination growth. Use small date batches and check size/headroom between them.

## Outlook 2007 workstation and credential prompts

Run as the logged-in Outlook user at the same elevation as Outlook, not SYSTEM. Use Windows PowerShell 5.1 x86 for the Outlook 2007
rehearsal, especially if using 32-bit Redemption. On 64-bit Windows its path is:

```text
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
```

The PSFoundation module manifest itself requires PowerShell 5.1 even though the script syntax follows the repository's 5.0 convention.
Verify the pinned modules are visible from this exact PowerShell executable. Repository initialization installs the versions in
`requirements.psd1`; modules installed only for a different PowerShell edition/user may not be visible here.

Open Outlook manually and put it in **Work Offline**, then cancel outstanding credential dialogs. Outlook 2007 exposes Work Offline in the
File menu; confirm its status before running the test. Microsoft's
[Outlook 2007/2010 offline article](https://support.microsoft.com/en-au/topic/outlook-2007-2010-status-is-always-offline-and-can-t-receive-or-send-mail-normally-easy-fix-articles-3e977242-7f66-2f0a-5956-51007986b38f)
also shows the status-bar control. For a local rehearsal, disconnecting networking after dependencies are available can keep obsolete
account connections out of the test. This does not fix the old account configuration, and an already-open dialog may still need dismissal.

Local PST archiving does not require authenticating the old POP mailbox. Do not delete the old account/profile merely to stop prompts before
you have captured its data paths and backups. The scripts reuse the running/default Outlook session; they do not select a profile or switch
it offline. `-DryRun` still connects to Outlook and can encounter profile or security dialogs.

Create the replacement IMAP configuration separately during migration and retain the original POP PST as a local data source. A server-side
mailbox migration is not evidence that every locally downloaded POP message is on the IMAP server. Check that explicitly before disposal.

## Rehearsal and production sequence

From the repository root, in the intended PowerShell host:

```powershell
.\winkit.ps1 init
.\winkit.ps1 test -Outlook
```

The Outlook integration suite attaches a uniquely named scratch PST to the active profile. It retains its temporary directory for
inspection. Treat failed archive/count checks as a stop condition. Skipped header/date-dependent tests do not verify those features; use a
disposable PST containing real dated mail or a properly licensed Redemption installation to complete the missing checks. Never run
`New-TestOutlookMessage.ps1` against the boss's production store.

Choose a unique source display name and a NEW destination path whose parent already exists. Substitute your actual paths, name and date:

```powershell
$archive = @{
  StoreName   = 'Boss PST - WORKING COPY'
  ArchivePath = 'D:\MailArchive\before-2024-rehearsal.pst'
  EndBefore   = [datetime]'2024-01-01'
  Mode        = 'Copy'
}
.\scripts\Office\New-OutlookArchive.ps1 @archive -DryRun -PassThru -Verbose
.\scripts\Office\New-OutlookArchive.ps1 @archive -PassThru -Confirm
```

Copy leaves the original messages in place, but Outlook's `Copy()` temporarily creates another item in the SOURCE PST before moving that
copy to the archive. **It is not a read-only operation and is unsuitable when the source has no write headroom.** A failed copy transfer may
leave the temporary duplicate in the source; the hardened script stops and reports this rather than deleting uncertain data. For a PST
already near its Outlook 2007 limit, prefer rehearsing on a working copy under the target classic Outlook installation with sufficient
configured headroom. Do not raise limits or launch a destructive move merely to get around a failing copy test.

After copying, reopen the archive in Outlook and compare eligible per-folder counts, first/last dates, Sent Items, nested folders, message
bodies and representative attachments. Check both source and archive; no failed results and a success counter alone are insufficient. The
archive is mail-only: calendar/contact/task data remain in the source PST and still need migration. Close Outlook before copying the archive
file elsewhere. A successful RemoveStore call is not proof that every process has released the file.

For production reduction after the rehearsal and backup verification, use a **different new PST path** and `-Mode Move`. Do not copy into an
archive and then rerun Move into the same file: existing destinations are intentionally refused. Existing Copy archives and newly created
Move archives will contain overlapping mail; label rehearsal artifacts accordingly. Preview each disjoint date batch first. Stop on any
failure, inspect the operation log and both PSTs, and reconcile the partially completed batch before retrying; there is no resume ledger,
transactional rollback or deduplication of previous archives.

Moving mail out does not necessarily shrink the PST's physical file immediately. Compact only after archive validation and another
recoverable backup, using Outlook's data-file settings. See
[Microsoft's compaction guidance](https://support.microsoft.com/en-us/outlook/reduce-the-size-of-your-mailbox-and-outlook-data-files-pst-and-ost).
Keep the original backup until the new installation's mail, contacts, calendar, attachments and send/receive behavior are verified.

Use `Repair-OutlookDataFile.ps1` only when a scan/repair is actually needed, with Outlook closed and a backup preserved. The tool is
interactive; its exit code does not prove that repair was performed or that the PST is healthy. Inspect ScanPST's own findings/log.

## Validation boundaries

The automated regression suite uses simulated Outlook objects to exercise collection mutation, transfer failures, date boundaries, store
selection, preview behavior, excluded subtrees, repair quoting and locks. It cannot establish Outlook 2007 compatibility, available PST
headroom, fixture property persistence, or real archive integrity. The gated integration suite must be run on the intended workstation. The
review machine has PSFoundation 1.3.0, Pester 6.0.1 and PSScriptAnalyzer 1.25.0; the repository pins 1.4.0, 5.5.0 and 1.22.0 respectively.

The 22 Office convention/regression checks passed under PowerShell 7.6 and Windows PowerShell 5.1 (with the existing module directory added
to that test process's module search path). The full non-Outlook suite could not pass because the Maintenance and Policy tests require the
missing PSFoundation 1.4.0. No Outlook COM integration run or real PST modification was performed during this review.
