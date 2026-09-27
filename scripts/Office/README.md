# Office

Scripts for deploying Microsoft Office and maintaining Outlook mail stores and data files on Windows.

| Script                                                   | Purpose                                                                                 |
| -------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| [Switch-OfficeVersion.ps1](Switch-OfficeVersion.ps1)     | Inventory Office, prepare installation media, and migrate to a selected Office product. |
| [New-OutlookArchive.ps1](New-OutlookArchive.ps1)         | Copy or move mail into a new Unicode PST archive.                                       |
| [Optimize-Outlook.ps1](Optimize-Outlook.ps1)             | Move duplicate messages into a review folder.                                           |
| [New-TestOutlookMessage.ps1](New-TestOutlookMessage.ps1) | Create synthetic messages for testing in an Outlook folder.                             |
| [Repair-OutlookDataFile.ps1](Repair-OutlookDataFile.ps1) | Open Microsoft's repair utility for a PST or OST file.                                  |

Examples below run from the repository root. Each script provides complete parameter help:

```powershell
Get-Help .\scripts\Office\Switch-OfficeVersion.ps1 -Full
Get-Help .\scripts\Office\New-OutlookArchive.ps1 -Full
```

## Requirements

Install the repository dependencies with `.\winkit.ps1 init`. Scripts declare their minimum PowerShell and PSFoundation versions in
`#Requires`. All support `-DryRun`, `-WhatIf`, and `-PassThru`; previews and required privileges depend on the operation.

Office migration requires elevated PowerShell and uses 64-bit PowerShell on a 64-bit OS, including when the installed Office suite is
32-bit. Outlook profile operations require an interactive user session and PowerShell matching Outlook's architecture. These are different
requirements: Outlook 2007 profile operations use 32-bit PowerShell, while migrating that installation on 64-bit Windows uses 64-bit
PowerShell.

Office desktop applications and these profile operations are not supported on Server Core. Back up affected data before making changes.

## Office migration

`Switch-OfficeVersion.ps1` uses the Microsoft Office Deployment Tool (ODT). It requires an explicit mode and destination; it does not select
or purchase a license, upgrade Windows, convert Outlook profiles, or provide automatic rollback.

### Modes

| Mode      | Behavior                                                                                                                       |
| --------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `Check`   | Read installed Click-to-Run and MSI Office registrations. With `-TargetProductId`, also report the target's activation status. |
| `Prepare` | Download the selected destination and record its build, configuration, file sizes, and SHA256 hashes in a media manifest.      |
| `Migrate` | Validate and stage prepared media locally, remove approved source products, install the destination, and verify the result.    |

`Check`, `-DryRun`, and `-WhatIf` do not write files, download media, terminate applications, launch installers, or change licensing.
Migration previews still require valid ODT and prepared media and must pass preflight checks. Actual Prepare and Migrate operations use
high-impact confirmation. Use `-Confirm:$false` for a reviewed unattended deployment.

### Supported installations

Sources include MSI Office 2007, 2010, 2013, and 2016, including Office 2007 Enterprise, and registered Click-to-Run products selected by
their exact product IDs. Unsupported or unknown MSI generations are blocked. Destinations are:

| Product family     | `TargetProductId`                         | Channel                                                        |
| ------------------ | ----------------------------------------- | -------------------------------------------------------------- |
| Office 2019 volume | `Standard2019Volume`, `ProPlus2019Volume` | `PerpetualVL2019`                                              |
| Office LTSC 2021   | `Standard2021Volume`, `ProPlus2021Volume` | `PerpetualVL2021`                                              |
| Office LTSC 2024   | `Standard2024Volume`, `ProPlus2024Volume` | `PerpetualVL2024`                                              |
| Microsoft 365 Apps | `O365ProPlusRetail`, `O365BusinessRetail` | `Current` by default; also `MonthlyEnterprise` or `SemiAnnual` |

The volume channel is derived from the destination. `-Architecture` accepts `32` or `64` and defaults to `64`. `-Language` accepts a list of
explicit language IDs and defaults to `de-de`. `-Version` optionally pins an exact `16.0` build; otherwise Prepare records the build it
downloads. `-ExcludeApp` omits selected destination applications; `-ExcludePublisher` adds Publisher to that list.

Verify the operating system, licensing entitlement, add-ins, VBA, and Outlook compatibility for the selected destination. The script's
Windows version guard is not a complete vendor support matrix, and availability of a product ID does not establish current vendor support.
Perform the migration on a recoverable pilot workstation before broader deployment.

### Removal scope

Run Check before choosing removal options:

```powershell
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check -PassThru
```

- `-RemoveProductId` names the exact installed Click-to-Run products to remove. Unapproved additional products, such as Visio or Project,
  stop the migration. Stale removal selections also stop it; refresh the list after a partial migration.
- `-RemoveMsi` authorizes removal of **all ODT-supported MSI Office products**, including supported Visio, Project, language packs,
  runtimes, and database engines. It does not select only the Office suite. Review applications that depend on these components first.
- Coexisting products are not automatically preserved. If the destination is already installed with other Office products, or with a
  different architecture, resolve the installation manually.

If the destination and architecture are already installed alone, Migrate verifies activation and skips installation. It does not reapply
languages, exclusions, channel, build, or a supplied key, and it does not serve as a repair or update command.

### Prepare installation media

Obtain an official Microsoft ODT `setup.exe`. The script validates its signature, publisher, executable metadata, and minimum version. Use a
dedicated local or UNC directory with no existing `Office` subdirectory. Restrict write access to deployment administrators and use a fresh
directory after an interrupted download.

```powershell
$target = @{
  TargetProductId = 'Standard2024Volume'
  SourcePath      = '\\srv\deploy\Office2024'
  OdtPath         = 'C:\ODT\setup.exe'
  Architecture   = '64'
  Language       = @('de-de')
}

.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Prepare @target
```

Use the same target settings on each workstation. Migration requires the generated `winkit-office-media.json` manifest and matching
payloads. It copies and verifies the media locally before removing Office. Staging requires twice the media size plus 4 GiB free on the
ProgramData drive; this allowance is not an exact installation-size estimate. Keep the source unchanged and accessible while staging. Hashes
detect changes to the prepared media; they do not replace access controls on the media and manifest.

### Migrate a workstation

Close Office applications in all sessions, complete pending reboots, back up user data, and retain the previous installation media and
licenses for recovery. Do not run concurrent Office deployments. The following examples use the `$target` settings above.

For Office 2007 Enterprise or another supported MSI source:

```powershell
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target -RemoveMsi -DryRun

# Run after reviewing the inventory, removal scope, and preview.
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target -RemoveMsi -PassThru
```

For a Click-to-Run source, specify the ID reported by Check:

```powershell
$source = @{ RemoveProductId = @('HomeBusiness2019Retail') }

.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source -DryRun
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source -PassThru
```

Running Office applications stop migration by default. `-ForceCloseApps` explicitly allows their termination across sessions and can discard
unsaved work. The script never schedules a reboot.

### Activation and results

Volume destinations use default KMS licensing when no key is supplied. To supply a MAK, obtain it as a SecureString before migration:

```powershell
$mak = Read-Host 'MAK for the destination' -AsSecureString
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target -RemoveMsi -ProductKey $mak -PassThru
```

ODT requires the key in temporary XML. The directory is restricted to Administrators and SYSTEM, the XML is removed in cleanup, and the key
is not passed on the process command line. This does not guarantee secure deletion or control the contents of ODT's own logs.

The script verifies installed product, architecture, build, and removal results. Volume activation must match the destination and report a
licensed status. Microsoft 365 returns `UserActivationRequired`; complete activation in the licensed user's session. Existing product keys
are not automatically removed.

Approved Prepare/Migrate operations write transcripts under `%ProgramData%\OfficeMigration` by default; use `-LogRoot` to choose another
directory. `-PassThru` returns structured results with the available inventory, activation status, failure phase, and reboot requirement.

| Exit code | Meaning                                                                                                                               |
| --------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| `0`       | Inventory, preparation, installation, or preview completed. Check the result's status; subscription activation may still be required. |
| `1`       | Preflight, execution, cleanup, or verification failed. Office may be partially migrated.                                              |
| `3010`    | Installation verified and a reboot is required. MSI removal always reports this requirement after a successful migration.             |

A reboot request during Click-to-Run removal stops installation and reports a failure with `RebootRequired`. Reboot, run Check, and refresh
the removal plan. If installation succeeds but volume activation is not verified, resolve KMS/MAK activation separately. A failed
installation after removal may require manual recovery. Inspect the reported phase, transcript, ODT logs, and any temporary directory
reported by a cleanup error.

After any requested reboot, check the installed inventory and target activation:

```powershell
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check -TargetProductId Standard2024Volume -PassThru
```

Microsoft references:
[ODT configuration](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options),
[MSI removal scope](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/upgrade-from-msi-version), and
[Office LTSC 2024 deployment](https://learn.microsoft.com/en-us/office/ltsc/2024/deploy).

## Outlook

The Outlook scripts work with desktop Outlook profiles or local data files. Profile operations use Outlook's COM object model and need
classic desktop Outlook, an interactive profile, and PowerShell matching Outlook's architecture. Outlook 2007 is 32-bit only. Run as the
profile's user; SYSTEM is not suitable for these operations.

### Compatibility and session requirements

| Script                       | Minimum classic Outlook         | Requirement                                                                |
| ---------------------------- | ------------------------------- | -------------------------------------------------------------------------- |
| `New-OutlookArchive.ps1`     | 2007 (12)                       | Unicode PST creation through `NameSpace.AddStoreEx`.                       |
| `Optimize-Outlook.ps1`       | 2007 (12)                       | Transport-header access through `Item.PropertyAccessor`.                   |
| `New-TestOutlookMessage.ps1` | 2007 (12)                       | Outlook object model and optional registered Redemption component.         |
| `Repair-OutlookDataFile.ps1` | Office 12 tool discovery onward | An available ScanPST/ScanOST executable; it does not connect to a profile. |

The three profile scripts reject Outlook versions below 12 before processing mail. Repair instead validates the data file and locates a
repair utility; it does not enforce a client-version check. These are script compatibility requirements, not a guarantee of vendor support
or PST health. New Outlook is not a target for these COM scripts: being able to open a PST does not establish automation compatibility. See
[Microsoft's Outlook automation guidance](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/vba-alternatives).

Run profile operations as the logged-in Outlook user at the same elevation as Outlook. For an Outlook 2007 rehearsal, use Windows PowerShell
5.1 x86, especially with 32-bit Redemption. On 64-bit Windows, its executable is:

```text
C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
```

An appropriate 32-bit PowerShell 7 host is another option on an OS that supports it; 64-bit Outlook 2010 or later uses 64-bit PowerShell.
Although the scripts use PowerShell 5.0-compatible syntax, the pinned PSFoundation module requires PowerShell 5.1. Run initialization and
verify module visibility in the intended host. Modules installed only for another PowerShell edition, architecture, or user may not be
available there; the exact dependency versions are in [requirements.psd1](../../requirements.psd1).

`-StoreName` selects a store by its display name. If omitted, the default delivery store is used. Use a unique display name and `-Verbose`
to inspect detected stores; duplicate display names are rejected. The scripts reuse the running/default Outlook session, do not select a
profile, and do not switch Outlook offline. Leave `-QuitOutlook` off when Outlook is already open interactively; previews do not quit it.

Previews avoid the requested mail changes or repair-tool launch. Outlook profile scripts can still connect to Outlook, and scripts may write
operation logs or an explicitly requested CSV report. Use closed-file backups and sufficient free space before processing real mail.

The scripts do not send mail, access address books, or use `Recipients`, avoiding those common Outlook Object Model Guard triggers. Profile,
credential, and security dialogs can still occur, including during `-DryRun`. Rehearse against a disposable profile first.

### Offline rehearsal and account migration

Open Outlook manually, select **Work Offline**, confirm its status, and cancel outstanding credential dialogs before a local rehearsal. In
Outlook 2007, Work Offline is available from the File menu; Microsoft's
[Outlook 2007/2010 offline article](https://support.microsoft.com/en-au/topic/outlook-2007-2010-status-is-always-offline-and-can-t-receive-or-send-mail-normally-easy-fix-articles-3e977242-7f66-2f0a-5956-51007986b38f)
also describes the status-bar control. Once dependencies are available, disconnecting networking can keep obsolete account connections out
of a local test. It does not repair the account configuration or dismiss an already-open dialog.

Local PST archiving does not require authenticating an old POP mailbox. Before deleting an account or profile to stop prompts, capture its
data paths and backups. Configure a replacement IMAP account separately and retain the original POP PST as a local data source. A
server-side mailbox migration does not prove that every locally downloaded POP message exists on the IMAP server; verify that before
disposing of the local data.

### Archive mail

`New-OutlookArchive.ps1` creates a new local Unicode PST, mirrors mail-folder hierarchy, and copies or moves messages into it. Contacts,
calendars, tasks, and search folders are skipped. The destination PST must not already exist; use a new path for every run.

The destination's parent directory must already exist, and source and destination must be distinct stores. Root-level mail is included;
non-mail folder subtrees and virtual search folders are skipped. Use `-DisplayName` to choose the mounted archive's display name.

```powershell
$archive = @{
  ArchivePath = 'D:\Archive\mail-2024.pst'
  StoreName   = 'user@example.com'
  StartDate   = '2024-01-01'
  EndBefore   = '2025-01-01'
  Mode        = 'Copy'
}

.\scripts\Office\New-OutlookArchive.ps1 @archive -DryRun -PassThru
.\scripts\Office\New-OutlookArchive.ps1 @archive -PassThru
```

`Copy` is the default and leaves source messages intact. It temporarily duplicates each message in its source store before moving the copy
into the archive, so allow space in both stores. `Move` removes successfully archived messages from the source.

`StartDate` is inclusive. Prefer the exclusive `EndBefore` bound for whole days or years. `EndDate` is inclusive of the exact supplied time;
a date without a time means midnight. `EndDate` and `EndBefore` cannot be combined. The archive is detached from the profile by default; use
`-DetachWhenDone:$false` to keep it mounted. Close Outlook before copying the PST elsewhere.

### PST migration and archive rehearsal

Before a production Move run, complete the scratch integration suite and manual archive checks on the actual Outlook installation. Simulated
tests cannot certify Outlook 2007 COM behavior or the health of an existing PST. Rehearse on disposable data first, then on a working copy
of the real PST.

#### Prepare recoverable copies

1. Record current PST paths and display names, Outlook version and architecture, account delivery locations, and folder counts. Record
   local-only contacts, calendars, tasks, rules, signatures, and other settings. The mail archiver does not migrate settings or non-mail
   data.
2. Close Outlook and confirm `OUTLOOK.EXE` has exited. Copy every relevant PST to a separate backup location. Compare source and backup with
   `Get-FileHash -Algorithm SHA256` while Outlook remains closed. Preserve one verified backup unopened and unchanged.
3. Make a separate rehearsal copy. Open it in classic Outlook under a distinct store name; avoid attaching copies with indistinguishable
   display names. Keep active PSTs on local disk outside synchronization folders, and do not process a network PST. See
   [Microsoft's network PST limitations](https://learn.microsoft.com/en-us/troubleshoot/outlook/data-files/limits-using-pst-files-over-lan-wan).
4. Allow disk space for the untouched backup, working copy, and all archives. There is no automatic archive splitting, source size-limit
   detection, or reliable estimate of destination growth. Use small date batches and check file sizes and available capacity between runs.

Outlook 2007 defaults to a 20 GB Unicode PST maximum and a 19 GB data threshold. Registry settings can alter these limits; the scripts
neither inspect nor change them. Check the actual format and configured limits before adding data. See
[Microsoft's PST/OST size-limit documentation](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/data-files/configure-size-limit-outlook-data-files).

**Copy writes to the source PST.** Outlook temporarily creates the duplicate there before transferring it to the archive, so Copy is
unsuitable when the source has no write headroom. A failed transfer can leave that temporary duplicate in the source. The script stops
rather than deleting uncertain data; inspect and reconcile both stores before retrying. For a PST near its Outlook 2007 limit, rehearse on a
working copy under the target classic Outlook installation with sufficient configured capacity. Do not raise limits or start a destructive
Move merely to bypass a failed Copy rehearsal.

#### Run and inspect a rehearsal

From the repository root in the intended PowerShell host:

```powershell
.\winkit.ps1 init
.\winkit.ps1 test -Outlook
```

Use a disposable profile whose default store is a scratch PST. The suite attaches its own uniquely named scratch store to the active
profile, detaches it afterward, and retains `%TEMP%\winkit-outlook-test-<run-id>` for inspection. It does not create or select a separate
profile for you. Failed archive/count assertions are a stop condition. A skipped header/date-dependent check is not a pass; complete missing
checks with real dated mail in a disposable PST or an appropriately licensed Redemption installation.

Choose a unique working-copy store name and a new destination in an existing directory:

```powershell
$rehearsal = @{
  StoreName   = 'PST - WORKING COPY'
  ArchivePath = 'D:\MailArchive\before-2024-rehearsal.pst'
  EndBefore   = [datetime]'2024-01-01'
  Mode        = 'Copy'
}

.\scripts\Office\New-OutlookArchive.ps1 @rehearsal -DryRun -PassThru -Verbose
.\scripts\Office\New-OutlookArchive.ps1 @rehearsal -PassThru -Confirm
```

Reopen the archive in Outlook and compare eligible per-folder counts, first and last dates, Sent Items, nested folders, message bodies, and
representative attachments. Inspect the source too. Success counters and the absence of failed results alone do not establish archive
integrity; declined transfers are not completed transfers. Calendar, contact, and task data remain in the source and need separate
migration. Close Outlook before copying the archive file elsewhere; detaching a store is not proof that every process has released its file.

#### Reduce the production PST

After rehearsal and backup verification, use a different new archive path with `-Mode Move`. An existing Copy archive cannot be reused as
the Move destination. Copy and Move archives can contain overlapping mail, so label rehearsal artifacts clearly. Preview each disjoint date
batch first. On any failure, inspect the operation log and both PSTs and reconcile the partially completed batch before retrying. There is
no resume ledger, transactional rollback, or deduplication of previous archives.

Moving mail out does not necessarily shrink the physical PST immediately. Compact only after validating the archive and taking another
recoverable backup, using Outlook's data-file settings. See
[Microsoft's compaction guidance](https://support.microsoft.com/en-us/outlook/reduce-the-size-of-your-mailbox-and-outlook-data-files-pst-and-ost).
Keep the original backup until mail, contacts, calendar, attachments, and send/receive behavior are verified on the new installation. Use
the repair script only when a scan or repair is needed, with Outlook closed and a preserved backup.

### Review duplicate messages

`Optimize-Outlook.ps1` compares transport Message-IDs within each mail folder. The first occurrence is kept; subsequent occurrences move to
the top-level `_Duplicates_Review` folder by default. It does not hard-delete messages, compare subjects or bodies, or remove the same
message from different folders. Messages without a usable Message-ID are skipped.

Message-IDs are compared ordinally and case-sensitively, without comparing bodies or attachments. The retained item is the first encountered
during traversal, not necessarily the oldest or newest message. Exclusions apply to complete subtrees; the review folder, search folders,
and non-mail folder subtrees are skipped. Candidates are consolidated into one review folder in the same store. Deduplication does not
compact a PST or free its storage by itself, and it is a separate reviewed task rather than a migration prerequisite.

```powershell
.\scripts\Office\Optimize-Outlook.ps1 -StoreName 'user@example.com' -ReportPath .\dedup-preview.csv -DryRun
.\scripts\Office\Optimize-Outlook.ps1 -StoreName 'user@example.com' -ReportPath .\dedup-run.csv -PassThru
```

Review the preview and resulting review folder before deleting anything manually. `-ReviewFolderName` changes the destination;
`-ExcludeFolders` replaces the default list of folder display names to skip. Defaults include Deleted Items, Junk Email/Junk E-mail, Outbox,
Sync Issues, Conflicts, Local Failures, and Server Failures. Localized folder names may require an explicit list.

### Create test messages

`New-TestOutlookMessage.ps1` creates deterministic synthetic messages in `WinkitTestData` by default. It does not send email. Use a
dedicated test profile or store; rerunning creates additional messages rather than replacing earlier test data.

```powershell
$messages = @{
  Count            = 20
  Seed             = 42
  DuplicateRatio   = 0.25
  StoreName        = 'Test Mail'
  TargetFolderName = 'WinkitTestData'
}

.\scripts\Office\New-TestOutlookMessage.ps1 @messages -DryRun
.\scripts\Office\New-TestOutlookMessage.ps1 @messages -UseRedemption -PassThru
```

`-DuplicateRatio` controls the fraction of items reusing an earlier Message-ID. Optional `-StartDate` and `-EndDate` bound synthetic
received times. Native Outlook transport-header and received-time writes are best-effort. For reliable Message-ID and backdated-time
injection, install the Redemption component and use `-UseRedemption`; this mode fails if `Redemption.RDOSession` is not registered. Check
header-injection results before using the messages to validate deduplication.

Successful fixture injection requires both the header and received-time writes; inspect `HeaderInjected` and verify persistence in the test
store. Redemption reuses Outlook's MAPI session rather than selecting another profile. Check
[Redemption's licensing](https://www.dimastr.com/redemption/) for your use; do not generate fixtures in a production mail store.

### Repair a data file

`Repair-OutlookDataFile.ps1` locates and launches ScanPST or ScanOST for the specified file. `Auto` chooses ScanPST for PST files and
prefers ScanOST for OST files when available, otherwise ScanPST. `-Tool` chooses a utility explicitly; `-ToolPath` supplies its executable.

```powershell
.\scripts\Office\Repair-OutlookDataFile.ps1 -Path 'D:\Mail\archive.pst' -DryRun
.\scripts\Office\Repair-OutlookDataFile.ps1 -Path 'D:\Mail\archive.pst' -Tool ScanPST -PassThru
```

Back up the file and close Outlook before repair. The script checks for running Outlook and attempts exclusive access to the data file. The
Microsoft utility may require interactive input. Paths containing spaces are supported, and nonzero tool exits are reported as failures. A
successful tool exit does not certify PST or OST health or establish that a repair was performed; inspect the utility's scan results and
log.

## Testing

Run `.\winkit.ps1 test` for logic and mocked safety tests without Outlook. Outlook regression tests cover collection mutation, transfer
failures, date boundaries, store selection, preview behavior, excluded subtrees, repair-path quoting, and locks. Transport Message-ID
parsing belongs to PSFoundation and is tested in that module's repository. Real Office deployment and activation require a recoverable pilot
workstation.

The optional `.\winkit.ps1 test -Outlook` suite generates deterministic subjects, Message-IDs, dates, and seeded duplicates in its scratch
store. Archive checks reopen output PSTs and compare actual mail counts. These tests require the intended Outlook installation; logic tests
alone cannot establish real COM compatibility, available PST capacity, fixture-property persistence, or archive integrity. See the
[Outlook integration testing guide](../../tests/Office/README.md) for profile setup, assertions, retained artifacts, and requirements.
