# Office

Scripts for deploying Microsoft Office and maintaining Outlook mail stores and data files on Windows.

| Script                                                   | Purpose                                                                                 |
| -------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| [Install-Office.ps1](Install-Office.ps1)                 | Prepare, install, verify, or recover an Office deployment.                              |
| [Remove-Office.ps1](Remove-Office.ps1)                   | Inventory or remove selected Click-to-Run Office products.                              |
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

Office deployment requires elevated PowerShell and uses 64-bit PowerShell on a 64-bit OS, including when the installed Office suite is
32-bit. Outlook profile operations require an interactive user session and PowerShell matching Outlook's architecture. These are different
requirements: Outlook 2007 profile operations use 32-bit PowerShell, while migrating that installation on 64-bit Windows uses 64-bit
PowerShell.

Office desktop applications and these profile operations are not supported on Server Core. Back up affected data before making changes.

All Office scripts show progress for their longer operations. Outlook archive and deduplication identify the current folder and item
processing phase; synthetic message generation shows the number processed. Item updates are throttled to avoid slowing large batches.
Deployment scripts show inventory, configuration, compatibility, media preparation, activation, execution, or recovery phases as applicable.
Their progress is indeterminate while a deployment call is running; it does not estimate download or installation completion percentages.
Data-file repair shows tool discovery and waiting status; follow ScanPST/ScanOST's own window for scan progress and any required
confirmation.

Progress displays clear when operations finish or fail. Set `$ProgressPreference = 'SilentlyContinue'` to hide them in unattended sessions;
ordinary status messages, warnings, and returned results remain available.

## Office deployment

The deployment scripts require **PSFoundation 1.6.1 or later** and an existing Microsoft-signed Office Deployment Tool (ODT) setup.exe. They
share PSFoundation's inventory, planning, media validation, execution, and recovery APIs. They do not purchase licenses, upgrade Windows,
convert Outlook profiles, or provide automatic rollback.

### Modes and responsibilities

| Script               | Mode    | Behavior                                                                                      |
| -------------------- | ------- | --------------------------------------------------------------------------------------------- |
| Install-Office       | Check   | Read inventory; with a target, assess installation eligibility and activation.                |
| Install-Office       | Prepare | Download and verify a reusable installation package.                                          |
| Install-Office       | Install | Install on a clean machine, or return an independently verified compliant no-op.              |
| Install-Office       | Recover | Verify or continue installation using its protected recovery journal.                         |
| Remove-Office        | Check   | Read inventory; with selected product IDs, assess removal eligibility.                        |
| Remove-Office        | Remove  | Remove exactly the selected supported Click-to-Run products.                                  |
| Switch-OfficeVersion | Check   | Read inventory; with a target/removal selection, assess migration eligibility and activation. |
| Switch-OfficeVersion | Prepare | Prepare media using the same API as Install-Office.                                           |
| Switch-OfficeVersion | Migrate | Stage the destination, remove approved sources, install, and verify.                          |
| Switch-OfficeVersion | Recover | Verify or continue the recorded migration without expanding removal authority.                |

Mode is mandatory. Check without a target/selection is inventory-only. A target Check returns a plan with State, Eligible, Blockers, and
language transitions; a blocked plan exits with code 1. Supply SourcePath when assessing a deployment that needs media. Check does not
perform every execution-time check: the module revalidates the host, applications, inventory, and media during deployment.

Check and previews do not download, write logs/journals, stop applications, launch installers, or change Office/licensing. -DryRun also sets
WhatIfPreference. Both -DryRun and -WhatIf return results even without -PassThru. Previews still run applicable validation and can return
blockers. Execution and preparation use high-impact confirmation; use -Confirm:$false for reviewed unattended runs. This does not bypass
validation or suppress language warnings.

Install has no removal parameters. Conflicting or incomplete installations are not automatically reconfigured. A compliant no-op requires
verification of the full configuration, including languages and application selection, and does not reapply a supplied key.

### Backend availability

The current native execution backend targets elevated **x64 Windows 11 desktop** hosts, including deployment of 32-bit Office there. Server,
ARM, and ordinary Windows 10 execution are outside that gate. The narrowly scoped pilot below is the only Windows 10 exception. Product IDs
identify targets; their availability is not a vendor lifecycle or product/OS support guarantee. Check OS support, licensing, add-ins, VBA,
and Outlook compatibility for the destination.

Native inventory currently reports Languages and PrimaryLanguage as verification limitations. Install/Migrate plans return
UnsupportedNativeVerification before mutation when these postconditions cannot be verified. Installed-Office automatic locale sourcing also
remains blocked where this evidence is unavailable. Explicit language selection sets the target; it does not bypass native verification. The
scripts preserve these module blockers by default. PilotMigration waives only the named language limitations for its exact profile.

Standalone MSI removal is unsupported. Recovery supports pre-launch continuation, verification of a completed deployment, and certain
migration checkpoints after verified Click-to-Run removal. Uncertain partial-installer states return UnsupportedRecoveryState. Recover does
not implement Quick Repair, Online Repair, journal-free mutation, or rollback. Validate supported operations on disposable pilot machines
before fleet deployment.

### Office Enterprise 2007 to Standard 2019 pilot

`Switch-OfficeVersion.ps1 -PilotMigration` requires **PSFoundation 1.6.1 or later** and is available in **Check and Migrate only**. The
wrapper passes this explicit authorization to both planning and execution; ordinary migrations retain their strict defaults.

The profile is restricted to x64 Windows 10 desktop build 19045, the reported Enterprise 2007 MSI suite/resources, and Standard2019Volume
x64 on PerpetualVL2019. German UI must be selected explicitly. Installed-office auto-discovery remains unavailable. The plan separately
records German, English, French and Italian companion proofing intent, which still requires post-install review. Additional full UI packs
are not silently installed. This pilot does not establish vendor support for Office 2019 or readiness for unattended workforce deployment.

Use a fresh ABB backup plus a short-lived pre-migration VM snapshot, with working hypervisor console/revert access. Keep users off the VM
through acceptance; reverting loses subsequent guest changes. Preparing media may be done before the snapshot. Example from the winkit root:

```powershell
$target = @{
  TargetProductId = 'Standard2019Volume'
  Architecture = '64'
  Language = @('de-de')
  SourcePath = 'C:\Media\Office2019'
}
$odt = 'C:\ODT\setup.exe' # Existing verified Microsoft ODT
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Prepare @target -OdtPath $odt -PassThru
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check @target -RemoveMsi -PilotMigration -PassThru
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target -OdtPath $odt -RemoveMsi -PilotMigration -DryRun -PassThru
# After reviewing the plan, removal scope, and rollback point:
$result = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target -OdtPath $odt `
  -RemoveMsi -PilotMigration -PassThru -Confirm
$result | ConvertTo-Json -Depth 30
```

Use the same ExcludeApp/ExcludePublisher settings for preparation, Check and Migrate if desired. Standard is a different application set
from Enterprise; review required applications and 32-bit add-in compatibility before accepting the migration. Prepare has no pilot flag. Do
not use Recover for this pilot: its schema-2 journal is inspectable evidence, and the module refuses replay.

Native exit 0 or 3010 with unresolved observations returns **AppliedUnverified**, wrapper exit **1**, and detailed `Verification.Unknowns`.
Native codes and `RebootRequired` survive unchanged; unverified exit 1 takes precedence over 3010. This is a manual review point, never a
signal to rerun the migration. Known mismatches and native errors remain failures. Journal/JSONL paths are in the result. Save these, the
result JSON, a fresh collector report and relevant native ODT logs **off the VM before any revert**; native logs may contain secrets. Check
Standard 2019 x64/build, German UI, all four proofing languages, activation, applications, add-ins and the existing Outlook profile. Reboot
explicitly if required and collect evidence again. Do not change the reference 2019 workstation.

### Products, languages, and configuration

| Product family     | TargetProductId                       | Channel                                             |
| ------------------ | ------------------------------------- | --------------------------------------------------- |
| Office 2019 volume | Standard2019Volume, ProPlus2019Volume | PerpetualVL2019                                     |
| Office LTSC 2021   | Standard2021Volume, ProPlus2021Volume | PerpetualVL2021                                     |
| Office LTSC 2024   | Standard2024Volume, ProPlus2024Volume | PerpetualVL2024                                     |
| Microsoft 365 Apps | O365ProPlusRetail, O365BusinessRetail | Current by default; MonthlyEnterprise or SemiAnnual |

-Architecture accepts 32 or 64 and defaults to 64. Volume channels are derived from the product. -Version optionally selects an exact 16.0
build; preparation otherwise resolves and pins a build. -ExcludeApp selects omitted applications; -ExcludePublisher adds Publisher.
PSFoundation validates allowed configuration values.

Language defaults to exactly **en-us**, independently of Windows or the execution account. Use -Language de-de for German, or an ordered
list such as -Language en-us,de-de. Order is preserved; the first language is the primary shell language. This does not change Windows
locale, keyboard layouts, or individual users' Office editing/display preferences.

ODT uses the same deployment engine with XML specifying products and languages. Media must contain the required language payloads. An
English/German package may serve an English-only, German-only, or bilingual request without installing every available language.

Automatic discovery requires -AutoSourceLocales. Its default source is InstalledOffice, requiring unambiguous installed-language and
primary-language evidence. -LocaleSource OperatingSystem instead reads the machine installation UI language, not the administrator's
culture. LocaleSource requires AutoSourceLocales; explicit Language and automatic sourcing are mutually exclusive. Discovery failures are
not silently replaced by en-us or OS language detection.

Switch reports language changes or unknown source languages in plans and warnings. Review additions, removals, and primary-language changes
before confirming. Explicitly select -Language de-de for German legacy MSI migrations where automatic preservation is unavailable. Include
needed language/proofing resources in the deployment plan; complete legacy preservation cannot be inferred from a product LCID.

### Prepare reusable media

Obtain an official Microsoft ODT setup.exe. PSFoundation checks its Microsoft signature and tool metadata. Use a dedicated local or UNC
package directory whose parent already exists. Valid compatible packages can be verified and reused; incompatible or incomplete packages
require a new directory.

The package and manifest must have Administrators/SYSTEM ownership and protected write access. Restrict share access appropriately and
ensure the actual execution identity can reach UNC media. SYSTEM or remote sessions may lack the operator's network access or mapped drives.
Keep the package unchanged while it is being staged.

```powershell
$target = @{
  TargetProductId = 'Standard2024Volume'
  Architecture   = '64'
  SourcePath     = '\\srv\deploy\Office2024'
  OdtPath        = 'C:\ODT\setup.exe'
}

# Include both language payloads in a reusable package.
.\scripts\Office\Install-Office.ps1 -Mode Prepare @target -Language en-us,de-de -PassThru
```

Preparation publishes a schema-2 psfoundation-office-media.json manifest with the build, available languages, tool version, payload sizes,
and hashes. It does not publish a partial download as ready. Old schema-1 winkit-office-media.json packages require preparation into a new
directory. Hashes detect changed payloads; they do not authenticate a manifest that an attacker can also replace.

Deployment verifies a protected local copy before installation/removal. Staging requires twice the media size plus 4 GiB free on the staging
drive; this is an allowance, not an exact installed-size estimate. Missing language payloads stop validation. Deployment does not silently
download missing files from the CDN or select a different build.

### Install a clean workstation

Close Office applications across sessions, complete pending reboots, and avoid concurrent deployments. These calls use the package settings
above. With Language omitted, the requested installation is English even when the package also contains German.

```powershell
.\scripts\Office\Install-Office.ps1 -Mode Check -PassThru

$check = @{
  TargetProductId = $target.TargetProductId
  SourcePath     = $target.SourcePath
}

.\scripts\Office\Install-Office.ps1 -Mode Check @check -PassThru
.\scripts\Office\Install-Office.ps1 -Mode Install @target -DryRun

# Run after reviewing the plan and applicable backend limitations.
.\scripts\Office\Install-Office.ps1 -Mode Install @target -Confirm:$false -PassThru
```

Use -Language de-de on target Check and Install calls for German. Running applications block execution unless -ForceCloseApps explicitly
authorizes termination, which can discard unsaved work. The module uses a shared deployment lock and checks native deployment activity. It
never schedules a reboot.

### Migrate an existing installation

Back up user data and retain previous installation media and licenses. Inventory before selecting removals:

```powershell
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check -PassThru
```

- -RemoveProductId names exact installed Click-to-Run IDs authorized for removal. Unapproved additional products, such as Visio or Project,
  block migration. Stale selections also block it; refresh the plan after partial work.
- -RemoveMsi authorizes **all supported MSI Office removals**, including supported Visio, Project, language packs, runtimes, and database
  engines. It does not mean only the suite. Review dependent applications. Unknown or unsupported MSI components block the plan.
- Installation and migration are separate operations. A desired configuration change on an existing target requires explicit planning; it is
  not an automatic repair, update, or architecture conversion by Install.

MSI sources recognized for migration include Office 2007, 2010, 2013, and 2016, including Office 2007 Enterprise. Recognition does not
override the host and verification gates above.

```powershell
# Explicitly request German for an MSI source.
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target -RemoveMsi -Language de-de -DryRun
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target -RemoveMsi -Language de-de -PassThru

# Select Click-to-Run products from the inventory.
$source = @{ RemoveProductId = @('HomeBusiness2019Retail') }

.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source -DryRun
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source -PassThru
```

The module validates and stages destination media before removal, then rechecks inventory and applications. Approved MSI removal is part of
the destination ODT configuration. A removal failure or reboot requirement stops continuation; inspect the result and recovery record before
retrying. Failure after removal may require manual recovery.

### Remove selected products

Remove-Office supports selected Click-to-Run products and their installed languages. It has no default selection and no RemoveMsi switch.
Unselected products must remain verifiably unchanged; uncertain shared-component effects can block removal. Already absent selected products
return AlreadyAbsent with Changed=false.

```powershell
$removal = @{
  RemoveProductId = @('O365ProPlusRetail')
  OdtPath         = 'C:\ODT\setup.exe'
}

.\scripts\Office\Remove-Office.ps1 -Mode Check -RemoveProductId $removal.RemoveProductId -PassThru
.\scripts\Office\Remove-Office.ps1 -Mode Remove @removal -DryRun
.\scripts\Office\Remove-Office.ps1 -Mode Remove @removal -Confirm:$false -PassThru
```

The scripts do not implement custom cleanup of profiles, PSTs, user documents, or product keys. Keep backups before removing software.

### Recover a recorded deployment

Use the RunId and LogRoot from the original operation. Recover reads the protected local journal and rechecks state and media. Its target,
languages, build, and removal scope come from the recorded operation. Target/removal overrides are rejected, including explicit Language or
AutoSourceLocales: recovery never redetects languages.

```powershell
$recovery = @{
  RunId   = '0123456789abcdef0123456789abcdef'
  OdtPath = 'C:\ODT\setup.exe'
  LogRoot = 'C:\ProgramData\PSFoundation-Office'
}

.\scripts\Office\Install-Office.ps1 -Mode Recover @recovery -DryRun
.\scripts\Office\Install-Office.ps1 -Mode Recover @recovery -PassThru

# Use Switch instead for an original migration journal.
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Recover @recovery -DryRun
```

Replace the example RunId with the actual identifier and use the matching script for the recorded action. Installation recovery cannot
inherit migration removal authority. Supply a needed MAK again as SecureString; journals never contain it. Changed media, unexpected
products, pending reboots, active deployments, and unsupported interrupted phases can block recovery. Repeatedly invoking Recover does not
make an unsupported partial install safe.

### Activation, logging, and automation results

Volume installations use default KMS licensing without a supplied key. Supply a MAK only as SecureString:

```powershell
$mak = Read-Host 'MAK for the destination' -AsSecureString
.\scripts\Office\Install-Office.ps1 -Mode Install @target -ProductKey $mak -PassThru
```

ODT needs the key in temporary XML. PSFoundation restricts that directory to Administrators/SYSTEM and removes the XML during cleanup; the
key is not passed on the process command line. This does not guarantee secure erasure or redaction of ODT's own logs.

Installation verification and activation are separate. Volume licensing must match the destination. Microsoft 365 reports
UserActivationRequired and needs activation in the licensed user's session. Resolve activation-only failures separately rather than
automatically reinstalling. Existing keys are not automatically removed.

Deployment operations keep protected journals and JSONL result logs under %ProgramData%\PSFoundation-Office by default. -LogRoot selects
another local root. Prepare returns a media assessment rather than a deployment journal; Check and previews write no operation logs.

-PassThru returns one structured outcome. Execution/recovery outcomes preserve PSFoundation's fields, including ReasonCode, Phase, Changed,
ChangeKnown, RebootRequired, NativeResults, Verification, Activation, RecoveryPath, LogPaths, and cleanup details. Changed=null with
ChangeKnown=false means the operation may have changed the machine. Check returns Inventory and an optional Plan; Prepare returns Media and
distinguishes MediaPrepared, AlreadyPrepared, and NotExecuted.

| Process exit code | Meaning                                                                                           |
| ----------------- | ------------------------------------------------------------------------------------------------- |
| 0                 | Completed, compliant/absent no-op, or preview. Inspect Status and activation.                     |
| 1                 | Blocked, failed, or unverified. Inspect ReasonCode, Phase, native outcomes, and possible changes. |
| 3010              | A reboot is required. This can also be a blocked recovery; inspect Status before continuing.      |

For automation, use explicit Mode, configuration, and -Confirm:$false; consume -PassThru objects rather than parsing console messages. Use
ConvertTo-Json -Depth 30 for nested results. Supply only parameters applicable to the selected mode: omit OdtPath from Check and deployment
settings from Recover. Missing mode-specific inputs fail rather than prompting partway through work.

Microsoft references:
[ODT configuration and languages](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options),
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
operation logs, archive JSON reports, or an explicitly requested CSV report. Use closed-file backups and sufficient free space before
processing real mail.

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

The progress display identifies connection, store selection, folder inspection, reading/filtering items, recording preview results, and
copying or moving messages. Item percentages apply to the current folder and phase; total folder counts are discovered during traversal.
Updates are throttled while processing items. Standard PowerShell `$ProgressPreference = 'SilentlyContinue'` suppresses the progress
display.

By default, each run writes a uniquely named, formatted UTF-8 JSON report under `%LOCALAPPDATA%\winkit\reports\Outlook` for the user running
the script. Use `-ReportDirectory` to choose another directory, or `-ReportPath .\archive-report.json` to specify an exact filename relative
to the current PowerShell location. Absolute report paths are also supported. The two parameters cannot be combined, and existing files are
never overwritten. The parent directory is created and a new report file is reserved before Outlook is opened; an unwritable report
destination stops the run before archive work. Reports are also written for previews and caught archive failures. They are finalized at the
end of the run; an interrupted process can leave an empty or incomplete report. Report-write failures produce a warning and exit code 1,
without undoing mail already copied or moved.

The console prints counts and the report path. `-PassThru`, `-DryRun`, and `-WhatIf` return one summary object containing `ReportPath`,
`Status`, `Preview`, `FoldersRead`, `FoldersSkipped`, `ItemsRead`, `ItemsMatched`, `Planned`, `Copied`, `Moved`, and `Failed`.
`FoldersSkipped` counts encountered search/non-mail folder roots, not every descendant in excluded subtrees. `ItemsRead` includes non-mail
items inspected in processed folders; `ItemsMatched` counts mail passing the date filters. A failed run's counts can be partial.

The report contains run timestamps, the source folder and filter settings, the summary, and a `Results` array with the individual operation
records. Each message record retains its subject (`Target`), folder (`Scope`), received date (`Received`, ISO 8601), action, status, and
detail. Preview entries use `Status = 'Skipped'` and `Detail = 'DryRun'`; no per-message WhatIf lines or result objects are printed. The
JSON report is the archive's detailed operation record. It contains message metadata, not message bodies or attachments.

```powershell
$summary = .\scripts\Office\New-OutlookArchive.ps1 @archive -DryRun -PassThru -ReportPath .\archive-report.json
notepad.exe $summary.ReportPath

# Read individual records for further analysis without flooding the console.
$report = Get-Content -LiteralPath $summary.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
$messages = $report.Results
```

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
batch first. On any failure, inspect the archive JSON report and both PSTs and reconcile the partially completed batch before retrying.
There is no resume ledger, transactional rollback, or deduplication of previous archives.

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

Office deployment wrapper tests cover mode validation, locale forwarding, confirmation/previews, recovery routing, and module result/exit
propagation. Native deployment, inventory, and media-verification tests belong to PSFoundation; wrapper tests never execute ODT.

The optional `.\winkit.ps1 test -Outlook` suite generates deterministic subjects, Message-IDs, dates, and seeded duplicates in its scratch
store. Archive checks reopen output PSTs and compare actual mail counts. These tests require the intended Outlook installation; logic tests
alone cannot establish real COM compatibility, available PST capacity, fixture-property persistence, or archive integrity. See the
[Outlook integration testing guide](../../tests/Office/README.md) for profile setup, assertions, retained artifacts, and requirements.
