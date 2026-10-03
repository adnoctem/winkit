# Outlook backup and repair

[Office overview](README.md) | [Requirements](Requirements.md)

Run examples from your **installed winkit directory**. Profile discovery and checkpoints run as the affected Outlook user, normally without
elevation. Use a separate backup destination with enough space and restricted access to personal data.

| Need                                                     | Tool                                                        |
| -------------------------------------------------------- | ----------------------------------------------------------- |
| Verified copies of entire PSTs                           | `Backup-Outlook.ps1`                                        |
| PST/OST preservation plus Office settings and signatures | `Checkpoint-Outlook.ps1`                                    |
| Portable mail archive from a live IMAP/Exchange store    | [Mail archival](Mail-Archival.md#imap-and-exchange-sources) |
| Microsoft's scan/repair utility                          | `Repair-OutlookDataFile.ps1`                                |

## Back up PST files

Backup copies whole PSTs, including mail, contacts, calendars, junk, and deleted items. It does not filter messages or export server
mailboxes. OST caches and stores without a PST path are skipped; selecting no usable PSTs fails.

```powershell
# Discover every attached PST. Omit AllStores for the default delivery store,
# or use StoreName for one uniquely named store.
.\scripts\Office\Backup-Outlook.ps1 -AllStores -Destination 'E:\OutlookBackups' -DryRun

# After preview, request graceful Outlook shutdown before copying.
$backup = .\scripts\Office\Backup-Outlook.ps1 -AllStores `
  -Destination 'E:\OutlookBackups' -QuitOutlook -PassThru
$backup | Format-List Status, Copied, Failed, BackupDirectory, ReportPath
```

For detached archives or explicit files, avoid opening Outlook entirely:

```powershell
$files = @(
  'D:\Mail\mail.pst'
  'D:\Archive\mail-2024.pst'
)
.\scripts\Office\Backup-Outlook.ps1 -PSTPath $files -Destination 'E:\OutlookBackups' -DryRun
$backup = .\scripts\Office\Backup-Outlook.ps1 -PSTPath $files -Destination 'E:\OutlookBackups' -PassThru

$manifest = Get-Content -LiteralPath $backup.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
$manifest | ConvertTo-Json -Depth 30
```

| Backup behavior    | Detail                                                                                                                                                                                                                            |
| ------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Discovery          | Reuses the interactive Outlook session. Releases COM references, then waits up to `WaitSeconds` (default 120) for Outlook to exit. Close it manually or use QuitOutlook; no force-kill.                                           |
| Explicit `PSTPath` | One or more literal filenames. Conflicts with StoreName, AllStores, QuitOutlook, IgnoreAdministrator, and WaitSeconds. No profile or Outlook installation required; close all file users first. Detached files require this mode. |
| Identity           | Discovery rejects elevation unless IgnoreAdministrator is deliberate for the same user. Direct copies have no Outlook identity requirement.                                                                                       |
| Locks              | Every source is opened exclusively before copying; handles remain held through verification. Any lock conflict stops the run.                                                                                                     |
| Destination        | Unique run directory, numbered filenames to avoid same-name collisions, no overwritten backups. Verified files end in `.pst`; failed copies retain `.partial`.                                                                    |
| Manifest           | Original paths, store names, byte counts, SHA-256 hashes, and results in `manifest.json`. Failure stops further copying but preserves verified files; manifest-write failures also fail the run.                                  |
| Verification       | Matching source/destination hashes establish copy integrity, not PST health. Keep a verified copy unopened; restore to a separate local working location before attaching.                                                        |
| Preview/output     | DryRun/WhatIf creates no files and does not close Outlook; discovery previews still connect. PassThru returns Status, Copied, Failed, BackupDirectory, ReportPath, Results. Progress covers discovery, waiting, copying, hashing. |

Account settings, rules outside the PST, and Windows profile settings are not captured by this PST-only tool.

## Capture data and settings

Checkpoint captures the current user's selected data files and Office settings. It rejects SYSTEM/service accounts. `IgnoreAdministrator`
allows deliberate elevation for that same user, never another user's HKCU/profile.

```powershell
# Profile discovery; omit ExcludeOst only when you also want raw cache preservation.
.\scripts\Office\Checkpoint-Outlook.ps1 -Destination 'E:\Checkpoints' -ExcludeOst -DryRun
$checkpoint = .\scripts\Office\Checkpoint-Outlook.ps1 -Destination 'E:\Checkpoints' `
  -ExcludeOst -QuitOutlook -PassThru
$checkpoint | Format-List Status, Copied, Failed, CheckpointDirectory, ReportPath, Warnings
```

Select explicit files without connecting to Outlook; settings still belong to the current Windows account:

```powershell
$capture = @{
  Destination = '\\server\backups'
  StorePaths  = @('D:\Mail\mail.pst', 'D:\Archive\archive.pst')
}
.\scripts\Office\Checkpoint-Outlook.ps1 @capture -DryRun
$checkpoint = .\scripts\Office\Checkpoint-Outlook.ps1 @capture -PassThru

$manifest = Get-Content -LiteralPath $checkpoint.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
$manifest | ConvertTo-Json -Depth 30
```

| Captured or checked      | Detail                                                                                                                                                                                                                                                |
| ------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Data-file discovery      | Local attached PST/OST paths in the current classic Outlook profile; it may start Outlook. No guessed directories. Detached files need StorePaths, which conflicts with QuitOutlook/WaitSeconds.                                                      |
| OST                      | Raw preservation only, not a portable mailbox backup or a guarantee of reattachment/restoration elsewhere. ExcludeOst omits caches. Server-only stores and credentials are not captured.                                                              |
| User files               | Signatures, templates, dictionaries, roaming Outlook data, AutoComplete cache, Ribbon/Quick Access Toolbar customizations. Settings scans do not silently add PST/OST files.                                                                          |
| Registry                 | Existing HKCU Office 12.0/14.0/15.0/16.0 keys and legacy Outlook Profiles key.                                                                                                                                                                        |
| Selection failures       | Missing optional settings are skipped; unreadable locations and selected missing files fail. Directory links are not traversed. Destination cannot be inside a captured settings tree. Local/UNC destinations are supported.                          |
| Shutdown and consistency | Close other Office applications. QuitOutlook requests graceful shutdown; otherwise close manually during the wait. Settings are enumerated after Outlook exits to include shutdown caches. No process is force-killed.                                |
| File locking             | All sources are locked exclusively before copying and through verification. Close/check/copy is not an atomic system snapshot; registry export is not a cross-file/registry transaction.                                                              |
| Verification             | Original path, relative checkpoint path, length, modification time, and matching SHA-256 hashes. Registry exports also receive output hashes. SkipHash checks lengths only, labeled Hashed=false and Verification=LengthOnly.                         |
| Failure                  | Unique run directories preserve earlier checkpoints; failed copies remain partial and the manifest records failure when writable. Missing/failed capture is never reported as success.                                                                |
| Manifest/output          | User identity, Office inventory, structured activation status without raw licensing output. Unknown activation metadata does not discard the checkpoint. PassThru returns Status, Copied, Failed, CheckpointDirectory, ReportPath, Warnings, Results. |
| Preview                  | Capture plan only; no files, registry export, or Outlook shutdown. Discovery can still connect to Outlook.                                                                                                                                            |

Restore manually: inspect manifest/hashes, close Office, and review each original path/registry export for the intended user and version.
Old settings can overwrite newer ones; an Office 2007 profile export is not automatically Office 2019-compatible. A checkpoint does not
replace a tested machine backup or preserve server mailbox state, passwords, activation entitlement, or every third-party add-in setting.

## Repair with ScanPST

Back up first and close Outlook. Tool discovery selects ScanPST for PST and OST files; `ToolPath` explicitly overrides it. There is no
automatic ScanOST discovery or `Tool` selector. File, tool, and operation-log paths use long filesystem names.

```powershell
$repair = @{ Path = 'D:\Mail\archive.pst' }
.\scripts\Office\Repair-OutlookDataFile.ps1 @repair -DryRun -PassThru
$result = .\scripts\Office\Repair-OutlookDataFile.ps1 @repair -PassThru
$result | Format-List *
```

| Tool capability                                       | Invocation                                                                                                            |
| ----------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| Office 16 ScanPST version `16.0.10325.20082` or later | Targeted mode; Path required, passed as `-file "<path>" -rescan 1`. One scan pass, visible UI, no force/silent flags. |
| Older/unknown or explicit legacy executable           | Interactive launch without arguments. Path optional; if supplied, displayed/checked but selected again in the UI.     |
| Legacy ScanOST via ToolPath                           | Interactive profile selection; does not receive a file argument. Run as the intended Outlook user.                    |

```powershell
# Explicit legacy utility, only on a workstation where this executable exists.
.\scripts\Office\Repair-OutlookDataFile.ps1 `
  -ToolPath 'C:\Program Files (x86)\Microsoft Office\Office12\SCANOST.EXE'
```

Capability is inferred conservatively from executable version metadata, not folder names or a runtime probe; MSI/Click-to-Run versions can
differ. Targeted mode rejects UNC, mapped-network, and unknown drive types even during preview; copy the data file to local storage.

The script checks running Outlook and exclusive file access when Path is supplied. Spaces in paths are supported; nonzero tool exits fail. A
successful launch/exit does not establish that repair occurred or prove data-file health: inspect ScanPST's results/log and respond to its
UI. Preview/results identify LaunchMode, FileVersion, ToolPath, RequestedPath; Action=LaunchRepairTool describes process execution, and an
interactive tool can work on another file selected in its UI.
