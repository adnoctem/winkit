# Office and Outlook requirements

[Office overview](README.md)

The winkit installer supplies the pinned PSFoundation dependency. Scripts check their required versions themselves; Windows PowerShell 5.1
or a supported PowerShell 7 host is needed.

Archiving, splitting PST archives, and duplicate review require PSFoundation 1.8.7 or later, including its PST attachment lifetime helpers.

## Choose the session

| Operation                                   | PowerShell session                                                    | Office requirement                                                  |
| ------------------------------------------- | --------------------------------------------------------------------- | ------------------------------------------------------------------- |
| Install, remove, migrate, or recover Office | Elevated; 64-bit on 64-bit Windows, even when replacing 32-bit Office | Supported x64 desktop host and verified ODT/media                   |
| Archive or deduplicate mail                 | Non-elevated, as the mailbox's Windows user                           | Classic Outlook 2007+ for PSTs; 2010+ for IMAP/Exchange archiving   |
| Discover PSTs for backup                    | Non-elevated, as the Outlook user                                     | Classic Outlook 2007+                                               |
| Copy explicit PST paths                     | Account with file access; no profile required                         | Outlook need not be installed; files must be closed                 |
| Capture a checkpoint                        | Affected Windows user; normally non-elevated                          | Classic Outlook 2007+ for discovery; explicit StorePaths avoids COM |
| Repair a data file                          | Account with file access; close Outlook                               | ScanPST discovery from Office 12 onward, or an explicit tool        |

Ordinary Outlook COM automation supports 64-bit PowerShell with 32-bit Outlook because Outlook runs in a separate process. See
[Microsoft's cross-architecture COM guidance](https://learn.microsoft.com/en-us/windows/win32/winprog64/process-interoperability). Optional
Redemption/MAPI use requires matching PowerShell, Redemption, and Outlook architecture.

Office deployment supports x64 Windows 10 22H2 (build 19045) and later desktop hosts, including Windows 11. Windows Server, ARM, and older
Windows 10 builds are blocked by the deployment backend. Product recognition does not establish vendor lifecycle or OS support. Office
desktop/profile operations are not supported on Server Core.

New Outlook is not supported by these COM workflows; PST support alone does not establish automation compatibility. Profile operations need
an interactive classic Outlook session and are unsuitable for SYSTEM. Checkpoint also rejects service accounts. See
[Microsoft's Outlook automation guidance](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/vba-alternatives).

## Use the correct Outlook identity

Open Outlook normally, then open PowerShell as that same Windows user. The scripts reuse the running/default Outlook session; they do not
select a profile. An administrator account with a non-elevated token is allowed. `-IgnoreAdministrator` only permits intentional elevation
under that same account; it cannot select another user's mailbox or settings. The elevation guard applies to previews too.

```powershell
# From your installed winkit directory; read-only archive preview in the default Inbox.
$archiveDirectory = Join-Path $env:LOCALAPPDATA 'winkit\archives'
New-Item -ItemType Directory -Path $archiveDirectory -Force | Out-Null

.\scripts\Office\New-OutlookArchive.ps1 `
  -ArchivePath (Join-Path $archiveDirectory 'preview.pst') `
  -DryRun -Verbose
```

Omitting `-StoreName` selects the default delivery store. Otherwise use a unique display name; duplicates are rejected. `-Verbose` helps
inspect detected stores. Leave `-QuitOutlook` off when Outlook should remain open; previews never quit it.

Profile previews can open Outlook, encounter profile/credential/security dialogs, and write reports or logs. They avoid the requested mail
changes. The scripts do not send mail or access address books/Recipients, avoiding those common Object Model Guard triggers.

## Dependencies, progress, and help

If a module cannot be found, check visibility in the intended host and account. A module installed only for another user or PowerShell
edition may be unavailable. Use the [installer's update procedure](../../../dist/README.md#installation-and-updates) to update winkit and
its dependency.

```powershell
Get-Module -ListAvailable PSFoundation | Select-Object Name, Version, Path
Get-Help .\scripts\Office\New-OutlookArchive.ps1 -Full

# Optional: hide progress bars while keeping warnings, status messages, and results.
$ProgressPreference = 'SilentlyContinue'
```

Progress identifies the current phase or folder and clears on completion/failure. Item updates are throttled; percentages describe the
current folder/phase. Deployment progress does not estimate download/install completion. ScanPST provides its own scan progress and prompts.
All scripts accept `-DryRun`, `-WhatIf`, and `-PassThru`; see each topic for its preview and result behavior.
