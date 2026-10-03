# Office

Deploy Microsoft Office, archive Outlook mail, and preserve data and settings with winkit.

| Task                                                                       | Guide                                                                 |
| -------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| Prepare media, install Office, remove products, or read deployment results | [Office installation and removal](Office-Installation-and-Removal.md) |
| Replace an Office edition or change between 32-bit and 64-bit Office       | [Office migrations](Office-Migrations.md)                             |
| Continue an interrupted installation or migration                          | [Office recovery](Office-Recovery.md)                                 |
| Archive mail, split existing PSTs, choose folders, or review duplicates    | [Mail archival](Mail-Archival.md)                                     |
| Copy PSTs, capture data and settings, or launch ScanPST                    | [Outlook backup and repair](Outlook-Backup-and-Repair.md)             |
| Choose the correct PowerShell session and Outlook profile                  | [Requirements](Requirements.md)                                       |

## Open the installed toolkit

The [winkit installer](../../../dist/README.md) installs the scripts and their pinned PSFoundation dependency.

Choose the directory matching your installation, then run the examples from there:

```powershell
# Default CurrentUser installation.
Set-Location -LiteralPath "$env:LOCALAPPDATA\Programs\winkit"
```

```powershell
# Default AllUsers installation; use native 64-bit PowerShell on 64-bit Windows.
Set-Location -LiteralPath "$env:ProgramFiles\winkit"
```

```powershell
# Example custom location chosen with WINKIT_INSTALL_PATH.
Set-Location -LiteralPath 'C:\Managed\Tools\winkit'
```

Use elevated PowerShell for Office deployment. Use a normal PowerShell session as the mailbox owner for Outlook operations. Elevating with
another account changes the user profile and its CurrentUser installation path; see [session requirements](Requirements.md).

Keep archives, reports, media, and checkpoints **outside the managed winkit directory**. The examples use `C:\Managed\Media`,
`C:\Managed\ODT`, user-local report paths, and an example backup drive `E:`; change these to suit your machine.

## Preview and help

```powershell
# Inventory only; run elevated.
.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check -PassThru

# Parameter help is available in an installed copy.
Get-Help .\scripts\Office\Switch-OfficeVersion.ps1 -Full
Get-Help .\scripts\Office\New-OutlookArchive.ps1 -Full
```

Use `-DryRun` or `-WhatIf` before changes and `-PassThru` for structured results. Outlook previews can connect to Outlook and write reports;
deployment previews do not launch installers or write deployment journals.
