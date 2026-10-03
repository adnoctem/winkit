# Office

Scripts for Office deployment and Outlook mail/data maintenance. Start with the [Office user overview](../../docs/www/Office/README.md) for
installed-toolkit usage, session requirements, and topic guides.

## Office deployment

| Script                                               | Purpose                                                  | Documentation                                                                                                 |
| ---------------------------------------------------- | -------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| [Install-Office.ps1](Install-Office.ps1)             | Prepare media, install, verify, or recover a deployment. | [Installation and removal](../../docs/www/Office/Office-Installation-and-Removal.md)                          |
| [Switch-OfficeVersion.ps1](Switch-OfficeVersion.ps1) | Inventory and migrate Office products or architectures.  | [Migrations](../../docs/www/Office/Office-Migrations.md)                                                      |
| [Remove-Office.ps1](Remove-Office.ps1)               | Inventory or remove selected Click-to-Run products.      | [Installation and removal](../../docs/www/Office/Office-Installation-and-Removal.md#remove-selected-products) |

Recorded installations and migrations use the [Office recovery workflow](../../docs/www/Office/Office-Recovery.md).

## Outlook

| Script                                                   | Purpose                                                            | Documentation                                                                                     |
| -------------------------------------------------------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------- |
| [New-OutlookArchive.ps1](New-OutlookArchive.ps1)         | Copy or move mail to new or existing PST archives.                 | [Mail archival](../../docs/www/Office/Mail-Archival.md)                                           |
| [Split-OutlookArchive.ps1](Split-OutlookArchive.ps1)     | Transfer one date range from an existing PST into another archive. | [Split an existing archive](../../docs/www/Office/Mail-Archival.md#split-an-existing-archive)     |
| [Optimize-Outlook.ps1](Optimize-Outlook.ps1)             | Move duplicate messages into a review folder.                      | [Duplicate review](../../docs/www/Office/Mail-Archival.md#review-duplicates)                      |
| [Backup-Outlook.ps1](Backup-Outlook.ps1)                 | Create verified closed-file PST copies.                            | [Backup and repair](../../docs/www/Office/Outlook-Backup-and-Repair.md#back-up-pst-files)         |
| [Checkpoint-Outlook.ps1](Checkpoint-Outlook.ps1)         | Capture data files, Office settings, and a manifest.               | [Backup and repair](../../docs/www/Office/Outlook-Backup-and-Repair.md#capture-data-and-settings) |
| [Repair-OutlookDataFile.ps1](Repair-OutlookDataFile.ps1) | Launch Microsoft's scan/repair utility.                            | [Backup and repair](../../docs/www/Office/Outlook-Backup-and-Repair.md#repair-with-scanpst)       |
| [New-TestOutlookMessage.ps1](New-TestOutlookMessage.ps1) | Generate synthetic fixtures for a disposable Outlook store.        | [Contributor fixture guide](../../docs/CONTRIBUTING.md#generate-outlook-test-messages)            |

## Source and parameter help

Each script exposes complete parameter help:

```powershell
Get-Help .\scripts\Office\Switch-OfficeVersion.ps1 -Full
Get-Help .\scripts\Office\New-OutlookArchive.ps1 -Full
```

Maintainers should read [Office and Outlook validation](../../docs/CONTRIBUTING.md#office-and-outlook-validation) and the
[integration test setup](../../tests/Office/README.md).
