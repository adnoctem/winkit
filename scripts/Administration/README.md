# Administration scripts

| Script                                                   | Purpose                                                                                                                   | Requirements                                                                                           |
| -------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| [Import-UserProfileData.ps1](Import-UserProfileData.ps1) | Import staged user data into an existing signed-out profile, verify contents, and apply the target account's permissions. | Elevated x64 Windows PowerShell 5.1 or PowerShell 7; local NTFS paths and the pinned Substrate module. |
| [New-HostnameScheme.ps1](New-HostnameScheme.ps1)         | Generate names following the OFC hostname schema.                                                                         | Company, building, device and workload naming inputs.                                                  |
| [Update-AutoDNSZones.ps1](Update-AutoDNSZones.ps1)       | Apply domain and DNS zone settings to AutoDNS from a JSON configuration.                                                  | AutoDNS credentials and a desired-state JSON file.                                                     |

See [Administration](../../docs/www/Administration/README.md) for installed paths and
[user data migration](../../docs/www/Administration/User-Data-Migration.md) for ABB preparation, previews, folder mappings, permissions, and
verification.

From the installed winkit directory, with an extracted backup at `C:\Migration\User1`:

```powershell
$import = @{
  SourcePath = 'C:\Migration\User1'
  TargetUser = 'NEWDOMAIN\User1'
}
.\scripts\Administration\Import-UserProfileData.ps1 @import -DryRun
Get-Help .\scripts\Administration\Import-UserProfileData.ps1 -Full
```

The importer changes only selected data folders. It resets custom permissions within those destination trees, including destination-only
files, and sets their owner to the target account. It does not import Windows profile hives, AppData, PSTs, OSTs, or installed applications.
ABB downloads and browser/Outlook migration are separate steps. Staging is retained unless `-RemoveSource` is requested.
