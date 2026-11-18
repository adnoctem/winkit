# User data migration

`Import-UserProfileData.ps1` copies ordinary data from an extracted backup into an existing Windows profile. It checks the target identity,
profile state, paths, disk space, and content conflicts; copies with robocopy; verifies SHA-256 hashes and file sizes; then resets
permissions and ownership within the selected destination folders. Reports record the operator, source, target SID, folder plan, hashes,
exclusions, and results.

This is a data-folder import. It does not recreate a Windows account, replace the profile registry, rewrite the machine SID, install
applications, or migrate browser and Outlook profiles.

## Prepare the backup and target

Run the read-only inventory on the **old** machine before downloading anything. Run it elevated to include the scan of fixed drives outside
the profile:

```powershell
$inventory = @{
  User          = 'OLDDOMAIN\User1'
  OutlookReport = 'C:\Reports'
  StagingRoot   = 'C:\Managed\Migration'
}
.\scripts\Administration\Show-UserDataInventory.ps1 @inventory -DryRun
```

It records known-folder redirection, public-folder content and shortcuts, network drives, data outside the profile, and destination path
lengths. Without `-DryRun`, it writes `manifest.json` and `folder-map.json` to a new folder below `%ProgramData%\winkit\UserDataInventory`.
The plan in the manifest contains:

- `AbbDownload`: each folder to download and the staging path to extract it to, for example `C:\Users\User1` to `C:\Managed\Migration\User1`
  and a redirected `C:\Downloads` to `C:\Managed\Migration\Downloads`.
- `Mappings`: the folder map for `Import-UserProfileData.ps1 -FolderMapPath`, also written to `folder-map.json`.
- `ManualCopy`: PSTs, Outlook signatures, and public desktop files that are not program shortcuts, which belong on the new computer's public
  desktop.
- `Notes`: unusual locations the map accounts for, such as a Desktop stored inside `Music`.
- `Decisions`: network drives other than the home drive, the logon script, unmapped data outside the profile, long paths, and Outlook report
  warnings.
- `SuggestedImportCall`: a preview command for the import.

Folder figures count each file once: a known folder stored inside another, such as a Desktop in `Music\Desktop`, is counted in its own
entry, and the outer entry names it under `NestedFolders`. Junctions and symbolic links are not followed; `ReparseItems` lists each one with
its target, because the import skips them too. Dates are ISO 8601 text.

Known folders are mapped from where they are stored into their standard location, so a Desktop redirected to `Music\Desktop` becomes
`Desktop` and is left out of the `Music` import. Empty known folders are noted instead of mapped, because a download may leave them out and
the import refuses a missing source; folders with only a few files are mapped and noted. `-Folder` limits the map to the named known
folders, with the same names as `Import-UserProfileData.ps1 -Folder`.

Public folders are shared by every user of the old computer and may hold application data, so they are listed as decisions with their size
and newest change. `-IncludePublic` imports the named ones below the matching user folder, for example `Documents\Public Documents`. The
public desktop is never imported: shortcuts to local programs are noted as recreated by the new computer's installers, shortcuts to network
paths are decisions, and any other files are listed under `ManualCopy`.

Other folders outside the profile are listed as decisions, with the programs that probably own them. A program counts when its install
location, install source, icon, or uninstaller, or a Start-menu or desktop shortcut, points into the folder; `ProgramMatches` and the
decision text say which. System, vendor, and winkit-managed folders such as `$WinREAgent`, `inetpub`, and `C:\Managed` are not scanned. Name
a folder with `-Include` to migrate it into `Documents\<folder name>`; this also works for a folder that is skipped by default:

```powershell
.\scripts\Administration\Show-UserDataInventory.ps1 @inventory -Folder Desktop, Documents, Downloads -IncludePublic Documents
.\scripts\Administration\Show-UserDataInventory.ps1 @inventory -Include 'C:\Business', 'D:\Scans'
```

The user's **home drive** is read from their signed-in session and, for domain accounts, from Active Directory (`homeDrive`,
`homeDirectory`), so it is found even when an administrator runs the inventory. It is mapped to `Private` in the new profile and staged as
`HomeDrive`; choose another target with `-HomeDriveTarget`. ABB does not back up the file server, so copy the share into the staging folder
yourself before the import, for example:

```powershell
robocopy '\\server\home\User1' 'C:\Managed\Migration\HomeDrive' /E /COPY:DAT /DCOPY:DAT /XJ /R:1 /W:1
```

Each drive in `NetworkDrives` records its `Source`: `HomeDrive`, `Persistent`, `Connected`, or `LogonScript`. Drives connected in the
current session are read only when the user runs the inventory without elevation; other sessions cannot see them, and a note says so. A
logon script from Active Directory is listed as a decision because it may connect further drives. `-MeasureNetwork` measures the home drive
and other drives through their UNC paths.

Pass the new profile folder as `-TargetRoot`. Every mapped folder, the home drive included, is checked at its mapped target, so a Desktop
moved from `Music\Desktop` gets shorter and `Documents\Public Documents` gets longer. Drives measured with `-MeasureNetwork` are checked as
if imported into `Documents\<share name>`. Paths longer than 247 (folders) or 259 (files) characters are listed as decisions, the same
limits the import enforces. A folder that cannot be scanned is reported as a decision and does not stop the inventory.

```powershell
.\scripts\Administration\Show-UserDataInventory.ps1 @inventory -TargetRoot 'C:\Users\User2'
```

When `-TargetRoot` is below the same profiles folder as the old profile, its folder name and the old user's domain fill in `-TargetUser` in
the suggested import call, with a note to check it. Without `-TargetRoot`, the old profile folder stands in, the inventory warns, and the
import call keeps a `<target>` placeholder.

1. Finish the old user's work and take a final successful ABB backup. Record the device, task, and backup timestamp you select.
2. Export browser passwords/bookmarks and complete the separate mail migration preparation while the old user session is available.
3. In **Active Backup for Business Portal**, select the correct device and backup version, then download the required user data folders.
   Synology documents this in its
   [recovery portal guide](https://kb.synology.com/en-br/DSM/help/ActiveBackup/activebackupforbusinessportal?version=7). Multiple selections
   may arrive in one archive. Download and extraction are manual; this script does not use an undocumented ABB API.
4. Extract to a local NTFS directory. With a folder map, extract every download to its `AbbDownload` staging path below one staging root,
   for example `C:\Managed\Migration\User1` and `C:\Managed\Migration\Public`. For a single folder without a map, use a directory such as
   `C:\Migration\User1` with `Documents`, `Desktop`, and other chosen folders directly beneath it. If the archive contains `C\Users\User1`,
   move or point to that extracted inner directory instead.
5. Create the new local/domain account, sign in once to initialize its profile, and sign out. Run the importer as a separate elevated
   administrator. Keep the user signed out and stop applications or sync services that could modify staging or the destination.

A full disk download is unnecessary for this workflow. You can retain one as recovery material, but only selected profile data folders are
imported. Inspect old folder redirection, OneDrive, other data drives, shared folders, and custom business-data directories separately.
Their contents may live outside the old profile and must be staged or mapped deliberately.

Use the [installed toolkit directory](README.md#open-the-installed-toolkit) for all commands below. ABB export and local extraction must
finish before running the script; a partly downloaded or changing source is not a valid staging tree.

## Preview and copy

```powershell
$import = @{
  SourcePath = 'C:\Migration\User1'
  TargetUser = 'NEWDOMAIN\User1'
}
.\scripts\Administration\Import-UserProfileData.ps1 @import -DryRun
```

For a local account, use `NEWPC\User1`. A user SID is also accepted, but it must resolve to an account with an existing profile. The script
cross-checks CIM and `ProfileList` and refuses loaded, special, temporary, or corrupted profiles. It rechecks the target during import.
These checks do not prevent a user from logging in, so arrange a maintenance window.

`-DryRun` and `-WhatIf` enumerate and hash source data and check conflicts without writing reports, copying files, changing permissions, or
deleting staging. After reviewing the plan:

```powershell
$results = .\scripts\Administration\Import-UserProfileData.ps1 @import -PassThru
$summary = $results | Where-Object Action -EQ ImportProfileData
$summary | Format-List Status, Error, ManifestPath, StagingPath
```

The default set is `Desktop`, `Documents`, `Downloads`, `Pictures`, `Music`, `Videos`, `Favorites`, `Contacts`, `Links`, `Saved Games`, and
`Searches`. Missing default folders are reported as skipped. Explicitly selected folders must exist:

```powershell
.\scripts\Administration\Import-UserProfileData.ps1 @import -Folder Desktop,Documents,Downloads -DryRun
```

Existing files with different content cause a failure before copying. Review those files first; add `-Overwrite` only when the staged
version should replace them. Destination-only files are kept. Repeat runs accept existing identical content, though robocopy may recopy it
to apply attributes and timestamps consistently. No `/MIR`, `/PURGE`, `/MOVE`, source ACL, or source-owner copying is used.

## Custom folders and redirection

Use on-disk paths, which may differ from Explorer's translated labels. `-Folder` replaces the default set. `-FolderMap` adds explicit
source-relative to target-relative directory mappings:

```powershell
$mappedImport = @{
  SourcePath = 'C:\Migration\User1'
  TargetUser = 'NEWDOMAIN\User1'
  Folder     = @('Desktop', 'Downloads')
  FolderMap  = @{
    Documents       = 'Documents'
    'Business Data' = 'Business Data'
    Pictures        = 'Pictures\Old computer'
  }
}
.\scripts\Administration\Import-UserProfileData.ps1 @mappedImport -DryRun
```

Use `-Folder @()` for mappings only. Relative paths must stay inside their respective roots. Absolute paths, parent traversal, AppData, and
system directories are refused. Overlapping mappings follow three rules:

- A source inside another selected source is copied only by its own mapping. With `Music\Desktop = 'Desktop'` and `Music = 'Music'`, the
  desktop files land in `Desktop` and are left out of the `Music` copy, its verification, and `-RemoveSource`.
- A destination inside another destination is allowed, for example `Public\Documents = 'Documents\Public Documents'` next to `Documents`,
  unless the outer source also contains an item at that path.
- Two mappings with the same source or the same destination are refused.

`-FolderMapPath` imports the `folder-map.json` written by `Show-UserDataInventory.ps1` instead of `-Folder` and `-FolderMap`; see
[Import a folder map](#import-a-folder-map).

The script does not discover Known Folder redirection or read the old `NTUSER.DAT`. For a fully local redirected folder beneath the new
profile, supply its actual relative path in `FolderMap`. It refuses reparse-point ancestors and cloud placeholders, so an active or
unhydrated OneDrive tree is not an import destination. External drive and UNC destinations require a separate migration workflow.

## Refresh staging from the live computer

An ABB download goes stale while the user keeps working. The inventory's plan contains a `FinalPass` block, and the run folder a
`final-pass.ps1`, that mirror each mapped folder from the old computer into the existing staging with
`robocopy /MIR /COPY:DAT /DCOPY:DAT /XJ`. Shortcuts (`.lnk`, `.url`) are copied; only `desktop.ini` and `Thumbs.db` are skipped. A folder
mapped separately, such as a Desktop inside `Music`, is excluded from the outer line with `/XD` and mirrored by its own line. The home drive
is read from its file server path.

On the old computer, as an administrator, create the hidden read-only share listed in `FinalPass.ShareCommands`. It is granted to the
account that ran the inventory when that is a domain account; replace the placeholder otherwise. The account also needs NTFS read access to
the mapped folders:

```powershell
net share MIGRO$=C:\ /GRANT:"OLDDOMAIN\Operator",READ
```

On the new computer, copy `final-pass.ps1` from the inventory run folder and preview it; without `-Apply` every line runs with `/L` and
copies nothing. Logs are written beside the script:

```powershell
.\final-pass.ps1
.\final-pass.ps1 -Apply
.\final-pass.ps1 -SourceComputer OLDPC -StagingRoot 'D:\Migration' -Apply
```

robocopy only reads the old computer. `/MIR` changes only the staging destination: it copies new and changed files and deletes staging files
that no longer exist on the old computer. Keep the user signed out and Outlook closed, then remove the share on the old computer:

```powershell
net share MIGRO$ /DELETE
```

## Import a folder map

`Show-UserDataInventory.ps1` writes `folder-map.json`. Each mapping has a `Source` relative to `-SourcePath`, a `Target` relative to the new
profile, and, for reference, the original `Origin` path and its `Kind`. Point `-SourcePath` at the staging root that holds every extracted
download, not at one user folder:

```powershell
$mapImport = @{
  SourcePath    = 'C:\Managed\Migration'
  TargetUser    = 'NEWDOMAIN\User1'
  FolderMapPath = 'C:\ProgramData\winkit\UserDataInventory\<run>\folder-map.json'
}
.\scripts\Administration\Import-UserProfileData.ps1 @mapImport -DryRun
```

Every mapped source must exist; download and extract all of them first, or remove an entry from a copy of the file. To change where data
lands, edit `Target` in a copy of the file. The map's notes are printed before the plan and recorded in the import manifest with the map's
path. The public desktop is not part of the map; copy its `ManualCopy` files into `%PUBLIC%\Desktop` on the new computer.

## Permissions and verification

The copy uses `/E /COPY:DAT /DCOPY:DAT /XJ /R:1 /W:1` plus file-selection and logging options. Source junctions and symbolic links inside
selected trees are excluded explicitly; destination links stop the run before copying or recursive ACL work. Source/destination roots and
their existing ancestors must not be links. Encrypted files and offline/cloud placeholders must be decrypted or hydrated first.

`/COPY:DAT` omits the old account's ACLs and owner. After copying, the script runs `icacls /reset` and `/setowner` for each selected
destination tree, then verifies an effective-on-the-object direct SID grant containing Modify rights and the exact target owner. It does not
calculate group-based effective access; deny entries are rejected conservatively. The profile root and any mapped parent must already
provide an inheritable Modify grant for the target SID.

**Custom permissions within selected destination trees are replaced**, including permissions on destination-only files. The profile root,
AppData, and unselected destination folders are not reset. Choose mappings narrowly when an existing folder contains specially secured data.
Changing ownership alone would not establish the required access; the inherited ACLs provide it. Microsoft documents these distinctions in
[robocopy](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy) and
[icacls](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/icacls).

Verification compares each included source file with its specific destination using length and SHA-256 of the main data stream. It also
checks expected directories, detects source file-list changes, and verifies owner and Modify permissions throughout each selected target
tree. Destination-only files do not invalidate the content check. It does not equate aggregate file counts with a successful migration.

Disk-space preflight conservatively requires the total included source size plus 256 MB free on the destination volume, even when some files
already exist. `-ReserveBytes` adjusts the reserve. Report storage and files added by other processes can still consume space during the
operation; a later error produces failure and retains staging unless cleanup has already begun.

## Reports and failure recovery

Each run creates a unique directory under `%ProgramData%\winkit\ProfileImports`, or a supplied `-ReportDirectory` outside source and target.
The run directory is restricted to the operator, Administrators, and SYSTEM. It contains:

- `manifest.json`: source, identity, plan, file hashes, exclusions, native exit codes, and overall state.
- `operations.jsonl`: structured operation results.
- `robocopy-*.log` and `permissions-*.log`: native copy and permission output.

Robocopy codes 0–7 still require the independent verification checks; codes 8 and above fail the run. The script uses Substrate's
`Invoke-SafeProcess`, `Convert-RobocopyExitCode`, and `Write-OperationResultLog` helpers. Exit code 1 indicates a failure. No reports are
written if preflight fails or the run is a preview; read the returned failure result in that case.

The import is not transactional. On failure, inspect the report and logs, correct the problem, and rerun against retained staging. Already
copied data or partially changed ACLs may remain. Do not delete staging because robocopy alone reported success. After a forced process
termination, a manifest that still says `Started` or `Copying` is not evidence of completion.

## Optional staging cleanup

Staging is retained by default. To request cleanup as part of a verified import:

```powershell
.\scripts\Administration\Import-UserProfileData.ps1 @import -RemoveSource -PassThru
```

Cleanup starts only after every selected tree passes verification and its verified audit is written. Source and destination hashes are
checked again before deletion. Only included verified source files and empty selected directories are removed, using literal paths and
non-recursive deletion. Excluded files, source links, unselected folders, and `SourcePath` itself remain. A cleanup failure can leave
staging partly cleaned; the verified destination remains the recovery copy. Keep the independent ABB backup.

PSTs, OSTs, `NTUSER.*`, `UsrClass.*`, and nested AppData directories are excluded. Inspect the manifest's excluded entries and finish the
separate browser and [mail migration](../Office/Mail-Archival.md) work before retiring the old machine. This script cannot validate that
those exports are current or complete.
