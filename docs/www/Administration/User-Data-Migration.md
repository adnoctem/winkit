# User data migration

`Import-UserProfileData.ps1` copies ordinary data from an extracted backup into an existing Windows profile. It checks the target identity,
profile state, paths, disk space, and content conflicts; copies with robocopy; verifies SHA-256 hashes and file sizes; then resets
permissions and ownership within the selected destination folders. Reports record the operator, source, target SID, folder plan, hashes,
exclusions, and results.

This is a data-folder import. It does not recreate a Windows account, replace the profile registry, rewrite the machine SID, install
applications, or migrate browser and Outlook profiles.

## Prepare the backup and target

1. Finish the old user's work and take a final successful ABB backup. Record the device, task, and backup timestamp you select.
2. Export browser passwords/bookmarks and complete the separate mail migration preparation while the old user session is available.
3. In **Active Backup for Business Portal**, select the correct device and backup version, then download the required user data folders.
   Synology documents this in its
   [recovery portal guide](https://kb.synology.com/en-br/DSM/help/ActiveBackup/activebackupforbusinessportal?version=7). Multiple selections
   may arrive in one archive. Download and extraction are manual; this script does not use an undocumented ABB API.
4. Extract to a local NTFS directory, for example `C:\Migration\User1`, so that `Documents`, `Desktop`, and other chosen folders are
   directly beneath it. If the archive contains `C\Users\User1`, point `SourcePath` at that extracted inner directory instead.
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

Use `-Folder @()` for mappings only. Source and target mappings must be non-overlapping; do not select both `Documents` and
`Documents\Business Data`. Relative paths must stay inside their respective roots. Absolute paths, parent traversal, AppData, and system
directories are refused.

The script does not discover Known Folder redirection or read the old `NTUSER.DAT`. For a fully local redirected folder beneath the new
profile, supply its actual relative path in `FolderMap`. It refuses reparse-point ancestors and cloud placeholders, so an active or
unhydrated OneDrive tree is not an import destination. External drive and UNC destinations require a separate migration workflow.

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
