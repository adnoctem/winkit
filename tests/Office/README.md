# winkit Outlook integration testing

The `tests/Office` directory holds the winkit testing solution for the
`scripts/Office` scripts. Two tiers exist:

| Tier            | Command                      | Needs Outlook       | Coverage                                                                                                                                                   |
| --------------- | ---------------------------- | ------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0 - logic       | `.\winkit.ps1 test`          | no                  | Script-independent logic. The transport Message-ID parser lives in PSFoundation (`Get-TransportMessageId`) and is tested in the PSFoundation repository.   |
| 1 - integration | `.\winkit.ps1 test -Outlook` | yes (2007 or later) | Full end-to-end runs of `New-TestOutlookMessage`, `Optimize-Outlook`, `New-OutlookArchive`, and `Repair-OutlookDataFile` against a disposable scratch PST. |

## Tier 0

```powershell
.\winkit.ps1 init   # installs the pinned modules, including Pester 5.5
.\winkit.ps1 test
```

Integration files are excluded automatically; the suite passes on machines
without Outlook.

## Tier 1 - integration suite

### Requirements on the test machine

- Outlook 2007 or later. **Outlook 2007 is 32-bit only**: run the suite from
  32-bit PowerShell (`Windows PowerShell 5.1 (x86)` or 32-bit PowerShell 7).
  Office 2010 64-bit and newer work from 64-bit PowerShell.
- Redemption (recommended, free for personal use):
  https://www.dimastr.com/redemption/ - enables guaranteed transport-header
  and backdated `ReceivedTime` injection. Without it, header-dependent
  assertions are skipped automatically.
- A normal MAPI profile. The suite attaches a scratch PST to that profile for
  the duration of the run and detaches it afterwards; the personal mailbox is
  never touched.

### Disposable test profile (recommended)

For an extra safety margin, run the suite against a dedicated Outlook profile
whose default store is a scratch PST, not the personal mailbox:

1. Windows Settings -> Mail -> Accounts -> Manage profiles -> Add, or
   Control Panel -> Mail (32-bit) on older systems.
2. Create a profile with a single POP3/IMAP account pointing at a scratch
   mailbox, or add a PST-only profile ("No e-mail" -> create data file).
   Outlook 2007: use `outlook.exe /manageprofiles`.
3. Make the new profile the default (`Always use this profile`) or launch the
   suite via `outlook.exe /profiles "TestProfile"` when starting Outlook.

The suite creates `%TEMP%\winkit-outlook-test\` with its own PSTs regardless
of the active profile.

### Running

```powershell
# from the winkit repository root, in 32-bit PowerShell when testing Outlook 2007
.\winkit.ps1 test -Outlook
```

### What the suite does

1. Attaches a scratch PST store (`winkit-test-store.pst`).
2. `New-TestOutlookMessage.ps1` generates 20 deterministic items (seed 42,
   25% duplicate Message-IDs) into `WinkitTestData`.
3. `Optimize-Outlook.ps1` walks the store: asserts the 20 items are seen,
   a dry run flags exactly 5 duplicates, and the real run moves them to the
   review folder. Skipped when PSFoundation 1.3.0 or Redemption is missing.
4. A second generation (seed 7) asserts deterministic Message-ID sequences.
5. `New-OutlookArchive.ps1` previews (no PST created), then copies the store
   into an archive PST; with Redemption present it also verifies
   `StartDate`/`EndDate` bounds select exactly the seeded date window, and a
   final `-Mode Move` run empties the store into a second PST.
6. `Repair-OutlookDataFile.ps1` previews a ScanPST run against the archive
   (dry run only - the real tool opens its own UI).
7. Detaches and deletes the scratch store and all generated PSTs.

### Interpreting results

- All green: the scripts work end-to-end on this Outlook version.
- Header-dependent tests skipped: Redemption is not installed; Message-ID
  injection could not be verified (see warning output of the generator).
- Failures: the failing `It` names the script and assertion; run the failing
  script manually against the disposable profile with `-DryRun -PassThru`.

### Testing an old Outlook version

The floor is Outlook 2007. To validate:

1. Install a Windows VM (XP SP3 / 7) with Office 2007.
2. Install PowerShell 5.1 (x86) and the pinned modules
   (`Save-Module PSFoundation,Pester -Path ...` and `Import-Module`).
3. Create the disposable profile and run `.\winkit.ps1 test -Outlook` from
   32-bit PowerShell.

### Extending

Add new `It` blocks following the existing patterns. Keep the order
constraints in mind: the dedup assertions count store-wide results and must
run while only `WinkitTestData` exists. New scenarios should generate their
fixtures into their own folder with their own seed.
