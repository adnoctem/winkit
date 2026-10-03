# winkit Outlook integration testing

The `tests/Office` directory holds the winkit testing solution for the `scripts/Office` scripts. Two tiers exist:

| Tier            | Command                      | Needs Outlook               | Coverage                                                                                            |
| --------------- | ---------------------------- | --------------------------- | --------------------------------------------------------------------------------------------------- |
| 0 - logic       | `.\winkit.ps1 test`          | no                          | Script conventions and mocked behavior, including Office migration and Outlook safety checks.       |
| 1 - integration | `.\winkit.ps1 test -Outlook` | yes (classic 2007 or later) | Fixture generation, deduplication, archiving, and a repair preview against disposable scratch PSTs. |

## Tier 0

```powershell
.\winkit.ps1 init   # installs the pinned modules, including Pester 5.5
.\winkit.ps1 test
```

Integration files are excluded automatically; these tests do not require Outlook. The transport Message-ID parser (`Get-TransportMessageId`)
is tested separately in the PSFoundation repository.

The Office suite requires PSFoundation 1.8.7, including ODT identity validation, archive append, repair metadata and PST lifetime helpers.
Its module tests cover standard identities, the Outlook 2007 MAPI fallback, search folders, exact path selection, and ancestor exclusions.
winkit tests cover implicit German/renamed Inbox selection, explicit literal names and store roots, failure before mutation for missing
identities or excluded ancestors, grouped inclusion switches, archive attachment, and JSON reports. Append checks cover existing and missing
files, attachment ownership, source-store rejection, destination folder conflicts, successive passes, path preservation and flattening,
previews, and partial failures. Repair checks cover targeted and interactive launches, long paths, local-storage validation, locked files,
and process results. Subject reporting checks cover null, empty, whitespace-only, and nonblank subjects in JSON and CSV without changing
source messages. Report-order checks cover `NewToOld`, `OldToNew`, and default received-date ordering, stable ties, undated failures,
CSV/log/output consistency, and unchanged duplicate selection. `OutlookBackup.Tests.ps1` covers profile discovery with mocked COM objects
and real temporary-file copies, locks, hashes, collision handling, previews, and conflicting parameter sets. These checks never open
Outlook.

`OfficeDeployment.Tests.ps1` covers the Install, Remove, and Switch wrappers: mode validation, ordered/default/automatic locales, module
dispatch, preview and confirmation forwarding, recovery scope, preparation results, and exit codes. Deployment commands are mocked; these
tests never invoke ODT. PSFoundation owns tests for native deployment behavior and inventory/media verification. The wrapper tests require
its Office API from version 1.8.1 or later; archiving, splitting and duplicate review require 1.8.7 for the PST lifetime helpers.

OutlookCheckpoint.Tests.ps1 covers explicit and discovered files, user identity, previews, locked/missing sources, OST selection, settings,
registry-export failures, manifest hashes, and preserved prior checkpoints using synthetic files and mocked Office/native APIs.

## Tier 1 - integration suite

### Requirements on the test machine

- Outlook 2007 or later. **Outlook 2007 is 32-bit only**: run the suite from 32-bit PowerShell (`Windows PowerShell 5.1 (x86)` or 32-bit
  PowerShell 7). Office 2010 64-bit and newer work from 64-bit PowerShell.
- The pinned modules must be available in that PowerShell host and user session. PSFoundation requires PowerShell 5.1 or later.
- [Redemption](https://www.dimastr.com/redemption/) is recommended for transport-header and backdated `ReceivedTime` injection; check
  licensing for your use. Native writes are best-effort, and fixture persistence must be verified.
- An interactive MAPI profile. The suite attaches its scratch PST to the active profile and targets that store. It does not create or select
  a separate profile; use a disposable profile and run as its user, with both Outlook and the test shell non-elevated.

### Disposable test profile (recommended)

For an extra safety margin, run the suite against a dedicated Outlook profile whose default store is a scratch PST, not the personal
mailbox:

1. Close Outlook. Open Control Panel -> Mail (32-bit for Outlook 2007) -> Show Profiles -> Add.
2. Configure a test profile with a scratch PST as its default delivery store. If configuring an account, use a disposable test mailbox.
3. Select that profile with `Always use this profile`, or enable profile selection and choose it when opening Outlook manually. Confirm the
   intended profile is open before starting the suite. Restore your normal profile selection after testing.

The suite creates `%TEMP%\winkit-outlook-test-<run-id>` with its own PSTs and a unique store display name. It retains this directory for
inspection. For offline rehearsal, backups, and manual archive checks, see the
[mail archival guide](../../docs/www/Office/Mail-Archival.md#rehearse-and-reduce-a-production-pst).

### Running

```powershell
# from the winkit repository root, in 32-bit PowerShell when testing Outlook 2007
.\winkit.ps1 test -Outlook
```

### What the suite does

1. Attaches a scratch PST store (`winkit-test-store.pst`).
2. `New-TestOutlookMessage.ps1` generates 20 deterministic items (seed 42, 25% duplicate Message-IDs) into `WinkitTestData`.
3. `Optimize-Outlook.ps1` walks the store: asserts the 20 items are seen, a dry run flags exactly 5 duplicates, and the real run moves them
   to the review folder. Deduplication checks require PSFoundation 1.8.7 or later. The preview assertion skips when transport headers are
   unavailable; the move assertion still requires five actual moves, so missing headers can also cause a failure.
4. A second generation (seed 7) asserts deterministic Message-ID sequences.
5. `New-OutlookArchive.ps1` previews (no PST created), then copies the store into an archive PST. Its JSON reports are retained in the
   scratch directory; assertions read the per-message `Results` array through the returned summary's `ReportPath`. When fixture injection
   succeeds, it also verifies `StartDate`/`EndBefore` bounds select the seeded date window. A final `-Mode Move` run empties the source
   store into another PST. Copy and Move checks reopen their archives and compare actual mail counts with the source counts.
6. `Repair-OutlookDataFile.ps1` previews a ScanPST run against the archive (dry run only - the real tool opens its own UI).
7. Attempts to detach the scratch store, releases COM references, and reports the retained artifact directory. Inspect the PSTs there;
   remove them manually only after closing Outlook and confirming the files are no longer in use.

Archive and optimizer calls explicitly use `-FolderName '' -Recurse` because their synthetic folders are under the scratch store root.
Normal script invocations default to Inbox alone. The suite does not automatically exercise profile shutdown or backup of real PST files.

### Interpreting results

- Passed assertions validate the tested environment and fixtures. They do not certify existing PST health, available capacity, or every
  Outlook configuration.
- Header/date-dependent tests skipped: required fixture properties could not be verified. Review generator warnings and `HeaderInjected`;
  complete missing checks with properly licensed Redemption or real dated mail in a disposable PST. A skip is not a pass.
- Repair preview skipped: no ScanPST executable was discovered. The suite does not perform an actual interactive repair.
- Failures: stop before production use. The failing `It` names the script and assertion; inspect retained PSTs and run the failing script
  manually against disposable data with `-DryRun -PassThru`. Complete the user guide's manual archive checks before production Move runs.

### Testing an old Outlook version

The floor is Outlook 2007. To validate:

1. Prepare an isolated Windows VM compatible with both Office 2007 and the required PowerShell host.
2. Install PowerShell 5.1 (x86), then run `.\winkit.ps1 init` in that host to install the pinned modules.
3. Create the disposable profile and run `.\winkit.ps1 test -Outlook` from 32-bit PowerShell.

### Extending

Add new `It` blocks following the existing patterns. Keep the order constraints in mind: the dedup assertions count store-wide results and
must run while only `WinkitTestData` exists. New scenarios should generate their fixtures into their own folder with their own seed.
