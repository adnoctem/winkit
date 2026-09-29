# Distribution and installation

This directory contains the maintained installer, local uninstaller, and their documentation. Release archives and SHA-256 checksums are
generated in the ignored `build/` directory. Both scripts are standalone and do not require PSFoundation to run.

## Installation and updates

Use native 64-bit Windows PowerShell 5.1 or PowerShell 7 on 64-bit Windows:

```powershell
irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
```

The installer resolves a GitHub release, verifies its ZIP against the published SHA-256 checksum, validates archive paths, installs its
pinned PSFoundation dependency, and activates the release. This verifies archive consistency; the repository and published checksum remain
the source of trust. Repeating the command updates the discovered registered installation. An exact version policy remains pinned until
changed; use `WINKIT_VERSION=latest` to resume following stable releases. Selecting an older exact version allows a downgrade.

On a fresh installation, scope defaults to `CurrentUser` in an unelevated session and `AllUsers` in an elevated session. The installer
announces the elevated default. Explicit `WINKIT_SCOPE=CurrentUser` remains valid when elevated. Existing registrations retain their scope;
elevation alone does not migrate them. Conflicting scope/path selections fail instead of creating a second installation.

Default destinations are `%LOCALAPPDATA%\Programs\winkit` and native `%ProgramFiles%\winkit`. The installer maintains one registration per
scope. When both scopes are registered, select one explicitly. HKCU always means the executing account; elevation with another account's
credentials does not grant discovery of the original user's HKCU installation.

```powershell
$env:WINKIT_INSTALL_PATH = 'C:\Managed\Tools\winkit'
$env:WINKIT_DRY_RUN = '1'
irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex

# After reviewing the plan, install with the same choices:
Remove-Item Env:WINKIT_DRY_RUN
irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
```

To inspect the source before execution, download it with `Invoke-RestMethod -OutFile`, review it, then run that file. The same environment
settings apply to file and pipeline execution. No public command-line parameters are accepted.

### Managed directory layout and relocation

For a common local layout, explicitly select `C:\Managed\Tools\winkit` for winkit, `C:\Managed\ODT` for the Office Deployment Tool, and
`C:\Managed\Media` for verified installation packages. This convention does not change the installer's scope-based default paths or Office's
application installation directory. Keep reports, checkpoints and personal files outside the managed winkit tree.

Repeating installation at the registered location updates it. Setting a different WINKIT_INSTALL_PATH in the same scope fails; there is no
same-scope relocation mode. To move such an installation, first ensure the desired release is available, preserve personal files, then
preview and uninstall the exact registered old path. Clear WINKIT_UNINSTALL, select the new path and scope, and preview/install there.
Shared PSFoundation modules remain installed. Do not manually move the directory or edit registration to bypass the ownership manifest. For
a simultaneous scope change, use the supported scope-change operation below with the new destination instead.

A Git checkout is separate from a managed installation. Preserve its local changes and use Git to establish it in the new location; do not
overlay a checkout on the installer's managed directory. The web installer always selects a published release, not local or unpushed code.

## Environment settings

Set these in the current process with `$env:NAME = 'value'`. Explicit choices override stored installation options, subject to identity and
scope checks. Unset/blank settings reuse registered choices where applicable, otherwise defaults. Boolean values accept `1/true/yes/on` or
`0/false/no/off`, case-insensitively. Other values fail before changes. Unprefixed variables such as `DRY_RUN` are ignored.

| Variable                 | Behavior                                                                                                       |
| ------------------------ | -------------------------------------------------------------------------------------------------------------- |
| `WINKIT_SCOPE`           | `CurrentUser` or `AllUsers`; default selected from registration or elevation.                                  |
| `WINKIT_INSTALL_PATH`    | Absolute local destination. For uninstall, an exact registered target; for scope change, the destination.      |
| `WINKIT_REPOSITORY`      | `OWNER/REPOSITORY`; defaults to `adnoctem/winkit`. Cannot change an existing installation's repository.        |
| `WINKIT_VERSION`         | Exact semantic release version, optionally prefixed with `v`, or `latest` to clear a stored pin.               |
| `WINKIT_NO_PATH`         | Skip adding the installation's `bin` to persistent and current-session PATH; stored, defaults to false.        |
| `WINKIT_FORCE`           | Reinstall the selected release and pinned dependency. Does not bypass ownership or file checks.                |
| `WINKIT_NON_INTERACTIVE` | Skip the explicit confirmation; conflicts and validation failures still fail.                                  |
| `WINKIT_DRY_RUN`         | Preview only; takes precedence over force and non-interactive settings.                                        |
| `WINKIT_PASS_THRU`       | Return a structured result, including status, path, scope, dependency/PATH changes, and pending cleanup paths. |
| `WINKIT_UNINSTALL`       | Remove a selected registered installation. Defaults to false.                                                  |
| `WINKIT_CHANGE_SCOPE`    | Move a registered installation to the explicitly specified opposite scope. Defaults to false.                  |

All invocation-only Boolean settings default to false and are never persisted. Environment variables remain in the current session until
changed or removed. Preview prints destinations, registry/PATH changes, dependency work, and release sources when applicable. Ordinary
installation preview reads GitHub release metadata; uninstall and scope-change preview require no network. None downloads assets or writes
files, registry values, dependencies, or PATH. AllUsers previews can run without elevation.

Removal and scope changes print their target and require typing `YES` unless `WINKIT_NON_INTERACTIVE=1`. Empty input means No. Ordinary
installation and updates retain PowerShell's standard confirmation preference. Non-interactive mode never suppresses validation. The scripts
do not alter the behavior of winkit's operational scripts.

## Uninstallation

Use Windows Installed Apps, the installed local `dist\uninstall.ps1`, or the web entry point:

```powershell
$env:WINKIT_UNINSTALL = '1'
$env:WINKIT_SCOPE = 'CurrentUser'
$env:WINKIT_DRY_RUN = '1'
irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
```

Remove `WINKIT_DRY_RUN` to perform removal. Clear `WINKIT_UNINSTALL` before a future install. A supplied `WINKIT_INSTALL_PATH` must match
exactly after normalization; the installer never falls back to another path. Otherwise registration identifies the target. Multiple matches
require an explicit scope or path. Missing, stale, or inconsistent registrations fail. Unregistered installations are not adopted or
removed.

The local uninstaller runs offline and pins the target to its own adjacent ownership manifest. It ignores inherited installation choices,
while honoring dry-run, non-interactive, and pass-through settings. Interactive AllUsers removal requests elevation if needed; unattended
AllUsers removal requires an already elevated session. Direct invocation from an existing native PowerShell session preserves result output.

Removal deletes the verified owned files, the exact matching PATH entry in the installation's scope and current process, and its
installation and Windows uninstall registration. Shared PSFoundation versions, NuGet, and PowerShellGet remain installed. Uninstall rejects
`WINKIT_FORCE`, `WINKIT_VERSION`, and `WINKIT_NO_PATH`; they do not describe removal. It cannot be combined with scope change.

## Changing scope

```powershell
# Run elevated, under the account owning the CurrentUser installation.
$env:WINKIT_CHANGE_SCOPE = '1'
$env:WINKIT_SCOPE = 'AllUsers'
$env:WINKIT_DRY_RUN = '1'
irm https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1 | iex
```

The destination defaults to that scope's standard directory; `WINKIT_INSTALL_PATH` can select another unused local directory. The source is
the single registration in the opposite scope. An occupied destination scope/path or overlapping source/destination directories fails. Both
directions require elevation. Scope changes copy and verify the existing release locally, retain its version-selection policy and
repository, and ensure its pinned PSFoundation dependency is available in the destination scope. They do not combine a version update or
forced reinstall with migration. Source and destination modules are retained as shared dependencies.

After validating the destination, the installer retires the source and transfers registry/PATH registration. On failure before commit it
attempts to restore source/destination files, registrations, and PATH. Clear `WINKIT_CHANGE_SCOPE` after use.

## Ownership, failure handling, and results

`.winkit-install.json` records schema version 2, installation identity, canonical path, scope, repository, installed version, installation
timestamp, and hashes of managed files. Registry identity and this manifest must agree. Added, missing, or modified files block updates,
scope changes, and uninstall. Preserve personal files outside the installation and restore managed files before retrying. Reparse points in
installation paths or trees are rejected; a matching registry entry alone never authorizes recursive deletion.

Maintenance operations use a machine-wide mutex and revalidate after acquiring it. Files are staged beside the destination. Updates and
removals retain the previous tree under a temporary sibling name until filesystem and metadata changes commit. Registry/PATH failures
trigger rollback; dependency installations are shared and are not rolled back. Rollback errors report retained paths for manual attention.

Normal result statuses are `Planned`, `Skipped`, `Installed`, `Updated`, `Reinstalled`, `Current`, `ScopeChanged`, and `Uninstalled`. Errors
terminate execution. A committed operation whose temporary-file cleanup fails reports `CleanupPending` and its `CleanupPaths`, and emits a
warning even without pass-through. Registration may already be removed after uninstall; inspect the reported retained directory rather than
rerunning removal against an unrelated target. A process crash/power loss is not an automatically recoverable transaction.

## Registry schema

Use HKCU for `CurrentUser` and HKLM for `AllUsers`, in the native registry view. Installer metadata is separate from Windows uninstall
registration. No `Settings` subkey is created until application settings exist.

```text
Software\AdNoctem\winkit
    SchemaVersion       REG_DWORD   2
    InstallId           REG_SZ      Installation GUID
    InstalledVersion    REG_SZ      Installed release version
    InstallOptions
        Scope           REG_SZ      CurrentUser or AllUsers
        InstallPath     REG_SZ      Canonical absolute path
        Repository      REG_SZ      OWNER/REPOSITORY
        Version         REG_SZ      Requested version, or empty for latest stable
        NoPath          REG_DWORD   0 or 1
```

Only persistent installation choices are recorded. Environment variables are not written to user/machine environment settings. Runtime flags
are never restored from registration. The ownership manifest records file identity; `InstallOptions` records reusable choices.

`Software\Microsoft\Windows\CurrentVersion\Uninstall\winkit` in the same hive contains `DisplayName`, `DisplayVersion`, `Publisher`,
`InstallLocation`, `InstallDate`, `UninstallString`, `URLInfoAbout`, `NoModify=1`, `NoRepair=1`, and the matching `InstallId`. The uninstall
command invokes the verified local uninstaller with native Windows PowerShell. It never downloads executable code from a URL. Registered
paths are read as data; registry command strings are never evaluated by the maintenance engine.
