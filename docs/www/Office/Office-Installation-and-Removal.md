# Office installation and removal

[Office overview](README.md) | [Requirements](Requirements.md)

Run these examples from your **installed winkit directory**, in elevated 64-bit PowerShell on 64-bit Windows. Install handles clean machines
and verified compliant no-ops. Use [migration](Office-Migrations.md) to replace an existing installation and [recovery](Office-Recovery.md)
to continue a recorded interrupted operation.

## Obtain and validate ODT

Download and extract the [Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117). `OdtPath` selects the
extracted `setup.exe`, not the `officedeploymenttool_*.exe` download. See
[Microsoft's ODT overview](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-office-deployment-tool).

Alternatively, acquire a verified copy through PSFoundation:

```powershell
Import-Module PSFoundation -Force
New-Item -ItemType Directory -Path 'C:\Managed' -Force | Out-Null

$source = Test-OfficeDeploymentToolSourceAvailability
if (-not $source.Available) {
  throw 'The reviewed ODT download is unavailable.'
}

Install-OfficeDeploymentTool -Destination 'C:\Managed\ODT' -DryRun
$tool = Install-OfficeDeploymentTool -Destination 'C:\Managed\ODT' -Confirm
if (-not $tool.Valid) {
  throw 'ODT acquisition did not return a verified tool.'
}
```

The availability check tests reachability only. Acquisition verifies Microsoft signatures on the extractor and extracted tool, rejects
incomplete extraction, and reuses an existing trusted tool. Its directory must permit protected ownership and permissions.

Validate the exact file to be used for deployment; this does not launch it or download Office:

```powershell
Import-Module PSFoundation -Force
$tool = Test-OfficeDeploymentTool -OdtPath 'C:\Managed\ODT\setup.exe'
$tool | Format-List Valid, Version, SignatureStatus, OriginalFilename, FileDescription, Detail
if (-not $tool.Valid) {
  throw "ODT validation failed: $($tool.Detail)"
}
```

Microsoft's tool can report `OriginalFilename = Bootstrapper.exe` and description `Microsoft 365 and Office`. Both identity and Microsoft
signature/version must validate; renaming another executable is insufficient. Update winkit/PSFoundation when using an old installation that
rejects this metadata; PSFoundation 1.7.1 and earlier lack the current ODT identity support. Do not bypass validation.

## Choose the target

`SourcePath` is an installation-media directory, **not** Office's application directory. Office uses its normal installation location. These
scripts require an explicit media path with an existing parent, even though raw ODT can download beside `setup.exe` without one. See
[Microsoft's SourcePath documentation](https://learn.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options).

| Product family     | TargetProductId                       | Channel                                             |
| ------------------ | ------------------------------------- | --------------------------------------------------- |
| Office 2019 volume | Standard2019Volume, ProPlus2019Volume | PerpetualVL2019                                     |
| Office LTSC 2021   | Standard2021Volume, ProPlus2021Volume | PerpetualVL2021                                     |
| Office LTSC 2024   | Standard2024Volume, ProPlus2024Volume | PerpetualVL2024                                     |
| Microsoft 365 Apps | O365ProPlusRetail, O365BusinessRetail | Current by default; MonthlyEnterprise or SemiAnnual |

```powershell
$target = @{
  TargetProductId = 'Standard2024Volume'
  Architecture    = '64'
  Language        = @('de-de')
  OdtPath         = 'C:\Managed\ODT\setup.exe'
  SourcePath      = 'C:\Managed\Media\Office2024'
}
New-Item -ItemType Directory -Path (Split-Path -Parent $target.SourcePath) -Force | Out-Null

# Optional application selection; use the same request through Check and installation.
$target.ExcludeApp = @('Groove', 'OneDrive')
```

| Option                            | Behavior                                                                                                                                |
| --------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| `Architecture`                    | `32` or `64`; defaults to `64`.                                                                                                         |
| `Version`                         | Optional exact `16.0` build; Prepare otherwise resolves and pins one.                                                                   |
| `ExcludeApp` / `ExcludePublisher` | Omit selected applications / add Publisher to the exclusions.                                                                           |
| `Language`                        | Defaults to exactly `en-us`, regardless of Windows or account locale. First entry is the primary shell language.                        |
| `AutoSourceLocales`               | Explicit opt-in to discovery; defaults to `LocaleSource InstalledOffice`. Requires unambiguous languages and primary-language evidence. |
| `LocaleSource OperatingSystem`    | With `AutoSourceLocales`, use the machine installation UI language, not the administrator's culture.                                    |

Explicit `Language` and automatic sourcing are mutually exclusive. Discovery failures do not silently fall back to English or OS detection.
Office languages do not change Windows locale, keyboards, or per-user editing/display preferences. ODT's XML chooses products/languages; one
media package can contain more languages than the installation requests.

```powershell
# Alternative: bilingual request, English primary.
$target.Language = @('en-us', 'de-de')

# Alternative: opt into installed-language discovery instead of an explicit list.
# Use only when the installed configuration has sufficient language evidence.
$automatic = $target.Clone()
$automatic.Remove('Language')
$automatic.AutoSourceLocales = $true
$automatic.LocaleSource = 'InstalledOffice'
```

OneDrive and Groove are separate exclusion IDs; neither uninstalls an independent OneDrive client. Teams application selection and its
Outlook meeting add-in need separate acceptance checks. See
[ODT options](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options),
[Groove controls](https://learn.microsoft.com/en-us/sharepoint/exclude-or-uninstall-previous-sync-client), and
[Office LTSC 2024 deployment](https://learn.microsoft.com/en-us/office/ltsc/2024/deploy).

## Prepare reusable media

Use the `$target` defined above. Select the intended installation language before continuing:

```powershell
$target.Language = @('de-de')

# Prepare English and German payloads without changing the installation request.
$media = $target.Clone()
$media.Language = @('en-us', 'de-de')
.\scripts\Office\Install-Office.ps1 -Mode Prepare @media -PassThru
```

| Media rule             | What to do                                                                                                                                                                                      |
| ---------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Local or UNC directory | A share such as `\\srv\deploy\Office2024` is supported; create its parent first and check access as the execution identity. SYSTEM/remote sessions may lack your mapped drives or share access. |
| Ownership and writes   | Package and manifest require Administrators/SYSTEM ownership and protected write access. Restrict share access and keep the package unchanged during staging.                                   |
| Existing media         | Reuse a verified compatible package. Use a new directory for incompatible/incomplete packages or old schema-1 `winkit-office-media.json` packages.                                              |
| Manifest               | Prepare publishes schema-2 `psfoundation-office-media.json` with build, languages, tool version, sizes, and hashes; partial downloads are not published as ready.                               |
| Integrity              | Hashes detect changed files, but cannot authenticate a manifest an attacker can replace. Deployment verifies a protected local copy before mutation.                                            |
| Space                  | Staging needs twice the media size plus 4 GiB on the staging drive; this is not an exact installed-size estimate.                                                                               |
| Missing payloads       | Validation stops. Deployment does not silently download missing languages or choose another build.                                                                                              |

## Install and verify

Close Office in all sessions, finish pending reboots, and avoid concurrent deployments. `Check` without a target inventories the machine;
with a target it returns eligibility, blockers, activation, and language transitions. `Check` takes no `OdtPath`.

```powershell
# Continue with the reviewed $target from above.
.\scripts\Office\Install-Office.ps1 -Mode Check -PassThru

$checkTarget = $target.Clone()
$checkTarget.Remove('OdtPath')
$check = .\scripts\Office\Install-Office.ps1 -Mode Check @checkTarget -PassThru
$check | ConvertTo-Json -Depth 30

.\scripts\Office\Install-Office.ps1 -Mode Install @target -DryRun

# After reviewing eligibility and the preview; this example uses default KMS licensing.
$result = .\scripts\Office\Install-Office.ps1 -Mode Install @target -Confirm -PassThru
$result | ConvertTo-Json -Depth 30
```

Install has no removal parameters and does not repair, reconfigure, or change the architecture of a conflicting/incomplete installation. A
compliant no-op verifies the full configuration, including languages and applications, without launching an installer or reapplying a key.
An eligible Check is not proof of execution success: host, inventory, applications, and media are rechecked during deployment.

Running applications block execution unless `-ForceCloseApps` explicitly authorizes termination and possible loss of unsaved work.
Deployment uses a shared lock, checks native installer activity, and never schedules a reboot.

### Use a MAK instead of KMS

Choose this execution alternative when a MAK is required; do not run both installation examples just to apply a key:

```powershell
$mak = Read-Host 'MAK for the destination' -AsSecureString
try {
  $result = .\scripts\Office\Install-Office.ps1 -Mode Install @target `
    -ProductKey $mak -Confirm -PassThru
}
finally {
  $mak.Dispose()
  Remove-Variable mak -ErrorAction SilentlyContinue
}
```

ODT needs the key in temporary XML, restricted to Administrators/SYSTEM and removed during cleanup. It is not a process argument; secure
erasure and redaction of ODT's own logs are not guaranteed. Existing keys are not automatically removed. Volume licensing must match the
target; Microsoft 365 returns `UserActivationRequired` for activation in the licensed user's session. Resolve activation-only failures
separately from installation failures.

## Remove selected products

`Remove-Office` removes only explicitly selected Click-to-Run products and their installed languages. There is no default selection or
standalone MSI removal switch. Unselected products must remain verifiably unchanged; uncertain shared-component effects can block removal.

```powershell
.\scripts\Office\Remove-Office.ps1 -Mode Check -PassThru

$removal = @{
  RemoveProductId = @('O365ProPlusRetail')
  OdtPath         = 'C:\Managed\ODT\setup.exe'
}
.\scripts\Office\Remove-Office.ps1 -Mode Check -RemoveProductId $removal.RemoveProductId -PassThru
.\scripts\Office\Remove-Office.ps1 -Mode Remove @removal -DryRun
.\scripts\Office\Remove-Office.ps1 -Mode Remove @removal -Confirm -PassThru
```

Already absent selections return `AlreadyAbsent` with `Changed=false`. Removal does not implement custom cleanup of profiles, PSTs,
documents, or keys. Keep backups before removing software.

## Results and automation

Mode is mandatory. `-DryRun` sets `WhatIfPreference`; both preview switches return results even without `-PassThru`. Deployment previews and
Check do not download, stop applications, write journals/logs, launch installers, or change Office/licensing. Validation can still block
them. Preparation and execution require confirmation; use `-Confirm:$false` only for reviewed unattended requests.

| Result                              | Meaning                                                                                                                                                                             |
| ----------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Check`                             | `Inventory` and optional `Plan` with `State`, `Eligible`, `Blockers`, and language transitions. Blocked plans exit 1.                                                               |
| `Prepare`                           | `Media`; reason distinguishes `MediaPrepared`, `AlreadyPrepared`, and `NotExecuted`. No deployment journal.                                                                         |
| Execution/recovery                  | One outcome with `ReasonCode`, `Phase`, `Changed`, `ChangeKnown`, `RebootRequired`, `NativeResults`, `Verification`, `Activation`, `RecoveryPath`, `LogPaths`, and cleanup details. |
| `Changed=null`, `ChangeKnown=false` | The operation may have changed the machine.                                                                                                                                         |
| Exit `0`                            | Completed, compliant/absent no-op, or preview; inspect status and activation too.                                                                                                   |
| Exit `1`                            | Blocked, failed, or unverified; inspect native outcomes and possible changes.                                                                                                       |
| Exit `3010`                         | Reboot required, possibly a blocked recovery; inspect status before continuing.                                                                                                     |

Execution keeps protected journals and JSONL logs under `%ProgramData%\PSFoundation-Office`; `-LogRoot` selects another local root. Consume
objects rather than console text, preserve nested results at depth 30, and pass only mode-appropriate parameters. Missing inputs fail
instead of prompting midway. `-Confirm:$false` does not bypass validation or language warnings.

```powershell
# After an execution example above, retain its outcome outside the winkit directory.
$reportDirectory = Join-Path $env:LOCALAPPDATA 'winkit\reports\Office'
New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
$reportPath = Join-Path $reportDirectory ('deployment-' + [guid]::NewGuid().ToString('N') + '.json')
$result | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $reportPath -Encoding UTF8
$result | Format-List Status, ReasonCode, Changed, ChangeKnown, RebootRequired, RecoveryPath, LogPaths
$reportPath
```
