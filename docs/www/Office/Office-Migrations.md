# Office migrations

[Office overview](README.md) | [Requirements](Requirements.md)

Use `Switch-OfficeVersion.ps1` to replace an Office product or change architecture. Run from your **installed winkit directory** in elevated
64-bit PowerShell on 64-bit Windows, even when the source Office is 32-bit. Prepare a tested machine backup and an
[Outlook checkpoint](Outlook-Backup-and-Repair.md#capture-data-and-settings) first; retain old media and licenses.

## Inventory and select removals

```powershell
$inventory = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check -PassThru
$inventory | ConvertTo-Json -Depth 30
```

| Source                                                    | Explicit removal authority                                                                                                                                                        |
| --------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Click-to-Run                                              | `-RemoveProductId` with exact installed IDs. Select the suite itself for a same-product architecture change.                                                                      |
| MSI Office 2007/2010/2013/2016, including Enterprise 2007 | `-RemoveMsi` authorizes **all supported MSI Office removals**, including supported Visio, Project, language packs, runtimes, and database engines. Review dependent applications. |
| Additional products                                       | Unapproved products such as Visio/Project block migration. Unknown/unsupported MSI components and stale selections also block it; refresh inventory after partial work.           |

Recognition does not override [host requirements](Requirements.md) or verification gates. Install and Migrate are separate operations:
`Install-Office` does not silently convert an existing configuration. See
[Microsoft's MSI removal scope](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/upgrade-from-msi-version).

## Prepare, check, and preview

This example changes Standard 2019 to 64-bit German Office. Select the actual installed product IDs and licensed destination. Obtain a
[verified ODT](Office-Installation-and-Removal.md#obtain-and-validate-odt) first. Keep product, build, languages, exclusions, and media
identical through Prepare, Check, preview, and execution.

```powershell
$target = @{
  TargetProductId = 'Standard2019Volume'
  Architecture    = '64'
  Language        = @('de-de')
  OdtPath         = 'C:\Managed\ODT\setup.exe'
  SourcePath      = 'C:\Managed\Media\Office2019'
}
$source = @{
  RemoveProductId = @('Standard2019Volume')
}
New-Item -ItemType Directory -Path (Split-Path -Parent $target.SourcePath) -Force | Out-Null

.\scripts\Office\Switch-OfficeVersion.ps1 -Mode Prepare @target -PassThru

$checkTarget = $target.Clone()
$checkTarget.Remove('OdtPath')
$check = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check @checkTarget @source -PassThru
$check | ConvertTo-Json -Depth 30

$preview = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source -DryRun -PassThru
$preview | ConvertTo-Json -Depth 30
```

For Home & Business 2019 or legacy MSI, replace `$source` **before Check and preview**, then run those checks again:

```powershell
# Alternative Click-to-Run source; confirm the exact ID in inventory.
$source = @{ RemoveProductId = @('HomeBusiness2019Retail') }
```

```powershell
# Alternative MSI source, such as Office 2007 Enterprise; broad removal scope.
$source = @{ RemoveMsi = $true }
```

Language defaults to `en-us` when omitted. Explicit `de-de` is appropriate for German MSI migrations whose language evidence is incomplete.
`-AutoSourceLocales` is opt-in and requires unambiguous evidence; see
[target configuration](Office-Installation-and-Removal.md#choose-the-target). Review language additions, removals, primary-language changes,
proofing resources, edition/application differences, and add-in architecture compatibility.

## Execute the reviewed request

Choose one licensing alternative. For MAK activation:

```powershell
$mak = Read-Host 'Office volume key' -AsSecureString
try {
  $result = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source `
    -ProductKey $mak -Confirm -PassThru
}
finally {
  $mak.Dispose()
  Remove-Variable mak -ErrorAction SilentlyContinue
}
```

For default KMS licensing, omit the key:

```powershell
$result = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source -Confirm -PassThru
```

Destination media is validated and staged before removal; inventory and applications are rechecked. MSI removal is part of the destination
ODT configuration. A removal failure or reboot requirement stops continuation. After any failure, inspect the result and journal before
retrying; a failure after removal may need [recovery](Office-Recovery.md).

## Verify and retain the result

```powershell
$reportDirectory = Join-Path $env:LOCALAPPDATA 'winkit\reports\Office'
New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
$reportPath = Join-Path $reportDirectory ('migration-' + [guid]::NewGuid().ToString('N') + '.json')
$result | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $reportPath -Encoding UTF8
$result | Format-List Status, ReasonCode, NativeResults, Verification, Activation, RebootRequired
$reportPath
```

Review installed applications, architecture, language/proofing, licensing, Outlook data/settings, add-ins, and send/receive behavior.
Preserve JSON, journals, and native diagnostics before reverting anything. Native installer success with incomplete verification is a manual
review point; activation-only failures are not a reason to reinstall. See
[results and exit codes](Office-Installation-and-Removal.md#results-and-automation).

After native verification and manual acceptance, preview the **identical** request:

```powershell
$repeat = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Migrate @target @source -DryRun -PassThru
$repeat | Format-List Status, ReasonCode, AlreadyCompliant, Changed, ChangeKnown, NativeResults
```

Expect `Completed`, `AlreadyCompliant`, `AlreadyCompliant=true`, `Changed=false`, `ChangeKnown=true`, and no native results. Only then
consider a real repeat; the compliant path launches no installer and does not reapply a key. Unknown verification, activation failures, or a
new removal/install confirmation require investigation.

## When inventory is incomplete

```powershell
$observed = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Check -PassThru
$observed.Inventory | ConvertTo-Json -Depth 30
```

| Observation                                                      | Interpretation                                                                                                                  |
| ---------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `Products[].VersionSource`, `Evidence`                           | Provenance of installed-build observations; preserved in Check and checkpoint manifests.                                        |
| `ClickToRunInventory`                                            | Documented installed inventory.                                                                                                 |
| `ActiveProductResources`                                         | Fallback from agreeing active-product resource versions only when documented inventory is absent, not conflicting or malformed. |
| `VersionToReport`                                                | Telemetry, not proof of installed build.                                                                                        |
| Missing languages, exclusions, proofing, or other postconditions | Unknown remains unknown. Resolve evidence gaps rather than bypassing blockers.                                                  |
| Legacy MSI LCID                                                  | A single LCID cannot establish complete locale preservation; choose explicit languages.                                         |

The scripts do not purchase licenses, upgrade Windows, convert Outlook profiles, or supply automatic rollback.
