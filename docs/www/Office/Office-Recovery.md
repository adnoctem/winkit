# Office recovery

[Office overview](README.md) | [Requirements](Requirements.md)

Recover continues or verifies a recorded installation/migration using its protected local journal. It is not Quick Repair, Online Repair,
rollback, or a journal-free reinstall. Use elevated 64-bit PowerShell from your **installed winkit directory** on 64-bit Windows.

## Read the original result

Use the original `RunId` and `LogRoot`, not a new target configuration. The default journal/log root is
`C:\ProgramData\PSFoundation-Office`. If you retained a result JSON:

```powershell
# Replace this with the actual retained result file.
$saved = Get-Content -LiteralPath 'E:\OfficeReports\migration-result.json' -Raw -Encoding UTF8 | ConvertFrom-Json
$saved | Format-List Status, ReasonCode, RunId, RecoveryPath, LogPaths, RebootRequired
```

Inspect native failures, verification, and activation separately. When a reboot is required, reboot before continuing in a fresh shell. Keep
the result, recovery journal, and native diagnostics available; repeatedly invoking Recover does not make an unsupported phase safe.

## Preview the recorded action

Replace the example RunId and choose the wrapper matching the original action:

```powershell
$recovery = @{
  RunId   = '0123456789abcdef0123456789abcdef'
  OdtPath = 'C:\Managed\ODT\setup.exe'
  LogRoot = 'C:\ProgramData\PSFoundation-Office'
}

# Original migration:
$preview = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Recover @recovery -DryRun -PassThru
$preview | ConvertTo-Json -Depth 30
```

```powershell
# Alternative for an original clean installation:
$preview = .\scripts\Office\Install-Office.ps1 -Mode Recover @recovery -DryRun -PassThru
$preview | ConvertTo-Json -Depth 30
```

| Boundary                                | Meaning                                                                                                                                     |
| --------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| Target, build, languages, removal scope | Loaded from the journal. Target/removal overrides, explicit `Language`, and `AutoSourceLocales` are rejected. Languages are not redetected. |
| Installation recovery                   | Cannot inherit migration removal authority.                                                                                                 |
| State and media                         | Revalidated. Changed media, unexpected products, active deployments, and pending reboots can block continuation.                            |
| Unsupported record/phase                | Obsolete authorization/schema and uncertain partial-installer phases are not converted into ordinary migrations.                            |
| Recovery authority                      | A journal or `RecoveryRequired` flag alone is not permission to replay a completed deployment.                                              |

## Continue after reviewing the preview

For a migration with default KMS licensing:

```powershell
$result = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Recover @recovery -Confirm -PassThru
$result | ConvertTo-Json -Depth 30
```

If the destination requires a MAK, use this alternative and supply it again; the journal never stores it:

```powershell
$mak = Read-Host 'MAK for the recorded destination' -AsSecureString
try {
  $result = .\scripts\Office\Switch-OfficeVersion.ps1 -Mode Recover @recovery `
    -ProductKey $mak -Confirm -PassThru
}
finally {
  $mak.Dispose()
  Remove-Variable mak -ErrorAction SilentlyContinue
}
```

For installation recovery, substitute `Install-Office.ps1` in the selected execution example. Retain and inspect the result using
[results and automation](Office-Installation-and-Removal.md#results-and-automation), then complete the
[migration acceptance checks](Office-Migrations.md#verify-and-retain-the-result). Exit `3010` may describe blocked recovery, not completed
installation; read `Status` before continuing.
