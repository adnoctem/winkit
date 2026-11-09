# AD account-map fixtures

`ad-account-map.synthetic.json` is a byte-level golden file for the synthetic test data in `ExportADAccountMap.Tests.ps1`. It is UTF-8
without BOM, uses CRLF line endings, and has one final CRLF.

`ad-account-map.golden.json` is the equivalent golden file for the owner-provided real AD export. The real-fixture test in
`ExportADAccountMap.Tests.ps1` is enabled and compares the exported bytes with it.

## `ad-records.csv`

A UTF-8 CSV in the `Export-AD.ps1` CSV format: semicolon-delimited, case-insensitive headers, and extra columns are allowed. The fixture
test reads it with `Import-Csv -Delimiter ';'`.

| Column              | Required | Meaning                                         |
| ------------------- | -------- | ----------------------------------------------- |
| `sAMAccountName`    | yes      | AD account key, including its original case     |
| `Enabled`           | yes      | `True` or `False`                               |
| `mail`              | yes      | AD mail attribute; may be empty                 |
| `displayName`       | yes      | Display name used in the audit report           |
| `DistinguishedName` | yes      | Distinguished name used for deterministic order |

The committed fixture contains nine enabled users with valid `adnoctem.co` addresses, including UTF-8 display names. Disabled users, empty
or outside-domain mail, and duplicate keys are covered by the synthetic unit tests in `ADDS.Tests.ps1`, not this real export.

## `account-map-overrides.csv`

A comma-delimited UTF-8 CSV with the production `samAccountName,email,note` schema. The committed file contains one outside-domain override
for `MGindorfer` to `admin@adnoctem.net`.

The real-fixture test verifies that this override is rejected with `OverrideDomainNotAllowed` unless `-AllowOverrideOutsideDomain` is
supplied, and that the override wins over the AD `mail` attribute when it is supplied. Intentional empty-email exclusions and unmatched
overrides are covered by the synthetic unit tests.

## `ad-account-map.golden.json`

The expected output of `New-ADAccountMap` plus `Export-ADAccountMap` for `ad-records.csv` and `account-map-overrides.csv`, using
`AllowedDomain adnoctem.co,example.com` and `-AllowOverrideOutsideDomain`. It must be UTF-8 without BOM, CRLF, and key-sorted exactly like
the synthetic golden file.
