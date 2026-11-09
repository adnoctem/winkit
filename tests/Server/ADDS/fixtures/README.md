# AD account-map fixtures

`ad-account-map.synthetic.json` is a byte-level golden file for the synthetic test data in `ExportADAccountMap.Tests.ps1`. It is UTF-8
without BOM, uses CRLF line endings, and has one final CRLF.

The real-AD golden test is intentionally skipped until the owner supplies an AD export fixture. Add these files under this directory:

## `ad-records.csv`

A UTF-8 CSV (BOM or no BOM) with case-insensitive headers:

| Column              | Required | Meaning                                         |
| ------------------- | -------- | ----------------------------------------------- |
| `sAMAccountName`    | yes      | AD account key, including its original case     |
| `Enabled`           | yes      | `True` or `False`                               |
| `mail`              | yes      | AD mail attribute; may be empty                 |
| `displayName`       | yes      | Display name used in the audit report           |
| `distinguishedName` | yes      | Distinguished name used for deterministic order |

Include at least these rows:

1. An enabled user with a valid `adnoctem.co` address.
2. An enabled user with a valid `example.com` address.
3. An enabled user with an empty `mail` value.
4. An enabled user whose `mail` domain is outside the allowlist.
5. A disabled user with a valid address.
6. Two rows with the same `sAMAccountName` and different `distinguishedName` values.
7. A case-variant `sAMAccountName` that must keep its AD casing in the output.

## `ad-account-map.golden.json`

The expected output of `New-ADAccountMap` plus `Export-ADAccountMap` for `ad-records.csv`, using `AllowedDomain adnoctem.co,example.com`. It
must be UTF-8 without BOM, CRLF, and key-sorted exactly like the synthetic golden file.

## `account-map-overrides.csv` (optional)

If supplied, the test also covers override precedence, an intentional empty email exclusion, an unmatched override, and an outside-domain
override. It uses the same `samAccountName,email,note` schema as the production override file.
