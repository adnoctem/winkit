# Active Directory exports and the account map

`Export-AD.ps1` writes directory exports. `Export-ADAccountMap.ps1` writes the `ad-account-map.json` object consumed by `plead`'s inbox
workflow.

Both scripts need the `ActiveDirectory` PowerShell module and connectivity to AD Web Services. They do not require elevation; the account
must be able to read the selected objects. Server Core is supported when the module and AD Web Services connectivity are available. SYSTEM
execution is supported when the module and network/AD permissions are available.

## Export users, computers, and groups

`Export-AD.ps1` writes one file per resource with a stable column order:

- `AD-Users.csv` or `AD-Users.json`
- `AD-Computers.csv` or `AD-Computers.json`
- `AD-Groups.csv` or `AD-Groups.json`

Users and computers include enabled accounts only by default. Use `-IncludeDisabled` to include disabled accounts. `-Omit` removes
attributes from either output format. Missing requested attributes stop the export by default; use `-Omit` for attributes you do not need,
or `-IncludeMissing` to export records with empty values.

```powershell
# Enabled users, computers, and groups as CSV.
.\scripts\Server\ADDS\Export-AD.ps1 -OutputPath C:\ADReports

# One OU as JSON, without noisy logon timestamps.
.\scripts\Server\ADDS\Export-AD.ps1 `
  -OutputPath C:\ADReports `
  -Resource Users `
  -Format Json `
  -SearchBase 'OU=Users,DC=adnoctem,DC=co' `
  -Omit lastLogon,lastLogonTimestamp

# Preview without writing files.
.\scripts\Server\ADDS\Export-AD.ps1 -OutputPath C:\ADReports -Resource Users -DryRun
```

Use `-Filter` or `-LDAPFilter` to narrow the query. `-Filter` defaults to `*`; the two options are mutually exclusive. Use `-Server` and
`-Credential` for an alternate domain controller or credential. `-Resource` accepts `Users`, `Computers`, `Groups`, or `All` (the default).

`-PassThru` returns one structured operation result per written file, including the path and record count. CSV output uses `;` as the
delimiter and honours `-Encoding`; JSON is always UTF-8 without BOM.

## Generate the account map

`Export-ADAccountMap.ps1` resolves selected enabled users to full email addresses and writes a JSON object of `samAccountName` to email. The
AD `mail` attribute is authoritative. An explicit override CSV can replace it when the attribute is unreliable. Addresses are never guessed
from a domain name or from `userPrincipalName`.

```json
{
  "MGindorfer": "maximilian.gindorfer@adnoctem.co",
  "MMustermann": "maximilian.mustermann@adnoctem.co",
  "EMustermann": "erika.mustermann@example.com"
}
```

The output is UTF-8 without BOM, uses CRLF line endings, and sorts keys with a stable ordinal comparison. Re-running the same selection
produces the same bytes. Existing files are only replaced with `-Force`; use `-DryRun` to resolve and preview without writing.

```powershell
# All enabled users, with the two accepted mail domains.
.\scripts\Server\ADDS\Export-ADAccountMap.ps1 `
  -AllowedDomain adnoctem.co,example.com `
  -OutputPath C:\ADReports\ad-account-map.json `
  -ReportPath C:\ADReports\ad-account-map-audit.csv

# Preview after changing the override file.
.\scripts\Server\ADDS\Export-ADAccountMap.ps1 `
  -AllowedDomain adnoctem.co,example.com `
  -OverridePath C:\ADReports\account-map-overrides.csv `
  -OutputPath C:\ADReports\ad-account-map.json `
  -ReportPath C:\ADReports\ad-account-map-audit.csv `
  -DryRun
```

### Selection

Choose one of these selections:

- `-Filter` (default `*`) for a whole-domain query.
- `-LDAPFilter` for an LDAP expression.
- `-SearchBase` with a filter to limit the query to an OU or container.
- `-Group` and `-Recursive` for group membership.
- `-Identity` for explicit users.

`-Identity` cannot be combined with a group, search base, or filter. Disabled users are excluded by default; use `-IncludeDisabled` to
include them.

### Allowed domains

`-AllowedDomain` is required and is the only source of truth for accepted mail domains. Matching is exact and case-insensitive:
`adnoctem.co` does not accept `mail.adnoctem.co`. Only ASCII/Punycode domain spellings are supported; Unicode domain spellings are rejected.
List every accepted domain explicitly.

### Override CSV

The override file is a UTF-8 CSV with case-insensitive headers:

```csv
samAccountName,email,note
MGindorfer,maximilian.gindorfer@adnoctem.co,mail attribute is stale
MMustermann,maximilian.mustermann@adnoctem.co,
EMustermann,erika.mustermann@example.com,second mail domain
ServiceAccount,,intentionally excluded
```

- `samAccountName` and `email` are required; `note` is optional.
- Duplicate account keys are rejected.
- An empty `email` marks an intentional exclusion.
- An override wins over the AD `mail` attribute.
- An override outside `-AllowedDomain` is unresolved unless `-AllowOverrideOutsideDomain` is supplied.
- An override that matches no selected user is reported and excluded by default. Use `-IncludeUnmatchedOverride` to include it.

### Missing mail, duplicates, and review

Missing `mail` fails the run by default. Supply `-IncludeMissing` to write the resolved subset and omit the unresolved accounts; the audit
report lists every omission and its reason. A present-but-invalid or wrong-domain `mail` value is not rewritten or replaced.

Duplicate `samAccountName` values fail by default. `-OnDuplicate First` or `-OnDuplicate Last` chooses deterministically by distinguished
name.

The optional `-ReportPath` CSV records every selected account plus unmatched overrides, with `samAccountName`, `email`, `source`, `status`,
`reason`, `rawMail`, `distinguishedName`, and `note`. It is written even when the map is not, so a failed run still leaves an audit
artifact.

## plead inbox contract

`plead`'s inbox schema v1 gives `email` precedence over `samAccountName` plus `domain`. Consuming scripts can write the mapped address
directly as the absence entry's `email` while retaining `samAccountName` and `domain` for traceability. The account map is only a
key-to-address object; it is not an absence file and contains no dates, templates, or schema version.

Accounts missing from the map have no entry in the consumer. That is the expected result for disabled accounts, intentional overrides with
an empty email, and unresolved accounts under `-IncludeMissing`.
