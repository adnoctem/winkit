# Export-AD structural API and AD account map

**Status:** approved and implemented in the Server/ADDS library and entry scripts. The owner decisions and implementation deviations are
recorded in section 9.

## 1. Purpose and scope

This plan turns `scripts/Server/ADDS/Export-AD.ps1` from a single monolithic export switch into a small, testable API made of advanced
functions, and adds a first-class generator for the JSON account map consumed by `plead`.

The work has two deliverables:

1. A structural refactor that keeps the current `Export-AD.ps1` command line and CSV output working, while adding proper parameter sets,
   pipeline input, typed records, comment-based help, preview support, and structured errors.
2. A new `Export-ADAccountMap.ps1` entry point and its supporting functions, which produce `ad-account-map.json` - a JSON object of
   `samAccountName` to full email address - with explicit source precedence, validation, auditing, and deterministic bytes.

Non-goals: writing to AD, provisioning mailboxes, talking to Plesk/`plead`, guessing addresses from a domain name, changing the `plead`
inbox schema, or adding a PowerShell module to winkit. The `delta4x4/ad` entry-creation script and `plead` are consumers of the generated
file; this plan does not change either repository.

## 2. Assessment of the current script

### 2.1 Public surface

`scripts/Server/ADDS/Export-AD.ps1` currently exposes three parameters:

| Parameter     | Type and default                                     | Actual behavior                                                                      |
| ------------- | ---------------------------------------------------- | ------------------------------------------------------------------------------------ |
| `-OutputPath` | `[string]`, position 0, mandatory                    | Directory that receives `AD-Users.csv`, `AD-Computers.csv`, `AD-Groups.csv`          |
| `-Resource`   | `[string]`, position 1, `Users/Computers/Groups/All` | Defaults to `All`; selects one export or all three                                   |
| `-Encoding`   | `[string]`, position 2, `UTF8`                       | Passed through to `Export-Csv -Encoding`; accepted values mirror `Export-Csv` in 5.1 |

The file has no `[CmdletBinding()]`, no `-WhatIf`/`-Confirm`, no `-DryRun`, no `-PassThru`, no pipeline input, no `-Server`/`-Credential`,
no `-SearchBase`/OU selection, and no `-Filter` parameter. The help text mentions a `Filter` parameter that does not exist. It also
documents `.PARAMETER OutputFile` while the actual parameter is `-OutputPath`, and its examples describe filtering that cannot be performed.

### 2.2 Data model

The script keeps a static `$properties` table per resource and calls `Get-ADUser`, `Get-ADComputer`, or `Get-ADGroup` with `-Filter *` and
the matching property list. It then writes the returned ActiveDirectory module objects directly through `Export-Csv`, using `;` as the
delimiter. Consequences:

- The CSV schema is whatever the ActiveDirectory module populates on `ADUser`, `ADComputer`, and `ADGroup`, in the order the module exposes
  it. It is not a contract the script controls and can differ between module versions or future queries.
- There is no intermediate record type. Anything that wants the same data must reimplement the query and the property table.
- `All` materializes all three result sets in one `PSCustomObject` before writing anything, so memory grows with the whole directory rather
  than with the largest single resource.
- File names use `(Get-Culture).TextInfo.ToTitleCase()`, so naming is tied to the host culture even though the resulting names currently
  happen to be `AD-Users.csv` and friends.

### 2.3 Limitations and defects

- No reusable functions, parameter sets, or pipeline support despite the repository's script conventions.
- `#Requires -Version 2.0` is obsolete; the PowerShell 2 compatibility branch is dead code. The repository runtime is PowerShell 5.1 or 7.
- `#Requires -RunAsAdministrator` is unrelated to AD read permissions, and the other AD report scripts in the same folder do not require
  elevation.
- `Import-Module ActiveDirectory -Force` is unconditional and not wrapped in a terminating error handler.
- AD queries are not wrapped in `try`/`catch` and do not pass `-ErrorAction Stop`, so partial failures can surface as raw module errors.
- `OutputPath` is validated only after the AD queries finish, and the directory is not created when missing.
- Empty results call `Write-Error` inside a `ForEach-Object`; the script continues and exits with code 0.
- Existing CSVs are overwritten silently; there is no preview or confirmation.
- `-Encoding` accepts `UTF7`/`UTF32` and similar legacy encodings and delegates their BOM behavior to the host PowerShell version.
- The script does not import `AdNoctem.Substrate.PowerShell`, does not use `Write-Log`, and has no `.NOTES` block covering Server Core or
  SYSTEM suitability.

### 2.4 Risks

- **Schema drift:** consumers parsing the CSVs are coupled to ActiveDirectory module internals rather than a documented contract.
- **Memory and runtime:** `-Filter *` on all three resources is expensive and unbounded in memory on large directories.
- **Sensitive data:** the export contains mail addresses, phone numbers, home directories, and account metadata, and is written without an
  ACL or overwrite guard.
- **Silent partial success:** a failed query can still leave a partial or missing file while the process exits 0.
- **Culture and encoding drift:** file naming and BOM behavior depend on host culture and PowerShell version.
- **No test coverage:** there are currently no Pester tests for this script under `tests/`.

### 2.5 What "structural API" should mean here

For winkit, a structural API means:

- Advanced functions with approved verbs, singular nouns, `[CmdletBinding()]`, `[OutputType()]`, parameter sets, pipeline input,
  comment-based help, and terminating errors with actionable messages.
- A small, typed data contract (`Winkit.ADDS.User`, `Winkit.ADDS.Computer`, `Winkit.ADDS.Group`, `Winkit.ADDS.AccountMapEntry`,
  `Winkit.ADDS.AccountMapResult`) instead of raw module objects.
- Thin entry scripts that own parameter parsing, logging, `ShouldProcess`, exit codes, and file output; the reusable logic lives in
  functions.
- Stable ordering, encoding, and file naming so results are reviewable in diffs and testable with golden files.
- No new module in this repository. The repository convention is that shared module functionality belongs to
  `AdNoctem.Substrate.PowerShell`; the proposal below therefore uses a dot-sourced script library, with promotion to Substrate called out as
  an explicit follow-up decision in section 8.

## 3. Proposed API

### 3.1 Layout

```text
scripts/Server/ADDS/
  lib/
    ADDS.ps1                        # dot-sourced function library; no workflow
  Export-AD.ps1                     # compatibility entry point, now a thin wrapper
  Export-ADAccountMap.ps1           # new mapping entry point
tests/Server/ADDS/
  ExportAD.Tests.ps1
  ADDS.Tests.ps1
  ExportADAccountMap.Tests.ps1
  fixtures/
    ad-account-map.golden.json
```

Rationale:

- The library travels with the installed `scripts/` tree (the installer copies `scripts/` recursively), so both entry scripts keep working
  from a normal winkit installation without Git or a module build step.
- It follows the existing sibling-script pattern (`Split-OutlookArchive.ps1` calls `New-OutlookArchive.ps1`) while keeping function calls
  in-process, with typed pipeline objects.
- The library contains no `#Requires -RunAsAdministrator` and no main workflow; entry scripts own requirements, elevation, and `exit` codes.
- If the owner prefers not to introduce a `lib/` folder, the fallback is to promote the reusable functions to
  `AdNoctem.Substrate.PowerShell` (section 8, question 1). The API names and schemas below remain the same.

All user-facing scripts keep the repository conventions: `#Requires -Version 5.1`, `#Requires -Module ActiveDirectory`,
`#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.1' }`,
`[CmdletBinding(SupportsShouldProcess = $true)]`, `-DryRun`, `-PassThru`, comment-based help, and `Write-Log` feedback.

### 3.2 Shared library functions

Every function below is an advanced function in `lib/ADDS.ps1`, uses PowerShell 5.0-compatible syntax, approved verbs, and a singular noun.
Parameters use one attribute/type/name per line per `CONTRIBUTING.md`.

#### `Get-ADResourceProperty`

```powershell
function Get-ADResourceProperty {
  [CmdletBinding()]
  [OutputType([string[]])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('Users', 'Computers', 'Groups')]
    [string]
    $Resource
  )
}
```

Returns the ordered attribute list used for the CSV row and for the mapping query. This centralizes the current `$properties` table and lets
tests assert it. It does not query AD.

#### `Get-ADDirectoryRecord`

```powershell
function Get-ADDirectoryRecord {
  [CmdletBinding(DefaultParameterSetName = 'Filter')]
  [OutputType('Winkit.ADDS.User', 'Winkit.ADDS.Computer', 'Winkit.ADDS.Group')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('Users', 'Computers', 'Groups', 'All')]
    [string]
    $Resource,

    [Parameter(ParameterSetName = 'Filter')]
    [string]
    $Filter = '*',

    [Parameter(Mandatory = $true, ParameterSetName = 'LDAPFilter')]
    [string]
    $LDAPFilter,

    [Parameter(ParameterSetName = 'Filter')]
    [Parameter(ParameterSetName = 'LDAPFilter')]
    [string]
    $SearchBase,

    [Parameter(
      Mandatory = $true,
      ParameterSetName = 'Identity',
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true
    )]
    [Alias('SamAccountName', 'DistinguishedName', 'ObjectGUID', 'Name')]
    [string[]]
    $Identity,

    [Parameter(Mandatory = $true, ParameterSetName = 'Group')]
    [string]
    $Group,

    [Parameter(ParameterSetName = 'Group')]
    [switch]
    $Recursive,

    [Parameter(ParameterSetName = 'Filter')]
    [Parameter(ParameterSetName = 'LDAPFilter')]
    [Parameter(ParameterSetName = 'Identity')]
    [Parameter(ParameterSetName = 'Group')]
    [string]
    $Server,

    [pscredential]
    $Credential,

    [string[]]
    $Properties
  )
}
```

Behavior:

- `Filter` (default) and `LDAPFilter` query the selected resource with `Get-ADUser`/`Get-ADComputer`/`Get-ADGroup`. `-SearchBase` limits the
  search to an OU or container.
- `Identity` accepts one or more identities from the pipeline; `ValueFromPipelineByPropertyName` with aliases lets
  `Get-ADUser | Get-ADDirectoryRecord -Resource Users` work. It resolves each identity with `-Identity`.
- `Group` selects members of one group. `-Recursive` uses `Get-ADGroupMember -Recursive` and then resolves each member to a full record;
  members whose `objectClass` does not match `-Resource` are skipped with `Write-Verbose`.
- `-Resource All` is valid only with `Filter`/`LDAPFilter`; combining it with `Identity` or `Group` terminates with a clear error because
  the result type would be ambiguous.
- All AD calls use `-ErrorAction Stop` inside `try`/`catch`; failures become terminating errors that name the resource and query scope. No
  partial result is returned silently.
- `-Properties` appends attributes to the resource defaults; duplicates are removed case-insensitively.
- `-Server` and `-Credential` are passed through to the AD cmdlets; no plaintext password parameter exists.

Output records are documented in section 3.3. The function streams records as the AD cmdlets return them; no whole-result array is allocated
by the function itself.

#### `Export-ADDirectoryRecord`

```powershell
function Export-ADDirectoryRecord {
  [CmdletBinding(SupportsShouldProcess = $true)]
  [OutputType('Winkit.ADDS.ExportResult')]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [psobject[]]
    $InputObject,

    [Parameter(Mandatory = $true)]
    [ValidateSet('Users', 'Computers', 'Groups')]
    [string]
    $Resource,

    [Parameter(Mandatory = $true, Position = 0)]
    [string]
    $OutputPath,

    [ValidateSet('UTF8', 'ASCII', 'Unicode', 'UTF7', 'UTF32', 'BigEndianUnicode')]
    [string]
    $Encoding = 'UTF8',

    [switch]
    $PassThru
  )
}
```

Behavior:

- Collects records for one resource, flattens each `Winkit.ADDS.*` record into an ordered row, and writes `AD-<Resource>.csv` in
  `-OutputPath` with `Export-Csv -NoTypeInformation -Delimiter ';' -Encoding $Encoding`.
- Uses the invariant culture for the title-cased file name, so names are always `AD-Users.csv`, `AD-Computers.csv`, and `AD-Groups.csv`.
- Validates `OutputPath` as an existing directory in `begin`, before any expensive work when the caller pipes from a query.
- An empty resource writes no file, emits a warning, and returns a `Skipped` result when `-PassThru` is used. It is not an error.
- `ShouldProcess` targets each CSV file. `-WhatIf` and `-DryRun` compute everything and write nothing.
- Existing files keep the current overwrite behavior (Export-AD historically overwrote them). If the owner wants an overwrite guard here,
  `-Force` semantics can be added in phase 2 without changing callers (section 8, question 2).

#### `Test-ADMailAddress`

```powershell
function Test-ADMailAddress {
  [CmdletBinding()]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string]
    $Address,

    [string[]]
    $AllowedDomain
  )
}
```

Validation rules:

- Trim the value; reject empty, whitespace, CR/LF, and an address longer than 254 characters.
- Parse with `System.Net.Mail.MailAddress` and require that `Address` round-trips exactly, so display names, comments, multiple addresses,
  and malformed input are rejected.
- Require a non-empty local part and host. Host comparison is case-insensitive.
- When `-AllowedDomain` is supplied, the host must exactly match one entry. Subdomain matching is deliberately not implicit; the owner must
  list every accepted domain (section 8, question 3).

The function returns `$true`/`$false` and does not write to the success stream or log. Callers add context.

#### `Import-ADAccountMapOverride`

```powershell
function Import-ADAccountMapOverride {
  [CmdletBinding()]
  [OutputType('Winkit.ADDS.AccountMapOverride')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]
    $Path
  )
}
```

Reads a UTF-8 CSV (BOM or no BOM) with these case-insensitive headers:

| Column           | Required | Meaning                                                        |
| ---------------- | -------- | -------------------------------------------------------------- |
| `samAccountName` | yes      | AD account key; trimmed, case-insensitive duplicate detection  |
| `email`          | yes      | Full address; empty means "intentionally exclude this account" |
| `note`           | no       | Free text copied into the audit report                         |

Rules:

- Missing file, missing required headers, an empty key, a malformed address, or a duplicate key terminates with an error that names the row.
- An empty `email` is a deliberate exclusion, not a missing value; it produces an `Excluded` entry so the intent is visible.
- The function does not check allowed domains; resolution does that, because the allowlist is a resolution parameter.
- JSON override files are intentionally not accepted in v1. The canonical manual file is CSV; the emitted JSON is an output artifact, not a
  hand-editing format. Section 8, question 4 asks whether JSON input is required.

#### `New-ADAccountMap`

```powershell
function New-ADAccountMap {
  [CmdletBinding(DefaultParameterSetName = 'Filter')]
  [OutputType('Winkit.ADDS.AccountMapResult')]
  param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Input', ValueFromPipeline = $true)]
    [psobject]
    $InputObject,

    [Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
    [string[]]
    $Identity,

    [Parameter(Mandatory = $true, ParameterSetName = 'Group')]
    [string]
    $Group,

    [Parameter(ParameterSetName = 'Group')]
    [switch]
    $Recursive,

    [Parameter(ParameterSetName = 'Filter')]
    [string]
    $Filter = '*',

    [Parameter(Mandatory = $true, ParameterSetName = 'LDAPFilter')]
    [string]
    $LDAPFilter,

    [Parameter(ParameterSetName = 'Filter')]
    [Parameter(ParameterSetName = 'LDAPFilter')]
    [string]
    $SearchBase,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $AllowedDomain,

    [string]
    $OverridePath,

    [string]
    $DefaultDomain,

    [ValidateSet('Fail', 'Omit')]
    [string]
    $OnMissingMail = 'Fail',

    [ValidateSet('Fail', 'First', 'Last')]
    [string]
    $OnDuplicate = 'Fail',

    [switch]
    $AllowOverrideOutsideDomain,

    [switch]
    $IncludeUnmatchedOverride,

    [switch]
    $EnabledOnly,

    [string]
    $Server,

    [pscredential]
    $Credential
  )
}
```

Behavior:

- With `InputObject`, the function accepts `Winkit.ADDS.User` records from the pipeline (or AD-like objects with `SamAccountName` and
  `mail`). With the other sets it queries AD through `Get-ADDirectoryRecord -Resource Users`.
- Selection is one of: explicit identities, one group (optionally recursive), an OU with an AD filter, or an LDAP filter.
- `-AllowedDomain` is mandatory and has no environment-specific default. It is the audit boundary for every resolved address.
- `-DefaultDomain` is optional and must be one of `-AllowedDomain`. If supplied, a user with no `mail` attribute and no override receives
  `<samAccountName>@<DefaultDomain>` with `Source = 'DefaultDomain'`. It is never used when `mail` exists but is invalid or points at a
  disallowed domain, and it is never used implicitly. This is the only sanctioned shorthand, and every use is visible in the result and
  report (section 8, question 5).
- `-OnMissingMail Fail` (default) marks the result `Failed`, so `Export-ADAccountMap` refuses to write a map. `Omit` writes the resolved
  subset and records the rest as unresolved.
- `-OnDuplicate Fail` (default) marks the run `Failed`; `First`/`Last` choose deterministically after sorting candidates by
  `DistinguishedName` with ordinal comparison.
- `-AllowOverrideOutsideDomain` permits an explicit override whose domain is outside the allowlist. It is off by default and every such
  entry is marked in the report.
- `-IncludeUnmatchedOverride` includes override rows whose `samAccountName` was not selected by the AD query. Off by default, so typos and
  stale leavers are reported rather than silently emitted.
- `-EnabledOnly` restricts the selection to `Enabled -eq $true`; it is off by default, matching the current `-Filter *` behavior that
  includes disabled users. (Section 8, question 6 asks whether the mapping should default to enabled accounts only.)
- The function does not open, write, or overwrite the output file. It returns one typed result object so the same resolution can be reviewed
  with `-WhatIf` and then exported without querying again.

Resolution order for each selected user:

1. An override row with the same case-insensitive `samAccountName` wins. An empty override email marks the account excluded.
2. Otherwise, a syntactically valid `mail` attribute whose domain is in `-AllowedDomain` is used with `Source = 'ADMail'`.
3. Otherwise, if the `mail` attribute is absent, `-DefaultDomain` was explicitly supplied, and no override exists, use
   `<samAccountName>@<DefaultDomain>` with `Source = 'DefaultDomain'`.
4. Otherwise, the entry is unresolved with a stable reason: `MissingMail`, `MailDomainNotAllowed`, `InvalidMail`, or
   `DuplicateSamAccountName`. A present-but-invalid or wrong-domain `mail` value is not replaced by `-DefaultDomain`.

`userPrincipalName`, `proxyAddresses`, and attribute-derived guesses are not used. `mail` is the authoritative AD attribute; the override
CSV is the authoritative source when present. This is the central rule that keeps the output explicit.

#### `Export-ADAccountMap`

```powershell
function Export-ADAccountMap {
  [CmdletBinding(SupportsShouldProcess = $true)]
  [OutputType('Winkit.ADDS.AccountMapResult')]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [psobject]
    $InputObject,

    [string]
    $OutputPath = 'ad-account-map.json',

    [string]
    $ReportPath,

    [switch]
    $Force,

    [switch]
    $DryRun,

    [switch]
    $PassThru
  )
}
```

Behavior:

- Accepts the `Winkit.ADDS.AccountMapResult` from `New-ADAccountMap`; a raw entry array is rejected with a clear error so there is one place
  that owns resolution policy.
- If `-DryRun` is set, `$WhatIfPreference = $true`.
- A result whose `Status` is `Failed` writes no map, but still writes `-ReportPath` when supplied so the failure is auditable; the function
  returns the Failed result. For a writable result, an existing target is refused unless `-Force` is set. The target must be a file path,
  not an existing directory.
- Builds the JSON from the resolved entries only. Unresolved and excluded entries never become keys.
- Writes to a uniquely named temporary file in the target directory, flushes and closes it, verifies the bytes, then
  `Move-Item -LiteralPath $temp -Destination $OutputPath -Force`. The final file is therefore never partially written; the temporary file is
  removed in `finally` on failure.
- `ShouldProcess` targets both `OutputPath` and `ReportPath`, so `-WhatIf` and `-Confirm` cover both artifacts.
- `-PassThru` returns the result with `Path`, `ReportPath`, `Sha256`, and final counts. The result is also emitted when `-WhatIf`/`-DryRun`
  is active, matching the preview behavior of `New-OutlookArchive.ps1`.

#### `ConvertTo-ADAccountMapJson`

```powershell
function ConvertTo-ADAccountMapJson {
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [System.Collections.IDictionary]
    $Map
  )
}
```

Serialization contract:

- One JSON object; keys are `samAccountName`, values are full email addresses.
- Keys sorted with `StringComparer.OrdinalIgnoreCase`; ties (keys differing only by case) are broken with `StringComparer.Ordinal` so the
  order is total and deterministic.
- Two-space indentation, `": "` between key and value, one entry per line, CRLF line endings, and one final CRLF. CRLF matches the
  repository's `*.json text eol=crlf` policy, so a checked-in golden file compares byte-for-byte.
- UTF-8 with no BOM when written by `Export-ADAccountMap`; the serializer returns a .NET string, and the writer uses
  `New-Object System.Text.UTF8Encoding($false)`.
- RFC 8259 escaping: `"` and `\` are escaped, control characters below U+0020 use fixed `\u00XX` sequences, and other characters are emitted
  as literal UTF-8. No timestamps, counts, host names, or paths appear in the file.
- An empty map serializes as `{}` plus the final CRLF.
- `ConvertTo-Json` is not used for the final file because Windows PowerShell 5.1 and PowerShell 7 differ in indentation, escaping, and
  key-order handling. A small purpose-built serializer is deterministic across both runtimes and is covered by the golden-file test.

### 3.3 Typed output objects

All records are `[pscustomobject]` instances with a custom type name, no classes, so they behave well in Pester and remain compatible with
Windows PowerShell 5.1.

`Winkit.ADDS.User`:

```text
PSTypeName        'Winkit.ADDS.User'
Resource          'Users'
DistinguishedName string
SamAccountName    string
UserPrincipalName string or $null
Mail              string or $null
Enabled           bool
ObjectGuid        guid or $null
Properties        ordered dictionary of requested AD attributes
```

`Winkit.ADDS.Computer`:

```text
PSTypeName        'Winkit.ADDS.Computer'
Resource          'Computers'
DistinguishedName string
SamAccountName    string
Name              string
DNSHostName       string or $null
Enabled           bool
ObjectGuid        guid or $null
Properties        ordered dictionary
```

`Winkit.ADDS.Group`:

```text
PSTypeName        'Winkit.ADDS.Group'
Resource          'Groups'
DistinguishedName string
SamAccountName    string
Name              string
GroupScope        string or $null
GroupCategory     string or $null
ObjectGuid        guid or $null
Properties        ordered dictionary
```

`Winkit.ADDS.ExportResult` (CSV writer):

```text
PSTypeName    'Winkit.ADDS.ExportResult'
Resource      'Users' | 'Computers' | 'Groups'
Path          string or $null
Count         int
Status        'Completed' | 'Skipped' | 'DryRun' | 'Failed'
Error         string or $null
```

`Export-AD.ps1 -PassThru` converts this record into a `New-OperationResult` object (`Target` = CSV path, `Source = 'ADDS'`,
`Action = 'Export'`, `Status`, and `Property = @{ Resource; Count }`) so the entry script follows the repository's structured-result
convention.

`Winkit.ADDS.AccountMapOverride` (override CSV row):

```text
PSTypeName     'Winkit.ADDS.AccountMapOverride'
SamAccountName string
Email          string or $null
Note           string or $null
RowNumber      int
Path           string
```

`Winkit.ADDS.AccountMapEntry`:

```text
PSTypeName        'Winkit.ADDS.AccountMapEntry'
SamAccountName    string
Email             string or $null
Source            'ADMail' | 'Override' | 'DefaultDomain' | 'None'
Status            'Resolved' | 'Unresolved' | 'Excluded' | 'Duplicate'
Reason            string or $null
DisplayName       string or $null
DistinguishedName string or $null
RawMail           string or $null
OverridePath      string or $null
```

`Winkit.ADDS.AccountMapResult`:

```text
PSTypeName      'Winkit.ADDS.AccountMapResult'
Status          'Completed' | 'DryRun' | 'Failed'
Selection       ordered dictionary (resource, filter/LDAP filter, search base, group, recursive, count)
Entries         Winkit.ADDS.AccountMapEntry[]
Unresolved      Winkit.ADDS.AccountMapEntry[]
Excluded        Winkit.ADDS.AccountMapEntry[]
Duplicates      Winkit.ADDS.AccountMapEntry[]
Map             ordered dictionary samAccountName -> email (resolved only, sorted)
SourceSummary   ordered dictionary (ADMail, Override, DefaultDomain, None)
EntryCount      int
ResolvedCount   int
UnresolvedCount int
Path            string or $null
ReportPath      string or $null
Sha256          string or $null
Errors          string[]
```

`-PassThru` on `Export-ADAccountMap` returns this object; the entry script does not need to invent a second result shape.

### 3.4 Entry scripts

#### `Export-AD.ps1` (backward-compatible)

The parameter block keeps the existing three parameters, their positions, defaults, and validate sets, and adds:

```powershell
[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Position = 0, Mandatory = $true)]
  [string]
  $OutputPath,

  [Parameter(Position = 1)]
  [ValidateSet('Users', 'Computers', 'Groups', 'All')]
  [string]
  $Resource = 'All',

  [Parameter(Position = 2)]
  [ValidateSet('UTF8', 'ASCII', 'Unicode', 'UTF7', 'UTF32', 'BigEndianUnicode')]
  [string]
  $Encoding = 'UTF8',

  [string]
  $Filter = '*',

  [string]
  $LDAPFilter,

  [string]
  $SearchBase,

  [string]
  $Server,

  [pscredential]
  $Credential,

  [string[]]
  $Properties,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)
```

Behavior:

- Dot-sources `lib/ADDS.ps1` after importing the substrate module, validates `OutputPath` as an existing directory, then loops over the
  requested resources. For each resource it calls `Get-ADDirectoryRecord` and pipes the records into `Export-ADDirectoryRecord`, so only one
  resource is buffered at a time instead of all three. A missing directory is a terminating error before any query; implicit creation is not
  added because it would widen the script's filesystem behavior.
- `-Resource All` still writes all three files in the same order. A missing resource is a warning plus a `Skipped` result, not a terminating
  error, preserving the "continue and exit 0" effect without writing to the error stream.
- `-PassThru` emits one `New-OperationResult` object per written file. Without it the script is quiet apart from `Write-Log` and warnings,
  matching current behavior.
- `-DryRun` sets `$WhatIfPreference = $true`; `-WhatIf` and `-Confirm` are available through `SupportsShouldProcess`.
- AD query failures terminate with exit code 1 and a `Write-Log -Color Red` message; the previous partial files are not deleted.
- The corrected `.NOTES` block documents Author, License, Server Core support (AD PowerShell module and network access required), and SYSTEM
  suitability (supported when the AD module and RSAT/AD Web Services access are available; no user profile needed).

Elevation is the one deliberate compatibility question: the current script requires it, but AD reads do not. The recommendation is to remove
`#Requires -RunAsAdministrator` in phase 2 as a visible, separately reviewed change, matching `Get-ADUserReport.ps1` and the other AD report
scripts. Until the owner approves, phase 1 keeps the directive.

#### `Export-ADAccountMap.ps1` (new)

```powershell
[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [string]
  $OutputPath = 'ad-account-map.json',

  [string]
  $ReportPath,

  [string]
  $SearchBase,

  [string]
  $Filter = '*',

  [string]
  $LDAPFilter,

  [string]
  $Group,

  [switch]
  $Recursive,

  [string[]]
  $Identity,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string[]]
  $AllowedDomain,

  [string]
  $OverridePath,

  [string]
  $DefaultDomain,

  [ValidateSet('Fail', 'Omit')]
  [string]
  $OnMissingMail = 'Fail',

  [ValidateSet('Fail', 'First', 'Last')]
  [string]
  $OnDuplicate = 'Fail',

  [switch]
  $AllowOverrideOutsideDomain,

  [switch]
  $IncludeUnmatchedOverride,

  [switch]
  $EnabledOnly,

  [string]
  $Server,

  [pscredential]
  $Credential,

  [switch]
  $Force,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)
```

Flow:

1. Import substrate, dot-source the library, validate the output and report paths (fail fast before querying AD).
2. Load overrides with `Import-ADAccountMapOverride` when `-OverridePath` is supplied.
3. Select exactly one parameter set and call `New-ADAccountMap` with a purpose-built splat. Never pass `-Filter` when `-Identity` or
   `-Group` is used, and reject `-SearchBase` with those selections instead of silently ignoring it.
4. Log the source summary, unresolved list count, excluded count, and duplicate count.
5. Call `Export-ADAccountMap` once. It enforces overwrite protection, `ShouldProcess`, the report-only path, atomic write, and deterministic
   bytes.
6. If the returned result is `Failed`, log the reasons, emit the result when `-PassThru` or `-DryRun` is used, and exit 1 without writing
   the map.
7. Otherwise log the output path, entry count, and SHA-256. Emit the result when `-PassThru` or `-DryRun` is used.
8. Exit 0 on success, including `-DryRun`; exit 1 on failure or unresolved entries under the `Fail` policy.

Examples for the user guide:

```powershell
# All users with an explicit allowlist and a curated override file.
.\scripts\Server\ADDS\Export-ADAccountMap.ps1 `
  -AllowedDomain delta4x4.com, delta4x4.de `
  -OverridePath C:\ADReports\account-map-overrides.csv `
  -OutputPath C:\ADReports\ad-account-map.json `
  -ReportPath C:\ADReports\ad-account-map-audit.csv `
  -DryRun
```

```powershell
# Only one OU, only enabled users, and replace the previous file.
.\scripts\Server\ADDS\Export-ADAccountMap.ps1 `
  -SearchBase 'OU=Users,OU=Delta4x4,DC=delta4x4,DC=local' `
  -Filter 'Enabled -eq $true' `
  -AllowedDomain delta4x4.com, delta4x4.de `
  -OutputPath C:\ADReports\ad-account-map.json `
  -Force
```

```powershell
# Reuse an existing export instead of querying AD again.
Get-ADDirectoryRecord -Resource Users -Filter * |
  New-ADAccountMap -AllowedDomain delta4x4.com, delta4x4.de |
  Export-ADAccountMap -OutputPath C:\ADReports\ad-account-map.json -Force -PassThru
```

### 3.5 Mapping of current behavior

| Current behavior                                    | New implementation                                                             |
| --------------------------------------------------- | ------------------------------------------------------------------------------ |
| `-OutputPath` mandatory, position 0, directory      | Same parameter and semantics                                                   |
| `-Resource` default `All`                           | Same ValidateSet and default                                                   |
| `-Encoding` ValidateSet and `Export-Csv` semantics  | Same values, same `Export-Csv` call, same delimiter `;`                        |
| `AD-Users.csv`, `AD-Computers.csv`, `AD-Groups.csv` | Same names, generated with invariant title casing                              |
| `Get-ADUser -Filter *` and per-resource properties  | `Get-ADDirectoryRecord -Resource Users -Filter *` with the same property table |
| All three results held in memory before writing     | One resource queried and written at a time                                     |
| Empty resource: `Write-Error`, no file, exit 0      | Warning, `Skipped` result, no file, exit 0                                     |
| Existing CSVs overwritten silently                  | Preserved in phase 1; overwrite guard is an owner decision                     |
| No preview                                          | `-DryRun`, `-WhatIf`, `-Confirm`                                               |
| No structured output                                | `-PassThru` returns `New-OperationResult` records                              |
| Help mentions a nonexistent `Filter`                | Real `-Filter`, `-LDAPFilter`, `-SearchBase`, `-Server`, `-Credential`         |
| `#Requires -Version 2.0`                            | `#Requires -Version 5.1`                                                       |
| No substrate import                                 | Pinned substrate import plus `Write-Log`                                       |

CSV columns: the current output is module-defined. The plan keeps the same logical attributes but fixes their order to the resource's
default columns followed by `Get-ADResourceProperty` order, duplicates removed. If byte-identical column order is a hard consumer
requirement, phase 1 must preserve the current order by flattening the original AD object instead; section 8, question 2 asks the owner to
choose.

## 4. AD account map feature

### 4.1 Output contract

The generated file is exactly the object map the consumer expects, with no wrapper or metadata:

```json
{
  "CFischer": "christopher.fischer@delta4x4.com",
  "LHippmann": "l.hippmann@delta4x4.com",
  "RRudorfer": "rene.rudorfer@delta4x4.de"
}
```

- Keys are AD `sAMAccountName` values; values are full email addresses.
- [`plead`'s inbox schema v1](https://github.com/adnoctem/plead/blob/main/docs/USAGE.md#the-inbox-absence-files) gives `email` precedence
  over `samAccountName` + `domain`, so consuming scripts can write the mapped address directly as the absence entry's `email` while keeping
  `samAccountName` and `domain` for traceability.
- The file is not an absence file and carries no `schemaVersion`, dates, or template data. Audit metadata belongs in the optional report or
  the `-PassThru` result, never in the map.

### 4.2 Input selection

One of the following parameter sets selects users:

| Selection       | Parameters                                              | Notes                                              |
| --------------- | ------------------------------------------------------- | -------------------------------------------------- |
| Whole domain    | `-Filter '*'` (default)                                 | Same default as today's `Get-ADUser -Filter *`     |
| Filtered OU     | `-SearchBase <DN>` plus `-Filter` or `-LDAPFilter`      | OU/container scoping without a full-domain query   |
| Group           | `-Group <identity>` and optional `-Recursive`           | Only users are resolved for the map                |
| Explicit users  | `-Identity <sAMAccountName/DN/UPN>` or pipeline records | Useful for testing and targeted regeneration       |
| Existing export | `Get-ADDirectoryRecord ... \| New-ADAccountMap ...`     | No second AD query; typed records carry the fields |

`-EnabledOnly` restricts the selection to enabled accounts. It is off by default, matching the current export's unfiltered query; the owner
can flip the default without changing the JSON format.

### 4.3 Authoritative attribute

`mail` is the authoritative AD attribute. It is the only automatic source. The explicit override CSV is the authoritative source when a row
exists. The resolution order is:

```text
override CSV > mail attribute (valid and allowed) > explicit DefaultDomain (only when mail is absent) > unresolved
```

`userPrincipalName`, `proxyAddresses`, `targetAddress`, and `<samAccountName>@<domain>` guesses are not consulted. This is deliberate: the
production domain has an unreliable `mail` attribute and users span `delta4x4.com` and `delta4x4.de`, so any automatic fallback would create
plausible but wrong addresses. An operator who wants a fallback must name it with `-DefaultDomain`, and every fallback is counted and
reported. When `mail` exists but is invalid or outside the allowlist, the entry is unresolved; `-DefaultDomain` does not mask a value that
looks intentional but is wrong.

### 4.4 Missing mail and wrong domain

| Condition                                        | Result                                                                              |
| ------------------------------------------------ | ----------------------------------------------------------------------------------- |
| `mail` missing or empty                          | `Status = Unresolved`, `Reason = MissingMail`                                       |
| `mail` is not a syntactically valid bare address | `Status = Unresolved`, `Reason = InvalidMail`                                       |
| `mail` domain is not in `-AllowedDomain`         | `Status = Unresolved`, `Reason = MailDomainNotAllowed`                              |
| Override exists                                  | Override wins; the AD value is still recorded as `RawMail` for the audit report     |
| Override domain outside allowlist and no opt-in  | `Status = Unresolved`, `Reason = OverrideDomainNotAllowed`                          |
| Override has empty email                         | `Status = Excluded`, not an error; the account is intentionally kept out of the map |

Policy:

- `-OnMissingMail Fail` (default): any unresolved entry sets `Result.Status = 'Failed'`; no map file is written, the optional report is
  still written so the operator can see every account that needs attention, and the script exits 1.
- `-OnMissingMail Omit`: resolved entries are written; unresolved entries are omitted from the map but remain in `Result.Unresolved` and the
  report. The script logs a warning and exits 0. This is the explicit "partial map" mode.
- Wrong-domain values are never silently rewritten, lower-cased, or mapped to another domain.

### 4.5 Overrides and manual mapping

The override CSV is the mechanism for the unreliable `mail` attribute. It is explicit, reviewable in Git or a change ticket, and its rows
are visible in the audit report. Example:

```csv
samAccountName,email,note
CFischer,christopher.fischer@delta4x4.com,mail attribute is stale
LHippmann,l.hippmann@delta4x4.com,
RRudorfer,rene.rudorfer@delta4x4.de,second mail domain
ServiceAccount,,intentionally excluded
```

Rules:

- Headers are case-insensitive; `samAccountName` and `email` are required, `note` is optional.
- Duplicate keys after trimming and case-folding terminate with an error.
- Empty email means intentional exclusion, not missing data.
- An override for a `samAccountName` not returned by the selection is reported as `UnmatchedOverride`. By default it is not emitted, because
  a typo or a stale leaver should not silently enter the map; `-IncludeUnmatchedOverride` opts in.
- Overrides do not need AD `mail`, so a user with no `mail` can still be mapped.
- If an override is outside `-AllowedDomain`, the default is an unresolved error; `-AllowOverrideOutsideDomain` is the audited escape hatch.

### 4.6 Duplicate samAccountName handling

AD enforces `sAMAccountName` uniqueness within a domain, but a global catalog query, multiple-domain selection, or an override file can
still produce duplicates. Handling:

- Keys are compared with ordinal case-insensitive comparison after trimming.
- Duplicate override rows always terminate with an error naming both rows; the file cannot be interpreted.
- Duplicate AD records default to `-OnDuplicate Fail`, because silently picking one account is unsafe.
- `First`/`Last` sort the duplicate candidates by `DistinguishedName` with ordinal comparison before choosing, so the result does not depend
  on AD enumeration order.
- The duplicate candidate that is not chosen is retained in `Result.Duplicates` and the report.
- The serializer independently rejects duplicate dictionary keys before producing bytes; this is a last-line defense against a JSON object
  that a parser would silently collapse.

### 4.7 Address validation

`Test-ADMailAddress` implements the rules in section 3.2. In addition:

- The final value must be the bare address, not `Name <address>`.
- The domain allowlist is compared case-insensitively and must match exactly; `delta4x4.com` does not automatically permit
  `mail.delta4x4.com`.
- Validation happens before the map is built, so an invalid value is reported against its `samAccountName` rather than failing later during
  serialization.
- The report records the raw `mail` value next to the final value, so a wrong-but-valid address can still be audited.

### 4.8 Deterministic output

Determinism is a testable contract:

- Keys sorted with ordinal-ignore-case, ties broken with ordinal comparison.
- Two-space indentation, one entry per line, CRLF line endings, one final CRLF.
- UTF-8 without BOM. The writer creates `System.Text.UTF8Encoding($false)` explicitly; it never relies on the host default.
- Fixed escaping: `"` and `\` escaped, control characters below U+0020 as fixed `\u00XX`, all other characters as literal UTF-8.
- No timestamps, machine names, counts, paths, or `Get-Date` output in the map.
- The same selected entries produce the same bytes regardless of input order, PowerShell version, or host culture.
- `Export-ADAccountMap` writes to a temporary file in the destination directory and moves it into place; a failed run never leaves a
  truncated final file, and a failed validation never touches an existing file.

### 4.9 Safety

- `[CmdletBinding(SupportsShouldProcess = $true)]` and `-DryRun` (which sets `$WhatIfPreference`) on both entry scripts and the export
  function.
- Overwrite protection: `Export-ADAccountMap` refuses an existing target unless `-Force`; `-WhatIf` writes nothing at all.
- Validation before the AD query where possible: output path, report path, override file, and allowed-domain syntax.
- `-Confirm` covers both the map and the report through `ShouldProcess`.
- No plaintext credentials: `-Credential` is `[pscredential]`; no password parameter is added.
- The map contains addresses, which are personal data. The docs recommend an ACL-restricted output directory outside the managed winkit
  installation. Logs contain counts and account names, not full address lists; the optional report contains the detail for authorized
  operators.
- Exit codes: 0 on success or preview; 1 on query failure, failed policy, or write failure.
- No `Invoke-Expression`, no string-built commands, and no unverified downloads.

### 4.10 Logging and auditability

`Write-Log -Message -Color` phases:

1. Selection summary: resource, filter/LDAP filter, search base, group, recursive flag, identity count, server.
2. Override load: path, total rows, excluded rows, duplicate errors.
3. Resolution summary: selected users, resolved by source (`ADMail`, `Override`, `DefaultDomain`), unresolved, excluded, duplicates,
   unmatched overrides.
4. Write summary: output path, report path, entry count, SHA-256.

Warnings use stable reasons from section 4.4 so operators can grep them. `-Verbose` adds per-account decisions without flooding the normal
log.

The optional `-ReportPath` CSV has one row per selected account plus each unmatched override:

| Column              | Meaning                                           |
| ------------------- | ------------------------------------------------- |
| `samAccountName`    | Account key                                       |
| `email`             | Final address or empty                            |
| `source`            | `ADMail`, `Override`, `DefaultDomain`, or `None`  |
| `status`            | `Resolved`, `Unresolved`, `Excluded`, `Duplicate` |
| `reason`            | Stable reason or empty                            |
| `rawMail`           | Original `mail` attribute value, if any           |
| `distinguishedName` | Account DN, if any                                |
| `note`              | Override note                                     |

The report is written even when the map is not, so a failed run still leaves an actionable audit artifact. It uses UTF-8 without BOM and a
`;` delimiter to match the CSV convention already used by Export-AD.

## 5. Testing strategy

### 5.1 Characterization first

Before changing behavior, add Pester tests that pin today's contract:

- `-OutputPath`, `-Resource`, and `-Encoding` parameter names, positions, defaults, and validate sets.
- `AD-Users.csv`, `AD-Computers.csv`, and `AD-Groups.csv` names, `;` delimiter, and the overwrite behavior.
- `-Resource Users` writes only the user file; `All` writes three; an empty resource writes no file.
- Incorrect help assumptions are corrected in the same change, so the tests assert the documented behavior, not the old prose.

These tests provide the safety net for the library extraction.

### 5.2 Unit tests with mocked AD

New files under `tests/Server/ADDS/`:

- `ADDS.Tests.ps1` - library functions.
- `ExportAD.Tests.ps1` - entry script and CSV writer.
- `ExportADAccountMap.Tests.ps1` - mapping, overrides, JSON, and safety. Tests must run without a domain, Outlook, or the `delta4x4`
  environment. Because Pester cannot mock a command that does not exist, `BeforeAll` defines lightweight stubs for `Get-ADUser`,
  `Get-ADComputer`, `Get-ADGroup`, and `Get-ADGroupMember` when the real ActiveDirectory module is absent, then mocks them per test. The
  tests dot-source `lib/ADDS.ps1` the same way the entry scripts do. Entry-script tests run the script through a generated harness with the
  `#Requires` lines removed (the pattern already used by `tests/Administration/UserProfileData.Tests.ps1`) or skip the test when the
  ActiveDirectory module is absent; the library tests never require the module.

Coverage matrix:

| Area                          | Cases                                                                                                                                                                                                    |
| ----------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Get-ADDirectoryRecord`       | Parameter sets resolve; `All` rejects `Identity`/`Group`; typed `PSTypeName`s; property bag order; pipeline identity; group recursion; query failure is terminating; `-Server`/`-Credential` passthrough |
| `Export-ADDirectoryRecord`    | File names; fixed column order; `;` delimiter; UTF8/ASCII encodings; empty input; `-WhatIf` writes nothing; `-PassThru` result                                                                           |
| `Test-ADMailAddress`          | Valid bare addresses; display names rejected; whitespace/CRLF rejected; allowlist exact match; case-insensitive host; overlong address                                                                   |
| `Import-ADAccountMapOverride` | BOM and no-BOM input; header case; missing header; duplicate key; empty email exclusion; malformed address                                                                                               |
| `New-ADAccountMap`            | Override precedence; missing mail; wrong domain; `DefaultDomain`; invalid mail; duplicate policies; unmatched overrides; disabled users; `OnMissingMail` semantics; sort stability for `First`/`Last`    |
| `ConvertTo-ADAccountMapJson`  | Empty map; escaping quote/backslash/control characters; non-ASCII literal UTF-8; key sorting with case variants; duplicate key rejection                                                                 |
| `Export-ADAccountMap`         | Overwrite refusal; `-Force`; `-WhatIf`; temp file cleanup; result object; report write; failed result refuses to write                                                                                   |
| `Export-AD.ps1`               | Backward-compatible surface; one resource at a time; preview; PassThru; error exit code                                                                                                                  |
| `Export-ADAccountMap.ps1`     | Selection parameter sets; exit codes for `Fail`/`Omit`; logging; `-DryRun`; report written on failure                                                                                                    |

### 5.3 Golden-file and determinism tests

- A checked-in fixture `tests/Server/ADDS/fixtures/ad-account-map.golden.json` contains the expected bytes for a fixed data set, including
  two mail domains and a case-variant key.
- The test writes the map to `TestDrive`, reads `[System.IO.File]::ReadAllBytes()`, asserts the first byte is `{` (0x7B) and the first three
  bytes are not `EF BB BF`, then compares the full byte array with the fixture.
- The test runs the export twice into two files and compares the hashes; both must be identical.
- Input entries are shuffled between runs; the bytes must not change.
- A culture test temporarily sets `de-DE` and `tr-TR` (restoring the previous culture in `finally`) and verifies identical bytes and file
  names.
- A failure-path test asserts that an existing valid output survives a failed run unchanged and that no `.tmp` file remains.

Because `.gitattributes` declares `*.json text eol=crlf`, the fixture and the generated CRLF bytes agree on checkout in Windows, Linux, and
CI. If the owner prefers LF output, the fixture must use a non-`.json` extension or an explicit `.gitattributes` exception.

### 5.4 Static analysis, help, and docs checks

- `.\winkit.ps1 format -Check` and `.\winkit.ps1 lint` must pass for the library, both entry scripts, and all tests.
- `Invoke-ScriptAnalyzer` covers approved verbs, singular nouns, PowerShell 5.0-compatible syntax, and the repository's formatting rules.
  The new function names avoid plural nouns, so no allowlist change is expected.
- Pester convention tests assert `#Requires` lines, pinned substrate version, `.SYNOPSIS`/`.NOTES`, `SupportsShouldProcess`, and `Get-Help`
  parameter names for both entry scripts.
- `.\winkit.ps1 test` runs the whole suite; the new tests need no Outlook and are part of the default logic suite.

## 6. Documentation, help, and release impact

### 6.1 Comment-based help

Every library function gets `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`, and, where relevant, `.NOTES`. Both entry scripts get
corrected help:

- `.PARAMETER OutputPath` (not `OutputFile`) and a description that matches the directory-versus-file difference between the two scripts.
- Examples that actually run, including the mapping examples from section 3.4.
- `.NOTES` with Author, License, Server Core support, and SYSTEM-account suitability. For AD reads: Server Core is supported when the
  ActiveDirectory PowerShell module and AD Web Services connectivity are available; SYSTEM execution is supported when the module and
  network/AD permissions are available, with no user profile required.

### 6.2 Repository docs

- Add `scripts/Server/README.md` (the area currently has no catalogue) with the script table, requirements, and links.
- Add `docs/www/Server/README.md` and a topic guide such as `docs/www/Server/Active-Directory-Exports.md` covering:
  - installed working directory assumptions and writable output locations outside the managed installation;
  - `Export-AD.ps1` parameters and examples;
  - `Export-ADAccountMap.ps1` selection, allowlist, override CSV schema, `Fail` versus `Omit`, report, and preview;
  - the `plead` consumer contract: the map is `samAccountName` -> email, `email` wins in the inbox file, disabled/excluded accounts are
    absent, and the map is UTF-8 without BOM with stable key order;
  - the fact that `mail` is authoritative and no address is guessed.
- Update `docs/CONTRIBUTING.md` or `docs/AGENTS.md` only if the owner accepts the `lib/` dot-source convention; the plan itself does not
  require a contributor-rule change.
- Do not put development history, test-run reports, or fixture-generation instructions in user docs; those stay with tests or this plan.

### 6.3 Release and versioning

- New script plus new parameters is a `feat(scripts): ...` change and releases a minor version under semantic-release.
- Refactoring `Export-AD.ps1` without changing its CSV contract is `refactor(scripts): ...`; if the CSV column order or overwrite behavior
  changes in a way existing consumers can observe, mark it as a breaking change and coordinate a major.
- No `requirements.psd1` change is needed for the local library. A future promotion of the library to `AdNoctem.Substrate.PowerShell` is a
  coordinated change: bump `requirements.psd1`, every `#Requires`/`Import-Module` pin, tests, and docs together, per `CONTRIBUTING.md`.
- The installer already copies `scripts/` recursively and its allowed roots include `scripts`, so `scripts/Server/ADDS/lib` needs no
  installer change. Verify with the existing installer tests and a `build` archive inspection.
- `bin/` launchers: add one only if the owner wants a simple mapped launcher for the new entry script; not required by this plan.

## 7. Phased implementation plan

Each phase is a separate reviewable change with its own tests and docs. No phase starts before the preceding phase's verification passes.

### Phase 0 - Owner review

Resolve the questions in section 8. Record the decisions in this document or the pull request.

### Phase 1 - Extract the library and pin current behavior

- Add `scripts/Server/ADDS/lib/ADDS.ps1` with `Get-ADResourceProperty`, `Get-ADDirectoryRecord` (filter/identity paths only),
  `New-ADDirectoryRecord` (internal), and `Export-ADDirectoryRecord`.
- Rewrite `Export-AD.ps1` as a thin wrapper with the same three public parameters and the same CSV files.
- Add characterization tests and library unit tests with mocked AD.
- Keep `#Requires -RunAsAdministrator` in this phase.
- Verify `format -Check`, `lint`, and `test`.

### Phase 2 - Expand the export surface

- Add `-Filter`, `-LDAPFilter`, `-SearchBase`, `-Server`, `-Credential`, `-Properties`, `-DryRun`, `-PassThru`, and `ShouldProcess` to
  `Export-AD.ps1`.
- Add group selection and the `All` validation to `Get-ADDirectoryRecord`.
- Decide and document the elevation change; if approved, remove `#Requires -RunAsAdministrator` and update the help.
- Test parameter sets, pipeline identity, preview, and failure exit codes.

### Phase 3 - Mapping resolution

- Add `Test-ADMailAddress`, `Import-ADAccountMapOverride`, and `New-ADAccountMap`.
- Implement selection, precedence, unresolved reasons, duplicate policy, default domain, unmatched overrides, and typed results.
- Unit tests cover the full resolution matrix; no file is written in this phase.

### Phase 4 - Deterministic JSON and the entry script

- Add `ConvertTo-ADAccountMapJson` and `Export-ADAccountMap`.
- Add `Export-ADAccountMap.ps1`, overwrite protection, atomic writes, report output, logging, and exit codes.
- Add the golden fixture, BOM/byte tests, culture tests, idempotence tests, and failure-path tests.
- Verify a generated file against `plead`'s documented inbox expectations in a scratch directory (no Plesk call).

### Phase 5 - Documentation and release

- Update script help, `scripts/Server/README.md`, `docs/www/Server/README.md`, and the AD export/account-map guide.
- Add usage examples for overrides, allowlists, and the two mail domains.
- Run the full suite, lint, format check, and a build archive inspection.
- Commit with conventional messages; release as a minor unless a breaking decision from section 8 requires otherwise.

### Phase 6 - Optional follow-up

- Evaluate promoting the generic AD primitives to `AdNoctem.Substrate.PowerShell`, with its own version pin and release.
- Add a `bin/` launcher or a `Get-ADDirectoryRecord.ps1` wrapper only if users need the acquisition API directly from the installed tree
  without dot-sourcing the library.

## 8. Open questions for the owner

1. **Library location:** approve `scripts/Server/ADDS/lib/ADDS.ps1` as a dot-sourced, in-repo function library, or should the reusable
   functions go into `AdNoctem.Substrate.PowerShell` (separate release and version-pin bump)? The recommendation is the local library for
   this change, with Substrate promotion as a follow-up.
2. **CSV contract:** must `Export-AD.ps1` produce byte-identical column order to today's module-defined output, or is a documented, stable
   column order acceptable? A stable order is recommended; it is the only way to test the schema.
3. **Domain matching:** is an exact allowlist (`delta4x4.com`, `delta4x4.de`) sufficient, or must subdomains and IDN forms be accepted? The
   recommendation is exact matching, with every accepted domain listed explicitly.
4. **Override format:** is CSV sufficient for v1, or is JSON override input required so the emitted map can be edited and fed back? The
   recommendation is CSV only, to keep duplicate-key detection reliable on Windows PowerShell 5.1.
5. **Default domain:** should `-DefaultDomain` exist at all? It is opt-in, audited, and disabled unless supplied; but the safest reading of
   "no guessing" is to require an override for every user without `mail`.
6. **Disabled accounts:** should the map include disabled users by default (current export behavior) or default to enabled users only? The
   design currently uses an opt-in `-EnabledOnly` switch; if the default flips, the switch should become `-IncludeDisabled`. The
   recommendation is enabled-only for the map, while keeping `Export-AD.ps1` unfiltered.
7. **Unmatched overrides:** default to report-and-exclude (recommended) or include them? Including stale or typo entries risks wrong
   mappings; excluding them can drop a deliberate manual entry if the AD query is scoped too narrowly.
8. **Override domain escape hatch:** should `-AllowOverrideOutsideDomain` exist, or must every override also be inside the allowlist? The
   recommendation is to keep the switch, require it explicitly, and report every use.
9. **Missing-mail policy:** is `Fail` the right default, or should the standard behavior be `Omit` with a warning because the consumer can
   tolerate missing accounts? The recommendation is `Fail`, because a silent omission is the failure mode this feature exists to prevent.
10. **Output path:** keep `-OutputPath` mandatory with no default for consistency with `Export-AD.ps1`, or default to `ad-account-map.json`
    in the current directory? The recommendation is the default name, with `-Force` required to overwrite.
11. **Report artifact:** is `-ReportPath` wanted as a first-class deliverable, and should it be CSV or JSON? The recommendation is an
    optional CSV because operators already consume CSV exports.
12. **Elevation:** remove `#Requires -RunAsAdministrator` from `Export-AD.ps1` (AD reads do not need it and sibling report scripts omit it),
    or keep it for historical compatibility?
13. **Key casing:** preserve AD `sAMAccountName` casing in the JSON keys (recommended) or normalize to lower case? Consumers should treat
    keys case-insensitively, but the file should be stable and traceable.
14. **Consumer verification:** should phase 4 include a checked-in sample map and a test that parses it with the `plead` schema rules, or is
    a byte-level golden file enough? The recommendation is both: byte determinism plus a schema-shape assertion.

## 9. Owner decisions and implementation record

The owner approved the structure and resolved the open questions as follows:

1. Keep the reusable functions in `scripts/Server/ADDS/lib/ADDS.ps1`; do not move anything into Substrate.
2. Adopt a new documented, stable CSV column order without byte compatibility to 3.0.0.
3. Use exact, ASCII/Punycode domain allowlist matching; reject Unicode domain spellings.
4. Support both CSV and JSON export through `Export-AD.ps1 -Format Csv|Json`.
5. Drop `-DefaultDomain`; the allowlist is the only domain source.
6. Default to enabled accounts and add `-IncludeDisabled`.
7. Report and exclude unmatched overrides by default.
8. Keep `-AllowOverrideOutsideDomain` as an explicit, audited switch.
9. Missing `mail` fails by default; `-IncludeMissing` writes the resolved subset and reports omissions. `Export-AD.ps1 -Omit` removes
   properties from the export and from the missing-value check.
10. Default the account-map output to `ad-account-map.json` in the current directory.
11. Provide an optional CSV audit report.
12. Remove `#Requires -RunAsAdministrator`; classify the change as a feature, not a breaking change.
13. Preserve the AD `sAMAccountName` casing in JSON keys.
14. Pause real-AD consumer/golden tests until the owner supplies the CSV fixture documented in `tests/Server/ADDS/fixtures/README.md`;
    synthetic golden and mocked-AD tests run now.

Implementation notes:

- `New-ADAccountMap` does not expose `-Omit`; the account map requires `sAMAccountName` and `mail`, so omitting either would defeat the
  mapping. `-Omit` belongs to the general `Export-AD.ps1` export surface.
- The synthetic golden fixture is `tests/Server/ADDS/fixtures/ad-account-map.synthetic.json`. The skipped real-AD test and the fixture
  contract live in `tests/Server/ADDS/ExportADAccountMap.Tests.ps1` and `tests/Server/ADDS/fixtures/README.md`.
- The library is dot-sourced by both entry scripts. It contains no elevation requirement and no main workflow.
