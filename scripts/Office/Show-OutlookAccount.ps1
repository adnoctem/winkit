#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.2.0' }

<#
.SYNOPSIS
  Reports Outlook accounts, their delivery stores, and data file headroom.
.DESCRIPTION
  Answers the question that decides whether an account can be removed safely:
  which store does each account deliver new mail into, and what else depends on
  that store. Lists every account in the current Outlook profile with its type,
  address, and delivery store, then every data file with its format, Exchange
  backing, size, and remaining headroom against the configured limit.

  Nothing is created, moved, copied, renamed, attached, or detached, and no
  folder is opened. Only account and store properties are read, so this runs on
  Outlook 2007 against Exchange and OST stores, where standard-folder discovery
  is unavailable and Show-OutlookState.ps1 cannot survey the store.

  Findings name the configurations that make removal destructive. The one that
  matters most is an account that delivers into a store it does not own: a POP
  or IMAP account delivering into an Exchange mailbox keeps its only copy of
  that mail inside the OST, and removing the Exchange account discards it.

  Store classification comes from the owning accounts, not the file extension.
  An Outlook 2007 IMAP account is backed by a PST that reports olNotExchange,
  so extension alone would read it as an ordinary local data file while its
  items can still be header-only stubs waiting on the server.

  Before Outlook 2013, an IMAP account saves sent mail and drafts into the
  profile's delivery store instead of the server's folders. That is reported,
  because it means a shared mailbox's sent mail is scattered across the
  workstations that carry it and is not archived with the server-side folders.

  A JSON report is written only when ReportPath or ReportDirectory is supplied.
  Collect those reports across a fleet to plan removals from one place.
.PARAMETER ReportDirectory
  Directory for a generated JSON report filename. Created when absent.
.PARAMETER ReportPath
  Exact JSON report path. Refused when the file already exists.
.PARAMETER IgnoreAdministrator
  Permit execution from an elevated session. Outlook must run as the same
  Windows user at the same elevation. This override does not switch users or
  profiles.
.PARAMETER PassThru
  Emit the inventory as an object in addition to printing it.
.EXAMPLE
  PS> .\Show-OutlookAccount.ps1
.EXAMPLE
  PS> .\Show-OutlookAccount.ps1 -ReportDirectory 'C:\Managed\Reports'
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not applicable - Outlook is a desktop client.
  SYSTEM-account execution: not applicable - requires an interactive Outlook profile.
  Outlook version: 2007 or later, including Exchange and OST profiles.
  Bitness: Outlook COM automation supports cross-architecture PowerShell clients.
#>

[CmdletBinding()]
[OutputType([PSCustomObject])]
param (
  [string]
  $ReportDirectory,

  [ValidateNotNullOrEmpty()]
  [string]
  $ReportPath,

  [switch]
  $IgnoreAdministrator,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.2.0 -Force

# Outlook automation must use the mailbox user's interactive security context.
$_outlookUser = Get-UserInfo
if ($_outlookUser.IsAdministrator) {
  if (-not $IgnoreAdministrator) {
    throw "PowerShell is elevated as '$($_outlookUser.UserName)'. Run this script from a non-elevated PowerShell window as the Windows user who runs Outlook. Use -IgnoreAdministrator only when Outlook intentionally runs elevated under that same user."
  }

  Write-Warning "IgnoreAdministrator permits elevated execution as '$($_outlookUser.UserName)'. Outlook must run under the same Windows user and elevation. This override does not switch users or profiles."
}

# -----------------------------------------------------------------------------

# Outlook 2007 Unicode defaults, in megabytes. Both are overridable per user
# and by policy, so the configured values are read before these are used.
$_defaultMaximumFileSizeMb = 20480
$_defaultWarningFileSizeMb = 19456

# Outlook 15 (2013) is the first release that stores IMAP sent mail and drafts
# in the server's folders instead of the local delivery store.
$_serverSentItemsMajor = 15

# Outlook 14 (2010) is the first release whose object model exposes
# Account.DeliveryStore. Below it the per-account delivery store cannot be read
# at all, so an unresolved delivery reports a limit of the client rather than a
# misconfigured profile.
$_deliveryStoreMajor = 14

$_startedAt = Get-Date

if ($PSBoundParameters.ContainsKey('ReportPath') -and $PSBoundParameters.ContainsKey('ReportDirectory')) {
  throw 'Use ReportPath or ReportDirectory, not both.'
}

$_reportPath = $null

# Reserve the report before opening Outlook so an unwritable destination stops
# the inventory early. No report is written unless one was explicitly requested.
if ($PSBoundParameters.ContainsKey('ReportPath')) {
  if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    throw 'ReportPath must specify a report filename.'
  }

  $_reportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
  $_reportDirectory = Split-Path -Path $_reportPath -Parent
}
elseif ($PSBoundParameters.ContainsKey('ReportDirectory')) {
  if ([string]::IsNullOrWhiteSpace($ReportDirectory)) {
    throw 'ReportDirectory must specify a directory.'
  }

  $_reportDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportDirectory)
  $_reportName = 'Show-OutlookAccount-{0}-{1}.json' -f $_startedAt.ToString('yyyyMMdd-HHmmss-fff'), [guid]::NewGuid().ToString('N')
  $_reportPath = Join-Path $_reportDirectory $_reportName
}

if ($_reportPath) {
  if (Test-Path -LiteralPath $_reportPath) {
    throw "Report path already exists: $_reportPath. Choose a new report filename."
  }

  $null = [IO.Directory]::CreateDirectory($_reportDirectory)
}

# -----------------------------------------------------------------------------

function Format-OutlookAccountSize {
  param (
    [Nullable[long]]
    $Bytes
  )

  if ($null -eq $Bytes) {
    return 'unknown'
  }

  if ($Bytes -ge 1GB) {
    return '{0:N1} GB' -f ($Bytes / 1GB)
  }

  if ($Bytes -ge 1MB) {
    return '{0:N1} MB' -f ($Bytes / 1MB)
  }

  return '{0:N0} KB' -f ($Bytes / 1KB)
}

function Get-OutlookAccountFileLimit {
  param (
    [string]
    $OutlookVersion
  )

  # Policy wins over the per-user preference; both are megabyte DWORDs. An
  # unreadable key leaves the documented default in place rather than failing.
  $_maximum = $_defaultMaximumFileSizeMb
  $_warning = $_defaultWarningFileSizeMb
  $_source = 'Default'

  $_paths = @(
    "HKCU:\Software\Microsoft\Office\$OutlookVersion\Outlook\PST",
    "HKCU:\Software\Policies\Microsoft\Office\$OutlookVersion\Outlook\PST"
  )

  foreach ($_path in $_paths) {
    $_configuredMaximum = Get-RegistryValue -Path $_path -Name 'MaxLargeFileSize' -ErrorAction SilentlyContinue
    $_configuredWarning = Get-RegistryValue -Path $_path -Name 'WarnLargeFileSize' -ErrorAction SilentlyContinue

    if ($null -ne $_configuredMaximum) {
      $_maximum = [int]$_configuredMaximum
      $_source = $_path
    }

    if ($null -ne $_configuredWarning) {
      $_warning = [int]$_configuredWarning
      $_source = $_path
    }
  }

  [PSCustomObject]@{
    MaximumMb = $_maximum
    WarningMb = $_warning
    Source    = $_source
  }
}

function Get-OutlookAccountStoreRecord {
  param (
    [object]
    $Store,

    [object]
    $FileLimit
  )

  # The configured limit is a registry fact, not a store property, so headroom
  # is derived here rather than read from the profile.
  $_headroomMb = $null
  $_overWarning = $false
  if ($null -ne $Store.SizeBytes) {
    $_headroomMb = [int]($FileLimit.MaximumMb - ($Store.SizeBytes / 1MB))
    $_overWarning = ($Store.SizeBytes / 1MB) -ge $FileLimit.WarningMb
  }

  [PSCustomObject]@{
    DisplayName               = [string]$Store.DisplayName
    StoreID                   = [string]$Store.StoreID
    FilePath                  = [string]$Store.FilePath
    DataFileFormat            = [string]$Store.DataFileFormat
    ExchangeStoreType         = [int]$Store.ExchangeStoreType
    ExchangeStoreName         = [string]$Store.ExchangeStoreName
    IsExchangeBacked          = [bool]$Store.IsExchangeBacked
    IsSynchronized            = [bool]$Store.IsSynchronized
    MayContainHeaderOnlyItems = [bool]$Store.MayContainHeaderOnlyItems
    SizeBytes                 = $Store.SizeBytes
    HeadroomMb                = $_headroomMb
    ExceedsWarning            = $_overWarning
    DeliveredFrom             = @($Store.DeliveredFrom)
    Evidence                  = [string]$Store.Evidence
  }
}

# -----------------------------------------------------------------------------

$_context = $null
$_accounts = @()
$_stores = @()
$_findings = New-Object System.Collections.ArrayList
$_outlookVersion = $null
$_outlookMajor = 0
$_fileLimit = $null

try {
  Write-Log -Message 'Connecting to Outlook and reading the profile...' -Color Cyan
  $_context = Connect-Outlook

  $_outlookVersion = [string]$_context.App.Version
  $_outlookMajor = [int](($_outlookVersion -split '\.')[0])
  if ($_outlookMajor -lt 12) {
    throw "Outlook 2007 (version 12) or later is required. Detected Outlook version: $_outlookVersion"
  }

  $_fileLimit = Get-OutlookAccountFileLimit -OutlookVersion "$_outlookMajor.0"

  # One correlated read of accounts and stores. Neither command opens a folder,
  # so this completes where standard-folder discovery is unavailable.
  $_accounts = @(Get-OutlookAccount -Namespace $_context.Namespace)
  $_stores = @(Get-OutlookStore -Namespace $_context.Namespace | ForEach-Object { Get-OutlookAccountStoreRecord -Store $_ -FileLimit $_fileLimit })
}
finally {
  if ($_context) {
    Remove-ComObject $_context.Namespace $_context.App
  }

  Invoke-ComGarbageCollection
}

# -----------------------------------------------------------------------------

function Add-OutlookAccountFinding {
  param (
    [string]
    $Severity,

    [string]
    $Code,

    [string]
    $Subject,

    [string]
    $Detail
  )

  $null = $_findings.Add([PSCustomObject]@{
      Severity = $Severity
      Code     = $Code
      Subject  = $Subject
      Detail   = $Detail
    })
}

# An Exchange-backed store in the profile is what makes an unresolved delivery
# dangerous rather than merely untidy: the store that could not be named may be
# the OST, and removing the Exchange account would take that mail with it.
$_hasExchangeBackedStore = @($_stores | Where-Object { $_.IsExchangeBacked }).Count -gt 0

# Stated once so the reports distinguish a profile that cannot be read from one
# that is configured wrongly. The two call for different work.
$_unresolvedCause = 'The delivery store did not resolve.'

if ($_outlookMajor -lt $_deliveryStoreMajor) {
  $_unresolvedCause = "Outlook $_outlookVersion does not expose a per-account delivery store; Outlook 2010 is the first version that does."
}

foreach ($_account in $_accounts) {
  $_target = @()

  if ($_account.DeliveryStoreID) {
    $_target = @($_stores | Where-Object { $_.StoreID -eq $_account.DeliveryStoreID })
  }

  # Reported ahead of the delivery findings because it depends only on the
  # account type and the client version. On a client that cannot name the
  # delivery store this is still the record of which workstation holds the
  # mailbox's local sent mail, and that record disappears on upgrade.
  if ($_account.AccountTypeName -eq 'IMAP' -and $_outlookMajor -lt $_serverSentItemsMajor) {
    $_sentItemsStore = 'the local delivery store'

    if ($_target.Count -eq 1) {
      $_sentItemsStore = "the local delivery store '$($_target[0].DisplayName)'"
    }

    Add-OutlookAccountFinding -Severity 'Warning' -Code 'LocalImapSentItems' -Subject $_account.DisplayName -Detail "Outlook $_outlookVersion saves this IMAP account's sent mail and drafts into $_sentItemsStore, not into the server's folders. Collect Sent Items and Drafts from this workstation before the mailbox counts as archived."
  }

  # Fail closed. An unresolved delivery store cannot be shown to be safe, so on
  # a profile holding an Exchange-backed store it is reported as critical and
  # the report stops being an exit gate. Exchange mail belongs in the mailbox
  # regardless, so only the other account types are at risk.
  $_unresolvedSeverity = 'Warning'

  if ($_hasExchangeBackedStore -and $_account.AccountType -ne 0) {
    $_unresolvedSeverity = 'Critical'
  }

  if (-not $_account.DeliveryStoreID) {
    $_detail = 'No delivery store resolved. Confirm the delivery location in Account Settings before changing this profile.'

    if ($_unresolvedSeverity -eq 'Critical') {
      $_detail = "$_unresolvedCause This profile holds an Exchange-backed store, so this $($_account.AccountTypeName) account may deliver into it. Read the delivery location in Account Settings, re-point it to a PST, and treat this account as unresolved rather than clear until a client that can report it confirms the change."
    }

    Add-OutlookAccountFinding -Severity $_unresolvedSeverity -Code 'NoDeliveryStore' -Subject $_account.DisplayName -Detail $_detail

    continue
  }

  if ($_target.Count -ne 1) {
    $_detail = "Delivery store '$($_account.DeliveryStoreName)' is not among the profile data files."

    if ($_unresolvedSeverity -eq 'Critical') {
      $_detail = "$_detail This profile holds an Exchange-backed store, so it cannot be ruled out as the destination. Confirm the delivery location in Account Settings before removing any account."
    }

    Add-OutlookAccountFinding -Severity $_unresolvedSeverity -Code 'UnknownDeliveryStore' -Subject $_account.DisplayName -Detail $_detail

    continue
  }

  # The configuration that makes account removal destructive: an account whose
  # mail lands in a store belonging to a different account.
  if ($_account.AccountType -ne 0 -and $_target[0].IsExchangeBacked) {
    Add-OutlookAccountFinding -Severity 'Critical' -Code 'DeliversIntoExchange' -Subject $_account.DisplayName -Detail "$($_account.AccountTypeName) mail is delivered into the Exchange-backed store '$($_target[0].DisplayName)'. Removing the Exchange account discards this account's mail with the OST. Re-point delivery to a PST first."
  }
}

foreach ($_store in $_stores) {
  if ($_store.DataFileFormat -eq 'OST') {
    Add-OutlookAccountFinding -Severity 'Information' -Code 'SynchronizedStore' -Subject $_store.DisplayName -Detail 'OST contents are not recoverable from the file alone. Archive through a connected profile, not by copying the file.'
  }
  elseif ($_store.IsSynchronized) {
    Add-OutlookAccountFinding -Severity 'Information' -Code 'SynchronizedStore' -Subject $_store.DisplayName -Detail "This data file mirrors server state. $($_store.Evidence) Removing mail from it can remove the server copy and reach every other client."
  }

  # Names the stores where an archive must verify each message body first.
  if ($_store.MayContainHeaderOnlyItems) {
    Add-OutlookAccountFinding -Severity 'Information' -Code 'HeaderOnlyItems' -Subject $_store.DisplayName -Detail 'Items here can be header-only stubs. Archiving verifies each message is fully downloaded and stops rather than moving a stub; download full items in Outlook first.'
  }

  if ($_store.ExceedsWarning) {
    Add-OutlookAccountFinding -Severity 'Critical' -Code 'FileSizeLimit' -Subject $_store.DisplayName -Detail "$(Format-OutlookAccountSize -Bytes $_store.SizeBytes) of a $($_fileLimit.MaximumMb) MB limit, $($_store.HeadroomMb) MB remaining. Outlook stops delivering into this file when the limit is reached."
  }
}

# -----------------------------------------------------------------------------

Write-Log -Message "`nOutlook $_outlookVersion | data file limit: $($_fileLimit.MaximumMb) MB (warn at $($_fileLimit.WarningMb) MB, source: $($_fileLimit.Source))" -Color Cyan

Write-Log -Message "`nAccounts: $(@($_accounts).Count)" -Color Cyan
Write-Log -Message ('  {0,-38} {1,-9} {2}' -f 'Account', 'Type', 'Delivers into') -Color Gray
foreach ($_account in $_accounts) {
  $_deliveryLabel = $_account.DeliveryStoreName
  if (-not $_deliveryLabel) {
    $_deliveryLabel = '(none resolved)'
  }

  Write-Log -Message ('  {0,-38} {1,-9} {2}' -f $_account.DisplayName, $_account.AccountTypeName, $_deliveryLabel) -Color Gray
}

Write-Log -Message "`nData files: $(@($_stores).Count)" -Color Cyan
Write-Log -Message ('  {0,-32} {1,-7} {2,-18} {3,-10} {4}' -f 'Store', 'Format', 'Exchange', 'Size', 'Path') -Color Gray
foreach ($_store in $_stores) {
  Write-Log -Message ('  {0,-32} {1,-7} {2,-18} {3,-10} {4}' -f $_store.DisplayName, $_store.DataFileFormat, $_store.ExchangeStoreName, (Format-OutlookAccountSize -Bytes $_store.SizeBytes), $_store.FilePath) -Color Gray

  if (@($_store.DeliveredFrom).Count -gt 0) {
    Write-Log -Message ('      receives delivery from: {0}' -f (@($_store.DeliveredFrom) -join ', ')) -Color Gray
  }
}

$_critical = @($_findings | Where-Object { $_.Severity -eq 'Critical' })
$_warnings = @($_findings | Where-Object { $_.Severity -eq 'Warning' })

if (@($_findings).Count -eq 0) {
  Write-Log -Message "`nFindings: none" -Color Green
}
else {
  $_findingColor = 'Yellow'
  if (@($_critical).Count -gt 0) {
    $_findingColor = 'Red'
  }

  Write-Log -Message "`nFindings: $(@($_critical).Count) critical, $(@($_warnings).Count) warning, $(@($_findings).Count - @($_critical).Count - @($_warnings).Count) informational" -Color $_findingColor
  foreach ($_finding in $_findings) {
    Write-Log -Message ('  [{0}] {1}: {2}' -f $_finding.Severity, $_finding.Subject, $_finding.Detail) -Color Gray
  }
}

if (@($_critical).Count -gt 0) {
  Write-Warning 'This profile has a configuration where removing an account destroys another account''s mail. Resolve the critical findings before changing accounts or data files.'
}

$_inventory = [PSCustomObject]@{
  SchemaVersion  = 2
  Script         = 'Show-OutlookAccount'
  StartedAt      = $_startedAt.ToString('o')
  FinishedAt     = (Get-Date).ToString('o')
  ComputerName   = $env:COMPUTERNAME
  UserName       = $_outlookUser.UserName
  OutlookVersion = $_outlookVersion
  FileLimit      = $_fileLimit
  Accounts       = @($_accounts)
  Stores         = @($_stores)
  Findings       = @($_findings)
  CriticalCount  = @($_critical).Count
  ReportPath     = $_reportPath
}

if ($_reportPath) {
  try {
    $_json = $_inventory | ConvertTo-Json -Depth 8
    $_encoding = New-Object Text.UTF8Encoding($true)
    [IO.File]::WriteAllText($_reportPath, $_json, $_encoding)
    Write-Log -Message "`nJSON report: $_reportPath" -Color Cyan
  }
  catch {
    $_inventory.ReportPath = $null
    Write-Warning "JSON report could not be written: $($_.Exception.Message)"
  }
}

$_inventoryResult = New-OperationResult -Target 'Outlook' -Source 'Outlook' -Action 'Inventory' -Status 'Completed' -Detail ('{0} account(s), {1} data file(s), {2} critical finding(s)' -f @($_accounts).Count, @($_stores).Count, @($_critical).Count)
$_operationLog = Write-OperationResultLog -Results @($_inventoryResult) -ScriptName 'Show-OutlookAccount' -Name 'winkit'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru) {
  $_inventory
}
