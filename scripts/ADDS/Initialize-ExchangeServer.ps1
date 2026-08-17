#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Installs Exchange Server 2019 (Mailbox role or Management Tools only) on a
  domain-joined Windows Server via WinRM.

.DESCRIPTION
  Orchestrates Exchange Server installation over WinRM (Invoke-Command). This
  script runs AFTER Setup-ADController style promotion: the target must be
  domain-joined, and schema/AD preparation requires a functioning domain.

  Two genuinely different scopes are supported and selected explicitly via
  -Mode (mandatory, no default):

    Full            - /PrepareSchema + /PrepareAD + /PrepareDomain, then
                      /Mode:Install /Roles:Mailbox. The complete on-prem
                      Exchange server path, including the forest-wide schema
                      extension (effectively irreversible - gated behind the
                      explicit -ConfirmSchemaExtension switch).
    ManagementTools - Skips all preparation stages (they structurally cannot
                      run here) and installs /Mode:Install
                      /Roles:ManagementTools only, for recipient management
                      against Exchange Online.

  Phases:

    Validate - Domain-joined check, pending-reboot detection, .NET version,
               media reachability, disk space, group-membership advisories,
               feature/installer inventory. Best-effort local validation: there
               is no equivalent of Test-ADDSForestInstallation for Exchange.
    Prepare  - Full mode only: the three Setup.exe preparation invocations in
               order (/PrepareSchema -> /PrepareAD -> /PrepareDomain). Each
               invocation is wrapped with a hard timeout and synthesized from
               exit code + ExchangeSetup.log tail.
    Install  - Windows feature install (versioned per Exchange CU), controlled
               reboot if required, prerequisite installers (VC++ 2012/2013,
               URL Rewrite, UCMA 4.0 from the media), then Setup.exe
               /Mode:Install. Setup.exe is re-runnable after failure - rely on
               its built-in resume behavior; do not manufacture state
               detection that the installer already provides.
    Verify   - Exchange services, Management Shell availability, server and
               database state.

  Without -Phase every phase runs in order. -Phase runs only the named phase
  (intended for crash recovery); combine -Phase <name> -Finish to resume the
  remaining phases through the end of the run.

  The installer is never transferred over WinRM: -ExchangeMediaPath must point
  at an already-staged Setup.exe reachable from the target (SMB share or
  pre-mounted ISO). Prerequisite installers (small, a few MB) are downloaded
  by the orchestrator and pushed via Copy-Item -ToSession.

  The Hybrid Configuration Wizard is explicitly out of scope: this script does
  installation and AD preparation only, not hybrid mail-flow configuration.

  All human-facing progress goes to the console via Write-Log; the output
  stream is reserved for the phase envelopes. With -OutputPath, every envelope
  is additionally appended to that file as one JSON line (JSONL).

.PARAMETER Server
  Target server (static IP or name).

.PARAMETER Credential
  Domain account. Full preparation steps need Schema Admins and Enterprise
  Admins membership. If the account was just added to those groups, open a
  fresh session: Kerberos group memberships are baked into the logon ticket,
  and a session opened before the membership change will not reflect it.

.PARAMETER Mode
  Installation scope: Full (Mailbox role) or ManagementTools (Exchange
  Management Tools only, no prep stages).

.PARAMETER DomainName
  Fully qualified domain name of the domain, e.g. company.com.

.PARAMETER OrganizationName
  Exchange organization name, required for /PrepareAD and the Mailbox role
  install.

.PARAMETER ExchangeMediaPath
  Full path to Setup.exe as reachable from the target, e.g.
  D:\Exchange\Setup.exe or \\fileserver\share\Exchange\Setup.exe. Never
  transferred over WinRM.

.PARAMETER ExchangeVersion
  Exchange version/CU. Selects the versioned data table (Windows features,
  minimum .NET, license-acceptance flag, prerequisite installer sources).
  Only 2019-CU14 is populated; add entries deliberately when targeting a new
  CU - do not assume the table stays current.

.PARAMETER Phase
  Run only the named phase. Defaults to running all phases in order.

.PARAMETER Finish
  Only meaningful together with -Phase: run the named phase, then continue
  through the remaining phases to the end of the run.

.PARAMETER ConfirmSchemaExtension
  Required to run /PrepareSchema. Schema extension is forest-wide and
  effectively irreversible; this switch is the explicit acknowledgment gate.

.PARAMETER OutputPath
  Optional file to append each phase envelope to as a JSON line.

.PARAMETER OperationTimeoutMinutes
  WinRM session operation timeout.

.PARAMETER SetupTimeoutMinutes
  Hard ceiling for each Setup.exe invocation. Treat "still running past the
  ceiling" as a condition to investigate, not as something that will finish.

.EXAMPLE
  PS> ./Initialize-ExchangeServer.ps1 -Server 192.0.2.11 -Credential $domainAdmin -Mode Full -DomainName company.com -OrganizationName 'AdNoctem' -ExchangeMediaPath '\\fileserver\media\Exchange2019CU14\Setup.exe' -ConfirmSchemaExtension
  Runs all four phases against a fresh domain-joined server, including the
  gated schema extension.

.EXAMPLE
  PS> ./Initialize-ExchangeServer.ps1 -Server 192.0.2.11 -Credential $domainAdmin -Mode ManagementTools -DomainName company.com -OrganizationName 'AdNoctem' -ExchangeMediaPath '\\fileserver\media\Exchange2019CU14\Setup.exe'
  Installs only the Exchange Management Tools; preparation stages are skipped
  structurally.

.EXAMPLE
  PS> ./Initialize-ExchangeServer.ps1 -Server 192.0.2.11 -Credential $domainAdmin -Mode Full -DomainName company.com -OrganizationName 'AdNoctem' -ExchangeMediaPath '\\fileserver\media\Exchange2019CU14\Setup.exe' -ConfirmSchemaExtension -Phase Install -Finish
  Resumes a crashed run at the Install phase and continues through Verify.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Carriage-return progress lines require Write-Host for in-place console updates.')]

# Parameters consumed by the phase/helper functions below (which live in their
# own scopes) are intentionally referenced only from those functions; the
# unused-parameter rule cannot see function-scope usage and needs the
# per-parameter exemptions.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Credential', Justification = 'Consumed by phase and helper functions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'OrganizationName', Justification = 'Consumed by phase functions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ExchangeMediaPath', Justification = 'Consumed by phase and helper functions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ConfirmSchemaExtension', Justification = 'Consumed by the Prepare phase.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'OutputPath', Justification = 'Consumed by Write-PhaseEnvelope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'OperationTimeoutMinutes', Justification = 'Consumed by New-TargetSession.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'SetupTimeoutMinutes', Justification = 'Consumed by Invoke-SetupExeStep.')]

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [string]$Server,

  [Parameter(Mandatory = $true)]
  [pscredential]$Credential,

  [Parameter(Mandatory = $true)]
  [ValidateSet('ManagementTools', 'Full')]
  [string]$Mode,

  [Parameter(Mandatory = $true)]
  [string]$DomainName,

  [Parameter(Mandatory = $true)]
  [string]$OrganizationName,

  [Parameter(Mandatory = $true)]
  [string]$ExchangeMediaPath,

  [Parameter(Mandatory = $false)]
  [ValidateSet('2019-CU14')]
  [string]$ExchangeVersion = '2019-CU14',

  [Parameter(Mandatory = $false)]
  [ValidateSet('Validate', 'Prepare', 'Install', 'Verify')]
  [string]$Phase,

  [Parameter(Mandatory = $false)]
  [switch]$Finish,

  [Parameter(Mandatory = $false)]
  [switch]$ConfirmSchemaExtension,

  [Parameter(Mandatory = $false)]
  [string]$OutputPath,

  [Parameter(Mandatory = $false)]
  [int]$OperationTimeoutMinutes = 45,

  [Parameter(Mandatory = $false)]
  [int]$SetupTimeoutMinutes = 60
)

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

# ---------------------------------------------------------------------------
# Versioned data table.
#
# Prerequisites drift across Exchange versions/CUs (Windows features, minimum
# .NET, license-acceptance flag, prerequisite installer sources). This table is
# the single place that data lives; review it deliberately when targeting a new
# CU instead of hardcoding values into the phase code.
# ---------------------------------------------------------------------------

$exchangeProfiles = @{
  '2019-CU14' = @{
    MinimumDotNetRelease = 528040
    LicenseFlag = '/IAcceptExchangeServerLicenseTerms_DiagnosticDataOFF'
    WindowsFeatures = @(
      'Server-Media-Foundation',
      'NET-Framework-45-Features',
      'RPC-over-HTTP-proxy',
      'RSAT-Clustering',
      'RSAT-Clustering-CmdInterface',
      'RSAT-Clustering-Management',
      'RSAT-Clustering-PowerShell',
      'WAS-Process-Model',
      'Web-Asp-Net45',
      'Web-Basic-Auth',
      'Web-Client-Auth',
      'Web-Digest-Auth',
      'Web-Dir-Browsing',
      'Web-Dyn-Compression',
      'Web-Http-Errors',
      'Web-Http-Logging',
      'Web-Http-Redirect',
      'Web-Http-Tracing',
      'Web-ISAPI-Ext',
      'Web-ISAPI-Filter',
      'Web-Metabase',
      'Web-Mgmt-Console',
      'Web-Mgmt-Service',
      'Web-Net-Ext45',
      'Web-Request-Monitor',
      'Web-Server',
      'Web-Stat-Compression',
      'Web-Static-Content',
      'Web-Windows-Auth',
      'Web-WMI',
      'Windows-Identity-Foundation',
      'RSAT-ADDS'
    )
    PrereqInstallers = @{
      VCRedist2012x64 = 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe'
      VCRedist2013x64 = 'https://aka.ms/highdpimfc2013x64enu'
      UrlRewrite = 'https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi'
    }
    Roles = @{
      ManagementTools = 'ManagementTools'
      Full = 'Mailbox'
    }
  }
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-PhaseEnvelope {
  <#
    Emits one phase result envelope on the output stream and optionally appends
    it to the -OutputPath file as a single JSON line. This is the only thing
    that writes to the success stream: everything else uses Write-Log.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    [string]$PhaseName,

    [Parameter(Mandatory = $true)]
    [bool]$Success,

    [Parameter(Mandatory = $true)]
    [string]$StartedAt,

    [Parameter(Mandatory = $true)]
    [string]$CompletedAt,

    [bool]$RebootRequired = $false,

    [System.Collections.ArrayList]$Diagnostics,

    [string]$ErrorText
  )

  $envelope = [pscustomobject][ordered]@{
    phase = $PhaseName
    success = $Success
    server = $Server
    domainName = $DomainName
    mode = $Mode
    exchangeVersion = $ExchangeVersion
    startedAt = $StartedAt
    completedAt = $CompletedAt
    rebootRequired = $RebootRequired
    diagnostics = @($Diagnostics)
    error = $ErrorText
  }

  $json = $envelope | ConvertTo-Json -Depth 6 -Compress
  if ($OutputPath) {
    Add-Content -LiteralPath $OutputPath -Value $json -Encoding UTF8
  }

  $envelope
}

function Get-RedactedError {
  <#
    Returns a scrubbed exception message: the domain password is the one
    secret this script holds in plaintext-capable form, and if it ever appears
    in an error string it is replaced before anything is logged or emitted.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.ErrorRecord]$ErrorRecord
  )

  $message = $ErrorRecord.Exception.Message
  $secret = $null
  try {
    $secret = $Credential.GetNetworkCredential().Password
  }
  catch {
    $secret = $null
  }

  if ($secret -and $message -and $message.Contains($secret)) {
    $message = $message.Replace($secret, '<redacted>')
  }

  if (-not $message) {
    $message = 'Unknown error (no message available).'
  }

  $message
}

function New-TargetSession {
  <#
    Creates a WinRM session with an explicit, generous operation timeout.
    Returns $null (after recording a diagnostic) when the connection fails.
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Session management only; New/Remove-PSSession do not warrant ShouldProcess gating.')]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [pscredential]$SessionCredential,

    [System.Collections.ArrayList]$Diagnostics
  )

  $sessionOption = New-PSSessionOption -OperationTimeout ([int]($OperationTimeoutMinutes * 60000)) -IdleTimeout ([int]($OperationTimeoutMinutes * 60000))
  $session = $null

  try {
    $session = New-PSSession -ComputerName $Server -Credential $SessionCredential -SessionOption $sessionOption -ErrorAction Stop
  }
  catch {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target $Server -Source 'WinRM' -Action 'Connect' -Status 'Failed' -Detail (Get-RedactedError $_)
    }
    return $null
  }

  if ($null -ne $Diagnostics) {
    $mechanism = $session.Runspace.ConnectionInfo.AuthenticationMechanism
    Add-OperationResult -Results $Diagnostics -Target $Server -Source 'WinRM' -Action 'Connect' -Status 'Completed' -Detail "Session established (auth: $mechanism)."
  }

  $session
}

function Test-TargetPendingReboot {
  <#
    Checks the standard pending-reboot indicators on the target before any
    state-changing step.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [System.Collections.ArrayList]$Diagnostics
  )

  $pendingScript = {
    $indicators = @()

    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
      $indicators += 'CBS RebootPending'
    }
    $cbs = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing' -ErrorAction SilentlyContinue
    if ($cbs -and (($cbs.PSObject.Properties.Name -contains 'RebootInProgress') -or ($cbs.PSObject.Properties.Name -contains 'PackagesPending'))) {
      $indicators += 'CBS servicing state'
    }
    $sessionManagerPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    if (Test-Path -LiteralPath (Join-Path -Path $sessionManagerPath -ChildPath 'PendingFileRenameOperations')) {
      $indicators += 'PendingFileRenameOperations'
    }
    if (Test-Path -LiteralPath (Join-Path -Path $sessionManagerPath -ChildPath 'PendingFileRenameOperations2')) {
      $indicators += 'PendingFileRenameOperations2'
    }
    $runOnce = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -ErrorAction SilentlyContinue
    if ($runOnce -and ($runOnce.PSObject.Properties.Name -contains 'DVDRebootSignal')) {
      $indicators += 'DVDRebootSignal'
    }
    $netlogon = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -ErrorAction SilentlyContinue
    if ($netlogon -and (($netlogon.PSObject.Properties.Name -contains 'JoinDomain') -or ($netlogon.PSObject.Properties.Name -contains 'AvoidSpnSet'))) {
      $indicators += 'Netlogon join'
    }
    $autoUpdate = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update' -ErrorAction SilentlyContinue
    if ($autoUpdate -and (($autoUpdate.PSObject.Properties.Name -contains 'RebootRequired') -or ($autoUpdate.PSObject.Properties.Name -contains 'PostRebootReporting'))) {
      $indicators += 'Windows Update'
    }
    try {
      $ccm = Get-CimInstance -Namespace 'root\ccm\clientsdk' -ClassName CCM_ClientUtilities -ErrorAction SilentlyContinue
      if ($ccm -and [bool]($ccm.DetermineIfRebootPending()).RebootPending) {
        $indicators += 'SCCM/MECM client'
      }
    }
    catch {
      $ccm = $null
    }
    $computerNameKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName'
    $activeName = (Get-ItemProperty "$computerNameKey\ActiveComputerName" -ErrorAction SilentlyContinue).ComputerName
    $configuredName = (Get-ItemProperty "$computerNameKey\ComputerName" -ErrorAction SilentlyContinue).ComputerName
    if ($activeName -and $configuredName -and ($activeName -ne $configuredName)) {
      $indicators += 'Computer rename'
    }
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    if ($computerSystem -and $computerSystem.PendingSystemReboot) {
      $indicators += 'CIM PendingSystemReboot'
    }

    [pscustomobject]@{
      PendingReboot = ($indicators.Count -gt 0)
      Indicators = $indicators
    }
  }

  try {
    $result = Invoke-Command -Session $Session -ScriptBlock $pendingScript -ErrorAction Stop
  }
  catch {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target $Server -Source 'PendingReboot' -Action 'Check' -Status 'Failed' -Detail (Get-RedactedError $_)
    }
    return $true
  }

  if ($null -ne $Diagnostics) {
    Add-OperationResult -Results $Diagnostics -Target $Server -Source 'PendingReboot' -Action 'Check' -Status $(if ($result.PendingReboot) { 'Failed' } else { 'Completed' }) -Detail $(if ($result.PendingReboot) { "Pending reboot detected - indicators: $($result.Indicators -join ', ')" } else { 'No pending reboot detected.' })
  }

  [bool]$result.PendingReboot
}

function Test-MediaReachable {
  <#
    Confirms the staged Setup.exe path is visible from the target. The
    installer itself is never pushed through WinRM.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [System.Collections.ArrayList]$Diagnostics
  )

  $mediaCheckScript = {
    param($path)
    Test-Path -LiteralPath $path -PathType Leaf
  }

  $exists = $false
  try {
    $exists = Invoke-Command -Session $Session -ArgumentList $ExchangeMediaPath -ScriptBlock $mediaCheckScript -ErrorAction Stop
  }
  catch {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target $ExchangeMediaPath -Source 'Media' -Action 'Check' -Status 'Failed' -Detail (Get-RedactedError $_)
    }
    return $false
  }

  if ($null -ne $Diagnostics) {
    Add-OperationResult -Results $Diagnostics -Target $ExchangeMediaPath -Source 'Media' -Action 'Check' -Status $(if ($exists) { 'Completed' } else { 'Failed' }) -Detail $(if ($exists) { 'Media reachable from target.' } else { 'Media path not reachable from target.' })
  }

  $exists
}

function Invoke-SetupExeStep {
  <#
    Runs one Setup.exe invocation remotely with a hard timeout and synthesizes
    the result from exit code plus the ExchangeSetup.log tail. Setup.exe is a
    native installer, not a cmdlet: the log is the actual source of truth for
    anything beyond bare success/failure.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [Parameter(Mandatory = $true)]
    [string]$SetupExe,

    [Parameter(Mandatory = $true)]
    [string[]]$Arguments,

    [Parameter(Mandatory = $true)]
    [string]$StepName,

    [Parameter(Mandatory = $true)]
    [System.Collections.ArrayList]$Diagnostics
  )

  if (-not $PSCmdlet.ShouldProcess($Server, "Run Exchange Setup: $StepName")) {
    Add-OperationResult -Results $Diagnostics -Target $StepName -Source 'ExchangeSetup' -Action 'Invoke' -Status 'Skipped' -Detail 'WhatIf - Setup.exe invocation skipped.'
    return [pscustomobject]@{ Success = $true; ExitCode = $null; LogExcerpt = '' }
  }

  # Setup.exe runs inside a background job in the remote session so the
  # orchestrator can poll with a visible progress bar and still harvest the
  # exit code + log tail when the job finishes.
  $setupJobScript = {
    param($setupPath, $setupArgs)
    $inner = {
      param($path, $argsList)
      $process = Start-Process -FilePath $path -ArgumentList $argsList -Wait -PassThru -NoNewWindow -ErrorAction Stop
      $logPath = Join-Path -Path "$env:SystemDrive\ExchangeSetupLogs" -ChildPath 'ExchangeSetup.log'
      $logTail = if (Test-Path -LiteralPath $logPath) { @(Get-Content -LiteralPath $logPath -Tail 25) } else { @() }
      [pscustomobject]@{
        ExitCode = $process.ExitCode
        LogTail = $logTail
      }
    }
    $job = Start-Job -ScriptBlock $inner -ArgumentList $setupPath, $setupArgs
    $job.Id
  }

  $setupJobId = $null
  try {
    $setupJobId = Invoke-Command -Session $Session -ArgumentList $SetupExe, $Arguments -ScriptBlock $setupJobScript -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target $StepName -Source 'ExchangeSetup' -Action 'Invoke' -Status 'Failed' -Detail (Get-RedactedError $_)
    return [pscustomobject]@{ Success = $false; ExitCode = $null; LogExcerpt = '' }
  }

  $setupJobStateScript = {
    param($id)
    $job = Get-Job -Id $id -ErrorAction SilentlyContinue
    if ($job) { $job.State } else { 'Missing' }
  }

  $pollStartedAt = Get-Date
  $pollDeadline = $pollStartedAt.AddMinutes($SetupTimeoutMinutes)
  $setupJobState = 'Running'
  $setupTimedOut = $false
  while ($setupJobState -notin @('Completed', 'Failed', 'Missing')) {
    try {
      $setupJobState = Invoke-Command -Session $Session -ArgumentList $setupJobId -ScriptBlock $setupJobStateScript -ErrorAction Stop
    }
    catch {
      $setupJobState = 'Missing'
    }
    $elapsedTime = (Get-Date) - $pollStartedAt
    $elapsedSeconds = [int]$elapsedTime.TotalSeconds
    $setupPct = [int][math]::Min(($elapsedSeconds / ($SetupTimeoutMinutes * 60)) * 100, 100)
    Write-Progress -Activity "Exchange Setup: $StepName" -Status "elapsed: $elapsedSeconds s / $SetupTimeoutMinutes min, job state: $setupJobState" -PercentComplete $setupPct
    Write-Host ("`r$StepName ({0}s / {1}min): {2,-40}" -f $elapsedSeconds, $SetupTimeoutMinutes, $setupJobState) -NoNewline -ForegroundColor Cyan
    if ((Get-Date) -gt $pollDeadline) {
      $setupTimedOut = $true
      break
    }
    Start-Sleep -Seconds 15
  }
  Write-Progress -Activity "Exchange Setup: $StepName" -Completed
  $setupTotalSeconds = [int]((Get-Date) - $pollStartedAt).TotalSeconds
  Write-Host ("`r$StepName finished ({0}s).{1}" -f $setupTotalSeconds, (' ' * 40)) -ForegroundColor Cyan

  if ($setupTimedOut -or $setupJobState -eq 'Missing') {
    Add-OperationResult -Results $Diagnostics -Target $StepName -Source 'ExchangeSetup' -Action 'Invoke' -Status 'Unknown' -Detail "Setup.exe still running past the $SetupTimeoutMinutes minute ceiling (job state: $setupJobState). Do not re-run blindly - check the log for progress first."
    return [pscustomobject]@{ Success = $false; ExitCode = $null; LogExcerpt = '' }
  }

  # Harvest the job result (best effort; Receive-Job re-throws remote errors).
  $harvestJobScript = {
    param($id)
    $job = Get-Job -Id $id -ErrorAction SilentlyContinue
    if (-not $job) {
      return $null
    }
    $output = @()
    try {
      $output = @(Receive-Job -Id $id -ErrorAction Stop)
    }
    catch {
      $output = @()
    }
    Remove-Job -Id $id -Force -ErrorAction SilentlyContinue
    if ($output.Count -gt 0) { $output[0] } else { $null }
  }

  $result = $null
  try {
    $result = Invoke-Command -Session $Session -ArgumentList $setupJobId -ScriptBlock $harvestJobScript -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target $StepName -Source 'ExchangeSetup' -Action 'Invoke' -Status 'Failed' -Detail "Could not harvest the Setup.exe result from the remote job: $(Get-RedactedError $_)"
    return [pscustomobject]@{ Success = $false; ExitCode = $null; LogExcerpt = '' }
  }

  if (-not $result) {
    Add-OperationResult -Results $Diagnostics -Target $StepName -Source 'ExchangeSetup' -Action 'Invoke' -Status 'Failed' -Detail 'Setup.exe job produced no result object - treat as failure and review the log.'
    return [pscustomobject]@{ Success = $false; ExitCode = $null; LogExcerpt = '' }
  }

  $excerpt = if ($result.LogTail) { ($result.LogTail -join "`n") } else { '(no ExchangeSetup.log content yet)' }

  $bannerFound = [bool]($result.LogTail -match 'completed successfully')
  $success = ($result.ExitCode -eq 0 -and $bannerFound)
  Add-OperationResult -Results $Diagnostics -Target $StepName -Source 'ExchangeSetup' -Action 'Invoke' -Status $(if ($success) { 'Completed' } else { 'Failed' }) -Detail "exit code: $($result.ExitCode), completion banner found: $bannerFound" -Property @{ LogExcerpt = $excerpt }

  [pscustomobject]@{ Success = $success; ExitCode = $result.ExitCode; LogExcerpt = $excerpt }
}

function Install-ExchangeFeature {
  <#
    Installs the Windows features declared for the selected Exchange CU,
    idempotently (missing features only).
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [Parameter(Mandatory = $true)]
    [string[]]$Features,

    [Parameter(Mandatory = $true)]
    [System.Collections.ArrayList]$Diagnostics
  )

  $featureJobScript = {
    param($featureNames)
    $inner = {
      param($names)
      $ConfirmPreference = 'None'
      $missing = @(Get-WindowsFeature -Name $names | Where-Object { -not $_.Installed })
      $rebootRequired = $false
      if ($missing.Count -gt 0) {
        $installResult = Install-WindowsFeature -Name ($missing | ForEach-Object { $_.Name }) -Confirm:$false -ErrorAction Stop
        $rebootRequired = [bool]$installResult.RestartNeeded
      }
      $stillMissing = @(Get-WindowsFeature -Name $names | Where-Object { -not $_.Installed } | ForEach-Object { $_.Name })
      [pscustomobject]@{
        RebootRequired = $rebootRequired
        StillMissing = $stillMissing
      }
    }
    $job = Start-Job -ScriptBlock $inner -ArgumentList (, $featureNames)
    $job.Id
  }

  $featureJobId = $null
  try {
    $featureJobId = Invoke-Command -Session $Session -ArgumentList (, $Features) -ScriptBlock $featureJobScript -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail (Get-RedactedError $_)
    return [pscustomobject]@{ Success = $false; RebootRequired = $false }
  }

  $featureJobStateScript = {
    param($id)
    $job = Get-Job -Id $id -ErrorAction SilentlyContinue
    if ($job) { $job.State } else { 'Missing' }
  }

  $featurePollStartedAt = Get-Date
  $featurePollDeadline = $featurePollStartedAt.AddMinutes($OperationTimeoutMinutes)
  $featureJobState = 'Running'
  $featureTimedOut = $false
  while ($featureJobState -notin @('Completed', 'Failed', 'Missing')) {
    try {
      $featureJobState = Invoke-Command -Session $Session -ArgumentList $featureJobId -ScriptBlock $featureJobStateScript -ErrorAction Stop
    }
    catch {
      $featureJobState = 'Missing'
    }
    $elapsedTime = (Get-Date) - $featurePollStartedAt
    $elapsedSeconds = [int]$elapsedTime.TotalSeconds
    $featurePct = [int][math]::Min(($elapsedSeconds / ($OperationTimeoutMinutes * 60)) * 100, 100)
    Write-Progress -Activity 'Installing Exchange prerequisites (Windows features)' -Status "elapsed: $elapsedSeconds s / $OperationTimeoutMinutes min, job state: $featureJobState" -PercentComplete $featurePct
    Write-Host ("`rInstalling Windows features ({0}s / {1}min): {2,-40}" -f $elapsedSeconds, $OperationTimeoutMinutes, $featureJobState) -NoNewline -ForegroundColor Cyan
    if ((Get-Date) -gt $featurePollDeadline) {
      $featureTimedOut = $true
      break
    }
    Start-Sleep -Seconds 15
  }
  Write-Progress -Activity 'Installing Exchange prerequisites (Windows features)' -Completed
  $featureTotalSeconds = [int]((Get-Date) - $featurePollStartedAt).TotalSeconds
  Write-Host ("`rWindows feature install finished ({0}s).{1}" -f $featureTotalSeconds, (' ' * 40)) -ForegroundColor Cyan

  if ($featureTimedOut -or $featureJobState -eq 'Missing') {
    Add-OperationResult -Results $Diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Install' -Status 'Unknown' -Detail "Feature install still running past the $OperationTimeoutMinutes minute ceiling (job state: $featureJobState)."
    return [pscustomobject]@{ Success = $false; RebootRequired = $false }
  }

  $featureHarvestScript = {
    param($id)
    $job = Get-Job -Id $id -ErrorAction SilentlyContinue
    if (-not $job) {
      return $null
    }
    $output = @()
    try {
      $output = @(Receive-Job -Id $id -ErrorAction Stop)
    }
    catch {
      $output = @()
    }
    Remove-Job -Id $id -Force -ErrorAction SilentlyContinue
    if ($output.Count -gt 0) { $output[0] } else { $null }
  }

  $result = $null
  try {
    $result = Invoke-Command -Session $Session -ArgumentList $featureJobId -ScriptBlock $featureHarvestScript -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail "Could not harvest the feature install result: $(Get-RedactedError $_)"
    return [pscustomobject]@{ Success = $false; RebootRequired = $false }
  }

  if (-not $result) {
    Add-OperationResult -Results $Diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail 'Feature install job produced no result object.'
    return [pscustomobject]@{ Success = $false; RebootRequired = $false }
  }

  if ($result.StillMissing.Count -gt 0) {
    Add-OperationResult -Results $Diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail ("Still missing: {0}" -f ($result.StillMissing -join ', '))
    return [pscustomobject]@{ Success = $false; RebootRequired = $false }
  }

  Add-OperationResult -Results $Diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Install' -Status 'Completed' -Detail "All $($Features.Count) features installed. rebootRequiredByFeature: $($result.RebootRequired)"
  [pscustomobject]@{ Success = $true; RebootRequired = [bool]$result.RebootRequired }
}

function Invoke-PrereqInstaller {
  <#
    Downloads the small prerequisite installers (VC++ runtimes, URL Rewrite)
    on the orchestrator, pushes them via Copy-Item -ToSession, and installs
    them silently on the target. This is fine over PSRP because they are a
    few MB; the multi-GB Exchange installer is never pushed this way.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [Parameter(Mandatory = $true)]
    [hashtable]$Installers,

    [Parameter(Mandatory = $true)]
    [System.Collections.ArrayList]$Diagnostics
  )

  $directorySetupScript = {
    param($directory)
    $null = New-Item -Path $directory -ItemType Directory -Force
  }

  $remoteDirectory = 'C:\Windows\Temp\winkit-exchange-prereqs'
  try {
    $null = Invoke-Command -Session $Session -ArgumentList $remoteDirectory -ScriptBlock $directorySetupScript -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target 'PrereqInstallers' -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail (Get-RedactedError $_)
    return [pscustomobject]@{ Success = $false }
  }

  $localWork = Join-Path $env:TEMP "winkit-exchange-prereqs-$(New-Guid)"
  $null = New-Item -Path $localWork -ItemType Directory -Force
  $overallSuccess = $true

  try {
    foreach ($name in $Installers.Keys) {
      $url = $Installers[$name]
      $extension = if ($url -like '*.msi') { '.msi' } else { '.exe' }
      $fileName = "$name$extension"
      $localFile = Join-Path $localWork $fileName

      try {
        Invoke-WebRequest -Uri $url -OutFile $localFile -UseBasicParsing -ErrorAction Stop
      }
      catch {
        Add-OperationResult -Results $Diagnostics -Target $name -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail "Download failed: $(Get-RedactedError $_)"
        $overallSuccess = $false
        break
      }

      try {
        Copy-Item -LiteralPath $localFile -Destination (Join-Path $remoteDirectory $fileName) -ToSession $Session -ErrorAction Stop
      }
      catch {
        Add-OperationResult -Results $Diagnostics -Target $name -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail "Copy to target failed: $(Get-RedactedError $_)"
        $overallSuccess = $false
        break
      }

      $installerScript = {
        param($directory, $file)
        $installer = Join-Path $directory $file
        $argumentList = if ($file -like '*.msi') { @('/quiet', '/norestart') } else { @('/install', '/quiet', '/norestart') }
        $process = Start-Process -FilePath $installer -ArgumentList $argumentList -Wait -PassThru -NoNewWindow -ErrorAction Stop
        [pscustomobject]@{ ExitCode = $process.ExitCode }
      }

      $installResult = $null
      try {
        $installResult = Invoke-Command -Session $Session -ArgumentList $remoteDirectory, $fileName -ScriptBlock $installerScript -ErrorAction Stop
      }
      catch {
        Add-OperationResult -Results $Diagnostics -Target $name -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail (Get-RedactedError $_)
        $overallSuccess = $false
        break
      }

      if ($installResult.ExitCode -notin @(0, 3010)) {
        Add-OperationResult -Results $Diagnostics -Target $name -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail "Installer exited with code $($installResult.ExitCode)."
        $overallSuccess = $false
        break
      }

      Add-OperationResult -Results $Diagnostics -Target $name -Source 'ExchangePrereq' -Action 'Install' -Status 'Completed' -Detail "Installed (exit code $($installResult.ExitCode))."
    }
  }
  finally {
    Remove-Item -LiteralPath $localWork -Recurse -Force -ErrorAction SilentlyContinue
  }

  [pscustomobject]@{ Success = $overallSuccess }
}

function Invoke-UcmaInstall {
  <#
    Installs UCMA 4.0 from the media staging directory (UCMARedist folder in
    the Exchange ISO).
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,

    [Parameter(Mandatory = $true)]
    [System.Collections.ArrayList]$Diagnostics
  )

  try {
    $ucmaScript = {
      param($installer)
      $process = Start-Process -FilePath $installer -ArgumentList @('/q') -Wait -PassThru -NoNewWindow -ErrorAction Stop
      [pscustomobject]@{ ExitCode = $process.ExitCode }
    }
    $result = Invoke-Command -Session $Session -ArgumentList $InstallerPath -ScriptBlock $ucmaScript -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target 'UCMA 4.0' -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail (Get-RedactedError $_)
    return $false
  }

  if ($result.ExitCode -notin @(0, 3010)) {
    Add-OperationResult -Results $Diagnostics -Target 'UCMA 4.0' -Source 'ExchangePrereq' -Action 'Install' -Status 'Failed' -Detail "UCMA installer exited with code $($result.ExitCode)."
    return $false
  }

  Add-OperationResult -Results $Diagnostics -Target 'UCMA 4.0' -Source 'ExchangePrereq' -Action 'Install' -Status 'Completed' -Detail "Installed (exit code $($result.ExitCode))."
  $true
}

function Invoke-RebootTarget {
  <#
    Restarts the target and blocks until WinRM is confirmed back (authenticated
    probe with the reboot credential), rendering live progress. "WinRM is
    back" is only the connectivity floor; callers re-verify the relevant state
    afterwards.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    [pscredential]$RebootCredential
  )

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  if (-not $PSCmdlet.ShouldProcess($Server, 'Reboot and wait for WinRM to return')) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Reboot' -Status 'Skipped' -Detail 'WhatIf - reboot skipped.'
    return Write-PhaseEnvelope -PhaseName 'reboot' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  try {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Reboot' -Status 'Completed' -Detail 'Restarting target and waiting for WinRM (up to 15 minutes).'
    Restart-Computer -ComputerName $Server -Credential $RebootCredential -Force -ErrorAction Stop

    $rebootStartedAt = Get-Date
    $rebootDeadline = $rebootStartedAt.AddSeconds(900)
    $winrmBack = $false
    while (-not $winrmBack -and (Get-Date) -lt $rebootDeadline) {
      try {
        $null = Invoke-Command -ComputerName $Server -Credential $RebootCredential -ScriptBlock { $true } -ErrorAction Stop
        $winrmBack = $true
      }
      catch {
        $winrmBack = $false
      }
      $elapsedTime = (Get-Date) - $rebootStartedAt
      $elapsedSeconds = [int]$elapsedTime.TotalSeconds
      $rebootPct = [int][math]::Min(($elapsedSeconds / 900) * 100, 100)
      Write-Progress -Activity 'Rebooting target' -Status "waiting for WinRM ($elapsedSeconds s / 900 s)" -PercentComplete $rebootPct
      Write-Host ("`rRebooting ({0}s / 900s): {1,-40}" -f $elapsedSeconds, $Server) -NoNewline -ForegroundColor Cyan
      if (-not $winrmBack) {
        Start-Sleep -Seconds 5
      }
    }
    Write-Progress -Activity 'Rebooting target' -Completed
    $rebootTotalSeconds = [int]((Get-Date) - $rebootStartedAt).TotalSeconds
    Write-Host ("`rReboot finished ({0}s).{1}" -f $rebootTotalSeconds, (' ' * 40)) -ForegroundColor Cyan

    if (-not $winrmBack) {
      throw 'WinRM did not come back within 900 seconds after the reboot.'
    }
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Reboot' -Status 'Completed' -Detail 'WinRM confirmed back after reboot.'
    Write-PhaseEnvelope -PhaseName 'reboot' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }
  catch {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Reboot' -Status 'Failed' -Detail (Get-RedactedError $_)
    Write-PhaseEnvelope -PhaseName 'reboot' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }
}

# ---------------------------------------------------------------------------
# Phase: Validate
# ---------------------------------------------------------------------------

function Invoke-ValidatePhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList
  $exchangeProfile = $exchangeProfiles[$ExchangeVersion]

  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  if (Test-TargetPendingReboot -Session $session -Diagnostics $diagnostics) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'A reboot is pending on the target. Reboot it and re-run before proceeding.'
  }

  $systemScript = {
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $dotNetRelease = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction SilentlyContinue).Release
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DeviceID="C:"' -ErrorAction SilentlyContinue
    $groupOutput = (whoami /groups) -join "`n"
    [pscustomobject]@{
      DomainRole = $computerSystem.DomainRole
      Domain = $computerSystem.Domain
      DotNetRelease = $dotNetRelease
      FreeGigabytes = if ($disk) { [math]::Round($disk.FreeSpace / 1GB, 1) } else { $null }
      SchemaAdmins = ($groupOutput -match 'Schema Admins')
      EnterpriseAdmins = ($groupOutput -match 'Enterprise Admins')
      DomainAdmins = ($groupOutput -match 'Domain Admins')
    }
  }

  $system = $null
  try {
    $system = Invoke-Command -Session $session -ScriptBlock $systemScript -ErrorAction Stop
  }
  catch {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Exchange' -Action 'Validate' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  if ($system.DomainRole -ge 4) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Exchange' -Action 'Validate' -Status 'Failed' -Detail 'Target is a domain controller.'
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Exchange cannot be installed on a domain controller.'
  }

  if ($system.DomainRole -notin @(1, 3)) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Exchange' -Action 'Validate' -Status 'Failed' -Detail "DomainRole: $($system.DomainRole) - target is not domain-joined."
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The target is not domain-joined. Exchange setup requires a functioning domain (run the AD controller script first).'
  }

  Add-OperationResult -Results $diagnostics -Target $Server -Source 'Exchange' -Action 'Validate' -Status 'Completed' -Detail "Target is domain-joined ($($system.Domain))."

  if (-not $system.DotNetRelease -or $system.DotNetRelease -lt $exchangeProfile.MinimumDotNetRelease) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target '.NET Framework' -Source 'Exchange' -Action 'Validate' -Status 'Failed' -Detail "Release $($system.DotNetRelease) is below the minimum $($exchangeProfile.MinimumDotNetRelease) (4.8)."
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The .NET Framework version on the target does not meet the Exchange CU minimum.'
  }
  Add-OperationResult -Results $diagnostics -Target '.NET Framework' -Source 'Exchange' -Action 'Validate' -Status 'Completed' -Detail "Release $($system.DotNetRelease) meets the minimum."

  if ($system.FreeGigabytes -lt 20) {
    Add-OperationResult -Results $diagnostics -Target 'C:' -Source 'Exchange' -Action 'Validate' -Status 'Warn' -Detail "Free space: $($system.FreeGigabytes) GB - below the recommended 20 GB."
  }
  else {
    Add-OperationResult -Results $diagnostics -Target 'C:' -Source 'Exchange' -Action 'Validate' -Status 'Completed' -Detail "Free space: $($system.FreeGigabytes) GB."
  }

  if ($Mode -eq 'Full') {
    if (-not $system.SchemaAdmins -or -not $system.EnterpriseAdmins) {
      Add-OperationResult -Results $diagnostics -Target 'GroupMembership' -Source 'Exchange' -Action 'Validate' -Status 'Warn' -Detail 'The session does not show Schema Admins/Enterprise Admins membership. /PrepareSchema and /PrepareAD will fail without it; note that memberships granted after this session opened will not appear (stale Kerberos ticket).'
    }
    else {
      Add-OperationResult -Results $diagnostics -Target 'GroupMembership' -Source 'Exchange' -Action 'Validate' -Status 'Completed' -Detail 'Schema Admins and Enterprise Admins membership present in the current session.'
    }
  }

  if (-not (Test-MediaReachable -Session $session -Diagnostics $diagnostics)) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Exchange media not reachable from the target at $ExchangeMediaPath. Stage Setup.exe on an SMB share or pre-mounted ISO reachable from the target; it is never transferred over WinRM."
  }

  if ($Mode -eq 'Full') {
    $ucmaPath = Join-Path (Split-Path $ExchangeMediaPath -Parent) 'UCMARedist\UcmaRuntimeSetup.exe'
    $pathExistsScript = {
      param($path)
      Test-Path -LiteralPath $path -PathType Leaf
    }
    $ucmaPresent = Invoke-Command -Session $session -ArgumentList $ucmaPath -ScriptBlock $pathExistsScript -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target 'UCMA 4.0' -Source 'ExchangePrereq' -Action 'Validate' -Status $(if ($ucmaPresent) { 'Completed' } else { 'Warn' }) -Detail $(if ($ucmaPresent) { 'UCMA staging present in media.' } else { "UCMA staging not found at $ucmaPath." })
  }

  $featureInventoryScript = {
    param($featureNames)
    $missing = @(Get-WindowsFeature -Name $featureNames | Where-Object { -not $_.Installed } | ForEach-Object { $_.Name })
    [pscustomobject]@{
      Missing = $missing
    }
  }
  $featureState = Invoke-Command -Session $session -ArgumentList (, $exchangeProfile.WindowsFeatures) -ScriptBlock $featureInventoryScript -ErrorAction SilentlyContinue
  if ($featureState -and $featureState.Missing.Count -gt 0) {
    Add-OperationResult -Results $diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Validate' -Status 'Warn' -Detail ("$($featureState.Missing.Count) feature(s) not yet installed: {0}" -f ($featureState.Missing -join ', '))
  }
  else {
    Add-OperationResult -Results $diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Validate' -Status 'Completed' -Detail 'All required Windows features present.'
  }

  Remove-PSSession -Session $session -ErrorAction SilentlyContinue
  Write-PhaseEnvelope -PhaseName 'validate' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Phase: Prepare
# ---------------------------------------------------------------------------

function Invoke-PreparePhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList
  $exchangeProfile = $exchangeProfiles[$ExchangeVersion]

  if ($Mode -eq 'ManagementTools') {
    Add-OperationResult -Results $diagnostics -Target 'Preparation' -Source 'ExchangeSetup' -Action 'Prepare' -Status 'Skipped' -Detail 'ManagementTools mode does not run schema/AD/domain preparation stages.'
    return Write-PhaseEnvelope -PhaseName 'prepare' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  if (-not $ConfirmSchemaExtension) {
    Add-OperationResult -Results $diagnostics -Target '/PrepareSchema' -Source 'ExchangeSetup' -Action 'Prepare' -Status 'Failed' -Detail 'Confirmation switch missing.'
    return Write-PhaseEnvelope -PhaseName 'prepare' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText '/PrepareSchema extends the schema for the entire forest and is effectively irreversible. Pass -ConfirmSchemaExtension to proceed, after verifying the target forest and Exchange version.'
  }

  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'prepare' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  if (-not (Test-MediaReachable -Session $session -Diagnostics $diagnostics)) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'prepare' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Exchange media not reachable from the target at $ExchangeMediaPath."
  }

  $licenseFlag = $exchangeProfile.LicenseFlag

  $schemaStep = Invoke-SetupExeStep -Session $session -SetupExe $ExchangeMediaPath -Arguments @('/PrepareSchema', $licenseFlag) -StepName '/PrepareSchema' -Diagnostics $diagnostics
  if (-not $schemaStep.Success) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'prepare' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText '/PrepareSchema failed. Review the log excerpt in diagnostics (or run CSS-Exchange SetupLogReviewer on the target for a deeper post-mortem). Setup.exe is resumable: fix the cause and re-run.'
  }

  $adStep = Invoke-SetupExeStep -Session $session -SetupExe $ExchangeMediaPath -Arguments @('/PrepareAD', "/OrganizationName:`"$OrganizationName`"", $licenseFlag) -StepName '/PrepareAD' -Diagnostics $diagnostics
  if (-not $adStep.Success) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'prepare' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText '/PrepareAD failed. Review the log excerpt in diagnostics (or run CSS-Exchange SetupLogReviewer on the target for a deeper post-mortem). Setup.exe is resumable: fix the cause and re-run.'
  }

  $domainStep = Invoke-SetupExeStep -Session $session -SetupExe $ExchangeMediaPath -Arguments @("/PrepareDomain:$DomainName", $licenseFlag) -StepName '/PrepareDomain' -Diagnostics $diagnostics
  Remove-PSSession -Session $session -ErrorAction SilentlyContinue

  if (-not $domainStep.Success) {
    return Write-PhaseEnvelope -PhaseName 'prepare' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText '/PrepareDomain failed. Review the log excerpt in diagnostics (or run CSS-Exchange SetupLogReviewer on the target for a deeper post-mortem). Setup.exe is resumable: fix the cause and re-run.'
  }

  Write-PhaseEnvelope -PhaseName 'prepare' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Phase: Install
# ---------------------------------------------------------------------------

function Invoke-InstallPhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList
  $exchangeProfile = $exchangeProfiles[$ExchangeVersion]
  $licenseFlag = $exchangeProfile.LicenseFlag
  $rebootRequired = $false

  if (-not $PSCmdlet.ShouldProcess($Server, "Install Exchange $ExchangeVersion ($Mode) prerequisites and role")) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Exchange' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf - install phase skipped.'
    return Write-PhaseEnvelope -PhaseName 'install' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  if (Test-TargetPendingReboot -Session $session -Diagnostics $diagnostics) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'A reboot is pending on the target. Reboot it and re-run before installing.'
  }

  if (-not (Test-MediaReachable -Session $session -Diagnostics $diagnostics)) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Exchange media not reachable from the target at $ExchangeMediaPath."
  }

  $featureResult = Install-ExchangeFeature -Session $session -Features $exchangeProfile.WindowsFeatures -Diagnostics $diagnostics
  if (-not $featureResult.Success) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Windows feature install failed.'
  }

  # Controlled reboot when the feature install demands it; never let an
  # installer's own auto-reboot fire mid-remote-session.
  if ($featureResult.RebootRequired) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    $rebootEnvelope = Invoke-RebootTarget -RebootCredential $Credential
    $rebootEnvelope | Write-Output
    if (-not $rebootEnvelope.success) {
      $rebootRequired = $true
      return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The controlled reboot after the feature install failed. Re-run with -Phase Install -Finish to resume.'
    }
    $rebootRequired = $true
    $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
    if (-not $session) {
      return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not re-establish the WinRM session after the reboot.'
    }
    $featureInventoryScript = {
      param($featureNames)
      $missing = @(Get-WindowsFeature -Name $featureNames | Where-Object { -not $_.Installed } | ForEach-Object { $_.Name })
      [pscustomobject]@{ Missing = $missing }
    }
    $featureCheck = Invoke-Command -Session $session -ArgumentList (, $exchangeProfile.WindowsFeatures) -ScriptBlock $featureInventoryScript -ErrorAction Stop
    if ($featureCheck.Missing.Count -gt 0) {
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
      return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Features did not persist after reboot; still missing: $($featureCheck.Missing -join ', ')"
    }
    Add-OperationResult -Results $diagnostics -Target 'WindowsFeatures' -Source 'ExchangePrereq' -Action 'Install' -Status 'Completed' -Detail 'Features confirmed present after reboot.'
  }

  $prereqResult = Invoke-PrereqInstaller -Session $session -Installers $exchangeProfile.PrereqInstallers -Diagnostics $diagnostics
  if (-not $prereqResult.Success) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Prerequisite installer failed. Fix the cause and re-run with -Phase Install -Finish.'
  }

  if ($Mode -eq 'Full') {
    $ucmaPath = Join-Path (Split-Path $ExchangeMediaPath -Parent) 'UCMARedist\UcmaRuntimeSetup.exe'
    $ucmaOk = Invoke-UcmaInstall -Session $session -InstallerPath $ucmaPath -Diagnostics $diagnostics
    if (-not $ucmaOk) {
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
      return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'UCMA 4.0 install failed.'
    }
  }

  $role = $exchangeProfile.Roles[$Mode]
  $setupArgs = @('/Mode:Install', "/Roles:$role", $licenseFlag)
  if ($Mode -eq 'Full') {
    $setupArgs += "/OrganizationName:`"$OrganizationName`""
  }

  $installStep = Invoke-SetupExeStep -Session $session -SetupExe $ExchangeMediaPath -Arguments $setupArgs -StepName "/Mode:Install /Roles:$role" -Diagnostics $diagnostics
  Remove-PSSession -Session $session -ErrorAction SilentlyContinue

  if (-not $installStep.Success) {
    return Write-PhaseEnvelope -PhaseName 'install' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Exchange role install failed. Review the log excerpt in diagnostics (or run CSS-Exchange SetupLogReviewer on the target for a deeper post-mortem). Setup.exe is resumable: fix the cause and re-run with -Phase Install -Finish."
  }

  Write-PhaseEnvelope -PhaseName 'install' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -RebootRequired $rebootRequired
}

# ---------------------------------------------------------------------------
# Phase: Verify
# ---------------------------------------------------------------------------

function Invoke-VerifyPhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  $verifyScript = {
    param($mode)
    $checks = @()
    if ($mode -eq 'Full') {
      $serviceNames = @('MSExchangeADTopology', 'MSExchangeIS', 'MSExchangeTransport', 'MSExchangeFrontEndTransport', 'MSExchangeMailboxAssistants', 'MSExchangeRepl', 'MSExchangeRPC', 'MSExchangeDagMgmt', 'W3SVC')
      foreach ($name in $serviceNames) {
        $service = Get-Service -Name $name -ErrorAction SilentlyContinue
        $checks += [pscustomobject]@{
          Name = $name
          Status = if ($service) { $service.Status.ToString() } else { 'Missing' }
        }
      }
    }
    $emsOk = $false
    $serverInfo = $null
    $databaseStates = @()
    try {
      Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
      $emsOk = $true
      if ($mode -eq 'Full') {
        $serverInfo = Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction Stop | Select-Object Name, AdminDisplayVersion, ServerRole
        $databaseStates = @(Get-MailboxDatabase -Server $env:COMPUTERNAME -Status -ErrorAction SilentlyContinue | Select-Object Name, Mounted)
      }
    }
    catch {
      $emsOk = $false
    }
    [pscustomobject]@{
      Services = $checks
      EMSSnapinLoaded = $emsOk
      ServerInfo = $serverInfo
      DatabaseStates = $databaseStates
    }
  }

  $verify = $null
  try {
    $verify = Invoke-Command -Session $session -ScriptBlock $verifyScript -ArgumentList $Mode -ErrorAction Stop
  }
  catch {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Exchange' -Action 'Verify' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }
  Remove-PSSession -Session $session -ErrorAction SilentlyContinue

  $serviceOk = $true
  foreach ($serviceEntry in $verify.Services) {
    Add-OperationResult -Results $diagnostics -Target $serviceEntry.Name -Source 'ExchangeService' -Action 'Verify' -Status $(if ($serviceEntry.Status -eq 'Running') { 'Completed' } else { 'Failed' }) -Detail $serviceEntry.Status
    if ($serviceEntry.Status -ne 'Running') {
      $serviceOk = $false
    }
  }

  Add-OperationResult -Results $diagnostics -Target 'ExchangeManagementShell' -Source 'Exchange' -Action 'Verify' -Status $(if ($verify.EMSSnapinLoaded) { 'Completed' } else { 'Failed' }) -Detail $(if ($verify.EMSSnapinLoaded) { 'Management Shell available.' } else { 'Management Shell snapin failed to load.' })

  if ($Mode -eq 'Full' -and $verify.ServerInfo) {
    Add-OperationResult -Results $diagnostics -Target $verify.ServerInfo.Name -Source 'ExchangeServer' -Action 'Verify' -Status 'Completed' -Detail "Version: $($verify.ServerInfo.AdminDisplayVersion), Roles: $($verify.ServerInfo.ServerRole)"
    $mounted = @($verify.DatabaseStates | Where-Object { $_.Mounted })
    Add-OperationResult -Results $diagnostics -Target 'MailboxDatabases' -Source 'ExchangeServer' -Action 'Verify' -Status $(if ($verify.DatabaseStates.Count -gt 0) { 'Completed' } else { 'Warn' }) -Detail "$($mounted.Count) of $($verify.DatabaseStates.Count) databases mounted."
  }

  $success = $verify.EMSSnapinLoaded
  if ($Mode -eq 'Full') {
    $success = ($success -and $serviceOk -and $null -ne $verify.ServerInfo)
  }

  if (-not $success) {
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Exchange verification failed - see diagnostics for details.'
  }

  Write-PhaseEnvelope -PhaseName 'verify' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------

$phaseOrder = @('Validate', 'Prepare', 'Install', 'Verify')

if ($Phase) {
  $startIndex = [array]::IndexOf($phaseOrder, $Phase)
  if ($Finish) {
    $phasesToRun = $phaseOrder[$startIndex..($phaseOrder.Length - 1)]
  }
  else {
    $phasesToRun = @($Phase)
  }
}
else {
  if ($Finish) {
    Write-Error '-Finish is only meaningful together with -Phase (it resumes the run from that phase to the end).'
    exit 1
  }
  $phasesToRun = @($phaseOrder)
}

$licenseFlag = $exchangeProfiles[$ExchangeVersion].LicenseFlag

Write-Log -Message "Target: $Server | Domain: $DomainName | Mode: $Mode | Version: $ExchangeVersion | Phases: $($phasesToRun -join ', ')" -Color Cyan

$overallSuccess = $true

foreach ($phaseName in $phasesToRun) {
  $envelope = $null
  try {
    switch ($phaseName) {
      'Validate' {
        $envelope = Invoke-ValidatePhase
      }
      'Prepare' {
        $envelope = Invoke-PreparePhase
      }
      'Install' {
        $envelope = Invoke-InstallPhase
      }
      'Verify' {
        $envelope = Invoke-VerifyPhase
      }
    }
  }
  catch {
    $diagnostics = New-Object System.Collections.ArrayList
    $errorText = Get-RedactedError $_
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Orchestrator' -Action 'Phase' -Status 'Failed' -Detail $errorText
    $envelope = Write-PhaseEnvelope -PhaseName $phaseName -Success $false -StartedAt (Get-Date).ToUniversalTime().ToString('o') -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $errorText
  }

  if ($envelope) {
    $envelope | Write-Output
  }

  if ($envelope -and -not $envelope.success) {
    $overallSuccess = $false
    break
  }
}

if ($overallSuccess) {
  Write-Log -Message 'All requested phases completed successfully.' -Color Green
  exit 0
}

Write-Log -Message 'One or more phases failed - see the envelopes above.' -Color Red
exit 1
