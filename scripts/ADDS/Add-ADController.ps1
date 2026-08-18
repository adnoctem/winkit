#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.1.0' }

<#
.SYNOPSIS
  Adds a domain controller to an existing domain (DC migration), with optional
  FSMO role transfer and old-domain-controller demotion.

.DESCRIPTION
  Companion to Initialize-ADController.ps1: that script creates a greenfield
  forest; this one joins a fresh server into an existing domain and promotes it
  to a replica domain controller - the "migration" workflow described in the
  windowspro article (see .LINK). Orchestrated over WinRM (Invoke-Command) with
  the same phase model, controlled reboots, progress bars, and JSONL envelope
  output as the rest of the ADDS family.

  Phases:

    Validate - WinRM prerequisites (TrustedHosts, negotiated authentication),
               pending-reboot detection, DSRM password complexity, domain
               functional-level check (a Windows Server 2025 DC requires
               Windows Server 2016 domain/forest functional level or higher;
               the script fails with guidance instead of raising the level
               implicitly), and domain-role state detection.
    Join     - Domain join via Add-Computer (with optional rename through
               -NewComputerName), controlled reboot. Skipped when the target
               is already a member of the expected domain.
    Promote  - Idempotent AD-Domain-Services feature install, then
               Install-ADDSDomainController with an explicitly supplied DSRM
               password (avoids the known interactive-prompt hang/bug), a
               controlled reboot, and - unless -SkipDnsClientConfig - the
               recommended post-promotion DNS client configuration (the new DC
               points at itself via loopback, with the existing DC as
               secondary).
    Verify   - AD readiness polling (SYSVOL/NETLOGON + Get-ADDomain),
               replication health (repadmin /replsummary), the _msdcs SRV
               record, core service states, and the domain controller list
               with Global Catalog confirmation.
    FSMO     - Only with -TransferFsmoRoles: moves all five FSMO roles
               (SchemaMaster, DomainNamingMaster, PDCEmulator, RIDMaster,
               InfrastructureMaster) to the new DC and re-verifies the
               holders. SchemaMaster transfer requires Schema Admins
               membership.
    Demote   - Only with -DemoteOldController <name> and -ConfirmDemotion:
               pre-checks that the old DC holds no FSMO roles (transfer them
               first), demotes it via Uninstall-ADDSDomainController (as a
               background job with progress), reboots it, and verifies it is
               gone from the domain controller list. Explicitly out of scope:
               computer-account cleanup (Remove-ADComputer) is a manual
               follow-up.

  Without -Phase every enabled phase runs in order. -Phase runs only the named
  phase (intended for crash recovery); combine -Phase <name> -Finish to resume
  the remaining phases through the end of the run. The FSMO and Demote phases
  run only when their switches/parameters are present, unless explicitly named
  via -Phase.

  Secrets are accepted only as SecureString/PSCredential, never logged, and
  never serialized into the envelopes. The -Credential must be a domain
  account with local administrator rights on the target and the permissions
  required for the selected phases.

.PARAMETER Server
  Target server (static IP or name). The new DC.

.PARAMETER Credential
  Domain account with rights to join the domain, promote a DC, and (for the
  optional phases) transfer FSMO roles / demote the old DC.

.PARAMETER DomainName
  Fully qualified domain name of the existing domain, e.g. company.com.

.PARAMETER SafeModeAdministratorPassword
  DSRM password as a SecureString. Validated for AD complexity requirements
  client-side in the Validate phase.

.PARAMETER ExistingDomainController
  Optional name of an existing domain controller; used as the secondary DNS
  server in the post-promotion DNS client configuration.

.PARAMETER NewComputerName
  Optional new hostname applied during the domain join.

.PARAMETER SiteName
  AD site for the new DC. Defaults to Default-First-Site-Name.

.PARAMETER TransferFsmoRoles
  Enable the FSMO phase: transfer all five operation master roles to the new
  DC.

.PARAMETER DemoteOldController
  Hostname of the old domain controller to demote. Enables the Demote phase
  (which additionally requires -ConfirmDemotion).

.PARAMETER ConfirmDemotion
  Explicit acknowledgment that the old domain controller will be demoted
  (removed from the domain as a DC) and rebooted.

.PARAMETER SkipDnsClientConfig
  Do not apply the post-promotion DNS client configuration (loopback + existing
  DC).

.PARAMETER Phase
  Run only the named phase. Defaults to running all enabled phases in order.

.PARAMETER Finish
  Only meaningful together with -Phase: run the named phase, then continue
  through the remaining phases to the end of the run.

.PARAMETER OutputPath
  Optional file to append each phase envelope to as a JSON line.

.PARAMETER OperationTimeoutMinutes
  WinRM session operation timeout.

.EXAMPLE
  PS> ./Add-ADController.ps1 -Server 192.0.2.11 -Credential $domainAdmin -DomainName company.com -SafeModeAdministratorPassword $dsrm -TransferFsmoRoles -DemoteOldController dc01 -ConfirmDemotion
  Full migration: join, promote, verify, transfer FSMO roles, demote dc01.

.EXAMPLE
  PS> ./Add-ADController.ps1 -Server 192.0.2.11 -Credential $domainAdmin -DomainName company.com -SafeModeAdministratorPassword $dsrm
  Adds a replica DC to company.com without touching FSMO roles or the old DC.

.EXAMPLE
  PS> ./Add-ADController.ps1 -Server 192.0.2.11 -Credential $domainAdmin -DomainName company.com -SafeModeAdministratorPassword $dsrm -Phase Promote -Finish
  Resumes a crashed run: promotes (if not already), verifies, and runs any
  enabled optional phases.

.LINK
  https://github.com/adnoctem/winkit
  https://www.windowspro.de/wolfgang-sommergut/domain-controller-windows-server-2025-migrieren
  https://blog.andreas-schreiner.de/2017/09/14/adds-forests-domains-erstellen-konfigurieren/
  https://learn.microsoft.com/en-us/powershell/module/addsdeployment/install-addsdomaincontroller
  https://learn.microsoft.com/en-us/powershell/module/activedirectory/move-addirectoryserveroperationmasterrole
  https://www.windowspro.de/tim-buntrock/domaenen-controller-herabstufen-server-2016-2019-gui-powershell

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Carriage-return progress lines require Write-Host for in-place console updates.')]

# Parameters consumed by the phase/helper functions below (which live in their
# own scopes) are intentionally referenced only from those functions; the
# unused-parameter rule cannot see function-scope usage and needs the
# per-parameter exemptions.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'SafeModeAdministratorPassword', Justification = 'Consumed by phase functions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ExistingDomainController', Justification = 'Consumed by the Promote phase.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'NewComputerName', Justification = 'Consumed by the Join phase.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'SiteName', Justification = 'Consumed by the Promote phase.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'ConfirmDemotion', Justification = 'Consumed by the Demote phase.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'SkipDnsClientConfig', Justification = 'Consumed by the Promote phase.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'OutputPath', Justification = 'Consumed by Write-PhaseEnvelope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'OperationTimeoutMinutes', Justification = 'Consumed by New-TargetSession.')]

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [string]$Server,

  [Parameter(Mandatory = $true)]
  [pscredential]$Credential,

  [Parameter(Mandatory = $true)]
  [string]$DomainName,

  [Parameter(Mandatory = $true)]
  [securestring]$SafeModeAdministratorPassword,

  [Parameter(Mandatory = $false)]
  [string]$ExistingDomainController,

  [Parameter(Mandatory = $false)]
  [string]$NewComputerName,

  [Parameter(Mandatory = $false)]
  [string]$SiteName = 'Default-First-Site-Name',

  [Parameter(Mandatory = $false)]
  [switch]$TransferFsmoRoles,

  [Parameter(Mandatory = $false)]
  [string]$DemoteOldController,

  [Parameter(Mandatory = $false)]
  [switch]$ConfirmDemotion,

  [Parameter(Mandatory = $false)]
  [switch]$SkipDnsClientConfig,

  [Parameter(Mandatory = $false)]
  [ValidateSet('Validate', 'Join', 'Promote', 'Verify', 'FSMO', 'Demote')]
  [string]$Phase,

  [Parameter(Mandatory = $false)]
  [switch]$Finish,

  [Parameter(Mandatory = $false)]
  [string]$OutputPath,

  [Parameter(Mandatory = $false)]
  [int]$OperationTimeoutMinutes = 30
)

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

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

function Test-DsrmPasswordComplexity {
  <#
    Client-side validation of the DSRM password against AD password complexity
    requirements (at least 8 characters, three of four character classes).
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [securestring]$Password
  )

  $plain = ([System.Net.NetworkCredential]::new('', $Password)).Password

  if ($plain.Length -lt 8) {
    return $false
  }

  $classes = 0
  if ($plain -cmatch '[a-z]') { $classes++ }
  if ($plain -cmatch '[A-Z]') { $classes++ }
  if ($plain -cmatch '[0-9]') { $classes++ }
  if ($plain -cmatch '[^a-zA-Z0-9]') { $classes++ }

  return ($classes -ge 3)
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

    [string]$TargetComputer = $Server,

    [System.Collections.ArrayList]$Diagnostics
  )

  $sessionOption = New-PSSessionOption -OperationTimeout ([int]($OperationTimeoutMinutes * 60000)) -IdleTimeout ([int]($OperationTimeoutMinutes * 60000))
  $session = $null

  try {
    $session = New-PSSession -ComputerName $TargetComputer -Credential $SessionCredential -SessionOption $sessionOption -ErrorAction Stop
  }
  catch {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target $TargetComputer -Source 'WinRM' -Action 'Connect' -Status 'Failed' -Detail (Get-RedactedError $_)
    }
    return $null
  }

  if ($null -ne $Diagnostics) {
    $mechanism = $session.Runspace.ConnectionInfo.AuthenticationMechanism
    Add-OperationResult -Results $Diagnostics -Target $TargetComputer -Source 'WinRM' -Action 'Connect' -Status 'Completed' -Detail "Session established (auth: $mechanism)."
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

    [string]$TargetComputer = $Server,

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
      Add-OperationResult -Results $Diagnostics -Target $TargetComputer -Source 'PendingReboot' -Action 'Check' -Status 'Failed' -Detail (Get-RedactedError $_)
    }
    return $true
  }

  if ($null -ne $Diagnostics) {
    Add-OperationResult -Results $Diagnostics -Target $TargetComputer -Source 'PendingReboot' -Action 'Check' -Status $(if ($result.PendingReboot) { 'Failed' } else { 'Completed' }) -Detail $(if ($result.PendingReboot) { "Pending reboot detected - indicators: $($result.Indicators -join ', ')" } else { 'No pending reboot detected.' })
  }

  [bool]$result.PendingReboot
}

function Get-FunctionalLevelName {
  <#
    Maps a domainFunctionality/forestFunctionality numeric level to a display
    name. Unknown levels are reported numerically.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [int]$Level
  )

  switch ($Level) {
    2 { return 'Windows Server 2003' }
    3 { return 'Windows Server 2008' }
    4 { return 'Windows Server 2008 R2' }
    5 { return 'Windows Server 2012' }
    6 { return 'Windows Server 2012 R2' }
    7 { return 'Windows Server 2016 or higher' }
    default { return "level $Level" }
  }
}

function Test-DomainFunctionalLevel {
  <#
    Queries the domain/forest functional levels through LDAP (rootDSE) from
    inside the target session and fails when they are below Windows Server
    2016, which a Windows Server 2025 DC requires. Returns $null when the
    levels could not be determined (callers treat that as best-effort).
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [System.Collections.ArrayList]$Diagnostics
  )

  $functionalScript = {
    param($domainName)
    try {
      $rootDse = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domainName/rootDSE")
      $rootDse.RefreshCache()
      [pscustomobject]@{
        Success = $true
        DomainFunctionality = [int]$rootDse.Properties['domainFunctionality'].Value
        ForestFunctionality = [int]$rootDse.Properties['forestFunctionality'].Value
      }
    }
    catch {
      [pscustomobject]@{
        Success = $false
        Detail = $_.Exception.Message
      }
    }
  }

  try {
    $result = Invoke-Command -Session $Session -ArgumentList $DomainName -ScriptBlock $functionalScript -ErrorAction Stop
  }
  catch {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target $DomainName -Source 'FunctionalLevel' -Action 'Check' -Status 'Warn' -Detail "Functional level query failed: $(Get-RedactedError $_)"
    }
    return $null
  }

  if (-not $result.Success) {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target $DomainName -Source 'FunctionalLevel' -Action 'Check' -Status 'Warn' -Detail "Functional levels could not be determined: $($result.Detail)"
    }
    return $null
  }

  [pscustomobject]@{
    DomainLevel = $result.DomainFunctionality
    ForestLevel = $result.ForestFunctionality
    DomainName = Get-FunctionalLevelName -Level $result.DomainFunctionality
    ForestName = Get-FunctionalLevelName -Level $result.ForestFunctionality
    MeetsRequirement = (($result.DomainFunctionality -ge 7) -and ($result.ForestFunctionality -ge 7))
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

  # 1. Orchestrator-side TrustedHosts prerequisite (workgroup scenario).
  $trustedHosts = ''
  try {
    $trustedHosts = (Get-Item 'WSMan:\localhost\Client\TrustedHosts' -ErrorAction Stop).Value
  }
  catch {
    $trustedHosts = ''
  }

  $isTrusted = ($trustedHosts.Trim() -eq '*') -or (($trustedHosts -split ',') | Where-Object { $_.Trim() -eq $Server })
  if (-not $isTrusted) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Validate' -Status 'Failed' -Detail "Target not in TrustedHosts. Remediate with: Set-Item WSMan:\localhost\Client\TrustedHosts -Value $Server -Concatenate"
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'WinRM TrustedHosts prerequisite not met.'
  }
  Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Validate' -Status 'Completed' -Detail 'TrustedHosts prerequisite met.'

  # 2. Session + negotiated authentication mechanism.
  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  $authMechanism = $session.Runspace.ConnectionInfo.AuthenticationMechanism
  if ($authMechanism -eq 'Basic') {
    $scheme = $null
    try {
      $scheme = $session.Runspace.ConnectionInfo.Scheme
    }
    catch {
      $scheme = $null
    }
    if ($scheme -ne 'https') {
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
      Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Validate' -Status 'Failed' -Detail 'Session negotiated Basic authentication without HTTPS transport.'
      return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Authentication is Basic over plain HTTP. SecureString parameters (DSRM password) require Kerberos/Negotiate or an HTTPS WinRM listener.'
    }
  }

  # 3. Pending-reboot gate.
  # Remote-management channel pre-flight (non-blocking diagnostic).
  try {
    $reachability = Test-RemoteHostReachability -ComputerName $Server -ErrorAction Stop
    $channelDetail = (($reachability.Channels | ForEach-Object { "$($_.Source)=$($_.Status)" }) -join ', ')
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Reachability' -Action 'Preflight' -Status 'Completed' -Detail $channelDetail
  }
  catch {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Reachability' -Action 'Preflight' -Status 'Warn' -Detail "Reachability probe failed: $(Get-RedactedError $_)"
  }
  $pendingReboot = Test-TargetPendingReboot -Session $session -Diagnostics $diagnostics
  if ($pendingReboot) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'A reboot is pending on the target. Reboot it and re-run before proceeding.'
  }

  # 4. DSRM password complexity, validated client-side.
  if (-not (Test-DsrmPasswordComplexity -Password $SafeModeAdministratorPassword)) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target 'DSRM' -Source 'ADDS' -Action 'Validate' -Status 'Failed' -Detail 'Password does not meet AD complexity requirements.'
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The DSRM password fails AD complexity requirements (at least 8 characters, three of four character classes).'
  }
  Add-OperationResult -Results $diagnostics -Target 'DSRM' -Source 'ADDS' -Action 'Validate' -Status 'Completed' -Detail 'DSRM password meets complexity requirements.'

  # 5. Domain-role state: already a DC -> short-circuit; partial state -> fail.
  $stateScript = {
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $ntds = Get-Service -Name NTDS -ErrorAction SilentlyContinue
    [pscustomobject]@{
      DomainRole = $computerSystem.DomainRole
      Domain = $computerSystem.Domain
      NTDSInstalled = ($null -ne $ntds)
    }
  }

  $state = $null
  try {
    $state = Invoke-Command -Session $session -ScriptBlock $stateScript -ErrorAction Stop
  }
  catch {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Validate' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  if ($state.DomainRole -ge 4) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    if ($state.Domain -eq $DomainName) {
      Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Validate' -Status 'Completed' -Detail "Target is already a domain controller of $DomainName; remaining checks skipped."
      return Write-PhaseEnvelope -PhaseName 'validate' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
    }
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Validate' -Status 'Failed' -Detail "Target is already a domain controller in $($state.Domain), expected $DomainName."
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Target is already a domain controller in $($state.Domain); expected $DomainName."
  }

  if ($state.NTDSInstalled) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Validate' -Status 'Failed' -Detail 'NTDS service present but machine is not a domain controller - partial promotion state.'
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The target is in a partially promoted state (NTDS present, not a DC). Do not re-attempt automatically; rebuild from a clean snapshot, then re-run.'
  }

  # 6. Domain/forest functional levels (best effort here; hard gate in Promote).
  $functionalLevel = Test-DomainFunctionalLevel -Session $session -Diagnostics $diagnostics
  if ($functionalLevel) {
    Add-OperationResult -Results $diagnostics -Target $DomainName -Source 'FunctionalLevel' -Action 'Check' -Status $(if ($functionalLevel.MeetsRequirement) { 'Completed' } else { 'Failed' }) -Detail "Domain: $($functionalLevel.DomainName) ($($functionalLevel.DomainLevel)), Forest: $($functionalLevel.ForestName) ($($functionalLevel.ForestLevel))"
    if (-not $functionalLevel.MeetsRequirement) {
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
      return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "The domain/forest functional level must be at least Windows Server 2016 for a Windows Server 2025 DC. Raise it explicitly (Set-ADDomainMode / Set-ADForestMode) and re-run."
    }
  }

  Remove-PSSession -Session $session -ErrorAction SilentlyContinue
  Write-PhaseEnvelope -PhaseName 'validate' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Phase: Join
# ---------------------------------------------------------------------------

function Invoke-JoinPhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'join' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  $stateScript = {
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    [pscustomobject]@{
      DomainRole = $computerSystem.DomainRole
      Domain = $computerSystem.Domain
    }
  }

  $state = $null
  try {
    $state = Invoke-Command -Session $session -ScriptBlock $stateScript -ErrorAction Stop
  }
  catch {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Join' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'join' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  if ($state.DomainRole -ge 4) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Join' -Status 'Failed' -Detail 'Target is already a domain controller; a DC cannot be domain-joined as a member.'
    return Write-PhaseEnvelope -PhaseName 'join' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The target is already a domain controller.'
  }

  if ($state.DomainRole -in @(1, 3) -and $state.Domain -eq $DomainName) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Join' -Status 'Skipped' -Detail "Already a member of $DomainName; join skipped."
    return Write-PhaseEnvelope -PhaseName 'join' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  if ($state.DomainRole -in @(1, 3)) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Join' -Status 'Failed' -Detail "Already a member of $($state.Domain), expected $DomainName."
    return Write-PhaseEnvelope -PhaseName 'join' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "The target is already joined to $($state.Domain). Unjoin it manually before re-running."
  }

  if (-not $PSCmdlet.ShouldProcess($Server, "Join domain $DomainName")) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Join' -Status 'Skipped' -Detail 'WhatIf - domain join skipped.'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'join' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  $joinScript = {
    param($domainName, [pscredential]$joinCredential, $newName)
    $ConfirmPreference = 'None'
    $arguments = @{
      DomainName = $domainName
      Credential = $joinCredential
      Restart = $false
      Force = $true
      Confirm = $false
    }
    if ($newName) {
      $arguments.NewName = $newName
    }
    Add-Computer @arguments
  }

  try {
    $null = Invoke-Command -Session $session -ScriptBlock $joinScript -ArgumentList $DomainName, $Credential, $NewComputerName -ErrorAction Stop
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Join' -Status 'Completed' -Detail "Joined to $DomainName; reboot required."
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'join' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -RebootRequired $true
  }
  catch {
    $message = Get-RedactedError $_
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Join' -Status 'Failed' -Detail $message
    return Write-PhaseEnvelope -PhaseName 'join' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $message
  }
}

# ---------------------------------------------------------------------------
# Phase: Promote
# ---------------------------------------------------------------------------

function Invoke-PromotePhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  # Re-check pending reboot: a partial run may have left a stale flag.
  if (Test-TargetPendingReboot -Session $session -Diagnostics $diagnostics) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'A reboot is pending on the target. Reboot it and re-run before promoting.'
  }

  # Hard functional-level gate (post-join the target must reach the domain).
  $functionalLevel = Test-DomainFunctionalLevel -Session $session -Diagnostics $diagnostics
  if ($functionalLevel -and -not $functionalLevel.MeetsRequirement) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "The domain/forest functional level must be at least Windows Server 2016 for a Windows Server 2025 DC. Raise it explicitly (Set-ADDomainMode / Set-ADForestMode) and re-run."
  }

  # State detection: never re-attempt promotion against an existing DC.
  $stateScript = {
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $ntds = Get-Service -Name NTDS -ErrorAction SilentlyContinue
    [pscustomobject]@{
      DomainRole = $computerSystem.DomainRole
      Domain = $computerSystem.Domain
      NTDSInstalled = ($null -ne $ntds)
    }
  }

  $state = $null
  try {
    $state = Invoke-Command -Session $session -ScriptBlock $stateScript -ErrorAction Stop
  }
  catch {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  if ($state.DomainRole -ge 4) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    if ($state.Domain -eq $DomainName) {
      Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Skipped' -Detail "Already promoted to $DomainName; no-op."
      return Write-PhaseEnvelope -PhaseName 'promote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
    }
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Failed' -Detail "Already a domain controller in $($state.Domain), expected $DomainName."
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Target is already a domain controller in $($state.Domain); expected $DomainName."
  }

  if ($state.NTDSInstalled) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Failed' -Detail 'NTDS service present but machine is not a domain controller - partial promotion state.'
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The target is in a partially promoted state (NTDS present, not a DC). Do not re-attempt automatically; rebuild from a clean snapshot, then re-run.'
  }

  # Capture the current primary DNS server for the post-promotion DNS client
  # configuration (the existing DC is used as secondary).
  $currentPrimaryDns = $null
  try {
    $dnsProbeScript = {
      $adapter = Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
      if (-not $adapter) {
        return $null
      }
      $servers = @(Get-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ServerAddresses)
      if ($servers.Count -gt 0) { $servers[0] } else { $null }
    }
    $currentPrimaryDns = Invoke-Command -Session $session -ScriptBlock $dnsProbeScript -ErrorAction SilentlyContinue
  }
  catch {
    $currentPrimaryDns = $null
  }

  # Idempotent feature install (guarded).
  if ($PSCmdlet.ShouldProcess($Server, 'Install the AD-Domain-Services feature if missing')) {
    $featureScript = {
      $feature = Get-WindowsFeature -Name AD-Domain-Services
      $installed = [bool]$feature.Installed
      $rebootRequired = $false
      if (-not $installed) {
        $installResult = Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools -Confirm:$false
        $installed = [bool]$installResult.Success
        $rebootRequired = [bool]$installResult.RestartNeeded
      }
      [pscustomobject]@{
        Installed = $installed
        RebootRequired = $rebootRequired
      }
    }

    try {
      $featureState = Invoke-Command -Session $session -ScriptBlock $featureScript -ErrorAction Stop
      Add-OperationResult -Results $diagnostics -Target 'AD-Domain-Services' -Source 'WindowsFeature' -Action 'Install' -Status $(if ($featureState.Installed) { 'Completed' } else { 'Failed' }) -Detail "installed: $($featureState.Installed), rebootRequiredByFeature: $($featureState.RebootRequired)"
      if (-not $featureState.Installed -or $featureState.RebootRequired) {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
        return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The AD-Domain-Services feature install failed or requires a reboot. Reboot the target and re-run.'
      }
    }
    catch {
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
      Add-OperationResult -Results $diagnostics -Target 'AD-Domain-Services' -Source 'WindowsFeature' -Action 'Install' -Status 'Failed' -Detail (Get-RedactedError $_)
      return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
    }
  }
  else {
    Add-OperationResult -Results $diagnostics -Target 'AD-Domain-Services' -Source 'WindowsFeature' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf - feature install skipped.'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  if (-not $PSCmdlet.ShouldProcess($Server, "Promote to domain controller of $DomainName")) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Skipped' -Detail 'WhatIf - promotion skipped.'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  $promoteScript = {
    param($domainName, $dsrmSecret, $siteName)
    $ConfirmPreference = 'None'
    $ErrorActionPreference = 'Stop'
    Import-Module ADDSDeployment -ErrorAction Stop
    Install-ADDSDomainController -DomainName $domainName -SafeModeAdministratorPassword $dsrmSecret -SiteName $siteName -NoRebootOnCompletion:$true -Confirm:$false -Force
  }

  $promoteJob = $null
  try {
    $promoteJob = Invoke-Command -Session $session -ScriptBlock $promoteScript -ArgumentList $DomainName, $SafeModeAdministratorPassword, $SiteName -AsJob -ErrorAction Stop
  }
  catch {
    $message = Get-RedactedError $_
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Failed' -Detail $message
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $message
  }

  $promoteStartedAt = Get-Date
  $promoteDeadline = $promoteStartedAt.AddMinutes($OperationTimeoutMinutes)
  $promoteTimedOut = $false
  while ($promoteJob.State -notin @('Completed', 'Failed')) {
    if ((Get-Date) -gt $promoteDeadline) {
      $promoteTimedOut = $true
      break
    }
    $elapsedTime = (Get-Date) - $promoteStartedAt
    $elapsedSeconds = [int]$elapsedTime.TotalSeconds
    $promotePct = [int][math]::Min(($elapsedSeconds / ($OperationTimeoutMinutes * 60)) * 100, 100)
    Write-Progress -Activity 'Promoting domain controller' -Status "elapsed: $elapsedSeconds s (watchdog $OperationTimeoutMinutes min), job state: $($promoteJob.State)" -PercentComplete $promotePct
    Write-Host ("`rPromoting ({0}s elapsed): {1,-40}" -f $elapsedSeconds, $Server) -NoNewline -ForegroundColor Cyan
    Start-Sleep -Seconds 10
  }
  Write-Progress -Activity 'Promoting domain controller' -Completed
  $promoteTotalSeconds = [int]((Get-Date) - $promoteStartedAt).TotalSeconds
  Write-Host ("`rPromotion finished ({0}s).{1}" -f $promoteTotalSeconds, (' ' * 40)) -ForegroundColor Cyan

  if ($promoteTimedOut) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Remove-Job -Job $promoteJob -Force -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Unknown' -Detail "Promotion did not complete within the $OperationTimeoutMinutes minute watchdog - state unknown."
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "The promotion did not report back within $OperationTimeoutMinutes minutes. It may have completed or partially completed; detect state with -Phase Validate and resume with -Phase Promote -Finish."
  }

  try {
    $null = Receive-Job -Job $promoteJob -ErrorAction Stop
  }
  catch {
    $message = Get-RedactedError $_
    Remove-Job -Job $promoteJob -Force -ErrorAction SilentlyContinue
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Failed' -Detail $message
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $message
  }
  Remove-Job -Job $promoteJob -Force -ErrorAction SilentlyContinue
  Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Completed' -Detail 'Domain controller promotion completed; reboot is required to finalize.'

  # Post-promotion DNS client configuration (article recommendation): self via
  # loopback, existing DC as secondary.
  if (-not $SkipDnsClientConfig) {
    try {
      $secondaryDns = if ($ExistingDomainController) { $ExistingDomainController } else { $currentPrimaryDns }
      $dnsConfigScript = {
        param($secondary)
        $adapter = Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1
        if (-not $adapter) {
          return [pscustomobject]@{ Success = $false; Detail = 'No active physical adapter found.' }
        }
        $servers = @('127.0.0.1')
        if ($secondary) {
          $servers += $secondary
        }
        Set-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ServerAddresses $servers -ErrorAction Stop
        [pscustomobject]@{ Success = $true; Detail = "DNS client set to $($servers -join ', ')." }
      }
      $dnsResult = Invoke-Command -Session $session -ScriptBlock $dnsConfigScript -ArgumentList $secondaryDns -ErrorAction Stop
      Add-OperationResult -Results $diagnostics -Target 'DNS Client' -Source 'DNS' -Action 'Configure' -Status $(if ($dnsResult.Success) { 'Completed' } else { 'Warn' }) -Detail $(if ($dnsResult.Success) { $dnsResult.Detail } else { $dnsResult.Detail })
    }
    catch {
      Add-OperationResult -Results $diagnostics -Target 'DNS Client' -Source 'DNS' -Action 'Configure' -Status 'Warn' -Detail "DNS client configuration failed: $(Get-RedactedError $_)"
    }
  }

  Remove-PSSession -Session $session -ErrorAction SilentlyContinue
  Write-PhaseEnvelope -PhaseName 'promote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -RebootRequired $true
}

# ---------------------------------------------------------------------------
# Phase: Verify
# ---------------------------------------------------------------------------

function Invoke-VerifyPhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  # Kerberos time-skew check first.
  try {
    $remoteTime = Invoke-Command -ComputerName $Server -Credential $Credential -ScriptBlock { (Get-Date).ToUniversalTime() } -ErrorAction Stop
  }
  catch {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'Time' -Action 'Check' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not query the target clock. Is the target up, and is the domain credential valid?'
  }

  $skewMinutes = [math]::Abs(((Get-Date).ToUniversalTime() - $remoteTime).TotalMinutes)
  Add-OperationResult -Results $diagnostics -Target $Server -Source 'Time' -Action 'Check' -Status $(if ($skewMinutes -gt 5) { 'Failed' } else { 'Completed' }) -Detail "Clock skew between orchestrator and target: $([math]::Round($skewMinutes, 1)) minutes."
  if ($skewMinutes -gt 5) {
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Clock skew of $([math]::Round($skewMinutes, 1)) minutes exceeds the Kerberos tolerance. Synchronize the target clock and re-run."
  }

  $verifySession = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $verifySession) {
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Authenticated session with the domain credential failed. Verify the credential is valid for $DomainName."
  }
  Add-OperationResult -Results $diagnostics -Target $Server -Source 'AD' -Action 'Authenticate' -Status 'Completed' -Detail 'Authenticated session established.'

  # AD-specific readiness poll (SYSVOL/NETLOGON + Get-ADDomain on localhost).
  $readinessScript = {
    $sysvol = Test-Path -LiteralPath '\\localhost\SYSVOL'
    $netlogon = Test-Path -LiteralPath '\\localhost\NETLOGON'
    $adOk = $false
    try {
      $null = Get-ADDomain -Server localhost -ErrorAction Stop
      $adOk = $true
    }
    catch {
      $adOk = $false
    }
    [pscustomobject]@{
      Sysvol = $sysvol
      Netlogon = $netlogon
      AD = $adOk
    }
  }

  $pollStarted = Get-Date
  $deadline = $pollStarted.AddMinutes($OperationTimeoutMinutes)
  $attempts = 0
  $ready = $false
  while (-not $ready -and ((Get-Date) -lt $deadline)) {
    $attempts++
    try {
      $check = Invoke-Command -Session $verifySession -ScriptBlock $readinessScript -ErrorAction Stop
      $ready = ($check.Sysvol -and $check.Netlogon -and $check.AD)
    }
    catch {
      $ready = $false
    }
    $elapsedTime = (Get-Date) - $pollStarted
    $elapsedSeconds = [int]$elapsedTime.TotalSeconds
    $readinessPct = [int][math]::Min(($elapsedSeconds / ($OperationTimeoutMinutes * 60)) * 100, 100)
    Write-Progress -Activity 'Waiting for AD readiness' -Status "attempt $attempts, $elapsedSeconds s / $OperationTimeoutMinutes min" -PercentComplete $readinessPct
    Write-Host ("`rWaiting for AD readiness ({0}s, attempt {1}): {2,-40}" -f $elapsedSeconds, $attempts, $Server) -NoNewline -ForegroundColor Cyan
    if (-not $ready) {
      Start-Sleep -Seconds 10
    }
  }
  Write-Progress -Activity 'Waiting for AD readiness' -Completed
  $readyAfterSeconds = [int]((Get-Date) - $pollStarted).TotalSeconds
  Write-Host ("`rAD readiness poll finished ({0}s).{1}" -f $readyAfterSeconds, (' ' * 40)) -ForegroundColor Cyan

  Add-OperationResult -Results $diagnostics -Target $Server -Source 'AD' -Action 'Readiness' -Status $(if ($ready) { 'Completed' } else { 'Failed' }) -Detail "attempts: $attempts, readyAfterSeconds: $readyAfterSeconds"

  if (-not $ready) {
    Remove-PSSession -Session $verifySession -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "AD was not ready after $attempts attempts ($readyAfterSeconds seconds). Check NTDS/Netlogon/DNS service state on the target."
  }

  # Deep checks: services, SRV record, replication, DC list + GC.
  $deepCheckScript = {
    param($domainName)
    $checks = @()

    $serviceNames = @('NTDS', 'Netlogon', 'DNS', 'KDC', 'DFSR', 'W32Time')
    foreach ($name in $serviceNames) {
      $service = Get-Service -Name $name -ErrorAction SilentlyContinue
      $status = if ($service) { $service.Status.ToString() } else { 'Missing' }
      $checks += [pscustomobject]@{
        Category = 'Service'
        Name = $name
        Status = $(if ($status -eq 'Running') { 'Pass' } else { 'Fail' })
        Detail = $status
      }
    }

    $srvRecords = @()
    try {
      $srvRecords = @(Resolve-DnsName -Name "_ldap._tcp.dc._msdcs.$domainName" -Type SRV -ErrorAction Stop)
    }
    catch {
      $srvRecords = @()
    }
    $checks += [pscustomobject]@{
      Category = 'DNS'
      Name = 'SRV'
      Status = $(if ($srvRecords.Count -gt 0) { 'Pass' } else { 'Fail' })
      Detail = "$($srvRecords.Count) _ldap._tcp.dc._msdcs SRV record(s) resolved."
    }

    $replOutput = ''
    $replFailed = $true
    try {
      $replOutput = ((& repadmin /replsummary 2>&1) | Out-String)
      $replFailed = ($replOutput -match 'FAIL')
    }
    catch {
      $replFailed = $true
    }
    $checks += [pscustomobject]@{
      Category = 'Replication'
      Name = 'repadmin'
      Status = $(if (-not $replFailed) { 'Pass' } else { 'Warn' })
      Detail = $(if (-not $replFailed) { 'repadmin /replsummary shows no failures.' } else { 'repadmin /replsummary reported failures - investigate.' })
    }

    $dcs = @(Get-ADDomainController -Filter * -ErrorAction SilentlyContinue)
    $thisDc = Get-ADDomainController -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue
    $checks += [pscustomobject]@{
      Category = 'AD'
      Name = 'DomainControllers'
      Status = 'Pass'
      Detail = ($dcs | ForEach-Object { $_.Name }) -join ', '
    }
    $checks += [pscustomobject]@{
      Category = 'AD'
      Name = 'ThisDC'
      Status = $(if ($thisDc) { 'Pass' } else { 'Fail' })
      Detail = $(if ($thisDc) { "Registered as DC; Global Catalog: $($thisDc.IsGlobalCatalog)" } else { 'Not found in the domain controller list.' })
    }

    [pscustomobject]@{
      Checks = $checks
      DcCount = $dcs.Count
      GlobalCatalog = if ($thisDc) { [bool]$thisDc.IsGlobalCatalog } else { $false }
    }
  }

  try {
    $deepCheck = Invoke-Command -Session $verifySession -ScriptBlock $deepCheckScript -ArgumentList $DomainName -ErrorAction Stop
    Remove-PSSession -Session $verifySession -ErrorAction SilentlyContinue
  }
  catch {
    Remove-PSSession -Session $verifySession -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'AD' -Action 'Verify' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  $failedChecks = @($deepCheck.Checks | Where-Object { $_.Status -eq 'Fail' })
  foreach ($entry in $deepCheck.Checks) {
    Add-OperationResult -Results $diagnostics -Target $entry.Name -Source $entry.Category -Action 'Verify' -Status $(if ($entry.Status -eq 'Pass') { 'Completed' } else { 'Warn' }) -Detail $entry.Detail
  }

  if ($failedChecks.Count -gt 0) {
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Verification failed on $($failedChecks.Count) check(s) - see diagnostics."
  }

  Write-PhaseEnvelope -PhaseName 'verify' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Phase: FSMO
# ---------------------------------------------------------------------------

function Invoke-FsmoPhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'fsmo' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not establish a WinRM session to the target.'
  }

  if (-not $PSCmdlet.ShouldProcess($Server, 'Transfer all five FSMO roles to the new DC')) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'FSMO' -Status 'Skipped' -Detail 'WhatIf - FSMO transfer skipped.'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'fsmo' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  $fsmoScript = {
    $forest = Get-ADForest -Server localhost -ErrorAction Stop
    $domain = Get-ADDomain -Server localhost -ErrorAction Stop
    $before = [ordered]@{
      SchemaMaster = $forest.SchemaMaster
      DomainNamingMaster = $forest.DomainNamingMaster
      PDCEmulator = $domain.PDCEmulator
      RIDMaster = $domain.RIDMaster
      InfrastructureMaster = $domain.InfrastructureMaster
    }

    $roles = @('SchemaMaster', 'DomainNamingMaster', 'PDCEmulator', 'RIDMaster', 'InfrastructureMaster')
    Move-ADDirectoryServerOperationMasterRole -Identity $env:COMPUTERNAME -OperationMasterRole $roles -Confirm:$false -Force -ErrorAction Stop

    $forestAfter = Get-ADForest -Server localhost -ErrorAction Stop
    $domainAfter = Get-ADDomain -Server localhost -ErrorAction Stop
    $after = [ordered]@{
      SchemaMaster = $forestAfter.SchemaMaster
      DomainNamingMaster = $forestAfter.DomainNamingMaster
      PDCEmulator = $domainAfter.PDCEmulator
      RIDMaster = $domainAfter.RIDMaster
      InfrastructureMaster = $domainAfter.InfrastructureMaster
    }

    [pscustomobject]@{
      Before = $before
      After = $after
    }
  }

  try {
    $result = Invoke-Command -Session $session -ScriptBlock $fsmoScript -ErrorAction Stop
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
  }
  catch {
    $message = Get-RedactedError $_
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'FSMO' -Status 'Failed' -Detail $message
    return Write-PhaseEnvelope -PhaseName 'fsmo' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $message
  }

  $allTransferred = $true
  foreach ($role in $result.Before.Keys) {
    $beforeHolder = $result.Before[$role]
    $afterHolder = $result.After[$role]
    $shortName = ($afterHolder -split '\.')[0]
    $isLocal = ($shortName -eq $env:COMPUTERNAME)
    if (-not $isLocal) {
      $allTransferred = $false
    }
    Add-OperationResult -Results $diagnostics -Target $role -Source 'ADDS' -Action 'FSMO' -Status $(if ($isLocal) { 'Completed' } else { 'Warn' }) -Detail "before: $beforeHolder, after: $afterHolder"
  }

  if (-not $allTransferred) {
    return Write-PhaseEnvelope -PhaseName 'fsmo' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Not all FSMO roles are held by the new DC after the transfer - see diagnostics.'
  }

  Write-PhaseEnvelope -PhaseName 'fsmo' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Phase: Demote
# ---------------------------------------------------------------------------

function Invoke-DemotePhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  if (-not $ConfirmDemotion) {
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Failed' -Detail 'Confirmation switch missing.'
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Demotion of $DemoteOldController is destructive and requires the explicit -ConfirmDemotion switch. Re-run with it after verifying the target."
  }

  $session = New-TargetSession -SessionCredential $Credential -TargetComputer $DemoteOldController -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Could not establish a WinRM session to the old controller $DemoteOldController."
  }

  $precheckScript = {
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    $domain = $null
    $forest = $null
    $holders = @{}
    try {
      $domain = Get-ADDomain -Server localhost -ErrorAction Stop
      $forest = Get-ADForest -Server localhost -ErrorAction Stop
      $holders = @{
        SchemaMaster = $forest.SchemaMaster
        DomainNamingMaster = $forest.DomainNamingMaster
        PDCEmulator = $domain.PDCEmulator
        RIDMaster = $domain.RIDMaster
        InfrastructureMaster = $domain.InfrastructureMaster
      }
    }
    catch {
      $holders = @{}
    }
    [pscustomobject]@{
      DomainRole = $computerSystem.DomainRole
      ComputerName = $env:COMPUTERNAME
      Holders = $holders
    }
  }

  $precheck = $null
  try {
    $precheck = Invoke-Command -Session $session -ScriptBlock $precheckScript -ErrorAction Stop
  }
  catch {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  if ($precheck.DomainRole -lt 4) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Skipped' -Detail 'Target is not a domain controller - nothing to demote.'
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  if ($precheck.ComputerName -eq $Server) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Failed' -Detail 'The demotion target is the same server as -Server.'
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "The demotion target $DemoteOldController is the new DC itself. Pass the hostname of the OLD controller."
  }

  $holdsRoles = $false
  foreach ($role in $precheck.Holders.Keys) {
    $holderShort = ($precheck.Holders[$role] -split '\.')[0]
    if ($holderShort -eq $precheck.ComputerName) {
      $holdsRoles = $true
      break
    }
  }
  if ($holdsRoles) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Failed' -Detail 'The old controller still holds FSMO roles.'
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "$DemoteOldController still holds FSMO roles. Run the FSMO phase first (-TransferFsmoRoles), then re-run the Demote phase."
  }

  if (-not $PSCmdlet.ShouldProcess($DemoteOldController, 'Demote (uninstall AD DS) and reboot')) {
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Skipped' -Detail 'WhatIf - demotion skipped.'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  $demoteScript = {
    $ConfirmPreference = 'None'
    $ErrorActionPreference = 'Stop'
    Import-Module ADDSDeployment -ErrorAction Stop
    Uninstall-ADDSDomainController -RemoveApplicationPartitions:$false -NoRebootOnCompletion:$true -Confirm:$false -Force
  }

  $demoteJob = $null
  try {
    $demoteJob = Invoke-Command -Session $session -ScriptBlock $demoteScript -AsJob -ErrorAction Stop
  }
  catch {
    $message = Get-RedactedError $_
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Failed' -Detail $message
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $message
  }

  $demoteStartedAt = Get-Date
  $demoteDeadline = $demoteStartedAt.AddMinutes($OperationTimeoutMinutes)
  $demoteTimedOut = $false
  while ($demoteJob.State -notin @('Completed', 'Failed')) {
    if ((Get-Date) -gt $demoteDeadline) {
      $demoteTimedOut = $true
      break
    }
    $elapsedTime = (Get-Date) - $demoteStartedAt
    $elapsedSeconds = [int]$elapsedTime.TotalSeconds
    $demotePct = [int][math]::Min(($elapsedSeconds / ($OperationTimeoutMinutes * 60)) * 100, 100)
    Write-Progress -Activity 'Demoting old domain controller' -Status "elapsed: $elapsedSeconds s (watchdog $OperationTimeoutMinutes min), job state: $($demoteJob.State)" -PercentComplete $demotePct
    Write-Host ("`rDemoting ({0}s elapsed): {1,-40}" -f $elapsedSeconds, $DemoteOldController) -NoNewline -ForegroundColor Cyan
    Start-Sleep -Seconds 10
  }
  Write-Progress -Activity 'Demoting old domain controller' -Completed
  $demoteTotalSeconds = [int]((Get-Date) - $demoteStartedAt).TotalSeconds
  Write-Host ("`rDemotion finished ({0}s).{1}" -f $demoteTotalSeconds, (' ' * 40)) -ForegroundColor Cyan

  if ($demoteTimedOut) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Remove-Job -Job $demoteJob -Force -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Unknown' -Detail "Demotion did not complete within the $OperationTimeoutMinutes minute watchdog - state unknown."
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "The demotion did not report back within $OperationTimeoutMinutes minutes. Detect state and re-run with -Phase Demote."
  }

  try {
    $null = Receive-Job -Job $demoteJob -ErrorAction Stop
  }
  catch {
    $message = Get-RedactedError $_
    Remove-Job -Job $demoteJob -Force -ErrorAction SilentlyContinue
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Failed' -Detail $message
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $message
  }
  Remove-Job -Job $demoteJob -Force -ErrorAction SilentlyContinue
  Remove-PSSession -Session $session -ErrorAction SilentlyContinue
  Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Completed' -Detail 'Demotion completed; reboot of the old controller is required.'

  # Controlled reboot of the OLD controller.
  $rebootStartedAt = Get-Date
  $rebootDeadline = $rebootStartedAt.AddSeconds(900)
  $winrmBack = $false
  try {
    Restart-Computer -ComputerName $DemoteOldController -Credential $Credential -Force -ErrorAction Stop
    while (-not $winrmBack -and (Get-Date) -lt $rebootDeadline) {
      try {
        $null = Invoke-Command -ComputerName $DemoteOldController -Credential $Credential -ScriptBlock { $true } -ErrorAction Stop
        $winrmBack = $true
      }
      catch {
        $winrmBack = $false
      }
      $elapsedTime = (Get-Date) - $rebootStartedAt
      $elapsedSeconds = [int]$elapsedTime.TotalSeconds
      $rebootPct = [int][math]::Min(($elapsedSeconds / 900) * 100, 100)
      Write-Progress -Activity 'Rebooting old domain controller' -Status "waiting for WinRM ($elapsedSeconds s / 900 s)" -PercentComplete $rebootPct
      Write-Host ("`rRebooting old DC ({0}s / 900s): {1,-40}" -f $elapsedSeconds, $DemoteOldController) -NoNewline -ForegroundColor Cyan
      if (-not $winrmBack) {
        Start-Sleep -Seconds 5
      }
    }
    Write-Progress -Activity 'Rebooting old domain controller' -Completed
    Write-Host ("`rOld DC reboot finished.{0}" -f (' ' * 40)) -ForegroundColor Cyan

    if (-not $winrmBack) {
      throw 'WinRM did not come back within 900 seconds after the old controller reboot.'
    }
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'WinRM' -Action 'Reboot' -Status 'Completed' -Detail 'WinRM confirmed back after reboot.'
  }
  catch {
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'WinRM' -Action 'Reboot' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  # Verify the old controller is gone from the domain controller list.
  $verifySession = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $verifySession) {
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not open a session to the new DC for post-demotion verification.'
  }
  try {
    $dcListScript = {
      param($oldName)
      $dcs = @(Get-ADDomainController -Filter * -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
      [pscustomobject]@{
        Dcs = $dcs
        OldGone = ($dcs -notcontains $oldName)
      }
    }
    $dcCheck = Invoke-Command -Session $verifySession -ScriptBlock $dcListScript -ArgumentList $DemoteOldController -ErrorAction Stop
    Remove-PSSession -Session $verifySession -ErrorAction SilentlyContinue
  }
  catch {
    Remove-PSSession -Session $verifySession -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status 'Warn' -Detail "Post-demotion verification failed: $(Get-RedactedError $_)"
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  Add-OperationResult -Results $diagnostics -Target $DemoteOldController -Source 'ADDS' -Action 'Demote' -Status $(if ($dcCheck.OldGone) { 'Completed' } else { 'Warn' }) -Detail $(if ($dcCheck.OldGone) { "Removed from the domain controller list; remaining DCs: $($dcCheck.Dcs -join ', ')" } else { 'Still present in the domain controller list - investigate.' })

  if (-not $dcCheck.OldGone) {
    return Write-PhaseEnvelope -PhaseName 'demote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "$DemoteOldController is still listed as a domain controller. Computer-account cleanup (Remove-ADComputer) is a manual follow-up."
  }

  Write-PhaseEnvelope -PhaseName 'demote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Transition: controlled reboot (new DC)
# ---------------------------------------------------------------------------

function Invoke-RebootTarget {
  <#
    Restarts the target and blocks until WinRM is confirmed back (authenticated
    probe), rendering live progress. "WinRM is back" is only the connectivity
    floor; callers re-verify the relevant state afterwards.
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
# Main flow
# ---------------------------------------------------------------------------

$allPhases = @('Validate', 'Join', 'Promote', 'Verify', 'FSMO', 'Demote')

if ($Phase) {
  $startIndex = [array]::IndexOf($allPhases, $Phase)
  $baseRun = if ($Finish) { $allPhases[$startIndex..($allPhases.Length - 1)] } else { @($Phase) }
}
else {
  if ($Finish) {
    Write-Error '-Finish is only meaningful together with -Phase (it resumes the run from that phase to the end).'
    exit 1
  }
  $baseRun = @($allPhases)
}

$phasesToRun = @($baseRun | Where-Object {
    if ($_ -eq 'FSMO') { $TransferFsmoRoles -or $Phase -eq 'FSMO' }
    elseif ($_ -eq 'Demote') { ([bool]$DemoteOldController) -or $Phase -eq 'Demote' }
    else { $true }
  })

Write-Log -Message "Target: $Server | Domain: $DomainName | Phases: $($phasesToRun -join ', ')" -Color Cyan

$overallSuccess = $true
$abort = $false

foreach ($phaseName in $phasesToRun) {
  if ($abort) {
    break
  }

  $envelope = $null
  try {
    switch ($phaseName) {
      'Validate' {
        $envelope = Invoke-ValidatePhase
      }
      'Join' {
        $envelope = Invoke-JoinPhase
        if ($envelope.success -and $envelope.rebootRequired) {
          $rebootEnvelope = Invoke-RebootTarget -RebootCredential $Credential
          $rebootEnvelope | Write-Output
          if (-not $rebootEnvelope.success) {
            $overallSuccess = $false
            $abort = $true
          }
        }
      }
      'Promote' {
        $envelope = Invoke-PromotePhase
        if ($envelope.success -and $envelope.rebootRequired) {
          $rebootEnvelope = Invoke-RebootTarget -RebootCredential $Credential
          $rebootEnvelope | Write-Output
          if (-not $rebootEnvelope.success) {
            $overallSuccess = $false
            $abort = $true
          }
        }
      }
      'Verify' {
        $envelope = Invoke-VerifyPhase
      }
      'FSMO' {
        $envelope = Invoke-FsmoPhase
      }
      'Demote' {
        $envelope = Invoke-DemotePhase
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
  }
}

if ($overallSuccess) {
  Write-Log -Message 'All requested phases completed successfully.' -Color Green
  exit 0
}

Write-Log -Message 'One or more phases failed - see the envelopes above.' -Color Red
exit 1
