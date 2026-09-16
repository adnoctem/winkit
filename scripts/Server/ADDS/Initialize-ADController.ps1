#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.1.0' }

<#
.SYNOPSIS
  Promotes a Windows Server to a new-forest domain controller over WinRM.

.DESCRIPTION
  Orchestrates greenfield Active Directory forest creation against a target
  server through WinRM (Invoke-Command). The actual promotion uses a single
  cmdlet, Install-ADDSForest, run with -NoRebootOnCompletion so the result is
  returned to the live session before anything restarts. The orchestrator
  controls the reboot explicitly afterwards.

  The script is organized into three phases, each of which emits exactly one
  structured JSON envelope to the output stream when it finishes:

    Validate - Orchestrator-side WinRM prerequisites (TrustedHosts, negotiated
               authentication mechanism), pending-reboot detection, idempotent
               AD-Domain-Services feature install, Test-ADDSForestInstallation,
               and client-side DSRM password complexity validation. Nothing
               directory-related changes here.
    Promote  - Domain-role state detection (already a DC -> idempotent no-op;
               half-promoted state -> fail with cleanup guidance), then
               Install-ADDSForest with -NoRebootOnCompletion.
    Verify   - Post-reboot readiness polling (SYSVOL/NETLOGON shares plus
               Get-ADDomain on localhost), preceded by a Kerberos time-skew
               check, authenticated with the inferred domain credential
               (NETBIOS\Administrator, same password as the pre-promotion
               local administrator).

  Without -Phase every phase runs in order. -Phase runs only the named phase
  (intended for crash recovery); combine -Phase <name> -Finish to resume the
  remaining phases through the end of the run.

  The Verify phase additionally applies the Windows Server 2025 domain
  controller firewall-profile mitigation (see .LINK): the NLA service can
  misclassify the DC network as Public/Unidentified after a restart, which
  makes the DC apply the public firewall profile and break domain traffic.
  On Server 2025+ targets the script sets the three mitigation registry keys,
  re-toggles the ms_tcpip6 adapter binding until the profile reports
  DomainAuthenticated, and registers an at-startup scheduled task so the fix
  survives future reboots. Failures here are recorded as warnings, never as a
  failed promotion.

  All human-facing progress goes to the console via Write-Log; the output
  stream is reserved for the phase envelopes. With -OutputPath, every envelope
  is additionally appended to that file as one JSON line (JSONL). Long waits
  (promotion, reboot, AD readiness) render a Write-Progress bar with an
  elapsed-time status line instead of silent polling.

  Secrets are accepted only as SecureString/PSCredential, never logged, and
  never serialized into the envelopes.

.PARAMETER Server
  Target server. Connect by static IP through the whole flow: the freshly
  promoted DC is becoming its own DNS server as part of the same operation,
  and its own-name resolution can be briefly unreliable right after promotion.

.PARAMETER Credential
  Local administrator credentials for the target before promotion. The same
  password is reused after promotion to authenticate as NETBIOS\Administrator.

.PARAMETER DomainName
  Fully qualified domain name of the new forest, e.g. company.com.

.PARAMETER DomainNetbiosName
  NetBIOS name of the new domain, e.g. COMPANY.

.PARAMETER SafeModeAdministratorPassword
  DSRM password as a SecureString. Validated against AD password complexity
  requirements client-side in the Validate phase; a weak password fails fast
  instead of hanging the remote promotion as an interactive prompt.

.PARAMETER Phase
  Run only the named phase. Defaults to running all phases in order.

.PARAMETER Finish
  Only meaningful together with -Phase: run the named phase, then continue
  through the remaining phases to the end of the run.

.PARAMETER OutputPath
  Optional file to append each phase envelope to as a JSON line.

.PARAMETER OperationTimeoutMinutes
  WinRM session operation timeout. Forest creation can legitimately take much
  longer than the default WinRM operation timeout, so this is raised
  explicitly on every session this script creates.

.PARAMETER VerifyTimeoutMinutes
  Upper bound for the post-reboot AD readiness poll.

.EXAMPLE
  PS> ./Initialize-ADController.ps1 -Server 192.0.2.10 -Credential $localAdmin -DomainName company.com -DomainNetbiosName COMPANY -SafeModeAdministratorPassword $dsrm -OutputPath results.jsonl
  Runs all three phases; appends one JSON envelope per phase to results.jsonl.

.EXAMPLE
  PS> ./Initialize-ADController.ps1 -Server 192.0.2.10 -Credential $localAdmin -DomainName company.com -DomainNetbiosName COMPANY -SafeModeAdministratorPassword $dsrm -Phase Promote -Finish
  Resumes a crashed run: promotes (if not already promoted), reboots, verifies.

.LINK
  https://github.com/adnoctem/winkit
  https://jans.cloud/2025/03/unidentified-network-public-firewall-profile-bei-windows-server-2025-domain-controller/
  https://learn.microsoft.com/en-us/windows/release-health/status-windows-server-2025#domain-controllers-manage-network-traffic-incorrectly-after-restarting

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
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'OutputPath', Justification = 'Consumed by Write-PhaseEnvelope.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'OperationTimeoutMinutes', Justification = 'Consumed by New-TargetSession.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'VerifyTimeoutMinutes', Justification = 'Consumed by the Verify phase.')]

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [string]$Server,

  [Parameter(Mandatory = $true)]
  [pscredential]$Credential,

  [Parameter(Mandatory = $true)]
  [string]$DomainName,

  [Parameter(Mandatory = $true)]
  [string]$DomainNetbiosName,

  [Parameter(Mandatory = $true)]
  [securestring]$SafeModeAdministratorPassword,

  [Parameter(Mandatory = $false)]
  [ValidateSet('Validate', 'Promote', 'Verify')]
  [string]$Phase,

  [Parameter(Mandatory = $false)]
  [switch]$Finish,

  [Parameter(Mandatory = $false)]
  [string]$OutputPath,

  [Parameter(Mandatory = $false)]
  [int]$OperationTimeoutMinutes = 30,

  [Parameter(Mandatory = $false)]
  [int]$VerifyTimeoutMinutes = 15
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
    domainNetbiosName = $DomainNetbiosName
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
    Returns a scrubbed exception message. The local admin password is the one
    secret this script holds in plaintext-capable form; if it ever appears in
    an error string it is replaced before anything is logged or emitted.
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
    Failing fast here prevents an interactive remote prompt hang in Phase 1.
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
    Checks the standard pending-reboot indicators on the target. A stale
    pending-reboot     flag makes feature installs queue behind it and can make
    Install-ADDSForest behave unpredictably, so this is checked before any
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

# ---------------------------------------------------------------------------
# Phase: Validate
# ---------------------------------------------------------------------------

function Invoke-ValidatePhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  # 1. Orchestrator-side TrustedHosts prerequisite (workgroup pre-domain
  #    scenario: NTLM-based WinRM requires the target in TrustedHosts).
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

  # 2. Session + negotiated authentication mechanism. Basic over plain HTTP
  #    cannot safely carry the SecureString DSRM password.
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

  # 4. DSRM password complexity, validated client-side before any remote work.
  if (-not (Test-DsrmPasswordComplexity -Password $SafeModeAdministratorPassword)) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target 'DSRM' -Source 'ADDS' -Action 'Validate' -Status 'Failed' -Detail 'Password does not meet AD complexity requirements.'
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The DSRM password fails AD complexity requirements (at least 8 characters, three of four character classes).'
  }
  Add-OperationResult -Results $diagnostics -Target 'DSRM' -Source 'ADDS' -Action 'Validate' -Status 'Completed' -Detail 'DSRM password meets complexity requirements.'

  # 5. Domain-role state: already promoted -> short-circuit; partial state
  #    (NTDS installed but not a DC) -> fail loudly, never auto-guess.
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
      Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Validate' -Status 'Completed' -Detail "Target is already a domain controller of $DomainName; feature install and prechecks skipped."
      return Write-PhaseEnvelope -PhaseName 'validate' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
    }
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Validate' -Status 'Failed' -Detail "Target is already a domain controller in $($state.Domain), expected $DomainName."
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Target is already a domain controller in $($state.Domain); expected $DomainName. Do not run this script against an existing environment."
  }

  # 6. Idempotent feature install (safe and reversible, but a real state
  #    change - the ADDSDeployment module only exists after this).
  $featureInstalled = $null
  $featureRebootRequired = $false
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
      $featureInstalled = [bool]$featureState.Installed
      $featureRebootRequired = [bool]$featureState.RebootRequired
      Add-OperationResult -Results $diagnostics -Target 'AD-Domain-Services' -Source 'WindowsFeature' -Action 'Install' -Status $(if ($featureInstalled) { 'Completed' } else { 'Failed' }) -Detail "installed: $featureInstalled, rebootRequiredByFeature: $featureRebootRequired"
    }
    catch {
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
      Add-OperationResult -Results $diagnostics -Target 'AD-Domain-Services' -Source 'WindowsFeature' -Action 'Install' -Status 'Failed' -Detail (Get-RedactedError $_)
      return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
    }
  }
  else {
    Add-OperationResult -Results $diagnostics -Target 'AD-Domain-Services' -Source 'WindowsFeature' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf - feature install skipped.'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  if (-not $featureInstalled) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'AD-Domain-Services feature install failed.'
  }

  if ($featureRebootRequired) {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WindowsFeature' -Action 'Install' -Status 'Failed' -Detail 'The feature install requires a reboot.'
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The AD-Domain-Services feature install requires a reboot. Reboot the target and re-run.'
  }

  # 7. Test-ADDSForestInstallation in a fresh session so the newly installed
  #    ADDSDeployment module is discoverable.
  Remove-PSSession -Session $session -ErrorAction SilentlyContinue
  $session = New-TargetSession -SessionCredential $Credential -Diagnostics $diagnostics
  if (-not $session) {
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Could not re-establish the WinRM session after the feature install.'
  }

  $precheckScript = {
    param($domainName, $netbiosName, $dsrmSecret)
    $ConfirmPreference = 'None'
    Import-Module ADDSDeployment -ErrorAction Stop
    $output = @()
    $precheckErrors = @()
    try {
      $output = @(Test-ADDSForestInstallation -DomainName $domainName -DomainNetbiosName $netbiosName -InstallDns:$true -SafeModeAdministratorPassword $dsrmSecret -NoRebootOnCompletion:$true -ErrorVariable precheckErrors -ErrorAction Continue)
    }
    catch {
      $precheckErrors = @($precheckErrors) + $_.Exception.Message
    }
    [pscustomobject]@{
      Success = ($precheckErrors.Count -eq 0)
      Messages = ($output + $precheckErrors)
    }
  }

  try {
    $precheck = Invoke-Command -Session $session -ScriptBlock $precheckScript -ArgumentList $DomainName, $DomainNetbiosName, $SafeModeAdministratorPassword -ErrorAction Stop
  }
  catch {
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Precheck' -Status 'Failed' -Detail (Get-RedactedError $_)
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText (Get-RedactedError $_)
  }

  $precheckMessages = @($precheck.Messages | Where-Object { $_ })
  Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Precheck' -Status $(if ($precheck.Success) { 'Completed' } else { 'Failed' }) -Detail $(if ($precheckMessages) { ($precheckMessages -join ' | ') } else { 'Test-ADDSForestInstallation returned no messages.' })
  Remove-PSSession -Session $session -ErrorAction SilentlyContinue

  if (-not $precheck.Success) {
    return Write-PhaseEnvelope -PhaseName 'validate' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'Test-ADDSForestInstallation reported failing checks.'
  }

  Write-PhaseEnvelope -PhaseName 'validate' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
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

  # State detection: never re-attempt promotion against an existing or
  # partially promoted domain controller.
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
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The target is in a partially promoted state (NTDS present, not a DC). Do not re-attempt automatically; demote with Uninstall-ADDSDomainController or rebuild from a clean snapshot, then re-run.'
  }

  if (-not $PSCmdlet.ShouldProcess($Server, "Promote to domain controller of $DomainName")) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Skipped' -Detail 'WhatIf - promotion skipped.'
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  $promoteScript = {
    param($domainName, $netbiosName, $dsrmSecret)
    $ConfirmPreference = 'None'
    $ErrorActionPreference = 'Stop'
    Import-Module ADDSDeployment -ErrorAction Stop
    Install-ADDSForest -DomainName $domainName -DomainNetbiosName $netbiosName -ForestMode Default -DomainMode Default -InstallDns:$true -SafeModeAdministratorPassword $dsrmSecret -NoRebootOnCompletion:$true -Confirm:$false -Force
  }

  # Run promotion as a job so the orchestrator can render live progress; the
  # job also survives long stretches without remote output.
  $promoteJob = $null
  try {
    $promoteJob = Invoke-Command -Session $session -ScriptBlock $promoteScript -ArgumentList $DomainName, $DomainNetbiosName, $SafeModeAdministratorPassword -AsJob -ErrorAction Stop
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
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Completed' -Detail 'Forest promotion completed; reboot is required to finalize.'
    Remove-Job -Job $promoteJob -Force -ErrorAction SilentlyContinue
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -RebootRequired $true
  }
  catch {
    $message = Get-RedactedError $_
    Remove-Job -Job $promoteJob -Force -ErrorAction SilentlyContinue
    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    if ($message -match 'timed out|operation timeout|I/O operation has been aborted|connection is closed|closed the connection') {
      Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Unknown' -Detail 'Connection lost during promotion - state unknown.'
      return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText 'The remote connection died during promotion; the promotion may have completed or partially completed. Detect state with -Phase Validate and resume with -Phase Promote -Finish.'
    }
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'ADDS' -Action 'Promote' -Status 'Failed' -Detail $message
    return Write-PhaseEnvelope -PhaseName 'promote' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText $message
  }
}

# ---------------------------------------------------------------------------
# Transition: controlled reboot
# ---------------------------------------------------------------------------

function Invoke-RebootTarget {
  <#
    Restarts the target and blocks until WinRM is confirmed back (authenticated
    probe with the reboot credential), rendering live progress. "WinRM is back"
    is only the connectivity floor; the Verify phase layers the AD-specific
    readiness poll on top.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    [pscredential]$DomainCredential
  )

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  if (-not $PSCmdlet.ShouldProcess($Server, 'Reboot and wait for WinRM to return')) {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Reboot' -Status 'Skipped' -Detail 'WhatIf - reboot skipped.'
    return Write-PhaseEnvelope -PhaseName 'reboot' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
  }

  try {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'WinRM' -Action 'Reboot' -Status 'Completed' -Detail 'Restarting target and waiting for WinRM (up to 15 minutes).'
    Restart-Computer -ComputerName $Server -Credential $DomainCredential -Force -ErrorAction Stop

    $rebootStartedAt = Get-Date
    $rebootDeadline = $rebootStartedAt.AddSeconds(900)
    $winrmBack = $false
    while (-not $winrmBack -and (Get-Date) -lt $rebootDeadline) {
      try {
        $null = Invoke-Command -ComputerName $Server -Credential $DomainCredential -ScriptBlock { $true } -ErrorAction Stop
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
# Windows Server 2025 DC firewall-profile mitigation
# ---------------------------------------------------------------------------

function Invoke-DcFirewallProfileMitigation {
  <#
    Mitigates the Windows Server 2025 domain controller firewall-profile bug
    (see .LINK): after a restart the NLA service can misclassify the DC
    network as Public/Unidentified, so the public firewall profile is applied
    and domain traffic breaks.

    The mitigation, mirroring the jans.cloud write-up:
      1. Three registry keys (AlwaysExpectDomainController,
         MaxNegativeCacheTtl, NegativeCachePeriod) that keep NLA/DNS/Netlogon
         from negative-caching the DC lookup during boot.
      2. Re-toggling the ms_tcpip6 adapter binding until Get-NetConnectionProfile
         reports DomainAuthenticated.
      3. An at-startup scheduled task (SYSTEM) that repeats step 2 on every
         future reboot, since the bug recurs on restart.

    Gated to Windows Server 2025 (build >= 26100). Failures are recorded as
    diagnostics only; they never fail the Verify phase.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter(Mandatory = $true)]
    $Session,

    [Parameter(Mandatory = $true)]
    [System.Collections.ArrayList]$Diagnostics
  )

  # 1. Only applies to Windows Server 2025+.
  $buildScript = {
    $build = [int](Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name CurrentBuild -ErrorAction SilentlyContinue).CurrentBuild
    [pscustomobject]@{ Build = $build }
  }
  $build = $null
  try {
    $build = (Invoke-Command -Session $Session -ScriptBlock $buildScript -ErrorAction Stop).Build
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target $Server -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Warn' -Detail "Could not query the target OS build: $(Get-RedactedError $_)"
    return
  }
  if ($build -lt 26100) {
    Add-OperationResult -Results $Diagnostics -Target $Server -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Skipped' -Detail "OS build $build - mitigation is Server 2025 (>= 26100) only."
    return
  }
  Add-OperationResult -Results $Diagnostics -Target $Server -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Completed' -Detail "OS build $build - Server 2025 firewall-profile mitigation applies."

  # 2. Registry keys.
  $registryScript = {
    $entries = @(
      @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\NlaSvc\Parameters'; Name = 'AlwaysExpectDomainController'; Value = 1 }
      @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters'; Name = 'MaxNegativeCacheTtl'; Value = 0 }
      @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'; Name = 'NegativeCachePeriod'; Value = 0 }
    )
    foreach ($entry in $entries) {
      if (-not (Test-Path -LiteralPath $entry.Path)) {
        $null = New-Item -Path $entry.Path -Force
      }
      Set-ItemProperty -Path $entry.Path -Name $entry.Name -Value $entry.Value -Type DWord -Force
    }
    $actual = @(foreach ($entry in $entries) { (Get-ItemProperty -Path $entry.Path -Name $entry.Name -ErrorAction Stop).$($entry.Name) })
    [pscustomobject]@{
      Applied = (($actual.Count -eq 3) -and ($actual[0] -eq 1) -and ($actual[1] -eq 0) -and ($actual[2] -eq 0))
    }
  }
  try {
    $registryResult = Invoke-Command -Session $Session -ScriptBlock $registryScript -ErrorAction Stop
    Add-OperationResult -Results $Diagnostics -Target 'FirewallProfileRegistry' -Source 'FirewallProfile' -Action 'Mitigate' -Status $(if ($registryResult.Applied) { 'Completed' } else { 'Warn' }) -Detail $(if ($registryResult.Applied) { 'All three mitigation registry keys applied.' } else { 'Registry keys could not be verified as applied.' })
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target 'FirewallProfileRegistry' -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Warn' -Detail "Registry key mitigation failed: $(Get-RedactedError $_)"
  }

  # 3. Profile check + ms_tcpip6 toggle loop, orchestrated from here so the
  #    progress bar renders (remote progress streams are not forwarded).
  $profileScript = {
    $profiles = @(Get-NetConnectionProfile -ErrorAction SilentlyContinue | Select-Object InterfaceAlias, NetworkCategory)
    [pscustomobject]@{
      Categories = @($profiles | ForEach-Object { $_.NetworkCategory.ToString() })
      Profiles = $profiles
    }
  }
  $toggleScript = {
    $null = Get-NetConnectionProfile -ErrorAction SilentlyContinue | ForEach-Object {
      Disable-NetAdapterBinding -Name $_.InterfaceAlias -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue
      Start-Sleep -Seconds 5
      Enable-NetAdapterBinding -Name $_.InterfaceAlias -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 5
  }

  $domainAuthenticated = $false
  $toggleAttempts = 0
  $toggleMaxAttempts = 7
  try {
    while (-not $domainAuthenticated -and ($toggleAttempts -lt $toggleMaxAttempts)) {
      $profileState = Invoke-Command -Session $Session -ScriptBlock $profileScript -ErrorAction Stop
      $domainAuthenticated = (($profileState.Categories -contains 'DomainAuthenticated') -and ($profileState.Categories.Count -gt 0))
      $toggleAttempts++
      $togglePct = [int][math]::Min(($toggleAttempts / $toggleMaxAttempts) * 100, 100)
      Write-Progress -Activity 'DC firewall profile fix' -Status "attempt $toggleAttempts / $toggleMaxAttempts - detected: $($profileState.Categories -join ', ')" -PercentComplete $togglePct
      Write-Host ("`rDC firewall profile ({0}/{1}): {2,-40}" -f $toggleAttempts, $toggleMaxAttempts, ($profileState.Categories -join ', ')) -NoNewline -ForegroundColor Cyan
      if (-not $domainAuthenticated) {
        $null = Invoke-Command -Session $Session -ScriptBlock $toggleScript -ErrorAction SilentlyContinue
      }
    }
    Write-Progress -Activity 'DC firewall profile fix' -Completed
    Write-Host ("`rDC firewall profile check finished.{0}" -f (' ' * 40)) -ForegroundColor Cyan

    Add-OperationResult -Results $Diagnostics -Target $Server -Source 'FirewallProfile' -Action 'Mitigate' -Status $(if ($domainAuthenticated) { 'Completed' } else { 'Warn' }) -Detail $(if ($domainAuthenticated) { "Network recognized as DomainAuthenticated after $toggleAttempts attempt(s)." } else { "Network not recognized as DomainAuthenticated after $toggleMaxAttempts attempts (last: $($profileState.Categories -join ', ')). Re-run later or check the scheduled task log." })
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target $Server -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Warn' -Detail "Network profile check failed: $(Get-RedactedError $_)"
  }

  # 4. At-startup scheduled task (SYSTEM) so the fix survives future reboots.
  $taskScript = {
    $taskName = 'Winkit-DcNetworkProfileFix'
    $scriptDirectory = 'C:\Windows\Temp\winkit-dc-fix'
    $null = New-Item -Path $scriptDirectory -ItemType Directory -Force
    $fixScript = @'
$i = 0
$maxLoop = 7
while (((Get-NetConnectionProfile).NetworkCategory -ne 'DomainAuthenticated') -and ($i -lt $maxLoop)) {
  $i++
  Get-NetConnectionProfile | ForEach-Object {
    Disable-NetAdapterBinding -Name $_.InterfaceAlias -ComponentID ms_tcpip6
    Start-Sleep -Seconds 5
    Enable-NetAdapterBinding -Name $_.InterfaceAlias -ComponentID ms_tcpip6
  }
  Start-Sleep -Seconds 5
}
'@
    $scriptPath = Join-Path $scriptDirectory 'Fix-DcNetworkProfile.ps1'
    Set-Content -LiteralPath $scriptPath -Value $fixScript -Encoding UTF8
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Description 'Toggles ms_tcpip6 until the DC network is DomainAuthenticated (Windows Server 2025 firewall-profile bug).' -Force | Out-Null
    [pscustomobject]@{ Registered = $true; TaskName = $taskName }
  }
  try {
    $taskResult = Invoke-Command -Session $Session -ScriptBlock $taskScript -ErrorAction Stop
    Add-OperationResult -Results $Diagnostics -Target $taskResult.TaskName -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Completed' -Detail 'At-startup remediation task registered.'
  }
  catch {
    Add-OperationResult -Results $Diagnostics -Target 'FirewallProfileTask' -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Warn' -Detail "Scheduled task registration failed: $(Get-RedactedError $_)"
  }
}

# ---------------------------------------------------------------------------
# Phase: Verify
# ---------------------------------------------------------------------------

function Invoke-VerifyPhase {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param()

  $startedAt = (Get-Date).ToUniversalTime().ToString('o')
  $diagnostics = New-Object System.Collections.ArrayList

  # Kerberos time-skew check first: a stale clock surfaces as an
  # authentication failure and sends you looking in the wrong place.
  try {
    $remoteTime = Invoke-Command -ComputerName $Server -Credential $domainCredential -ScriptBlock { (Get-Date).ToUniversalTime() } -ErrorAction Stop
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

  # Opening a session with the inferred domain credential IS the domain
  # authentication proof: same password as the pre-promotion local admin,
  # qualified down-level as NETBIOS\Administrator.
  $verifySession = New-TargetSession -SessionCredential $domainCredential -Diagnostics $diagnostics
  if (-not $verifySession) {
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "Authenticated session with $DomainNetbiosName\Administrator failed - the domain credential was rejected. Verify the password matches the pre-promotion local admin password."
  }
  Add-OperationResult -Results $diagnostics -Target $Server -Source 'AD' -Action 'Authenticate' -Status 'Completed' -Detail "Authenticated as $DomainNetbiosName\Administrator."

  # AD-specific readiness poll: SYSVOL/NETLOGON only exist once replication
  # has initialized, and Get-ADDomain on localhost confirms the directory is
  # serving. WinRM answering alone is not sufficient.
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
  $deadline = $pollStarted.AddMinutes($VerifyTimeoutMinutes)
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
    $readinessPct = [int][math]::Min(($elapsedSeconds / ($VerifyTimeoutMinutes * 60)) * 100, 100)
    Write-Progress -Activity 'Waiting for AD readiness' -Status "attempt $attempts, $elapsedSeconds s / $VerifyTimeoutMinutes min" -PercentComplete $readinessPct
    Write-Host ("`rWaiting for AD readiness ({0}s, attempt {1}): {2,-40}" -f $elapsedSeconds, $attempts, $Server) -NoNewline -ForegroundColor Cyan
    if (-not $ready) {
      Start-Sleep -Seconds 10
    }
  }
  Write-Progress -Activity 'Waiting for AD readiness' -Completed
  $readyAfterSeconds = [int]((Get-Date) - $pollStarted).TotalSeconds
  Write-Host ("`rAD readiness poll finished ({0}s).{1}" -f $readyAfterSeconds, (' ' * 40)) -ForegroundColor Cyan

  Add-OperationResult -Results $diagnostics -Target $Server -Source 'AD' -Action 'Readiness' -Status $(if ($ready) { 'Completed' } else { 'Failed' }) -Detail "attempts: $attempts, readyAfterSeconds: $readyAfterSeconds"
  Remove-PSSession -Session $verifySession -ErrorAction SilentlyContinue

  if (-not $ready) {
    return Write-PhaseEnvelope -PhaseName 'verify' -Success $false -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics -ErrorText "AD was not ready after $attempts attempts ($readyAfterSeconds seconds). Check NTDS/Netlogon/DNS service state on the target."
  }

  # Windows Server 2025 DC firewall-profile mitigation (see .LINK): applied
  # after readiness so NLA has the best chance of classifying the network.
  $mitigationSession = New-TargetSession -SessionCredential $domainCredential -Diagnostics $diagnostics
  if ($mitigationSession) {
    Invoke-DcFirewallProfileMitigation -Session $mitigationSession -Diagnostics $diagnostics
    Remove-PSSession -Session $mitigationSession -ErrorAction SilentlyContinue
  }
  else {
    Add-OperationResult -Results $diagnostics -Target $Server -Source 'FirewallProfile' -Action 'Mitigate' -Status 'Warn' -Detail 'Could not open a session for the DC firewall-profile mitigation.'
  }

  Write-PhaseEnvelope -PhaseName 'verify' -Success $true -StartedAt $startedAt -CompletedAt (Get-Date).ToUniversalTime().ToString('o') -Diagnostics $diagnostics
}

# ---------------------------------------------------------------------------
# Main flow
# ---------------------------------------------------------------------------

$phaseOrder = @('Validate', 'Promote', 'Verify')

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

$domainCredential = New-Object System.Management.Automation.PSCredential("$DomainNetbiosName\Administrator", $Credential.Password)

Write-Log -Message "Target: $Server | Domain: $DomainName | NetBIOS: $DomainNetbiosName | Phases: $($phasesToRun -join ', ')" -Color Cyan

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
      'Promote' {
        $envelope = Invoke-PromotePhase
        if ($envelope.success -and ('Verify' -in $phasesToRun)) {
          $rebootEnvelope = Invoke-RebootTarget -DomainCredential $domainCredential
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
