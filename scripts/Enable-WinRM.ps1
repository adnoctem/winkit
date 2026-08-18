#Requires -Version 5.0
#Requires -RunAsAdministrator
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.2.0' }

<#
.SYNOPSIS
  Enables WinRM (PowerShell remoting) with secure defaults and optional
  test-VM-only relaxations.

.DESCRIPTION
  Configures the WinRM stack so the machine can be driven remotely via
  Invoke-Command - the capability the winkit ADDS/Exchange orchestrator
  scripts depend on. The script is production-safe by default and mirrors the
  secure baseline (Kerberos/Negotiate authentication, encrypted transport,
  explicit TrustedHosts only when required by a workgroup scenario).

  Performs:
    1. WinRM service set to Automatic and started.
    2. HTTP listener on port 5985 created when none exists (via
       Set-WSManQuickConfig).
    3. 'Windows Remote Management' firewall rules enabled.
    4. TrustedHosts configured when requested: -TrustedHosts replaces the
       value with the given host list (the workgroup/pre-domain scenario the
       ADDS scripts remediate with Set-Item WSMan:\localhost\Client\
       TrustedHosts), or -TrustedHostsAll sets '*' - only for isolated
       environments.
    5. Verification of service, listener, and local connectivity.

  Test-VM-only relaxations, all OFF by default and explicitly opt-in:
    -AllowUnencrypted - allows unencrypted WinRM traffic (listener
    AllowUnencrypted=true).
    -EnableBasicAuth - enables Basic authentication on the WinRM service.
  These are acceptable ONLY on isolated, disposable test infrastructure with
  no sensitive data or network exposure - never on domain-joined production
  machines. The reference this script derives from (Enable-TrustedRemoting in
  stevencohn/WindowsPowerShell) hardcoded these insecure settings; here they
  are deliberately flipped to opt-in.

  Requires administrator elevation.

.PARAMETER TrustedHosts
  Replace the client TrustedHosts value with this comma-separated host list.
  Needed when the orchestrator and target are both in a workgroup (no Kerberos
  realm yet). Pass the host(s) the orchestrator connects FROM.

.PARAMETER TrustedHostsAll
  Set TrustedHosts to '*'. Only for isolated test environments - trusts every
  host for NTLM-based WinRM.

.PARAMETER AllowUnencrypted
  TEST-VM ONLY: allow unencrypted WinRM traffic on the listener.

.PARAMETER EnableBasicAuth
  TEST-VM ONLY: enable Basic authentication on the WinRM service.

.PARAMETER DryRun
  Preview changes without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Enable-WinRM.ps1
  Enables WinRM with secure defaults (service, listener, firewall).

.EXAMPLE
  PS> ./Enable-WinRM.ps1 -TrustedHosts 192.0.2.10
  Enables WinRM and trusts the given orchestrator for NTLM-based remoting.

.EXAMPLE
  PS> ./Enable-WinRM.ps1 -TrustedHostsAll -AllowUnencrypted -EnableBasicAuth
  TEST VM ONLY: full insecure WinRM bootstrap for a disposable automation VM.

.LINK
  https://github.com/adnoctem/winkit
  https://learn.microsoft.com/en-us/powershell/scripting/learn/ps101/08-powershell-remoting

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - remoting is the primary Server Core administration path.
  SYSTEM-account execution: works - no user-context dependency.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Preview changes without applying them.'
  )]
  [switch]
  $DryRun,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Replace the client TrustedHosts value with this host list (workgroup scenario).'
  )]
  [string[]]
  $TrustedHosts,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'TEST VM ONLY: set TrustedHosts to * (trust every host).'
  )]
  [switch]
  $TrustedHostsAll,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'TEST VM ONLY: allow unencrypted WinRM traffic.'
  )]
  [switch]
  $AllowUnencrypted,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'TEST VM ONLY: enable Basic authentication on the WinRM service.'
  )]
  [switch]
  $EnableBasicAuth,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no WinRM changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

function Test-WinRmListenerPresent {
  <#
    Checks whether a WinRM listener is configured via winrm enumerate (the
    WSMan: drive does not reliably enumerate the automatic listener). Requires
    elevation - guaranteed by the script's #Requires -RunAsAdministrator.
  #>
  [CmdletBinding()]
  param()

  $output = (& winrm enumerate winrm/config/listener 2>&1) | Out-String
  return (($LASTEXITCODE -eq 0) -and (-not [string]::IsNullOrWhiteSpace($output)))
}

if ($AllowUnencrypted -or $EnableBasicAuth -or $TrustedHostsAll) {
  Write-Log -Message 'WARNING: one or more test-VM-only relaxations were requested. These are acceptable ONLY on isolated, disposable test infrastructure - never on domain-joined production machines.' -Color Red
}

# 1. WinRM service: Automatic + running.
Write-Log -Message 'Configuring the WinRM service...' -Color Yellow
if ($PSCmdlet.ShouldProcess('WinRM', 'Set startup type to Automatic and start the service')) {
  try {
    $null = Set-ServiceStartupState -Name WinRM -StartupType Automatic
    $service = Get-Service -Name WinRM -ErrorAction Stop
    if ($service.Status -ne 'Running') {
      Start-Service -Name WinRM -ErrorAction Stop
    }
    $service = Get-Service -Name WinRM -ErrorAction Stop
    Add-OperationResult -Results $_results -Target 'WinRM' -Source 'WinRM' -Action 'Configure' -Status $(if ($service.Status -eq 'Running') { 'Completed' } else { 'Failed' }) -Detail "startup: $($service.StartType), status: $($service.Status)"
  }
  catch {
    Add-OperationResult -Results $_results -Target 'WinRM' -Source 'WinRM' -Action 'Configure' -Status 'Failed' -Detail $_.Exception.Message
  }
}
else {
  Add-OperationResult -Results $_results -Target 'WinRM' -Source 'WinRM' -Action 'Configure' -Status 'Skipped' -Detail 'WhatIf'
}

# 2. HTTP listener on port 5985 when none exists.
Write-Log -Message 'Ensuring the WinRM listener...' -Color Yellow
if (-not (Test-WinRmListenerPresent)) {
  if ($PSCmdlet.ShouldProcess('WinRM listener', 'Create the HTTP listener via Set-WSManQuickConfig')) {
    try {
      Set-WSManQuickConfig -Force -ErrorAction Stop
      $listenerOk = Test-WinRmListenerPresent
      Add-OperationResult -Results $_results -Target 'WinRM Listener' -Source 'WinRM' -Action 'Configure' -Status $(if ($listenerOk) { 'Completed' } else { 'Failed' }) -Detail 'HTTP listener created.'
    }
    catch {
      Add-OperationResult -Results $_results -Target 'WinRM Listener' -Source 'WinRM' -Action 'Configure' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  else {
    Add-OperationResult -Results $_results -Target 'WinRM Listener' -Source 'WinRM' -Action 'Configure' -Status 'Skipped' -Detail 'WhatIf'
  }
}
else {
  Add-OperationResult -Results $_results -Target 'WinRM Listener' -Source 'WinRM' -Action 'Configure' -Status 'Completed' -Detail 'Listener already present.'
}

# 3. Firewall rules for Windows Remote Management.
Write-Log -Message 'Enabling Windows Remote Management firewall rules...' -Color Yellow
if ($PSCmdlet.ShouldProcess('Windows Remote Management firewall rules', 'Enable')) {
  try {
    $rules = @(Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction SilentlyContinue)
    if ($rules.Count -eq 0) {
      Add-OperationResult -Results $_results -Target 'Firewall' -Source 'WinRM' -Action 'Configure' -Status 'Warn' -Detail 'No Windows Remote Management firewall rules found.'
    }
    else {
      $null = $rules | Enable-NetFirewallRule
      $enabled = @(Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction SilentlyContinue | Where-Object { $_.Enabled -eq 'True' }).Count
      Add-OperationResult -Results $_results -Target 'Firewall' -Source 'WinRM' -Action 'Configure' -Status $(if ($enabled -eq $rules.Count) { 'Completed' } else { 'Warn' }) -Detail "$enabled of $($rules.Count) Windows Remote Management rule(s) enabled."
    }
  }
  catch {
    Add-OperationResult -Results $_results -Target 'Firewall' -Source 'WinRM' -Action 'Configure' -Status 'Failed' -Detail $_.Exception.Message
  }
}
else {
  Add-OperationResult -Results $_results -Target 'Firewall' -Source 'WinRM' -Action 'Configure' -Status 'Skipped' -Detail 'WhatIf'
}

# 4. TrustedHosts.
if ($TrustedHostsAll -or $TrustedHosts) {
  Write-Log -Message 'Configuring TrustedHosts...' -Color Yellow
  $trustedValue = if ($TrustedHostsAll) { '*' } else { ($TrustedHosts -join ',') }

  if ($PSCmdlet.ShouldProcess('TrustedHosts', "Set value to '$trustedValue'")) {
    try {
      Set-Item 'WSMan:\localhost\Client\TrustedHosts' -Value $trustedValue -Force -ErrorAction Stop
      $current = (Get-Item 'WSMan:\localhost\Client\TrustedHosts' -ErrorAction Stop).Value
      Add-OperationResult -Results $_results -Target 'TrustedHosts' -Source 'WinRM' -Action 'Configure' -Status $(if ($current -eq $trustedValue) { 'Completed' } else { 'Failed' }) -Detail "TrustedHosts set to: $current"
    }
    catch {
      Add-OperationResult -Results $_results -Target 'TrustedHosts' -Source 'WinRM' -Action 'Configure' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  else {
    Add-OperationResult -Results $_results -Target 'TrustedHosts' -Source 'WinRM' -Action 'Configure' -Status 'Skipped' -Detail 'WhatIf'
  }
}

# 5. Test-VM-only relaxations.
if ($AllowUnencrypted) {
  if ($PSCmdlet.ShouldProcess('WinRM AllowUnencrypted', 'Set to true (test VM only)')) {
    try {
      Set-Item 'WSMan:\localhost\Service\AllowUnencrypted' -Value $true -Force -ErrorAction Stop
      Add-OperationResult -Results $_results -Target 'AllowUnencrypted' -Source 'WinRM' -Action 'Configure' -Status 'Completed' -Detail 'Unencrypted WinRM traffic allowed (test VM only).'
    }
    catch {
      Add-OperationResult -Results $_results -Target 'AllowUnencrypted' -Source 'WinRM' -Action 'Configure' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  else {
    Add-OperationResult -Results $_results -Target 'AllowUnencrypted' -Source 'WinRM' -Action 'Configure' -Status 'Skipped' -Detail 'WhatIf'
  }
}

if ($EnableBasicAuth) {
  if ($PSCmdlet.ShouldProcess('WinRM Basic authentication', 'Enable (test VM only)')) {
    try {
      Set-Item 'WSMan:\localhost\Service\Auth\Basic' -Value $true -Force -ErrorAction Stop
      Add-OperationResult -Results $_results -Target 'BasicAuth' -Source 'WinRM' -Action 'Configure' -Status 'Completed' -Detail 'Basic authentication enabled (test VM only).'
    }
    catch {
      Add-OperationResult -Results $_results -Target 'BasicAuth' -Source 'WinRM' -Action 'Configure' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
  else {
    Add-OperationResult -Results $_results -Target 'BasicAuth' -Source 'WinRM' -Action 'Configure' -Status 'Skipped' -Detail 'WhatIf'
  }
}

# 6. Verification.
Write-Log -Message 'Verifying WinRM...' -Color Yellow
try {
  $service = Get-Service -Name WinRM -ErrorAction Stop
  Add-OperationResult -Results $_results -Target 'WinRM Service' -Source 'WinRM' -Action 'Verify' -Status $(if ($service.Status -eq 'Running') { 'Completed' } else { 'Failed' }) -Detail $service.Status.ToString()

  $listenerOk = Test-WinRmListenerPresent
  Add-OperationResult -Results $_results -Target 'WinRM Listener' -Source 'WinRM' -Action 'Verify' -Status $(if ($listenerOk) { 'Completed' } else { 'Failed' }) -Detail $(if ($listenerOk) { 'Listener present.' } else { 'No listener detected.' })

  $null = Test-WSMan -ComputerName localhost -ErrorAction Stop
  Add-OperationResult -Results $_results -Target 'Test-WSMan' -Source 'WinRM' -Action 'Verify' -Status 'Completed' -Detail 'Local WinRM connectivity confirmed.'

  $trustedHostsValue = (Get-Item 'WSMan:\localhost\Client\TrustedHosts' -ErrorAction SilentlyContinue).Value
  Add-OperationResult -Results $_results -Target 'TrustedHosts' -Source 'WinRM' -Action 'Verify' -Status 'Completed' -Detail $(if ($trustedHostsValue) { "Current value: $trustedHostsValue" } else { 'Not configured (empty).' })
}
catch {
  Add-OperationResult -Results $_results -Target 'WinRM' -Source 'WinRM' -Action 'Verify' -Status 'Failed' -Detail $_.Exception.Message
}

# 7. Summary.
$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
$_skipped = @($_results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_completed = @($_results | Where-Object { $_.Status -notin @('Failed', 'Skipped') }).Count
$_color = if ($_failed -gt 0) { 'Yellow' } else { 'Green' }
Write-Log -Message "`nWinRM enablement workflow complete. Completed: $_completed | Skipped: $_skipped | Failed: $_failed" -Color $_color
$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Enable-WinRM'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failed -gt 0 -and -not $DryRun) {
  exit 1
}
exit 0
