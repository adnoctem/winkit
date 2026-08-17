#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Enables Remote Desktop Services (RDP) locally or on remote machines.

.DESCRIPTION
  Enables Remote Desktop on the local machine (default) or on one or more
  remote machines (-ComputerName, driven over WinRM).

  Per machine, the script:
    1. Sets fDenyTSConnections to 0 - the master RDP switch.
    2. Configures Network Level Authentication (UserAuthentication on
       RDP-Tcp) per -EnableNLA (default: required).
    3. Ensures the TermService (Terminal Services) is set to Automatic and
       running. No restart is performed: the registry switch is read
       dynamically, so a restart would only drop existing sessions without
       being required.
    4. Enables the Remote Desktop firewall rules. Rules are matched by rule
       name (RemoteDesktop*) instead of display group, because the display
       group name is locale-dependent (e.g. 'Remote Desktop' on English
       systems vs 'RemoteDesktop' on German systems).
    5. Adds the specified user(s) to the local 'Remote Desktop Users' group
       (idempotent). Defaults to the account the script is running as; use
       -SkipUser to leave group membership untouched.

  Local mode requires an elevated PowerShell session. Remote mode requires
  WinRM access to the targets with an account that has local administrator
  rights there; the orchestrator itself does not need elevation.

  All steps are recorded as structured results (Add-OperationResult);
  -PassThru returns them, -DryRun previews without making changes.

.PARAMETER ComputerName
  Remote target(s). Omitting this enables RDP on the local machine.

.PARAMETER Credential
  Credentials for the remote WinRM session(s). Ignored in local mode.

.PARAMETER User
  Account(s) to add to the local 'Remote Desktop Users' group, e.g.
  COMPANY\user, user@company.com, or a local account name. Defaults to the
  account the script is running as.

.PARAMETER SkipUser
  Do not touch the 'Remote Desktop Users' group membership.

.PARAMETER EnableNLA
  Whether Network Level Authentication is required (default $true).
  Set to $false to allow legacy, non-NLA clients.

.PARAMETER DryRun
  Preview the operations without applying them.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Enable-RemoteDesktopServices.ps1
  Enables RDP locally and adds the current user to Remote Desktop Users.

.EXAMPLE
  PS> ./Enable-RemoteDesktopServices.ps1 -ComputerName dc01, exch01 -Credential $admin -User 'COMPANY\Administrator'
  Enables RDP on two machines and grants COMPANY\Administrator RDP access.

.EXAMPLE
  PS> ./Enable-RemoteDesktopServices.ps1 -ComputerName pc01 -Credential $admin -SkipUser -EnableNLA:$false
  Enables RDP without group changes and without the NLA requirement.

.EXAMPLE
  PS> ./Enable-RemoteDesktopServices.ps1 -ComputerName pc01 -Credential $admin -DryRun

.LINK
  https://github.com/adnoctem/winkit
  https://nt4admins.de/powershell/rdp-aktivieren-remote-per-powershell/
  https://sid-500.com/2021/03/22/enable-remote-desktop-remotely-with-powershell-enable-remotedesktop/
  https://www.der-windows-papst.de/2018/03/30/powershell-remotedesktop-aktivieren-und-firewall/

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string[]]
  $ComputerName,

  [Parameter(Mandatory = $false)]
  [pscredential]
  $Credential,

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string[]]
  $User,

  [Parameter(Mandatory = $false)]
  [switch]
  $SkipUser,

  [Parameter(Mandatory = $false)]
  [bool]
  $EnableNLA = $true,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$ProgressPreference = 'SilentlyContinue'

$isRemoteMode = [bool]$ComputerName

$users = @()
if (-not $SkipUser) {
  if ($User) {
    $users = @($User)
  }
  else {
    $users = @("$env:USERDOMAIN\$env:USERNAME")
  }
}

# ---- Per-machine operations -------------------------------------------------
# Runs in-process for local mode and via Invoke-Command for remote mode, so
# the exact same steps (and result entries) apply in both modes.

$rdpEnableScript = {
  param($enableNla, $userList)
  $entries = @()

  # 1. Registry - master RDP switch.
  try {
    $terminalServerPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    Set-ItemProperty -Path $terminalServerPath -Name 'fDenyTSConnections' -Value 0 -Type DWord -Force
    $fDeny = (Get-ItemProperty -Path $terminalServerPath -Name 'fDenyTSConnections').fDenyTSConnections
    $entries += [pscustomobject]@{
      Target = 'RDP Registry'
      Action = 'Enable'
      Status = $(if ($fDeny -eq 0) { 'Completed' } else { 'Failed' })
      Detail = "fDenyTSConnections=$fDeny (0 = RDP enabled)"
    }
  }
  catch {
    $entries += [pscustomobject]@{ Target = 'RDP Registry'; Action = 'Enable'; Status = 'Failed'; Detail = $_.Exception.Message }
  }

  # 2. Registry - Network Level Authentication.
  try {
    $rdpTcpPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'
    Set-ItemProperty -Path $rdpTcpPath -Name 'UserAuthentication' -Value ([int]$enableNla) -Type DWord -Force
    $userAuth = (Get-ItemProperty -Path $rdpTcpPath -Name 'UserAuthentication').UserAuthentication
    $entries += [pscustomobject]@{
      Target = 'NLA'
      Action = 'Configure'
      Status = $(if ($userAuth -eq ([int]$enableNla)) { 'Completed' } else { 'Failed' })
      Detail = "UserAuthentication=$userAuth (1 = NLA required)"
    }
  }
  catch {
    $entries += [pscustomobject]@{ Target = 'NLA'; Action = 'Configure'; Status = 'Failed'; Detail = $_.Exception.Message }
  }

  # 3. Service - TermService running and set to Automatic.
  try {
    Set-Service -Name TermService -StartupType Automatic -ErrorAction Stop
    $termService = Get-Service -Name TermService -ErrorAction Stop
    if ($termService.Status -ne 'Running') {
      Start-Service -Name TermService -ErrorAction Stop
      $termService = Get-Service -Name TermService -ErrorAction Stop
    }
    $entries += [pscustomobject]@{
      Target = 'TermService'
      Action = 'Ensure'
      Status = $(if ($termService.Status -eq 'Running') { 'Completed' } else { 'Failed' })
      Detail = "startup: $($termService.StartType), status: $($termService.Status)"
    }
  }
  catch {
    $entries += [pscustomobject]@{ Target = 'TermService'; Action = 'Ensure'; Status = 'Failed'; Detail = $_.Exception.Message }
  }

  # 4. Firewall - Remote Desktop rules by rule name (locale-independent).
  try {
    $rdpRules = @(Get-NetFirewallRule -Name 'RemoteDesktop*' -ErrorAction SilentlyContinue)
    if ($rdpRules.Count -eq 0) {
      $entries += [pscustomobject]@{ Target = 'Firewall'; Action = 'Enable'; Status = 'Warn'; Detail = 'No RemoteDesktop* firewall rules found.' }
    }
    else {
      $null = $rdpRules | Enable-NetFirewallRule
      $enabledRules = @(Get-NetFirewallRule -Name 'RemoteDesktop*' -ErrorAction SilentlyContinue | Where-Object { $_.Enabled -eq 'True' })
      $entries += [pscustomobject]@{
        Target = 'Firewall'
        Action = 'Enable'
        Status = $(if ($enabledRules.Count -eq $rdpRules.Count) { 'Completed' } else { 'Warn' })
        Detail = "$($enabledRules.Count) of $($rdpRules.Count) RemoteDesktop* rules enabled."
      }
    }
  }
  catch {
    $entries += [pscustomobject]@{ Target = 'Firewall'; Action = 'Enable'; Status = 'Failed'; Detail = $_.Exception.Message }
  }

  # 5. Group membership - Remote Desktop Users.
  $groupName = 'Remote Desktop Users'
  if ($userList.Count -eq 0) {
    $entries += [pscustomobject]@{ Target = $groupName; Action = 'AddMember'; Status = 'Skipped'; Detail = 'No users specified.' }
  }
  else {
    $existingMembers = @(Get-LocalGroupMember -Group $groupName -ErrorAction SilentlyContinue)
    foreach ($account in $userList) {
      $memberSid = $null
      try {
        $memberSid = ([System.Security.Principal.NTAccount]$account).Translate([System.Security.Principal.SecurityIdentifier])
      }
      catch {
        $memberSid = $null
      }

      $alreadyMember = $false
      foreach ($member in $existingMembers) {
        if ($memberSid -and $member.SID -eq $memberSid) {
          $alreadyMember = $true
          break
        }
        if (-not $memberSid -and $member.Name -eq $account) {
          $alreadyMember = $true
          break
        }
      }

      if ($alreadyMember) {
        $entries += [pscustomobject]@{ Target = $account; Action = 'AddMember'; Status = 'Skipped'; Detail = "Already a member of $groupName." }
        continue
      }

      try {
        Add-LocalGroupMember -Group $groupName -Member $account -ErrorAction Stop
        $entries += [pscustomobject]@{ Target = $account; Action = 'AddMember'; Status = 'Completed'; Detail = "Added to $groupName." }
      }
      catch {
        try {
          (& net localgroup $groupName $account /add 2>&1) | Out-Null
          if ($LASTEXITCODE -eq 0) {
            $entries += [pscustomobject]@{ Target = $account; Action = 'AddMember'; Status = 'Completed'; Detail = "Added to $groupName (via net localgroup)." }
          }
          else {
            $entries += [pscustomobject]@{ Target = $account; Action = 'AddMember'; Status = 'Failed'; Detail = "Add-LocalGroupMember and net localgroup both failed (exit $LASTEXITCODE)." }
          }
        }
        catch {
          $entries += [pscustomobject]@{ Target = $account; Action = 'AddMember'; Status = 'Failed'; Detail = $_.Exception.Message }
        }
      }
    }
  }

  $entries
}

# ---- Local mode -------------------------------------------------------------

if (-not $isRemoteMode) {
  if (-not $PSCmdlet.ShouldProcess('localhost', 'Enable Remote Desktop Services')) {
    Add-OperationResult -Results $_results -Target 'localhost' -Source 'RDP' -Action 'Enable' -Status 'Skipped' -Detail 'WhatIf - no changes applied.'
    if ($PassThru -or $DryRun) { $_results }
    exit 0
  }

  $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) {
    Write-Log -Message 'Local mode requires an elevated PowerShell session (Run as Administrator). Use -ComputerName for remote operation.' -Color Red
    Add-OperationResult -Results $_results -Target 'localhost' -Source 'RDP' -Action 'Enable' -Status 'Failed' -Detail 'RequiresAdministrator'
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }

  Write-Log -Message "Enabling Remote Desktop Services locally. Adding user(s): $($users -join ', ')" -Color Cyan
  $entries = & $rdpEnableScript -enableNla $EnableNLA -userList $users
  foreach ($entry in $entries) {
    Add-OperationResult -Results $_results -Target $entry.Target -Source 'RDP' -Action $entry.Action -Status $entry.Status -Detail $entry.Detail -Property @{ Machine = 'localhost' }
  }
}

# ---- Remote mode ------------------------------------------------------------

else {
  foreach ($target in $ComputerName) {
    if (-not $PSCmdlet.ShouldProcess($target, 'Enable Remote Desktop Services')) {
      Add-OperationResult -Results $_results -Target 'RDP' -Source 'RDP' -Action 'Enable' -Status 'Skipped' -Detail 'WhatIf - no changes applied.' -Property @{ Machine = $target }
      continue
    }

    Write-Log -Message "Enabling Remote Desktop Services on $target..." -Color Cyan
    $session = $null
    $entries = @()
    try {
      $session = New-PSSession -ComputerName $target -Credential $Credential -ErrorAction Stop
      $entries = @(Invoke-Command -Session $session -ScriptBlock $rdpEnableScript -ArgumentList $EnableNLA, $users -ErrorAction Stop)
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
      $session = $null
    }
    catch {
      $entries = @([pscustomobject]@{ Target = 'RDP'; Action = 'Enable'; Status = 'Failed'; Detail = "Remote operation failed: $($_.Exception.Message)" })
      Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }

    foreach ($entry in $entries) {
      Add-OperationResult -Results $_results -Target $entry.Target -Source 'RDP' -Action $entry.Action -Status $entry.Status -Detail $entry.Detail -Property @{ Machine = $target }
    }
  }
}

# ---- Summary ----------------------------------------------------------------

$failedCount = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Enable-RemoteDesktopServices'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($failedCount -gt 0) {
  Write-Log -Message "Completed with $failedCount failed step(s)." -Color Red
  exit 1
}

Write-Log -Message 'Remote Desktop Services enabled successfully.' -Color Green
exit 0
