#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Removes selected bundled, sponsored, and policy-backed Windows bloatware.
.DESCRIPTION
  Uses built-in grouped package patterns to remove matching UPF AppX/MSIX
  packages from both the online provisioning store and installed user profiles.
  It then runs guarded follow-up procedures for known bloatware surfaces that
  cannot be handled by Remove-AppxPackage alone, including WinGet/ARP entries,
  Windows capabilities, scheduled tasks, registry policy suppressions, Start
  ghost-pin prevention policies, and safe shortcut cleanup.

  Provisioned package removal is enabled by default because many sponsored
  Windows 11 stubs are provisioned but not installed until first sign-in.
.PARAMETER Group
  Built-in group names to remove. Defaults to Default and OEM. Risky remains
  explicit opt-in.
.PARAMETER Pattern
  Additional package wildcard patterns to include.
.PARAMETER Config
  Optional PSD1 or JSON file with grouped package patterns. Values override or
  add to the built-in groups.
.PARAMETER AllUsers
  Remove installed packages for all users. Requires elevation on most builds.
.PARAMETER SkipProvisioned
  Do not remove provisioned packages from the Windows image.
.PARAMETER SkipWinGet
  Skip WinGet/ARP Win32 removal targets.
.PARAMETER SkipSpecialProcedures
  Skip registry, task, capability, shortcut, and policy follow-up procedures.
.PARAMETER ConfigureStartPins
  Set the Windows 11 ConfigureStartPins CSP to an empty pinned list for future
  profiles. This does not modify the current user's start2.bin.
.PARAMETER IncludeProtected
  Include packages normally protected by winkit safety checks.
.PARAMETER Force
  Required together with -IncludeProtected to actually remove protected matches.
.PARAMETER DryRun
  Preview matching packages and removals without changing the system.
.PARAMETER ExportConfig
  Export the default package pattern groups as JSON and exit.
.PARAMETER ExportPath
  Output path used with -ExportConfig.
.PARAMETER PassThru
  Return structured lifecycle and follow-up results.
.EXAMPLE
  PS> ./Remove-Bloatware.ps1 -DryRun

  Preview removal for the Default and OEM groups.
.EXAMPLE
  PS> ./Remove-Bloatware.ps1 -Group Default,OEM -AllUsers -DryRun

  Preview the default non-risky package selection across all users.
.EXAMPLE
  PS> ./Remove-Bloatware.ps1 -Group Default,Risky -ConfigureStartPins

  Include preference-sensitive risky removals and configure empty future Start pins.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [Parameter(
    Position = 0,
    Mandatory = $false,
    HelpMessage = 'Built-in group names to remove. Defaults to Default and OEM.'
  )]
  [string[]]
  $Group = @('Default', 'OEM'),

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Additional package wildcard patterns to include.'
  )]
  [string[]]
  $Pattern,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Optional PSD1 or JSON file with grouped package patterns.'
  )]
  [string]
  $Config,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Remove installed packages for all users.'
  )]
  [switch]
  $AllUsers,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Do not remove provisioned packages from the Windows image.'
  )]
  [switch]
  $SkipProvisioned,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Skip WinGet/ARP Win32 removal targets.'
  )]
  [switch]
  $SkipWinGet,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Skip registry, task, capability, shortcut, and policy follow-up procedures.'
  )]
  [switch]
  $SkipSpecialProcedures,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Set the Windows 11 ConfigureStartPins CSP to an empty pinned list.'
  )]
  [switch]
  $ConfigureStartPins,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Include packages normally protected by winkit safety checks.'
  )]
  [switch]
  $IncludeProtected,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Required together with -IncludeProtected to actually remove protected matches.'
  )]
  [switch]
  $Force,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Preview matching packages and removals without changing the system.'
  )]
  [switch]
  $DryRun,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Export the default package pattern groups as JSON and exit.'
  )]
  [switch]
  $ExportConfig,

  [Parameter(
    Mandatory = $false,
    HelpMessage = 'Output path used with -ExportConfig.'
  )]
  [string]
  $ExportPath,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no bloatware changes will be applied`n" -Color Yellow
}

$_packageGroups = @{
  Default = @(
    'Microsoft.BingWeather*'
    'Microsoft.BingNews*'
    'Microsoft.BingSearch*'
    'Microsoft.GetHelp*'
    'Microsoft.Getstarted*'
    'Microsoft.MicrosoftOfficeHub*'
    'Microsoft.MicrosoftSolitaireCollection*'
    'Microsoft.MicrosoftStickyNotes*'
    'Microsoft.OutlookForWindows*'
    'Microsoft.People*'
    'Microsoft.PowerAutomateDesktop*'
    'Microsoft.SkypeApp*'
    'Microsoft.Todos*'
    'Microsoft.WindowsFeedbackHub*'
    'Microsoft.WindowsMaps*'
    'Microsoft.WindowsSoundRecorder*'
    'Microsoft.ZuneMusic*'
    'Microsoft.ZuneVideo*'
    'MicrosoftCorporationII.MicrosoftFamily*'
    'MicrosoftCorporationII.QuickAssist*'
    'MicrosoftTeams*'
    'MSTeams*'
    'microsoft.windowscommunicationsapps*'
    'Microsoft.StartExperiencesApp*'
    'Microsoft.Windows.ParentalControls*'
    'Microsoft.Edge.GameAssist*'
    'Microsoft.YourPhone*'
    'Clipchamp.Clipchamp*'
    'Disney*'
    '*Disney*'
    'Netflix*'
    '*Netflix*'
    'Spotify*'
    'SpotifyAB*'
    '*Roblox*'
    'ROBLOXCORPORATION.ROBLOX*'
    'Amazon*'
    'AmazonVideo*'
    '*Instagram*'
    'Facebook.317180B0BB486*'
    'Facebook.InstagramBeta*'
    'Keeper*'
    'KeeperSecurityInc*'
    '*CandyCrush*'
    'king.com.*'
    '*BubbleWitch*'
    'AdobeSystemsIncorporated.AdobePhotoshopExpress*'
    '*Duolingo*'
    'PandoraMediaInc*'
    'Flipboard.Flipboard*'
    'LinkedInforWindows*'
    '*TikTok*'
    'BytedancePte.Ltd.TikTok*'
    '*WhatsApp*'
    '5319275A.WhatsAppDesktop*'
    'Microsoft.GamingApp*'
    'Microsoft.XboxGamingOverlay*'
    'Microsoft.XboxGameOverlay*'
    'Microsoft.Xbox.TCUI*'
    'Microsoft.XboxIdentityProvider*'
    'Microsoft.XboxSpeechToTextOverlay*'
    'Microsoft.XboxApp*'
    'Microsoft.XboxGameCallableUI*'
    'Microsoft.Copilot*'
    'Microsoft.Windows.Ai.Copilot.Provider*'
    'Microsoft.Office.OneNote*'
    'Microsoft.RemoteDesktop*'
    'Microsoft.Whiteboard*'
    'Microsoft.WindowsAlarms*'
    'Microsoft.549981C3F5F10*'
    'Microsoft.MicrosoftJournal*'
    'Microsoft.PowerBI*'
    'MicrosoftCorporationII.MicrosoftEdgeGameAssist*'
  )

  OEM = @(
    'DolbyLaboratories.DolbyAccess*'
    'NVIDIACorp.NVIDIAControlPanel*'
    'Microsoft.DrawboardPDF*'
    'ActiproSoftwareLLC*'
    'AppUp.IntelGraphicsExperience*'
    'RealtekSemiconductorCorp.RealtekAudioControl*'
  )

  Risky = @(
    'MicrosoftWindows.Client.WebExperience*'
    'Microsoft.Windows.DevHome*'
    'MicrosoftWindows.UndockedDevKit*'
  )
}

$_results = New-Object System.Collections.ArrayList
$_useDetailedPackageConfirm = $PSBoundParameters.ContainsKey('Confirm') -and [bool]$PSBoundParameters['Confirm']
$_wingetAgreementArguments = @('--accept-source-agreements')
$_wingetUninstallArguments = @('--silent', '--accept-source-agreements', '--accept-package-agreements', '--force')

function Test-BloatwarePattern {
  param (
    [Parameter(Mandatory = $true)]
    [string[]]$Patterns,

    [Parameter(Mandatory = $true)]
    [string[]]$Match
  )

  foreach ($_pattern in $Patterns) {
    foreach ($_match in $Match) {
      if ($_pattern -like $_match -or $_match -like $_pattern) {
        return $true
      }
    }
  }

  return $false
}

function Invoke-BloatwarePackageSetRemove {
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,

    [Parameter(Mandatory = $true)]
    [string[]]$Patterns,

    [string]$BatchTarget,
    [string]$BatchAction,
    [int]$MatchCount,
    [switch]$Installed,
    [switch]$Provisioned,
    [switch]$AllUsers,
    [switch]$IncludeProtected,
    [switch]$Force,
    [switch]$DryRun
  )

  if ($DryRun -or $WhatIfPreference) {
    $_removeResults = @(Uninstall-UPFAppxPackageSet -Pattern $Patterns -Installed:$Installed -Provisioned:$Provisioned -AllUsers:$AllUsers -IncludeProtected:$IncludeProtected -Force:$Force -DryRun:$DryRun -PassThru -WhatIf:$WhatIfPreference -Confirm:$false)
    foreach ($_result in $_removeResults) { [void]$Results.Add($_result) }
    return
  }

  if ($_useDetailedPackageConfirm) {
    $_removeResults = @(Uninstall-UPFAppxPackageSet -Pattern $Patterns -Installed:$Installed -Provisioned:$Provisioned -AllUsers:$AllUsers -IncludeProtected:$IncludeProtected -Force:$Force -PassThru -Confirm:$true)
    foreach ($_result in $_removeResults) { [void]$Results.Add($_result) }
    return
  }

  if ($MatchCount -gt 0 -and -not $PSCmdlet.ShouldProcess($BatchTarget, $BatchAction)) {
    Add-OperationResult -Results $Results -Target $BatchTarget -Source 'UPFAppxPackage' -Action 'Uninstall' -Status 'Skipped' -Detail 'BatchConfirmationDeclined'
    return
  }

  $_removeResults = @(Uninstall-UPFAppxPackageSet -Pattern $Patterns -Installed:$Installed -Provisioned:$Provisioned -AllUsers:$AllUsers -IncludeProtected:$IncludeProtected -Force:$Force -PassThru -Confirm:$false)
  foreach ($_result in $_removeResults) { [void]$Results.Add($_result) }
}

function Invoke-BloatwareRegistrySet {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$Path,
    [string]$Name,
    [object]$Value,
    [string]$Type,
    [string]$Detail
  )

  $_target = "$Path\$Name"
  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'SetValue' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_target, "Set registry value to $Value")) {
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'SetValue' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  try {
    $_result = Set-RegistryValue -Path $Path -Name $Name -Value $Value -Type $Type
    $_status = if ($_result) { $_result.Status } else { 'Completed' }
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'SetValue' -Status $_status -Detail $Detail
  }
  catch {
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'SetValue' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-BloatwareRegistryRemove {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$Path,
    [string]$Name,
    [string]$Detail
  )

  $_target = "$Path\$Name"
  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'RemoveValue' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_target, 'Remove registry value')) {
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'RemoveValue' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  try {
    $_result = Remove-RegistryValue -Path $Path -Name $Name
    $_status = if ($_result) { $_result.Status } else { 'Completed' }
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'RemoveValue' -Status $_status -Detail $Detail
  }
  catch {
    Add-OperationResult -Results $Results -Target $_target -Source 'Registry' -Action 'RemoveValue' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-BloatwareShortcutRemove {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$Path
  )

  $_resolvedPath = [Environment]::ExpandEnvironmentVariables($Path)
  if (-not (Test-Path -LiteralPath $_resolvedPath -PathType Leaf)) {
    Add-OperationResult -Results $Results -Target $_resolvedPath -Source 'FileSystem' -Action 'RemoveShortcut' -Status 'Skipped' -Detail 'NotFound'
    return
  }

  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $_resolvedPath -Source 'FileSystem' -Action 'RemoveShortcut' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_resolvedPath, 'Remove shortcut')) {
    Add-OperationResult -Results $Results -Target $_resolvedPath -Source 'FileSystem' -Action 'RemoveShortcut' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  try {
    Remove-Item -LiteralPath $_resolvedPath -Force
    Add-OperationResult -Results $Results -Target $_resolvedPath -Source 'FileSystem' -Action 'RemoveShortcut' -Status 'Removed' -Detail 'Shortcut removed.'
  }
  catch {
    Add-OperationResult -Results $Results -Target $_resolvedPath -Source 'FileSystem' -Action 'RemoveShortcut' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-BloatwareTaskDisable {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$TaskPath,
    [string]$TaskName
  )

  $_target = "$TaskPath$TaskName"
  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $_target -Source 'ScheduledTask' -Action 'Disable' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_target, 'Disable scheduled task')) {
    Add-OperationResult -Results $Results -Target $_target -Source 'ScheduledTask' -Action 'Disable' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  try {
    Disable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
    Add-OperationResult -Results $Results -Target $_target -Source 'ScheduledTask' -Action 'Disable' -Status 'Disabled' -Detail 'Scheduled task disabled.'
  }
  catch {
    Add-OperationResult -Results $Results -Target $_target -Source 'ScheduledTask' -Action 'Disable' -Status 'Skipped' -Detail $_.Exception.Message
  }
}

function Invoke-BloatwareCapabilityRemove {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$Pattern
  )

  if (-not (Test-Elevation)) {
    Add-OperationResult -Results $Results -Target $Pattern -Source 'Capability' -Action 'Remove' -Status 'Skipped' -Detail 'RequiresElevation'
    return
  }

  $_capabilities = @(Get-WindowsCapability -Online -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $Pattern })
  if ($_capabilities.Count -eq 0) {
    Add-OperationResult -Results $Results -Target $Pattern -Source 'Capability' -Action 'Remove' -Status 'Skipped' -Detail 'NoMatch'
    return
  }

  foreach ($_capability in $_capabilities) {
    if ($DryRun) {
      Add-OperationResult -Results $Results -Target $_capability.Name -Source 'Capability' -Action 'Remove' -Status 'Skipped' -Detail 'DryRun'
      continue
    }

    if (-not $PSCmdlet.ShouldProcess($_capability.Name, 'Remove Windows capability')) {
      Add-OperationResult -Results $Results -Target $_capability.Name -Source 'Capability' -Action 'Remove' -Status 'Skipped' -Detail 'WhatIf'
      continue
    }

    try {
      Remove-WindowsCapability -Online -Name $_capability.Name -ErrorAction Stop | Out-Null
      Add-OperationResult -Results $Results -Target $_capability.Name -Source 'Capability' -Action 'Remove' -Status 'Removed' -Detail 'Windows capability removed.'
    }
    catch {
      Add-OperationResult -Results $Results -Target $_capability.Name -Source 'Capability' -Action 'Remove' -Status 'Failed' -Detail $_.Exception.Message
    }
  }
}

function Invoke-BloatwareWinGetRemove {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string]$Id
  )

  $_winget = Get-Command -Name winget.exe -ErrorAction SilentlyContinue
  if (-not $_winget) {
    Add-OperationResult -Results $Results -Target $Id -Source 'WinGet' -Action 'Uninstall' -Status 'Skipped' -Detail 'WinGetUnavailable'
    return
  }

  $_listArguments = @('list', '--id', $Id, '--exact') + $_wingetAgreementArguments
  $_list = & $_winget.Source @_listArguments 2>&1
  $_listExitCode = $LASTEXITCODE
  if ($_listExitCode -ne 0) {
    $_detail = ($_list | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($_detail)) {
      $_detail = "winget list exited with $_listExitCode."
    }
    Add-OperationResult -Results $Results -Target $Id -Source 'WinGet' -Action 'Resolve' -Status 'Failed' -Detail $_detail
    return
  }

  if (-not (($_list | Out-String) -match [regex]::Escape($Id))) {
    Add-OperationResult -Results $Results -Target $Id -Source 'WinGet' -Action 'Uninstall' -Status 'Skipped' -Detail 'NoMatch'
    return
  }

  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $Id -Source 'WinGet' -Action 'Uninstall' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  if (-not $PSCmdlet.ShouldProcess($Id, 'Uninstall WinGet package')) {
    Add-OperationResult -Results $Results -Target $Id -Source 'WinGet' -Action 'Uninstall' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  $_uninstallArguments = @('uninstall', '--id', $Id, '--exact') + $_wingetUninstallArguments
  $_output = & $_winget.Source @_uninstallArguments 2>&1
  $_uninstallExitCode = $LASTEXITCODE
  if ($_uninstallExitCode -eq 0) {
    Add-OperationResult -Results $Results -Target $Id -Source 'WinGet' -Action 'Uninstall' -Status 'Removed' -Detail 'WinGet package removed.'
  }
  else {
    $_detail = ($_output | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($_detail)) {
      $_detail = "winget uninstall exited with $_uninstallExitCode."
    }
    Add-OperationResult -Results $Results -Target $Id -Source 'WinGet' -Action 'Uninstall' -Status 'Failed' -Detail $_detail
  }
}

function Invoke-CopilotArpFallback {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results
  )

  $_program = Find-Win32Program -RegistryPath 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Copilot'
  if (-not $_program) {
    Add-OperationResult -Results $Results -Target 'Microsoft Copilot' -Source 'Win32Program' -Action 'UninstallFallback' -Status 'Skipped' -Detail 'NoMatch'
    return
  }

  $_command = $_program.QuietUninstallString
  if ([string]::IsNullOrWhiteSpace($_command)) {
    $_command = $_program.UninstallString
  }

  if ([string]::IsNullOrWhiteSpace($_command)) {
    Add-OperationResult -Results $Results -Target $_program.DisplayName -Source 'Win32Program' -Action 'UninstallFallback' -Status 'Skipped' -Detail 'NoUninstallString'
    return
  }

  if ($_command -notmatch '--force-uninstall') {
    $_command = "$_command --force-uninstall"
  }

  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $_program.DisplayName -Source 'Win32Program' -Action 'UninstallFallback' -Status 'Skipped' -Detail "DryRun: $_command"
    return
  }

  if (-not $PSCmdlet.ShouldProcess($_program.DisplayName, "Run uninstall command: $_command")) {
    Add-OperationResult -Results $Results -Target $_program.DisplayName -Source 'Win32Program' -Action 'UninstallFallback' -Status 'Skipped' -Detail 'WhatIf'
    return
  }

  if ($_command -match '^\s*"([^"]+)"\s*(.*)$') {
    $_file = $Matches[1]
    $_args = $Matches[2]
  }
  else {
    $_parts = $_command.Trim() -split '\s+', 2
    $_file = $_parts[0]
    $_args = if ($_parts.Count -gt 1) { $_parts[1] } else { '' }
  }

  try {
    $_process = Start-Process -FilePath $_file -ArgumentList $_args -Wait -PassThru -ErrorAction Stop
    Add-OperationResult -Results $Results -Target $_program.DisplayName -Source 'Win32Program' -Action 'UninstallFallback' -Status "ExitCode:$($_process.ExitCode)" -Detail 'ARP uninstall command executed.'
  }
  catch {
    Add-OperationResult -Results $Results -Target $_program.DisplayName -Source 'Win32Program' -Action 'UninstallFallback' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Invoke-BloatwareRegistryPolicies {
  param (
    [AllowEmptyCollection()]
    [System.Collections.ArrayList]$Results,
    [string[]]$Patterns,
    [switch]$ApplyStartPins
  )

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('Microsoft.OutlookForWindows*')) {
    Invoke-BloatwareRegistryRemove -Results $Results -Path 'HKLM:\Software\Microsoft\WindowsUpdate\Orchestrator\UScheduler_Oobe' -Name 'OutlookUpdate' -Detail 'Prevent new Outlook scheduled reinstall.'
  }

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('Microsoft.Copilot*', 'Microsoft.Windows.Ai.Copilot.Provider*')) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Windows\WindowsCopilot' -Name 'TurnOffWindowsCopilot' -Value 1 -Type DWord -Detail 'Disable Windows Copilot machine policy.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot' -Name 'TurnOffWindowsCopilot' -Value 1 -Type DWord -Detail 'Disable Windows Copilot current-user policy.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name 'HubsSidebarEnabled' -Value 0 -Type DWord -Detail 'Disable Edge sidebar/Copilot surface.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name 'CopilotPageContext' -Value 0 -Type DWord -Detail 'Disable Edge Copilot page context.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name 'GenAIDataAnalysisEnabled' -Value 0 -Type DWord -Detail 'Disable Edge GenAI data analysis.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Edge' -Name 'GenAILocalFoundationalModelSettings' -Value 1 -Type DWord -Detail 'Disable Edge local GenAI model settings.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\AdvancedKeyboardSettings' -Name 'HasOverriddenCopilotKey' -Value 1 -Type DWord -Detail 'Mark Copilot key as overridden.'
  }

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('Microsoft.Xbox*', 'Microsoft.GamingApp*')) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR' -Name 'AppCaptureEnabled' -Value 0 -Type DWord -Detail 'Disable Game DVR capture.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Windows\GameDVR' -Name 'AllowGameDVR' -Value 0 -Type DWord -Detail 'Disable Game DVR policy.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKCU:\System\GameConfigStore' -Name 'GameDVR_Enabled' -Value 0 -Type DWord -Detail 'Disable Game DVR user setting.'
  }

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('MicrosoftTeams*', 'MSTeams*')) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Windows\Windows Chat' -Name 'ChatIcon' -Value 3 -Type DWord -Detail 'Hide consumer Teams Chat taskbar icon.'
  }

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('Microsoft.YourPhone*')) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Windows\System' -Name 'EnableMmx' -Value 0 -Type DWord -Detail 'Disable Cross Device phone integration.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\CDP' -Name 'CdpSessionUserAuthzPolicy' -Value 0 -Type DWord -Detail 'Disable current-user CDP phone integration.'
  }

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('Microsoft.MicrosoftOfficeHub*')) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Office\16.0\Common\OfficeUpdate' -Name 'HideEnableDisableUpdates' -Value 1 -Type DWord -Detail 'Best-effort OfficeHub restage suppression.'
  }

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('MicrosoftWindows.Client.WebExperience*')) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Dsh' -Name 'AllowNewsAndInterests' -Value 0 -Type DWord -Detail 'Disable Widgets board when Web Experience remains installed.'
  }

  if (Test-BloatwarePattern -Patterns $Patterns -Match @('*WhatsApp*', '*Disney*', '*Spotify*', '*Netflix*', '*Instagram*', '*TikTok*', '*CandyCrush*', '*Roblox*', '*Duolingo*', '*Flipboard*', '*LinkedIn*', '*Pandora*')) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Windows\CloudContent' -Name 'DisableCloudOptimizedContent' -Value 1 -Type DWord -Detail 'Prevent promoted Start pins for future profiles.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Windows\CloudContent' -Name 'DisableConsumerAccountStateContent' -Value 1 -Type DWord -Detail 'Prevent account-state promoted content.'
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Policies\Microsoft\Windows\CloudContent' -Name 'DisableWindowsConsumerFeatures' -Value 1 -Type DWord -Detail 'Prevent Windows consumer feature promotions.'
  }

  if ($ApplyStartPins) {
    Invoke-BloatwareRegistrySet -Results $Results -Path 'HKLM:\Software\Microsoft\PolicyManager\current\device\Start\ConfigureStartPins' -Name 'value' -Value '{"pinnedList":[]}' -Type String -Detail 'Configure empty Start pinned list for future profiles.'
  }
}

if ($PSBoundParameters.ContainsKey('Config')) {
  if ([string]::IsNullOrWhiteSpace($Config)) {
    Write-Log -Message '-Config requires a PSD1 or JSON file path.' -Color Red
    exit 1
  }

  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) {
    Write-Log -Message "Config file not found: $_configPath" -Color Red
    exit 1
  }

  try {
    $_extension = [System.IO.Path]::GetExtension($_configPath)
    if ($_extension -eq '.psd1') {
      $_configGroups = Import-PowerShellDataFile -Path $_configPath
    }
    else {
      $_configGroups = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $_configPath -Raw -ErrorAction Stop) -ErrorAction Stop
    }

    foreach ($_property in $_configGroups.PSObject.Properties) {
      $_packageGroups[$_property.Name] = @($_property.Value)
    }
  }
  catch {
    Write-Log -Message "Failed to import config '$_configPath': $_" -Color Red
    exit 1
  }
}

if ($ExportConfig) {
  if ($DryRun) {
    Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red
    exit 1
  }
  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_packageGroups | ConvertTo-Json -Depth 4 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default bloatware package groups exported to: $_exportPath" -Color Green
  }
  else {
    $_packageGroups | ConvertTo-Json -Depth 4
  }
  exit 0
}

$_patterns = New-Object System.Collections.ArrayList
foreach ($_group in $Group) {
  if (-not $_packageGroups.ContainsKey($_group)) {
    Write-Log -Message "Unknown package group: $_group" -Color Red
    Write-Log -Message "Available groups: $($_packageGroups.Keys -join ', ')" -Color Gray
    exit 1
  }

  foreach ($_entry in @($_packageGroups[$_group])) {
    [void]$_patterns.Add($_entry)
  }
}

foreach ($_entry in @($Pattern)) {
  if (-not [string]::IsNullOrWhiteSpace($_entry)) {
    [void]$_patterns.Add($_entry)
  }
}

$_patterns = @($_patterns | Sort-Object -Unique)
if ($_patterns.Count -eq 0) {
  Write-Log -Message 'No package patterns selected.' -Color Yellow
  exit 0
}

Write-Log -Message "Selected bloatware groups: $($Group -join ', ')" -Color Yellow
Write-Log -Message "Pattern count: $($_patterns.Count)" -Color Gray

if (-not $SkipProvisioned) {
  Write-Log -Message 'Pass 1/7: removing provisioned UPF AppX/MSIX packages ...' -Color Yellow
  if (-not (Test-Elevation)) {
    Add-OperationResult -Results $_results -Target 'ProvisionedPackages' -Source 'UPFAppxPackage' -Action 'Uninstall' -Status 'Skipped' -Detail 'RequiresElevation'
  }
  else {
    $_provisionedMatches = @(Find-UPFAppxPackage -Pattern $_patterns -Installed:$false -Provisioned -IncludeProtected:$IncludeProtected)
    $_provisionedMatched = @($_provisionedMatches | Where-Object { $_.Matched -and $_.Package })
    $_provisionedOnly = @($_provisionedMatched | Where-Object { $_.Package.Source -eq 'Provisioned' })
    Write-Log -Message "Provisioned matches: $($_provisionedOnly.Count)" -Color Gray

    if ($DryRun -and $_provisionedOnly.Count -gt 0) {
      $_provisionedPreview = $_provisionedOnly | Select-Object Pattern, Protected, ProtectedReason, @{
        Name = 'PackageName'
        Expression = { $_.Package.PackageName }
      }
      $_provisionedPreview | Format-Table -AutoSize
    }

    Invoke-BloatwarePackageSetRemove -Results $_results -Patterns $_patterns -BatchTarget 'ProvisionedPackages' -BatchAction "Remove $($_provisionedOnly.Count) provisioned UPF AppX/MSIX package match(es)" -MatchCount $_provisionedOnly.Count -Installed:$false -Provisioned -IncludeProtected:$IncludeProtected -Force:$Force -DryRun:$DryRun
  }
}
else {
  Add-OperationResult -Results $_results -Target 'ProvisionedPackages' -Source 'UPFAppxPackage' -Action 'Uninstall' -Status 'Skipped' -Detail 'SkipProvisioned'
}

Write-Log -Message 'Pass 2/7: removing installed UPF AppX/MSIX packages ...' -Color Yellow
if ($AllUsers -and -not (Test-Elevation)) {
  Add-OperationResult -Results $_results -Target 'InstalledPackages' -Source 'UPFAppxPackage' -Action 'Uninstall' -Status 'Skipped' -Detail 'AllUsersRequiresElevation'
}
else {
  $_installedMatches = @(Find-UPFAppxPackage -Pattern $_patterns -Installed -AllUsers:$AllUsers -Provisioned:$false -IncludeProtected:$IncludeProtected)
  $_installedMatched = @($_installedMatches | Where-Object { $_.Matched -and $_.Package })
  $_skippedProtected = @($_installedMatches | Where-Object { $_.Protected })
  $_unmatched = @($_installedMatches | Where-Object { -not $_.Matched })
  Write-Log -Message "Installed matches: $($_installedMatched.Count) | Protected matches: $($_skippedProtected.Count) | Unmatched patterns: $($_unmatched.Count)" -Color Gray

  if ($DryRun -and $_installedMatched.Count -gt 0) {
    $_installedPreview = $_installedMatched | Select-Object Pattern, Protected, ProtectedReason, @{
      Name = 'PackageName'
      Expression = { $_.Package.PackageName }
    }
    $_installedPreview | Format-Table -AutoSize
  }

  Invoke-BloatwarePackageSetRemove -Results $_results -Patterns $_patterns -BatchTarget 'InstalledPackages' -BatchAction "Remove $($_installedMatched.Count) installed UPF AppX/MSIX package match(es)" -MatchCount $_installedMatched.Count -Installed -Provisioned:$false -AllUsers:$AllUsers -IncludeProtected:$IncludeProtected -Force:$Force -DryRun:$DryRun
}

if (-not $SkipWinGet) {
  Write-Log -Message 'Pass 3/7: removing WinGet/ARP bloatware targets ...' -Color Yellow
  $_wingetTargets = @()
  if (Test-BloatwarePattern -Patterns $_patterns -Match @('Microsoft.Copilot*', 'Microsoft.Windows.Ai.Copilot.Provider*')) {
    $_wingetTargets += 'ARP\Machine\X86\Microsoft Copilot'
  }

  foreach ($_target in @($_wingetTargets | Sort-Object -Unique)) {
    Invoke-BloatwareWinGetRemove -Results $_results -Id $_target
  }

  if (Test-BloatwarePattern -Patterns $_patterns -Match @('Microsoft.Copilot*', 'Microsoft.Windows.Ai.Copilot.Provider*')) {
    Invoke-CopilotArpFallback -Results $_results
  }
}
else {
  Add-OperationResult -Results $_results -Target 'WinGetTargets' -Source 'WinGet' -Action 'Uninstall' -Status 'Skipped' -Detail 'SkipWinGet'
}

if (-not $SkipSpecialProcedures) {
  Write-Log -Message 'Pass 4/7: removing Windows capabilities ...' -Color Yellow
  if (Test-BloatwarePattern -Patterns $_patterns -Match @('MicrosoftCorporationII.QuickAssist*')) {
    Invoke-BloatwareCapabilityRemove -Results $_results -Pattern 'App.Support.QuickAssist*'
  }

  Write-Log -Message 'Pass 5/7: disabling scheduled tasks ...' -Color Yellow
  if (Test-BloatwarePattern -Patterns $_patterns -Match @('Microsoft.WindowsMaps*')) {
    Invoke-BloatwareTaskDisable -Results $_results -TaskPath '\Microsoft\Windows\Maps\' -TaskName 'MapsToastTask'
    Invoke-BloatwareTaskDisable -Results $_results -TaskPath '\Microsoft\Windows\Maps\' -TaskName 'MapsUpdateTask'
  }

  Write-Log -Message 'Pass 6/7: applying registry policy follow-ups ...' -Color Yellow
  Invoke-BloatwareRegistryPolicies -Results $_results -Patterns $_patterns -ApplyStartPins:$ConfigureStartPins

  Write-Log -Message 'Pass 7/7: removing safe shortcut leftovers ...' -Color Yellow
  if (Test-BloatwarePattern -Patterns $_patterns -Match @('Microsoft.Copilot*', 'Microsoft.Windows.Ai.Copilot.Provider*')) {
    Invoke-BloatwareShortcutRemove -Results $_results -Path '%ProgramData%\Microsoft\Windows\Start Menu\Programs\Copilot.lnk'
    Invoke-BloatwareShortcutRemove -Results $_results -Path '%APPDATA%\Microsoft\Windows\Start Menu\Programs\Copilot.lnk'
  }
}
else {
  Add-OperationResult -Results $_results -Target 'SpecialProcedures' -Source 'Bloatware' -Action 'FollowUp' -Status 'Skipped' -Detail 'SkipSpecialProcedures'
}

$_removed = @($_results | Where-Object { $_.Status -in @('Removed', 'Disabled') }).Count
$_skipped = @($_results | Where-Object { $_.Status -eq 'Skipped' }).Count
$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
Write-Log -Message "Bloatware removal complete. Removed/disabled: $_removed | Skipped: $_skipped | Failed: $_failed" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Remove-Bloatware'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}
Write-Log -Message 'Sign out/in or reboot before judging Settings > Installed apps; that catalogue is cached.' -Color Gray

if ($PassThru -or $DryRun) {
  $_results
}
