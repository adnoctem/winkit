#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', 'Read-JsonFileWithComments', Justification = 'The helper intentionally strips multiple JSON comment lines.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', 'Get-TerminalProfileEntries', Justification = 'The helper returns context-menu entries, not a single Terminal profile object.')]

<#
.SYNOPSIS
  Configures Windows Terminal Explorer context menu entries.
.DESCRIPTION
  Adds idempotent Windows Terminal context menu entries for directory and
  directory-background right-clicks. It can optionally build a profile submenu
  from Windows Terminal settings, add a file context menu entry that opens the
  file's parent directory, and launch an editor inside the new terminal.

  By default the entries are written under HKCU:\Software\Classes. Use
  -Scope Machine to write under HKLM:\Software\Classes for all users; this
  requires elevation and will relaunch through Request-AdministratorPrivilege
  when needed.
.PARAMETER Undo
  Remove the registry keys owned by this script.
.PARAMETER Scope
  Write current-user or machine-wide Explorer class registrations.
.PARAMETER Label
  Label for the folder/background context menu entry.
.PARAMETER MenuId
  Registry-safe identifier for keys created by this script.
.PARAMETER IncludeProfiles
  Build a submenu from Windows Terminal profiles instead of a single direct
  "open here" command.
.PARAMETER IncludeFiles
  Add a file context menu entry for all file types.
.PARAMETER FileLabel
  Label for the file context menu entry.
.PARAMETER EditorCommand
  Optional editor command to run inside Windows Terminal for file clicks. Use
  "{file}" as a placeholder for the selected file. If omitted, file clicks open
  a terminal in the selected file's parent directory.
.PARAMETER Config
  Optional JSON config containing an array of profile menu overrides. Each
  entry may set Guid, Name, Label, Icon, and Hidden. Relative Icon paths are
  resolved relative to the config file.
.PARAMETER ExportConfig
  Export a profile menu config template generated from Windows Terminal
  settings. Cannot be combined with -DryRun.
.PARAMETER ExportPath
  When used with -ExportConfig, writes the JSON template to this path instead
  of printing it to the console.
.PARAMETER Extended
  Show entries only when Shift is held.
.PARAMETER TerminalSettingsPath
  Explicit Windows Terminal settings.json path. When omitted, stable and
  preview package locations are probed.
.PARAMETER TerminalCommand
  Terminal executable to launch. Defaults to wt.exe or the discovered wt.exe.
.PARAMETER DryRun
  Preview changes without applying them.
.PARAMETER PassThru
  Return structured operation result objects.
.EXAMPLE
  PS> .\Set-TerminalContextMenu.ps1
  Adds "Open in Windows Terminal" for folders and folder backgrounds.
.EXAMPLE
  PS> .\Set-TerminalContextMenu.ps1 -ExportConfig -ExportPath .\terminal-context-menu.json
  Exports a profile submenu config template with empty Icon fields.
.EXAMPLE
  PS> .\Set-TerminalContextMenu.ps1 -IncludeProfiles -Config .\terminal-context-menu.json
  Adds a profile submenu, using labels and optional icons from the config.
.EXAMPLE
  PS> .\Set-TerminalContextMenu.ps1 -IncludeFiles -EditorCommand '$env:EDITOR {file}'
  Adds a file context menu entry that opens the selected file in $env:EDITOR
  inside Windows Terminal.
.EXAMPLE
  PS> .\Set-TerminalContextMenu.ps1 -Undo
  Removes the context menu keys created by this script.
.LINK
  https://github.com/adnoctem/winkit
.LINK
  https://github.com/kerol2r20/Windows-terminal-context-menu
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT

  Idea reference:
  The Windows Terminal context-menu concept used here was inspired by
  kerol2r20/Windows-terminal-context-menu. This script keeps the registry
  integration native to winkit and adds profile/config/result handling around
  that original Explorer context-menu idea.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [switch]
  $Undo,

  [ValidateSet('CurrentUser', 'Machine')]
  [string]
  $Scope = 'CurrentUser',

  [string]
  $Label = 'Open in Windows Terminal',

  [ValidatePattern('^[A-Za-z0-9_.-]+$')]
  [string]
  $MenuId = 'WinkitWindowsTerminal',

  [switch]
  $IncludeProfiles,

  [switch]
  $IncludeFiles,

  [string]
  $FileLabel,

  [string]
  $EditorCommand,

  [string]
  $Config,

  [switch]
  $ExportConfig,

  [string]
  $ExportPath,

  [switch]
  $Extended,

  [string]
  $TerminalSettingsPath,

  [string]
  $TerminalCommand,

  [switch]
  $DryRun,

  [switch]
  $PassThru,

  [Parameter(DontShow = $true)]
  [string]
  $OpenFile,

  [Parameter(DontShow = $true)]
  [switch]
  $Elevated
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

function Resolve-TerminalCommand {
  param([string]$Preferred)

  if (-not [string]::IsNullOrWhiteSpace($Preferred)) {
    return $Preferred
  }

  $_command = Get-Command -Name wt.exe -ErrorAction SilentlyContinue
  if ($_command) {
    return $_command.Source
  }

  return 'wt.exe'
}

function ConvertTo-RegistryCommandArgument {
  param([string]$Value)

  return '"' + ($Value -replace '"', '\"') + '"'
}

function ConvertTo-PowerShellSingleQuotedString {
  param([string]$Value)

  return "'" + ($Value -replace "'", "''") + "'"
}

function Start-TerminalForFile {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal context-menu dispatch helper; the setup script itself supports ShouldProcess.')]
  param(
    [string]$Path,
    [string]$Terminal,
    [string]$Editor
  )

  if ([string]::IsNullOrWhiteSpace($Path)) {
    exit 1
  }

  $_filePath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
  $_directory = Split-Path -LiteralPath $_filePath -Parent
  $_arguments = [System.Collections.Generic.List[string]]::new()
  $_arguments.Add('-d')
  $_arguments.Add($_directory)

  if (-not [string]::IsNullOrWhiteSpace($Editor)) {
    $_escapedFile = ConvertTo-PowerShellSingleQuotedString -Value $_filePath
    $_editorCommand = if ($Editor -like '*{file}*') {
      $Editor.Replace('{file}', $_escapedFile)
    }
    else {
      "$Editor $_escapedFile"
    }

    $_arguments.Add('--')
    $_arguments.Add('pwsh.exe')
    $_arguments.Add('-NoExit')
    $_arguments.Add('-Command')
    $_arguments.Add($_editorCommand)
  }

  Start-Process -FilePath $Terminal -ArgumentList $_arguments.ToArray()
}

function Read-JsonFileWithComments {
  param([string]$Path)

  if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
    return $null
  }

  $_content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
  $_content = ($_content -split "`r?`n" | Where-Object { $_ -notmatch '^\s*//' }) -join "`n"
  return ConvertFrom-Json -InputObject $_content -ErrorAction Stop
}

function Resolve-TerminalSettingsPath {
  param([string]$ExplicitPath)

  if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
    if (Test-Path -LiteralPath $ExplicitPath) {
      return (Resolve-Path -LiteralPath $ExplicitPath).ProviderPath
    }
    return $ExplicitPath
  }

  $_candidates = @(
    Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json'
    Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json'
  )

  foreach ($_candidate in $_candidates) {
    if (Test-Path -LiteralPath $_candidate -ErrorAction SilentlyContinue) {
      return $_candidate
    }
  }

  return $null
}

function Resolve-ContextMenuIcon {
  param(
    [string]$Icon,
    [string]$ConfigRoot,
    [string]$Fallback
  )

  if (-not [string]::IsNullOrWhiteSpace($Icon)) {
    $_expanded = [Environment]::ExpandEnvironmentVariables($Icon)
    if ([System.IO.Path]::IsPathRooted($_expanded) -and (Test-Path -LiteralPath $_expanded -ErrorAction SilentlyContinue)) {
      return $_expanded
    }

    if (-not [string]::IsNullOrWhiteSpace($ConfigRoot)) {
      $_candidate = Join-Path -Path $ConfigRoot -ChildPath $_expanded
      if (Test-Path -LiteralPath $_candidate -ErrorAction SilentlyContinue) {
        return $_candidate
      }
    }
  }

  return $Fallback
}

function Get-TerminalProfileList {
  param(
    [string]$SettingsPath
  )

  if ([string]::IsNullOrWhiteSpace($SettingsPath) -or -not (Test-Path -LiteralPath $SettingsPath -ErrorAction SilentlyContinue)) {
    return @()
  }

  $_settings = Read-JsonFileWithComments -Path $SettingsPath
  if (-not $_settings) { return @() }

  $_profiles = if ($_settings.profiles.PSObject.Properties.Name -contains 'list') {
    @($_settings.profiles.list)
  }
  else {
    @($_settings.profiles)
  }

  return @($_profiles)
}

function ConvertTo-TerminalProfileConfig {
  param([object[]]$Profiles)

  $_entries = New-Object System.Collections.ArrayList
  foreach ($_profile in $Profiles) {
    if (-not $_profile) { continue }
    $_props = $_profile.PSObject.Properties.Name
    $_guid = if ($_props -contains 'guid') { [string]$_profile.guid } else { $null }
    $_name = if ($_props -contains 'name') { [string]$_profile.name } else { $null }
    $_hidden = ($_props -contains 'hidden') -and [bool]$_profile.hidden

    if ([string]::IsNullOrWhiteSpace($_name)) { continue }

    [void]$_entries.Add([PSCustomObject]@{
        Guid = $_guid
        Name = $_name
        Label = $_name
        Icon = ''
        Hidden = $_hidden
      })
  }

  return @($_entries)
}

function Merge-TerminalProfileConfig {
  param(
    [object[]]$Base,
    [object[]]$Overrides
  )

  if (-not $Overrides) { return @($Base) }

  $_merged = New-Object System.Collections.ArrayList
  foreach ($_entry in $Base) {
    $_entryProps = $_entry.PSObject.Properties.Name
    $_entryGuid = if ($_entryProps -contains 'Guid') { [string]$_entry.Guid } else { $null }
    $_entryName = if ($_entryProps -contains 'Name') { [string]$_entry.Name } else { $null }
    $_override = $null

    foreach ($_candidate in $Overrides) {
      if (-not $_candidate) { continue }
      $_candidateProps = $_candidate.PSObject.Properties.Name
      $_candidateGuid = if ($_candidateProps -contains 'Guid') { [string]$_candidate.Guid } else { $null }
      $_candidateName = if ($_candidateProps -contains 'Name') { [string]$_candidate.Name } else { $null }

      if (-not [string]::IsNullOrWhiteSpace($_entryGuid) -and $_candidateGuid -eq $_entryGuid) {
        $_override = $_candidate
        break
      }
      if ([string]::IsNullOrWhiteSpace($_entryGuid) -and -not [string]::IsNullOrWhiteSpace($_entryName) -and $_candidateName -eq $_entryName) {
        $_override = $_candidate
        break
      }
    }

    $_copy = [ordered]@{}
    foreach ($_property in $_entry.PSObject.Properties) {
      $_copy[$_property.Name] = $_property.Value
    }

    if ($_override) {
      foreach ($_property in $_override.PSObject.Properties) {
        if ($_copy.Contains($_property.Name)) {
          $_copy[$_property.Name] = $_property.Value
        }
      }
    }

    [void]$_merged.Add([PSCustomObject]$_copy)
  }

  return @($_merged)
}

function Get-TerminalProfileEntries {
  param(
    [object[]]$ProfileSettings,
    [string]$ConfigRootPath,
    [string]$DefaultIcon
  )

  $_entries = New-Object System.Collections.ArrayList
  foreach ($_profile in $ProfileSettings) {
    if (-not $_profile) { continue }
    $_props = $_profile.PSObject.Properties.Name
    $_hidden = ($_props -contains 'Hidden') -and [bool]$_profile.Hidden
    $_name = if ($_props -contains 'Name') { [string]$_profile.Name } else { $null }
    if ($_hidden -or [string]::IsNullOrWhiteSpace($_name)) { continue }

    $_label = if ($_props -contains 'Label' -and -not [string]::IsNullOrWhiteSpace([string]$_profile.Label)) {
      [string]$_profile.Label
    }
    else {
      $_name
    }

    $_iconValue = if ($_props -contains 'Icon') { [string]$_profile.Icon } else { '' }
    $_icon = Resolve-ContextMenuIcon -Icon $_iconValue -ConfigRoot $ConfigRootPath -Fallback $DefaultIcon

    [void]$_entries.Add([PSCustomObject]@{
        Name = $_name
        Label = $_label
        Icon = $_icon
      })
  }

  return @($_entries)
}

function Add-ResultFromRegistryOperation {
  param(
    [System.Collections.IList]$Results,
    [string]$Target,
    [string]$Action,
    [object]$Operation,
    [string]$Detail
  )

  $_status = if ($Operation -and $Operation.Status) { $Operation.Status } else { 'Failed' }
  Add-OperationResult -Results $Results -Target $Target -Source 'Registry' -Scope $Scope -Action $Action -Status $_status -Detail $Detail
}

function Set-ContextMenuValue {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper delegates registry writes to winkit ShouldProcess-aware helpers.')]
  [CmdletBinding()]
  param(
    [System.Collections.IList]$Results,
    [string]$Path,
    [string]$Name = '',
    [AllowNull()]
    [object]$Value,
    [Microsoft.Win32.RegistryValueKind]$Type = [Microsoft.Win32.RegistryValueKind]::String
  )

  if ($DryRun) {
    Add-OperationResult -Results $Results -Target "$Path\$Name" -Source 'Registry' -Scope $Scope -Action 'SetValue' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  $_operation = Set-RegistryValue -Path $Path -Name $Name -Value $Value -Type $Type -WhatIf:$WhatIfPreference
  Add-ResultFromRegistryOperation -Results $Results -Target "$Path\$Name" -Action 'SetValue' -Operation $_operation -Detail $Value
}

function Remove-ContextMenuKey {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper delegates registry removal to winkit ShouldProcess-aware helpers.')]
  [CmdletBinding()]
  param(
    [System.Collections.IList]$Results,
    [string]$Path
  )

  if ($DryRun) {
    Add-OperationResult -Results $Results -Target $Path -Source 'Registry' -Scope $Scope -Action 'RemoveKey' -Status 'Skipped' -Detail 'DryRun'
    return
  }

  $_operation = Remove-RegistryKey -Path $Path -Recurse -WhatIf:$WhatIfPreference
  Add-ResultFromRegistryOperation -Results $Results -Target $Path -Action 'RemoveKey' -Operation $_operation -Detail 'Terminal context menu key'
}

function Set-DirectCommandMenu {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper delegates registry writes to winkit ShouldProcess-aware helpers.')]
  param(
    [System.Collections.IList]$Results,
    [string]$Path,
    [string]$MenuLabel,
    [string]$Icon,
    [string]$Command
  )

  Set-ContextMenuValue -Results $Results -Path $Path -Name 'MUIVerb' -Value $MenuLabel
  Set-ContextMenuValue -Results $Results -Path $Path -Name 'Icon' -Value $Icon
  if ($Extended) {
    Set-ContextMenuValue -Results $Results -Path $Path -Name 'Extended' -Value ''
  }
  else {
    $null = Remove-RegistryValue -Path $Path -Name 'Extended' -WhatIf:$WhatIfPreference -ErrorAction SilentlyContinue
  }
  Set-ContextMenuValue -Results $Results -Path "$Path\command" -Name '' -Value $Command
}

function Set-SubmenuRoot {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper delegates registry writes to winkit ShouldProcess-aware helpers.')]
  param(
    [System.Collections.IList]$Results,
    [string]$Path,
    [string]$MenuLabel,
    [string]$Icon,
    [string]$SubCommandsKey
  )

  Set-ContextMenuValue -Results $Results -Path $Path -Name 'MUIVerb' -Value $MenuLabel
  Set-ContextMenuValue -Results $Results -Path $Path -Name 'Icon' -Value $Icon
  Set-ContextMenuValue -Results $Results -Path $Path -Name 'ExtendedSubCommandsKey' -Value $SubCommandsKey
  if ($Extended) {
    Set-ContextMenuValue -Results $Results -Path $Path -Name 'Extended' -Value ''
  }
  else {
    $null = Remove-RegistryValue -Path $Path -Name 'Extended' -WhatIf:$WhatIfPreference -ErrorAction SilentlyContinue
  }
}

function Set-SubmenuCommand {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Internal helper delegates registry writes to winkit ShouldProcess-aware helpers.')]
  param(
    [System.Collections.IList]$Results,
    [string]$Path,
    [string]$MenuLabel,
    [string]$Icon,
    [string]$Command
  )

  Set-ContextMenuValue -Results $Results -Path $Path -Name 'MUIVerb' -Value $MenuLabel
  Set-ContextMenuValue -Results $Results -Path $Path -Name 'Icon' -Value $Icon
  Set-ContextMenuValue -Results $Results -Path "$Path\command" -Name '' -Value $Command
}

$_terminal = Resolve-TerminalCommand -Preferred $TerminalCommand

if (-not [string]::IsNullOrWhiteSpace($OpenFile)) {
  Start-TerminalForFile -Path $OpenFile -Terminal $_terminal -Editor $EditorCommand
  exit 0
}

if ($Scope -eq 'Machine' -and -not (Test-Elevation)) {
  Request-AdministratorPrivilege -BoundParameters $PSBoundParameters -ArgumentList $args -IsElevatedRelaunch:$Elevated
}

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no terminal context menu changes will be applied`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_classesRoot = if ($Scope -eq 'Machine') { 'HKLM:\Software\Classes' } else { 'HKCU:\Software\Classes' }
$_directoryMenuPath = "$_classesRoot\Directory\shell\$MenuId"
$_backgroundMenuPath = "$_classesRoot\Directory\Background\shell\$MenuId"
$_fileMenuPath = "$_classesRoot\*\shell\$MenuId"
$_submenuRelativePath = "Directory\ContextMenus\$MenuId"
$_submenuRoot = "$_classesRoot\Directory\ContextMenus\$MenuId"
$_submenuShell = "$_submenuRoot\shell"
$_scriptPath = $PSCommandPath
$_terminalIcon = if ($_terminal -match '\.exe$' -or $_terminal -match '\\') { "$_terminal,0" } else { 'wt.exe,0' }
$_settingsPath = Resolve-TerminalSettingsPath -ExplicitPath $TerminalSettingsPath
$_terminalProfiles = @(Get-TerminalProfileList -SettingsPath $_settingsPath)
$_profileSettings = @(ConvertTo-TerminalProfileConfig -Profiles $_terminalProfiles)

if ($ExportConfig) {
  if ($DryRun) {
    Write-Log -Message '-DryRun cannot be combined with -ExportConfig.' -Color Red
    exit 1
  }

  if ($PSBoundParameters.ContainsKey('ExportPath') -and -not [string]::IsNullOrWhiteSpace($ExportPath)) {
    $_exportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExportPath)
    $_profileSettings | ConvertTo-Json -Depth 4 | Out-File -FilePath $_exportPath -Encoding utf8
    Write-Log -Message "Default Terminal context menu config exported to: $_exportPath" -Color Green
  }
  else {
    $_profileSettings | ConvertTo-Json -Depth 4
  }

  return
}

$_configRoot = $null
if ($Config) {
  $_configPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config)
  if (-not (Test-Path -LiteralPath $_configPath)) {
    Write-Log -Message "Config file not found: $_configPath" -Color Red
    exit 1
  }
  $_configRoot = Split-Path -Path $_configPath -Parent
  $_overrides = Read-JsonFileWithComments -Path $_configPath
  if ($_overrides -and $_overrides -isnot [array]) {
    $_overrides = @($_overrides)
  }
  $_profileSettings = @(Merge-TerminalProfileConfig -Base $_profileSettings -Overrides $_overrides)
}

if ($Undo) {
  foreach ($_path in @($_directoryMenuPath, $_backgroundMenuPath, $_fileMenuPath, $_submenuRoot)) {
    Remove-ContextMenuKey -Results $_results -Path $_path
  }
}
else {
  if ($IncludeProfiles) {
    $_profiles = @(Get-TerminalProfileEntries -ProfileSettings $_profileSettings -ConfigRootPath $_configRoot -DefaultIcon $_terminalIcon)

    Set-SubmenuRoot -Results $_results -Path $_directoryMenuPath -MenuLabel $Label -Icon $_terminalIcon -SubCommandsKey $_submenuRelativePath
    Set-SubmenuRoot -Results $_results -Path $_backgroundMenuPath -MenuLabel $Label -Icon $_terminalIcon -SubCommandsKey $_submenuRelativePath

    if ($_profiles.Count -eq 0) {
      $_defaultCommand = "$(ConvertTo-RegistryCommandArgument $_terminal) -d `"%V\.`""
      Set-SubmenuCommand -Results $_results -Path "$_submenuShell\00OpenHere" -MenuLabel 'Default profile' -Icon $_terminalIcon -Command $_defaultCommand
    }
    else {
      $_index = 0
      foreach ($_profile in $_profiles) {
        $_index++
        $_safeName = ($_profile.Label -replace '[^A-Za-z0-9_.-]', '-').Trim('-')
        if ([string]::IsNullOrWhiteSpace($_safeName)) { $_safeName = "Profile$_index" }
        $_profilePath = "$_submenuShell\$('{0:00}' -f $_index)$_safeName"
        $_command = "$(ConvertTo-RegistryCommandArgument $_terminal) -p `"$($_profile.Name)`" -d `"%V\.`""
        Set-SubmenuCommand -Results $_results -Path $_profilePath -MenuLabel $_profile.Label -Icon $_profile.Icon -Command $_command
      }
    }
  }
  else {
    $_folderCommand = "$(ConvertTo-RegistryCommandArgument $_terminal) -d `"%1`""
    $_backgroundCommand = "$(ConvertTo-RegistryCommandArgument $_terminal) -d `"%V`""
    Set-DirectCommandMenu -Results $_results -Path $_directoryMenuPath -MenuLabel $Label -Icon $_terminalIcon -Command $_folderCommand
    Set-DirectCommandMenu -Results $_results -Path $_backgroundMenuPath -MenuLabel $Label -Icon $_terminalIcon -Command $_backgroundCommand
  }

  if ($IncludeFiles) {
    $_resolvedFileLabel = if (-not [string]::IsNullOrWhiteSpace($FileLabel)) {
      $FileLabel
    }
    elseif (-not [string]::IsNullOrWhiteSpace($EditorCommand)) {
      'Open with editor in Windows Terminal'
    }
    else {
      'Open parent in Windows Terminal'
    }

    $_scriptCommand = @(
      'powershell.exe'
      '-NoProfile'
      '-ExecutionPolicy'
      'Bypass'
      '-File'
      (ConvertTo-RegistryCommandArgument $_scriptPath)
      '-OpenFile'
      '"%1"'
    )

    if (-not [string]::IsNullOrWhiteSpace($TerminalCommand)) {
      $_scriptCommand += '-TerminalCommand'
      $_scriptCommand += (ConvertTo-RegistryCommandArgument $TerminalCommand)
    }
    if (-not [string]::IsNullOrWhiteSpace($EditorCommand)) {
      $_scriptCommand += '-EditorCommand'
      $_scriptCommand += (ConvertTo-RegistryCommandArgument $EditorCommand)
    }

    Set-DirectCommandMenu -Results $_results -Path $_fileMenuPath -MenuLabel $_resolvedFileLabel -Icon $_terminalIcon -Command ($_scriptCommand -join ' ')
  }
}

$_applied = @($_results | Where-Object { $_.Status -in @('Created', 'Updated', 'Applied', 'Removed', 'Moved') }).Count
$_skipped = @($_results | Where-Object { $_.Status -in @('Skipped', 'AlreadyExists', 'NotFound') }).Count
$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
Write-Log -Message "Terminal context menu complete. Applied: $_applied | Skipped: $_skipped | Failed: $_failed" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Set-TerminalContextMenu'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}
