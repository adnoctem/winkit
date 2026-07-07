#Requires -Version 5.0

$script:UPFProtectedPackagePatterns = @(
  'Microsoft.DesktopAppInstaller*',
  'Microsoft.StorePurchaseApp*',
  'Microsoft.WindowsStore*',
  'Microsoft.SecHealthUI*',
  'Microsoft.UI.Xaml*',
  'Microsoft.VCLibs*',
  'Microsoft.NET.Native*',
  'Microsoft.WindowsAppRuntime*',
  'Microsoft.Services.Store.Engagement*',
  'Microsoft.LockApp*',
  'Microsoft.CredDialogHost*',
  'Microsoft.ECApp*'
)

function Test-PackageAdministrator {
  <#
    .SYNOPSIS
      Checks whether the current PowerShell session is running with Administrator rights.
    .DESCRIPTION
      Returns a Boolean indicating whether the current Windows identity belongs
      to the local Administrators role. Package operations that read or mutate
      the system provisioning store use this helper to fail softly when a
      non-elevated shell cannot perform the requested work.
    .EXAMPLE
      PS> Test-PackageAdministrator
      Returns True when the current process is elevated.
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([bool])]
  [CmdletBinding()]
  param ()

  $_identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $_principal = New-Object Security.Principal.WindowsPrincipal($_identity)
  return $_principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-PackageLifecycleResult {
  <#
    .SYNOPSIS
      Creates a normalized result object for package lifecycle operations.
    .DESCRIPTION
      Builds the common structured output used by UPF AppX/MSIX, Win32 program,
      and WinGet-backed helpers. Callers can consume Target, Source, Action,
      Status, SkippedReason, and Error without parsing console text.
    .PARAMETER Target
      Package, program, file path, or registry entry targeted by the operation.
    .PARAMETER Source
      Package source category, such as UPFAppxPackage, Win32Program, WinGet,
      Installed, or Provisioned.
    .PARAMETER Action
      Lifecycle action represented by the result.
    .PARAMETER Status
      Outcome string such as Completed, Removed, Skipped, Failed, or ExitCode:n.
    .PARAMETER SkippedReason
      Optional reason when Status is Skipped.
    .PARAMETER ErrorMessage
      Optional error detail when Status is Failed.
    .EXAMPLE
      PS> New-PackageLifecycleResult -Target 'Microsoft.GetHelp' -Source 'UPFAppxPackage' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'NoMatch'
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Factory helper only creates a result object.')]
  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [string]$Target,
    [string]$Source,
    [string]$Action,
    [string]$Status,
    [string]$SkippedReason,
    [string]$ErrorMessage
  )

  New-OperationResult -Target $Target -Source $Source -Action $Action -Status $Status -SkippedReason $SkippedReason -ErrorMessage $ErrorMessage
}

function Get-InstalledProgramCount {
  <#
    .SYNOPSIS
      Counts visible Win32 programs registered in the Windows uninstall registry keys.
    .DESCRIPTION
      Uses Get-Win32Program to count unique, user-visible classic Win32
      application registrations from HKLM 64-bit, HKLM 32-bit, and HKCU
      uninstall locations. Registry entries marked as SystemComponent are
      excluded by default through the inventory helper.
    .EXAMPLE
      PS> Get-InstalledProgramCount
      Returns the number of registered Win32 programs visible to normal
      uninstall inventory views.
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([int])]
  [CmdletBinding()]
  param ()

  return @(Get-Win32Program).Count
}

function Get-AppxPackageCount {
  <#
    .SYNOPSIS
      Counts installed UPF AppX/MSIX packages for the current user.
    .DESCRIPTION
      Uses Get-UPFAppxPackage to count installed AppX/MSIX packages while
      excluding framework, resource, and bundle packages by default. This keeps
      fetch-style summaries focused on user-facing applications instead of
      runtime dependencies.
    .EXAMPLE
      PS> Get-AppxPackageCount
      Returns the current user's user-facing UPF AppX/MSIX package count.
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([int])]
  [CmdletBinding()]
  param ()

  return @(Get-UPFAppxPackage -Installed).Count
}

function Get-PackageCount {
  <#
    .SYNOPSIS
      Returns reusable Win32 and UPF AppX/MSIX package counts for system summaries.
    .DESCRIPTION
      Combines Get-InstalledProgramCount and Get-AppxPackageCount into a single
      object with Programs, Appx, and Total properties. This is intentionally
      small and fast enough for fetch-style scripts that should not duplicate
      registry or package inventory logic.
    .EXAMPLE
      PS> Get-PackageCount
      Returns Programs, Appx, and Total package counts.
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param ()

  $_programs = Get-InstalledProgramCount
  $_appx = Get-AppxPackageCount

  return [PSCustomObject]@{
    Programs = $_programs
    Appx = $_appx
    Total = ($_programs + $_appx)
  }
}

function Get-Win32Program {
  <#
    .SYNOPSIS
      Returns normalized Win32 program inventory from Windows uninstall registry keys.
    .DESCRIPTION
      Reads machine-wide 64-bit, machine-wide 32-bit and current-user uninstall
      keys using winkit registry helpers, and returns normalized program
      metadata suitable for lifecycle operations.
    .PARAMETER Name
      Optional wildcard DisplayName filter. When omitted, all visible Win32
      uninstall registrations are returned.
    .PARAMETER IncludeSystemComponent
      Include entries marked with SystemComponent=1. These are normally hidden
      from user-facing uninstall inventory and are excluded by default.
    .EXAMPLE
      PS> Get-Win32Program
      Lists visible Win32 programs registered on the machine and current user.
    .EXAMPLE
      PS> Get-Win32Program -Name '*OneDrive*' -IncludeSystemComponent
      Returns OneDrive-related uninstall entries, including hidden system
      component registrations.
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [string]$Name,
    [switch]$IncludeSystemComponent
  )

  $_uninstallRoots = @(
    @{
      Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
      Scope = 'Machine'
      View = '64-bit'
    },
    @{
      Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
      Scope = 'Machine'
      View = '32-bit'
    },
    @{
      Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
      Scope = 'CurrentUser'
      View = 'Default'
    }
  )

  $_programs = foreach ($_root in $_uninstallRoots) {
    $_key = Get-RegistryKey -Path $_root.Path -ErrorAction SilentlyContinue
    if (-not $_key) { continue }

    foreach ($_subKey in $_key.SubKeys) {
      $_path = Join-Path -Path $_root.Path -ChildPath $_subKey
      $_displayName = Get-RegistryValue -Path $_path -Name 'DisplayName' -ErrorAction SilentlyContinue
      if ([string]::IsNullOrWhiteSpace($_displayName)) { continue }
      if ($Name -and $_displayName -notlike $Name) { continue }

      $_systemComponent = Get-RegistryValue -Path $_path -Name 'SystemComponent' -ErrorAction SilentlyContinue
      if ($_systemComponent -and -not $IncludeSystemComponent) { continue }

      [PSCustomObject]@{
        Source = 'Win32Program'
        Name = $_displayName
        DisplayName = $_displayName
        DisplayVersion = Get-RegistryValue -Path $_path -Name 'DisplayVersion' -ErrorAction SilentlyContinue
        Publisher = Get-RegistryValue -Path $_path -Name 'Publisher' -ErrorAction SilentlyContinue
        InstallLocation = Get-RegistryValue -Path $_path -Name 'InstallLocation' -ErrorAction SilentlyContinue
        InstallDate = Get-RegistryValue -Path $_path -Name 'InstallDate' -ErrorAction SilentlyContinue
        UninstallString = Get-RegistryValue -Path $_path -Name 'UninstallString' -ErrorAction SilentlyContinue
        QuietUninstallString = Get-RegistryValue -Path $_path -Name 'QuietUninstallString' -ErrorAction SilentlyContinue
        ModifyPath = Get-RegistryValue -Path $_path -Name 'ModifyPath' -ErrorAction SilentlyContinue
        EstimatedSize = Get-RegistryValue -Path $_path -Name 'EstimatedSize' -ErrorAction SilentlyContinue
        RegistryPath = $_path
        RegistryScope = $_root.Scope
        RegistryView = $_root.View
        SystemComponent = [bool]$_systemComponent
      }
    }
  }

  $_programs | Sort-Object -Property DisplayName, RegistryPath -Unique
}

function Find-Win32Program {
  <#
    .SYNOPSIS
      Finds Win32 program uninstall registrations by wildcard name or exact registry path.
    .DESCRIPTION
      Resolves classic Win32 application registrations from uninstall registry
      locations. Use -Name for broad wildcard matching or -RegistryPath when a
      caller has already selected a concrete uninstall key.
    .PARAMETER Name
      Wildcard DisplayName pattern to resolve.
    .PARAMETER RegistryPath
      Exact uninstall registry key path to normalize into a Win32 program object.
    .PARAMETER IncludeSystemComponent
      Include hidden SystemComponent entries when resolving by name.
    .EXAMPLE
      PS> Find-Win32Program -Name '*Visual Studio Code*'
    .EXAMPLE
      PS> Find-Win32Program -RegistryPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SomeApp'
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
    [string]$Name,

    [Parameter(Mandatory = $true, ParameterSetName = 'RegistryPath')]
    [string]$RegistryPath,

    [switch]$IncludeSystemComponent
  )

  if ($PSCmdlet.ParameterSetName -eq 'RegistryPath') {
    $_displayName = Get-RegistryValue -Path $RegistryPath -Name 'DisplayName' -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($_displayName)) { return }

    return [PSCustomObject]@{
      Source = 'Win32Program'
      Name = $_displayName
      DisplayName = $_displayName
      DisplayVersion = Get-RegistryValue -Path $RegistryPath -Name 'DisplayVersion' -ErrorAction SilentlyContinue
      Publisher = Get-RegistryValue -Path $RegistryPath -Name 'Publisher' -ErrorAction SilentlyContinue
      InstallLocation = Get-RegistryValue -Path $RegistryPath -Name 'InstallLocation' -ErrorAction SilentlyContinue
      InstallDate = Get-RegistryValue -Path $RegistryPath -Name 'InstallDate' -ErrorAction SilentlyContinue
      UninstallString = Get-RegistryValue -Path $RegistryPath -Name 'UninstallString' -ErrorAction SilentlyContinue
      QuietUninstallString = Get-RegistryValue -Path $RegistryPath -Name 'QuietUninstallString' -ErrorAction SilentlyContinue
      ModifyPath = Get-RegistryValue -Path $RegistryPath -Name 'ModifyPath' -ErrorAction SilentlyContinue
      EstimatedSize = Get-RegistryValue -Path $RegistryPath -Name 'EstimatedSize' -ErrorAction SilentlyContinue
      RegistryPath = $RegistryPath
      RegistryScope = $null
      RegistryView = $null
      SystemComponent = [bool](Get-RegistryValue -Path $RegistryPath -Name 'SystemComponent' -ErrorAction SilentlyContinue)
    }
  }

  Get-Win32Program -Name $Name -IncludeSystemComponent:$IncludeSystemComponent
}

function Install-Win32Program {
  <#
    .SYNOPSIS
      Runs a local Win32 installer executable or package with explicit arguments.
    .DESCRIPTION
      Starts a local installer path using Start-Process and returns a normalized
      package lifecycle result. This helper deliberately requires an explicit
      local path and argument list so scripts do not hide downloaded installer
      execution behind a generic package verb.
    .PARAMETER Path
      Local executable or installer package path to run.
    .PARAMETER ArgumentList
      Installer arguments, such as silent install flags.
    .PARAMETER NoWait
      Start the installer and return immediately instead of waiting for it to exit.
    .PARAMETER PassThru
      Return process exit code information in the lifecycle status.
    .PARAMETER DryRun
      Preview the installer invocation without starting the process.
    .EXAMPLE
      PS> Install-Win32Program -Path 'C:\Installers\AppSetup.exe' -ArgumentList '/quiet','/norestart'
    .EXAMPLE
      PS> Install-Win32Program -Path 'C:\Installers\AppSetup.exe' -ArgumentList '/quiet' -DryRun
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$Path,

    [string[]]$ArgumentList,

    [switch]$NoWait,
    [switch]$PassThru,
    [switch]$DryRun
  )

  if ($DryRun) { $WhatIfPreference = $true }

  $_target = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
  $_argumentText = if ($ArgumentList) { $ArgumentList -join ' ' } else { '' }

  if (-not $PSCmdlet.ShouldProcess($_target, "Install Win32 program $_argumentText")) {
    return New-PackageLifecycleResult -Target $_target -Source 'Win32Program' -Action 'Install' -Status 'Skipped' -SkippedReason 'WhatIf'
  }

  try {
    $_params = @{
      FilePath = $_target
      ErrorAction = 'Stop'
    }
    if ($ArgumentList) { $_params.ArgumentList = $ArgumentList }
    if (-not $NoWait) { $_params.Wait = $true }
    if ($PassThru) { $_params.PassThru = $true }

    $_process = Start-Process @_params
    $_status = 'Installed'
    if ($PassThru -and $_process) {
      $_status = "ExitCode:$($_process.ExitCode)"
    }
    New-PackageLifecycleResult -Target $_target -Source 'Win32Program' -Action 'Install' -Status $_status
  }
  catch {
    New-PackageLifecycleResult -Target $_target -Source 'Win32Program' -Action 'Install' -Status 'Failed' -ErrorMessage $_.Exception.Message
  }
}

function Uninstall-Win32Program {
  <#
    .SYNOPSIS
      Uninstalls Win32 programs using registered uninstall commands.
    .DESCRIPTION
      Resolves one or more Win32 uninstall registry entries and runs the
      selected uninstall command. Interactive uninstall strings require -Force;
      quiet uninstall strings are preferred when -Quiet is supplied. MSI
      install commands using /I are converted to /X for uninstall behavior.
    .PARAMETER Name
      Wildcard DisplayName pattern to resolve and uninstall.
    .PARAMETER InputObject
      Win32 program object from Get-Win32Program or Find-Win32Program.
    .PARAMETER Quiet
      Prefer QuietUninstallString when available.
    .PARAMETER Force
      Permit execution of a non-quiet uninstall command.
    .PARAMETER DryRun
      Preview resolved uninstall commands without running them.
    .EXAMPLE
      PS> Uninstall-Win32Program -Name '*OneDrive*' -Quiet -DryRun
    .EXAMPLE
      PS> Find-Win32Program -Name '*Example App*' | Uninstall-Win32Program -Quiet
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
    [string]$Name,

    [Parameter(Mandatory = $true, ParameterSetName = 'InputObject', ValueFromPipeline = $true)]
    [psobject]$InputObject,

    [switch]$Quiet,
    [switch]$Force,
    [switch]$DryRun
  )

  begin {
    if ($DryRun) { $WhatIfPreference = $true }
    $_targets = @()
  }

  process {
    if ($PSCmdlet.ParameterSetName -eq 'InputObject') {
      $_targets += $InputObject
    }
  }

  end {
    if ($PSCmdlet.ParameterSetName -eq 'Name') {
      $_targets = @(Find-Win32Program -Name $Name)
    }

    if ($_targets.Count -eq 0) {
      return New-PackageLifecycleResult -Target $Name -Source 'Win32Program' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'NoMatch'
    }

    foreach ($_program in $_targets) {
      $_command = if ($Quiet -and $_program.QuietUninstallString) { $_program.QuietUninstallString } else { $_program.UninstallString }
      if ([string]::IsNullOrWhiteSpace($_command)) {
        New-PackageLifecycleResult -Target $_program.DisplayName -Source 'Win32Program' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'NoUninstallString'
        continue
      }

      if (-not $Force -and -not $Quiet) {
        New-PackageLifecycleResult -Target $_program.DisplayName -Source 'Win32Program' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'ForceRequired'
        continue
      }

      if (-not $PSCmdlet.ShouldProcess($_program.DisplayName, "Run uninstall command: $_command")) {
        New-PackageLifecycleResult -Target $_program.DisplayName -Source 'Win32Program' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'WhatIf'
        continue
      }

      try {
        if ($_command -match '^\s*"([^"]+)"\s*(.*)$') {
          $_file = $Matches[1]
          $_args = $Matches[2]
        }
        else {
          $_parts = $_command.Trim() -split '\s+', 2
          $_file = $_parts[0]
          $_args = if ($_parts.Count -gt 1) { $_parts[1] } else { '' }
        }

        if ($_file -match '(?i)msiexec(\.exe)?$' -and $_args -match '(?i)\s/I\s*') {
          $_args = $_args -replace '(?i)\s/I\s*', ' /X '
        }

        $_params = @{
          FilePath = $_file
          Wait = $true
          PassThru = $true
          ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($_args)) { $_params.ArgumentList = $_args }
        $_process = Start-Process @_params
        New-PackageLifecycleResult -Target $_program.DisplayName -Source 'Win32Program' -Action 'Uninstall' -Status "ExitCode:$($_process.ExitCode)"
      }
      catch {
        New-PackageLifecycleResult -Target $_program.DisplayName -Source 'Win32Program' -Action 'Uninstall' -Status 'Failed' -ErrorMessage $_.Exception.Message
      }
    }
  }
}

function Get-UPFAppxPackage {
  <#
    .SYNOPSIS
      Returns normalized installed and provisioned UPF AppX/MSIX package inventory.
    .DESCRIPTION
      Wraps Get-AppxPackage and Get-AppxProvisionedPackage into a common object
      shape so callers can reason about installed user packages and provisioned
      Windows image packages together. Framework, resource, and bundle packages
      are hidden by default to keep normal inventory focused on user-facing apps.
    .PARAMETER Name
      App package name or wildcard pattern. Defaults to all packages.
    .PARAMETER Installed
      Include installed packages for the current user, or for all users with
      -AllUsers. If neither -Installed nor -Provisioned is supplied, installed
      current-user inventory is returned.
    .PARAMETER Provisioned
      Include packages provisioned into the online Windows image. This requires
      elevation and returns no entries from a non-elevated session.
    .PARAMETER AllUsers
      Query installed packages for all user profiles. Requires elevation on
      many Windows builds.
    .PARAMETER IncludeFramework
      Include framework packages such as Microsoft.UI.Xaml and VCLibs.
    .PARAMETER IncludeResource
      Include resource packages.
    .PARAMETER IncludeBundle
      Include bundle packages.
    .EXAMPLE
      PS> Get-UPFAppxPackage -Name 'Microsoft.GetHelp'
    .EXAMPLE
      PS> Get-UPFAppxPackage -Name 'Microsoft.Bing*' -Installed -Provisioned
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [string]$Name = '*',
    [switch]$Installed,
    [switch]$Provisioned,
    [switch]$AllUsers,
    [switch]$IncludeFramework,
    [switch]$IncludeResource,
    [switch]$IncludeBundle
  )

  if (-not $Installed -and -not $Provisioned) {
    $Installed = $true
  }

  if ($Installed) {
    $_installed = if ($AllUsers) {
      @(Get-AppxPackage -AllUsers -Name $Name -ErrorAction SilentlyContinue)
    }
    else {
      @(Get-AppxPackage -Name $Name -ErrorAction SilentlyContinue)
    }

    foreach ($_package in $_installed) {
      $_packageProperties = $_package.PSObject.Properties.Name
      $_isFramework = ($_packageProperties -contains 'IsFramework') -and [bool]$_package.IsFramework
      $_isResourcePackage = ($_packageProperties -contains 'IsResourcePackage') -and [bool]$_package.IsResourcePackage
      $_isBundle = ($_packageProperties -contains 'IsBundle') -and [bool]$_package.IsBundle
      $_nonRemovable = ($_packageProperties -contains 'NonRemovable') -and [bool]$_package.NonRemovable
      $_userSecurityId = if ($_packageProperties -contains 'UserSecurityId') { $_package.UserSecurityId } else { $null }

      if (-not $IncludeFramework -and $_isFramework) { continue }
      if (-not $IncludeResource -and $_isResourcePackage) { continue }
      if (-not $IncludeBundle -and $_isBundle) { continue }

      [PSCustomObject]@{
        Source = 'Installed'
        Name = $_package.Name
        DisplayName = $_package.Name
        PackageFullName = $_package.PackageFullName
        PackageName = $_package.PackageFullName
        PackageFamilyName = $_package.PackageFamilyName
        Publisher = $_package.Publisher
        Version = $_package.Version
        Architecture = $_package.Architecture
        IsFramework = $_isFramework
        IsResourcePackage = $_isResourcePackage
        IsBundle = $_isBundle
        NonRemovable = $_nonRemovable
        InstallLocation = $_package.InstallLocation
        User = $_userSecurityId
      }
    }
  }

  if ($Provisioned) {
    if (-not (Test-PackageAdministrator)) {
      Write-Verbose 'Provisioned UPF AppX/MSIX package inventory requires elevation.'
      return
    }

    $_provisioned = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like $Name -or $_.PackageName -like $Name })

    foreach ($_package in $_provisioned) {
      [PSCustomObject]@{
        Source = 'Provisioned'
        Name = $_package.DisplayName
        DisplayName = $_package.DisplayName
        PackageFullName = $null
        PackageName = $_package.PackageName
        PackageFamilyName = $null
        Publisher = $null
        Version = $_package.Version
        Architecture = $_package.Architecture
        IsFramework = $false
        IsResourcePackage = $false
        IsBundle = $false
        NonRemovable = $false
        InstallLocation = $null
        User = $null
      }
    }
  }
}

function Find-UPFAppxPackage {
  <#
    .SYNOPSIS
      Resolves UPF AppX/MSIX wildcard patterns to concrete package inventory records.
    .DESCRIPTION
      Looks up installed and/or provisioned AppX/MSIX packages for each pattern
      and annotates every result with match and protection metadata. Unmatched
      patterns are returned as structured records so removal scripts can report
      what did and did not resolve.
    .PARAMETER Pattern
      One or more package name, full name, family name, display name, or
      provisioned package wildcard patterns.
    .PARAMETER Installed
      Search installed packages.
    .PARAMETER Provisioned
      Search provisioned packages in the online Windows image.
    .PARAMETER AllUsers
      Search installed packages for all users.
    .PARAMETER IncludeProtected
      Return protected package matches as removable candidates instead of
      marking them as protected-only results.
    .PARAMETER IncludeFramework
      Include framework packages during discovery.
    .PARAMETER IncludeResource
      Include resource packages during discovery.
    .PARAMETER IncludeBundle
      Include bundle packages during discovery.
    .EXAMPLE
      PS> Find-UPFAppxPackage -Pattern 'Microsoft.Zune*' -Installed
    .EXAMPLE
      PS> 'Microsoft.GetHelp','Microsoft.BingWeather*' | Find-UPFAppxPackage -Installed
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string[]]$Pattern,

    [switch]$Installed,
    [switch]$Provisioned,
    [switch]$AllUsers,
    [switch]$IncludeProtected,
    [switch]$IncludeFramework,
    [switch]$IncludeResource,
    [switch]$IncludeBundle
  )

  process {
    foreach ($_pattern in $Pattern) {
      $_inventory = @(Get-UPFAppxPackage -Name $_pattern -Installed:$Installed -Provisioned:$Provisioned -AllUsers:$AllUsers -IncludeFramework:$IncludeFramework -IncludeResource:$IncludeResource -IncludeBundle:$IncludeBundle)
      if ($_inventory.Count -eq 0) {
        [PSCustomObject]@{
          Pattern = $_pattern
          Matched = $false
          Protected = $false
          ProtectedReason = $null
          Package = $null
        }
        continue
      }

      foreach ($_package in $_inventory) {
        $_safety = Test-UPFAppxPackageRemovalSafety -InputObject $_package
        if ($_safety.Protected -and -not $IncludeProtected) {
          [PSCustomObject]@{
            Pattern = $_pattern
            Matched = $true
            Protected = $true
            ProtectedReason = $_safety.Reason
            Package = $_package
          }
          continue
        }

        [PSCustomObject]@{
          Pattern = $_pattern
          Matched = $true
          Protected = $_safety.Protected
          ProtectedReason = $_safety.Reason
          Package = $_package
        }
      }
    }
  }
}

function Test-UPFAppxPackageRemovalSafety {
  <#
    .SYNOPSIS
      Classifies whether a UPF AppX/MSIX package should be protected from removal by default.
    .DESCRIPTION
      Checks a normalized package inventory object for non-removable,
      framework, resource, and protected-name patterns. Removal helpers use this
      classification to skip Store, App Installer, framework runtimes, and
      shell/security-adjacent packages unless the caller deliberately opts into
      protected removal.
    .PARAMETER InputObject
      Package object returned by Get-UPFAppxPackage or embedded in a
      Find-UPFAppxPackage result.
    .EXAMPLE
      PS> Get-UPFAppxPackage -Name 'Microsoft.WindowsStore' | Test-UPFAppxPackageRemovalSafety
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [psobject]$InputObject
  )

  process {
    $_nameCandidates = @(
      $InputObject.Name,
      $InputObject.DisplayName,
      $InputObject.PackageName,
      $InputObject.PackageFullName,
      $InputObject.PackageFamilyName
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $_reason = $null
    if ($InputObject.NonRemovable) { $_reason = 'NonRemovable' }
    elseif ($InputObject.IsFramework) { $_reason = 'Framework' }
    elseif ($InputObject.IsResourcePackage) { $_reason = 'ResourcePackage' }
    else {
      foreach ($_pattern in $script:UPFProtectedPackagePatterns) {
        if (@($_nameCandidates | Where-Object { $_ -like $_pattern }).Count -gt 0) {
          $_reason = "ProtectedPattern:$_pattern"
          break
        }
      }
    }

    [PSCustomObject]@{
      Target = $InputObject.PackageName
      Protected = -not [string]::IsNullOrWhiteSpace($_reason)
      Reason = $_reason
    }
  }
}

function Install-UPFAppxPackage {
  <#
    .SYNOPSIS
      Installs or provisions a local UPF AppX/MSIX package file.
    .DESCRIPTION
      Installs a local .appx, .msix, .appxbundle, or .msixbundle file for the
      current user, or provisions it into the online Windows image with
      -Provisioned. This function intentionally uses a UPF-prefixed noun to
      avoid shadowing platform cmdlets such as Add-AppxPackage.
    .PARAMETER Path
      Local UPF AppX/MSIX package file path.
    .PARAMETER DependencyPath
      Optional dependency package paths.
    .PARAMETER LicensePath
      Optional license path used with -Provisioned.
    .PARAMETER Provisioned
      Add the package to the online Windows image for future users.
    .PARAMETER SkipLicense
      Skip license processing when provisioning.
    .PARAMETER ForceUpdateFromAnyVersion
      Permit updating from any package version when installing for the current user.
    .PARAMETER DryRun
      Preview the install/provision operation without applying it.
    .EXAMPLE
      PS> Install-UPFAppxPackage -Path 'C:\Packages\Example.msix'
    .EXAMPLE
      PS> Install-UPFAppxPackage -Path 'C:\Packages\Example.msixbundle' -Provisioned -SkipLicense
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$Path,

    [string[]]$DependencyPath,
    [string]$LicensePath,
    [switch]$Provisioned,
    [switch]$SkipLicense,
    [switch]$ForceUpdateFromAnyVersion,
    [switch]$DryRun
  )

  if ($DryRun) { $WhatIfPreference = $true }

  $_path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
  $_action = if ($Provisioned) { 'Provision' } else { 'Install' }

  if (-not $PSCmdlet.ShouldProcess($_path, "$_action UPF AppX/MSIX package")) {
    return New-PackageLifecycleResult -Target $_path -Source 'UPFAppxPackage' -Action $_action -Status 'Skipped' -SkippedReason 'WhatIf'
  }

  try {
    if ($Provisioned) {
      $_params = @{
        Online = $true
        PackagePath = $_path
        ErrorAction = 'Stop'
      }
      if ($DependencyPath) { $_params.DependencyPackagePath = $DependencyPath }
      if ($LicensePath) { $_params.LicensePath = $LicensePath }
      if ($SkipLicense) { $_params.SkipLicense = $true }
      $null = Add-AppxProvisionedPackage @_params
    }
    else {
      $_params = @{
        Path = $_path
        ErrorAction = 'Stop'
      }
      if ($DependencyPath) { $_params.DependencyPath = $DependencyPath }
      if ($ForceUpdateFromAnyVersion) { $_params.ForceUpdateFromAnyVersion = $true }
      $null = Add-AppxPackage @_params
    }

    New-PackageLifecycleResult -Target $_path -Source 'UPFAppxPackage' -Action $_action -Status 'Completed'
  }
  catch {
    New-PackageLifecycleResult -Target $_path -Source 'UPFAppxPackage' -Action $_action -Status 'Failed' -ErrorMessage $_.Exception.Message
  }
}

function Update-UPFAppxPackage {
  <#
    .SYNOPSIS
      Updates an installed UPF AppX/MSIX package from a local package file.
    .DESCRIPTION
      Uses Install-UPFAppxPackage with ForceUpdateFromAnyVersion to install a
      local AppX/MSIX package over an existing package registration. This helper
      is for local package files; Store-delivered update orchestration is left
      to Windows Store or WinGet provider workflows.
    .PARAMETER Path
      Local UPF AppX/MSIX package file path.
    .PARAMETER DependencyPath
      Optional dependency package paths.
    .PARAMETER DryRun
      Preview the update without applying it.
    .EXAMPLE
      PS> Update-UPFAppxPackage -Path 'C:\Packages\Example-2.0.msix'
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$Path,

    [string[]]$DependencyPath,
    [switch]$DryRun
  )

  Install-UPFAppxPackage -Path $Path -DependencyPath $DependencyPath -ForceUpdateFromAnyVersion -DryRun:$DryRun -WhatIf:$WhatIfPreference
}

function Repair-UPFAppxPackage {
  <#
    .SYNOPSIS
      Repairs installed UPF AppX/MSIX packages by re-registering their manifests.
    .DESCRIPTION
      Resolves installed packages by name and runs Add-AppxPackage -Register
      against each package's AppxManifest.xml. This is useful for restoring
      package registrations without deleting app data.
    .PARAMETER Name
      Package name or wildcard pattern to repair.
    .PARAMETER AllUsers
      Resolve installed packages for all users.
    .PARAMETER DryRun
      Preview manifest registrations without applying them.
    .EXAMPLE
      PS> Repair-UPFAppxPackage -Name 'Microsoft.WindowsStore'
    .EXAMPLE
      PS> Repair-UPFAppxPackage -Name 'Microsoft.DesktopAppInstaller*' -DryRun
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [switch]$AllUsers,
    [switch]$DryRun
  )

  if ($DryRun) { $WhatIfPreference = $true }
  $_packages = @(Get-UPFAppxPackage -Name $Name -Installed -AllUsers:$AllUsers -IncludeFramework -IncludeResource -IncludeBundle |
      Where-Object { -not [string]::IsNullOrWhiteSpace($_.InstallLocation) })

  if ($_packages.Count -eq 0) {
    return New-PackageLifecycleResult -Target $Name -Source 'UPFAppxPackage' -Action 'Repair' -Status 'Skipped' -SkippedReason 'NoMatch'
  }

  foreach ($_package in $_packages) {
    $_manifest = Join-Path -Path $_package.InstallLocation -ChildPath 'AppxManifest.xml'
    if (-not (Test-Path -LiteralPath $_manifest)) {
      New-PackageLifecycleResult -Target $_package.PackageFullName -Source 'UPFAppxPackage' -Action 'Repair' -Status 'Skipped' -SkippedReason 'ManifestNotFound'
      continue
    }

    if (-not $PSCmdlet.ShouldProcess($_package.PackageFullName, 'Re-register UPF AppX/MSIX package')) {
      New-PackageLifecycleResult -Target $_package.PackageFullName -Source 'UPFAppxPackage' -Action 'Repair' -Status 'Skipped' -SkippedReason 'WhatIf'
      continue
    }

    try {
      $null = Add-AppxPackage -DisableDevelopmentMode -Register $_manifest -ErrorAction Stop
      New-PackageLifecycleResult -Target $_package.PackageFullName -Source 'UPFAppxPackage' -Action 'Repair' -Status 'Completed'
    }
    catch {
      New-PackageLifecycleResult -Target $_package.PackageFullName -Source 'UPFAppxPackage' -Action 'Repair' -Status 'Failed' -ErrorMessage $_.Exception.Message
    }
  }
}

function Reset-UPFAppxPackage {
  <#
    .SYNOPSIS
      Resets app data for installed UPF AppX/MSIX packages when the platform supports it.
    .DESCRIPTION
      Resolves installed packages by name and calls the platform
      Reset-AppxPackage cmdlet when it is available. The helper does not delete
      package folders manually; unsupported platforms return a skipped result.
    .PARAMETER Name
      Package name or wildcard pattern to reset.
    .PARAMETER DryRun
      Preview reset operations without applying them.
    .EXAMPLE
      PS> Reset-UPFAppxPackage -Name 'Microsoft.GetHelp'
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [switch]$DryRun
  )

  if ($DryRun) { $WhatIfPreference = $true }
  if (-not (Get-Command -Name Reset-AppxPackage -ErrorAction SilentlyContinue)) {
    return New-PackageLifecycleResult -Target $Name -Source 'UPFAppxPackage' -Action 'Reset' -Status 'Skipped' -SkippedReason 'ResetAppxPackageUnavailable'
  }

  $_packages = @(Get-UPFAppxPackage -Name $Name -Installed)
  if ($_packages.Count -eq 0) {
    return New-PackageLifecycleResult -Target $Name -Source 'UPFAppxPackage' -Action 'Reset' -Status 'Skipped' -SkippedReason 'NoMatch'
  }

  foreach ($_package in $_packages) {
    if (-not $PSCmdlet.ShouldProcess($_package.PackageFullName, 'Reset UPF AppX/MSIX package data')) {
      New-PackageLifecycleResult -Target $_package.PackageFullName -Source 'UPFAppxPackage' -Action 'Reset' -Status 'Skipped' -SkippedReason 'WhatIf'
      continue
    }

    try {
      $null = Reset-AppxPackage -Package $_package.PackageFullName -ErrorAction Stop
      New-PackageLifecycleResult -Target $_package.PackageFullName -Source 'UPFAppxPackage' -Action 'Reset' -Status 'Completed'
    }
    catch {
      New-PackageLifecycleResult -Target $_package.PackageFullName -Source 'UPFAppxPackage' -Action 'Reset' -Status 'Failed' -ErrorMessage $_.Exception.Message
    }
  }
}

function Uninstall-UPFAppxPackage {
  <#
    .SYNOPSIS
      Uninstalls a single normalized UPF AppX/MSIX package inventory record.
    .DESCRIPTION
      Removes one installed or provisioned package object produced by
      Get-UPFAppxPackage or Find-UPFAppxPackage. Protected packages are skipped
      unless both -IncludeProtected and -Force are supplied.
    .PARAMETER InputObject
      Normalized package record to uninstall.
    .PARAMETER AllUsers
      Remove installed package registrations for all users.
    .PARAMETER IncludeProtected
      Allow protected package records into removal evaluation.
    .PARAMETER Force
      Required with -IncludeProtected to remove protected package records.
    .PARAMETER DryRun
      Preview the uninstall operation without applying it.
    .EXAMPLE
      PS> Find-UPFAppxPackage -Pattern 'Microsoft.ZuneMusic' -Installed | Where-Object Matched | ForEach-Object Package | Uninstall-UPFAppxPackage -DryRun
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [psobject]$InputObject,

    [switch]$AllUsers,
    [switch]$IncludeProtected,
    [switch]$Force,
    [switch]$DryRun
  )

  process {
    if ($DryRun) { $WhatIfPreference = $true }

    $_safety = Test-UPFAppxPackageRemovalSafety -InputObject $InputObject
    if ($_safety.Protected -and (-not $IncludeProtected -or -not $Force)) {
      return New-PackageLifecycleResult -Target $InputObject.PackageName -Source $InputObject.Source -Action 'Uninstall' -Status 'Skipped' -SkippedReason $_safety.Reason
    }

    if ($InputObject.Source -eq 'Provisioned') {
      $_target = $InputObject.PackageName
      if (-not $PSCmdlet.ShouldProcess($_target, 'Remove provisioned UPF AppX/MSIX package')) {
        return New-PackageLifecycleResult -Target $_target -Source 'Provisioned' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'WhatIf'
      }
      try {
        $null = Remove-AppxProvisionedPackage -Online -PackageName $_target -ErrorAction Stop
        return New-PackageLifecycleResult -Target $_target -Source 'Provisioned' -Action 'Uninstall' -Status 'Removed'
      }
      catch {
        return New-PackageLifecycleResult -Target $_target -Source 'Provisioned' -Action 'Uninstall' -Status 'Failed' -ErrorMessage $_.Exception.Message
      }
    }

    $_target = $InputObject.PackageFullName
    if (-not $PSCmdlet.ShouldProcess($_target, 'Remove installed UPF AppX/MSIX package')) {
      return New-PackageLifecycleResult -Target $_target -Source 'Installed' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'WhatIf'
    }

    try {
      if ($AllUsers) {
        $null = Remove-AppxPackage -Package $_target -AllUsers -ErrorAction Stop
      }
      else {
        $null = Remove-AppxPackage -Package $_target -ErrorAction Stop
      }
      New-PackageLifecycleResult -Target $_target -Source 'Installed' -Action 'Uninstall' -Status 'Removed'
    }
    catch {
      New-PackageLifecycleResult -Target $_target -Source 'Installed' -Action 'Uninstall' -Status 'Failed' -ErrorMessage $_.Exception.Message
    }
  }
}

function Uninstall-UPFAppxPackageSet {
  <#
    .SYNOPSIS
      Resolves and uninstalls a set of UPF AppX/MSIX package wildcard patterns.
    .DESCRIPTION
      Sequentially resolves package patterns through Find-UPFAppxPackage and
      sends each concrete package to Uninstall-UPFAppxPackage. This set helper
      intentionally reuses the singular uninstall function so safety checks,
      WhatIf behavior, and result objects stay consistent.
    .PARAMETER Pattern
      One or more package wildcard patterns to resolve and remove.
    .PARAMETER Installed
      Remove installed packages. Enabled by default.
    .PARAMETER Provisioned
      Also remove provisioned packages from the online Windows image.
    .PARAMETER AllUsers
      Remove installed packages for all user profiles.
    .PARAMETER IncludeProtected
      Include protected package matches in removal evaluation.
    .PARAMETER Force
      Required with -IncludeProtected to remove protected matches.
    .PARAMETER PassThru
      Return structured lifecycle results instead of only writing a summary.
    .PARAMETER DryRun
      Preview all removals without changing the system.
    .EXAMPLE
      PS> Uninstall-UPFAppxPackageSet -Pattern 'Microsoft.Zune*' -DryRun
    .EXAMPLE
      PS> Uninstall-UPFAppxPackageSet -Pattern 'Microsoft.GetHelp','Microsoft.BingWeather*' -Provisioned -PassThru
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string[]]$Pattern,

    [switch]$Installed,
    [switch]$Provisioned,
    [switch]$AllUsers,
    [switch]$IncludeProtected,
    [switch]$Force,
    [switch]$PassThru,
    [switch]$DryRun
  )

  if ($DryRun) {
    $WhatIfPreference = $true
    Write-Log -Message "DRY RUN - no UPF AppX/MSIX packages will be removed`n" -Color Yellow
  }

  $_includeInstalled = $Installed -or -not $PSBoundParameters.ContainsKey('Installed')
  $_results = New-Object System.Collections.ArrayList
  $_matches = @(Find-UPFAppxPackage -Pattern $Pattern -Installed:$_includeInstalled -Provisioned:$Provisioned -AllUsers:$AllUsers -IncludeProtected:$IncludeProtected)

  foreach ($_match in $_matches) {
    if (-not $_match.Matched) {
      [void]$_results.Add((New-PackageLifecycleResult -Target $_match.Pattern -Source 'UPFAppxPackage' -Action 'Resolve' -Status 'Skipped' -SkippedReason 'NoMatch'))
      continue
    }

    if ($_match.Protected -and (-not $IncludeProtected -or -not $Force)) {
      [void]$_results.Add((New-PackageLifecycleResult -Target $_match.Package.PackageName -Source $_match.Package.Source -Action 'Uninstall' -Status 'Skipped' -SkippedReason $_match.ProtectedReason))
      continue
    }

    $_removeResult = Uninstall-UPFAppxPackage -InputObject $_match.Package -AllUsers:$AllUsers -IncludeProtected:$IncludeProtected -Force:$Force -DryRun:$DryRun -WhatIf:$WhatIfPreference
    [void]$_results.Add($_removeResult)
  }

  if ($PassThru) { return $_results }

  $_removed = @($_results | Where-Object { $_.Status -eq 'Removed' }).Count
  $_skipped = @($_results | Where-Object { $_.Status -eq 'Skipped' }).Count
  $_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
  Write-Log -Message "UPF AppX/MSIX package removal complete. Removed: $_removed | Skipped: $_skipped | Failed: $_failed" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
}

function Install-UPFAppxPackageSet {
  <#
    .SYNOPSIS
      Sequentially installs a set of local UPF AppX/MSIX package files.
    .DESCRIPTION
      Calls Install-UPFAppxPackage once for each supplied path and optionally
      returns the collected lifecycle results. This avoids duplicating install
      logic while giving scripts a convenient set operation for manifests or
      prepared package folders.
    .PARAMETER Path
      One or more local package file paths.
    .PARAMETER DependencyPath
      Optional dependency package paths passed to each install operation.
    .PARAMETER Provisioned
      Provision each package into the online Windows image.
    .PARAMETER SkipLicense
      Skip license processing when provisioning.
    .PARAMETER ForceUpdateFromAnyVersion
      Permit package updates from any version during install.
    .PARAMETER PassThru
      Return lifecycle result objects.
    .PARAMETER DryRun
      Preview installs without applying them.
    .EXAMPLE
      PS> Install-UPFAppxPackageSet -Path 'C:\Packages\App1.msix','C:\Packages\App2.msix' -DryRun
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string[]]$Path,

    [string[]]$DependencyPath,
    [switch]$Provisioned,
    [switch]$SkipLicense,
    [switch]$ForceUpdateFromAnyVersion,
    [switch]$PassThru,
    [switch]$DryRun
  )

  $_results = foreach ($_path in $Path) {
    Install-UPFAppxPackage -Path $_path -DependencyPath $DependencyPath -Provisioned:$Provisioned -SkipLicense:$SkipLicense -ForceUpdateFromAnyVersion:$ForceUpdateFromAnyVersion -DryRun:$DryRun -WhatIf:$WhatIfPreference
  }

  if ($PassThru) { return $_results }
}

function Install-Win32ProgramFromWinGet {
  <#
    .SYNOPSIS
      Installs a Win32 program through the Microsoft.WinGet.Client provider.
    .DESCRIPTION
      Uses Microsoft.WinGet.Client when available to install an application by
      name or exact package ID. The winkit function name is intentionally
      Win32ProgramFromWinGet so it does not shadow provider cmdlets exported by
      Microsoft.WinGet.Client.
    .PARAMETER Name
      Package name to install through WinGet.
    .PARAMETER Id
      Exact WinGet package identifier to install.
    .PARAMETER Version
      Optional package version.
    .PARAMETER Source
      Optional WinGet source.
    .PARAMETER DryRun
      Preview the install without invoking the provider command.
    .EXAMPLE
      PS> Install-Win32ProgramFromWinGet -Id 'Microsoft.VisualStudioCode'
    .EXAMPLE
      PS> Install-Win32ProgramFromWinGet -Name 'Visual Studio Code' -DryRun
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
    [string]$Name,

    [Parameter(Mandatory = $true, ParameterSetName = 'Id')]
    [string]$Id,

    [string]$Version,
    [string]$Source,
    [switch]$DryRun
  )

  if ($DryRun) { $WhatIfPreference = $true }
  if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client -ErrorAction SilentlyContinue)) {
    return New-PackageLifecycleResult -Target $(if ($Id) { $Id } else { $Name }) -Source 'WinGet' -Action 'Install' -Status 'Skipped' -SkippedReason 'Microsoft.WinGet.ClientUnavailable'
  }

  Import-Module Microsoft.WinGet.Client -Force -ErrorAction Stop
  $_cmd = Get-Command -Name Install-WinGetPackage -Module Microsoft.WinGet.Client -ErrorAction Stop
  $_target = if ($PSCmdlet.ParameterSetName -eq 'Id') { $Id } else { $Name }

  if (-not $PSCmdlet.ShouldProcess($_target, 'Install WinGet package')) {
    return New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Install' -Status 'Skipped' -SkippedReason 'WhatIf'
  }

  try {
    $_params = @{ ErrorAction = 'Stop' }
    if ($PSCmdlet.ParameterSetName -eq 'Id') { $_params.Id = $Id } else { $_params.Name = $Name }
    if ($Version) { $_params.Version = $Version }
    if ($Source) { $_params.Source = $Source }
    $null = & $_cmd @_params
    New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Install' -Status 'Completed'
  }
  catch {
    New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Install' -Status 'Failed' -ErrorMessage $_.Exception.Message
  }
}

function Update-Win32ProgramFromWinGet {
  <#
    .SYNOPSIS
      Updates one or all Win32 programs through the Microsoft.WinGet.Client provider.
    .DESCRIPTION
      Uses Microsoft.WinGet.Client when available to update a specific package
      by name or ID, or every eligible package with -All. The function name
      avoids shadowing Microsoft.WinGet.Client's own Update-WinGetPackage cmdlet.
    .PARAMETER Id
      Exact WinGet package identifier to update.
    .PARAMETER Name
      Package name to update.
    .PARAMETER All
      Update all eligible packages.
    .PARAMETER IncludeUnknown
      Include packages with unknown installed versions when the provider supports it.
    .PARAMETER Source
      Optional WinGet source.
    .PARAMETER DryRun
      Preview the update without invoking the provider command.
    .EXAMPLE
      PS> Update-Win32ProgramFromWinGet -Id 'Microsoft.VisualStudioCode'
    .EXAMPLE
      PS> Update-Win32ProgramFromWinGet -All -DryRun
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(ParameterSetName = 'Id')]
    [string]$Id,

    [Parameter(ParameterSetName = 'Name')]
    [string]$Name,

    [Parameter(ParameterSetName = 'All')]
    [switch]$All,

    [switch]$IncludeUnknown,
    [string]$Source,
    [switch]$DryRun
  )

  if ($DryRun) { $WhatIfPreference = $true }
  if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client -ErrorAction SilentlyContinue)) {
    return New-PackageLifecycleResult -Target 'WinGet' -Source 'WinGet' -Action 'Update' -Status 'Skipped' -SkippedReason 'Microsoft.WinGet.ClientUnavailable'
  }

  Import-Module Microsoft.WinGet.Client -Force -ErrorAction Stop
  $_cmd = Get-Command -Name Update-WinGetPackage -Module Microsoft.WinGet.Client -ErrorAction SilentlyContinue
  if (-not $_cmd) {
    return New-PackageLifecycleResult -Target 'WinGet' -Source 'WinGet' -Action 'Update' -Status 'Skipped' -SkippedReason 'UpdateWinGetPackageUnavailable'
  }

  $_target = if ($Id) { $Id } elseif ($Name) { $Name } elseif ($All) { 'All' } else { 'All' }
  if (-not $PSCmdlet.ShouldProcess($_target, 'Update WinGet package')) {
    return New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Update' -Status 'Skipped' -SkippedReason 'WhatIf'
  }

  try {
    $_params = @{ ErrorAction = 'Stop' }
    if ($Id) { $_params.Id = $Id }
    if ($Name) { $_params.Name = $Name }
    if ($All) { $_params.All = $true }
    if ($IncludeUnknown) { $_params.IncludeUnknown = $true }
    if ($Source) { $_params.Source = $Source }
    $null = & $_cmd @_params
    New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Update' -Status 'Completed'
  }
  catch {
    New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Update' -Status 'Failed' -ErrorMessage $_.Exception.Message
  }
}

function Uninstall-Win32ProgramFromWinGet {
  <#
    .SYNOPSIS
      Uninstalls a Win32 program through the Microsoft.WinGet.Client provider.
    .DESCRIPTION
      Uses Microsoft.WinGet.Client when available to uninstall an application by
      name or exact package ID. The wrapper returns winkit lifecycle result
      objects while avoiding command-name collisions with the provider module.
    .PARAMETER Name
      Package name to uninstall through WinGet.
    .PARAMETER Id
      Exact WinGet package identifier to uninstall.
    .PARAMETER DryRun
      Preview the uninstall without invoking the provider command.
    .EXAMPLE
      PS> Uninstall-Win32ProgramFromWinGet -Id 'Microsoft.OneDrive' -DryRun
    .EXAMPLE
      PS> Uninstall-Win32ProgramFromWinGet -Name 'Microsoft OneDrive'
    .LINK
      https://github.com/adnoctem/winkit/lib/packages.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Name')]
    [string]$Name,

    [Parameter(Mandatory = $true, ParameterSetName = 'Id')]
    [string]$Id,

    [switch]$DryRun
  )

  if ($DryRun) { $WhatIfPreference = $true }
  if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client -ErrorAction SilentlyContinue)) {
    return New-PackageLifecycleResult -Target $(if ($Id) { $Id } else { $Name }) -Source 'WinGet' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'Microsoft.WinGet.ClientUnavailable'
  }

  Import-Module Microsoft.WinGet.Client -Force -ErrorAction Stop
  $_cmd = Get-Command -Name Uninstall-WinGetPackage -Module Microsoft.WinGet.Client -ErrorAction Stop
  $_target = if ($PSCmdlet.ParameterSetName -eq 'Id') { $Id } else { $Name }

  if (-not $PSCmdlet.ShouldProcess($_target, 'Uninstall WinGet package')) {
    return New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'WhatIf'
  }

  try {
    $_params = @{ ErrorAction = 'Stop' }
    if ($PSCmdlet.ParameterSetName -eq 'Id') { $_params.Id = $Id } else { $_params.Name = $Name }
    $null = & $_cmd @_params
    New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Uninstall' -Status 'Completed'
  }
  catch {
    New-PackageLifecycleResult -Target $_target -Source 'WinGet' -Action 'Uninstall' -Status 'Failed' -ErrorMessage $_.Exception.Message
  }
}
