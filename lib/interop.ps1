#Requires -Version 5.0

function Remove-ComObject {
  <#
    .SYNOPSIS
      Releases one or more COM objects.
    .DESCRIPTION
      Wraps Marshal.ReleaseComObject with null checks and best-effort error
      handling so scripts can safely clean up Outlook and Office interop
      objects from finally blocks.
    .EXAMPLE
      PS> Remove-ComObject $items $folder
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'COM cleanup must be best-effort during script teardown.')]
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Releases local COM references only; it does not change external system state.')]
  [CmdletBinding()]
  param (
    [Parameter(ValueFromRemainingArguments = $true)]
    [AllowNull()]
    [object[]]
    $InputObject
  )

  foreach ($_object in $InputObject) {
    if ($null -eq $_object) { continue }

    try {
      if ([System.Runtime.InteropServices.Marshal]::IsComObject($_object)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($_object)
      }
    }
    catch { }
  }
}

function Invoke-ComGarbageCollection {
  <#
    .SYNOPSIS
      Runs final COM cleanup garbage collection passes.
    .DESCRIPTION
      Forces garbage collection and waits for pending finalizers. This is useful
      after releasing Office COM references so Outlook can close PST files and
      exit cleanly when requested.
    .EXAMPLE
      PS> Invoke-ComGarbageCollection
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param ()

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Get-OutlookInstallation {
  <#
    .SYNOPSIS
      Finds local Outlook installation directories.
    .DESCRIPTION
      Discovers common Microsoft Office and Microsoft 365 installation roots
      that may contain Outlook.exe or Outlook data-file repair tools such as
      ScanPST.exe and ScanOST.exe. The function checks App Paths registry
      entries first, then common Office directory layouts under Program Files.
    .EXAMPLE
      PS> Get-OutlookInstallation
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param ()

  $_candidateDirectories = New-Object System.Collections.Generic.List[string]
  $_registryPaths = @(
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
    'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
  )

  foreach ($_registryPath in $_registryPaths) {
    if (-not (Test-Path -LiteralPath $_registryPath)) { continue }

    $_property = Get-ItemProperty -LiteralPath $_registryPath -ErrorAction SilentlyContinue
    if (-not $_property) { continue }

    $_outlookPath = $_property.'(default)'
    if ([string]::IsNullOrWhiteSpace($_outlookPath) -and ($_property.PSObject.Properties.Name -contains 'Path')) {
      $_outlookPath = Join-Path -Path $_property.Path -ChildPath 'OUTLOOK.EXE'
    }

    if (-not [string]::IsNullOrWhiteSpace($_outlookPath)) {
      $_directory = Split-Path -Path ([Environment]::ExpandEnvironmentVariables($_outlookPath)) -Parent
      if (-not [string]::IsNullOrWhiteSpace($_directory)) {
        $_candidateDirectories.Add($_directory)
      }
    }
  }

  $_programRoots = @(
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)}
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

  $_officeVersions = @('Office16', 'Office15', 'Office14', 'Office12', 'Office11')
  foreach ($_programRoot in $_programRoots) {
    $_officeRoot = Join-Path -Path $_programRoot -ChildPath 'Microsoft Office'

    foreach ($_version in $_officeVersions) {
      $_candidateDirectories.Add((Join-Path -Path $_officeRoot -ChildPath $_version))
      $_candidateDirectories.Add((Join-Path -Path $_officeRoot -ChildPath "root\$_version"))
    }
  }

  $_seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($_directory in $_candidateDirectories) {
    if ([string]::IsNullOrWhiteSpace($_directory)) { continue }
    if (-not (Test-Path -LiteralPath $_directory -PathType Container)) { continue }

    $_resolvedDirectory = (Resolve-Path -LiteralPath $_directory).ProviderPath
    if (-not $_seen.Add($_resolvedDirectory)) { continue }

    $_outlookPath = Join-Path -Path $_resolvedDirectory -ChildPath 'OUTLOOK.EXE'
    $_scanPstPath = Join-Path -Path $_resolvedDirectory -ChildPath 'SCANPST.EXE'
    $_scanOstPath = Join-Path -Path $_resolvedDirectory -ChildPath 'SCANOST.EXE'

    if (
      -not (Test-Path -LiteralPath $_outlookPath -PathType Leaf) -and
      -not (Test-Path -LiteralPath $_scanPstPath -PathType Leaf) -and
      -not (Test-Path -LiteralPath $_scanOstPath -PathType Leaf)
    ) {
      continue
    }

    [PSCustomObject]@{
      Path = $_resolvedDirectory
      OutlookPath = if (Test-Path -LiteralPath $_outlookPath -PathType Leaf) { $_outlookPath } else { $null }
      ScanPstPath = if (Test-Path -LiteralPath $_scanPstPath -PathType Leaf) { $_scanPstPath } else { $null }
      ScanOstPath = if (Test-Path -LiteralPath $_scanOstPath -PathType Leaf) { $_scanOstPath } else { $null }
    }
  }
}

function Find-OutlookRepairTool {
  <#
    .SYNOPSIS
      Finds an Outlook data-file repair tool.
    .DESCRIPTION
      Resolves ScanPST.exe or ScanOST.exe from discovered Outlook installation
      directories. ScanPST is present in modern Outlook installs; ScanOST exists
      only in older Outlook versions.
    .PARAMETER Name
      Repair tool executable to find.
    .EXAMPLE
      PS> Find-OutlookRepairTool -Name ScanPST
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param (
    [ValidateSet('ScanPST', 'ScanOST')]
    [string]
    $Name = 'ScanPST'
  )

  $_propertyName = if ($Name -eq 'ScanOST') { 'ScanOstPath' } else { 'ScanPstPath' }
  $_fileName = if ($Name -eq 'ScanOST') { 'SCANOST.EXE' } else { 'SCANPST.EXE' }

  foreach ($_installation in Get-OutlookInstallation) {
    $_path = $_installation.$_propertyName
    if ([string]::IsNullOrWhiteSpace($_path)) { continue }
    if (-not (Test-Path -LiteralPath $_path -PathType Leaf)) { continue }

    [PSCustomObject]@{
      Name = $Name
      Path = $_path
      InstallationPath = $_installation.Path
    }
  }

  $_command = Get-Command -Name $_fileName -ErrorAction SilentlyContinue
  if ($_command -and (Test-Path -LiteralPath $_command.Source -PathType Leaf)) {
    [PSCustomObject]@{
      Name = $Name
      Path = $_command.Source
      InstallationPath = Split-Path -Path $_command.Source -Parent
    }
  }
}

function Connect-Outlook {
  <#
    .SYNOPSIS
      Connects to an Outlook COM application and MAPI namespace.
    .DESCRIPTION
      Reuses a running Outlook instance when available, otherwise starts one,
      then logs on to the MAPI namespace without prompting.
    .EXAMPLE
      PS> $context = Connect-Outlook
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'Falling back to a new Outlook COM instance is intentional.')]
  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param ()

  try {
    $_application = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
  }
  catch {
    $_application = New-Object -ComObject Outlook.Application
  }

  $_namespace = $_application.GetNamespace('MAPI')
  $_namespace.Logon($null, $null, $false, $false)

  [PSCustomObject]@{
    App = $_application
    Namespace = $_namespace
  }
}

function Get-OutlookStoreRoot {
  <#
    .SYNOPSIS
      Gets the root folder for an Outlook store.
    .DESCRIPTION
      Resolves a named Outlook store by DisplayName, or the default delivery
      store when no name is supplied. Store COM objects are released as they are
      inspected; the returned root folder is owned by the caller.
    .PARAMETER Namespace
      Outlook MAPI namespace returned by Connect-Outlook.
    .PARAMETER Name
      Optional store display name.
    .EXAMPLE
      PS> Get-OutlookStoreRoot -Namespace $context.Namespace -Name 'user@example.com'
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Namespace,

    [string]
    $Name
  )

  $_stores = $Namespace.Stores
  try {
    Write-Verbose 'Available Outlook stores:'
    for ($_index = 1; $_index -le $_stores.Count; $_index++) {
      $_store = $_stores.Item($_index)
      try {
        Write-Verbose ("  - {0}" -f $_store.DisplayName)
        if ([string]::IsNullOrWhiteSpace($Name) -and $_store.IsDefault) {
          return $_store.GetRootFolder()
        }

        if ($_store.DisplayName -eq $Name) {
          return $_store.GetRootFolder()
        }
      }
      finally {
        Remove-ComObject $_store
      }
    }
  }
  finally {
    Remove-ComObject $_stores
  }

  throw "Outlook store '$Name' not found. Run with -Verbose to list available stores."
}

function Add-OutlookStoreRoot {
  <#
    .SYNOPSIS
      Adds a Unicode PST store and returns its root folder.
    .DESCRIPTION
      Calls Outlook Namespace.AddStoreEx with OlStoreType.olStoreUnicode and
      locates the newly attached store by FilePath. The returned root folder is
      owned by the caller.
    .PARAMETER Namespace
      Outlook MAPI namespace returned by Connect-Outlook.
    .PARAMETER Path
      Full path to the PST file to attach or create.
    .EXAMPLE
      PS> Add-OutlookStoreRoot -Namespace $context.Namespace -Path 'D:\Archive\mail.pst'
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Namespace,

    [Parameter(Mandatory = $true)]
    [string]
    $Path
  )

  $_olStoreUnicode = 2
  $_resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
  $Namespace.AddStoreEx($_resolvedPath, $_olStoreUnicode)

  $_stores = $Namespace.Stores
  try {
    for ($_index = 1; $_index -le $_stores.Count; $_index++) {
      $_store = $_stores.Item($_index)
      try {
        if ($_store.FilePath -eq $_resolvedPath) {
          return $_store.GetRootFolder()
        }
      }
      finally {
        Remove-ComObject $_store
      }
    }
  }
  finally {
    Remove-ComObject $_stores
  }

  throw "Outlook store was added but could not be located by path: $_resolvedPath"
}

function Get-OutlookSubFolder {
  <#
    .SYNOPSIS
      Gets or creates an Outlook child folder.
    .DESCRIPTION
      Searches a parent folder's Folders collection by display name. When
      -Create is set, the folder is created if missing. The returned folder is
      owned by the caller.
    .PARAMETER ParentFolder
      Outlook parent folder.
    .PARAMETER Name
      Child folder display name.
    .PARAMETER Create
      Create the folder when it does not exist.
    .EXAMPLE
      PS> Get-OutlookSubFolder -ParentFolder $root -Name '_Review' -Create
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $ParentFolder,

    [Parameter(Mandatory = $true)]
    [string]
    $Name,

    [switch]
    $Create
  )

  $_folders = $ParentFolder.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_folder = $_folders.Item($_index)
      if ($_folder.Name -eq $Name) {
        return $_folder
      }

      Remove-ComObject $_folder
    }

    if ($Create) {
      return $_folders.Add($Name)
    }
  }
  finally {
    Remove-ComObject $_folders
  }

  return $null
}
