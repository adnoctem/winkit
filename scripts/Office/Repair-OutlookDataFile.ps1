#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.8.2' }

<#
.SYNOPSIS
  Launches ScanPST to inspect or repair an Outlook data file.
.DESCRIPTION
  Discovers ScanPST only. Supported Office 16 builds receive the selected file
  through -file with -rescan 1, keeping the utility's normal visible UI.
  Older or unrecognized executables launch interactively without arguments;
  select the file or profile in their window. No forced or silent mode is used.

  ToolPath permits an explicit legacy utility. Paths are expanded to their
  long filesystem names. Close Outlook first. When Path is supplied, the
  script also checks exclusive file access before starting the utility.
  A successful process exit does not certify data-file health or that a
  repair occurred. Review the utility's results and log.
.PARAMETER Path
  Existing PST or OST file. Required for supported ScanPST file targeting.
  Targeted mode requires local storage; UNC paths and mapped network drives
  are rejected, including during preview.
  Optional for interactive tools: displayed and checked, but not forwarded.
  Interactive selection remains the user's responsibility.
.PARAMETER ToolPath
  Explicit executable path instead of automatic ScanPST discovery. Legacy
  alternatives launch interactively unless recognized as a supported ScanPST.
.PARAMETER DryRun
  Preview the executable, launch mode, and actual arguments without starting it.
.PARAMETER PassThru
  Return operation results including LaunchMode, ToolPath, and FileVersion.
.EXAMPLE
  PS> .\Repair-OutlookDataFile.ps1 -Path 'C:\Users\User\Documents\Outlook Files\archive.pst'
.EXAMPLE
  PS> .\Repair-OutlookDataFile.ps1 -ToolPath 'C:\Program Files (x86)\Microsoft Office\Office12\SCANOST.EXE'
  Launches an explicitly selected legacy utility for interactive profile selection.
.EXAMPLE
  PS> .\Repair-OutlookDataFile.ps1 -Path D:\Mail\archive.pst -DryRun -PassThru
.LINK
  https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/data-files/scanpst-exe-runs-multiple-passes
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: not supported - the repair utility requires an interactive desktop.
  SYSTEM-account execution: not supported - repair requires an interactive desktop.
  Outlook version: 2007 or later for discovery; targeting depends on executable version.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [ValidateNotNullOrEmpty()]
  [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
  [string]
  $Path,

  [ValidateNotNullOrEmpty()]
  [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
  [string]
  $ToolPath,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - Outlook data-file repair will not be started`n" -Color Yellow
}

$_results = New-Object Collections.ArrayList
$_target = $Path
$_dataFilePath = $null
$_toolPath = $null
$_toolName = $null
$_launchMode = $null
$_fileVersion = $null
$_property = @{}

try {
  Write-Log -Message 'Checking the data file and locating ScanPST...' -Color Cyan
  Write-Progress -Id 40 -Activity 'Outlook data-file repair' -Status 'Validating paths and locating ScanPST' -PercentComplete -1
  if ($PSBoundParameters.ContainsKey('Path')) {
    $_dataFilePath = Resolve-LongPath -LiteralPath $Path
    $_target = $_dataFilePath
    if ([IO.Path]::GetExtension($_dataFilePath) -notin @('.pst', '.ost')) {
      throw 'Expected an existing .pst or .ost file.'
    }
  }

  if ($ToolPath) {
    $_repairTool = Get-OutlookRepairToolInfo -LiteralPath $ToolPath
  }
  else {
    $_repairTool = Find-OutlookRepairTool -Name ScanPST | Select-Object -First 1
  }
  if (-not $_repairTool) {
    throw 'No ScanPST.exe installation was found. Supply ToolPath to select an executable explicitly.'
  }

  $_toolPath = Resolve-LongPath -LiteralPath $_repairTool.Path
  $_toolName = $_repairTool.Name
  $_fileVersion = $_repairTool.FileVersion
  $_launchMode = if ($_repairTool.SupportsFileArgument) { 'Targeted' } else { 'Interactive' }
  if (-not $_dataFilePath) {
    $_target = $_toolPath
  }

  $_property = @{
    Tool          = $_toolName
    ToolPath      = $_toolPath
    FileVersion   = [string]$_fileVersion
    LaunchMode    = $_launchMode
    RequestedPath = $_dataFilePath
  }
  $_processArguments = @{
    FilePath    = $_toolPath
    Wait        = $true
    PassThru    = $true
    ErrorAction = 'Stop'
  }
  $_commandLine = '"' + $_toolPath + '"'

  if ($_launchMode -eq 'Targeted') {
    if (-not $_dataFilePath) {
      throw 'This ScanPST supports file targeting. Supply Path to select the PST or OST to inspect.'
    }

    $_localPath = $_dataFilePath
    if ($_localPath.StartsWith('\\?\', [StringComparison]::OrdinalIgnoreCase)) {
      $_localPath = $_localPath.Substring(4)
    }
    if ($_localPath.StartsWith('\\') -or $_localPath.StartsWith('UNC\', [StringComparison]::OrdinalIgnoreCase)) {
      throw 'Targeted ScanPST requires a local data file. UNC paths and mapped network drives are not supported.'
    }
    if ($_localPath -notmatch '^[A-Za-z]:\\') {
      throw 'Cannot verify local storage for the selected data file.'
    }

    $_driveId = $_localPath.Substring(0, 2)
    $_disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$_driveId'" -ErrorAction Stop
    if (-not $_disk -or $_disk.DriveType -notin @(2, 3, 6)) {
      throw 'Targeted ScanPST requires verified local storage. UNC paths and mapped network drives are not supported.'
    }

    # -rescan activates the documented command-line mode. One pass preserves
    # a bounded run; no -force or -silent switches bypass the normal UI.
    $_processArguments.ArgumentList = '-file "' + $_dataFilePath + '" -rescan 1'
    $_commandLine += ' ' + $_processArguments.ArgumentList
  }
  else {
    Write-Log -Message 'Interactive mode: select the file or profile in the repair utility. No file argument will be passed.' -Color Yellow
    if ($_dataFilePath) {
      Write-Log -Message "Requested file (select manually): $_dataFilePath" -Color Gray
    }
  }

  if ($WhatIfPreference) {
    Write-Log -Message "[DRY RUN] Would run ($_launchMode): $_commandLine" -Color Yellow
    Add-OperationResult -Results $_results -Target $_target -Source 'OutlookRepair' -Action 'LaunchRepairTool' -Status 'Skipped' -Detail "DryRun: $_commandLine" -Property $_property
  }
  elseif ($PSCmdlet.ShouldProcess($_target, "Launch $_toolName in $_launchMode mode")) {
    Write-Progress -Id 40 -Activity 'Outlook data-file repair' -Status 'Checking Outlook processes and exclusive file access' -PercentComplete -1
    if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
      throw 'Close Outlook before launching the repair tool.'
    }
    if ($_dataFilePath) {
      $_lock = [IO.File]::Open($_dataFilePath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
      $_lock.Dispose()
    }

    Write-Log -Message "Starting $_toolName ($_launchMode): $_commandLine" -Color Yellow
    Write-Log -Message 'Review and confirm the scan or repair in the utility window.' -Color Cyan
    Write-Progress -Id 40 -Activity 'Outlook data-file repair' -Status "Waiting for $_toolName; follow the instructions in its window" -PercentComplete -1
    $_process = Start-Process @_processArguments
    $_property.ExitCode = $_process.ExitCode
    $_status = if ($_process.ExitCode -eq 0) { 'Completed' } else { 'Failed' }
    $_color = if ($_process.ExitCode -eq 0) { 'Green' } else { 'Yellow' }
    $_detail = "Repair tool exited with code $($_process.ExitCode). Review its log; process exit does not verify file health or that the requested file was repaired."
    Write-Log -Message $_detail -Color $_color
    Add-OperationResult -Results $_results -Target $_target -Source 'OutlookRepair' -Action 'LaunchRepairTool' -Status $_status -Detail $_detail -Property $_property
  }
  else {
    Add-OperationResult -Results $_results -Target $_target -Source 'OutlookRepair' -Action 'LaunchRepairTool' -Status 'Skipped' -Detail 'Declined' -Property $_property
  }
}
catch {
  Write-Warning $_.Exception.Message
  Add-OperationResult -Results $_results -Target $_target -Source 'OutlookRepair' -Action 'LaunchRepairTool' -Status 'Failed' -Detail $_.Exception.Message -Property $_property
}
finally {
  Write-Progress -Id 40 -Activity 'Outlook data-file repair' -Completed
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Repair-OutlookDataFile' -Name 'winkit'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $WhatIfPreference) {
  $_results
}

if (@($_results | Where-Object Status -EQ Failed).Count -gt 0) {
  exit 1
}
