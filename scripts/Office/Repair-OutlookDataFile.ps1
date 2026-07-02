#Requires -Version 5.0

<#
.SYNOPSIS
  Repairs an Outlook PST or OST data file with Microsoft repair tools.
.DESCRIPTION
  Locates the installed Outlook data-file repair utilities and runs ScanPST.exe
  or ScanOST.exe against the file passed through -Path. By default, PST files
  use ScanPST and OST files prefer ScanOST when available, falling back to
  ScanPST for modern Outlook installations where ScanOST is no longer present.

  Close Outlook before repairing an attached data file. The Microsoft repair
  tools may display their own UI and can require interactive confirmation.
.PARAMETER Path
  PST or OST data file to scan or repair.
.PARAMETER Tool
  Repair tool to use. Auto selects based on file extension and installed tools.
.PARAMETER ToolPath
  Explicit path to ScanPST.exe or ScanOST.exe. When supplied, this overrides
  automatic Outlook tool discovery.
.PARAMETER DryRun
  Preview the repair command without executing it.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> .\Repair-OutlookDataFile.ps1 -Path C:\Users\User\Documents\Outlook Files\archive.pst
.EXAMPLE
  PS> .\Repair-OutlookDataFile.ps1 -Path C:\Users\User\AppData\Local\Microsoft\Outlook\mail.ost -Tool ScanOST
.EXAMPLE
  PS> .\Repair-OutlookDataFile.ps1 -Path D:\Mail\archive.pst -ToolPath 'C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE' -DryRun
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Parameter(Mandatory = $true)]
  [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
  [string]
  $Path,

  [ValidateSet('Auto', 'ScanPST', 'ScanOST')]
  [string]
  $Tool = 'Auto',

  [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
  [string]
  $ToolPath,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - Outlook data-file repair will not be started`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$_dataFilePath = (Resolve-Path -LiteralPath $Path).ProviderPath
$_extension = [System.IO.Path]::GetExtension($_dataFilePath).ToLowerInvariant()

function Resolve-OutlookDataFileRepairTool {
  param (
    [string]
    $RequestedTool,

    [string]
    $RequestedToolPath,

    [string]
    $DataFileExtension
  )

  if (-not [string]::IsNullOrWhiteSpace($RequestedToolPath)) {
    $_resolvedToolPath = (Resolve-Path -LiteralPath $RequestedToolPath).ProviderPath
    $_toolName = [System.IO.Path]::GetFileNameWithoutExtension($_resolvedToolPath)
    return [PSCustomObject]@{
      Name = $_toolName
      Path = $_resolvedToolPath
      Detail = 'ExplicitToolPath'
    }
  }

  $_preferredTools = switch ($RequestedTool) {
    'ScanPST' { @('ScanPST') }
    'ScanOST' { @('ScanOST') }
    default {
      if ($DataFileExtension -eq '.ost') {
        @('ScanOST', 'ScanPST')
      }
      else {
        @('ScanPST')
      }
    }
  }

  foreach ($_toolName in $_preferredTools) {
    $_tool = Find-OutlookRepairTool -Name $_toolName | Select-Object -First 1
    if ($_tool) {
      return [PSCustomObject]@{
        Name = $_tool.Name
        Path = $_tool.Path
        Detail = $_tool.InstallationPath
      }
    }
  }

  return $null
}

if ($_extension -notin @('.pst', '.ost')) {
  Write-Log -Message "Unsupported Outlook data-file extension: $_extension" -Color Red
  Add-OperationResult -Results $_results -Target $_dataFilePath -Source 'OutlookRepair' -Action 'Repair' -Status 'Failed' -Detail 'Expected .pst or .ost file.'

  if ($PassThru -or $DryRun) {
    $_results
  }

  exit 1
}

$_repairTool = Resolve-OutlookDataFileRepairTool -RequestedTool $Tool -RequestedToolPath $ToolPath -DataFileExtension $_extension
if (-not $_repairTool) {
  $_detail = if ($Tool -eq 'Auto') {
    'No ScanPST.exe or compatible ScanOST.exe installation was found.'
  }
  else {
    "No $Tool repair tool was found."
  }

  Write-Log -Message $_detail -Color Red
  Add-OperationResult -Results $_results -Target $_dataFilePath -Source 'OutlookRepair' -Action 'Repair' -Status 'Failed' -Detail $_detail

  if ($PassThru -or $DryRun) {
    $_results
  }

  exit 1
}

$_toolPath = $_repairTool.Path
$_toolName = $_repairTool.Name
$_commandLine = "`"$_toolPath`" `"$_dataFilePath`""

if ($DryRun) {
  Write-Log -Message "[DRY RUN] Would run: $_commandLine" -Color Yellow
  Add-OperationResult -Results $_results -Target $_dataFilePath -Source 'OutlookRepair' -Action 'Repair' -Status 'Skipped' -Detail "DryRun: $_commandLine" -Property @{
    Tool = $_toolName
    ToolPath = $_toolPath
  }
}
elseif ($PSCmdlet.ShouldProcess($_dataFilePath, "Repair with $_toolName")) {
  Write-Log -Message "Starting $_toolName for Outlook data file..." -Color Yellow
  Write-Log -Message "  File: $_dataFilePath" -Color Gray
  Write-Log -Message "  Tool: $_toolPath" -Color Gray

  try {
    $_process = Start-Process -FilePath $_toolPath -ArgumentList @($_dataFilePath) -Wait -PassThru -ErrorAction Stop
    $_status = if ($_process.ExitCode -eq 0) { 'Completed' } else { "ExitCode:$($_process.ExitCode)" }
    $_color = if ($_process.ExitCode -eq 0) { 'Green' } else { 'Yellow' }

    Write-Log -Message "Outlook data-file repair finished with exit code $($_process.ExitCode)." -Color $_color
    Add-OperationResult -Results $_results -Target $_dataFilePath -Source 'OutlookRepair' -Action 'Repair' -Status $_status -Detail "Tool: $_toolName" -Property @{
      Tool = $_toolName
      ToolPath = $_toolPath
      ExitCode = $_process.ExitCode
    }
  }
  catch {
    Write-Log -Message "FAILED - could not start ${_toolName}: $($_.Exception.Message)" -Color Red
    Add-OperationResult -Results $_results -Target $_dataFilePath -Source 'OutlookRepair' -Action 'Repair' -Status 'Failed' -Detail $_.Exception.Message -Property @{
      Tool = $_toolName
      ToolPath = $_toolPath
    }
  }
}
else {
  Add-OperationResult -Results $_results -Target $_dataFilePath -Source 'OutlookRepair' -Action 'Repair' -Status 'Skipped' -Detail 'WhatIf' -Property @{
    Tool = $_toolName
    ToolPath = $_toolPath
  }
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Repair-OutlookDataFile'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
if ($_failed -gt 0) {
  exit 1
}
