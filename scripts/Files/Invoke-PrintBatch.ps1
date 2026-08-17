#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Prints a batch of files (PDF by default) to a printer.

.DESCRIPTION
  Collects files matching -Filter from -Path and prints them one after
  another. Two print methods are supported:

    Default (recommended for interactive use) - Uses the registered file
    handler through the ShellExecute Print / PrintTo verb (Start-Process
    -Verb). No PDF reader installation is required: on current Windows the
    handler is Microsoft Edge, which prints silently. Adobe Acrobat/Reader
    is only involved if it happens to be the registered handler.

    Sumatra - Uses a portable SumatraPDF.exe (-silent -print-to), the
    battle-tested approach for headless or service contexts where launching
    a GUI handler (Edge/Acrobat) fails or is unreliable. Point at an
    existing executable with -SumatraPath or let the script download the
    portable build once via -InstallSumatraPortable.

  Without -Printer the system default printer is used. A short delay
  between jobs (-PrintDelaySeconds) avoids the intermittent failures that
  occur when jobs are submitted in a tight loop; -Wait additionally blocks
  on the handler process (meaningful with Sumatra/Acrobat; Microsoft Edge
  tends to linger, so waiting is of limited use there).

  The script needs no elevation. Use -DryRun to preview the batch.

.PARAMETER Path
  Directory (or single file) to scan. Defaults to the current directory.

.PARAMETER Filter
  File pattern. Defaults to *.pdf.

.PARAMETER Recurse
  Include files in subdirectories of -Path.

.PARAMETER Printer
  Printer name to print to. Defaults to the system default printer.

.PARAMETER Method
  Print method: Default (registered handler via Print/PrintTo verb) or
  Sumatra (portable SumatraPDF.exe).

.PARAMETER SumatraPath
  Path to SumatraPDF.exe for the Sumatra method. When omitted,
  $env:LOCALAPPDATA\winkit\bin\sumatra\SumatraPDF.exe is assumed.

.PARAMETER InstallSumatraPortable
  Download and extract the portable SumatraPDF build to
  $env:LOCALAPPDATA\winkit\bin\sumatra when the Sumatra method is selected
  and the executable is missing.

.PARAMETER PrintDelaySeconds
  Pause between print jobs. Defaults to 2 seconds.

.PARAMETER Wait
  Block on the handler process until it exits before continuing.

.PARAMETER DryRun
  Preview the batch without printing anything.

.PARAMETER PassThru
  Return structured operation results.

.EXAMPLE
  PS> ./Invoke-PrintBatch.ps1 -Path C:\Reports
  Prints all PDFs in C:\Reports to the default printer.

.EXAMPLE
  PS> ./Invoke-PrintBatch.ps1 -Path C:\Reports -Filter *.pdf -Recurse -Printer 'Reception Laser'
  Prints all PDFs below C:\Reports to the named printer.

.EXAMPLE
  PS> ./Invoke-PrintBatch.ps1 -Path C:\Reports -Method Sumatra -InstallSumatraPortable -Printer 'Reception Laser' -Wait
  Downloads portable SumatraPDF, then prints the batch headless with
  process-wait between jobs.

.EXAMPLE
  PS> ./Invoke-PrintBatch.ps1 -Path C:\Reports -DryRun

.LINK
  https://github.com/adnoctem/winkit
  https://brndmp.olafritman.com/batch-print-pdf-with-powershell/
  https://pipe.how/invoke-print/
  https://medium.com/@mayberryjalin/powershell-streamlining-batch-pdf-printing-301f25c1cd03
  https://community.spiceworks.com/t/trying-to-send-pdf-print-jobs-to-a-network-printer/957742

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Carriage-return progress lines require Write-Host for in-place console updates.')]

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string]
  $Path = '.',

  [Parameter(Mandatory = $false)]
  [string]
  $Filter = '*.pdf',

  [Parameter(Mandatory = $false)]
  [switch]
  $Recurse,

  [Parameter(Mandatory = $false)]
  [string]
  $Printer,

  [Parameter(Mandatory = $false)]
  [ValidateSet('Default', 'Sumatra')]
  [string]
  $Method = 'Default',

  [Parameter(Mandatory = $false)]
  [string]
  $SumatraPath,

  [Parameter(Mandatory = $false)]
  [switch]
  $InstallSumatraPortable,

  [Parameter(Mandatory = $false)]
  [int]
  $PrintDelaySeconds = 2,

  [Parameter(Mandatory = $false)]
  [switch]
  $Wait,

  [Parameter(Mandatory = $false)]
  [switch]
  $DryRun,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force
$ProgressPreference = 'SilentlyContinue'

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no files will be printed`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList

$sumatraDirectory = Join-Path $env:LOCALAPPDATA 'winkit\bin\sumatra'
$sumatraDefaultPath = Join-Path $sumatraDirectory 'SumatraPDF.exe'
$sumatraDownloadUrl = 'https://www.sumatrapdfreader.org/dl/rel/SumatraPDF-3.5.2-64.zip'

# ---- Helpers ----------------------------------------------------------------

function Get-PdfPrintHandler {
  <#
    Reports the registered command line that would handle PDF printing,
    purely as a diagnostic (best effort; may be empty).
  #>
  [CmdletBinding()]
  param()

  $command = $null
  foreach ($keyPath in @(
      'Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\pdf\shell\Printto\command',
      'Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\pdf\shell\open\command',
      'Registry::HKEY_CLASSES_ROOT\.pdf\shell\open\command'
    )) {
    try {
      $item = Get-ItemProperty -Path $keyPath -ErrorAction Stop
      if ($item.'(default)') {
        $command = $item.'(default)'
        break
      }
    }
    catch {
      $command = $null
    }
  }

  if ($command) {
    return ($command -replace '"', '')
  }
  return $null
}

function Get-SelectedPrinterName {
  <#
    Returns the printer name to print to: the explicitly requested one
    (validated against the installed print devices) or the system default.
    Returns $null when no -Printer was given and no default could be found.
  #>
  [CmdletBinding()]
  param(
    [System.Collections.ArrayList]$Diagnostics
  )

  $printDevices = @(Get-PrintDevice -ErrorAction SilentlyContinue)

  if ($Printer) {
    $match = $printDevices | Where-Object { $_.Name -eq $Printer } | Select-Object -First 1
    if (-not $match) {
      if ($null -ne $Diagnostics) {
        Add-OperationResult -Results $Diagnostics -Target $Printer -Source 'Print' -Action 'Validate' -Status 'Failed' -Detail 'Requested printer not found among the installed print devices.'
      }
      return $null
    }
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target $Printer -Source 'Print' -Action 'Validate' -Status 'Completed' -Detail 'Requested printer found.'
    }
    return $Printer
  }

  $defaultDevice = $printDevices | Where-Object { $_.Default } | Select-Object -First 1
  if (-not $defaultDevice) {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target 'DefaultPrinter' -Source 'Print' -Action 'Validate' -Status 'Warn' -Detail 'No default printer detected; the handler will use whatever printer Windows selects.'
    }
    return $null
  }

  if ($null -ne $Diagnostics) {
    Add-OperationResult -Results $Diagnostics -Target $defaultDevice.Name -Source 'Print' -Action 'Validate' -Status 'Completed' -Detail 'Using the system default printer.'
  }
  return $defaultDevice.Name
}

function Install-SumatraPortable {
  <#
    Downloads and extracts the portable SumatraPDF build into
    $env:LOCALAPPDATA\winkit\bin\sumatra. Returns the executable path, or
    $null on failure.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [System.Collections.ArrayList]$Diagnostics
  )

  if (-not $PSCmdlet.ShouldProcess('SumatraPDF portable', 'Download and extract')) {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target 'SumatraPDF' -Source 'Print' -Action 'Install' -Status 'Skipped' -Detail 'WhatIf - download skipped.'
    }
    return $null
  }

  $work = Join-Path $env:TEMP "sumatra-download-$(New-Guid)"
  $null = New-Item -Path $work -ItemType Directory -Force -ErrorAction SilentlyContinue

  try {
    $zipPath = Join-Path $work 'SumatraPDF.zip'
    Write-Log -Message 'Downloading portable SumatraPDF...' -Color Yellow
    Invoke-WebRequest -Uri $sumatraDownloadUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop

    $null = New-Item -Path $sumatraDirectory -ItemType Directory -Force -ErrorAction SilentlyContinue
    Expand-Archive -Path $zipPath -DestinationPath $sumatraDirectory -Force -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $sumatraDefaultPath -PathType Leaf)) {
      throw "SumatraPDF.exe not found after extraction in $sumatraDirectory."
    }

    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target 'SumatraPDF' -Source 'Print' -Action 'Install' -Status 'Completed' -Detail "Portable SumatraPDF installed at $sumatraDefaultPath."
    }
    return $sumatraDefaultPath
  }
  catch {
    if ($null -ne $Diagnostics) {
      Add-OperationResult -Results $Diagnostics -Target 'SumatraPDF' -Source 'Print' -Action 'Install' -Status 'Failed' -Detail $_.Exception.Message
    }
    return $null
  }
  finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
  }
}

# ---- Collect the batch ------------------------------------------------------

$files = @(Get-ChildItem -LiteralPath $Path -Filter $Filter -File -Recurse:$Recurse -ErrorAction SilentlyContinue | Sort-Object -Property Name)

if ($files.Count -eq 0) {
  Write-Log -Message "No files matching '$Filter' found in '$Path'." -Color Yellow
  Add-OperationResult -Results $_results -Target $Filter -Source 'Print' -Action 'Collect' -Status 'Skipped' -Detail 'NoFilesFound'
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Invoke-PrintBatch'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

Write-Log -Message "Found $($files.Count) file(s) matching '$Filter' in '$Path'." -Color Cyan

# ---- Diagnostics ------------------------------------------------------------

$handler = Get-PdfPrintHandler
Add-OperationResult -Results $_results -Target 'PDF Handler' -Source 'Print' -Action 'Inspect' -Status 'Completed' -Detail $(if ($handler) { "Registered PDF print handler: $handler" } else { 'No PDF print handler command found in the registry.' })

$printerName = Get-SelectedPrinterName -Diagnostics $_results
if ($Printer -and -not $printerName) {
  Write-Log -Message "Printer '$Printer' was not found. Verify the printer name and re-run." -Color Red
  if ($PassThru -or $DryRun) { $_results }
  exit 1
}

# ---- Method preparation -----------------------------------------------------

$sumatraExecutable = $null
if ($Method -eq 'Sumatra') {
  $sumatraExecutable = if ($SumatraPath) { $SumatraPath } else { $sumatraDefaultPath }

  if (-not (Test-Path -LiteralPath $sumatraExecutable -PathType Leaf)) {
    if ($InstallSumatraPortable) {
      $sumatraExecutable = Install-SumatraPortable -Diagnostics $_results
    }
    else {
      Add-OperationResult -Results $_results -Target 'SumatraPDF' -Source 'Print' -Action 'Validate' -Status 'Failed' -Detail "SumatraPDF.exe not found at $sumatraExecutable. Pass -SumatraPath or re-run with -InstallSumatraPortable."
    }
  }

  if (-not $sumatraExecutable -or -not (Test-Path -LiteralPath $sumatraExecutable -PathType Leaf)) {
    Write-Log -Message 'SumatraPDF executable is not available for the Sumatra method.' -Color Red
    if ($PassThru -or $DryRun) { $_results }
    exit 1
  }
}

# ---- Print the batch --------------------------------------------------------

if (-not $PSCmdlet.ShouldProcess("$($files.Count) file(s)", "Print to $(if ($printerName) { $printerName } else { 'the default printer' }) via $Method")) {
  foreach ($file in $files) {
    Add-OperationResult -Results $_results -Target $file.Name -Source 'Print' -Action 'Print' -Status 'Skipped' -Detail 'WhatIf - print skipped.'
  }
  $_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Invoke-PrintBatch'
  if ($_operationLog) { Write-Log -Message "Operation log: $_operationLog" -Color Gray }
  if ($PassThru -or $DryRun) { $_results }
  exit 0
}

$total = $files.Count
$index = 0
$failedCount = 0

foreach ($file in $files) {
  $index++
  $pct = [int](($index / $total) * 100)
  Write-Progress -Activity 'Printing batch' -Status "$($file.Name) ($index/$total)" -PercentComplete $pct
  Write-Host ("`rPrinting ({0}/{1}): {2,-60}" -f $index, $total, $file.Name) -NoNewline -ForegroundColor Cyan

  try {
    if ($Method -eq 'Sumatra') {
      $arguments = @('-silent')
      if ($printerName) {
        $arguments += @('-print-to', "`"$printerName`"")
      }
      else {
        $arguments += @('-print-to-default')
      }
      $arguments += @('-print-settings', 'shrink', "`"$($file.FullName)`"")

      $process = Start-Process -FilePath $sumatraExecutable -ArgumentList $arguments -PassThru -Wait:$Wait -ErrorAction Stop
      if ($Wait -and $process.ExitCode -ne 0) {
        throw "SumatraPDF exited with code $($process.ExitCode)."
      }
      Add-OperationResult -Results $_results -Target $file.Name -Source 'Print' -Action 'Print' -Status 'Completed' -Detail "Printed via Sumatra to $(if ($printerName) { $printerName } else { 'default printer' })."
    }
    else {
      if ($printerName) {
        Start-Process -FilePath $file.FullName -Verb PrintTo -ArgumentList "`"$printerName`"" -Wait:$Wait -ErrorAction Stop
      }
      else {
        Start-Process -FilePath $file.FullName -Verb Print -Wait:$Wait -ErrorAction Stop
      }
      Add-OperationResult -Results $_results -Target $file.Name -Source 'Print' -Action 'Print' -Status 'Completed' -Detail "Submitted via the registered handler to $(if ($printerName) { $printerName } else { 'default printer' })."
    }
  }
  catch {
    $failedCount++
    Add-OperationResult -Results $_results -Target $file.Name -Source 'Print' -Action 'Print' -Status 'Failed' -Detail $_.Exception.Message
  }

  if ($index -lt $total -and $PrintDelaySeconds -gt 0) {
    Start-Sleep -Seconds $PrintDelaySeconds
  }
}

Write-Progress -Activity 'Printing batch' -Completed
Write-Host ("`rPrinted {0}/{1} file(s).{2}" -f ($total - $failedCount), $total, (' ' * 40)) -ForegroundColor Cyan

# ---- Summary ----------------------------------------------------------------

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Invoke-PrintBatch'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($failedCount -gt 0) {
  Write-Log -Message "Completed with $failedCount failed print job(s)." -Color Red
  exit 1
}

Write-Log -Message "All $total file(s) submitted for printing." -Color Green
exit 0
