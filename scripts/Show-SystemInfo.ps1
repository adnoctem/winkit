#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

<#
.SYNOPSIS
  Displays a neofetch-style summary of the current system.
.DESCRIPTION
  Renders a compact "fetch" view of the local machine - user, hostname, OS,
  host model, kernel/build, uptime, shell, installed package counts, memory
  and disk usage - next to a small two-tone rendition of the Windows flag.
  All reusable data is gathered through winkit functions; presentation-only
  details are resolved locally.

  The output is produced by Show-SystemInfo, which is also aliased to
  'fetch'. Running this script directly prints the summary once. Dot-source it
  to make both Show-SystemInfo and the 'fetch' alias available for the rest of
  the session without printing immediately.
.EXAMPLE
  PS> .\Show-SystemInfo.ps1
.EXAMPLE
  PS> . .\scripts\Show-SystemInfo.ps1
  PS> fetch
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

# -----------------------------------------------------------------------------

function Show-SystemInfo {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingWriteHost", "", Justification = "Fetch-style output needs per-segment foreground colors.")]

  <#
    .SYNOPSIS
      Prints a neofetch-style summary of the current system.
    .DESCRIPTION
      Combines Get-OSVersionInfo, Get-SystemMemory, Get-SystemDisk,
      Get-Hostname, Get-SystemUptime, Get-UserInfo and Get-PackageCount with a
      few presentation-only lookups and renders them next to a small ANSI
      rendition of the Windows flag.
    .EXAMPLE
      PS> Show-SystemInfo
    .EXAMPLE
      PS> fetch
    .LINK
      https://github.com/adnoctem/winkit/scripts/Show-SystemInfo.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([void])]
  [CmdletBinding()]
  [Alias('fetch')]
  param()

  Import-Module PSFoundation -Force

  $invariant = [System.Globalization.CultureInfo]::InvariantCulture

  # ---- Gather data --------------------------------------------------------------
  $os = Get-OSVersionInfo
  $mem = Get-SystemMemory
  $disks = Get-SystemDisk
  $hostInfo = Get-Hostname
  $uptime = Get-SystemUptime
  $user = Get-UserInfo

  $shortUser = ($user.UserName -split '\\')[-1]

  $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
  $hostModel = if ($computerSystem) {
    "$($computerSystem.Manufacturer) $($computerSystem.Model)".Trim()
  }
  else {
    'Unknown'
  }

  $osVersion = if ($os.DisplayVersion) { $os.DisplayVersion } else { $os.ReleaseId }
  $osLine = "$($os.ProductName) $osVersion (Build $($os.CurrentBuild).$($os.UBR))"
  $kernelLine = "10.0.$($os.CurrentBuild).$($os.UBR)"
  $shellLine = "PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
  $packageCount = Get-PackageCount
  $pkgsLine = "$($packageCount.Programs) (Programs), $($packageCount.Appx) (Appx)"

  $memLine = "$($mem.UsedGiB) GiB / $($mem.TotalGiB) GiB ($($mem.LoadPercent)% used)"

  $systemDisk = $disks | Where-Object { $_.Name -eq $env:SystemDrive } | Select-Object -First 1
  if (-not $systemDisk) { $systemDisk = $disks | Select-Object -First 1 }

  $diskLine = if ($systemDisk) {
    $percentUsed = (100 - [double]$systemDisk.PercentFree)
    "$($systemDisk.Name) $($systemDisk.UsedGiB) GiB / $($systemDisk.TotalGiB) GiB ($($percentUsed.ToString('0.#', $invariant))% used)"
  }
  else {
    'Unknown'
  }

  # ---- Logo ----------------------------------------------------------------------
  # A small two-pane rendition of the Windows flag, classic four-color scheme.
  $pane = ([char]0x2588).ToString() * 11
  $blank = ' ' * 11

  $logoRows = @(
    @{
      L = $blank
      LC = 'White'
      R = $blank
      RC = 'White'
    }
    @{
      L = $pane
      LC = 'Red'
      R = $pane
      RC = 'Green'
    }
    @{
      L = $pane
      LC = 'Red'
      R = $pane
      RC = 'Green'
    }
    @{
      L = $pane
      LC = 'Red'
      R = $pane
      RC = 'Green'
    }
    @{
      L = $pane
      LC = 'Red'
      R = $pane
      RC = 'Green'
    }
    @{
      L = $blank
      LC = 'White'
      R = $blank
      RC = 'White'
    }
    @{
      L = $pane
      LC = 'Blue'
      R = $pane
      RC = 'Yellow'
    }
    @{
      L = $pane
      LC = 'Blue'
      R = $pane
      RC = 'Yellow'
    }
    @{
      L = $pane
      LC = 'Blue'
      R = $pane
      RC = 'Yellow'
    }
    @{
      L = $pane
      LC = 'Blue'
      R = $pane
      RC = 'Yellow'
    }
  )

  # ---- Info lines ------------------------------------------------------------------
  $headerText = "$shortUser@$($hostInfo.Hostname)"

  $infoRows = @(
    @{ Type = 'Header' }
    @{ Type = 'Separator' }
    @{
      Type = 'Field'
      Label = 'os'
      Value = $osLine
      Color = 'Green'
    }
    @{
      Type = 'Field'
      Label = 'host'
      Value = $hostModel
      Color = 'Yellow'
    }
    @{
      Type = 'Field'
      Label = 'kernel'
      Value = $kernelLine
      Color = 'Cyan'
    }
    @{
      Type = 'Field'
      Label = 'uptime'
      Value = $uptime.Display
      Color = 'Magenta'
    }
    @{
      Type = 'Field'
      Label = 'shell'
      Value = $shellLine
      Color = 'Red'
    }
    @{
      Type = 'Field'
      Label = 'pkgs'
      Value = $pkgsLine
      Color = 'Blue'
    }
    @{
      Type = 'Field'
      Label = 'memory'
      Value = $memLine
      Color = 'DarkYellow'
    }
    @{
      Type = 'Field'
      Label = 'disk'
      Value = $diskLine
      Color = 'DarkCyan'
    }
  )

  # ---- Render --------------------------------------------------------------------
  Write-Host ''

  for ($i = 0; $i -lt $logoRows.Count; $i++) {
    $logo = $logoRows[$i]
    $info = $infoRows[$i]

    Write-Host -NoNewline " $($logo.L) " -ForegroundColor $logo.LC
    Write-Host -NoNewline "$($logo.R)  " -ForegroundColor $logo.RC

    switch ($info.Type) {
      'Header' {
        Write-Host $headerText -ForegroundColor Cyan
      }
      'Separator' {
        Write-Host (([char]0x2500).ToString() * $headerText.Length) -ForegroundColor DarkGray
      }
      'Field' {
        Write-Host -NoNewline ("{0,-7}" -f $info.Label) -ForegroundColor $info.Color
        Write-Host $info.Value -ForegroundColor White
      }
    }
  }

  Write-Host ''
  Write-Host -NoNewline '  '
  foreach ($color in [enum]::GetValues([System.ConsoleColor])) {
    Write-Host -NoNewline '###' -ForegroundColor $color
  }
  Write-Host ''
  Write-Host ''
}

if ($MyInvocation.InvocationName -ne '.') {
  Show-SystemInfo
}
