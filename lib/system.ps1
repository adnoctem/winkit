#Requires -Version 5.0

# Shared native methods - compiled once and reused by Get-SystemMemory / Get-SystemUptime
if ($null -eq ('SysInfoNative' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class SysInfoNative {
    [StructLayout(LayoutKind.Sequential)]
    public struct MEMORYSTATUSEX {
        public uint dwLength;
        public uint dwMemoryLoad;
        public ulong ullTotalPhys;
        public ulong ullAvailPhys;
        public ulong ullTotalPageFile;
        public ulong ullAvailPageFile;
        public ulong ullTotalVirtual;
        public ulong ullAvailVirtual;
        public ulong ullAvailExtendedVirtual;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GlobalMemoryStatusEx(ref MEMORYSTATUSEX lpBuffer);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern ulong GetTickCount64();
}
'@ -ErrorAction Stop
}

$script:SysInfoInvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture

function Format-SysInfoInvariant {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]
    $Format,

    [Parameter(ValueFromRemainingArguments = $true)]
    [object[]]
    $ArgumentList
  )

  return [string]::Format($script:SysInfoInvariantCulture, $Format, $ArgumentList)
}

function Resolve-WindowsProductName {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [string]
    $ProductName,

    [Parameter(Mandatory = $false)]
    [int]
    $CurrentBuild = 0
  )

  if ([string]::IsNullOrWhiteSpace($ProductName)) {
    return $ProductName
  }

  if ($CurrentBuild -ge 22000 -and $ProductName -like 'Windows 10*') {
    return $ProductName -replace '^Windows 10', 'Windows 11'
  }

  return $ProductName
}

function Get-OSBuildNumber {
  <#
    .SYNOPSIS
      Returns the Windows build number as an integer.
    .DESCRIPTION
      Reads CurrentBuild from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
      This is a single cheap registry read - far faster than Get-CimInstance
      or any WMI-based approach.  Returns e.g. 22621 (22H2), 22631 (23H2).
    .EXAMPLE
      PS> Get-OSBuildNumber
      22621
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([int])]
  [CmdletBinding()]
  param()

  $build = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'CurrentBuild' -ErrorAction Stop
  return [int]$build
}

function Get-OSDisplayVersion {
  <#
    .SYNOPSIS
      Returns the Windows feature-update display name (e.g. "22H2", "23H2").
    .DESCRIPTION
      Reads DisplayVersion from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
      Falls back to ReleaseId on older builds that lack DisplayVersion.
    .EXAMPLE
      PS> Get-OSDisplayVersion
      23H2
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param()

  try {
    $display = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'DisplayVersion' -ErrorAction Stop
    if ($display) { return $display }
  }
  catch {
    Write-Verbose "DisplayVersion was not available; falling back to ReleaseId."
  }

  $releaseId = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'ReleaseId' -ErrorAction Stop
  return $releaseId
}

function Get-OSEdition {
  <#
    .SYNOPSIS
      Returns the Windows edition SKU (e.g. "Professional", "Enterprise").
    .DESCRIPTION
      Reads EditionID from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
    .EXAMPLE
      PS> Get-OSEdition
      Professional
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param()

  return Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'EditionID' -ErrorAction Stop
}

function Get-OSProductName {
  <#
    .SYNOPSIS
      Returns the full Windows product name string.
    .DESCRIPTION
      Reads ProductName from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
      Returns e.g. "Windows 11 Pro" or "Windows 10 Enterprise".
    .EXAMPLE
      PS> Get-OSProductName
      Windows 11 Pro
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param()

  $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
  $props = Get-ItemProperty -LiteralPath $key -ErrorAction Stop
  return Resolve-WindowsProductName -ProductName $props.ProductName -CurrentBuild ([int]$props.CurrentBuild)
}

function Get-OSVersionInfo {
  <#
    .SYNOPSIS
      Returns a complete snapshot of Windows version metadata from the registry.
    .DESCRIPTION
      Performs a single Get-ItemProperty call against HKLM:\SOFTWARE\Microsoft\
      Windows NT\CurrentVersion and returns all relevant fields in one object.
      Far cheaper than calling the individual Get-OS* functions when you need
      multiple values.
    .EXAMPLE
      PS> Get-OSVersionInfo
    .EXAMPLE
      PS> Get-OSVersionInfo | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
  $props = Get-ItemProperty -LiteralPath $key -ErrorAction Stop

  $installDate = $null
  if ($props.InstallDate) {
    try {
      $unixEpochUtc = [DateTime]::SpecifyKind([DateTime]'1970-01-01T00:00:00Z', [DateTimeKind]::Utc)
      $installDate = $unixEpochUtc.AddSeconds([int64]$props.InstallDate).ToLocalTime()
    }
    catch {
      Write-Verbose "Unable to convert InstallDate registry value: $($props.InstallDate)"
    }
  }

  $currentBuild = [int]$props.CurrentBuild

  [PSCustomObject]@{
    ProductName = Resolve-WindowsProductName -ProductName $props.ProductName -CurrentBuild $currentBuild
    EditionID = $props.EditionID
    DisplayVersion = $props.DisplayVersion
    CurrentBuild = $currentBuild
    UBR = if ($null -ne $props.UBR) { [int]$props.UBR } else { 0 }
    ReleaseId = $props.ReleaseId
    BuildBranch = $props.BuildBranch
    InstallDate = $installDate
    RegisteredOwner = $props.RegisteredOwner
  }
}

function Get-SystemMemory {
  <#
    .SYNOPSIS
      Returns physical memory statistics - total, available, used, and load
      percentage.
    .DESCRIPTION
      Uses kernel32!GlobalMemoryStatusEx via P/Invoke (no CIM/WMI overhead).
      Returns an object with human-readable GiB values and the raw bytes.
    .EXAMPLE
      PS> Get-SystemMemory
    .EXAMPLE
      PS> Get-SystemMemory | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $memInfo = New-Object SysInfoNative+MEMORYSTATUSEX
  $memInfo.dwLength = [System.Runtime.InteropServices.Marshal]::SizeOf($memInfo)

  if (-not [SysInfoNative]::GlobalMemoryStatusEx([ref]$memInfo)) {
    Write-Error 'GlobalMemoryStatusEx failed.'
    return $null
  }

  $totalGiB = [math]::Round($memInfo.ullTotalPhys / 1GB, 2)
  $availableGiB = [math]::Round($memInfo.ullAvailPhys / 1GB, 2)
  $usedGiB = [math]::Round(($memInfo.ullTotalPhys - $memInfo.ullAvailPhys) / 1GB, 2)

  [PSCustomObject]@{
    TotalBytes = $memInfo.ullTotalPhys
    AvailableBytes = $memInfo.ullAvailPhys
    UsedBytes = $memInfo.ullTotalPhys - $memInfo.ullAvailPhys
    LoadPercent = $memInfo.dwMemoryLoad
    TotalGiB = Format-SysInfoInvariant '{0:0.##}' $totalGiB
    AvailableGiB = Format-SysInfoInvariant '{0:0.##}' $availableGiB
    UsedGiB = Format-SysInfoInvariant '{0:0.##}' $usedGiB
  }
}

function Get-SystemDisk {
  <#
    .SYNOPSIS
      Returns disk usage information for fixed volumes backed by physical disks.
    .DESCRIPTION
      Uses Win32_DiskDrive associations to include only volumes that resolve
      back to a physical disk. Provider-backed and cloud-mounted drives such as
      Google Drive are ignored. Filters to fixed drives by default and returns
      total size, free space, used space, and the filesystem type.
    .PARAMETER All
      Include non-fixed volumes when they are backed by a physical disk.
    .EXAMPLE
      PS> Get-SystemDisk
    .EXAMPLE
      PS> Get-SystemDisk -All
    .EXAMPLE
      PS> Get-SystemDisk | Format-Table -AutoSize
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [switch]
    $All = $false
  )

  $logicalDisksByDeviceId = @{}

  try {
    $physicalDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop

    foreach ($physicalDisk in $physicalDisks) {
      $partitions = Get-CimAssociatedInstance -InputObject $physicalDisk -Association Win32_DiskDriveToDiskPartition -ErrorAction Stop
      foreach ($partition in $partitions) {
        $logicalDisks = Get-CimAssociatedInstance -InputObject $partition -Association Win32_LogicalDiskToPartition -ErrorAction Stop
        foreach ($logicalDisk in $logicalDisks) {
          if (-not $All -and [int]$logicalDisk.DriveType -ne 3) { continue }
          if ([string]::IsNullOrWhiteSpace($logicalDisk.DeviceID)) { continue }
          $logicalDisksByDeviceId[$logicalDisk.DeviceID] = $logicalDisk
        }
      }
    }
  }
  catch {
    Write-Verbose "Physical disk association inventory failed: $($_.Exception.Message)"
    $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue
    foreach ($logicalDisk in $logicalDisks) {
      if (-not $All -and [int]$logicalDisk.DriveType -ne 3) { continue }
      if ([string]::IsNullOrWhiteSpace($logicalDisk.DeviceID)) { continue }
      $logicalDisksByDeviceId[$logicalDisk.DeviceID] = $logicalDisk
    }
  }

  foreach ($deviceId in ($logicalDisksByDeviceId.Keys | Sort-Object)) {
    $disk = $logicalDisksByDeviceId[$deviceId]
    if ($null -eq $disk.Size -or [uint64]$disk.Size -eq 0) { continue }

    $totalBytes = [uint64]$disk.Size
    $freeBytes = [uint64]$disk.FreeSpace
    $totalGiB = [math]::Round($totalBytes / 1GB, 2)
    $freeGiB = [math]::Round($freeBytes / 1GB, 2)
    $usedGiB = [math]::Round(($totalBytes - $freeBytes) / 1GB, 2)
    $percentFree = [math]::Round($freeBytes * 100.0 / $totalBytes, 1)

    [PSCustomObject]@{
      Name = $disk.DeviceID
      Label = $disk.VolumeName
      Type = switch ([int]$disk.DriveType) {
        2 { 'Removable' }
        3 { 'Fixed' }
        4 { 'Network' }
        5 { 'CDRom' }
        6 { 'Ram' }
        default { 'Unknown' }
      }
      FileSystem = $disk.FileSystem
      TotalGiB = Format-SysInfoInvariant '{0:0.##}' $totalGiB
      FreeGiB = Format-SysInfoInvariant '{0:0.##}' $freeGiB
      UsedGiB = Format-SysInfoInvariant '{0:0.##}' $usedGiB
      PercentFree = Format-SysInfoInvariant '{0:0.#}' $percentFree
      TotalBytes = $totalBytes
      FreeBytes = $freeBytes
    }
  }
}

function Get-Hostname {
  <#
    .SYNOPSIS
      Returns the computer hostname.
    .DESCRIPTION
      Uses [System.Net.Dns]::GetHostName() and resolves it to a fully-qualified
      domain name when joined to a domain.  Returns both the short hostname and,
      if different, the FQDN.
    .EXAMPLE
      PS> Get-Hostname
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $hostname = [System.Net.Dns]::GetHostName()

  $fqdn = $null
  try {
    $entry = [System.Net.Dns]::GetHostEntry($hostname)
    $fqdn = $entry.HostName
    if ($fqdn -eq $hostname) { $fqdn = $null }
  }
  catch {
    Write-Verbose "Unable to resolve FQDN for hostname: $hostname"
  }

  [PSCustomObject]@{
    Hostname = $hostname
    FQDN = $fqdn
  }
}

function Get-SystemUptime {
  <#
    .SYNOPSIS
      Returns the system uptime (time since last boot).
    .DESCRIPTION
      Uses kernel32!GetTickCount64 via P/Invoke for a non-wrapping, high-
      precision uptime value - no CIM/WMI overhead.  Returns the raw tick
      count, total milliseconds, and a human-readable breakdown.
    .EXAMPLE
      PS> Get-SystemUptime
    .EXAMPLE
      PS> Get-SystemUptime | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $ticksMs = [SysInfoNative]::GetTickCount64()
  $span = [TimeSpan]::FromMilliseconds($ticksMs)

  [PSCustomObject]@{
    TotalMilliseconds = $ticksMs
    Days = $span.Days
    Hours = $span.Hours
    Minutes = $span.Minutes
    Seconds = $span.Seconds
    TotalHours = Format-SysInfoInvariant '{0:0.#}' ([math]::Round($span.TotalHours, 1))
    TotalDays = Format-SysInfoInvariant '{0:0.#}' ([math]::Round($span.TotalDays, 1))
    Display = Format-SysInfoInvariant '{0}d {1:D2}h {2:D2}m {3:D2}s' $span.Days $span.Hours $span.Minutes $span.Seconds
  }
}

function Get-SystemInfo {
  <#
    .SYNOPSIS
      Returns a comprehensive system information snapshot (fetch-style).
    .DESCRIPTION
      Assembles OS version, memory, disk, hostname, and uptime into a single
      structured object.  Disk data is resolved through physical disk
      associations; other data sources use registry reads, .NET APIs, and
      lightweight P/Invoke where possible.
    .EXAMPLE
      PS> Get-SystemInfo | Format-List
    .EXAMPLE
      PS> Get-SystemInfo
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $_os = Get-OSVersionInfo
  $_mem = Get-SystemMemory
  $_disks = Get-SystemDisk
  $_host = Get-Hostname
  $_up = Get-SystemUptime

  [PSCustomObject]@{
    OSProductName = $_os.ProductName
    OSEdition = $_os.EditionID
    OSVersion = $_os.DisplayVersion
    OSBuild = $_os.CurrentBuild
    OSUBRev = $_os.UBR
    Hostname = $_host.Hostname
    FQDN = $_host.FQDN
    TotalMemoryGiB = $_mem.TotalGiB
    MemoryLoadPct = $_mem.LoadPercent
    Disks = ($_disks | ForEach-Object { Format-SysInfoInvariant '{0} {1}GiB/{2}GiB ({3}% free)' $_.Name $_.FreeGiB $_.TotalGiB $_.PercentFree }) -join ' | '
    Uptime = $_up.Display
    InstallDate = $_os.InstallDate
  }
}

function Get-SystemPaths {
  <#
    .SYNOPSIS
      Returns standard winkit directory paths for tools, config, cache, data, and logs.
    .DESCRIPTION
      Provides a single structured lookup for the five canonical winkit folders.
      The -Name parameter controls the subdirectory name used under each root.
      Defaults to 'winkit' so callers can omit it for standard usage.
    .PARAMETER Name
      Subdirectory name under each root. Defaults to 'winkit'.
    .EXAMPLE
      PS> $paths = Get-SystemPaths
      PS> $paths.Data
      C:\ProgramData\winkit
    .EXAMPLE
      PS> Get-SystemPaths -Name 'myapp' | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [string]
    $Name = 'winkit'
  )

  return [PSCustomObject]@{
    Home = Join-Path -Path $env:USERPROFILE -ChildPath $Name
    Config = Join-Path -Path $env:APPDATA -ChildPath $Name
    Cache = Join-Path -Path $env:LOCALAPPDATA -ChildPath $Name
    Data = Join-Path -Path $env:ProgramData -ChildPath $Name
    Logs = Join-Path -Path $env:LOCALAPPDATA -ChildPath "$Name\logs"
  }
}

function Test-HostApplicability {
  <#
    .SYNOPSIS
      Returns $true if the current host meets all supplied applicability constraints.
    .DESCRIPTION
      Gates a setting or operation by build number, OS edition, and processor
      bitness.  Any constraint that is not supplied is treated as "no
      restriction."  All supplied constraints must pass — the check is a
      logical AND across them.

      Builds on the existing Get-OSBuildNumber and Get-OSEdition helpers in
      this file so it shares the same cheap registry-read path.
    .PARAMETER MinBuild
      Minimum Windows build number required (inclusive).  Example: 22000 for
      Windows 11+.
    .PARAMETER MaxBuild
      Maximum Windows build number allowed (inclusive).  Example: 22621 for
      Windows 11 22H2 and below.
    .PARAMETER Edition
      OS edition(s) that qualify.  Accepts an array of strings such as
      @('Enterprise', 'Professional').
    .PARAMETER Bitness
      Required processor architecture: x64 or x86.
    .OUTPUTS
      [bool]
    .EXAMPLE
      PS> if (-not (Test-HostApplicability -MinBuild 22000)) { Write-Log -Message 'Requires Windows 11+'; exit 0 }
    .EXAMPLE
      PS> Test-HostApplicability -Edition @('ServerStandard', 'ServerDatacenter') -Bitness x64
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([bool])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [int]
    $MinBuild,

    [Parameter(Mandatory = $false)]
    [int]
    $MaxBuild,

    [Parameter(Mandatory = $false)]
    [string[]]
    $Edition,

    [Parameter(Mandatory = $false)]
    [ValidateSet('x64', 'x86')]
    [string]
    $Bitness
  )

  $build = Get-OSBuildNumber

  if ($PSBoundParameters.ContainsKey('MinBuild') -and $build -lt $MinBuild) {
    return $false
  }

  if ($PSBoundParameters.ContainsKey('MaxBuild') -and $build -gt $MaxBuild) {
    return $false
  }

  if ($PSBoundParameters.ContainsKey('Edition') -and $Edition.Count -gt 0) {
    $currentEdition = Get-OSEdition
    if ($currentEdition -notin $Edition) {
      return $false
    }
  }

  if ($PSBoundParameters.ContainsKey('Bitness')) {
    $is64Bit = [Environment]::Is64BitOperatingSystem
    if (($Bitness -eq 'x64' -and -not $is64Bit) -or ($Bitness -eq 'x86' -and $is64Bit)) {
      return $false
    }
  }

  return $true
}
