#Requires -Version 5.0

function Test-ObjectProperty {
  param (
    [AllowNull()]
    [object]
    $InputObject,

    [Parameter(Mandatory = $true)]
    [string]
    $Name
  )

  if ($null -eq $InputObject) { return $false }
  return ($InputObject.PSObject.Properties.Name -contains $Name)
}

function Get-ObjectPropertyValue {
  param (
    [AllowNull()]
    [object]
    $InputObject,

    [Parameter(Mandatory = $true)]
    [string]
    $Name,

    [AllowNull()]
    [object]
    $DefaultValue = $null
  )

  if (-not (Test-ObjectProperty -InputObject $InputObject -Name $Name)) {
    return $DefaultValue
  }

  $_value = $InputObject.$Name
  if ($null -eq $_value) {
    return $DefaultValue
  }

  return $_value
}

function ConvertTo-PrintDevice {
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Printer,

    [AllowNull()]
    [bool]
    $Default,

    [string]
    $Source = 'PrintManagement'
  )

  $_name = [string](Get-ObjectPropertyValue -InputObject $Printer -Name 'Name')
  $_deviceId = Get-ObjectPropertyValue -InputObject $Printer -Name 'DeviceID' -DefaultValue $_name

  if (-not $PSBoundParameters.ContainsKey('Default') -or $null -eq $Default) {
    $Default = [bool](Get-ObjectPropertyValue -InputObject $Printer -Name 'Default' -DefaultValue $false)
  }

  [PSCustomObject]@{
    Name = $_name
    DeviceId = $_deviceId
    DriverName = Get-ObjectPropertyValue -InputObject $Printer -Name 'DriverName'
    PortName = Get-ObjectPropertyValue -InputObject $Printer -Name 'PortName'
    Type = Get-ObjectPropertyValue -InputObject $Printer -Name 'Type'
    Shared = [bool](Get-ObjectPropertyValue -InputObject $Printer -Name 'Shared' -DefaultValue $false)
    Published = [bool](Get-ObjectPropertyValue -InputObject $Printer -Name 'Published' -DefaultValue $false)
    Network = [bool](Get-ObjectPropertyValue -InputObject $Printer -Name 'Network' -DefaultValue $false)
    Default = [bool]$Default
    Source = $Source
  }
}

function Get-PrintDeviceDefaultMap {
  [CmdletBinding()]
  param ()

  $_defaultByName = @{}

  try {
    $_cimPrinters = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)
    foreach ($_printer in $_cimPrinters) {
      $_name = [string](Get-ObjectPropertyValue -InputObject $_printer -Name 'Name')
      if ([string]::IsNullOrWhiteSpace($_name)) { continue }
      $_defaultByName[$_name] = [bool](Get-ObjectPropertyValue -InputObject $_printer -Name 'Default' -DefaultValue $false)
    }
  }
  catch {
    Write-Verbose "Could not read Win32_Printer default state: $($_.Exception.Message)"
  }

  return $_defaultByName
}

function Get-PrintDevice {
  <#
    .SYNOPSIS
      Gets installed local print devices.
    .DESCRIPTION
      Uses the PrintManagement module when available and enriches results with
      default-printer state from Win32_Printer. If PrintManagement is
      unavailable or fails, falls back to Win32_Printer directly.
    .EXAMPLE
      PS> Get-PrintDevice
    .LINK
      https://github.com/adnoctem/winkit/lib/devices.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param ()

  $_defaultByName = Get-PrintDeviceDefaultMap
  $_getPrinterCommand = Get-Command -Name Get-Printer -ErrorAction SilentlyContinue

  if ($_getPrinterCommand) {
    try {
      $_printers = @(Get-Printer -ErrorAction Stop)
      foreach ($_printer in $_printers) {
        $_name = [string](Get-ObjectPropertyValue -InputObject $_printer -Name 'Name')
        $_isDefault = if ($_defaultByName.ContainsKey($_name)) { $_defaultByName[$_name] } else { $false }
        ConvertTo-PrintDevice -Printer $_printer -Default $_isDefault -Source 'PrintManagement'
      }

      return
    }
    catch {
      Write-Verbose "Get-Printer failed; falling back to Win32_Printer: $($_.Exception.Message)"
    }
  }

  try {
    $_cimPrinters = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)
    foreach ($_printer in $_cimPrinters) {
      ConvertTo-PrintDevice -Printer $_printer -Source 'CIM'
    }
  }
  catch {
    Write-Verbose "Could not enumerate print devices: $($_.Exception.Message)"
  }
}

function Get-DefaultPrintDevice {
  <#
    .SYNOPSIS
      Gets the default local print device.
    .DESCRIPTION
      Queries Win32_Printer for the printer where Default is true and returns
      the same normalized object shape as Get-PrintDevice.
    .EXAMPLE
      PS> Get-DefaultPrintDevice
    .LINK
      https://github.com/adnoctem/winkit/lib/devices.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param ()

  try {
    $_printers = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | Where-Object { $_.Default -eq $true })
    foreach ($_printer in $_printers) {
      ConvertTo-PrintDevice -Printer $_printer -Source 'CIM'
    }
  }
  catch {
    Write-Verbose "Could not read default print device: $($_.Exception.Message)"
  }
}

function Set-DefaultPrintDevice {
  <#
    .SYNOPSIS
      Sets the default local print device.
    .DESCRIPTION
      Resolves a printer by exact name through Win32_Printer and calls its
      SetDefaultPrinter CIM method. The function returns a structured operation
      result indicating whether the default printer was changed or skipped.
    .PARAMETER Name
      Exact printer name to make the default.
    .PARAMETER PassThru
      Accepted for consistency with scripts that request operation-result
      output. The result object is returned in all cases.
    .EXAMPLE
      PS> Set-DefaultPrintDevice -Name 'Brother MFC-L2860DWE LAN'
    .LINK
      https://github.com/adnoctem/winkit/lib/devices.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Name,

    [switch]
    $PassThru
  )

  [void]$PassThru
  $_printers = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | Where-Object { $_.Name -eq $Name })

  if ($_printers.Count -eq 0) {
    Write-Error -Message "Print device '$Name' was not found." -ErrorAction Stop
  }

  if ($_printers.Count -gt 1) {
    Write-Error -Message "Multiple print devices named '$Name' were found. Refusing to choose one." -ErrorAction Stop
  }

  $_printer = $_printers[0]
  if ([bool](Get-ObjectPropertyValue -InputObject $_printer -Name 'Default' -DefaultValue $false)) {
    return New-OperationResult -Target $Name -Source 'Win32_Printer' -Action 'SetDefault' -Status 'Skipped' -Detail 'AlreadyDefault'
  }

  if (-not $PSCmdlet.ShouldProcess($Name, 'Set default print device')) {
    return New-OperationResult -Target $Name -Source 'Win32_Printer' -Action 'SetDefault' -Status 'Skipped' -Detail 'WhatIf'
  }

  try {
    $_result = Invoke-CimMethod -InputObject $_printer -MethodName 'SetDefaultPrinter' -ErrorAction Stop
    $_returnValue = Get-ObjectPropertyValue -InputObject $_result -Name 'ReturnValue' -DefaultValue 0

    if ([int]$_returnValue -eq 0) {
      return New-OperationResult -Target $Name -Source 'Win32_Printer' -Action 'SetDefault' -Status 'Completed' -Detail 'Default printer changed.'
    }

    return New-OperationResult -Target $Name -Source 'Win32_Printer' -Action 'SetDefault' -Status 'Failed' -Detail "SetDefaultPrinter returned $_returnValue."
  }
  catch {
    return New-OperationResult -Target $Name -Source 'Win32_Printer' -Action 'SetDefault' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Get-WiaPropertyValue {
  param (
    [AllowNull()]
    [object]
    $DeviceInfo,

    [Parameter(Mandatory = $true)]
    [string]
    $Name
  )

  if ($null -eq $DeviceInfo) { return $null }

  foreach ($_property in @($DeviceInfo.Properties)) {
    if ($_property.Name -eq $Name) {
      return $_property.Value
    }
  }

  return $null
}

function ConvertTo-ScanDevice {
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $DeviceInfo
  )

  [PSCustomObject]@{
    Name = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'Name'
    DeviceId = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'Unique Device ID'
    Manufacturer = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'Manufacturer'
    Type = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'Type'
    Port = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'Port'
    Server = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'Server'
    DriverVersion = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'Driver Version'
    WiaVersion = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'WIA Version'
    PnpId = Get-WiaPropertyValue -DeviceInfo $DeviceInfo -Name 'PnP ID String'
    Source = 'WIA'
  }
}

function Get-ScanDevice {
  <#
    .SYNOPSIS
      Gets WIA-compatible scan devices.
    .DESCRIPTION
      Enumerates Windows Image Acquisition device metadata through
      WIA.DeviceManager. This is discovery-only; Windows does not provide a
      universal default scanner setting equivalent to the default printer.
    .EXAMPLE
      PS> Get-ScanDevice
    .LINK
      https://github.com/adnoctem/winkit/lib/devices.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param ()

  try {
    $_deviceManager = New-Object -ComObject WIA.DeviceManager
    foreach ($_deviceInfo in @($_deviceManager.DeviceInfos)) {
      ConvertTo-ScanDevice -DeviceInfo $_deviceInfo
    }
  }
  catch {
    Write-Verbose "Could not enumerate WIA scan devices: $($_.Exception.Message)"
  }
  finally {
    if (Get-Command -Name Remove-ComObject -ErrorAction SilentlyContinue) {
      Remove-ComObject $_deviceManager
    }
  }
}
