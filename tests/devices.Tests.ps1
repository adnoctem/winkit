#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../lib/common.ps1
  . $PSScriptRoot/../lib/devices.ps1
}

Describe 'Get-PrintDevice' {
  It 'uses PrintManagement and enriches default state from CIM' {
    Mock Get-Command { [PSCustomObject]@{ Name = 'Get-Printer' } } -ParameterFilter { $Name -eq 'Get-Printer' }
    Mock Get-Printer {
      @(
        [PSCustomObject]@{
          Name = 'Office'
          DriverName = 'Office Driver'
          PortName = 'IP_10.0.0.5'
          Type = 'Local'
          Shared = $true
          Published = $false
        }
        [PSCustomObject]@{
          Name = 'PDF'
          DriverName = 'PDF Driver'
          PortName = 'PORTPROMPT:'
          Type = 'Local'
          Shared = $false
          Published = $false
        }
      )
    }
    Mock Get-CimInstance {
      @(
        [PSCustomObject]@{ Name = 'Office'; DeviceID = 'Office'; Default = $true }
        [PSCustomObject]@{ Name = 'PDF'; DeviceID = 'PDF'; Default = $false }
      )
    } -ParameterFilter { $ClassName -eq 'Win32_Printer' }

    $result = @(Get-PrintDevice)

    $result.Count | Should -Be 2
    $result[0].Name | Should -Be 'Office'
    $result[0].DeviceId | Should -Be 'Office'
    $result[0].Default | Should -BeTrue
    $result[0].Source | Should -Be 'PrintManagement'
    $result[1].Default | Should -BeFalse
  }

  It 'falls back to CIM when PrintManagement is unavailable' {
    Mock Get-Command { $null } -ParameterFilter { $Name -eq 'Get-Printer' }
    Mock Get-CimInstance {
      @([PSCustomObject]@{
          Name = 'Fallback'
          DeviceID = 'Fallback'
          DriverName = 'Fallback Driver'
          PortName = 'LPT1:'
          Network = $false
          Shared = $false
          Default = $true
        })
    } -ParameterFilter { $ClassName -eq 'Win32_Printer' }

    $result = @(Get-PrintDevice)

    $result.Count | Should -Be 1
    $result[0].Name | Should -Be 'Fallback'
    $result[0].Default | Should -BeTrue
    $result[0].Source | Should -Be 'CIM'
  }
}

Describe 'Get-DefaultPrintDevice' {
  It 'returns the printer marked as default by CIM' {
    Mock Get-CimInstance {
      @(
        [PSCustomObject]@{ Name = 'Office'; DeviceID = 'Office'; Default = $true; DriverName = 'Office Driver' }
        [PSCustomObject]@{ Name = 'PDF'; DeviceID = 'PDF'; Default = $false; DriverName = 'PDF Driver' }
      )
    } -ParameterFilter { $ClassName -eq 'Win32_Printer' }

    $result = @(Get-DefaultPrintDevice)

    $result.Count | Should -Be 1
    $result[0].Name | Should -Be 'Office'
    $result[0].Default | Should -BeTrue
    $result[0].Source | Should -Be 'CIM'
  }
}

Describe 'Set-DefaultPrintDevice' {
  It 'sets the default printer by exact name' {
    Mock Get-CimInstance {
      @([PSCustomObject]@{ Name = 'Office'; DeviceID = 'Office'; Default = $false })
    } -ParameterFilter { $ClassName -eq 'Win32_Printer' }
    Mock Invoke-CimMethod { [PSCustomObject]@{ ReturnValue = 0 } } -ParameterFilter { $MethodName -eq 'SetDefaultPrinter' }

    $result = Set-DefaultPrintDevice -Name 'Office'

    $result.Target | Should -Be 'Office'
    $result.Action | Should -Be 'SetDefault'
    $result.Status | Should -Be 'Completed'
    Should -Invoke Invoke-CimMethod -Times 1 -Exactly
  }

  It 'honors WhatIf' {
    Mock Get-CimInstance {
      @([PSCustomObject]@{ Name = 'Office'; DeviceID = 'Office'; Default = $false })
    } -ParameterFilter { $ClassName -eq 'Win32_Printer' }
    Mock Invoke-CimMethod { [PSCustomObject]@{ ReturnValue = 0 } }

    $result = Set-DefaultPrintDevice -Name 'Office' -WhatIf

    $result.Status | Should -Be 'Skipped'
    $result.Detail | Should -Be 'WhatIf'
    Should -Invoke Invoke-CimMethod -Times 0 -Exactly
  }

  It 'skips when the requested printer is already default' {
    Mock Get-CimInstance {
      @([PSCustomObject]@{ Name = 'Office'; DeviceID = 'Office'; Default = $true })
    } -ParameterFilter { $ClassName -eq 'Win32_Printer' }
    Mock Invoke-CimMethod { [PSCustomObject]@{ ReturnValue = 0 } }

    $result = Set-DefaultPrintDevice -Name 'Office'

    $result.Status | Should -Be 'Skipped'
    $result.Detail | Should -Be 'AlreadyDefault'
    Should -Invoke Invoke-CimMethod -Times 0 -Exactly
  }

  It 'fails when the printer name is missing' {
    Mock Get-CimInstance { @() } -ParameterFilter { $ClassName -eq 'Win32_Printer' }

    { Set-DefaultPrintDevice -Name 'Missing' } | Should -Throw
  }

  It 'fails when exact name resolution returns more than one printer' {
    Mock Get-CimInstance {
      @(
        [PSCustomObject]@{ Name = 'Office'; DeviceID = 'Office1'; Default = $false }
        [PSCustomObject]@{ Name = 'Office'; DeviceID = 'Office2'; Default = $false }
      )
    } -ParameterFilter { $ClassName -eq 'Win32_Printer' }

    { Set-DefaultPrintDevice -Name 'Office' } | Should -Throw
  }
}

Describe 'Get-ScanDevice' {
  It 'normalizes WIA device info objects' {
    $deviceInfo = [PSCustomObject]@{
      Properties = @(
        [PSCustomObject]@{ Name = 'Name'; Value = 'Brother MFC' }
        [PSCustomObject]@{ Name = 'Unique Device ID'; Value = 'wia-device-1' }
        [PSCustomObject]@{ Name = 'Manufacturer'; Value = 'Brother' }
        [PSCustomObject]@{ Name = 'Type'; Value = 65537 }
        [PSCustomObject]@{ Name = 'Port'; Value = 'AUTO' }
        [PSCustomObject]@{ Name = 'Server'; Value = 'local' }
        [PSCustomObject]@{ Name = 'Driver Version'; Value = '1.2.3' }
        [PSCustomObject]@{ Name = 'WIA Version'; Value = '2.0' }
        [PSCustomObject]@{ Name = 'PnP ID String'; Value = 'root#image#0000' }
      )
    }

    $result = ConvertTo-ScanDevice -DeviceInfo $deviceInfo

    $result.Name | Should -Be 'Brother MFC'
    $result.DeviceId | Should -Be 'wia-device-1'
    $result.Manufacturer | Should -Be 'Brother'
    $result.Port | Should -Be 'AUTO'
    $result.WiaVersion | Should -Be '2.0'
    $result.Source | Should -Be 'WIA'
  }

  It 'does not throw when WIA is unavailable' {
    Mock New-Object { throw 'WIA unavailable' } -ParameterFilter { $ComObject -eq 'WIA.DeviceManager' }

    { Get-ScanDevice } | Should -Not -Throw
    @(Get-ScanDevice).Count | Should -Be 0
  }
}
