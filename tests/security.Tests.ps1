#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../lib/common.ps1
  . $PSScriptRoot/../lib/security.ps1
}

Describe 'Add-DefenderExclusion' {
  It 'adds a Defender path exclusion' {
    Mock Add-MpPreference { }

    $result = Add-DefenderExclusion -Type Path -Value 'C:\Program Files\JetBrains'

    $result.Status | Should -Be 'Completed'
    $result.Source | Should -Be 'Defender'
    Should -Invoke Add-MpPreference -Times 1 -Exactly -ParameterFilter {
      $ExclusionPath -eq 'C:\Program Files\JetBrains'
    }
  }

  It 'returns a failed result when Add-MpPreference fails' {
    Mock Add-MpPreference { throw 'denied' }

    $result = Add-DefenderExclusion -Type Process -Value 'wsl.exe'

    $result.Status | Should -Be 'Failed'
    $result.Detail | Should -Match 'denied'
  }

  It 'honors WhatIf without calling Add-MpPreference' {
    Mock Add-MpPreference { }

    $result = Add-DefenderExclusion -Type Extension -Value 'vhdx' -WhatIf

    $result.Status | Should -Be 'Skipped'
    $result.Detail | Should -Be 'WhatIf'
    Should -Invoke Add-MpPreference -Times 0 -Exactly
  }
}

Describe 'Enable-WSLFirewallRule' {
  It 'creates the WSL inbound firewall rule when missing' {
    Mock Get-NetFirewallRule { @() }
    Mock New-NetFirewallRule { [PSCustomObject]@{ DisplayName = 'WSL' } }

    $result = Enable-WSLFirewallRule

    $result.Status | Should -Be 'Completed'
    Should -Invoke New-NetFirewallRule -Times 1 -Exactly -ParameterFilter {
      $DisplayName -eq 'WSL' -and $Direction -eq 'Inbound' -and $InterfaceAlias -eq 'vEthernet (WSL)' -and $Action -eq 'Allow'
    }
  }

  It 'skips when the WSL rule already exists' {
    Mock Get-NetFirewallRule { @([PSCustomObject]@{ DisplayName = 'WSL' }) }
    Mock New-NetFirewallRule { }

    $result = Enable-WSLFirewallRule

    $result.Status | Should -Be 'Skipped'
    $result.Detail | Should -Be 'AlreadyExists'
    Should -Invoke New-NetFirewallRule -Times 0 -Exactly
  }
}

Describe 'Disable-JetBrainsFirewallRule' {
  It 'disables matching public-profile JetBrains firewall rules' {
    $rule = [PSCustomObject]@{ DisplayName = 'WebStorm 2026.1' }
    Mock Get-NetFirewallProfile { [PSCustomObject]@{ Name = 'Public' } }
    Mock Get-NetFirewallRule { @($rule) }
    Mock Disable-NetFirewallRule { }

    $result = @(Disable-JetBrainsFirewallRule -Prefix @('WebStorm'))

    $result.Count | Should -Be 1
    $result[0].Status | Should -Be 'Completed'
    Should -Invoke Disable-NetFirewallRule -Times 1 -Exactly
  }

  It 'returns skipped results during WhatIf without querying firewall rules' {
    Mock Get-NetFirewallProfile { throw 'should not query' }
    Mock Get-NetFirewallRule { throw 'should not query' }
    Mock Disable-NetFirewallRule { }

    $result = @(Disable-JetBrainsFirewallRule -Prefix @('WebStorm', 'Rider') -WhatIf)

    $result.Count | Should -Be 2
    $result[0].Status | Should -Be 'Skipped'
    $result[0].Detail | Should -Be 'WhatIf'
    Should -Invoke Get-NetFirewallProfile -Times 0 -Exactly
    Should -Invoke Get-NetFirewallRule -Times 0 -Exactly
  }
}
