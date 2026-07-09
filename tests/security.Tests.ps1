#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }


BeforeAll {
  param()
  . $PSScriptRoot/../lib/common.ps1
  . $PSScriptRoot/../lib/security.ps1

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Config variable is used across nested Describe blocks.')]
  $config = Import-SecurityEventConfiguration -Path (Join-Path $PSScriptRoot '../lib/security.psd1') -Force
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

Describe 'Import-SecurityEventConfiguration' {
  It 'imports the PSD1 configuration successfully' {
    $imported = Import-SecurityEventConfiguration -Force
    $imported | Should -Not -BeNullOrEmpty
    $imported.Keys | Should -Contain 'FieldMaps'
    $imported.Keys | Should -Contain 'Groups'
    $imported.Keys | Should -Contain 'Events'
  }

  It 'returns cached config on second call without Force' {
    $null = Import-SecurityEventConfiguration -Force
    $second = Import-SecurityEventConfiguration
    $second | Should -Not -BeNullOrEmpty
  }
}

Describe 'FieldMaps' {
  It 'contains expected LogonType entries' {
    $config.FieldMaps.LogonType[2].Name | Should -Be 'Interactive'
    $config.FieldMaps.LogonType[3].Name | Should -Be 'Network'
    $config.FieldMaps.LogonType[7].Name | Should -Be 'Unlock'
    $config.FieldMaps.LogonType[10].Name | Should -Be 'RemoteInteractive'
    $config.FieldMaps.LogonType[11].Name | Should -Be 'CachedInteractive'
  }

  It 'contains ImpersonationLevel entries' {
    $config.FieldMaps.ImpersonationLevel['%%1834'].Name | Should -Be 'Impersonation'
    $config.FieldMaps.ImpersonationLevel['%%1835'].Name | Should -Be 'Delegation'
  }
}

Describe 'Get-SecurityEventGroup' {
  It 'retrieves the Logon group with expected event IDs' {
    $group = Get-SecurityEventGroup -Name 'Logon' -Configuration $config
    $group.Name | Should -Be 'Logon'
    $group.EventIds | Should -Contain 4624
    $group.EventIds | Should -Contain 4625
    $group.DefaultSuppressions.TargetUserName | Should -Contain 'ANONYMOUS LOGON'
  }

  It 'throws for unknown group names' {
    { Get-SecurityEventGroup -Name 'NonExistentGroup' -Configuration $config } | Should -Throw
  }
}

Describe 'Get-SecurityEventDefinition' {
  It 'finds 4624 by ID' {
    $defs = @(Get-SecurityEventDefinition -Id 4624 -Configuration $config)
    $defs.Count | Should -BeGreaterThan 0
    $defs[0].Name | Should -Be 'SuccessfulLogon'
    $defs[0].UtilityGroup | Should -Be 'Logon'
  }

  It 'finds all Logon group events' {
    $defs = @(Get-SecurityEventDefinition -Group 'Logon' -Configuration $config)
    $defs.Count | Should -BeGreaterThan 10
  }

  It 'finds events by LogName and Id' {
    $defs = @(Get-SecurityEventDefinition -LogName 'System' -Id 7045 -Configuration $config)
    $defs.Count | Should -Be 1
    $defs[0].Name | Should -Be 'ServiceInstalled'
  }
}

Describe 'Resolve-WindowsEventMappedField' {
  It 'resolves LogonType 10 to RemoteInteractive' {
    $result = Resolve-WindowsEventMappedField -MapName 'LogonType' -Value 10 -Configuration $config
    $result.Name | Should -Be 'RemoteInteractive'
  }

  It 'resolves LogonType 2 to Interactive' {
    $result = Resolve-WindowsEventMappedField -MapName 'LogonType' -Value 2 -Configuration $config
    $result.Name | Should -Be 'Interactive'
  }

  It 'returns null for unknown map name' {
    $result = Resolve-WindowsEventMappedField -MapName 'NonExistent' -Value 5 -Configuration $config
    $result | Should -Be $null
  }

  It 'returns null for unmapped value' {
    $result = Resolve-WindowsEventMappedField -MapName 'LogonType' -Value 999 -Configuration $config
    $result | Should -Be $null
  }

  It 'resolves ImpersonationLevel %%1834' {
    $result = Resolve-WindowsEventMappedField -MapName 'ImpersonationLevel' -Value '%%1834' -Configuration $config
    $result.Name | Should -Be 'Impersonation'
  }
}

Describe 'Test-WindowsEventLogChannel' {
  It 'returns false for a non-existent channel' {
    $result = Test-WindowsEventLogChannel -LogName 'Fake-Channel-That-DoesNotExist'
    $result | Should -Be $false
  }
}

Describe 'ConvertFrom-WinEvent' {
  It 'parses a synthetic event object with ToXml successfully' {
    $xml = [xml]@'
<Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
  <System>
    <Provider Name="Microsoft-Windows-Security-Auditing" Guid="{...}" />
    <EventID>4624</EventID>
    <Level>0</Level>
    <Task>12544</Task>
    <Keywords>0x8020000000000000</Keywords>
    <TimeCreated SystemTime="2026-07-10T10:00:00.000000000Z" />
    <Computer>TEST-HOST</Computer>
  </System>
  <EventData>
    <Data Name="SubjectUserSid">S-1-5-18</Data>
    <Data Name="SubjectUserName">SYSTEM</Data>
    <Data Name="SubjectDomainName">NT AUTHORITY</Data>
    <Data Name="TargetUserSid">S-1-5-21-999</Data>
    <Data Name="TargetUserName">jdoe</Data>
    <Data Name="TargetDomainName">CONTOSO</Data>
    <Data Name="TargetLogonId">0xabc123</Data>
    <Data Name="LogonType">10</Data>
    <Data Name="IpAddress">192.168.1.100</Data>
  </EventData>
</Event>
'@

    $mockEvent = [PSCustomObject]@{
      TimeCreated = [datetime]'2026-07-10T10:00:00'
      Id = 4624
      ProviderName = 'Microsoft-Windows-Security-Auditing'
      LogName = 'Security'
      MachineName = 'TEST-HOST'
      RecordId = 12345
      LevelDisplayName = 'Information'
    }
    $mockEvent | Add-Member -MemberType ScriptMethod -Name 'ToXml' -Value { $xml.OuterXml } -Force

    $result = $mockEvent | ConvertFrom-WinEvent -Configuration $config

    $result.Id | Should -Be 4624
    $result.TargetUserName | Should -Be 'jdoe'
    $result.IpAddress | Should -Be '192.168.1.100'
    $result.LogonType | Should -Be 10
    $result.LogonTypeName | Should -Be 'RemoteInteractive'
    $result.RawData | Should -Not -BeNullOrEmpty
    $result.RawData['TargetUserName'] | Should -Be 'jdoe'
  }

  It 'handles event without LogonType field gracefully' {
    $xml = [xml]@'
<Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
  <System>
    <Provider Name="Service Control Manager" />
    <EventID>7036</EventID>
    <TimeCreated SystemTime="2026-07-10T10:00:00.000000000Z" />
    <Computer>TEST-HOST</Computer>
  </System>
  <EventData>
    <Data Name="param1">BITS</Data>
    <Data Name="param2">running</Data>
  </EventData>
</Event>
'@

    $mockEvent = [PSCustomObject]@{
      TimeCreated = [datetime]'2026-07-10T10:00:00'
      Id = 7036
      ProviderName = 'Service Control Manager'
      LogName = 'System'
      MachineName = 'TEST-HOST'
      RecordId = 1
      LevelDisplayName = 'Information'
    }
    $mockEvent | Add-Member -MemberType ScriptMethod -Name 'ToXml' -Value { $xml.OuterXml } -Force

    $result = $mockEvent | ConvertFrom-WinEvent -Configuration $config

    $result.Id | Should -Be 7036
    $result.LogonType | Should -Be $null
    $result.LogonTypeName | Should -Be $null
    $result.RawData['param1'] | Should -Be 'BITS'
  }
}

Describe 'Get-WindowsEventByDefinition' {
  It 'throws when called with no definitions and no Group/Id' {
    { Get-WindowsEventByDefinition -Configuration $config -ErrorAction Stop } | Should -Throw
  }
}

Describe 'Get-WindowsSysmonEvent' {
  It 'returns silently when Sysmon channel is missing' {
    function Test-WindowsEventLogChannel { param([string]$LogName) $null = $LogName; return $false }
    $result = Get-WindowsSysmonEvent -Configuration $config -ErrorAction SilentlyContinue
    $result | Should -BeNullOrEmpty
  }
}

Describe 'Get-WindowsLogonEvent - integration' {
  It 'filters events by Id parameter' {
    Mock -CommandName 'Get-WinEvent' -MockWith { @() }

    $null = Get-WindowsLogonEvent -Id 4625 -Configuration $config -ErrorAction SilentlyContinue

    Should -Invoke -CommandName 'Get-WinEvent' -Times 1 -Exactly
  }
}

Describe 'Get-WindowsLogonEvent - system account suppression' {
  It 'suppresses SYSTEM when IncludeSystem is not set' {
    $xml = [xml]@'
<Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
  <System><Provider Name="Microsoft-Windows-Security-Auditing" /><EventID>4624</EventID><TimeCreated SystemTime="2026-07-10T10:00:00.000000000Z" /><Computer>TEST-HOST</Computer></System>
  <EventData><Data Name="TargetUserName">SYSTEM</Data><Data Name="LogonType">5</Data></EventData>
</Event>
'@
    $mockEvent = [PSCustomObject]@{
      TimeCreated = [datetime]'2026-07-10T10:00:00'
      Id = 4624
      ProviderName = 'Microsoft-Windows-Security-Auditing'
      LogName = 'Security'
      MachineName = 'TEST-HOST'
      RecordId = 1
      LevelDisplayName = 'Information'
    }
    $mockEvent | Add-Member -MemberType ScriptMethod -Name 'ToXml' -Value { $xml.OuterXml } -Force

    Mock -CommandName 'Get-WinEvent' -MockWith { @($mockEvent) }

    $result = Get-WindowsLogonEvent -Configuration $config -ErrorAction SilentlyContinue
    $result | Should -BeNullOrEmpty
  }

  It 'includes SYSTEM when -IncludeSystem is supplied' {
    $xml = [xml]@'
<Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event">
  <System><Provider Name="Microsoft-Windows-Security-Auditing" /><EventID>4624</EventID><TimeCreated SystemTime="2026-07-10T10:00:00.000000000Z" /><Computer>TEST-HOST</Computer></System>
  <EventData><Data Name="TargetUserName">SYSTEM</Data><Data Name="LogonType">5</Data></EventData>
</Event>
'@
    $mockEvent = [PSCustomObject]@{
      TimeCreated = [datetime]'2026-07-10T10:00:00'
      Id = 4624
      ProviderName = 'Microsoft-Windows-Security-Auditing'
      LogName = 'Security'
      MachineName = 'TEST-HOST'
      RecordId = 1
      LevelDisplayName = 'Information'
    }
    $mockEvent | Add-Member -MemberType ScriptMethod -Name 'ToXml' -Value { $xml.OuterXml } -Force

    Mock -CommandName 'Get-WinEvent' -MockWith { @($mockEvent) }

    $result = Get-WindowsLogonEvent -IncludeSystem -Configuration $config -ErrorAction SilentlyContinue
    $result | Should -Not -BeNullOrEmpty
    $result.TargetUserName | Should -Be 'SYSTEM'
  }
}
