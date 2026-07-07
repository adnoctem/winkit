#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../lib/user.ps1
}

Describe 'Get-UserInfo' {
  It 'returns a hashtable' {
    $result = Get-UserInfo
    $result | Should -BeOfType [hashtable]
  }

  It 'contains a non-empty UserName' {
    $result = Get-UserInfo
    $result.UserName | Should -Not -BeNullOrEmpty
  }

  It 'contains a boolean IsAdministrator' {
    $result = Get-UserInfo
    $result.IsAdministrator | Should -BeOfType [bool]
  }

  It 'contains a non-empty SID' {
    $result = Get-UserInfo
    $result.SID | Should -Not -BeNullOrEmpty
  }

  It 'SID starts with S-1-' {
    $result = Get-UserInfo
    $result.SID | Should -Match '^S-1-'
  }
}

Describe 'Get-UserSID' {
  It 'returns a SID string for the current user' {
    $username = (Get-UserInfo).UserName
    $result = Get-UserSID -UserName $username
    $result | Should -Match '^S-1-'
  }

  It 'returns null for a non-existent user' {
    $result = Get-UserSID -UserName 'NONEXISTENT_USER_12345' -ErrorAction SilentlyContinue
    $result | Should -BeNullOrEmpty
  }
}
