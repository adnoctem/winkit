#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Outlook mocks run inside the invoked script; the globals are removed after each test.')]
param ()

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.4.1 -Force
  $script:ScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../scripts/Office/Show-OutlookAccount.ps1'))
  $script:Ast = [Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$null, [ref]$null)
  $definitions = @($script:Ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false))
  foreach ($definition in $definitions) {
    . ([scriptblock]::Create($definition.Extent.Text))
  }

  # A store entry ID in the shape Outlook 2007 records for a PST: wrapped-store marker, mspst.dll, then the path.
  $script:EntryIdHex = '0000000038a1bb1005e5101aa1bb08002b2a56c200006d737073742e646c6c00000000'
  $script:EntryIdBytes = [byte[]]($script:EntryIdHex -split '(..)' | Where-Object { $_ } | ForEach-Object { [Convert]::ToByte($_, 16) })
}

Describe 'Show-OutlookAccount profile evidence' {
  BeforeEach {
    $script:RegistryRoot = 'Software\winkit-tests\' + [guid]::NewGuid().ToString('N')
    $accountKey = "HKCU:\$($script:RegistryRoot)\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000003"
    $null = New-Item -Path $accountKey -Force
    $null = New-ItemProperty -Path $accountKey -Name 'Account Name' -PropertyType Binary -Value ([Text.Encoding]::Unicode.GetBytes("user@example.com`0"))
    $null = New-ItemProperty -Path $accountKey -Name 'Delivery Store EntryID' -PropertyType Binary -Value $script:EntryIdBytes
    $null = New-ItemProperty -Path $accountKey -Name 'POP3 Password' -PropertyType Binary -Value ([byte[]](1..40))
    $null = New-ItemProperty -Path $accountKey -Name 'Leave on Server' -PropertyType DWord -Value 1
    $null = New-Item -Path "HKCU:\$($script:RegistryRoot)\Microsoft\Office\12.0\Outlook" -Force
    $null = New-ItemProperty -Path "HKCU:\$($script:RegistryRoot)\Microsoft\Office\12.0\Outlook" -Name 'DefaultProfile' -Value 'Outlook'
  }

  AfterEach {
    Remove-Item -LiteralPath "HKCU:\$($script:RegistryRoot)" -Recurse -Force -ErrorAction SilentlyContinue
  }

  AfterAll {
    if (@(Get-ChildItem -LiteralPath 'HKCU:\Software\winkit-tests' -ErrorAction SilentlyContinue).Count -eq 0) {
      Remove-Item -LiteralPath 'HKCU:\Software\winkit-tests' -Force -ErrorAction SilentlyContinue
    }
  }

  It 'matches each recorded delivery store to the data file with the same entry ID' {
    $stores = @(
      [PSCustomObject]@{ DisplayName = 'Postfach'; StoreID = $script:EntryIdHex.ToUpperInvariant(); FilePath = 'C:\Users\User1\outlook.pst' }
      [PSCustomObject]@{ DisplayName = 'Archive'; StoreID = '00ff'; FilePath = 'C:\Users\User1\archive.pst' }
    )

    $evidence = Get-OutlookAccountProfileEvidence -Store $stores -RegistryRoot $script:RegistryRoot

    $evidence.DefaultProfiles[0].DefaultProfile | Should -Be 'Outlook'
    @($evidence.Profiles).Count | Should -Be 1
    $evidence.DeliveryStores.Count | Should -Be 1
    $evidence.DeliveryStores[0].AccountName | Should -Be 'user@example.com'
    $evidence.DeliveryStores[0].DeliveryStoreEntryID | Should -Be $script:EntryIdHex
    $evidence.DeliveryStores[0].MatchedStore | Should -Be 'Postfach'
    $evidence.DeliveryStores[0].MatchedFilePath | Should -Be 'C:\Users\User1\outlook.pst'
  }

  It 'withholds credential values and emits unknown names only as name and length' {
    $evidence = Get-OutlookAccountProfileEvidence -Store @() -RegistryRoot $script:RegistryRoot
    $node = $evidence.Profiles[0].Tree.Children[0].Children[0]
    $password = @($node.Values | Where-Object Name -EQ 'POP3 Password')[0]
    $other = @($node.Values | Where-Object Name -EQ 'Leave on Server')[0]

    $password.Emitted | Should -BeFalse
    $password.Length | Should -BeNullOrEmpty
    $password.Hex | Should -BeNullOrEmpty
    $other.Emitted | Should -BeFalse
    $other.Hex | Should -BeNullOrEmpty
    $other.Text | Should -BeNullOrEmpty
    $evidence.DeliveryStores[0].MatchedStore | Should -BeNullOrEmpty
  }
}

Describe 'Show-OutlookAccount -IncludeProfileEvidence' {
  BeforeEach {
    Mock Import-Module { }
    Mock Get-UserInfo { @{ UserName = 'TEST\MailboxUser'; IsAdministrator = $false } }
    Mock Write-Log { }
    Mock Write-OperationResultLog { }
    Mock Remove-ComObject { }
    Mock Invoke-ComGarbageCollection { }
    Mock Get-RegistryValue { $null }
    Mock Connect-Outlook {
      [PSCustomObject]@{
        App       = [PSCustomObject]@{ Version = '12.0.6785.5000' }
        Namespace = 'synthetic-namespace'
      }
    }
    Mock Get-OutlookAccount { @() }
    Mock Get-OutlookStore { @() }
  }

  It 'adds profile evidence only when requested' {
    $without = & $script:ScriptPath -PassThru -WarningAction SilentlyContinue
    $with = & $script:ScriptPath -PassThru -IncludeProfileEvidence -WarningAction SilentlyContinue

    $without.ProfileEvidence | Should -BeNullOrEmpty
    $with.ProfileEvidence | Should -Not -BeNullOrEmpty
    $with.ProfileEvidence.PSObject.Properties.Name | Should -Contain 'DeliveryStores'
  }
}
