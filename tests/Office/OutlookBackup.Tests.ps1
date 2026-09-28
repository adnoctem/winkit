#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Mock context crosses script scope and is removed after each test.')]
param ()

BeforeAll {
  Import-Module PSFoundation -Force
  $script:BackupScript = Join-Path $PSScriptRoot '../../scripts/Office/Backup-OutlookDataFile.ps1'
}

Describe 'Outlook backup profile discovery' {
  BeforeEach {
    Mock Import-Module { }
    Mock Write-Log { }
    Mock Write-Progress { }
    Mock Remove-ComObject { }
    Mock Invoke-ComGarbageCollection { }
    Mock Get-UserInfo { @{ IsAdministrator = $false; UserName = 'TEST\MailUser' } }
    Mock Get-Process { }
    $source = Join-Path $TestDrive 'profile.pst'
    [IO.File]::WriteAllText($source, 'whole mailbox')
    $script:BackupDestination = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $store = [PSCustomObject]@{ StoreID = 'one'; DisplayName = 'Mailbox'; FilePath = $source }
    $stores = [PSCustomObject]@{ Values = @($store); Count = 1 }
    $stores | Add-Member ScriptMethod Item { param($Index) $this.Values[$Index - 1] }
    $app = [PSCustomObject]@{ Version = '12.0'; QuitCalled = $false }
    $app | Add-Member ScriptMethod Quit { $this.QuitCalled = $true }
    $global:WinkitBackupContext = [PSCustomObject]@{
      App       = $app
      Namespace = [PSCustomObject]@{ Stores = $stores; DefaultStore = $store }
    }
    Mock Connect-Outlook { $global:WinkitBackupContext }
  }

  AfterEach {
    Remove-Variable -Name WinkitBackupContext -Scope Global -ErrorAction SilentlyContinue
  }

  It 'discovers the default PST and shuts down only when requested' {
    $result = & $script:BackupScript -Destination $script:BackupDestination -QuitOutlook -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    $result.Results[0].StoreName | Should -Be Mailbox
    $global:WinkitBackupContext.App.QuitCalled | Should -BeTrue
  }

  It 'does not quit Outlook during preview' {
    $result = & $script:BackupScript -Destination $script:BackupDestination -AllStores -QuitOutlook -WhatIf
    $result.Status | Should -Be Preview
    $global:WinkitBackupContext.App.QuitCalled | Should -BeFalse
    Test-Path -LiteralPath $script:BackupDestination | Should -BeFalse
  }

  It 'fails safely when Outlook does not exit' {
    Mock Get-Process { [PSCustomObject]@{ Name = 'OUTLOOK' } }
    $result = & $script:BackupScript -Destination $script:BackupDestination -WaitSeconds 0 -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Results[-1].Detail | Should -Match 'still running'
    Test-Path -LiteralPath $script:BackupDestination | Should -BeFalse
    $global:WinkitBackupContext.App.QuitCalled | Should -BeFalse
  }

  It 'rejects ambiguous store names before shutdown or copying' {
    $stores = $global:WinkitBackupContext.Namespace.Stores
    $stores.Values = @($stores.Values[0], $stores.Values[0])
    $stores.Count = 2
    $result = & $script:BackupScript -StoreName Mailbox -Destination $script:BackupDestination -QuitOutlook -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Results[-1].Detail | Should -Match 'matches 2 stores'
    $global:WinkitBackupContext.App.QuitCalled | Should -BeFalse
  }

  It 'blocks elevated discovery before opening Outlook' {
    Mock Get-UserInfo { @{ IsAdministrator = $true; UserName = 'TEST\Administrator' } }
    { & $script:BackupScript -Destination $script:BackupDestination -DryRun } | Should -Throw '*non-elevated*'
    Should -Invoke Connect-Outlook -Times 0
  }
}

Describe 'Closed-file Outlook backups' {
  BeforeEach {
    Mock Import-Module { }
    Mock Connect-Outlook { throw 'Direct file backup must not connect to Outlook' }
    Mock Get-UserInfo { throw 'Direct file backup must not inspect Outlook identity' }
    Mock Write-Log { }
    Mock Write-Progress { }
    $source = Join-Path $TestDrive 'source.pst'
    $script:BackupDestination = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    [IO.File]::WriteAllBytes($source, [byte[]](0..255))
  }

  It 'copies exact bytes and records matching hashes without Outlook' {
    $result = & $script:BackupScript -PSTPath $source -Destination $script:BackupDestination -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    $result.Copied | Should -Be 1
    $manifest = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $entry = $manifest.Results[0]
    $entry.SourceSHA256 | Should -Be (Get-FileHash -LiteralPath $source).Hash
    $entry.BackupSHA256 | Should -Be (Get-FileHash -LiteralPath $entry.Target).Hash
    $entry.OriginalPath | Should -Be $source
    $entry.Bytes | Should -Be 256
    Should -Invoke Connect-Outlook -Times 0
    Should -Invoke Get-UserInfo -Times 0
  }

  It 'previews without creating the destination or opening Outlook' {
    $result = & $script:BackupScript -PSTPath $source -Destination $script:BackupDestination -DryRun
    $result.Status | Should -Be Preview
    $result.Results[0].Detail | Should -Be DryRun
    Test-Path -LiteralPath $script:BackupDestination | Should -BeFalse
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'rejects every Outlook-specific option in direct file mode' {
    foreach ($option in @('StoreName', 'AllStores', 'QuitOutlook', 'IgnoreAdministrator', 'WaitSeconds')) {
      $arguments = @{
        PSTPath     = $source
        Destination = $script:BackupDestination
        DryRun      = $true
      }
      $arguments[$option] = switch ($option) {
        StoreName { 'Mailbox' }
        WaitSeconds { 1 }
        default { $true }
      }
      { & $script:BackupScript @arguments } | Should -Throw
    }
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'rejects a locked PST before creating backup artifacts' {
    $lock = [IO.File]::Open($source, 'Open', 'ReadWrite', 'None')
    try {
      $result = & $script:BackupScript -PSTPath $source -Destination $script:BackupDestination -PassThru -Confirm:$false -WarningAction SilentlyContinue
      $result.Status | Should -Be Failed
      $result.Copied | Should -Be 0
      Test-Path -LiteralPath $script:BackupDestination | Should -BeFalse
    }
    finally {
      $lock.Dispose()
    }
  }

  It 'rejects OST caches and preserves existing backups on repeated runs' {
    $ost = Join-Path $TestDrive 'mail.ost'
    [IO.File]::WriteAllText($ost, 'cache')
    $invalid = & $script:BackupScript -PSTPath $ost -Destination $script:BackupDestination -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $invalid.Status | Should -Be Failed
    Test-Path -LiteralPath $script:BackupDestination | Should -BeFalse
    $first = & $script:BackupScript -PSTPath $source -Destination $script:BackupDestination -PassThru -Confirm:$false
    $second = & $script:BackupScript -PSTPath $source -Destination $script:BackupDestination -PassThru -Confirm:$false
    $first.BackupDirectory | Should -Not -Be $second.BackupDirectory
    Test-Path -LiteralPath $first.Results[0].Target | Should -BeTrue
  }

  It 'handles duplicate filenames and deduplicates repeated source paths' {
    $other = Join-Path $TestDrive 'other'
    $null = [IO.Directory]::CreateDirectory($other)
    $second = Join-Path $other 'source.pst'
    [IO.File]::WriteAllText($second, 'second file')
    $result = & $script:BackupScript -PSTPath $source, $second, $source -Destination $script:BackupDestination -PassThru -Confirm:$false
    $result.Copied | Should -Be 2
    $result.Results[0].Target | Should -Not -Be $result.Results[1].Target
    @($result.Results | Where-Object { $_.SourceSHA256 -eq $_.BackupSHA256 }).Count | Should -Be 2
  }
}
