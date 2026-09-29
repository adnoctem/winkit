#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'COM test double crosses script scope and is removed after each test.')]
param ()

BeforeAll {
  Import-Module PSFoundation -Force
  $script:CheckpointScript = Join-Path $PSScriptRoot '../../scripts/Office/Checkpoint-Outlook.ps1'
}

Describe 'Outlook user checkpoints' {
  BeforeEach {
    $script:SavedAppData = $env:APPDATA
    $script:SavedLocalAppData = $env:LOCALAPPDATA
    $env:APPDATA = Join-Path $TestDrive 'Roaming'
    $env:LOCALAPPDATA = Join-Path $TestDrive 'Local'
    $script:Destination = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $script:Source = Join-Path $TestDrive 'mail.pst'
    [IO.File]::WriteAllText($script:Source, 'Synthetic mail data')
    $signature = Join-Path $env:APPDATA 'Microsoft\Signatures'
    $null = [IO.Directory]::CreateDirectory($signature)
    [IO.File]::WriteAllText((Join-Path $signature 'signature.htm'), '<p>Test signature</p>')

    Mock Import-Module { }
    Mock Write-Log { }
    Mock Write-Progress { }
    Mock Remove-ComObject { }
    Mock Invoke-ComGarbageCollection { }
    Mock Get-UserInfo { @{ UserName = 'TEST\MailUser'; SID = 'S-1-5-21-1000'; IsAdministrator = $false } }
    Mock Get-Process { }
    Mock Get-Process { [PSCustomObject]@{ SessionId = 4 } } -ParameterFilter { $Id -eq $PID }
    Mock Get-OfficeInventory { [PSCustomObject]@{ Products = @([PSCustomObject]@{ ProductId = 'Standard2019Volume' }); Msi = @(); Unknowns = @() } }
    Mock Get-OfficeActivationStatus { [PSCustomObject]@{ TargetProductId = $TargetProductId; Status = 'Licensed' } }
    Mock Test-RegistryPath { $Path -eq 'HKCU\Software\Microsoft\Office\16.0' }
    Mock Invoke-SafeProcess {
      [IO.File]::WriteAllText($ArgumentList[2], 'Windows Registry Editor Version 5.00')
      [PSCustomObject]@{ ExitCode = 0; TimedOut = $false; Cancelled = $false }
    }
    Mock Connect-Outlook { throw 'Explicit StorePaths must not open Outlook.' }
  }

  AfterEach {
    $env:APPDATA = $script:SavedAppData
    $env:LOCALAPPDATA = $script:SavedLocalAppData
    Remove-Variable -Name WinkitCheckpointContext -Scope Global -ErrorAction SilentlyContinue
  }

  It 'captures data, user settings and registry exports with hashes and restore paths' {
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    $result.Copied | Should -Be 2
    $manifest = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest.Status | Should -Be Completed
    $manifest.User | Should -Be 'TEST\MailUser'
    $manifest.Activation[0].Status | Should -Be Licensed
    $manifest.Hashed | Should -BeTrue
    $copies = @($manifest.Results | Where-Object { $_.Status -eq 'Completed' -and $_.Category -ne 'Registry' })
    foreach ($copy in $copies) {
      $copy.SourceSHA256 | Should -Be (Get-FileHash -LiteralPath $copy.OriginalPath).Hash
      $copy.BackupSHA256 | Should -Be (Get-FileHash -LiteralPath $copy.Target).Hash
      (Join-Path $result.CheckpointDirectory $copy.RelativePath) | Should -Be $copy.Target
    }
    $registry = @($manifest.Results | Where-Object Category -EQ Registry)
    $registry.Count | Should -Be 1
    $registry[0].BackupSHA256 | Should -Be (Get-FileHash -LiteralPath $registry[0].Target).Hash
    Should -Invoke Connect-Outlook -Times 0
    Should -Invoke Invoke-SafeProcess -Times 1
  }

  It 'previews with <Option> without writing or exporting' -ForEach @(
    @{ Option = 'DryRun' }
    @{ Option = 'WhatIf' }
  ) {
    $arguments = @{ Destination = $script:Destination; StorePaths = $script:Source }
    $arguments[$Option] = $true
    $result = & $script:CheckpointScript @arguments
    $result.Status | Should -Be Preview
    $result.ReportPath | Should -BeNullOrEmpty
    Test-Path -LiteralPath $script:Destination | Should -BeFalse
    Should -Invoke Invoke-SafeProcess -Times 0
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'rejects missing selected files before creating checkpoint artifacts' {
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths (Join-Path $TestDrive 'missing.pst') -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    Test-Path -LiteralPath $script:Destination | Should -BeFalse
    Should -Invoke Invoke-SafeProcess -Times 0
  }

  It 'locks every selected source before copying any file' {
    $locked = Join-Path $TestDrive 'locked.pst'
    [IO.File]::WriteAllText($locked, 'Locked file')
    $stream = [IO.File]::Open($locked, 'Open', 'ReadWrite', 'None')
    try {
      $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source, $locked -PassThru -Confirm:$false -WarningAction SilentlyContinue
      $result.Status | Should -Be Failed
      $result.Copied | Should -Be 0
      Test-Path -LiteralPath $script:Destination | Should -BeFalse
    }
    finally {
      $stream.Dispose()
    }
    # The earlier source handle must also have been released after the failure.
    $stream = [IO.File]::Open($script:Source, 'Open', 'ReadWrite', 'None')
    $stream.Dispose()
  }

  It 'records incomplete capture when a registry export fails' {
    Mock Invoke-SafeProcess { throw 'Synthetic registry export failure' }
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $manifest = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest.Status | Should -Be Failed
    $manifest.Failed | Should -Be 1
    $manifest.Copied | Should -Be 2
    $manifest.Results[-1].Detail | Should -Match 'registry export failure'
  }

  It 'rejects a nonzero native export even if it left a nonempty output file' {
    Mock Invoke-SafeProcess {
      [IO.File]::WriteAllText($ArgumentList[2], 'Incomplete registry output')
      [PSCustomObject]@{ ExitCode = 1; TimedOut = $false; Cancelled = $false }
    }
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Results[-1].Detail | Should -Match 'exit code 1'
  }

  It 'fails unreadable settings enumeration instead of silently omitting files' {
    Mock Get-ChildItem { throw 'Synthetic access denied' } -ParameterFilter { $LiteralPath -like '*Signatures' }
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    Test-Path -LiteralPath $script:Destination | Should -BeFalse
  }

  It 'distinguishes length-only copies from hash verification' {
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -SkipHash -PassThru -Confirm:$false
    $manifest = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest.Hashed | Should -BeFalse
    $store = $manifest.Results | Where-Object Category -EQ Store
    $store.Verification | Should -Be LengthOnly
    $store.SourceSHA256 | Should -BeNullOrEmpty
  }

  It 'excludes OSTs from both explicit selections and settings folders' {
    $ost = Join-Path $TestDrive 'mail.ost'
    [IO.File]::WriteAllText($ost, 'OST cache')
    $outlook = Join-Path $env:APPDATA 'Microsoft\Outlook'
    $null = [IO.Directory]::CreateDirectory($outlook)
    [IO.File]::WriteAllText((Join-Path $outlook 'other.ost'), 'Another cache')
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source, $ost -ExcludeOst -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    @(Get-ChildItem -LiteralPath $result.CheckpointDirectory -Recurse -Filter '*.ost').Count | Should -Be 0
    @($result.Results | Where-Object Detail -EQ 'OST excluded by request.').Count | Should -Be 2
  }

  It 'retains OSTs as explicitly selected cache copies without ExcludeOst' {
    $ost = Join-Path $TestDrive 'mail.ost'
    [IO.File]::WriteAllText($ost, 'OST cache')
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $ost -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    @($result.Results | Where-Object Category -EQ Store).Count | Should -Be 1
  }

  It 'rejects a destination within the captured settings tree' {
    $nested = Join-Path $env:APPDATA 'Microsoft\Signatures\Backup'
    $result = & $script:CheckpointScript -Destination $nested -StorePaths $script:Source -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    Test-Path -LiteralPath $nested | Should -BeFalse
  }

  It 'rejects elevation by default and service identities even with an override' {
    Mock Get-UserInfo { @{ UserName = 'TEST\Admin'; SID = 'S-1-5-21-1000'; IsAdministrator = $true } }
    { & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -DryRun } | Should -Throw '*non-elevated*'
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -IgnoreAdministrator -DryRun -WarningAction SilentlyContinue
    $result.Status | Should -Be Preview
    Mock Get-UserInfo { @{ UserName = 'SYSTEM'; SID = 'S-1-5-18'; IsAdministrator = $true } }
    { & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -IgnoreAdministrator -DryRun } | Should -Throw '*service identity*'
  }

  It 'blocks capture while another Office application runs in this session' {
    Mock Get-Process { [PSCustomObject]@{ SessionId = 4; Name = 'WINWORD' } } -ParameterFilter { $Name -contains 'WINWORD' }
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    Test-Path -LiteralPath $script:Destination | Should -BeFalse
  }

  It 'uses distinct runs and distinct names for same-named source files' {
    $other = Join-Path $TestDrive 'other'
    $null = [IO.Directory]::CreateDirectory($other)
    $second = Join-Path $other 'mail.pst'
    [IO.File]::WriteAllText($second, 'Second mailbox')
    $first = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source, $second, $script:Source -PassThru -Confirm:$false
    $again = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false
    $first.CheckpointDirectory | Should -Not -Be $again.CheckpointDirectory
    @($first.Results | Where-Object Category -EQ Store).Count | Should -Be 2
    @($first.Results | Where-Object Category -EQ Store | Select-Object -ExpandProperty Target -Unique).Count | Should -Be 2
  }

  It 'records activation as unknown without losing a usable user-data checkpoint' {
    Mock Get-OfficeActivationStatus { throw 'Synthetic license query failure' }
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Completed
    $result.Warnings.Count | Should -Be 1
    $manifest = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest.Activation[0].Status | Should -Be Unknown
  }

  It 'uses Outlook discovery for the profile and never quits it during preview' {
    $store = [PSCustomObject]@{ DisplayName = 'Mailbox'; FilePath = $script:Source }
    $stores = [PSCustomObject]@{ Values = @($store); Count = 1 }
    $stores | Add-Member ScriptMethod Item { param($Index) $this.Values[$Index - 1] }
    $app = [PSCustomObject]@{
      Version      = '16.0'
      QuitCalled   = $false
      ShutdownFile = (Join-Path $env:APPDATA 'Microsoft\Signatures\saved-on-exit.htm')
    }
    $app | Add-Member ScriptMethod Quit {
      $this.QuitCalled = $true
      [IO.File]::WriteAllText($this.ShutdownFile, 'Settings saved during shutdown')
    }
    $global:WinkitCheckpointContext = [PSCustomObject]@{
      App       = $app
      Namespace = [PSCustomObject]@{ Stores = $stores }
    }
    Mock Connect-Outlook { $global:WinkitCheckpointContext }
    $preview = & $script:CheckpointScript -Destination $script:Destination -QuitOutlook -DryRun
    $preview.Status | Should -Be Preview
    $global:WinkitCheckpointContext.App.QuitCalled | Should -BeFalse
    Test-Path -LiteralPath $global:WinkitCheckpointContext.App.ShutdownFile | Should -BeFalse
    $result = & $script:CheckpointScript -Destination $script:Destination -QuitOutlook -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    $global:WinkitCheckpointContext.App.QuitCalled | Should -BeTrue
    @($result.Results | Where-Object { $_.OriginalPath -eq $global:WinkitCheckpointContext.App.ShutdownFile -and $_.Status -eq 'Completed' }).Count | Should -Be 1
  }

  It 'does not broaden explicit store selection through settings-directory discovery' {
    $outlook = Join-Path $env:APPDATA 'Microsoft\Outlook'
    $null = [IO.Directory]::CreateDirectory($outlook)
    $detached = Join-Path $outlook 'detached.pst'
    [IO.File]::WriteAllText($detached, 'Unselected mail')
    $result = & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    @($result.Results | Where-Object Category -EQ Store).Count | Should -Be 1
    @($result.Results | Where-Object { $_.Target -eq $detached -and $_.Status -eq 'Skipped' }).Count | Should -Be 1
  }

  It 'rejects discovery-only options alongside explicit StorePaths' {
    { & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -QuitOutlook -DryRun } | Should -Throw
    { & $script:CheckpointScript -Destination $script:Destination -StorePaths $script:Source -WaitSeconds 0 -DryRun } | Should -Throw
  }
}
