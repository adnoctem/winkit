#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Gated suite: run with .\winkit.ps1 test -Outlook on a machine with Outlook
# installed. A scratch PST store is attached for the duration of the run and
# detached afterwards, so the personal mailbox is never touched.
#
# Test order matters: the dedup assertions count exact store-wide results, so
# they run while only the first fixture folder exists.

$_outlookRegistryPaths = @(
  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
  'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
)

$_outlookInstalled = @($_outlookRegistryPaths | Where-Object { Test-Path -LiteralPath $_ }).Count -gt 0

if ($env:WINKIT_TEST_OUTLOOK -ne '1' -or -not $_outlookInstalled) {
  Describe 'Outlook integration' {
    It 'runs only when Outlook is installed and -Outlook was passed' {
      Set-ItResult -Skipped -Because 'Outlook integration tests require the winkit -Outlook flag and an Outlook installation.'
    }
  }
  return
}

BeforeAll {
  Import-Module PSFoundation -Force

  $script:OfficeScripts = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '..\..\scripts\Office'))
  $script:Generator = Join-Path -Path $script:OfficeScripts -ChildPath 'New-TestOutlookMessage.ps1'
  $script:Optimizer = Join-Path -Path $script:OfficeScripts -ChildPath 'Optimize-Outlook.ps1'
  $script:Archiver = Join-Path -Path $script:OfficeScripts -ChildPath 'New-OutlookArchive.ps1'
  $script:Repairer = Join-Path -Path $script:OfficeScripts -ChildPath 'Repair-OutlookDataFile.ps1'

  $script:TempRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'winkit-outlook-test'
  $null = New-Item -Path $script:TempRoot -ItemType Directory -Force
  $script:TestStorePst = Join-Path -Path $script:TempRoot -ChildPath 'winkit-test-store.pst'
  $script:ArchiveCopyPst = Join-Path -Path $script:TempRoot -ChildPath 'winkit-test-archive-copy.pst'
  $script:ArchiveMovePst = Join-Path -Path $script:TempRoot -ChildPath 'winkit-test-archive-move.pst'

  foreach ($_file in @($script:TestStorePst, $script:ArchiveCopyPst, $script:ArchiveMovePst)) {
    if (Test-Path -LiteralPath $_file) {
      Remove-Item -LiteralPath $_file -Force
    }
  }

  $script:RedemptionAvailable = $false
  try {
    $null = New-Object -ComObject Redemption.RDOSession
    $script:RedemptionAvailable = $true
  }
  catch {
    Write-Verbose 'Redemption is not available; header-dependent assertions will be skipped.'
  }

  $script:RepairToolAvailable = $false
  try {
    $script:RepairToolAvailable = [bool](Find-OutlookRepairTool -Name ScanPST)
  }
  catch {
    Write-Verbose 'No ScanPST discovery result; repair tests will be skipped.'
  }

  $script:OptimizeSupported = [bool](Get-Module PSFoundation -ListAvailable | Where-Object { $_.Version -ge [version]'1.3.0' })

  $script:Context = Connect-Outlook
  $script:StoreRoot = Add-OutlookStoreRoot -Namespace $script:Context.Namespace -Path $script:TestStorePst
  $script:StoreName = $script:StoreRoot.Name
}

Describe 'New-TestOutlookMessage' {
  It 'creates the requested number of deterministic items' {
    $_generatorArgs = @{
      Count = 20
      Seed = 42
      DuplicateRatio = 0.25
      StoreName = $script:StoreName
      PassThru = $true
    }
    if ($script:RedemptionAvailable) { $_generatorArgs['UseRedemption'] = $true }

    $_results = & $script:Generator @_generatorArgs
    $_created = @($_results | Where-Object { $_.Status -eq 'Created' })
    $_created.Count | Should -Be 20
    $_created[0].Target | Should -Be 'Winkit synthetic 1 (seed 42)'
    $_created[-1].Target | Should -Be 'Winkit synthetic 20 (seed 42)'
  }
}

Describe 'Optimize-Outlook deduplication' {
  It 'walks the generated folder and counts mail items' {
    if (-not $script:OptimizeSupported) {
      Set-ItResult -Skipped -Because 'Optimize-Outlook requires PSFoundation 1.3.0 (Get-TransportMessageId).'
      return
    }

    $_context = Connect-Outlook
    try {
      $_store = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $script:StoreName
      $_folder = Get-OutlookSubFolder -ParentFolder $_store -Name 'WinkitTestData'
      $_items = $_folder.Items
      try {
        $_mailCount = 0
        for ($_index = 1; $_index -le $_items.Count; $_index++) {
          $_item = $_items.Item($_index)
          try {
            if ($_item.Class -eq 43) { $_mailCount++ }
          }
          finally {
            Remove-ComObject $_item
          }
        }
        $_mailCount | Should -Be 20
      }
      finally {
        Remove-ComObject $_items
      }
    }
    finally {
      Remove-ComObject $_folder $_store
      Remove-ComObject $_context.Namespace $_context.App
    }
  }

  It 'flags exactly the seeded duplicates in a dry run' {
    if (-not $script:OptimizeSupported) {
      Set-ItResult -Skipped -Because 'Optimize-Outlook requires PSFoundation 1.3.0 (Get-TransportMessageId).'
      return
    }

    $_context = Connect-Outlook
    try {
      $_store = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $script:StoreName
      $_folder = Get-OutlookSubFolder -ParentFolder $_store -Name 'WinkitTestData'
      $_items = $_folder.Items
      $_headersSeen = $false
      try {
        for ($_index = 1; $_index -le $_items.Count; $_index++) {
          $_item = $_items.Item($_index)
          try {
            $_accessor = $null
            try {
              $_accessor = $_item.PropertyAccessor
              if (-not [string]::IsNullOrWhiteSpace([string]($_accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x007D001E')))) {
                $_headersSeen = $true
              }
            }
            catch {
              Write-Verbose 'Could not read transport headers for a test item.'
            }
            finally {
              Remove-ComObject $_accessor
            }
          }
          finally {
            Remove-ComObject $_item
          }
        }
      }
      finally {
        Remove-ComObject $_items
      }

      if (-not $_headersSeen) {
        Set-ItResult -Skipped -Because 'Transport headers were not injected (install Redemption or rerun with -UseRedemption).'
        return
      }

      $_results = & $script:Optimizer -StoreName $script:StoreName -DryRun -PassThru
      @($_results | Where-Object { $_.Action -eq 'MoveDuplicate' -and $_.Detail -eq 'DryRun' }).Count | Should -Be 5
      @($_results | Where-Object { $_.Status -eq 'Kept' }).Count | Should -Be 15
    }
    finally {
      Remove-ComObject $_store
      Remove-ComObject $_context.Namespace $_context.App
    }
  }

  It 'moves the seeded duplicates to the review folder' {
    if (-not $script:OptimizeSupported) {
      Set-ItResult -Skipped -Because 'Optimize-Outlook requires PSFoundation 1.3.0 (Get-TransportMessageId).'
      return
    }

    $_results = & $script:Optimizer -StoreName $script:StoreName -Confirm:$false -PassThru
    @($_results | Where-Object { $_.Action -eq 'MoveDuplicate' -and $_.Status -eq 'Moved' }).Count | Should -Be 5
  }
}

Describe 'New-TestOutlookMessage determinism' {
  It 'seeds a deterministic Message-ID sequence with duplicates' {
    $_generatorArgs = @{
      Count = 8
      Seed = 7
      DuplicateRatio = 0.25
      StoreName = $script:StoreName
      TargetFolderName = 'WinkitIds'
      PassThru = $true
    }
    if ($script:RedemptionAvailable) { $_generatorArgs['UseRedemption'] = $true }

    $_results = & $script:Generator @_generatorArgs
    $_ids = @($_results | Where-Object { $_.Status -eq 'Created' } | ForEach-Object { $_.MessageId }) -join ';'
    $_ids | Should -Be '<winkit-7-1@synthetic.local>;<winkit-7-2@synthetic.local>;<winkit-7-3@synthetic.local>;<winkit-7-4@synthetic.local>;<winkit-7-5@synthetic.local>;<winkit-7-6@synthetic.local>;<winkit-7-1@synthetic.local>;<winkit-7-2@synthetic.local>'
  }
}

Describe 'New-OutlookArchive' {
  It 'previews the copy without creating a PST' {
    $_results = & $script:Archiver -ArchivePath $script:ArchiveCopyPst -StoreName $script:StoreName -Mode Copy -DryRun -PassThru
    @($_results | Where-Object { $_.Detail -eq 'DryRun' }).Count | Should -BeGreaterThan 0
    Test-Path -LiteralPath $script:ArchiveCopyPst | Should -BeFalse
  }

  It 'copies the generated mail into the archive PST' {
    $_results = & $script:Archiver -ArchivePath $script:ArchiveCopyPst -StoreName $script:StoreName -Mode Copy -PassThru
    @($_results | Where-Object { $_.Status -eq 'Copied' }).Count | Should -BeGreaterThan 0
    Test-Path -LiteralPath $script:ArchiveCopyPst | Should -BeTrue
  }

  It 'honours StartDate bounds when ReceivedTime was injected' {
    $_generatorArgs = @{
      Count = 10
      Seed = 9
      DuplicateRatio = 0.0
      StoreName = $script:StoreName
      TargetFolderName = 'WinkitDated'
      StartDate = '2024-01-01'
      EndDate = '2024-12-31'
      PassThru = $true
    }
    if ($script:RedemptionAvailable) { $_generatorArgs['UseRedemption'] = $true }

    $_created = @(& $script:Generator @_generatorArgs | Where-Object { $_.Status -eq 'Created' })
    $_injected = @($_created | Where-Object { $_.HeaderInjected }).Count
    if ($_injected -eq 0) {
      Set-ItResult -Skipped -Because 'ReceivedTime injection requires Redemption (rerun with -UseRedemption).'
      return
    }

    $_results = & $script:Archiver -ArchivePath $script:ArchiveCopyPst -StoreName $script:StoreName -Mode Copy -StartDate '2024-07-01' -EndDate '2024-12-31' -PassThru
    @($_results | Where-Object { $_.Status -eq 'Copied' }).Count | Should -Be 5
  }

  It 'moves the generated mail into the archive PST' {
    $_results = & $script:Archiver -ArchivePath $script:ArchiveMovePst -StoreName $script:StoreName -Mode Move -PassThru
    @($_results | Where-Object { $_.Status -eq 'Moved' }).Count | Should -BeGreaterThan 0
    Test-Path -LiteralPath $script:ArchiveMovePst | Should -BeTrue
  }
}

Describe 'Repair-OutlookDataFile' {
  It 'previews the repair command for the archive PST' {
    if (-not $script:RepairToolAvailable) {
      Set-ItResult -Skipped -Because 'No ScanPST.exe installation was discovered.'
      return
    }

    $_results = & $script:Repairer -Path $script:ArchiveCopyPst -DryRun -PassThru
    @($_results | Where-Object { $_.Status -eq 'Skipped' -and $_.Detail -like 'DryRun:*' }).Count | Should -Be 1
  }
}

AfterAll {
  if ($script:StoreRoot) {
    try {
      $script:Context.Namespace.RemoveStore($script:StoreRoot)
    }
    catch {
      Write-Warning "Could not detach test PST: $($_.Exception.Message)"
    }
  }

  Remove-ComObject $script:StoreRoot
  if ($script:Context) {
    Remove-ComObject $script:Context.Namespace $script:Context.App
  }

  Invoke-ComGarbageCollection

  Remove-Item -LiteralPath $script:TempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
