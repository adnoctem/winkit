#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Cross-script mock context is removed after each test; Pester mocks execute in the invoked script scope.')]
param ()

# Exercise real script functions with mutable Outlook-shaped objects. No COM,
# profile, or real mail is opened. Script-level guards are tested with mocks.
BeforeAll {
  Import-Module PSFoundation -Force
  $script:OfficePath = Join-Path $PSScriptRoot '../../scripts/Office'
  foreach ($name in @('New-OutlookArchive', 'Optimize-Outlook', 'New-TestOutlookMessage')) {
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:OfficePath "$name.ps1"), [ref]$null, [ref]$null)
    foreach ($definition in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
      . ([scriptblock]::Create($definition.Extent.Text))
    }
  }

  function New-FakeCollection {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only in-memory test fixtures.')]
    param ([object[]]$Values = @())
    $collection = [PSCustomObject]@{ Values = (New-Object Collections.ArrayList) }
    foreach ($value in $Values) { $null = $collection.Values.Add($value) }
    $collection | Add-Member ScriptProperty Count { $this.Values.Count }
    $collection | Add-Member ScriptMethod Item { param($Index) $this.Values[$Index - 1] }
    return $collection
  }

  function New-FakeMail {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only in-memory test fixtures.')]
    param ([string]$Id, [string]$MessageId = '<same@test>')
    $mail = [PSCustomObject]@{
      EntryID = $Id; Subject = $Id; ReceivedTime = [datetime]'2024-12-31T12:00:00'
      Class = 43; MessageId = $MessageId; FailMove = $false; Copies = 0; Moves = 0
      SourceItems = $null
    }
    $mail | Add-Member ScriptMethod Copy {
      $this.Copies++
      $copy = New-FakeMail -Id ($this.EntryID + '-copy')
      $copy.FailMove = $this.FailMove
      $copy.SourceItems = $this.SourceItems
      # Outlook Copy changes the live source collection. Insert ahead of the
      # current item to catch reliance on mutable numeric positions.
      $this.SourceItems.Values.Insert(0, $copy)
      return $copy
    }
    $mail | Add-Member ScriptMethod Move {
      param($Destination)
      if ($this.FailMove) { throw 'Disk full' }
      $this.Moves++
      $this.SourceItems.Values.Remove($this)
      $null = $Destination.Items.Values.Add($this)
      return $this
    }
    return $mail
  }

  function New-FakeFolder {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only in-memory test fixtures.')]
    param ([string]$Name, [object[]]$Mail = @(), [int]$FolderType = 1)
    $accessor = [PSCustomObject]@{ FolderType = $FolderType }
    $accessor | Add-Member ScriptMethod GetProperty {
      param($Tag)
      if ($Tag -ne 'http://schemas.microsoft.com/mapi/proptag/0x36010003') { throw 'Unexpected property tag' }
      $this.FolderType
    }
    $items = New-FakeCollection -Values $Mail
    foreach ($item in $Mail) {
      $item.SourceItems = $items
    }

    [PSCustomObject]@{
      Name = $Name; FolderPath = "\\Test\$Name"; StoreID = 'source-store'
      Items = $items; Folders = (New-FakeCollection)
      DefaultItemType = 0; PropertyAccessor = $accessor
    }
  }
}

Describe 'Outlook archive safety' {
  BeforeEach {
    Mock Remove-ComObject { }
    $script:StartDate = $null
    $script:EndDate = $null
    $script:EndBefore = $null
    $script:OutlookArchiveCopied = 0
    $script:OutlookArchiveMoved = 0
    $script:OutlookArchiveFoldersRead = 0
    $script:OutlookArchiveFoldersSkipped = 0
    $script:OutlookArchiveItemsRead = 0
    $script:OutlookArchiveItemsMatched = 0
    $script:OutlookArchiveProgressTimer = [Diagnostics.Stopwatch]::StartNew()
    Mock Write-Progress { }
    $script:Source = New-FakeFolder -Name Inbox -Mail @((New-FakeMail A), (New-FakeMail B))
    $script:Destination = New-FakeFolder -Name Archive
    $namespace = [PSCustomObject]@{}
    $namespace | Add-Member ScriptMethod GetItemFromID {
      param($Id, $StoreId)
      if ($StoreId -ne $script:Source.StoreID) { throw 'Wrong store identifier' }
      @($script:Source.Items.Values | Where-Object EntryID -EQ $Id)[0]
    }
    $script:_context = [PSCustomObject]@{ Namespace = $namespace }
    $script:_archivePath = 'unused.pst'
    $script:Results = New-Object Collections.ArrayList
  }

  It 'processes each original once even when Copy reorders the source collection' {
    Copy-OutlookFolderItem -SourceFolder $script:Source -DestinationFolder $script:Destination -ArchiveMode Copy -Results $script:Results -Confirm:$false
    $script:Source.Items.Count | Should -Be 2
    $script:Destination.Items.Count | Should -Be 2
    @($script:Source.Items.Values | Where-Object Copies -NE 1).Count | Should -Be 0
    @($script:Results | Where-Object Status -EQ Copied).Count | Should -Be 2
  }

  It 'stops on a failed copy transfer without reporting success or retrying the orphan' {
    $script:Source.Items.Item(1).FailMove = $true
    { Copy-OutlookFolderItem -SourceFolder $script:Source -DestinationFolder $script:Destination -ArchiveMode Copy -Results $script:Results -Confirm:$false } | Should -Throw '*Disk full*'
    $script:Destination.Items.Count | Should -Be 0
    $script:Source.Items.Count | Should -Be 3
    $script:OutlookArchiveCopied | Should -Be 0
    $script:Results.Count | Should -Be 0
  }

  It 'records successful moves only after the transfer' {
    Copy-OutlookFolderItem -SourceFolder $script:Source -DestinationFolder $script:Destination -ArchiveMode Move -Results $script:Results -Confirm:$false
    $script:Source.Items.Count | Should -Be 0
    $script:Destination.Items.Count | Should -Be 2
    @($script:Results | Where-Object Status -EQ Moved).Count | Should -Be 2
  }

  It 'does not modify either folder during WhatIf' {
    Copy-OutlookFolderItem -SourceFolder $script:Source -DestinationFolder $script:Destination -ArchiveMode Move -Results $script:Results -WhatIf
    $script:Source.Items.Count | Should -Be 2
    $script:Destination.Items.Count | Should -Be 0
    @($script:Results | Where-Object Detail -EQ DryRun).Count | Should -Be 2
    Should -Invoke Write-Progress -ParameterFilter { $Status -eq 'Reading folder contents and applying date filters' }
    Should -Invoke Write-Progress -ParameterFilter { $Status -eq 'Recording preview results' }
  }

  It 'includes the whole final day with an exclusive next-day cutoff' {
    $script:StartDate = [datetime]'2024-01-01'
    $script:EndBefore = [datetime]'2025-01-01'
    $mail = New-FakeMail A
    Test-OutlookItemInRange $mail | Should -BeTrue
    $mail.ReceivedTime = [datetime]'2025-01-01'
    Test-OutlookItemInRange $mail | Should -BeFalse
    $mail.ReceivedTime = [datetime]'2023-12-31T23:59:59'
    Test-OutlookItemInRange $mail | Should -BeFalse
  }

  It 'never treats contacts as archive mail even without date filters' {
    $mail = New-FakeMail A
    $mail.Class = 40
    Test-OutlookItemInRange $mail | Should -BeFalse
  }

  It 'skips virtual search folders before creating destinations' {
    Mock Get-OutlookSubFolder { throw 'Must not create a folder' }
    $folder = New-FakeFolder -Name Search -FolderType 2
    Copy-OutlookFolderTree -SourceFolder $folder -DestinationParent $script:Destination -ArchiveMode Move -Results $script:Results -Confirm:$false
    Should -Invoke Get-OutlookSubFolder -Times 0
  }
}

Describe 'Outlook deduplication safety' {
  BeforeEach {
    Mock Remove-ComObject { }
    Mock Get-MessageId { param($Item) $Item.MessageId }
    $script:OL_MAIL = 43
    $script:Results = New-Object Collections.ArrayList
    $script:Source = New-FakeFolder -Name Inbox -Mail @((New-FakeMail A), (New-FakeMail B))
    $script:Destination = New-FakeFolder -Name Review
  }

  It 'does not report a failed duplicate move as Moved' {
    $script:Source.Items.Item(1).FailMove = $true
    { Optimize-OutlookFolder -Folder $script:Source -ReviewFolder $script:Destination -Results $script:Results -Confirm:$false } | Should -Throw '*Disk full*'
    @($script:Results | Where-Object Status -EQ Moved).Count | Should -Be 0
    @($script:Results | Where-Object Status -EQ Failed).Count | Should -Be 1
    $script:Source.Items.Count | Should -Be 2
  }

  It 'does not collapse case-distinct message identifiers' {
    $script:Source.Items.Item(1).MessageId = '<Case@test>'
    $script:Source.Items.Item(2).MessageId = '<case@test>'
    Optimize-OutlookFolder -Folder $script:Source -ReviewFolder $script:Destination -Results $script:Results -Confirm:$false
    $script:Source.Items.Count | Should -Be 2
    $script:Destination.Items.Count | Should -Be 0
  }

  It 'excludes the entire subtree under an excluded folder' {
    Mock Optimize-OutlookFolder { throw 'Must not process excluded descendants' }
    $excluded = New-FakeFolder -Name 'Deleted Items'
    $null = $excluded.Folders.Values.Add($script:Source)
    Invoke-OutlookFolderTree -Folder $excluded -ReviewFolder $script:Destination -ReviewName Review -Exclude @('Deleted Items') -Results $script:Results -Confirm:$false
    Should -Invoke Optimize-OutlookFolder -Times 0
  }
}

Describe 'Archive script preflight' {
  BeforeEach {
    Mock Import-Module { }
    Mock Connect-Outlook { throw 'Must not open Outlook' }
    Mock Remove-ComObject { }
    Mock Invoke-ComGarbageCollection { }
    Mock Write-OperationResultLog { }
    Mock Write-Log { }
  }

  It 'refuses an existing PST before connecting to Outlook' {
    $path = Join-Path $TestDrive 'existing.pst'
    [IO.File]::WriteAllText($path, 'existing data')
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath $path -ReportDirectory $TestDrive -PassThru -WarningAction SilentlyContinue
    $LASTEXITCODE | Should -Be 1
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'already exists'
    [IO.File]::ReadAllText($path) | Should -Be 'existing data'
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'refuses conflicting date bounds before connecting to Outlook' {
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportDirectory $TestDrive -EndDate '2024-12-31' -EndBefore '2025-01-01' -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'not both'
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'stops before Outlook when the report directory is not writable' {
    $blocked = Join-Path $TestDrive 'not-a-directory'
    [IO.File]::WriteAllText($blocked, 'keep')
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportDirectory $blocked -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.ReportPath | Should -BeNullOrEmpty
    $LASTEXITCODE | Should -Be 1
    Should -Invoke Connect-Outlook -Times 0
    [IO.File]::ReadAllText($blocked) | Should -Be 'keep'
  }

  It 'preserves an existing explicitly named report' {
    $reportPath = Join-Path $TestDrive 'existing-report.json'
    [IO.File]::WriteAllText($reportPath, 'original report')
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportPath $reportPath -DryRun -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'already exists'
    $result.ReportPath | Should -BeNullOrEmpty
    [IO.File]::ReadAllText($reportPath) | Should -Be 'original report'
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'rejects conflicting report options without creating either destination' {
    $reportPath = Join-Path $TestDrive 'conflict.json'
    $reportDirectory = Join-Path $TestDrive 'conflict-directory'
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportPath $reportPath -ReportDirectory $reportDirectory -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'not both'
    Test-Path -LiteralPath $reportPath | Should -BeFalse
    Test-Path -LiteralPath $reportDirectory | Should -BeFalse
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'rejects a report path that resolves to the archive destination' {
    $path = Join-Path $TestDrive 'same.pst'
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath $path -ReportPath $path -DryRun -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'different files'
    Test-Path -LiteralPath $path | Should -BeFalse
    Should -Invoke Connect-Outlook -Times 0
  }
}

Describe 'Repair tool launch safety' {
  BeforeEach {
    Mock Import-Module { }
    Mock Write-Log { }
    Mock Write-OperationResultLog { }
    Mock Get-Process { }
    Mock Start-Process { [PSCustomObject]@{ ExitCode = 0 } }
    $script:DataPath = Join-Path $TestDrive 'mail with spaces.pst'
    $script:ToolPath = Join-Path $TestDrive 'SCANPST.EXE'
    [IO.File]::WriteAllText($script:DataPath, 'fake pst')
    [IO.File]::WriteAllText($script:ToolPath, 'fake executable')
  }

  It 'quotes the data file as one native argument' {
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    Should -Invoke Start-Process -Times 1 -ParameterFilter { $ArgumentList -eq ('"' + $script:DataPath + '"') }
  }

  It 'does not launch while Outlook is running' {
    Mock Get-Process { [PSCustomObject]@{ Name = 'OUTLOOK' } }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -PassThru -Confirm:$false
    $result.Status | Should -Be Failed
    Should -Invoke Start-Process -Times 0
  }

  It 'returns a failure when the repair tool exits nonzero' {
    Mock Start-Process { [PSCustomObject]@{ ExitCode = 3 } }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -PassThru -Confirm:$false
    $result.Status | Should -Be Failed
    $LASTEXITCODE | Should -Be 1
  }

  It 'does not launch against a locked PST' {
    $lock = [IO.File]::Open($script:DataPath, 'Open', 'ReadWrite', 'None')
    try {
      $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -PassThru -Confirm:$false
      $result.Status | Should -Be Failed
      Should -Invoke Start-Process -Times 0
    }
    finally { $lock.Dispose() }
  }
}

Describe 'Outlook store selection and preview' {
  BeforeEach {
    Mock Write-Progress { }
    Mock Import-Module { }
    Mock Remove-ComObject { }
    Mock Invoke-ComGarbageCollection { }
    Mock Write-OperationResultLog { }
    Mock Write-Log { }
    Mock Get-OutlookStoreRoot { throw 'Default selection must use Namespace.DefaultStore' }
    Mock Add-OutlookStoreRoot { throw 'Preview must not mount a store' }
    Mock Get-OutlookSubFolder { throw 'Preview must not create a folder' }
    $script:Source = New-FakeFolder -Name Root
    $store = [PSCustomObject]@{ DisplayName = 'Duplicate name'; FilePath = 'source.pst'; Root = $script:Source }
    $store | Add-Member ScriptMethod GetRootFolder { $this.Root }
    $app = [PSCustomObject]@{ Version = '12.0'; QuitCalled = $false }
    $app | Add-Member ScriptMethod Quit { $this.QuitCalled = $true }
    $script:FakeContext = [PSCustomObject]@{
      App       = $app
      Namespace = [PSCustomObject]@{ DefaultStore = $store; Stores = (New-FakeCollection @($store, $store)) }
    }
    $global:WinkitSafetyTestContext = $script:FakeContext
    Mock Connect-Outlook { $global:WinkitSafetyTestContext }
  }

  AfterEach {
    Remove-Variable -Name WinkitSafetyTestContext -Scope Global -ErrorAction SilentlyContinue
  }

  It 'previews all three scripts using the default store without creating folders or quitting Outlook' {
    $archiveResults = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'preview.pst') -ReportDirectory $TestDrive -DryRun -QuitOutlook
    $dedupResults = & (Join-Path $script:OfficePath 'Optimize-Outlook.ps1') -DryRun -QuitOutlook
    $generatorResults = & (Join-Path $script:OfficePath 'New-TestOutlookMessage.ps1') -Count 1 -DryRun -QuitOutlook
    @($archiveResults | Where-Object Status -EQ Failed).Count | Should -Be 0
    @($dedupResults | Where-Object Status -EQ Failed).Count | Should -Be 0
    $generatorResults.Status | Should -Be Skipped
    $generatorResults.Detail | Should -Be DryRun
    Should -Invoke Connect-Outlook -Times 3
    Should -Invoke Get-OutlookStoreRoot -Times 0
    Should -Invoke Add-OutlookStoreRoot -Times 0
    Should -Invoke Get-OutlookSubFolder -Times 0
    $script:FakeContext.App.QuitCalled | Should -BeFalse
    Test-Path (Join-Path $TestDrive 'preview.pst') | Should -BeFalse
  }

  It 'refuses ambiguous store names before mutation' {
    foreach ($name in @('New-OutlookArchive', 'Optimize-Outlook', 'New-TestOutlookMessage')) {
      $arguments = @{ StoreName = 'Duplicate name'; PassThru = $true; WarningAction = 'SilentlyContinue' }
      if ($name -eq 'New-OutlookArchive') {
        $arguments.ArchivePath = Join-Path $TestDrive 'ambiguous.pst'
        $arguments.ReportDirectory = $TestDrive
      }
      $result = & (Join-Path $script:OfficePath "$name.ps1") @arguments
      $result.Status | Should -Be Failed
      $result.Detail | Should -Match 'matches 2 stores'
    }
    Should -Invoke Add-OutlookStoreRoot -Times 0
    Should -Invoke Get-OutlookSubFolder -Times 0
  }

  It 'writes an empty Results array and returns one preview summary' {
    $result = @(& (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'empty.pst') -ReportDirectory $TestDrive -WhatIf -PassThru)
    $result.Count | Should -Be 1
    $result[0].Status | Should -Be Preview
    $result[0].Planned | Should -Be 0
    $json = Get-Content -LiteralPath $result[0].ReportPath -Raw -Encoding UTF8
    $json | Should -Match '"Results":\s*\[\s*\]'
    ($json | ConvertFrom-Json).Summary.FoldersRead | Should -Be 1
    Should -Invoke Write-Progress -ParameterFilter { $Completed }
  }

  It 'resolves an exact relative report path and creates its parent during WhatIf' {
    Push-Location $TestDrive
    try {
      $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'relative.pst') -ReportPath '.\reports\archive-report.json' -WhatIf
      $result.Status | Should -Be Preview
      $result.ReportPath | Should -Be (Join-Path $TestDrive 'reports\archive-report.json')
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $report.Summary.ReportPath | Should -Be $result.ReportPath
      Test-Path -LiteralPath (Join-Path $TestDrive 'relative.pst') | Should -BeFalse
    }
    finally {
      Pop-Location
    }
  }

  It 'keeps a large filtered preview out of the console and preserves Unicode mail in JSON' {
    for ($index = 0; $index -lt 300; $index++) {
      $mail = New-FakeMail -Id "mail-$index"
      $mail.Subject = "Grüße aus Köln — $index"
      $null = $script:Source.Items.Values.Add($mail)
    }

    $excluded = New-FakeMail -Id 'outside-date-range'
    $excluded.ReceivedTime = [datetime]'2025-01-01'
    $null = $script:Source.Items.Values.Add($excluded)
    $script:FakeContext.Namespace | Add-Member ScriptMethod GetItemFromID {
      param($Id, $StoreId)
      if ($StoreId -ne $this.DefaultStore.Root.StoreID) {
        throw 'Wrong store identifier'
      }

      @($this.DefaultStore.Root.Items.Values | Where-Object EntryID -EQ $Id)[0]
    }

    $arguments = @{
      ArchivePath     = Join-Path $TestDrive 'large.pst'
      ReportDirectory = Join-Path $TestDrive 'reports'
      StartDate       = '2024-01-01'
      EndBefore       = '2025-01-01'
      DryRun          = $true
      PassThru        = $true
    }

    $captured = @(& (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') @arguments *>&1)
    $captured.Count | Should -Be 1
    $result = $captured[0]
    $result.ItemsRead | Should -Be 301
    $result.ItemsMatched | Should -Be 300
    $result.Planned | Should -Be 300
    $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $report.Results.Count | Should -Be 300
    $report.Results[0].Target | Should -Be 'Grüße aus Köln — 0'
    $report.Results[0].Scope | Should -Be '\\Test\Root'
    $report.Results[0].Action | Should -Be Copy
    $report.Results[0].Detail | Should -Be DryRun
    $report.Results[0].Received | Should -Not -BeNullOrEmpty
    $report.Settings.EndBefore | Should -Not -BeNullOrEmpty
    Test-Path -LiteralPath $arguments.ArchivePath | Should -BeFalse
    @($script:Source.Items.Values | Where-Object { $_.Copies -or $_.Moves }).Count | Should -Be 0

    $again = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') @arguments
    $again.ReportPath | Should -Not -Be $result.ReportPath
    Test-Path -LiteralPath $result.ReportPath | Should -BeTrue
  }

  It 'reports JSON serialization failure instead of claiming a report was saved' {
    Mock ConvertTo-Json { throw 'Report write failed' }
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'failed-report.pst') -ReportDirectory $TestDrive -DryRun -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.ReportPath | Should -BeNullOrEmpty
    $result.Detail | Should -Match 'Report write failed'
    $LASTEXITCODE | Should -Be 1
    Should -Invoke Write-Progress -ParameterFilter { $Completed }
  }

  It 'preserves successful <Mode> records and partial failures in the report' -ForEach @(
    @{ Mode = 'Copy'; ResultStatus = 'Copied' }
    @{ Mode = 'Move'; ResultStatus = 'Moved' }
  ) {
    $goodMail = New-FakeMail A
    $goodMail.SourceItems = $script:Source.Items
    $null = $script:Source.Items.Values.Add($goodMail)
    $badMail = New-FakeMail B
    $badMail.FailMove = $true
    $badMail.SourceItems = $script:Source.Items
    $null = $script:Source.Items.Values.Add($badMail)
    $destination = New-FakeFolder -Name Archive
    $destination.StoreID = 'archive-store'
    $script:FakeContext | Add-Member NoteProperty ArchiveRoot $destination
    $script:FakeContext.Namespace | Add-Member ScriptMethod GetItemFromID {
      param($Id, $StoreId)
      if ($StoreId -ne $this.DefaultStore.Root.StoreID) {
        throw 'Wrong store identifier'
      }

      @($this.DefaultStore.Root.Items.Values | Where-Object EntryID -EQ $Id)[0]
    }

    Mock Add-OutlookStoreRoot { $global:WinkitSafetyTestContext.ArchiveRoot }
    $arguments = @{
      ArchivePath     = Join-Path $TestDrive 'partial.pst'
      ReportDirectory = $TestDrive
      Mode            = $Mode
      DetachWhenDone  = $false
      PassThru        = $true
      Confirm         = $false
      WarningAction   = 'SilentlyContinue'
    }

    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') @arguments
    $result.Status | Should -Be Failed
    $result.Failed | Should -Be 1
    $result.$ResultStatus | Should -Be 1 -Because $result.Detail
    $LASTEXITCODE | Should -Be 1
    $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $report.Results.Count | Should -Be 2
    $report.Results[0].Status | Should -Be $ResultStatus
    $report.Results[0].Target | Should -Be A
    $report.Results[1].Status | Should -Be Failed
    $report.Results[1].Detail | Should -Match 'Disk full'
    $destination.Items.Count | Should -Be 1
    Should -Invoke Write-Progress -ParameterFilter { $Completed }
  }
}
