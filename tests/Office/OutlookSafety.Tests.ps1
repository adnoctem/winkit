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
      EntryID = 'folder-' + $Name
      Items = $items; Folders = (New-FakeCollection)
      DefaultItemType = 0; PropertyAccessor = $accessor
    }
  }
}

Describe 'Outlook report subjects' -ForEach @(
  @{ ResultFunction = 'Add-OutlookArchiveResult' }
  @{ ResultFunction = 'Add-OutlookItemResult' }
) {
  It 'labels <Case> subjects without changing the source message' -ForEach @(
    @{ Case = 'null'; Subject = $null; Expected = '<No Subject>' }
    @{ Case = 'empty'; Subject = ''; Expected = '<No Subject>' }
    @{ Case = 'whitespace'; Subject = " `t "; Expected = '<No Subject>' }
    @{ Case = 'nonblank'; Subject = '  Grüße aus Köln  '; Expected = '  Grüße aus Köln  ' }
  ) {
    $mail = New-FakeMail -Id 'subject-test'
    $mail.Subject = $Subject
    $results = New-Object Collections.ArrayList

    & $ResultFunction -Results $results -Item $mail -Action Copy -Status Skipped -Folder '\\Test\Inbox' -Detail DryRun

    $results.Count | Should -Be 1
    $json = ConvertTo-Json -InputObject @($results) -Depth 5 | ConvertFrom-Json
    $json[0].Target | Should -BeExactly $Expected
    $csv = $results | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv
    $csv.Target | Should -BeExactly $Expected
    $mail.Subject | Should -BeExactly $Subject
  }
}

Describe 'Outlook profile elevation guard' -ForEach @(
  @{ ScriptName = 'New-OutlookArchive' }
  @{ ScriptName = 'Optimize-Outlook' }
  @{ ScriptName = 'New-TestOutlookMessage' }
) {
  BeforeEach {
    Mock Import-Module { }
    Mock Get-UserInfo { @{ UserName = 'TEST\Administrator'; IsAdministrator = $true } }
    Mock Connect-Outlook { throw 'Connection reached for guard test' }
    Mock Write-Log { }
    Mock Write-OperationResultLog { }
    Mock Remove-ComObject { }
    Mock Invoke-ComGarbageCollection { }
    $script:ProfileGuardScriptPath = Join-Path $script:OfficePath "$ScriptName.ps1"
    $arguments = @{ PassThru = $true }
    if ($ScriptName -eq 'New-OutlookArchive') {
      $arguments.ArchivePath = Join-Path $TestDrive 'guard.pst'
      $arguments.ReportDirectory = Join-Path $TestDrive 'reports'
    }
  }

  It 'blocks elevated execution before COM access or report creation, including previews' {
    foreach ($preview in @($false, $true)) {
      { & $script:ProfileGuardScriptPath @arguments -DryRun:$preview } | Should -Throw '*non-elevated*IgnoreAdministrator*'
    }

    Should -Invoke Connect-Outlook -Times 0
    Should -Invoke Write-OperationResultLog -Times 0
    Test-Path -LiteralPath (Join-Path $TestDrive 'reports') | Should -BeFalse
  }

  It 'allows the explicit override with a context warning' {
    $result = & $script:ProfileGuardScriptPath @arguments -IgnoreAdministrator -DryRun -WarningVariable warnings -WarningAction SilentlyContinue
    Should -Invoke Connect-Outlook -Times 1 -Exactly
    ($warnings -join ' ') | Should -Match 'IgnoreAdministrator permits elevated execution'
    ($warnings -join ' ') | Should -Match 'same Windows user and elevation'
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'Connection reached for guard test'
  }

  It 'allows a non-elevated user without an override warning' {
    Mock Get-UserInfo { @{ UserName = 'TEST\MailboxUser'; IsAdministrator = $false } }
    $result = & $script:ProfileGuardScriptPath @arguments -DryRun -WarningVariable warnings -WarningAction SilentlyContinue
    Should -Invoke Connect-Outlook -Times 1 -Exactly
    ($warnings -join ' ') | Should -Not -Match 'IgnoreAdministrator permits'
    $result.Detail | Should -Match 'Connection reached for guard test'
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

}

Describe 'Archive script preflight' {
  BeforeEach {
    Mock Get-UserInfo { @{ UserName = 'TEST\MailboxUser'; IsAdministrator = $false } }
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
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath $path -ReportDirectory $TestDrive -PassThru -WarningAction SilentlyContinue
    $LASTEXITCODE | Should -Be 1
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'already exists'
    [IO.File]::ReadAllText($path) | Should -Be 'existing data'
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'refuses conflicting date bounds before connecting to Outlook' {
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportDirectory $TestDrive -EndDate '2024-12-31' -EndBefore '2025-01-01' -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'not both'
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'refuses contradictory attachment options before connecting to Outlook' {
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'conflict.pst') -AddDataFile -DetachWhenDone:$true -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'conflicts'
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'stops before Outlook when the report directory is not writable' {
    $blocked = Join-Path $TestDrive 'not-a-directory'
    [IO.File]::WriteAllText($blocked, 'keep')
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportDirectory $blocked -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.ReportPath | Should -BeNullOrEmpty
    $LASTEXITCODE | Should -Be 1
    Should -Invoke Connect-Outlook -Times 0
    [IO.File]::ReadAllText($blocked) | Should -Be 'keep'
  }

  It 'preserves an existing explicitly named report' {
    $reportPath = Join-Path $TestDrive 'existing-report.json'
    [IO.File]::WriteAllText($reportPath, 'original report')
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportPath $reportPath -DryRun -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'already exists'
    $result.ReportPath | Should -BeNullOrEmpty
    [IO.File]::ReadAllText($reportPath) | Should -Be 'original report'
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'rejects conflicting report options without creating either destination' {
    $reportPath = Join-Path $TestDrive 'conflict.json'
    $reportDirectory = Join-Path $TestDrive 'conflict-directory'
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'new.pst') -ReportPath $reportPath -ReportDirectory $reportDirectory -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'not both'
    Test-Path -LiteralPath $reportPath | Should -BeFalse
    Test-Path -LiteralPath $reportDirectory | Should -BeFalse
    Should -Invoke Connect-Outlook -Times 0
  }

  It 'rejects a report path that resolves to the archive destination' {
    $path = Join-Path $TestDrive 'same.pst'
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath $path -ReportPath $path -DryRun -WarningAction SilentlyContinue
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
    Mock Get-OutlookStandardFolderIdentity -ModuleName PSFoundation { @() }
    Mock Get-UserInfo { @{ UserName = 'TEST\MailboxUser'; IsAdministrator = $false } }
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
    $script:FakeContext.Namespace | Add-Member ScriptMethod GetFolderFromID {
      param($Id, $StoreId)
      if ($StoreId -ne $this.DefaultStore.Root.StoreID) {
        throw 'Wrong store identifier'
      }
      $queue = New-Object Collections.Queue
      $queue.Enqueue($this.DefaultStore.Root)
      while ($queue.Count) {
        $folder = $queue.Dequeue()
        if ($folder.EntryID -eq $Id) {
          return $folder
        }
        foreach ($child in $folder.Folders.Values) {
          $queue.Enqueue($child)
        }
      }
      throw 'Folder not found'
    }

    $global:WinkitSafetyTestContext = $script:FakeContext
    Mock Connect-Outlook { $global:WinkitSafetyTestContext }
  }

  AfterEach {
    Remove-Variable -Name WinkitSafetyTestContext -Scope Global -ErrorAction SilentlyContinue
  }

  It 'defaults to Inbox alone and visits descendants only with Recurse' {
    $inbox = New-FakeFolder -Name Inbox
    $nested = New-FakeFolder -Name Nested
    $protected = New-FakeFolder -Name Protected
    $inbox.FolderPath = $script:Source.FolderPath + '\Inbox'
    $nested.FolderPath = $inbox.FolderPath + '\Nested'
    $protected.FolderPath = $script:Source.FolderPath + '\Protected'
    Mock Get-OutlookStandardFolderIdentity -ModuleName PSFoundation {
      [PSCustomObject]@{ Kind = 'Inbox'; EntryID = 'folder-Inbox'; StoreID = 'source-store'; State = 'Resolved' }
    }
    $null = $inbox.Folders.Values.Add($nested)
    $null = $script:Source.Folders.Values.Add($inbox)
    $null = $script:Source.Folders.Values.Add($protected)
    Mock Get-OutlookSubFolder {
      param($ParentFolder, $Name)
      @($ParentFolder.Folders.Values | Where-Object Name -EQ $Name)[0]
    }
    $single = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'inbox.pst') -ReportDirectory $TestDrive -DryRun
    $single.Status | Should -Be Preview
    $single.FoldersRead | Should -Be 1
    $report = Get-Content -LiteralPath $single.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $report.Settings.SourceFolder | Should -Be '\\Test\Root\Inbox'
    $report.Settings.FolderName | Should -BeNullOrEmpty
    $report.Settings.FolderSelection | Should -Be DefaultInbox
    $report.Settings.Recurse | Should -BeFalse

    $recursive = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -ArchivePath (Join-Path $TestDrive 'recursive.pst') -ReportDirectory $TestDrive -Recurse -DryRun
    $recursive.Status | Should -Be Preview
    $recursive.FoldersRead | Should -Be 2

    $all = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -Recurse -ArchivePath (Join-Path $TestDrive 'all.pst') -ReportDirectory $TestDrive -DryRun
    $all.FoldersRead | Should -Be 4
    $rootReport = Get-Content -LiteralPath $all.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $rootReport.Settings.SourceFolder | Should -Be '\\Test\Root'
    $rootReport.Settings.FolderName | Should -BeExactly ''
    $rootReport.Settings.FolderSelection | Should -Be ExplicitPath
    Should -Invoke Get-OutlookSubFolder -Times 0 -ParameterFilter { $Create }
  }

  It 'selects the implicit <InboxName> Inbox independently of a literal Inbox for <ScriptName>' -ForEach @(
    @{ ScriptName = 'New-OutlookArchive'; InboxName = 'Posteingang' }
    @{ ScriptName = 'New-OutlookArchive'; InboxName = 'Renamed' }
    @{ ScriptName = 'Optimize-Outlook'; InboxName = 'Posteingang' }
    @{ ScriptName = 'Optimize-Outlook'; InboxName = 'Renamed' }
  ) {
    $inbox = New-FakeFolder -Name $InboxName -Mail @((New-FakeMail 'standard-inbox-mail'))
    $inbox.EntryID = 'standard-inbox'
    $inbox.FolderPath = $script:Source.FolderPath + '\' + $InboxName
    $literal = New-FakeFolder -Name Inbox -Mail @((New-FakeMail 'literal-inbox-mail'))
    $literal.FolderPath = $script:Source.FolderPath + '\Inbox'
    $null = $script:Source.Folders.Values.Add($inbox)
    $null = $script:Source.Folders.Values.Add($literal)
    Mock Get-OutlookStandardFolderIdentity -ModuleName PSFoundation {
      [PSCustomObject]@{ Kind = 'Inbox'; EntryID = 'standard-inbox'; StoreID = 'source-store'; State = 'Resolved' }
    }
    $script:FakeContext.Namespace | Add-Member ScriptMethod GetItemFromID {
      param($Id, $StoreId)
      foreach ($folder in $this.DefaultStore.Root.Folders.Values) {
        if ($folder.StoreID -eq $StoreId) {
          foreach ($mail in $folder.Items.Values) {
            if ($mail.EntryID -eq $Id) { return $mail }
          }
        }
      }
      throw 'Item not found'
    }

    $arguments = @{ DryRun = $true; PassThru = $true }
    if ($ScriptName -eq 'New-OutlookArchive') {
      $arguments.ArchivePath = Join-Path $TestDrive 'identity.pst'
      $arguments.ReportDirectory = $TestDrive
    }
    $implicit = & (Join-Path $script:OfficePath "$ScriptName.ps1") @arguments
    $explicit = & (Join-Path $script:OfficePath "$ScriptName.ps1") @arguments -FolderName Inbox
    if ($ScriptName -eq 'New-OutlookArchive') {
      $implicit.Status | Should -Be Preview
      $explicit.Status | Should -Be Preview
      $implicitReport = Get-Content -LiteralPath $implicit.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $explicitReport = Get-Content -LiteralPath $explicit.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $implicitReport.Settings.SourceFolder | Should -Be $inbox.FolderPath
      $explicitReport.Settings.SourceFolder | Should -Be $literal.FolderPath
      $explicitReport.Settings.FolderSelection | Should -Be ExplicitPath
      $explicitReport.Settings.FolderName | Should -Be Inbox
      $implicit = $implicitReport.Results
      $explicit = $explicitReport.Results
    }
    @($implicit).Count | Should -Be 1
    $implicit.Target | Should -Be 'standard-inbox-mail'
    @($explicit).Count | Should -Be 1
    $explicit.Target | Should -Be 'literal-inbox-mail'
    Should -Invoke Add-OutlookStoreRoot -Times 0
    Should -Invoke Get-OutlookSubFolder -Times 0 -ParameterFilter { $Create }
  }

  It 'stops <ScriptName> before mutation when implicit selection is <Case>' -ForEach @(
    @{ ScriptName = 'New-OutlookArchive'; Case = 'unresolved'; Excluded = $null }
    @{ ScriptName = 'Optimize-Outlook'; Case = 'unresolved'; Excluded = $null }
    @{ ScriptName = 'New-OutlookArchive'; Case = 'excluded'; Excluded = 'Parent\Posteingang' }
    @{ ScriptName = 'Optimize-Outlook'; Case = 'excluded'; Excluded = 'Parent\Posteingang' }
    @{ ScriptName = 'New-OutlookArchive'; Case = 'ancestor excluded'; Excluded = 'Parent' }
    @{ ScriptName = 'Optimize-Outlook'; Case = 'ancestor excluded'; Excluded = 'Parent' }
  ) {
    $parent = New-FakeFolder -Name Parent
    $parent.FolderPath = $script:Source.FolderPath + '\Parent'
    $inbox = New-FakeFolder -Name Posteingang
    $inbox.FolderPath = $parent.FolderPath + '\Posteingang'
    $null = $parent.Folders.Values.Add($inbox)
    $null = $script:Source.Folders.Values.Add($parent)
    if ($Excluded) {
      Mock Get-OutlookStandardFolderIdentity -ModuleName PSFoundation {
        [PSCustomObject]@{ Kind = 'Inbox'; EntryID = 'folder-Posteingang'; StoreID = 'source-store'; State = 'Resolved' }
      }
    }
    $arguments = @{ PassThru = $true; Confirm = $false; WarningAction = 'SilentlyContinue' }
    if ($Excluded) {
      $arguments.Exclusions = @($Excluded)
    }
    if ($ScriptName -eq 'New-OutlookArchive') {
      $arguments.ArchivePath = Join-Path $TestDrive 'blocked.pst'
      $arguments.ReportDirectory = $TestDrive
    }
    $result = & (Join-Path $script:OfficePath "$ScriptName.ps1") @arguments
    $result.Status | Should -Be Failed
    if ($Excluded) {
      $result.Detail | Should -Match 'CustomExclusion'
    }
    else {
      $result.Detail | Should -Match 'Cannot resolve.*Inbox identity'
    }
    Should -Invoke Add-OutlookStoreRoot -Times 0
    Should -Invoke Get-OutlookSubFolder -Times 0 -ParameterFilter { $Create }
  }

  It 'resolves localized nested paths and refuses missing folders without broadening scope' {
    $inbox = New-FakeFolder -Name Posteingang
    $nested = New-FakeFolder -Name Kunden
    $null = $inbox.Folders.Values.Add($nested)
    $null = $script:Source.Folders.Values.Add($inbox)
    Mock Get-OutlookSubFolder {
      param($ParentFolder, $Name)
      $found = @($ParentFolder.Folders.Values | Where-Object Name -EQ $Name)
      if ($found.Count) {
        $found[0]
      }
    }

    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName 'Posteingang\Kunden' -ArchivePath (Join-Path $TestDrive 'nested.pst') -ReportDirectory $TestDrive -DryRun
    $result.Status | Should -Be Preview
    $result.FoldersRead | Should -Be 1
    $missing = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName 'missing' -ArchivePath (Join-Path $TestDrive 'missing.pst') -ReportDirectory $TestDrive -DryRun -WarningAction SilentlyContinue
    $missing.Status | Should -Be Failed
    $missing.Detail | Should -Match 'not found'
    $missing.FoldersRead | Should -Be 0
    Should -Invoke Add-OutlookStoreRoot -Times 0
  }

  It 'forwards inclusion flags while custom exclusions still protect localized folders' {
    $inbox = New-FakeFolder -Name Posteingang
    $sent = New-FakeFolder -Name Gesendet
    $junk = New-FakeFolder -Name Unerwuenscht
    foreach ($folder in @($inbox, $sent, $junk)) {
      $null = $script:Source.Folders.Values.Add($folder)
    }
    Mock Get-OutlookStandardFolderIdentity -ModuleName PSFoundation {
      [PSCustomObject]@{ Kind = 'Inbox'; EntryID = 'folder-Posteingang'; State = 'Resolved' }
      [PSCustomObject]@{ Kind = 'SentItems'; EntryID = 'folder-Gesendet'; State = 'Resolved' }
      [PSCustomObject]@{ Kind = 'Junk'; EntryID = 'folder-Unerwuenscht'; State = 'Resolved' }
    }

    $arguments = @{
      FolderName       = ''
      Recurse          = $true
      IncludeInbox     = $false
      IncludeSentItems = $true
      IncludeJunk      = $true
      Exclusions       = @('Unerwuenscht')
      DryRun           = $true
    }
    $archive = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') @arguments -ArchivePath (Join-Path $TestDrive 'policy.pst') -ReportDirectory $TestDrive
    $archive.Status | Should -Be Preview
    $archive.FoldersRead | Should -Be 2
    $report = Get-Content -LiteralPath $archive.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    ($report.FolderPlan | Where-Object StandardKind -EQ Inbox).Process | Should -BeFalse
    ($report.FolderPlan | Where-Object StandardKind -EQ SentItems).Process | Should -BeTrue
    ($report.FolderPlan | Where-Object StandardKind -EQ Junk).Reason | Should -Be CustomExclusion

    $mail = New-FakeMail -Id 'sent-mail'
    $null = $sent.Items.Values.Add($mail)
    $csvPath = Join-Path $TestDrive 'policy.csv'
    $optimizer = @(& (Join-Path $script:OfficePath 'Optimize-Outlook.ps1') @arguments -ReportPath $csvPath)
    @($optimizer | Where-Object Detail -EQ CustomExclusion).Count | Should -Be 1
    @($optimizer | Where-Object { $_.Detail -like '*IncludeInbox*' }).Count | Should -Be 1
    $csvMail = Import-Csv -LiteralPath $csvPath | Where-Object Target -EQ 'sent-mail'
    $csvMail.PSObject.Properties.Name | Should -Contain MessageId
    $csvMail.Received | Should -Not -BeNullOrEmpty
    Should -Invoke Add-OutlookStoreRoot -Times 0
  }

  It 'keeps a requested data file attached with its filename or explicit display name' {
    $destination = New-FakeFolder -Name Archive
    $destination.StoreID = 'archive-store'
    $script:FakeContext | Add-Member NoteProperty ArchiveRoot $destination
    Mock Add-OutlookStoreRoot { $global:WinkitSafetyTestContext.ArchiveRoot }

    $arguments = @{
      ArchivePath     = Join-Path $TestDrive 'Archive - 2018.pst'
      ReportDirectory = $TestDrive
      FolderName      = ''
      AddDataFile     = $true
      PassThru        = $true
      Confirm         = $false
    }
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') @arguments
    $result.Status | Should -Be Completed
    $destination.Name | Should -Be 'Archive - 2018'
    $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $report.Settings.DetachWhenDone | Should -BeFalse
    $report.Settings.AddDataFile | Should -BeTrue

    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') @arguments -DataFileName 'Custom archive'
    $result.Status | Should -Be Completed
    $destination.Name | Should -Be 'Custom archive'
  }

  It 'previews all three scripts using the default store without creating folders or quitting Outlook' {
    $archiveResults = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'preview.pst') -ReportDirectory $TestDrive -DryRun -QuitOutlook
    $dedupResults = & (Join-Path $script:OfficePath 'Optimize-Outlook.ps1') -FolderName '' -DryRun -QuitOutlook
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
    $result = @(& (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'empty.pst') -ReportDirectory $TestDrive -WhatIf -PassThru)
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
      $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'relative.pst') -ReportPath '.\reports\archive-report.json' -WhatIf
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
      FolderName      = ''
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
    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') -FolderName '' -ArchivePath (Join-Path $TestDrive 'failed-report.pst') -ReportDirectory $TestDrive -DryRun -WarningAction SilentlyContinue
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
      FolderName      = ''
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
