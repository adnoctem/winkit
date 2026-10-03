#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Cross-script mock context is removed after each test; Pester mocks execute in the invoked script scope.')]
param ()

# Exercise real script functions with mutable Outlook-shaped objects. No COM,
# profile, or real mail is opened. Script-level guards are tested with mocks.
BeforeAll {
  Import-Module PSFoundation -Force
  # The PST helper release can lag this consumer branch. Declare only mock
  # boundaries when unavailable; helper implementation tests belong to PSF.
  if (-not (Get-Command Open-OutlookPstStore -ErrorAction SilentlyContinue)) {
    function Open-OutlookPstStore {
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Mock boundary only; no attachment is opened.')]
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'The parameters declare the API consumed by Pester mocks; the stub never opens Outlook.')]
      [CmdletBinding(SupportsShouldProcess = $true)]
      param (
        [object]
        $Namespace,

        [string]
        $LiteralPath
      )

      throw 'Open-OutlookPstStore must be mocked in consumer tests.'
    }
  }
  if (-not (Get-Command Close-OutlookPstStore -ErrorAction SilentlyContinue)) {
    function Close-OutlookPstStore {
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Mock boundary only; no attachment is closed.')]
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'The parameter declares the API consumed by Pester mocks; the stub never closes Outlook.')]
      [CmdletBinding()]
      param (
        [object]
        $Context
      )

      throw 'Close-OutlookPstStore must be mocked in consumer tests.'
    }
  }

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
      SourceItems   = $null
      DownloadState = 1
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
    Mock Get-CimInstance { [PSCustomObject]@{ DriveType = 3 } }
    Mock Import-Module { }
    Mock Write-Log { }
    Mock Write-OperationResultLog { }
    Mock Get-Process { }
    Mock Start-Process { [PSCustomObject]@{ ExitCode = 0 } }
    $script:DataPath = Join-Path $TestDrive 'mail with spaces.pst'
    $script:ToolPath = Join-Path $TestDrive 'SCANPST.EXE'
    [IO.File]::WriteAllText($script:DataPath, 'fake pst')
    [IO.File]::WriteAllText($script:ToolPath, 'fake executable')
    Mock Get-OutlookRepairToolInfo {
      param($LiteralPath)
      [PSCustomObject]@{
        Name                 = 'ScanPST'
        Path                 = $LiteralPath
        FileVersion          = [version]'16.0.10325.20082'
        SupportsFileArgument = $true
      }
    }
  }

  It 'targets supported ScanPST with a quoted file and one rescan without forcing or hiding its UI' {
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    $result.LaunchMode | Should -Be Targeted
    Should -Invoke Start-Process -Times 1 -ParameterFilter { $ArgumentList -eq ('-file "' + $script:DataPath + '" -rescan 1') }
  }

  It 'launches an explicit legacy tool without arguments or a required data-file path' {
    Mock Get-OutlookRepairToolInfo {
      param($LiteralPath)
      [PSCustomObject]@{
        Name                 = 'ScanOST'
        Path                 = $LiteralPath
        FileVersion          = [version]'12.0.6650.5000'
        SupportsFileArgument = $false
      }
    }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -ToolPath $script:ToolPath -PassThru -Confirm:$false
    $result.Status | Should -Be Completed
    $result.LaunchMode | Should -Be Interactive
    Should -Invoke Start-Process -Times 1 -ParameterFilter { -not $ArgumentList }
  }

  It 'previews legacy ScanPST without forwarding the requested path' {
    Mock Get-OutlookRepairToolInfo {
      param($LiteralPath)
      [PSCustomObject]@{
        Name                 = 'ScanPST'
        Path                 = $LiteralPath
        FileVersion          = [version]'12.0.6650.5000'
        SupportsFileArgument = $false
      }
    }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -DryRun
    $result.LaunchMode | Should -Be Interactive
    $result.RequestedPath | Should -Be $script:DataPath
    $result.Detail | Should -Be ('DryRun: "' + $script:ToolPath + '"')
    Should -Invoke Start-Process -Times 0
  }

  It 'requires a data-file path for supported targeting before launching' {
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -ToolPath $script:ToolPath -PassThru -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'Supply Path'
    Should -Invoke Start-Process -Times 0
  }

  It 'rejects UNC targets during preview' {
    Mock Resolve-LongPath { param($LiteralPath) $LiteralPath }
    Mock Resolve-LongPath { '\\server\share\mail.pst' } -ParameterFilter { $LiteralPath -like '*.pst' }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -WhatIf -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'local data file'
    Should -Invoke Start-Process -Times 0
    Should -Invoke Get-CimInstance -Times 0
  }

  It 'rejects extended UNC targets during preview' {
    Mock Resolve-LongPath { param($LiteralPath) $LiteralPath }
    Mock Resolve-LongPath { '\\?\UNC\server\share\mail.pst' } -ParameterFilter { $LiteralPath -like '*.pst' }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -WhatIf -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'local data file'
    Should -Invoke Start-Process -Times 0
    Should -Invoke Get-CimInstance -Times 0
  }

  It 'rejects mapped network drives before launching' {
    Mock Resolve-LongPath { param($LiteralPath) $LiteralPath }
    Mock Resolve-LongPath { 'Z:\mail.pst' } -ParameterFilter { $LiteralPath -like '*.pst' }
    Mock Get-CimInstance { [PSCustomObject]@{ DriveType = 4 } }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'mapped network drives'
    Should -Invoke Get-CimInstance -Times 1 -ParameterFilter { $ClassName -eq 'Win32_LogicalDisk' -and $Filter -eq "DeviceID='Z:'" }
    Should -Invoke Start-Process -Times 0
  }

  It 'does not launch when local storage cannot be verified' {
    Mock Get-CimInstance { $null }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $script:DataPath -ToolPath $script:ToolPath -PassThru -Confirm:$false -WarningAction SilentlyContinue
    $result.Status | Should -Be Failed
    $result.Detail | Should -Match 'verified local storage'
    Should -Invoke Start-Process -Times 0
  }

  It 'discovers only ScanPST even for OST files and honors WhatIf' {
    $ostPath = Join-Path $TestDrive 'cache.ost'
    [IO.File]::WriteAllText($ostPath, 'fake ost')
    Mock Find-OutlookRepairTool {
      [PSCustomObject]@{
        Name                 = 'ScanPST'
        Path                 = 'TestDrive:\SCANPST.EXE'
        FileVersion          = [version]'16.0.10325.20082'
        SupportsFileArgument = $true
      }
    }
    $result = & (Join-Path $script:OfficePath 'Repair-OutlookDataFile.ps1') -Path $ostPath -WhatIf
    $result.Status | Should -Be Skipped
    $result.LaunchMode | Should -Be Targeted
    $result.Detail | Should -Match '\-file.*\-rescan 1'
    Should -Invoke Find-OutlookRepairTool -Times 1 -Exactly -ParameterFilter { $Name -eq 'ScanPST' }
    Should -Invoke Start-Process -Times 0
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
    $sourcePath = Join-Path $TestDrive 'source.pst'
    [IO.File]::WriteAllText($sourcePath, 'source fixture')
    $store = [PSCustomObject]@{
      DisplayName       = 'Duplicate name'
      FilePath          = $sourcePath
      Root              = $script:Source
      StoreID           = 'source-store'
      ExchangeStoreType = 3
    }
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

  Context 'Existing archive destinations' {
    BeforeEach {
      $script:ArchivePath = Join-Path $TestDrive 'existing.pst'
      [IO.File]::WriteAllText($script:ArchivePath, 'existing archive fixture')
      $script:Destination = New-FakeFolder -Name 'User archive name'
      $script:Destination.StoreID = 'archive-store'
      $archiveStore = [PSCustomObject]@{
        DisplayName = $script:Destination.Name
        FilePath    = $script:ArchivePath
        StoreID     = 'archive-store'
        Root        = $script:Destination
      }
      $archiveStore | Add-Member ScriptMethod GetRootFolder { $this.Root }
      $script:FakeContext | Add-Member NoteProperty ArchiveRoot $script:Destination
      $script:FakeContext.Namespace.Stores = New-FakeCollection @($script:FakeContext.Namespace.DefaultStore, $archiveStore)
      $script:FakeContext.Namespace | Add-Member NoteProperty Detached (New-Object Collections.ArrayList)
      $script:FakeContext.Namespace | Add-Member ScriptMethod RemoveStore {
        param($Root)
        $null = $this.Detached.Add($Root.StoreID)
      }
      $script:FakeContext.Namespace | Add-Member ScriptMethod GetItemFromID {
        param($Id, $StoreId)
        $queue = New-Object Collections.Queue
        $queue.Enqueue($this.DefaultStore.Root)
        while ($queue.Count) {
          $folder = $queue.Dequeue()
          foreach ($mail in $folder.Items.Values) {
            if ($mail.EntryID -eq $Id -and $folder.StoreID -eq $StoreId) {
              return $mail
            }
          }
          foreach ($child in $folder.Folders.Values) {
            $queue.Enqueue($child)
          }
        }
        throw 'Item not found'
      }
      $script:AppendArguments = @{
        ArchivePath     = $script:ArchivePath
        Append          = $true
        FolderName      = ''
        ReportDirectory = $TestDrive
        Mode            = 'Move'
        PassThru        = $true
        Confirm         = $false
        WarningAction   = 'SilentlyContinue'
      }
      $script:ArchiveScript = Join-Path $script:OfficePath 'New-OutlookArchive.ps1'
    }

    Context 'PST sources selected by file path' {
      BeforeEach {
        $script:Source | Add-Member NoteProperty Store $script:FakeContext.Namespace.DefaultStore
        $script:PstContext = [PSCustomObject]@{
          Path           = $script:FakeContext.Namespace.DefaultStore.FilePath
          Root           = $script:Source
          Namespace      = $script:FakeContext.Namespace
          StoreId        = $script:Source.StoreID
          DisplayName    = 'Duplicate name'
          AttachedByCall = $true
          Closed         = $false
        }
        $script:FakeContext | Add-Member NoteProperty PstContext $script:PstContext
        Mock Open-OutlookPstStore { $global:WinkitSafetyTestContext.PstContext }
        Mock Close-OutlookPstStore {
          param($Context)
          $Context.Closed = $true
          $Context.Root = $null
        }
        $script:AppendArguments.SourceArchivePath = $script:PstContext.Path
      }

      It 'previews an explicit PST and closes the source without creating a destination' {
        $result = & $script:ArchiveScript @script:AppendArguments -DryRun
        $result.Status | Should -Be Preview
        $result.SourceFilePath | Should -Be $script:PstContext.Path
        $result.DestinationFilePath | Should -Be $script:ArchivePath
        $report = Get-Content -LiteralPath $result.ReportPath -Raw | ConvertFrom-Json
        $report.Settings.SourceArchivePath | Should -Be $script:PstContext.Path
        $report.Settings.SourceAttachmentCreated | Should -BeTrue
        $script:PstContext.Closed | Should -BeTrue
        Should -Invoke Open-OutlookPstStore -Times 1 -ParameterFilter { -not $WhatIf -and -not $Confirm }
        Should -Invoke Close-OutlookPstStore -Times 1
        Should -Invoke Add-OutlookStoreRoot -Times 0
        Should -Invoke Get-OutlookStoreRoot -Times 0
      }

      It 'preserves both the planning failure and a source cleanup failure' {
        Mock Get-OutlookFolderPlan { throw 'Source folder planning failed' }
        Mock Close-OutlookPstStore { throw 'Source attachment cleanup failed' }
        $result = & $script:ArchiveScript @script:AppendArguments -DryRun
        $result.Status | Should -Be Failed
        $report = Get-Content -LiteralPath $result.ReportPath -Raw | ConvertFrom-Json
        @($report.Results | Where-Object Status -EQ Failed).Count | Should -Be 2
        $report.Results.Detail | Should -Contain 'Source folder planning failed'
        $report.Results.Detail | Should -Contain 'Source attachment cleanup failed'
        Should -Invoke Close-OutlookPstStore -Times 1
      }

      It 'refuses the same source and destination before opening Outlook' {
        $script:AppendArguments.ArchivePath = $script:PstContext.Path
        $result = & $script:ArchiveScript @script:AppendArguments -DryRun
        $result.Status | Should -Be Failed
        $result.Detail | Should -BeLike '*must be different files*'
        Should -Invoke Connect-Outlook -Times 0
        Should -Invoke Open-OutlookPstStore -Times 0
      }

      It 'refuses a missing source without creating a report at that path' {
        $script:AppendArguments.SourceArchivePath = Join-Path $TestDrive 'missing-source.pst'
        $script:AppendArguments.Remove('ReportDirectory')
        $result = & $script:ArchiveScript @script:AppendArguments -ReportPath $script:AppendArguments.SourceArchivePath -DryRun
        $result.Status | Should -Be Failed
        Test-Path -LiteralPath $script:AppendArguments.SourceArchivePath | Should -BeFalse
        Should -Invoke Connect-Outlook -Times 0
      }

      It 'refuses competing source selectors during parameter binding' {
        { & $script:ArchiveScript @script:AppendArguments -StoreName 'Mailbox' -DryRun } | Should -Throw
        {
          & (Join-Path $script:OfficePath 'Optimize-Outlook.ps1') -PSTPath $script:PstContext.Path -StoreName 'Mailbox' -DryRun
        } | Should -Throw
        Should -Invoke Connect-Outlook -Times 0
      }

      It 'requires one explicit date range for Split' {
        $command = Get-Command (Join-Path $script:OfficePath 'Split-OutlookArchive.ps1')
        foreach ($parameterSet in $command.ParameterSets) {
          foreach ($name in @('ArchivePath', 'PSTPath', 'StartDate', $parameterSet.Name)) {
            ($parameterSet.Parameters | Where-Object Name -EQ $name).IsMandatory | Should -BeTrue
          }
        }
        {
          & $command -ArchivePath $script:PstContext.Path -PSTPath $script:ArchivePath `
            -StartDate '2018-01-01' -EndDate '2019-01-01' -EndBefore '2019-01-01' -DryRun
        } | Should -Throw
        Should -Invoke Connect-Outlook -Times 0
      }

      It 'explains an unavailable PST helper before opening Outlook' {
        Mock Get-Command { $null } -ParameterFilter { $Name -eq 'Open-OutlookPstStore' }
        $result = & $script:ArchiveScript @script:AppendArguments -DryRun
        $result.Status | Should -Be Failed
        $result.Detail | Should -BeLike '*Update PSFoundation before retrying*'
        Should -Invoke Connect-Outlook -Times 0
      }

      It 'splits exactly the requested range through the archive workflow using <UpperBound>' -ForEach @(
        @{ UpperBound = 'EndBefore'; Expected = 1 }
        @{ UpperBound = 'EndDate'; Expected = 2 }
      ) {
        $before = New-FakeMail -Id 'before'
        $before.ReceivedTime = [datetime]'2017-12-31T23:59:59'
        $start = New-FakeMail -Id 'start'
        $start.ReceivedTime = [datetime]'2018-01-01'
        $end = New-FakeMail -Id 'end'
        $end.ReceivedTime = [datetime]'2019-01-01'
        $script:Source.Items = New-FakeCollection @($before, $start, $end)
        $arguments = @{
          ArchivePath     = $script:PstContext.Path
          PSTPath         = $script:ArchivePath
          Append          = $true
          StartDate       = [datetime]'2018-01-01'
          FolderName      = ''
          Recurse         = $true
          IncludeSentMail = $true
          ExcludeFolders  = @('Protected')
          Sort            = 'OldToNew'
          ReportDirectory = $TestDrive
          DryRun          = $true
        }
        $arguments[$UpperBound] = [datetime]'2019-01-01'
        $result = @(& (Join-Path $script:OfficePath 'Split-OutlookArchive.ps1') @arguments)
        $result.Count | Should -Be 1
        $result[0].Status | Should -Be Preview
        $result[0].Planned | Should -Be $Expected
        $report = Get-Content -LiteralPath $result[0].ReportPath -Raw | ConvertFrom-Json
        $report.Settings.Mode | Should -Be Copy
        $report.Settings.ArchivePath | Should -Be $script:ArchivePath
        $report.Settings.SourceArchivePath | Should -Be $script:PstContext.Path
        $report.Settings.Recurse | Should -BeTrue
        $report.Settings.Include | Should -Contain SentItems
        $report.Settings.Exclusions | Should -Contain Protected
        $report.Settings.Sort | Should -Be OldToNew
        $script:Source.Items.Count | Should -Be 3
        Should -Invoke Close-OutlookPstStore -Times 1
      }

      It 'forwards an omitted FolderName as identity selection and propagates failure from Split' {
        Mock Get-OutlookFolderPlan { throw 'Default Inbox unavailable' }
        $result = & (Join-Path $script:OfficePath 'Split-OutlookArchive.ps1') `
          -ArchivePath $script:PstContext.Path -PSTPath $script:ArchivePath -Append `
          -StartDate '2018-01-01' -EndBefore '2019-01-01' -ReportDirectory $TestDrive -WhatIf -WarningAction SilentlyContinue
        $LASTEXITCODE | Should -Be 1
        $result.Status | Should -Be Failed
        $report = Get-Content -LiteralPath $result.ReportPath -Raw | ConvertFrom-Json
        $report.Settings.FolderSelection | Should -Be DefaultInbox
        $report.Settings.FolderName | Should -BeNullOrEmpty
        Should -Invoke Close-OutlookPstStore -Times 1
      }

      It 'moves only the selected range into an existing PST through Split' {
        $selected = New-FakeMail -Id 'selected'
        $selected.ReceivedTime = [datetime]'2018-06-01'
        $retained = New-FakeMail -Id 'retained'
        $retained.ReceivedTime = [datetime]'2019-01-01'
        $script:Source.Items = New-FakeCollection @($selected, $retained)
        $selected.SourceItems = $script:Source.Items
        $retained.SourceItems = $script:Source.Items
        $result = & (Join-Path $script:OfficePath 'Split-OutlookArchive.ps1') `
          -ArchivePath $script:PstContext.Path -PSTPath $script:ArchivePath -Append `
          -StartDate '2018-01-01' -EndBefore '2019-01-01' -FolderName '' `
          -Mode Move -SkipPathPreservation -AddDataFile -DataFileName 'Year 2018' `
          -ReportDirectory $TestDrive -Confirm:$false -PassThru
        $result.Status | Should -Be Completed
        $result.Moved | Should -Be 1
        $script:Source.Items.Values.EntryID | Should -Be @('retained')
        $script:Destination.Items.Values.EntryID | Should -Be @('selected')
        $script:Destination.Name | Should -Be 'Year 2018'
        $script:PstContext.Closed | Should -BeTrue
        $script:FakeContext.Namespace.Detached.Count | Should -Be 0
      }

      It 'keeps partial source-open cleanup details in the report' {
        Mock Open-OutlookPstStore {
          $failure = New-Object InvalidOperationException('Source root lookup failed')
          $failure.Data['OutlookPstCleanupError'] = 'Temporary source could not be detached'
          throw $failure
        }
        $result = & $script:ArchiveScript @script:AppendArguments -DryRun
        $result.Status | Should -Be Failed
        $result.Detail | Should -BeLike '*Source root lookup failed*Temporary source could not be detached*'
        Should -Invoke Close-OutlookPstStore -Times 0
      }

      It 'previews Optimize on a PST without changing its deduplication workflow' {
        $csvPath = Join-Path $TestDrive 'pst-review.csv'
        $null = & (Join-Path $script:OfficePath 'Optimize-Outlook.ps1') `
          -PSTPath $script:PstContext.Path -FolderName '' -DryRun -ReportPath $csvPath
        $script:PstContext.Closed | Should -BeTrue
        Should -Invoke Open-OutlookPstStore -Times 1 -ParameterFilter { -not $WhatIf -and -not $Confirm }
        Should -Invoke Close-OutlookPstStore -Times 1
        Should -Invoke Get-OutlookSubFolder -Times 0
        Should -Invoke Get-OutlookStoreRoot -Times 0
      }

      It 'reports Optimize source cleanup failures with the PST path' {
        Mock Close-OutlookPstStore { throw 'Source attachment cleanup failed' }
        $results = @(& (Join-Path $script:OfficePath 'Optimize-Outlook.ps1') `
            -PSTPath $script:PstContext.Path -FolderName '' -DryRun -WarningAction SilentlyContinue)
        $failure = @($results | Where-Object Status -EQ Failed)
        $failure.Count | Should -Be 1
        $failure[0].Target | Should -Be $script:PstContext.Path
        $failure[0].Action | Should -Be DetachSourceStore
      }
    }

    Context 'Real PSFoundation PST helper integration' {
      BeforeEach {
        $script:Source | Add-Member NoteProperty Store $script:FakeContext.Namespace.DefaultStore
        $script:FakeContext.Namespace | Add-Member NoteProperty Added 0
        $script:FakeContext.Namespace | Add-Member ScriptMethod GetStoreFromID {
          param($Id)
          foreach ($store in $this.Stores.Values) {
            if ($store.StoreID -eq $Id) {
              return $store
            }
          }
          throw 'Store is not attached'
        }
        $script:FakeContext.Namespace | Add-Member ScriptMethod AddStore {
          param($Path)
          if ($Path -ne $this.DefaultStore.FilePath) {
            throw 'Unexpected source attachment'
          }
          $this.Added++
          $null = $this.Stores.Values.Add($this.DefaultStore)
        }
        $script:FakeContext.Namespace | Add-Member ScriptMethod RemoveStore {
          param($Root)
          $store = $this.GetStoreFromID($Root.StoreID)
          $null = $this.Detached.Add($Root.StoreID)
          $this.Stores.Values.Remove($store)
        } -Force
      }

      It 'runs <ScriptName> preview with real helpers when AlreadyAttached=<AlreadyAttached>' -ForEach @(
        @{ ScriptName = 'New-OutlookArchive'; AlreadyAttached = $true }
        @{ ScriptName = 'New-OutlookArchive'; AlreadyAttached = $false }
        @{ ScriptName = 'Split-OutlookArchive'; AlreadyAttached = $true }
        @{ ScriptName = 'Split-OutlookArchive'; AlreadyAttached = $false }
        @{ ScriptName = 'Optimize-Outlook'; AlreadyAttached = $true }
        @{ ScriptName = 'Optimize-Outlook'; AlreadyAttached = $false }
      ) {
        if (-not (Get-Command Open-OutlookPstStore -Module PSFoundation -ErrorAction SilentlyContinue)) {
          Set-ItResult -Skipped -Because 'The installed PSFoundation does not yet export the PST lifetime helpers.'
          return
        }

        $sourcePath = $script:FakeContext.Namespace.DefaultStore.FilePath
        if (-not $AlreadyAttached) {
          $script:FakeContext.Namespace.Stores.Values.Remove($script:FakeContext.Namespace.DefaultStore)
        }
        $arguments = @{
          FolderName = ''
          WhatIf     = $true
          PassThru   = $true
        }
        if ($ScriptName -eq 'Optimize-Outlook') {
          $arguments.PSTPath = $sourcePath
        }
        else {
          $arguments.Append = $true
          $arguments.ReportDirectory = $TestDrive
          $arguments.Mode = 'Move'
          if ($ScriptName -eq 'Split-OutlookArchive') {
            $arguments.ArchivePath = $sourcePath
            $arguments.PSTPath = $script:ArchivePath
            $arguments.StartDate = [datetime]'2018-01-01'
            $arguments.EndBefore = [datetime]'2019-01-01'
          }
          else {
            $arguments.SourceArchivePath = $sourcePath
            $arguments.ArchivePath = $script:ArchivePath
          }
        }

        $results = @(& (Join-Path $script:OfficePath "$ScriptName.ps1") @arguments)
        @($results | Where-Object Status -EQ Failed).Count | Should -Be 0
        if ($ScriptName -ne 'Optimize-Outlook') {
          $results.Count | Should -Be 1
          $results[0].Status | Should -Be Preview
          $report = Get-Content -LiteralPath $results[0].ReportPath -Raw | ConvertFrom-Json
          $report.Settings.SourceAttachmentCreated | Should -Be (-not $AlreadyAttached)
          $report.Settings.SourceStore.FilePath | Should -Be $sourcePath
        }
        $script:FakeContext.Namespace.Added | Should -Be ([int](-not $AlreadyAttached))
        $script:FakeContext.Namespace.Detached.Count | Should -Be ([int](-not $AlreadyAttached))
        @($script:FakeContext.Namespace.Stores.Values | Where-Object StoreID -EQ 'source-store').Count | Should -Be ([int]$AlreadyAttached)
        $script:FakeContext.Namespace.GetStoreFromID('archive-store').Root.Name | Should -Be 'User archive name'
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
        Should -Invoke Add-OutlookStoreRoot -Times 0
        Should -Invoke Get-OutlookSubFolder -Times 0
      }
    }

    Context 'Synchronized OST sources' {
      BeforeEach {
        $script:FakeContext.App.Version = '16.0'
        $script:FakeContext.Namespace.DefaultStore.FilePath = Join-Path $TestDrive 'source.ost'
        [IO.File]::WriteAllText($script:FakeContext.Namespace.DefaultStore.FilePath, 'source OST fixture')
        Mock Get-OutlookSubFolder {
          param($ParentFolder, $Name, $Create)
          $matchingFolders = @($ParentFolder.Folders.Values | Where-Object Name -EQ $Name)
          if ($matchingFolders.Count) {
            return $matchingFolders[0]
          }
          if ($Create) {
            $folder = New-FakeFolder -Name $Name
            $folder.StoreID = $ParentFolder.StoreID
            $folder.FolderPath = $ParentFolder.FolderPath + '\' + $Name
            $null = $ParentFolder.Folders.Values.Add($folder)
            return $folder
          }
        }
      }

      It 'archives <Provider> mail with <ArchiveMode>, filters, preserved paths and disjoint append passes' -ForEach @(
        @{ Provider = 'IMAP'; ExchangeType = 3; ArchiveMode = 'Copy' }
        @{ Provider = 'Exchange'; ExchangeType = 0; ArchiveMode = 'Copy' }
        @{ Provider = 'IMAP'; ExchangeType = 3; ArchiveMode = 'Move' }
        @{ Provider = 'Exchange'; ExchangeType = 0; ArchiveMode = 'Move' }
      ) {
        $script:FakeContext.Namespace.DefaultStore.ExchangeStoreType = $ExchangeType
        $inbox = New-FakeFolder -Name Posteingang
        $inbox.FolderPath = $script:Source.FolderPath + '\Posteingang'
        $leaf = New-FakeFolder -Name Amazon -Mail @((New-FakeMail 'first'), (New-FakeMail 'second'))
        $leaf.FolderPath = $inbox.FolderPath + '\Amazon'
        $leaf.Items.Item(2).ReceivedTime = [datetime]'2025-06-01'
        $excluded = New-FakeFolder -Name Protected -Mail @((New-FakeMail 'excluded'))
        $excluded.FolderPath = $inbox.FolderPath + '\Protected'
        # An excluded header-only message must not prevent selected mail archiving.
        $excluded.Items.Item(1).DownloadState = 0
        $null = $inbox.Folders.Values.Add($leaf)
        $null = $inbox.Folders.Values.Add($excluded)
        $null = $script:Source.Folders.Values.Add($inbox)
        Mock Get-OutlookStandardFolderIdentity -ModuleName PSFoundation {
          [PSCustomObject]@{
            Kind    = 'Inbox'
            StoreID = 'source-store'
            EntryID = 'folder-Posteingang'
            State   = 'Resolved'
          }
        }
        $script:AppendArguments.Remove('FolderName')
        $script:AppendArguments.Mode = $ArchiveMode
        $script:AppendArguments.Recurse = $true
        $script:AppendArguments.Exclusions = @('Posteingang\Protected')
        $script:AppendArguments.EndBefore = [datetime]'2025-01-01'
        $first = & $script:ArchiveScript @script:AppendArguments
        $first.Status | Should -Be Completed
        ($first.Copied + $first.Moved) | Should -Be 1

        $script:AppendArguments.Remove('EndBefore')
        $script:AppendArguments.StartDate = [datetime]'2025-01-01'
        $second = & $script:ArchiveScript @script:AppendArguments
        $second.Status | Should -Be Completed
        ($second.Copied + $second.Moved) | Should -Be 1
        $leaf.Items.Count | Should -Be $(if ($ArchiveMode -eq 'Copy') { 2 } else { 0 })
        $excluded.Items.Count | Should -Be 1
        $target = $script:Destination.Folders.Item(1).Folders.Item(1)
        $target.Name | Should -Be Amazon
        $target.Items.Count | Should -Be 2
        $report = Get-Content -LiteralPath $second.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $report.Settings.SourceStore.DataFileFormat | Should -Be OST
        $report.Settings.SourceStore.ExchangeStoreType | Should -Be $ExchangeType
        $report.Settings.SourceStore.MaySynchronize | Should -BeTrue
        $report.Settings.FolderSelection | Should -Be DefaultInbox
        $report.Results[0].DestinationFolderPath | Should -Be ($script:ArchivePath + '::\Posteingang\Amazon')
        $report.SourceWarnings -join ' ' | Should -Match 'Server-mailbox completeness is not verified'
        $script:FakeContext.Namespace.Detached.Count | Should -Be 0
      }

      It 'previews <Provider> <ArchiveMode> without transfers and reports synchronization consequences' -ForEach @(
        @{ Provider = 'IMAP'; ExchangeType = 3; ArchiveMode = 'Copy' }
        @{ Provider = 'Exchange'; ExchangeType = 0; ArchiveMode = 'Copy' }
        @{ Provider = 'IMAP'; ExchangeType = 3; ArchiveMode = 'Move' }
        @{ Provider = 'Exchange'; ExchangeType = 0; ArchiveMode = 'Move' }
      ) {
        $script:FakeContext.Namespace.DefaultStore.ExchangeStoreType = $ExchangeType
        $mail = New-FakeMail 'preview'
        $mail.SourceItems = $script:Source.Items
        $null = $script:Source.Items.Values.Add($mail)
        $script:AppendArguments.Mode = $ArchiveMode
        $script:AppendArguments.Append = $false
        $script:AppendArguments.ArchivePath = Join-Path $TestDrive 'new.pst'
        $result = & $script:ArchiveScript @script:AppendArguments -DryRun -WarningVariable warnings
        $result.Status | Should -Be Preview
        $result.Planned | Should -Be 1
        $mail.Copies | Should -Be 0
        $mail.Moves | Should -Be 0
        Test-Path -LiteralPath $script:AppendArguments.ArchivePath | Should -BeFalse
        Should -Invoke Add-OutlookStoreRoot -Times 0
        $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $report.SourceWarnings.Count | Should -Be 2
        $expected = if ($ArchiveMode -eq 'Copy') { 'temporarily creates duplicates' } else { 'Removal can synchronize' }
        ($warnings -join ' ') | Should -Match $expected
        ($report.SourceWarnings -join ' ') | Should -Match $expected
      }

      It 'stops on <Case> download state during enumeration in preview and execution' -ForEach @(
        @{ Case = 'header-only'; State = 0 }
        @{ Case = 'null'; State = $null }
        @{ Case = 'unknown'; State = 99 }
        @{ Case = 'missing'; State = 1 }
        @{ Case = 'unreadable'; State = 1 }
      ) {
        $mail = New-FakeMail 'incomplete'
        $mail.SourceItems = $script:Source.Items
        $mail.DownloadState = $State
        if ($Case -eq 'missing') {
          $mail.PSObject.Properties.Remove('DownloadState')
        }
        elseif ($Case -eq 'unreadable') {
          $mail | Add-Member ScriptProperty DownloadState { throw 'Provider unavailable' } -Force
        }
        $null = $script:Source.Items.Values.Add($mail)
        foreach ($preview in @($true, $false)) {
          $result = & $script:ArchiveScript @script:AppendArguments -DryRun:$preview
          $result.Status | Should -Be Failed
          $result.Detail | Should -Match 'DownloadState'
          $result.Detail | Should -Match 'incomplete'
          $result.Detail | Should -Match ([regex]::Escape($script:Source.FolderPath))
          ($result.Copied + $result.Moved) | Should -Be 0
          $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
          $report.Results[-1].Detail | Should -Match 'DownloadState'
        }
        $mail.Copies | Should -Be 0
        $mail.Moves | Should -Be 0
        $script:Source.Items.Count | Should -Be 1
        $script:Destination.Items.Count | Should -Be 0
      }

      It 'rechecks download state immediately before <ArchiveMode> and retains earlier successful transfers' -ForEach @(
        @{ ArchiveMode = 'Copy' }
        @{ ArchiveMode = 'Move' }
      ) {
        foreach ($id in @('good', 'changed')) {
          $mail = New-FakeMail $id
          $mail.SourceItems = $script:Source.Items
          $null = $script:Source.Items.Values.Add($mail)
        }
        $changed = $script:Source.Items.Item(2)
        $changed | Add-Member NoteProperty DownloadReads 0
        $changed | Add-Member ScriptProperty DownloadState {
          $this.DownloadReads++
          if ($this.DownloadReads -eq 1) { return 1 }
          return 0
        } -Force
        $script:AppendArguments.Mode = $ArchiveMode
        $result = & $script:ArchiveScript @script:AppendArguments
        $result.Status | Should -Be Failed
        ($result.Copied + $result.Moved) | Should -Be 1
        $result.Detail | Should -Match 'changed'
        $changed.Copies | Should -Be 0
        $changed.Moves | Should -Be 0
        $script:Destination.Items.Count | Should -Be 1
        $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
        @($report.Results | Where-Object { $_.Status -in @('Copied', 'Moved') }).Count | Should -Be 1
      }

      It 'retains partial results and a source duplicate when a downloaded OST copy fails to transfer' {
        foreach ($id in @('good', 'bad')) {
          $mail = New-FakeMail $id
          $mail.SourceItems = $script:Source.Items
          $mail.FailMove = $id -eq 'bad'
          $null = $script:Source.Items.Values.Add($mail)
        }
        $script:AppendArguments.Mode = 'Copy'
        $result = & $script:ArchiveScript @script:AppendArguments
        $result.Status | Should -Be Failed
        $result.Copied | Should -Be 1
        $result.Detail | Should -Match 'Disk full'
        $script:Source.Items.Count | Should -Be 3
        $script:Destination.Items.Count | Should -Be 1
      }
    }

    It 'does not require download metadata for a PST source' {
      $mail = New-FakeMail 'local'
      $mail.SourceItems = $script:Source.Items
      $mail | Add-Member ScriptProperty DownloadState { throw 'PST download metadata must not be used' } -Force
      $null = $script:Source.Items.Values.Add($mail)
      $result = & $script:ArchiveScript @script:AppendArguments
      $result.Status | Should -Be Completed
      $result.Moved | Should -Be 1
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $report.Settings.SourceStore.DataFileFormat | Should -Be PST
      $report.Settings.SourceStore.MaySynchronize | Should -BeFalse
      $report.SourceWarnings.Count | Should -Be 0
    }

    It 'captures the named source observations instead of classifying the default store' {
      $script:FakeContext.App.Version = '16.0'
      $script:FakeContext.Namespace.DefaultStore.FilePath = Join-Path $TestDrive 'selected.ost'
      [IO.File]::WriteAllText($script:FakeContext.Namespace.DefaultStore.FilePath, 'source OST fixture')
      $selectedStore = $script:FakeContext.Namespace.DefaultStore
      $script:FakeContext.Namespace.DefaultStore = $script:FakeContext.Namespace.Stores.Item(2)
      $script:AppendArguments.StoreName = $selectedStore.DisplayName
      Mock Get-OutlookStoreRoot {
        $global:WinkitSafetyTestContext.Namespace.Stores.Item(1).Root
      }
      Mock Get-OutlookFolderPlan {
        $root = $global:WinkitSafetyTestContext.Namespace.Stores.Item(1).Root
        [PSCustomObject]@{
          EntryID      = $root.EntryID
          StoreID      = $root.StoreID
          FolderPath   = $root.FolderPath
          RelativePath = ''
          Process      = $false
          Reason       = 'Empty source fixture'
        }
      }
      $result = & $script:ArchiveScript @script:AppendArguments -DryRun
      $result.Status | Should -Be Preview
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $report.Settings.SourceStore.StoreID | Should -Be 'source-store'
      $report.Settings.SourceStore.DataFileFormat | Should -Be OST
      $report.Settings.SourceStore.ExchangeStoreType | Should -Be 3
      $report.Settings.SourceStore.MaySynchronize | Should -BeTrue
    }

    It 'requires Append for an existing file and refuses a missing Append target' {
      $script:AppendArguments.Append = $false
      $result = & $script:ArchiveScript @script:AppendArguments
      $result.Status | Should -Be Failed
      $result.Detail | Should -Match 'already exists'
      $script:AppendArguments.Append = $true
      $script:AppendArguments.ArchivePath = Join-Path $TestDrive 'absent.pst'
      $result = & $script:ArchiveScript @script:AppendArguments
      $result.Status | Should -Be Failed
      $result.Detail | Should -Match 'existing PST'
      Test-Path -LiteralPath $script:AppendArguments.ArchivePath | Should -BeFalse
      Should -Invoke Connect-Outlook -Times 0
      Should -Invoke Add-OutlookStoreRoot -Times 0
    }

    It 'refuses a source store destination without renaming or detaching it' {
      $script:AppendArguments.ArchivePath = $script:FakeContext.Namespace.DefaultStore.FilePath
      $result = & $script:ArchiveScript @script:AppendArguments -DisplayName Changed
      $result.Status | Should -Be Failed
      $result.Detail | Should -Match 'different stores'
      $script:Source.Name | Should -Be Root
      $script:FakeContext.Namespace.Detached.Count | Should -Be 0
      Should -Invoke Add-OutlookStoreRoot -Times 0
    }

    It 'preserves an existing attachment and name across successive approved Move passes' {
      foreach ($id in @('first-pass', 'second-pass')) {
        $mail = New-FakeMail -Id $id
        $mail.SourceItems = $script:Source.Items
        $null = $script:Source.Items.Values.Add($mail)
        $result = & $script:ArchiveScript @script:AppendArguments
        $result.Status | Should -Be Completed
        $result.Moved | Should -Be 1
      }
      $script:Destination.Items.Count | Should -Be 2
      $script:Destination.Name | Should -Be 'User archive name'
      $script:FakeContext.Namespace.Detached.Count | Should -Be 0
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $report.Settings.ArchiveAttachedInitially | Should -BeTrue
      $report.Settings.AttachmentCreated | Should -BeFalse
      $report.Settings.Append | Should -BeTrue
      $report.Results.SourceEntryID | Should -Be 'second-pass'
      Should -Invoke Add-OutlookStoreRoot -Times 0
    }

    It 'renames an existing store only when explicitly requested and refuses explicit detachment' {
      $result = & $script:ArchiveScript @script:AppendArguments -DataFileName 'Chosen name'
      $result.Status | Should -Be Completed
      $script:Destination.Name | Should -Be 'Chosen name'
      $result = & $script:ArchiveScript @script:AppendArguments -DetachWhenDone:$true
      $result.Status | Should -Be Failed
      $result.Detail | Should -Match 'already attached'
      $script:FakeContext.Namespace.Detached.Count | Should -Be 0
    }

    It 'previews detached archives without mounting, renaming, or changing files' {
      $script:FakeContext.Namespace.Stores = New-FakeCollection @($script:FakeContext.Namespace.DefaultStore)
      $result = & $script:ArchiveScript @script:AppendArguments -DryRun -DisplayName Changed
      $result.Status | Should -Be Preview
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $report.Settings.DestinationValidation | Should -Be FileOnly
      $script:Destination.Name | Should -Be 'User archive name'
      [IO.File]::ReadAllText($script:ArchivePath) | Should -Be 'existing archive fixture'
      Should -Invoke Add-OutlookStoreRoot -Times 0
    }

    It 'detaches only a new attachment unless AddDataFile is requested (<KeepAttached>)' -ForEach @(
      @{ KeepAttached = $false }
      @{ KeepAttached = $true }
    ) {
      $script:FakeContext.Namespace.Stores = New-FakeCollection @($script:FakeContext.Namespace.DefaultStore)
      Mock Add-OutlookStoreRoot { $global:WinkitSafetyTestContext.ArchiveRoot }
      $result = & $script:ArchiveScript @script:AppendArguments -AddDataFile:$KeepAttached
      $result.Status | Should -Be Completed
      $script:Destination.Name | Should -Be 'User archive name'
      $script:FakeContext.Namespace.Detached.Count | Should -Be ([int](-not $KeepAttached))
      Should -Invoke Add-OutlookStoreRoot -Times 1
    }

    It 'preserves attachment ownership when a path alias resolves to an already attached store' {
      $script:FakeContext.Namespace.Stores.Item(2).FilePath = $script:FakeContext.Namespace.DefaultStore.FilePath
      Mock Add-OutlookStoreRoot { $global:WinkitSafetyTestContext.ArchiveRoot }
      $result = & $script:ArchiveScript @script:AppendArguments
      $result.Status | Should -Be Completed
      $script:FakeContext.Namespace.Detached.Count | Should -Be 0
      Should -Invoke Add-OutlookStoreRoot -Times 1
    }

    It 'rejects a late source-store identity match before detachment or message changes' {
      $script:FakeContext.Namespace.Stores = New-FakeCollection @($script:FakeContext.Namespace.DefaultStore)
      Mock Add-OutlookStoreRoot { $global:WinkitSafetyTestContext.Namespace.DefaultStore.Root }
      $result = & $script:ArchiveScript @script:AppendArguments
      $result.Status | Should -Be Failed
      $result.Detail | Should -Match 'different stores'
      $script:FakeContext.Namespace.Detached.Count | Should -Be 0
    }

    It 'stops on a <Conflict> destination folder without transferring messages' -ForEach @(
      @{ Conflict = 'duplicate name' }
      @{ Conflict = 'non-mail' }
      @{ Conflict = 'search' }
    ) {
      $inbox = New-FakeFolder -Name Posteingang -Mail @((New-FakeMail 'mail'))
      $inbox.FolderPath = $script:Source.FolderPath + '\Posteingang'
      $null = $script:Source.Folders.Values.Add($inbox)
      $folder = New-FakeFolder -Name Posteingang
      $null = $script:Destination.Folders.Values.Add($folder)
      if ($Conflict -eq 'duplicate name') {
        $null = $script:Destination.Folders.Values.Add((New-FakeFolder -Name Posteingang))
      }
      elseif ($Conflict -eq 'non-mail') {
        $folder.DefaultItemType = 1
      }
      else {
        $folder.PropertyAccessor.FolderType = 2
      }
      Mock Get-OutlookSubFolder {
        param($ParentFolder, $Name)
        @($ParentFolder.Folders.Values | Where-Object Name -EQ $Name)[0]
      }
      $script:AppendArguments.FolderName = 'Posteingang'
      $result = & $script:ArchiveScript @script:AppendArguments
      $result.Status | Should -Be Failed
      $inbox.Items.Count | Should -Be 1
      $script:FakeContext.Namespace.Detached.Count | Should -Be 0
    }

    It 'keeps successful transfers and the existing file after an append failure (<InitiallyAttached>)' -ForEach @(
      @{ InitiallyAttached = $true }
      @{ InitiallyAttached = $false }
    ) {
      foreach ($id in @('good', 'bad')) {
        $mail = New-FakeMail -Id $id
        $mail.SourceItems = $script:Source.Items
        $mail.FailMove = $id -eq 'bad'
        $null = $script:Source.Items.Values.Add($mail)
      }
      if (-not $InitiallyAttached) {
        $script:FakeContext.Namespace.Stores = New-FakeCollection @($script:FakeContext.Namespace.DefaultStore)
        Mock Add-OutlookStoreRoot { $global:WinkitSafetyTestContext.ArchiveRoot }
      }
      $result = & $script:ArchiveScript @script:AppendArguments
      $result.Status | Should -Be Failed
      $result.Moved | Should -Be 1
      $script:Destination.Items.Count | Should -Be 1
      $script:Source.Items.Count | Should -Be 1
      $script:FakeContext.Namespace.Detached.Count | Should -Be ([int](-not $InitiallyAttached))
      Test-Path -LiteralPath $script:ArchivePath | Should -BeTrue
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      @($report.Results | Where-Object Status -EQ Moved).Count | Should -Be 1
    }

    It 'creates missing destination paths and reuses them on the next append' {
      $inbox = New-FakeFolder -Name Posteingang
      $inbox.FolderPath = $script:Source.FolderPath + '\Posteingang'
      $leaf = New-FakeFolder -Name Amazon
      $leaf.FolderPath = $inbox.FolderPath + '\Amazon'
      $null = $inbox.Folders.Values.Add($leaf)
      $null = $script:Source.Folders.Values.Add($inbox)
      Mock Get-OutlookSubFolder {
        param($ParentFolder, $Name, $Create)
        $matchingFolders = @($ParentFolder.Folders.Values | Where-Object Name -EQ $Name)
        if ($matchingFolders.Count) {
          return $matchingFolders[0]
        }
        if ($Create) {
          $folder = New-FakeFolder -Name $Name
          $folder.StoreID = $ParentFolder.StoreID
          $folder.FolderPath = $ParentFolder.FolderPath + '\' + $Name
          $null = $ParentFolder.Folders.Values.Add($folder)
          return $folder
        }
      }
      $script:AppendArguments.FolderName = 'Posteingang\Amazon'
      foreach ($id in @('one', 'two')) {
        $mail = New-FakeMail -Id $id
        $mail.SourceItems = $leaf.Items
        $null = $leaf.Items.Values.Add($mail)
        $result = & $script:ArchiveScript @script:AppendArguments
        $result.Status | Should -Be Completed
      }
      $script:Destination.Folders.Count | Should -Be 1
      $targetInbox = $script:Destination.Folders.Item(1)
      $targetInbox.Name | Should -Be Posteingang
      $targetInbox.Folders.Count | Should -Be 1
      $targetInbox.Folders.Item(1).Items.Count | Should -Be 2
    }

    It 'uses preserved or flattened destinations in real transfers and previews (<Flatten>, <Preview>)' -ForEach @(
      @{ Flatten = $false; Preview = $false }
      @{ Flatten = $true; Preview = $false }
      @{ Flatten = $false; Preview = $true }
      @{ Flatten = $true; Preview = $true }
    ) {
      $inbox = New-FakeFolder -Name Posteingang
      $inbox.FolderPath = $script:Source.FolderPath + '\Posteingang'
      $leaf = New-FakeFolder -Name Amazon -Mail @((New-FakeMail 'order'))
      $leaf.FolderPath = $inbox.FolderPath + '\Amazon'
      $null = $inbox.Folders.Values.Add($leaf)
      $null = $script:Source.Folders.Values.Add($inbox)
      $destinationInbox = New-FakeFolder -Name Posteingang
      $destinationInbox.StoreID = 'archive-store'
      $destinationLeaf = New-FakeFolder -Name Amazon
      $destinationLeaf.StoreID = 'archive-store'
      $null = $destinationInbox.Folders.Values.Add($destinationLeaf)
      $null = $script:Destination.Folders.Values.Add($destinationInbox)
      Mock Get-OutlookSubFolder {
        param($ParentFolder, $Name)
        @($ParentFolder.Folders.Values | Where-Object Name -EQ $Name)[0]
      }
      $script:AppendArguments.FolderName = 'Posteingang'
      $result = & $script:ArchiveScript @script:AppendArguments -Recurse -SkipPathPreservation:$Flatten -DryRun:$Preview
      $result.Status | Should -Be $(if ($Preview) { 'Preview' } else { 'Completed' })
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $expectedPath = $script:ArchivePath + '::\'
      if (-not $Flatten) {
        $expectedPath += 'Posteingang\Amazon'
      }
      $report.Results.DestinationFolderPath | Should -Be $expectedPath
      $report.Results.SourceFolderPath | Should -Be $leaf.FolderPath
      $report.Settings.SkipPathPreservation | Should -Be $Flatten
      if ($Preview) {
        $leaf.Items.Count | Should -Be 1
        Should -Invoke Get-OutlookSubFolder -Times 0 -ParameterFilter { $Create }
      }
      elseif ($Flatten) {
        $script:Destination.Items.Count | Should -Be 1
        $destinationLeaf.Items.Count | Should -Be 0
      }
      else {
        $script:Destination.Items.Count | Should -Be 0
        $destinationLeaf.Items.Count | Should -Be 1
      }
    }
  }

  AfterEach {
    Remove-Variable -Name WinkitSafetyTestContext -Scope Global -ErrorAction SilentlyContinue
  }

  It 'applies the <Group> groups without widening scope for <ScriptName>' -ForEach @(
    @{ ScriptName = 'New-OutlookArchive'; Group = 'default'; Media = $false; Failures = $false }
    @{ ScriptName = 'Optimize-Outlook'; Group = 'default'; Media = $false; Failures = $false }
    @{ ScriptName = 'New-OutlookArchive'; Group = 'Media'; Media = $true; Failures = $false }
    @{ ScriptName = 'Optimize-Outlook'; Group = 'Media'; Media = $true; Failures = $false }
    @{ ScriptName = 'New-OutlookArchive'; Group = 'Failures'; Media = $false; Failures = $true }
    @{ ScriptName = 'Optimize-Outlook'; Group = 'Failures'; Media = $false; Failures = $true }
    @{ ScriptName = 'New-OutlookArchive'; Group = 'both'; Media = $true; Failures = $true }
    @{ ScriptName = 'Optimize-Outlook'; Group = 'both'; Media = $true; Failures = $true }
  ) {
    $mediaKinds = @('Calendar', 'Contacts', 'Journal', 'Notes', 'Tasks', 'AllPublicFolders', 'RssFeeds', 'ToDo', 'ManagedEmail', 'SuggestedContacts')
    $failureKinds = @('Conflicts', 'SyncIssues', 'LocalFailures', 'ServerFailures')
    $identities = @()
    foreach ($kind in @('Inbox', 'SentItems', 'DeletedItems', 'Junk', 'Drafts', 'Outbox') + $mediaKinds + $failureKinds) {
      $folder = New-FakeFolder -Name "Localized-$kind" -Mail @((New-FakeMail -Id $kind))
      $folder.FolderPath = $script:Source.FolderPath + '\' + $folder.Name
      $null = $script:Source.Folders.Values.Add($folder)
      $identities += [PSCustomObject]@{
        Kind    = $kind
        EntryID = $folder.EntryID
        StoreID = 'source-store'
        State   = 'Resolved'
      }
    }
    $script:FakeContext | Add-Member NoteProperty Identities $identities
    Mock Get-OutlookStandardFolderIdentity -ModuleName PSFoundation { $global:WinkitSafetyTestContext.Identities }
    $script:FakeContext.Namespace | Add-Member ScriptMethod GetItemFromID {
      param($Id, $StoreId)
      foreach ($folder in $this.DefaultStore.Root.Folders.Values) {
        foreach ($mail in $folder.Items.Values) {
          if ($mail.EntryID -eq $Id -and $folder.StoreID -eq $StoreId) {
            return $mail
          }
        }
      }
      throw 'Item not found'
    }
    $arguments = @{
      FolderName      = ''
      Recurse         = $true
      IncludeMedia    = $Media
      IncludeFailures = $Failures
      ExcludeFolders  = @('Localized-RssFeeds', 'Localized-LocalFailures')
      DryRun          = $true
    }
    if ($ScriptName -eq 'New-OutlookArchive') {
      $arguments.ArchivePath = Join-Path $TestDrive 'groups.pst'
      $arguments.ReportDirectory = $TestDrive
    }
    $result = & (Join-Path $script:OfficePath "$ScriptName.ps1") @arguments
    if ($ScriptName -eq 'New-OutlookArchive') {
      $result.Status | Should -Be Preview
      $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
      $rows = @($report.Results)
      $reasons = @($report.FolderPlan | ForEach-Object { $_.Reason })
    }
    else {
      $rows = @($result | Where-Object Action -NE SelectFolder)
      $reasons = @($result | Where-Object Action -EQ SelectFolder | ForEach-Object { $_.Detail })
    }
    $expected = @('Inbox')
    if ($Media) {
      $expected += @($mediaKinds | Where-Object { $_ -ne 'RssFeeds' })
    }
    if ($Failures) {
      $expected += @($failureKinds | Where-Object { $_ -ne 'LocalFailures' })
    }
    @($rows.Target | Sort-Object) | Should -Be @($expected | Sort-Object)
    ($reasons -join ';') | Should -Not -Match 'IncludeCalendar|IncludeConflicts'
    Should -Invoke Add-OutlookStoreRoot -Times 0
    Should -Invoke Get-OutlookSubFolder -Times 0 -ParameterFilter { $Create }
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
      IncludeSentItems = $true
      IncludeJunk      = $true
      Exclusions       = @('Unerwuenscht', 'Posteingang')
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
    @($optimizer | Where-Object Detail -EQ CustomExclusion).Count | Should -Be 2
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

  It 'sorts archive dates <Order> while preserving ties and putting failures last' -ForEach @(
    @{ Order = 'OldToNew'; Expected = @('oldest', 'tie-first', 'tie-second', 'newest') }
    @{ Order = 'NewToOld'; Expected = @('newest', 'tie-first', 'tie-second', 'oldest') }
    @{ Order = 'Default'; Expected = @('newest', 'tie-first', 'tie-second', 'oldest') }
  ) {
    $messages = @(
      @{ Id = 'tie-first'; Date = '2024-06-01' }
      @{ Id = 'oldest'; Date = '2024-01-01' }
      @{ Id = 'tie-second'; Date = '2024-06-01' }
      @{ Id = 'newest'; Date = '2025-01-01' }
    )
    foreach ($message in $messages) {
      $mail = New-FakeMail -Id $message.Id
      $mail.ReceivedTime = [datetime]$message.Date
      $null = $script:Source.Items.Values.Add($mail)
    }
    $script:FakeContext.Namespace | Add-Member ScriptMethod GetItemFromID {
      param($Id, $StoreId)
      if ($StoreId -ne $this.DefaultStore.Root.StoreID) {
        throw 'Wrong store identifier'
      }
      @($this.DefaultStore.Root.Items.Values | Where-Object EntryID -EQ $Id)[0]
    }

    # A later folder fails after the first folder's results have been recorded.
    $broken = New-FakeMail -Id 'unreadable-date'
    $broken.ReceivedTime = $null
    $child = New-FakeFolder -Name Later -Mail @($broken)
    $child.FolderPath = $script:Source.FolderPath + '\Later'
    $null = $script:Source.Folders.Values.Add($child)
    $arguments = @{
      ArchivePath     = Join-Path $TestDrive 'sorted.pst'
      ReportDirectory = $TestDrive
      StartDate       = '2024-01-01'
      FolderName      = ''
      Recurse         = $true
      DryRun          = $true
      WarningAction   = 'SilentlyContinue'
    }
    if ($Order -ne 'Default') {
      $arguments.Sort = $Order
    }

    $result = & (Join-Path $script:OfficePath 'New-OutlookArchive.ps1') @arguments
    $report = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $report.Results.Count | Should -Be 5
    @($report.Results[0..3].Target) | Should -Be $Expected
    $report.Results[-1].Status | Should -Be Failed
    $report.Summary.Planned | Should -Be 4
    $report.Settings.Sort | Should -Be $(if ($Order -eq 'Default') { 'NewToOld' } else { $Order })
    @($script:Source.Items.Values.EntryID) | Should -Be @('tie-first', 'oldest', 'tie-second', 'newest')
    @($script:Source.Items.Values | Where-Object { $_.Copies -or $_.Moves }).Count | Should -Be 0
    Should -Invoke Add-OutlookStoreRoot -Times 0
  }

  It 'sorts optimizer CSV, log, and output <Order> without changing duplicate selection' -ForEach @(
    @{ Order = 'OldToNew'; Expected = @('oldest', 'tie-second', 'tie-first', 'newest') }
    @{ Order = 'NewToOld'; Expected = @('newest', 'tie-second', 'tie-first', 'oldest') }
    @{ Order = 'Default'; Expected = @('newest', 'tie-second', 'tie-first', 'oldest') }
  ) {
    Mock Get-MessageId {
      param($Item)
      $Item.MessageId
    }

    $messages = @(
      @{ Id = 'tie-first'; Date = '2024-06-01' }
      @{ Id = 'oldest'; Date = '2024-01-01' }
      @{ Id = 'tie-second'; Date = '2024-06-01' }
      @{ Id = 'newest'; Date = '2025-01-01' }
    )
    foreach ($message in $messages) {
      $mail = New-FakeMail -Id $message.Id
      $mail.ReceivedTime = [datetime]$message.Date
      $null = $script:Source.Items.Values.Add($mail)
    }
    $child = New-FakeFolder -Name Excluded
    $child.FolderPath = $script:Source.FolderPath + '\Excluded'
    $null = $script:Source.Folders.Values.Add($child)
    $arguments = @{
      FolderName = ''
      Recurse    = $true
      Exclusions = @('Excluded')
      ReportPath = Join-Path $TestDrive ('sorted-' + $Order + '.csv')
      DryRun     = $true
    }
    if ($Order -ne 'Default') {
      $arguments.Sort = $Order
    }

    $result = @(& (Join-Path $script:OfficePath 'Optimize-Outlook.ps1') @arguments)
    $result.Count | Should -Be 5
    @($result[0..3].Target) | Should -Be $Expected
    $result[-1].Action | Should -Be SelectFolder
    ($result | Where-Object Status -EQ Kept).Target | Should -Be newest
    $csv = @(Import-Csv -LiteralPath $arguments.ReportPath)
    @($csv.Target) | Should -Be @($result.Target)
    $script:ExpectedSortedTargets = @($result.Target)
    Should -Invoke Write-OperationResultLog -Times 1 -Exactly -ParameterFilter {
      (@($Results.Target) -join '|') -eq ($script:ExpectedSortedTargets -join '|')
    }
    @($script:Source.Items.Values | Where-Object { $_.Copies -or $_.Moves }).Count | Should -Be 0
    Should -Invoke Get-OutlookSubFolder -Times 0 -ParameterFilter { $Create }
  }

  It 'sorts synthetic-message reports <Order> using planned received dates' -ForEach @(
    @{ Order = 'OldToNew'; First = 'Winkit synthetic 1 (seed 1)' }
    @{ Order = 'NewToOld'; First = 'Winkit synthetic 3 (seed 1)' }
    @{ Order = 'Default'; First = 'Winkit synthetic 3 (seed 1)' }
  ) {
    $arguments = @{
      Count     = 3
      StartDate = '2024-01-01'
      EndDate   = '2024-12-31'
      DryRun    = $true
    }
    if ($Order -ne 'Default') {
      $arguments.Sort = $Order
    }

    $result = @(& (Join-Path $script:OfficePath 'New-TestOutlookMessage.ps1') @arguments)
    $result.Count | Should -Be 3
    $result[0].Target | Should -Be $First
    $script:ExpectedSortedTargets = @($result.Target)
    Should -Invoke Write-OperationResultLog -Times 1 -Exactly -ParameterFilter {
      (@($Results.Target) -join '|') -eq ($script:ExpectedSortedTargets -join '|')
    }
    Should -Invoke Get-OutlookSubFolder -Times 0
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
