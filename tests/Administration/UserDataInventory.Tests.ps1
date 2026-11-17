#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

param ()

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.4.0 -Force
  $script:ScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../scripts/Administration/Show-UserDataInventory.ps1'))
  $script:Content = Get-Content -LiteralPath $script:ScriptPath -Raw
  $script:Ast = [Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$null, [ref]$null)
  $definitions = @($script:Ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false))
  foreach ($definition in $definitions) {
    . ([scriptblock]::Create($definition.Extent.Text))
  }
}

Describe 'Show-UserDataInventory conventions' {
  It 'pins the module and exposes the inventory surface' {
    $script:Content | Should -Match '#Requires -Version 5.1'
    $script:Content | Should -Match "RequiredVersion = '0.4.0'"
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    $script:Content | Should -Match 'Server Core: supported'
    $script:Content | Should -Match 'SYSTEM-account execution: not supported'
    (Get-Help $script:ScriptPath).Synopsis | Should -Match 'Inventories a user'
  }
}

Describe 'Show-UserDataInventory JSON output' {
  It 'keeps quotes and angle brackets readable and round-trips literal escape text' {
    $value = [PSCustomObject]@{
      Call    = ".\Import.ps1 -SourcePath 'C:\Migration\<user>' -Mode A&B"
      Literal = 'C:\u0027\data'
    }

    $json = ConvertTo-DataInventoryJson -InputObject $value -Depth 4

    $json | Should -Not -Match '(?<!\\)\\u00(27|3c|3e|26)'
    $json | Should -Match "-SourcePath 'C:"
    $roundTrip = $json | ConvertFrom-Json
    $roundTrip.Call | Should -BeExactly $value.Call
    $roundTrip.Literal | Should -BeExactly 'C:\u0027\data'
  }
}

Describe 'Show-UserDataInventory helpers' {
  It 'summarizes a data tree' {
    $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path (Join-Path $root 'Sub') -Force
    [IO.File]::WriteAllText((Join-Path $root 'one.txt'), '12345')
    [IO.File]::WriteAllText((Join-Path $root 'Sub\two.txt'), '1234567890')

    $summary = Get-DataInventoryTreeSummary -Path $root

    $summary.Exists | Should -BeTrue
    $summary.Files | Should -Be 2
    $summary.Directories | Should -Be 1
    $summary.Bytes | Should -Be 15
    $summary.Error | Should -BeNullOrEmpty
  }

  It 'resolves URL shortcuts and flags network targets' {
    $local = Join-Path $TestDrive 'local.url'
    $network = Join-Path $TestDrive 'network.url'
    [IO.File]::WriteAllText($local, "[InternetShortcut]`r`nURL=https://example.com/path`r`n")
    [IO.File]::WriteAllText($network, "[InternetShortcut]`r`nURL=\\server\share\file`r`n")

    (Get-DataInventoryShortcut -Path $local).Target | Should -Be 'https://example.com/path'
    (Get-DataInventoryShortcut -Path $local).IsNetwork | Should -BeFalse
    (Get-DataInventoryShortcut -Path $network).IsNetwork | Should -BeTrue
  }

  It 'validates a matching Outlook report and extracts PST files' {
    $reportPath = Join-Path $TestDrive 'Show-OutlookAccount-test.json'
    $report = [ordered]@{
      SchemaVersion  = 2
      Script         = 'Show-OutlookAccount'
      StartedAt      = (Get-Date).ToUniversalTime().ToString('o')
      FinishedAt     = (Get-Date).ToUniversalTime().ToString('o')
      ComputerName   = $env:COMPUTERNAME
      UserName       = 'EXAMPLE\User1'
      OutlookVersion = '16.0'
      Accounts       = @()
      Stores         = @(
        [ordered]@{ DisplayName = 'Mailbox'; FilePath = 'C:\Data\mail.pst'; DataFileFormat = 'PST'; IsExchangeBacked = $false }
        [ordered]@{ DisplayName = 'IMAP'; FilePath = ''; DataFileFormat = 'None'; IsExchangeBacked = $false }
      )
      Findings       = @()
      CriticalCount  = 0
    }
    [IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))

    $result = Get-DataInventoryOutlookReport -Path $reportPath -Computer $env:COMPUTERNAME -UserName 'EXAMPLE\User1' -MaxAgeDays 7 -SignaturePath 'C:\Users\User1\AppData\Roaming\Microsoft\Signatures'

    $result.Valid | Should -BeTrue
    @($result.DataFiles).Count | Should -Be 1
    $result.DataFiles[0].Path | Should -Be 'C:\Data\mail.pst'
    $result.Signature | Should -Match 'Signatures'
  }

  It 'marks another computer, another user, or an old report invalid' {
    $reportPath = Join-Path $TestDrive 'Show-OutlookAccount-other.json'
    $report = [ordered]@{
      SchemaVersion = 2
      Script        = 'Show-OutlookAccount'
      FinishedAt    = (Get-Date).AddDays(-30).ToUniversalTime().ToString('o')
      ComputerName  = 'OTHERPC'
      UserName      = 'EXAMPLE\Other'
      Stores        = @()
      Findings      = @()
    }
    [IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))

    $result = Get-DataInventoryOutlookReport -Path $reportPath -Computer $env:COMPUTERNAME -UserName 'EXAMPLE\User1' -MaxAgeDays 7 -SignaturePath 'C:\Signatures'

    $result.Valid | Should -BeFalse
    ($result.Warnings -join "`n") | Should -Match 'OTHERPC'
    ($result.Warnings -join "`n") | Should -Match 'EXAMPLE\\Other'
    ($result.Warnings -join "`n") | Should -Match 'older than'
  }

  It 'builds a migration plan with folder mappings, manual copies, and decisions' {
    $known = @(
      [PSCustomObject]@{ Name = 'Desktop'; Path = 'C:\Users\User1\Desktop'; Exists = $true; Bytes = 10; IsUnc = $false; IsOneDrive = $false; IsReparse = $false; IsOutsideProfile = $false }
      [PSCustomObject]@{ Name = 'Documents'; Path = '\\server\docs'; Exists = $true; Bytes = 20; IsUnc = $true; IsOneDrive = $false; IsReparse = $false; IsOutsideProfile = $true }
    )
    $public = @(
      [PSCustomObject]@{ Name = 'Desktop'; Path = 'C:\Users\Public\Desktop'; Files = 1; Shortcuts = @() }
    )
    $outside = @(
      [PSCustomObject]@{ Path = 'D:\Business'; Files = 1; Bytes = 30; Programs = @('Example App') }
    )
    $long = @(
      [PSCustomObject]@{ Source = 'D:\Business\long'; Destination = 'C:\Users\User1\Business\long'; Length = 260; Limit = 259; IsDirectory = $false; WordWebPair = $false }
    )
    $outlook = [PSCustomObject]@{
      DataFiles = @([PSCustomObject]@{ Path = 'C:\Data\mail.pst'; Format = 'PST' })
      Signature = 'C:\Users\User1\AppData\Roaming\Microsoft\Signatures'
      Warnings  = @()
      Error     = $null
    }

    $plan = Get-DataInventoryPlan -KnownFolder $known -PublicFolder $public -NetworkDrive @() -OutsideProfile $outside -LongPath $long -Outlook $outlook

    $plan.Folder | Should -Contain 'Desktop'
    $plan.FolderMap.Keys | Should -Contain 'Business'
    $plan.FolderMap.Keys | Should -Contain 'Public Desktop'
    @($plan.ManualCopy | Where-Object Path -EQ 'C:\Data\mail.pst').Count | Should -Be 1
    @($plan.ManualCopy | Where-Object Path -Match 'Signatures').Count | Should -Be 1
    ($plan.Decisions -join "`n") | Should -Match 'too long'
    ($plan.Decisions -join "`n") | Should -Match 'outside the profile'
  }

  It 'carries Outlook report warnings into the plan decisions' {
    $outlook = [PSCustomObject]@{
      DataFiles = @()
      Signature = $null
      Warnings  = @("The Outlook report is from computer 'OTHERPC', not 'THISPC'.")
      Error     = $null
    }

    $plan = Get-DataInventoryPlan -KnownFolder @() -PublicFolder @() -NetworkDrive @() -OutsideProfile @() -LongPath @() -Outlook $outlook

    $plan.Decisions | Should -Contain "The Outlook report is from computer 'OTHERPC', not 'THISPC'."
    @($plan.Decisions | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count | Should -Be 0
  }

  It 'scans only fixed lettered volumes outside the profile' {
    $fixedRoot = Join-Path $TestDrive 'FixedRoot'
    $removableRoot = Join-Path $TestDrive 'RemovableRoot'
    $business = Join-Path $fixedRoot 'Business'
    $null = New-Item -ItemType Directory -Path $business -Force
    $null = New-Item -ItemType Directory -Path (Join-Path $removableRoot 'Camera') -Force
    [IO.File]::WriteAllText((Join-Path $business 'ledger.txt'), '12345')

    # Back the mocked volumes with temporary PowerShell drives on unused letters.
    $usedLetters = @(Get-PSDrive -PSProvider FileSystem | ForEach-Object { $_.Name.ToUpperInvariant() })
    $freeLetters = @([char[]]'QRSTUVWXYZ' | ForEach-Object { [string]$_ } | Where-Object { $_ -notin $usedLetters })
    $script:FixedLetter = $freeLetters[0]
    $script:RemovableLetter = $freeLetters[1]
    $null = New-PSDrive -Name $script:FixedLetter -PSProvider FileSystem -Root $fixedRoot -Scope Global
    $null = New-PSDrive -Name $script:RemovableLetter -PSProvider FileSystem -Root $removableRoot -Scope Global

    Mock Get-Volume {
      @(
        [PSCustomObject]@{ DriveLetter = $script:FixedLetter; DriveType = 'Fixed' }
        [PSCustomObject]@{ DriveLetter = $script:RemovableLetter; DriveType = 'Removable' }
        [PSCustomObject]@{ DriveLetter = $null; DriveType = 'Fixed' }
      )
    }
    Mock Get-ItemProperty { }

    try {
      $outside = @(Get-DataInventoryOutsideProfile -ProfilePath 'C:\Users\User1')

      Should -Invoke Get-Volume -Times 1 -Exactly
      $outside.Count | Should -Be 1
      $outside[0].Path | Should -Be $business
      $outside[0].Files | Should -Be 1
    }
    finally {
      Remove-PSDrive -Name $script:FixedLetter -Scope Global -Force -ErrorAction SilentlyContinue
      Remove-PSDrive -Name $script:RemovableLetter -Scope Global -Force -ErrorAction SilentlyContinue
    }
  }

  It 'detects a destination path over the supported length' {
    $source = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $target = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $source -Force
    $long = Join-Path $source 'Documents'
    for ($index = 0; $index -lt 12; $index++) {
      $long = Join-Path $long ('segment-{0:D2}-{1}' -f $index, ('x' * 24))
    }
    $longFull = '\\?\' + $long
    $null = [IO.Directory]::CreateDirectory($longFull)
    [IO.File]::WriteAllText(($longFull + '\payload.txt'), 'x')
    try {
      $paths = @(Get-DataInventoryLongPath -SourceRoot $source -TargetRoot $target)

      $paths.Count | Should -BeGreaterThan 0
      $paths[0].Length | Should -BeGreaterThan $paths[0].Limit
    }
    finally {
      [IO.Directory]::Delete($longFull, $true)
      Remove-Item -LiteralPath $source -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
