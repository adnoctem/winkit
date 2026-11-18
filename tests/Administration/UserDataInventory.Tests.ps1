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

  It 'does not count Explorer metadata as data' {
    $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $root -Force
    [IO.File]::WriteAllText((Join-Path $root 'desktop.ini'), '[.ShellClassInfo]')
    [IO.File]::WriteAllText((Join-Path $root 'Thumbs.db'), 'cache')

    $summary = Get-DataInventoryTreeSummary -Path $root

    $summary.Files | Should -Be 0
    $summary.Bytes | Should -Be 0
  }

  It 'writes the newest change as ISO 8601 text' {
    $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $root -Force
    [IO.File]::WriteAllText((Join-Path $root 'one.txt'), 'x')

    $summary = Get-DataInventoryTreeSummary -Path $root
    $json = ConvertTo-DataInventoryJson -InputObject $summary -Depth 4

    $summary.Newest | Should -BeOfType [string]
    [DateTime]::Parse($summary.Newest, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind) |
      Should -BeOfType [DateTime]
    $json | Should -Not -Match '/Date\('
  }

  It 'lists reparse points with their targets without following them' {
    $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $outside = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $root, $outside -Force
    [IO.File]::WriteAllText((Join-Path $outside 'elsewhere.txt'), 'not counted')
    [IO.File]::WriteAllText((Join-Path $root 'own.txt'), 'counted')
    $link = Join-Path $root 'linked'
    $null = New-Item -ItemType Junction -Path $link -Target $outside
    try {
      $summary = Get-DataInventoryTreeSummary -Path $root

      $summary.Files | Should -Be 1
      $summary.ReparsePoints | Should -Be 1
      $summary.ReparseItems[0].Path | Should -Be $link
      $summary.ReparseItems[0].LinkType | Should -Be 'Junction'
      $summary.ReparseItems[0].Target | Should -Match ([regex]::Escape((Split-Path $outside -Leaf)))
    }
    finally {
      [IO.Directory]::Delete($link)
    }
  }

  It 'counts a known folder stored inside another known folder only once' {
    $profileRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $music = Join-Path $profileRoot 'Music'
    $desktop = Join-Path $music 'Desktop'
    $null = New-Item -ItemType Directory -Path $desktop -Force
    [IO.File]::WriteAllText((Join-Path $music 'song.txt'), 'music')
    [IO.File]::WriteAllText((Join-Path $desktop 'one.txt'), 'desk')
    [IO.File]::WriteAllText((Join-Path $desktop 'two.txt'), 'desk')
    $script:ShellFolders = [PSCustomObject]@{
      'My Music' = $music
      Desktop    = $desktop
    }
    Mock Get-ItemProperty { $script:ShellFolders } -ParameterFilter { $LiteralPath -like '*User Shell Folders' }

    $known = @(Get-DataInventoryKnownFolder -RegistryRoot 'HKCU:' -ProfilePath $profileRoot)

    $musicEntry = @($known | Where-Object Name -EQ 'Music')[0]
    $desktopEntry = @($known | Where-Object Name -EQ 'Desktop')[0]
    $musicEntry.Files | Should -Be 1
    $musicEntry.NestedFolders | Should -Be @('Desktop')
    $desktopEntry.Files | Should -Be 2
    @($desktopEntry.NestedFolders).Count | Should -Be 0
  }
}

Describe 'Show-UserDataInventory mapping' {
  BeforeAll {
    function Get-KnownFolderFixture {
      param (
        [string]
        $Name,

        [string]
        $Path,

        [switch]
        $Outside,

        [switch]
        $Unc
      )

      return [PSCustomObject]@{
        Name             = $Name
        Path             = $Path
        Exists           = $true
        Files            = 2
        Bytes            = 20
        IsUnc            = [bool]$Unc
        IsOutsideProfile = [bool]$Outside
        IsOneDrive       = $false
        IsReparse        = $false
      }
    }

    # Mirrors a real profile: Desktop redirected into Music, Searches inside Favorites, Downloads on C:\.
    $script:Known = @(
      Get-KnownFolderFixture -Name 'Desktop' -Path 'C:\Users\User1\Music\Desktop'
      Get-KnownFolderFixture -Name 'Documents' -Path 'C:\Users\User1\Documents'
      Get-KnownFolderFixture -Name 'Downloads' -Path 'C:\Downloads' -Outside
      Get-KnownFolderFixture -Name 'Music' -Path 'C:\Users\User1\Music'
      Get-KnownFolderFixture -Name 'Favorites' -Path 'C:\Users\User1\Favorites'
      Get-KnownFolderFixture -Name 'Searches' -Path 'C:\Users\User1\Favorites\Searches'
      Get-KnownFolderFixture -Name 'Videos' -Path '\\server\videos' -Outside -Unc
    )
    $script:Public = @(
      [PSCustomObject]@{ Name = 'Desktop'; Path = 'C:\Users\Public\Desktop'; Files = 9; Shortcuts = @() }
      [PSCustomObject]@{ Name = 'Documents'; Path = 'C:\Users\Public\Documents'; Files = 5; Shortcuts = @() }
      [PSCustomObject]@{ Name = 'Pictures'; Path = 'C:\Users\Public\Pictures'; Files = 0; Shortcuts = @() }
    )
    $script:Outside = @(
      [PSCustomObject]@{ Path = 'C:\Downloads'; Files = 2; Bytes = 20; Programs = @(); Included = $false }
      [PSCustomObject]@{ Path = 'C:\asfibu'; Files = 12; Bytes = 2MB; Programs = @('Example Accounting'); Included = $false }
      [PSCustomObject]@{ Path = 'D:\Business'; Files = 3; Bytes = 30; Programs = @(); Included = $true }
    )
    $script:Mapping = Get-DataInventoryMapping `
      -KnownFolder $script:Known `
      -PublicFolder $script:Public `
      -OutsideProfile $script:Outside `
      -ProfilePath 'C:\Users\User1' `
      -PublicRoot 'C:\Users\Public'

    function Get-MappingTarget {
      param (
        [string]
        $Source
      )

      return @($script:Mapping.Mappings | Where-Object Source -EQ $Source | ForEach-Object Target)
    }
  }

  It 'maps known folders from where they are stored into their standard location' {
    Get-MappingTarget -Source 'User1\Music\Desktop' | Should -Be 'Desktop'
    Get-MappingTarget -Source 'User1\Music' | Should -Be 'Music'
    Get-MappingTarget -Source 'User1\Favorites\Searches' | Should -Be 'Searches'
    Get-MappingTarget -Source 'User1\Documents' | Should -Be 'Documents'
    Get-MappingTarget -Source 'Downloads' | Should -Be 'Downloads'
    @($script:Mapping.Mappings | Where-Object Target -EQ 'Videos').Count | Should -Be 0
    ($script:Mapping.Decisions -join "`n") | Should -Match 'Videos is redirected to a network'
  }

  It 'explains nested and redirected folders in notes' {
    $notes = $script:Mapping.Notes -join "`n"

    $notes | Should -Match 'Desktop is stored inside Music'
    $notes | Should -Match 'left out of the Music import'
    $notes | Should -Match 'Searches is stored inside Favorites'
    $notes | Should -Match 'Downloads is redirected outside the profile to C:\\Downloads'
  }

  It 'maps public data below the user folders and lists the public desktop for a manual copy' {
    Get-MappingTarget -Source 'Public\Documents' | Should -Be 'Documents\Public Documents'
    @($script:Mapping.Mappings | Where-Object Source -Like 'Public\Desktop*').Count | Should -Be 0
    @($script:Mapping.Mappings | Where-Object Source -Like 'Public\Pictures*').Count | Should -Be 0
    @($script:Mapping.ManualCopy | Where-Object Path -EQ 'C:\Users\Public\Desktop').Count | Should -Be 1
  }

  It 'maps only included outside folders and lists the rest as decisions' {
    Get-MappingTarget -Source 'Business' | Should -Be 'Documents\Business'
    @($script:Mapping.Mappings | Where-Object Origin -EQ 'C:\asfibu').Count | Should -Be 0
    ($script:Mapping.Decisions -join "`n") | Should -Match "-Include 'C:\\asfibu'"
    ($script:Mapping.Decisions -join "`n") | Should -Match 'Example Accounting'
    ($script:Mapping.Decisions -join "`n") | Should -Not -Match 'C:\\Downloads'
  }

  It 'lists each download root once with a unique staging name' {
    $roots = @($script:Mapping.Roots | ForEach-Object Path)

    $roots | Should -Be @('C:\Users\User1', 'C:\Downloads', 'C:\Users\Public', 'D:\Business')
    $collision = Get-DataInventoryMapping `
      -KnownFolder @(Get-KnownFolderFixture -Name 'Downloads' -Path 'C:\Downloads' -Outside) `
      -PublicFolder @() `
      -OutsideProfile @([PSCustomObject]@{ Path = 'D:\Downloads'; Files = 1; Bytes = 1; Programs = @(); Included = $true }) `
      -ProfilePath 'C:\Users\User1' `
      -PublicRoot 'C:\Users\Public'
    @($collision.Roots | ForEach-Object StagingName) | Should -Be @('Downloads', 'Downloads (2)')
    @($collision.Mappings | ForEach-Object Source) | Should -Be @('Downloads', 'Downloads (2)')
    ($collision.Notes -join "`n") | Should -Match "Extract D:\\Downloads as 'Downloads \(2\)'"
  }

  It 'builds the plan with staging paths, manual copies, decisions, and a folder-map import call' {
    $long = @(
      [PSCustomObject]@{ Source = 'D:\Business\long'; Destination = 'C:\Users\User1\Documents\Business\long'; Length = 260; Limit = 259; IsDirectory = $false; WordWebPair = $false }
    )
    $outlook = [PSCustomObject]@{
      DataFiles = @([PSCustomObject]@{ Path = 'C:\Data\mail.pst'; Format = 'PST' })
      Signature = 'C:\Users\User1\AppData\Roaming\Microsoft\Signatures'
      Warnings  = @("The Outlook report is from computer 'OTHERPC', not 'THISPC'.")
      Error     = $null
    }

    $plan = Get-DataInventoryPlan `
      -Mapping $script:Mapping `
      -NetworkDrive @() `
      -LongPath $long `
      -Outlook $outlook `
      -StagingRoot "C:\Managed\O'Brien Migration" `
      -FolderMapPath 'C:\Reports\run\folder-map.json'

    @($plan.AbbDownload | Where-Object Path -EQ 'C:\Downloads')[0].StagingPath | Should -Be "C:\Managed\O'Brien Migration\Downloads"
    @($plan.ManualCopy | Where-Object Path -EQ 'C:\Data\mail.pst').Count | Should -Be 1
    @($plan.ManualCopy | Where-Object Path -Match 'Signatures').Count | Should -Be 1
    $plan.Decisions | Should -Contain "The Outlook report is from computer 'OTHERPC', not 'THISPC'."
    ($plan.Decisions -join "`n") | Should -Match 'too long'
    @($plan.Decisions | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count | Should -Be 0
    $plan.SuggestedImportCall | Should -Match "-SourcePath 'C:\\Managed\\O''Brien Migration'"
    $plan.SuggestedImportCall | Should -Match "-FolderMapPath 'C:\\Reports\\run\\folder-map.json'"
  }

  It 'writes a folder map the importer can read' {
    $plan = Get-DataInventoryPlan -Mapping $script:Mapping -NetworkDrive @() -LongPath @() -Outlook $null -StagingRoot 'C:\Managed\Migration' -FolderMapPath $null

    $document = ConvertTo-DataInventoryFolderMap -Plan $plan -RunId 'run' -Computer 'OLDPC' -UserName 'EXAMPLE\User1'
    $parsed = ConvertTo-DataInventoryJson -InputObject $document -Depth 6 | ConvertFrom-Json

    $parsed.SchemaVersion | Should -Be 1
    @($parsed.Mappings).Count | Should -Be @($script:Mapping.Mappings).Count
    @($parsed.Mappings | Where-Object Source -EQ 'User1\Music\Desktop')[0].Target | Should -Be 'Desktop'
    ($parsed.Notes -join "`n") | Should -Match 'Desktop is stored inside Music'
    $plan.SuggestedImportCall | Should -Match "-FolderMapPath '<folder-map.json>'"
  }

  It 'accepts only local folders outside the profile and public folders for -Include' {
    $business = Join-Path $TestDrive 'Business'
    $null = New-Item -ItemType Directory -Path $business -Force

    $resolved = @(Resolve-DataInventoryInclude -Path @($business, "$business\") -ProfilePath 'C:\Users\User1' -PublicRoot 'C:\Users\Public')

    $resolved | Should -Be @($business)
    { Resolve-DataInventoryInclude -Path @('\\server\share') -ProfilePath 'C:\Users\User1' -PublicRoot 'C:\Users\Public' } | Should -Throw '*local drive path*'
    { Resolve-DataInventoryInclude -Path @($TestDrive) -ProfilePath $TestDrive -PublicRoot 'C:\Users\Public' } | Should -Throw '*outside the profile*'
    { Resolve-DataInventoryInclude -Path @((Join-Path $TestDrive 'Missing')) -ProfilePath 'C:\Users\User1' -PublicRoot 'C:\Users\Public' } | Should -Throw '*not found*'
  }

}

Describe 'Show-UserDataInventory scans' {
  It 'scans fixed lettered volumes, skips system and managed folders, and honors -Include' {
    $fixedRoot = Join-Path $TestDrive 'FixedRoot'
    $removableRoot = Join-Path $TestDrive 'RemovableRoot'
    $business = Join-Path $fixedRoot 'Business'
    $managed = Join-Path $fixedRoot 'Managed'
    $servicing = Join-Path $fixedRoot '$WinREAgent'
    $vendor = Join-Path $fixedRoot 'inetpub'
    $included = Join-Path $TestDrive 'Elsewhere\Accounting'
    foreach ($directory in @($business, $managed, $servicing, $vendor, $included, (Join-Path $removableRoot 'Camera'))) {
      $null = New-Item -ItemType Directory -Path $directory -Force
    }
    [IO.File]::WriteAllText((Join-Path $business 'ledger.txt'), '12345')
    [IO.File]::WriteAllText((Join-Path $vendor 'site.txt'), 'web')
    [IO.File]::WriteAllText((Join-Path $included 'books.txt'), 'books')

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
      $scan = @{
        ProfilePath = 'C:\Users\User1'
        ScanVolumes = $true
        ManagedRoot = $managed
      }
      $outside = @(Get-DataInventoryOutsideProfile @scan)
      $withInclude = @(Get-DataInventoryOutsideProfile @scan -Include @($vendor, $included))

      Should -Invoke Get-Volume -Times 2 -Exactly
      $outside.Count | Should -Be 1
      $outside[0].Path | Should -Be $business
      $outside[0].Files | Should -Be 1
      $outside[0].Included | Should -BeFalse

      @($withInclude | ForEach-Object Path) | Should -Be @($business, $vendor, $included)
      @($withInclude | Where-Object Included | ForEach-Object Path) | Should -Be @($vendor, $included)
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
