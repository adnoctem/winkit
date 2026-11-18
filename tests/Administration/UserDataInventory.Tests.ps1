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

  It 'lists public folders as decisions by default and lists the public desktop files for a manual copy' {
    @($script:Mapping.Mappings | Where-Object Kind -EQ 'Public').Count | Should -Be 0
    ($script:Mapping.Decisions -join "`n") | Should -Match 'Public Documents holds 5 files .* Use -IncludePublic Documents'
    ($script:Mapping.Decisions -join "`n") | Should -Not -Match 'Public Pictures'
    @($script:Mapping.ManualCopy | Where-Object Path -EQ 'C:\Users\Public\Desktop').Count | Should -Be 1
  }

  It 'maps public folders named with -IncludePublic below the user folders' {
    $included = Get-DataInventoryMapping `
      -KnownFolder @() `
      -PublicFolder $script:Public `
      -OutsideProfile @() `
      -ProfilePath 'C:\Users\User1' `
      -PublicRoot 'C:\Users\Public' `
      -IncludePublic Documents

    @($included.Mappings | Where-Object Source -EQ 'Public\Documents')[0].Target | Should -Be 'Documents\Public Documents'
    ($included.Decisions -join "`n") | Should -Not -Match 'Public Documents holds'
  }

  It 'maps only known folders named with -Folder and notes the others' {
    $selected = Get-DataInventoryMapping `
      -KnownFolder $script:Known `
      -PublicFolder @() `
      -OutsideProfile @() `
      -ProfilePath 'C:\Users\User1' `
      -PublicRoot 'C:\Users\Public' `
      -Folder Desktop, Documents

    @($selected.Mappings | ForEach-Object Target) | Should -Be @('Desktop', 'Documents')
    ($selected.Notes -join "`n") | Should -Match 'did not select and are not mapped: Downloads, Music, Favorites, Searches'
  }

  It 'notes empty known folders instead of mapping them and flags nearly empty ones' {
    $empty = Get-KnownFolderFixture -Name 'Music' -Path 'C:\Users\User1\Music'
    $empty.Files = 0
    $empty | Add-Member -NotePropertyName NestedFolders -NotePropertyValue @('Desktop')
    $videos = Get-KnownFolderFixture -Name 'Videos' -Path 'C:\Users\User1\Videos'
    $videos.Files = 0
    $few = Get-KnownFolderFixture -Name 'Pictures' -Path 'C:\Users\User1\Pictures'
    $few.Files = 1
    $many = Get-KnownFolderFixture -Name 'Documents' -Path 'C:\Users\User1\Documents'
    $many.Files = 30

    $result = Get-DataInventoryMapping -KnownFolder @($empty, $videos, $few, $many) -PublicFolder @() -OutsideProfile @() -ProfilePath 'C:\Users\User1' -PublicRoot 'C:\Users\Public'

    @($result.Mappings | ForEach-Object Target) | Should -Be @('Pictures', 'Documents')
    $result.Notes | Should -Contain 'Known folder Music is empty apart from Desktop and is not mapped.'
    $result.Notes | Should -Contain 'Known folder Videos is empty and is not mapped.'
    ($result.Notes -join "`n") | Should -Match 'Pictures holds only 1 file'
    ($result.Notes -join "`n") | Should -Not -Match 'Documents holds only'
  }

  It 'notes program shortcuts on the public desktop instead of copying them' {
    $desktop = [PSCustomObject]@{
      Name      = 'Desktop'
      Path      = 'C:\Users\Public\Desktop'
      Files     = 4
      Shortcuts = @(
        [PSCustomObject]@{ Path = 'C:\Users\Public\Desktop\Accounting.lnk'; Target = 'C:\as5win\as5.exe'; IsNetwork = $false }
        [PSCustomObject]@{ Path = 'C:\Users\Public\Desktop\Word.lnk'; Target = ''; IsNetwork = $false }
        [PSCustomObject]@{ Path = 'C:\Users\Public\Desktop\Share.lnk'; Target = '\\server\share'; IsNetwork = $true }
        [PSCustomObject]@{ Path = 'C:\Users\Public\Desktop\Portal.url'; Target = 'https://example.com'; IsNetwork = $false }
      )
    }
    $onlyPrograms = [PSCustomObject]@{
      Name      = 'Desktop'
      Path      = 'C:\Users\Public\Desktop'
      Files     = 2
      Shortcuts = @($desktop.Shortcuts[0], $desktop.Shortcuts[1])
    }
    $base = @{
      KnownFolder    = @()
      OutsideProfile = @()
      ProfilePath    = 'C:\Users\User1'
      PublicRoot     = 'C:\Users\Public'
    }

    $mixed = Get-DataInventoryMapping @base -PublicFolder @($desktop)
    $programs = Get-DataInventoryMapping @base -PublicFolder @($onlyPrograms)

    ($programs.Notes -join "`n") | Should -Match '2 shortcut\(s\) to local programs \(Accounting, Word\)'
    @($programs.ManualCopy).Count | Should -Be 0
    @($programs.Roots).Count | Should -Be 0
    @($mixed.ManualCopy)[0].Reason | Should -Match 'its 1 file\(s\) that are not program shortcuts'
    ($mixed.Decisions -join "`n") | Should -Match 'Share.lnk -> \\\\server\\share'
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
      -LongPath $long `
      -Outlook $outlook `
      -StagingRoot "C:\Managed\O'Brien Migration" `
      -FolderMapPath 'C:\Reports\run\folder-map.json' `
      -TargetUser 'EXAMPLE\NewUser'

    @($plan.AbbDownload | Where-Object Path -EQ 'C:\Downloads')[0].StagingPath | Should -Be "C:\Managed\O'Brien Migration\Downloads"
    @($plan.ManualCopy | Where-Object Path -EQ 'C:\Data\mail.pst').Count | Should -Be 1
    @($plan.ManualCopy | Where-Object Path -Match 'Signatures').Count | Should -Be 1
    $plan.Decisions | Should -Contain "The Outlook report is from computer 'OTHERPC', not 'THISPC'."
    ($plan.Decisions -join "`n") | Should -Match 'too long'
    @($plan.Decisions | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count | Should -Be 0
    $plan.SuggestedImportCall | Should -Match "-SourcePath 'C:\\Managed\\O''Brien Migration'"
    $plan.SuggestedImportCall | Should -Match "-FolderMapPath 'C:\\Reports\\run\\folder-map.json'"
    $plan.SuggestedImportCall | Should -Match "-TargetUser 'EXAMPLE\\NewUser'"
  }

  It 'writes a folder map the importer can read' {
    $plan = Get-DataInventoryPlan -Mapping $script:Mapping -LongPath @() -Outlook $null -StagingRoot 'C:\Managed\Migration' -FolderMapPath $null

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

Describe 'Show-UserDataInventory drives' {
  BeforeEach {
    $script:RegistryRoot = 'Registry::HKEY_USERS\S-1-5-21-1-2-3-1001'
    $script:OtherUser = [PSCustomObject]@{
      Name      = 'EXAMPLE\User1'
      SID       = 'S-1-5-21-1-2-3-1001'
      IsCurrent = $false
    }
    $script:Volatile = [PSCustomObject]@{
      HOMEDRIVE = 'X:'
      HOMESHARE = '\\server\home\User1'
    }
    $script:Directory = [PSCustomObject]@{
      HomeDrive     = 'X:'
      HomeDirectory = '\\server\home\User1\'
      ScriptPath    = 'logon.cmd'
      Error         = $null
    }

    Mock Get-ItemProperty { $script:Volatile } -ParameterFilter { $LiteralPath -like '*\Volatile Environment' }
    Mock Get-ItemProperty { $null } -ParameterFilter { $LiteralPath -like 'HKLM:*' }
    Mock Test-Path { $false } -ParameterFilter { $LiteralPath -like '*\Network' }
    Mock Get-DataInventoryDirectoryAccount { $script:Directory }
    Mock Test-DataInventoryAdministrator { $true }
    Mock Get-PSDrive {
      [PSCustomObject]@{ Name = 'Z'; DisplayRoot = '\\server\admin-session' }
    }
  }

  It 'reads the home drive and logon script of another user without session drives' {
    $result = Get-DataInventoryNetworkDrive -RegistryRoot $script:RegistryRoot -InventoryUser $script:OtherUser -Measure $false

    $homeDrive = @($result.Drives | Where-Object Source -EQ 'HomeDrive')
    $homeDrive.Count | Should -Be 1
    $homeDrive[0].Letter | Should -Be 'X'
    $homeDrive[0].Target | Should -Be '\\server\home\User1'
    $homeDrive[0].Origin | Should -Be 'Volatile Environment'
    @($result.Drives | Where-Object Source -EQ 'LogonScript')[0].Target | Should -Be 'logon.cmd'
    @($result.Drives | Where-Object Source -EQ 'Connected').Count | Should -Be 0
    ($result.Notes -join "`n") | Should -Match 'Connected drives were not read'
    Should -Invoke Get-PSDrive -Times 0 -Exactly
  }

  It 'falls back to Active Directory when the user is not signed in' {
    $script:Volatile = $null

    $result = Get-DataInventoryNetworkDrive -RegistryRoot $script:RegistryRoot -InventoryUser $script:OtherUser -Measure $false

    $homeDrive = @($result.Drives | Where-Object Source -EQ 'HomeDrive')
    $homeDrive.Count | Should -Be 1
    $homeDrive[0].Target | Should -Be '\\server\home\User1'
    $homeDrive[0].Origin | Should -Be 'Active Directory'
  }

  It 'skips Active Directory for local accounts' {
    $local = [PSCustomObject]@{
      Name      = "$env:COMPUTERNAME\User1"
      SID       = 'S-1-5-21-1-2-3-1001'
      IsCurrent = $false
    }

    $null = Get-DataInventoryNetworkDrive -RegistryRoot $script:RegistryRoot -InventoryUser $local -Measure $false

    Should -Invoke Get-DataInventoryDirectoryAccount -Times 0 -Exactly
  }

  It 'reads connected drives only for the current user in a non-elevated session' {
    $current = [PSCustomObject]@{
      Name      = 'EXAMPLE\User1'
      SID       = 'S-1-5-21-1-2-3-1001'
      IsCurrent = $true
    }

    $elevated = Get-DataInventoryNetworkDrive -RegistryRoot 'HKCU:' -InventoryUser $current -Measure $false
    Mock Test-DataInventoryAdministrator { $false }
    $plain = Get-DataInventoryNetworkDrive -RegistryRoot 'HKCU:' -InventoryUser $current -Measure $false

    ($elevated.Notes -join "`n") | Should -Match 'EnableLinkedConnections'
    @($elevated.Drives | Where-Object Source -EQ 'Connected').Count | Should -Be 0
    @($plain.Drives | Where-Object Source -EQ 'Connected')[0].Target | Should -Be '\\server\admin-session'
  }

  It 'maps the home drive and lists other drives and the logon script as decisions' {
    $drives = @(
      [PSCustomObject]@{ Letter = 'X'; Target = '\\server\home\User1'; Source = 'HomeDrive'; Measured = $false; Error = $null }
      [PSCustomObject]@{ Letter = 'G'; Target = '\\server\production'; Source = 'Persistent'; Measured = $false; Error = $null }
      [PSCustomObject]@{ Letter = $null; Target = 'logon.cmd'; Source = 'LogonScript'; Measured = $false; Error = $null }
    )
    $mappingParameters = @{
      KnownFolder    = @()
      PublicFolder   = @()
      OutsideProfile = @()
      ProfilePath    = 'C:\Users\User1'
      PublicRoot     = 'C:\Users\Public'
      NetworkDrive   = $drives
      DriveNote      = @('Connected drives were not read.')
    }

    $default = Get-DataInventoryMapping @mappingParameters
    $custom = Get-DataInventoryMapping @mappingParameters -HomeDriveTarget 'Documents\Home'

    $default.Mappings[0].Source | Should -Be 'HomeDrive'
    $default.Mappings[0].Target | Should -Be 'Private'
    $default.Mappings[0].Kind | Should -Be 'HomeDrive'
    $default.Roots[0].StagingName | Should -Be 'HomeDrive'
    $custom.Mappings[0].Target | Should -Be 'Documents\Home'
    ($default.Notes -join "`n") | Should -Match "Home drive X: .*'HomeDrive' staging folder"
    $default.Notes | Should -Contain 'Connected drives were not read.'
    ($default.Decisions -join "`n") | Should -Match 'Network drive G: .* from Persistent'
    ($default.Decisions -join "`n") | Should -Match 'Logon script logon.cmd'
  }
}

Describe 'Show-UserDataInventory program owners' {
  It 'reads directories from uninstall values and ignores commands without a path' {
    ConvertTo-DataInventoryProgramDirectory -Value '"C:\asfibu\bin\fibu.exe",0' -IsFile | Should -Be 'C:\asfibu\bin'
    ConvertTo-DataInventoryProgramDirectory -Value 'C:\as5win\as5.exe,1' -IsFile | Should -Be 'C:\as5win'
    ConvertTo-DataInventoryProgramDirectory -Value '"C:\VR-NetWorld\unins000.exe" /SILENT' -IsFile | Should -Be 'C:\VR-NetWorld'
    ConvertTo-DataInventoryProgramDirectory -Value 'MsiExec.exe /X{2A360EF1-FFA5-4BF4-8709-EF9D2FB3669E}' -IsFile | Should -BeNullOrEmpty
    ConvertTo-DataInventoryProgramDirectory -Value 'C:\Install\Accounting 2.0\' | Should -Be 'C:\Install\Accounting 2.0'
    ConvertTo-DataInventoryProgramDirectory -Value '' | Should -BeNullOrEmpty
  }

  It 'collects evidence from uninstall entries without InstallLocation and from shortcuts' {
    $script:Uninstall = [PSCustomObject]@{
      DisplayName     = 'Example Accounting'
      InstallLocation = ''
      DisplayIcon     = '"C:\asfibu\bin\fibu.exe",0'
      UninstallString = 'MsiExec.exe /X{2A360EF1-FFA5-4BF4-8709-EF9D2FB3669E}'
    }
    Mock Get-ItemProperty { $script:Uninstall } -ParameterFilter { $Path -like '*Uninstall*' }
    $menu = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $menu -Force
    $shell = New-Object -ComObject WScript.Shell
    try {
      $link = $shell.CreateShortcut((Join-Path $menu 'Banking.lnk'))
      $link.TargetPath = 'C:\VR-NetWorld\vrnw.exe'
      $link.Save()
    }
    finally {
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }

    $evidence = @(Get-DataInventoryProgramEvidence -RegistryRoot 'HKCU:' -ShortcutRoot @($menu))

    @($evidence | Where-Object { $_.MatchedBy -eq 'DisplayIcon' -and $_.Path -eq 'C:\asfibu\bin' })[0].Name | Should -Be 'Example Accounting'
    @($evidence | Where-Object MatchedBy -EQ 'UninstallString').Count | Should -Be 0
    @($evidence | Where-Object MatchedBy -EQ 'Shortcut')[0].Path | Should -Be 'C:\VR-NetWorld'
  }

  It 'names owners whose evidence points into the folder and reports how they matched' {
    $folder = Join-Path $TestDrive 'asfibu'
    $null = New-Item -ItemType Directory -Path $folder -Force
    [IO.File]::WriteAllText((Join-Path $folder 'ledger.dat'), 'x')
    $evidence = @(
      [PSCustomObject]@{ Name = 'Example Accounting'; MatchedBy = 'DisplayIcon'; Path = "$folder\bin"; Evidence = 'icon' }
      [PSCustomObject]@{ Name = 'Example Accounting'; MatchedBy = 'Shortcut'; Path = $folder; Evidence = 'link' }
      [PSCustomObject]@{ Name = 'Example Accounting'; MatchedBy = 'Shortcut'; Path = "$folder\tools"; Evidence = 'second link' }
      [PSCustomObject]@{ Name = 'Broad Tool'; MatchedBy = 'InstallLocation'; Path = $TestDrive; Evidence = 'parent' }
    )

    $outside = @(Get-DataInventoryOutsideProfile -ProfilePath 'C:\Users\User1' -Include @($folder) -ScanVolumes $false -ProgramEvidence $evidence)
    $mapping = Get-DataInventoryMapping -KnownFolder @() -PublicFolder @() -OutsideProfile @($outside[0] | Select-Object * -ExcludeProperty Included) -ProfilePath 'C:\Users\User1' -PublicRoot 'C:\Users\Public'

    $outside[0].Programs | Should -Be @('Example Accounting')
    @($outside[0].ProgramMatches | ForEach-Object MatchedBy) | Should -Be @('DisplayIcon', 'Shortcut')
    ($mapping.Decisions -join "`n") | Should -Match 'probable programs: Example Accounting \(DisplayIcon, Shortcut\)'
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
      $paths = @(Get-DataInventoryLongPath -SourceRoot $source -DestinationRoot $target)

      $paths.Count | Should -BeGreaterThan 0
      $paths[0].Length | Should -BeGreaterThan $paths[0].Limit
    }
    finally {
      [IO.Directory]::Delete($longFull, $true)
      Remove-Item -LiteralPath $source -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  It 'flags a relocated folder whose mapped target is too long although its source is not' {
    $source = Join-Path $TestDrive ([guid]::NewGuid().ToString('N').Substring(0, 8))
    $file = Join-Path $source ('a' * 120)
    $null = New-Item -ItemType Directory -Path $file -Force
    [IO.File]::WriteAllText((Join-Path $file 'b.txt'), 'x')
    $target = 'C:\Users\User1\Documents\Public Documents\' + ('t' * 100)

    $sourceOnly = @(Get-DataInventoryLongPath -SourceRoot $source)
    $mapped = @(Get-DataInventoryLongPath -SourceRoot $source -DestinationRoot $target)

    $sourceOnly.Count | Should -Be 0
    @($mapped | Where-Object { -not $_.IsDirectory }).Count | Should -Be 1
    @($mapped | Where-Object { -not $_.IsDirectory })[0].Destination | Should -BeLike "$target\*\b.txt"
  }

  It 'reads UNC listings from robocopy and applies the import limits exactly' {
    $share = '\\server\home\User1'
    $deep = $share + '\' + ('d' * 200)
    $text = @(
      ("`t`t" + $deep + '\'),
      ("`t`t`t" + $deep + '\' + ('f' * 30) + '.txt'),
      ("`t`t`t" + $deep + '\' + ('g' * 100) + '.txt')
    ) -join "`r`n"
    $destination = 'C:\Users\User1\Private'

    $paths = @(ConvertFrom-DataInventoryRobocopyList -Text $text -SourceRoot $share -DestinationRoot $destination)

    # 'Private\' + 200 = 223 chars for the directory; the 30-char file reaches 258 (allowed), the 100-char file 328.
    $paths.Count | Should -Be 1
    $paths[0].Source | Should -Be ($deep + '\' + ('g' * 100) + '.txt')
    $paths[0].Destination | Should -Be ($destination + '\' + ('d' * 200) + '\' + ('g' * 100) + '.txt')
    $paths[0].Limit | Should -Be 259
  }

  It 'derives the new user from -TargetRoot or warns that it was assumed' {
    $user = [PSCustomObject]@{
      Name        = 'EXAMPLE\OldUser'
      ProfilePath = 'C:\Users\OldUser'
    }

    $assumed = Resolve-DataInventoryTarget -TargetRoot $null -InventoryUser $user
    $derived = Resolve-DataInventoryTarget -TargetRoot 'C:\Users\NewUser' -InventoryUser $user
    $elsewhere = Resolve-DataInventoryTarget -TargetRoot 'D:\Profiles\NewUser' -InventoryUser $user

    $assumed.Assumed | Should -BeTrue
    $assumed.Root | Should -Be 'C:\Users\OldUser'
    $assumed.User | Should -BeNullOrEmpty
    ($assumed.Notes -join "`n") | Should -Match '-TargetRoot was not given'
    $derived.Assumed | Should -BeFalse
    $derived.User | Should -Be 'EXAMPLE\NewUser'
    ($derived.Notes -join "`n") | Should -Match 'derived from -TargetRoot'
    $elsewhere.User | Should -BeNullOrEmpty
    ($elsewhere.Notes -join "`n") | Should -Match 'cannot be determined'
  }
}
