#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Cross-script mock context is removed after each test; mocks run in the invoked script scope.')]
param ()

# Real robocopy runs only against TestDrive. Identity/profile discovery and ACL
# mutations are mocked in workflow tests. No Windows profile is created or changed.
BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.4.1 -Force
  $script:ImportPath = Join-Path $PSScriptRoot '../../scripts/Administration/Import-UserProfileData.ps1'
  $script:NativeProcess = Get-Command Invoke-SafeProcess -CommandType Cmdlet -Module AdNoctem.Substrate.PowerShell
  $script:Content = Get-Content -LiteralPath $script:ImportPath -Raw
  $script:Ast = [Management.Automation.Language.Parser]::ParseFile($script:ImportPath, [ref]$null, [ref]$null)
  $definitions = @($script:Ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false))
  foreach ($definition in $definitions) {
    . ([scriptblock]::Create($definition.Extent.Text))
  }

  # Use the same entry point with helpers loaded separately so Pester can replace
  # profile discovery without exposing a production bypass parameter.
  $script:HarnessText = $script:Content
  foreach ($definition in @($definitions | Sort-Object { $_.Extent.StartOffset } -Descending)) {
    $script:HarnessText = $script:HarnessText.Remove($definition.Extent.StartOffset, $definition.Extent.EndOffset - $definition.Extent.StartOffset)
  }
  $script:HarnessText = $script:HarnessText -replace '(?m)^#Requires -RunAsAdministrator\r?\n', ''
  $script:HarnessText = $script:HarnessText -replace '(?m)^Import-Module AdNoctem.Substrate.PowerShell[^\r\n]+', ''

  function Write-ProfileFixture {
    param (
      [string]
      $Path,

      [string]
      $Text = 'fixture data'
    )

    $null = New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force
    [IO.File]::WriteAllText($Path, $Text)
  }

  function Invoke-ProfileImport {
    param (
      [hashtable]
      $Extra = @{}
    )

    $parameters = @{
      SourcePath      = $global:WinkitProfileTestContext.Source
      TargetUser      = 'TEST\User1'
      ReportDirectory = $script:Reports
      Folder          = @('Documents')
      PassThru        = $true
    }
    foreach ($key in $Extra.Keys) {
      # A $null value removes a default parameter, for example Folder when a folder map is used.
      if ($null -eq $Extra[$key]) {
        $parameters.Remove($key)
        continue
      }

      $parameters[$key] = $Extra[$key]
    }
    $global:LASTEXITCODE = 0
    $results = @(& $script:Harness @parameters 6>$null)
    return [PSCustomObject]@{
      ExitCode = $global:LASTEXITCODE
      Results  = $results
      Summary  = $results | Where-Object Action -EQ 'ImportProfileData'
    }
  }
}

Describe 'Import-UserProfileData conventions' {
  It 'pins the module, requires elevation, and exposes normal preview/help conventions' {
    $script:Content | Should -Match '#Requires -RunAsAdministrator'
    $script:Content | Should -Match "RequiredVersion = '0.4.1'"
    $script:Content | Should -Match 'Server Core: supported'
    $script:Content | Should -Match 'SYSTEM-account execution: supported'
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    (Get-Help $script:ImportPath).Synopsis | Should -Match 'Imports staged user data'
  }
}

Describe 'Import-UserProfileData JSON output' {
  It 'keeps quotes and angle brackets readable and round-trips literal escape text' {
    $value = [PSCustomObject]@{
      Call    = ".\Import.ps1 -SourcePath 'C:\Migration\<user>' -Mode A&B"
      Literal = 'C:\u0027\data'
    }

    $json = ConvertTo-ImportJson -InputObject $value -Depth 4

    $json | Should -Not -Match '(?<!\\)\\u00(27|3c|3e|26)'
    $json | Should -Match "-SourcePath 'C:"
    $roundTrip = $json | ConvertFrom-Json
    $roundTrip.Call | Should -BeExactly $value.Call
    $roundTrip.Literal | Should -BeExactly 'C:\u0027\data'
  }
}

Describe 'Import-UserProfileData workflow' {
  BeforeEach {
    $global:WinkitProfileTestContext = @{
      NativeProcess = $script:NativeProcess
    }
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $global:WinkitProfileTestContext.Source = Join-Path $script:Root 'staged user'
    $global:WinkitProfileTestContext.Target = Join-Path $script:Root 'new user'
    $script:Reports = Join-Path $script:Root 'reports'
    $script:Harness = Join-Path $script:Root 'Invoke-TestImport.ps1'
    $global:WinkitProfileTestContext.TargetSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') -Text 'hello from backup'
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'NTUSER.DAT') -Text 'do not change'
    [IO.File]::WriteAllText($script:Harness, $script:HarnessText, (New-Object Text.UTF8Encoding($true)))

    Mock Test-Elevation { $true }
    Mock Get-ImportProfile {
      [PSCustomObject]@{
        Path = $global:WinkitProfileTestContext.Target
        SID  = $global:WinkitProfileTestContext.TargetSid
        User = 'TEST\User1'
      }
    }
    Mock Get-ImportVolume {
      [PSCustomObject]@{ AvailableFreeSpace = [long]10GB }
    }
    Mock Assert-ImportPermission {}
    Mock Invoke-ImportPermission {}
    Mock Initialize-ImportReport {
      $null = New-Item -ItemType Directory -Path $Path
    }
    Mock Invoke-SafeProcess {
      if ($FilePath -notlike '*\robocopy.exe') {
        throw 'Workflow must not invoke unmocked permission tools.'
      }
      & $global:WinkitProfileTestContext.NativeProcess -FilePath $FilePath -ArgumentList $ArgumentList -AsResult
    }
  }

  AfterEach {
    Remove-Variable -Name WinkitProfileTestContext -Scope Global -ErrorAction SilentlyContinue
  }

  It 'copies with real robocopy, keeps extras and staging, and writes a hash manifest' {
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'Documents\existing.txt') -Text 'keep me'
    $null = New-Item -ItemType Directory -Path (Join-Path $global:WinkitProfileTestContext.Source 'Documents\Empty')
    $run = Invoke-ProfileImport
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    $run.Summary.Status | Should -Be 'Completed'
    Get-Content -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\report [1].txt') | Should -Be 'hello from backup'
    Get-Content -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\existing.txt') | Should -Be 'keep me'
    Get-Content -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'NTUSER.DAT') | Should -Be 'do not change'
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\Empty') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
    $manifest = Get-Content -LiteralPath $run.Summary.ManifestPath -Raw | ConvertFrom-Json
    $manifest.Status | Should -Be 'Completed'
    $manifest.Plan[0].Inventory.Files[0].SHA256 | Should -Match '^[0-9A-F]{64}$'
    $manifest.Target.SID | Should -Be $global:WinkitProfileTestContext.TargetSid
  }

  It 'passes data-only copy, bounded retries, and no mirror/move switches' {
    $run = Invoke-ProfileImport
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    Should -Invoke Invoke-SafeProcess -Times 1 -Exactly -ParameterFilter {
      $ArgumentList -contains '/COPY:DAT' -and $ArgumentList -contains '/DCOPY:DAT' -and
      $ArgumentList -contains '/XJ' -and $ArgumentList -contains '/R:1' -and
      $ArgumentList -notcontains '/MIR' -and $ArgumentList -notcontains '/MOVE' -and
      $ArgumentList -notcontains '/COPYALL'
    }
  }

  It 'previews without copying, ACL changes, or logs' {
    $run = Invoke-ProfileImport -Extra @{ DryRun = $true; RemoveSource = $true }
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    $run.Summary.Status | Should -Be 'Preview'
    Test-Path -LiteralPath $script:Reports | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents') | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
    Should -Invoke Invoke-ImportPermission -Times 0 -Exactly
  }

  It 'refuses different destination content before making any changes' {
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'Documents\report [1].txt') -Text 'newer work'
    $run = Invoke-ProfileImport
    $run.ExitCode | Should -Be 1
    $run.Summary.Error | Should -Match 'content conflict'
    Get-Content -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\report [1].txt') | Should -Be 'newer work'
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
    Test-Path -LiteralPath $script:Reports | Should -BeFalse
  }

  It 'overwrites only with an explicit switch, even with matching size and timestamp' {
    $source = Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt'
    $target = Join-Path $global:WinkitProfileTestContext.Target 'Documents\report [1].txt'
    Write-ProfileFixture -Path $target -Text 'different content'
    (Get-Item -LiteralPath $target).LastWriteTimeUtc = (Get-Item -LiteralPath $source).LastWriteTimeUtc
    $run = Invoke-ProfileImport -Extra @{ Overwrite = $true }
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    (Get-FileHash -LiteralPath $target).Hash | Should -Be (Get-FileHash -LiteralPath $source).Hash
  }

  It 'skips Explorer metadata instead of reporting a destination conflict' {
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Documents\desktop.ini') -Text 'source desktop'
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'Documents\desktop.ini') -Text 'target desktop'
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Documents\Thumbs.db') -Text 'source thumbs'
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'Documents\Thumbs.db') -Text 'target thumbs'

    $run = Invoke-ProfileImport

    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    Get-Content -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\desktop.ini') | Should -Be 'target desktop'
    Get-Content -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\Thumbs.db') | Should -Be 'target thumbs'
    $manifest = Get-Content -LiteralPath $run.Summary.ManifestPath -Raw | ConvertFrom-Json
    @($manifest.Plan[0].Inventory.Excluded | Where-Object Reason -EQ 'ExplorerMetadata').Count | Should -Be 2
  }

  It 'stops before copying when a destination path exceeds the supported length' {
    $long = Join-Path $global:WinkitProfileTestContext.Source 'Documents'
    for ($index = 0; $index -lt 12; $index++) {
      $long = Join-Path $long ('segment-{0:D2}-{1}' -f $index, ('x' * 24))
    }
    $longFull = '\\?\' + $long
    $null = [IO.Directory]::CreateDirectory($longFull)
    [IO.File]::WriteAllText(($longFull + '\payload.txt'), 'x')
    try {
      $run = Invoke-ProfileImport

      $run.ExitCode | Should -Be 1
      $run.Summary.Error | Should -Match 'exceed the supported length'
      Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents') | Should -BeFalse
      Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
    }
    finally {
      [IO.Directory]::Delete($longFull, $true)
    }
  }

  It 'can repeat an already verified import without deleting target-only files' {
    (Invoke-ProfileImport).ExitCode | Should -Be 0
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'Documents\extra.txt')
    (Invoke-ProfileImport).ExitCode | Should -Be 0
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\extra.txt') | Should -BeTrue
  }

  It 'excludes mail, hives and AppData, and cleanup retains every excluded source item' {
    foreach ($relative in @('Documents\archive.pst', 'Documents\mail.ost', 'Documents\NTUSER.DAT', 'Documents\AppData\secret')) {
      Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source $relative)
    }
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Desktop\unselected.txt')
    $run = Invoke-ProfileImport -Extra @{ RemoveSource = $true }
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\archive.pst') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\AppData\secret') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Desktop\unselected.txt') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\archive.pst') | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\AppData') | Should -BeFalse
  }

  It 'preserves all staging when a later folder copy fails' {
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Desktop\second.txt')
    Mock Invoke-SafeProcess {
      if ($ArgumentList -contains '/COPY:DAT' -and $ArgumentList[0] -like '*\Desktop') {
        return [PSCustomObject]@{ ExitCode = 8; TimedOut = $false; Cancelled = $false }
      }
      & $global:WinkitProfileTestContext.NativeProcess -FilePath $FilePath -ArgumentList $ArgumentList -AsResult
    }
    $run = Invoke-ProfileImport -Extra @{ Folder = @('Documents', 'Desktop'); RemoveSource = $true }
    $run.ExitCode | Should -Be 1
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
    (Get-Content -LiteralPath $run.Summary.ManifestPath -Raw | ConvertFrom-Json).Status | Should -Be 'Failed'
  }

  It 'rejects missing copy output even when robocopy claims success' {
    Mock Invoke-SafeProcess {
      $null = New-Item -ItemType Directory -Path $ArgumentList[1]
      [PSCustomObject]@{ ExitCode = 1; TimedOut = $false; Cancelled = $false }
    }
    $run = Invoke-ProfileImport -Extra @{ RemoveSource = $true }
    $run.ExitCode | Should -Be 1
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
    Should -Invoke Invoke-ImportPermission -Times 0 -Exactly
  }

  It 'retains staging when permission repair fails' {
    Mock Invoke-ImportPermission { throw 'ownership failed' }
    $run = Invoke-ProfileImport -Extra @{ RemoveSource = $true }
    $run.ExitCode | Should -Be 1
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
  }

  It 'does not copy when the initial audit cannot be written' {
    Mock Write-ImportAudit { throw 'audit disk unavailable' }
    $run = Invoke-ProfileImport
    $run.ExitCode | Should -Be 1
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
  }

  It 'does not clean up when writing the verified audit fails' {
    $global:WinkitProfileTestContext.RealAudit = Get-Command Write-ImportAudit
    Mock Write-ImportAudit {
      if ($Manifest.Status -eq 'Verified') {
        throw 'cannot persist verification'
      }
      & $global:WinkitProfileTestContext.RealAudit -Manifest $Manifest -Results $Results -Directory $Directory
    }
    $run = Invoke-ProfileImport -Extra @{ RemoveSource = $true }
    $run.ExitCode | Should -Be 1
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
  }

  It 'detects a profile becoming loaded after copy and retains staging' {
    $global:WinkitProfileTestContext.ProfileChecks = 0
    Mock Get-ImportProfile {
      $global:WinkitProfileTestContext.ProfileChecks++
      if ($global:WinkitProfileTestContext.ProfileChecks -ge 4) {
        throw 'The target profile is loaded.'
      }
      [PSCustomObject]@{
        Path = $global:WinkitProfileTestContext.Target
        SID  = $global:WinkitProfileTestContext.TargetSid
        User = 'TEST\User1'
      }
    }
    $run = Invoke-ProfileImport -Extra @{ RemoveSource = $true }
    $run.ExitCode | Should -Be 1
    Should -Invoke Invoke-ImportPermission -Times 0 -Exactly
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
  }

  It 'refuses a loaded target before copying' {
    Mock Get-ImportProfile { throw 'The target profile is loaded.' }
    $run = Invoke-ProfileImport
    $run.ExitCode | Should -Be 1
    $run.Summary.Error | Should -Match 'loaded'
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly
  }

  It 'refuses insufficient free space before creating a report or copying' {
    Mock Get-ImportVolume { [PSCustomObject]@{ AvailableFreeSpace = 1 } }
    $run = Invoke-ProfileImport
    $run.ExitCode | Should -Be 1
    $run.Summary.Error | Should -Match 'disk space'
    Test-Path -LiteralPath $script:Reports | Should -BeFalse
  }

  It 'imports explicit non-overlapping nested folder mappings' {
    $run = Invoke-ProfileImport -Extra @{
      Folder    = @()
      FolderMap = @{ Documents = 'Data\Imported Documents' }
    }
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Data\Imported Documents\report [1].txt') | Should -BeTrue
  }

  Context 'folder map from Show-UserDataInventory' {
    BeforeEach {
      $source = $global:WinkitProfileTestContext.Source
      Write-ProfileFixture -Path (Join-Path $source 'Old\Music\Desktop\desk.txt') -Text 'desktop file'
      Write-ProfileFixture -Path (Join-Path $source 'Old\Music\song.txt') -Text 'music file'
      Write-ProfileFixture -Path (Join-Path $source 'Old\Documents\letter.txt') -Text 'document'
      Write-ProfileFixture -Path (Join-Path $source 'Public\Documents\shared.txt') -Text 'shared document'
      Write-ProfileFixture -Path (Join-Path $source 'Downloads\setup.txt') -Text 'download'

      $script:FolderMapPath = Join-Path $script:Root 'folder-map.json'
      $script:FolderMap = [ordered]@{
        SchemaVersion = 1
        GeneratedBy   = 'Show-UserDataInventory'
        Mappings      = @(
          [ordered]@{ Source = 'Old\Music\Desktop'; Target = 'Desktop' }
          [ordered]@{ Source = 'Old\Music'; Target = 'Music' }
          [ordered]@{ Source = 'Old\Documents'; Target = 'Documents' }
          [ordered]@{ Source = 'Public\Documents'; Target = 'Documents\Public Documents' }
          [ordered]@{ Source = 'Downloads'; Target = 'Downloads' }
        )
        Notes         = @('Desktop is stored in Music\Desktop on the old computer.')
      }
      [IO.File]::WriteAllText($script:FolderMapPath, ($script:FolderMap | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
    }

    It 'imports nested sources once, places nested destinations, and cleans up verified staging' {
      $run = Invoke-ProfileImport -Extra @{ Folder = $null; FolderMapPath = $script:FolderMapPath; RemoveSource = $true }

      $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
      $target = $global:WinkitProfileTestContext.Target
      Get-Content -LiteralPath (Join-Path $target 'Desktop\desk.txt') | Should -Be 'desktop file'
      Get-Content -LiteralPath (Join-Path $target 'Music\song.txt') | Should -Be 'music file'
      Test-Path -LiteralPath (Join-Path $target 'Music\Desktop') | Should -BeFalse
      Get-Content -LiteralPath (Join-Path $target 'Documents\letter.txt') | Should -Be 'document'
      Get-Content -LiteralPath (Join-Path $target 'Documents\Public Documents\shared.txt') | Should -Be 'shared document'
      Get-Content -LiteralPath (Join-Path $target 'Downloads\setup.txt') | Should -Be 'download'

      $source = $global:WinkitProfileTestContext.Source
      Test-Path -LiteralPath (Join-Path $source 'Old\Music\Desktop\desk.txt') | Should -BeFalse
      Test-Path -LiteralPath (Join-Path $source 'Old\Music\song.txt') | Should -BeFalse
      Test-Path -LiteralPath (Join-Path $source 'Documents\report [1].txt') | Should -BeTrue

      $manifest = Get-Content -LiteralPath $run.Summary.ManifestPath -Raw | ConvertFrom-Json
      $manifest.FolderMapPath | Should -Be $script:FolderMapPath
      $manifest.FolderMapNote | Should -Contain 'Desktop is stored in Music\Desktop on the old computer.'
      $music = @($manifest.Plan | Where-Object { $_.Destination -like '*\Music' })[0]
      @($music.Inventory.Excluded | Where-Object Reason -EQ 'NestedMapping').Count | Should -Be 1
      @($music.Inventory.Files).Count | Should -Be 1
    }

    It 'refuses a nested destination that collides with data the outer mapping copies' {
      Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Old\Documents\Public Documents\clash.txt')

      $run = Invoke-ProfileImport -Extra @{ Folder = $null; FolderMapPath = $script:FolderMapPath }

      $run.ExitCode | Should -Be 1
      $run.Summary.Error | Should -Match 'collides'
      Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
    }

    It 'refuses a folder map combined with folder selections' {
      $run = Invoke-ProfileImport -Extra @{ FolderMapPath = $script:FolderMapPath }

      $run.ExitCode | Should -Be 1
      $run.Summary.Error | Should -Match 'on its own'
    }

    It 'refuses an unknown folder map schema or a missing mapped source' {
      $script:FolderMap.SchemaVersion = 2
      [IO.File]::WriteAllText($script:FolderMapPath, ($script:FolderMap | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
      $schemaRun = Invoke-ProfileImport -Extra @{ Folder = $null; FolderMapPath = $script:FolderMapPath }

      $script:FolderMap.SchemaVersion = 1
      $script:FolderMap.Mappings += [ordered]@{ Source = 'Old\Missing'; Target = 'Missing' }
      [IO.File]::WriteAllText($script:FolderMapPath, ($script:FolderMap | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false)))
      $missingRun = Invoke-ProfileImport -Extra @{ Folder = $null; FolderMapPath = $script:FolderMapPath }

      $schemaRun.ExitCode | Should -Be 1
      $schemaRun.Summary.Error | Should -Match 'schema'
      $missingRun.ExitCode | Should -Be 1
      $missingRun.Summary.Error | Should -Match 'does not exist'
    }
  }

  It 'rejects traversal, forbidden folders, overlapping mappings and missing explicit folders' {
    $cases = @(
      @{ Folder = @('..\staged user\Documents') },
      @{ Folder = @('AppData') },
      @{ Folder = @('Missing') },
      @{ FolderMap = @{ Documents = 'Other' } },
      @{ FolderMap = @{ Documents = '..\escaped' } }
    )
    foreach ($extra in $cases) {
      (Invoke-ProfileImport -Extra $extra).ExitCode | Should -Be 1
    }
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
  }

  It 'rejects source/target and report/source overlap' {
    (Invoke-ProfileImport -Extra @{ ReportDirectory = (Join-Path $global:WinkitProfileTestContext.Source 'logs') }).ExitCode | Should -Be 1
    Mock Get-ImportProfile {
      [PSCustomObject]@{ Path = $global:WinkitProfileTestContext.Source; SID = $global:WinkitProfileTestContext.TargetSid; User = 'TEST\User1' }
    }
    (Invoke-ProfileImport).ExitCode | Should -Be 1
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly
  }

  It 'skips source junctions without copying or deleting their targets' {
    $outside = Join-Path $script:Root 'outside'
    Write-ProfileFixture -Path (Join-Path $outside 'keep.txt')
    $link = Join-Path $global:WinkitProfileTestContext.Source 'Documents\loop'
    $null = New-Item -ItemType Junction -Path $link -Target $outside
    try {
      $run = Invoke-ProfileImport -Extra @{ RemoveSource = $true }
      $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
      Test-Path -LiteralPath (Join-Path $outside 'keep.txt') | Should -BeTrue
      Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\loop') | Should -BeFalse
    }
    finally {
      [IO.Directory]::Delete($link)
    }
  }

  It 'rejects destination junctions even on otherwise unrelated destination children' {
    $outside = Join-Path $script:Root 'outside'
    Write-ProfileFixture -Path (Join-Path $outside 'keep.txt')
    $null = New-Item -ItemType Directory -Path (Join-Path $global:WinkitProfileTestContext.Target 'Documents')
    $link = Join-Path $global:WinkitProfileTestContext.Target 'Documents\linked'
    $null = New-Item -ItemType Junction -Path $link -Target $outside
    try {
      $run = Invoke-ProfileImport
      $run.ExitCode | Should -Be 1
      Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
    }
    finally {
      [IO.Directory]::Delete($link)
    }
  }
}

Describe 'Windows system junctions in the destination profile' {
  BeforeAll {
    # Windows creates Documents\My Pictures (German: Eigene Bilder) and its siblings as Hidden+System junctions to
    # the profile's own folders, with an Everyone deny entry that stops browsing into them.
    function New-SystemJunction {
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only TestDrive fixtures.')]
      param (
        [string]
        $Path,

        [string]
        $Target,

        [switch]
        $Plain
      )

      $null = New-Item -ItemType Directory -Path $Target -Force
      $null = New-Item -ItemType Junction -Path $Path -Target $Target
      if (-not $Plain) {
        $null = & "$env:SystemRoot\System32\attrib.exe" +H +S $Path /L
        $null = & "$env:SystemRoot\System32\icacls.exe" $Path /deny '*S-1-1-0:(RD)' /L /Q
      }

      $null = $global:WinkitJunctionCleanup.Add($Path)
    }

    function Get-JunctionSddl {
      param (
        [string]
        $Path
      )

      $saved = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.acl')
      $null = & "$env:SystemRoot\System32\icacls.exe" $Path /save $saved /L /Q
      return @(Get-Content -LiteralPath $saved -Encoding Unicode | Where-Object { $_ -match '^[OGDS]:' })[0]
    }
  }

  BeforeEach {
    $global:WinkitJunctionCleanup = New-Object System.Collections.ArrayList
    $global:WinkitIcaclsCalls = New-Object System.Collections.ArrayList
    $global:WinkitProfileTestContext = @{
      NativeProcess = $script:NativeProcess
      TargetSid     = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    }
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $global:WinkitProfileTestContext.Source = Join-Path $script:Root 'staged user'
    $global:WinkitProfileTestContext.Target = Join-Path $script:Root 'new user'
    $script:Reports = Join-Path $script:Root 'reports'
    $script:Harness = Join-Path $script:Root 'Invoke-TestImport.ps1'
    $script:Documents = Join-Path $global:WinkitProfileTestContext.Target 'Documents'
    $script:Junction = Join-Path $script:Documents 'My Pictures'
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') -Text 'hello from backup'
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'NTUSER.DAT') -Text 'do not change'
    $null = New-Item -ItemType Directory -Path $script:Documents -Force
    New-SystemJunction -Path $script:Junction -Target (Join-Path $global:WinkitProfileTestContext.Target 'Pictures')
    [IO.File]::WriteAllText($script:Harness, $script:HarnessText, (New-Object Text.UTF8Encoding($true)))

    Mock Test-Elevation { $true }
    Mock Get-ImportProfile {
      [PSCustomObject]@{
        Path = $global:WinkitProfileTestContext.Target
        SID  = $global:WinkitProfileTestContext.TargetSid
        User = 'TEST\User1'
      }
    }
    Mock Get-ImportVolume {
      [PSCustomObject]@{ AvailableFreeSpace = [long]10GB }
    }
    Mock Assert-ImportPermission {}
    Mock Initialize-ImportReport {
      $null = New-Item -ItemType Directory -Path $Path
    }
    # Real robocopy and icacls on TestDrive only; every icacls call is recorded.
    Mock Invoke-SafeProcess {
      if ($FilePath -like '*\icacls.exe') {
        $null = $global:WinkitIcaclsCalls.Add(@($ArgumentList))
      }
      elseif ($FilePath -notlike '*\robocopy.exe') {
        throw "Unexpected native tool: $FilePath"
      }

      & $global:WinkitProfileTestContext.NativeProcess -FilePath $FilePath -ArgumentList $ArgumentList -AsResult
    }
  }

  AfterEach {
    foreach ($link in @($global:WinkitJunctionCleanup)) {
      if (Test-Path -LiteralPath $link) {
        [IO.Directory]::Delete($link)
      }
    }

    foreach ($name in @('WinkitProfileTestContext', 'WinkitJunctionCleanup', 'WinkitIcaclsCalls')) {
      Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    }
  }

  It 'lists them as SystemJunction with their target and accepts them in a dry run' {
    $inventory = Get-ImportInventory -Root $script:Documents -Destination -ProfileRoot $global:WinkitProfileTestContext.Target
    $run = Invoke-ProfileImport -Extra @{ DryRun = $true }

    $junction = @($inventory.Excluded | Where-Object Reason -EQ 'SystemJunction')
    $junction.Count | Should -Be 1
    $junction[0].Path | Should -Be $script:Junction
    $junction[0].Target | Should -Be (Join-Path $global:WinkitProfileTestContext.Target 'Pictures')
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    $run.Summary.Status | Should -Be 'Preview'
    @($run.Summary.Plan)[0].SystemJunctions | Should -Be @($script:Junction)
  }

  It 'imports and resets permissions while leaving the junctions untouched' {
    $before = Get-JunctionSddl -Path $script:Junction
    $attributesBefore = (Get-Item -LiteralPath $script:Junction -Force).Attributes

    $run = Invoke-ProfileImport

    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    $run.Summary.Status | Should -Be 'Completed'
    Get-Content -LiteralPath (Join-Path $script:Documents 'report [1].txt') | Should -Be 'hello from backup'
    # LinkType is empty on Windows PowerShell 5.1 for a junction that denies listing, so read the reparse data.
    (Get-ImportReparseInfo -Path $script:Junction).Kind | Should -Be 'Junction'
    Get-ImportLinkTarget -Item (Get-Item -LiteralPath $script:Junction -Force) | Should -Be (Join-Path $global:WinkitProfileTestContext.Target 'Pictures')
    (Get-Item -LiteralPath $script:Junction -Force).Attributes | Should -Be $attributesBefore
    Get-JunctionSddl -Path $script:Junction | Should -BeExactly $before
    $before | Should -Match '\(D;'

    # The root is never reset recursively, and the junction is never passed to a resetting call.
    $resets = @($global:WinkitIcaclsCalls | Where-Object { $_ -contains '/reset' -or $_ -contains '/setowner' })
    $resets.Count | Should -BeGreaterThan 0
    @($resets | Where-Object { $_[0] -eq $script:Documents -and $_ -contains '/T' }).Count | Should -Be 0
    @($resets | Where-Object { $_[0] -eq $script:Documents }).Count | Should -Be 2
    @($resets | Where-Object { $_[0] -eq $script:Junction }).Count | Should -Be 0
    @($resets | Where-Object { $_[0] -eq (Join-Path $script:Documents 'report [1].txt') -and $_ -contains '/T' }).Count | Should -Be 2
  }

  It 'still refuses <Case>' -ForEach @(
    @{ Case = 'a junction without Hidden and System'; Kind = 'Plain' }
    @{ Case = 'a system junction one level deeper'; Kind = 'Deeper' }
    @{ Case = 'a system junction pointing outside the profile'; Kind = 'Outside' }
  ) {
    switch ($Kind) {
      'Plain' {
        New-SystemJunction -Path (Join-Path $script:Documents 'Plain') -Target (Join-Path $global:WinkitProfileTestContext.Target 'Videos') -Plain
        $expected = 'without the Hidden and System attributes'
      }
      'Deeper' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:Documents 'Sub') -Force
        New-SystemJunction -Path (Join-Path $script:Documents 'Sub\My Music') -Target (Join-Path $global:WinkitProfileTestContext.Target 'Music')
        $expected = 'outside the allowed position'
      }
      'Outside' {
        New-SystemJunction -Path (Join-Path $script:Documents 'My Videos') -Target (Join-Path $script:Root 'elsewhere')
        $expected = 'not a folder inside the target profile'
      }
    }

    $run = Invoke-ProfileImport -Extra @{ DryRun = $true }

    $run.ExitCode | Should -Be 1
    $run.Summary.Error | Should -Match 'Unsupported destination link'
    $run.Summary.Error | Should -Match $expected
  }

  It 'still refuses a symbolic link in the destination' {
    $link = Join-Path $script:Documents 'Linked'
    try {
      $null = New-Item -ItemType SymbolicLink -Path $link -Target (Join-Path $global:WinkitProfileTestContext.Target 'Pictures') -ErrorAction Stop
      $null = $global:WinkitJunctionCleanup.Add($link)
    }
    catch {
      Set-ItResult -Skipped -Because 'creating a symbolic link needs Developer Mode or elevation'
      return
    }

    $run = Invoke-ProfileImport -Extra @{ DryRun = $true }

    $run.ExitCode | Should -Be 1
    $run.Summary.Error | Should -Match 'unsupported link \(SymbolicLink\)'
  }

  It 'stops before copying when a staged folder has the name of a system junction' {
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Source 'Documents\My Pictures\photo.txt') -Text 'staged'

    $run = Invoke-ProfileImport

    $run.ExitCode | Should -Be 1
    $run.Summary.Error | Should -Match 'same name as the Windows system junction'
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly -ParameterFilter { $ArgumentList -contains '/COPY:DAT' }
    Test-Path -LiteralPath (Join-Path $script:Documents 'report [1].txt') | Should -BeFalse
  }
}

Describe 'Import-UserProfileData dry run in a fresh session' {
  It 'prints no WhatIf lines while modules load and still reports Preview in <Engine>' -ForEach @(
    @{ Engine = 'pwsh' }
    @{ Engine = 'powershell.exe' }
  ) {
    # A fresh process has not loaded CimCmdlets yet; the profile lookup below autoloads it like the real one.
    $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $source = Join-Path $root 'staged user'
    $target = Join-Path $root 'new user'
    Write-ProfileFixture -Path (Join-Path $source 'Documents\report.txt') -Text 'data'
    Write-ProfileFixture -Path (Join-Path $target 'NTUSER.DAT') -Text 'hive'
    $definitions = @($script:Ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false) | ForEach-Object { $_.Extent.Text })
    $harness = Join-Path $root 'Invoke-FreshImport.ps1'
    [IO.File]::WriteAllText($harness, $script:HarnessText, (New-Object Text.UTF8Encoding($true)))
    $wrapper = @(
      'Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.4.1 -Force'
      ($definitions -join "`r`n`r`n")
      'function Test-Elevation { $true }'
      'function Get-ImportVolume { [PSCustomObject]@{ AvailableFreeSpace = [long]10GB } }'
      'function Assert-ImportPermission { }'
      ('function Get-ImportProfile {{ $null = Get-CimInstance -ClassName Win32_OperatingSystem; [PSCustomObject]@{{ Path = ''{0}''; SID = ''{1}''; User = ''TEST\User1'' }} }}' -f $target, [Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
      ("`$null = Get-Module CimCmdlets | ForEach-Object {{ throw 'CimCmdlets is already loaded.' }}")
      ("& '{0}' -SourcePath '{1}' -TargetUser 'TEST\User1' -ReportDirectory '{2}' -Folder Documents -DryRun *>&1 | ForEach-Object {{ [string]`$_ }}" -f $harness, $source, (Join-Path $root 'reports'))
    ) -join "`r`n"
    $wrapperPath = Join-Path $root 'Run-Fresh.ps1'
    [IO.File]::WriteAllText($wrapperPath, $wrapper, (New-Object Text.UTF8Encoding($true)))

    $output = (& $Engine -NoProfile -ExecutionPolicy Bypass -File $wrapperPath 2>&1 | ForEach-Object { [string]$_ }) -join "`n"

    $output | Should -Not -Match 'WhatIf:|What if:'
    $output | Should -Match 'Profile import: Preview'
    $output | Should -Match 'dry run: no report written'
    Test-Path -LiteralPath (Join-Path $root 'reports') | Should -BeFalse
  }
}

Describe 'Profile discovery guards' {
  BeforeEach {
    $global:WinkitProfileTestContext = @{
      Target = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
      SID    = 'S-1-5-21-1-2-3-1001'
      Loaded = $false
      Status = 0
    }
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'NTUSER.DAT')
    Mock Resolve-ImportIdentity {
      [PSCustomObject]@{ SID = $global:WinkitProfileTestContext.SID; User = 'TEST\User1' }
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{
        LocalPath = $global:WinkitProfileTestContext.Target
        Special   = $false
        Loaded    = $global:WinkitProfileTestContext.Loaded
        Status    = $global:WinkitProfileTestContext.Status
      }
    }
    Mock Get-ItemProperty {
      [PSCustomObject]@{ ProfileImagePath = $global:WinkitProfileTestContext.Target }
    }
    Mock Test-Path { $false } -ParameterFilter { $LiteralPath -like 'Registry::*' }
  }

  AfterEach {
    Remove-Variable -Name WinkitProfileTestContext -Scope Global -ErrorAction SilentlyContinue
  }

  It 'resolves the registered unloaded ordinary profile' {
    (Get-ImportProfile -User 'TEST\User1').SID | Should -Be 'S-1-5-21-1-2-3-1001'
  }

  It 'rejects CIM loaded state and a loaded registry hive independently' {
    $global:WinkitProfileTestContext.Loaded = $true
    { Get-ImportProfile -User 'TEST\User1' } | Should -Throw '*loaded*'
    $global:WinkitProfileTestContext.Loaded = $false
    Mock Test-Path { $true } -ParameterFilter { $LiteralPath -like 'Registry::*' }
    { Get-ImportProfile -User 'TEST\User1' } | Should -Throw '*loaded*'
  }

  It 'rejects temporary or corrupted profiles' {
    foreach ($status in @(1, 8)) {
      $global:WinkitProfileTestContext.Status = $status
      { Get-ImportProfile -User 'TEST\User1' } | Should -Throw '*ordinary local profile*'
    }
  }

  It 'rejects disagreement between ProfileList and CIM' {
    Mock Get-ItemProperty { [PSCustomObject]@{ ProfileImagePath = 'C:\OtherProfile' } }
    { Get-ImportProfile -User 'TEST\User1' } | Should -Throw '*same existing profile*'
  }
}

Describe 'Permission tools on disposable data' {
  It 'resets inherited ACLs, assigns the exact SID, and verifies Modify rights on scratch files' {
    $root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $data = Join-Path $root 'data'
    Write-ProfileFixture -Path (Join-Path $data 'child\file.txt')
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
    $acl = Get-Acl -LiteralPath $root
    $rule = New-Object Security.AccessControl.FileSystemAccessRule(
      $sid,
      [Security.AccessControl.FileSystemRights]::FullControl,
      [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
      [Security.AccessControl.PropagationFlags]::None,
      [Security.AccessControl.AccessControlType]::Allow
    )
    $acl.AddAccessRule($rule)
    Set-Acl -LiteralPath $root -AclObject $acl
    Assert-ImportPermission -Path $root -SID $sid.Value -RequireInheritance
    Invoke-ImportPermission -Path $data -SID $sid.Value -LogPrefix (Join-Path $root 'acl')
    Assert-ImportPermission -Path $data -SID $sid.Value -RequireOwner -RequireInheritance
    Assert-ImportPermission -Path (Join-Path $data 'child\file.txt') -SID $sid.Value -RequireOwner
  }

  It 'creates reports accessible only to the operator, administrators and SYSTEM' {
    $path = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    Initialize-ImportReport -Path $path
    $acl = Get-Acl -LiteralPath $path
    $acl.AreAccessRulesProtected | Should -BeTrue
    $rules = @($acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier]))
    $allowed = @('S-1-5-18', 'S-1-5-32-544', [Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
    foreach ($rule in $rules) {
      $rule.IdentityReference.Value | Should -BeIn $allowed
    }
  }
}
