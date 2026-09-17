#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Tier-0 logic tests for the Group Policy pipeline (Export-LocalPolicy ->
# Build-GroupPolicyBackup -> Import-GroupPolicyBackup). Every script runs end to
# end against fixtures under TestDrive; no elevation, domain, RSAT, or LGPO.exe
# is required. The LGPO text rules asserted here were verified against LGPO 3.0.

BeforeAll {
  Import-Module PSFoundation -MinimumVersion 1.4.0 -Force

  $script:PolicyScripts = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '..\..\scripts\Policy'))
  $script:ExportScript = Join-Path $script:PolicyScripts 'Export-LocalPolicy.ps1'
  $script:BuildScript = Join-Path $script:PolicyScripts 'Build-GroupPolicyBackup.ps1'
  $script:ImportScript = Join-Path $script:PolicyScripts 'Import-GroupPolicyBackup.ps1'
  $script:Unexportable = @('**SecureKey', '**soft.Soft')

  function Initialize-PolicyFixture {
    param ([string]$Root)

    $_key = 'Software\Policies\winkit\Fixture'
    $_machine = @(
      [PSCustomObject]@{ Key = $_key; ValueName = 'Dword'; Type = 4; Data = [uint32]4294967295 }
      [PSCustomObject]@{ Key = $_key; ValueName = 'Qword'; Type = 11; Data = [uint64]5000000000 }
      [PSCustomObject]@{ Key = $_key; ValueName = 'SzPath'; Type = 1; Data = 'C:\path\file' }
      [PSCustomObject]@{ Key = $_key; ValueName = 'SzLiteralZero'; Type = 1; Data = 'a\0b' }
      [PSCustomObject]@{ Key = $_key; ValueName = 'SzTrailing'; Type = 1; Data = 'x; "q" ' }
      [PSCustomObject]@{ Key = $_key; ValueName = 'NonAscii'; Type = 1; Data = "Gr$([char]0xFC)$([char]0xDF)e $([char]0x2713)" }
      [PSCustomObject]@{ Key = $_key; ValueName = 'Expand'; Type = 2; Data = '%SystemRoot%\x' }
      [PSCustomObject]@{ Key = $_key; ValueName = 'Multi'; Type = 7; Data = @('x\y', 'z z') }
      [PSCustomObject]@{ Key = $_key; ValueName = 'Binary'; Type = 3; Data = [byte[]](0, 255, 16) }
      [PSCustomObject]@{ Key = $_key; ValueName = ''; Type = 1; Data = 'default value' }
      [PSCustomObject]@{ Key = "$_key\Sub"; ValueName = '**delvals.'; Type = 1; Data = ' ' }
      [PSCustomObject]@{ Key = "$_key\Sub"; ValueName = 'AfterDelvals'; Type = 4; Data = [uint32]1 }
      [PSCustomObject]@{ Key = "$_key\Sub"; ValueName = '**del.Gone'; Type = 1; Data = ' ' }
      [PSCustomObject]@{ Key = "$_key\Created"; ValueName = ''; Type = 0; Data = $null }
      [PSCustomObject]@{ Key = $_key; ValueName = '**DeleteKeys'; Type = 1; Data = 'K1;K2' }
      [PSCustomObject]@{ Key = $_key; ValueName = '**SecureKey'; Type = 4; Data = [uint32]1 }
      [PSCustomObject]@{ Key = 'Software\Policies\winkit\Other'; ValueName = '**soft.Soft'; Type = 4; Data = [uint32]1 }
    )
    $_user = @(
      [PSCustomObject]@{ Key = 'Software\Policies\winkit\UserSide'; ValueName = 'Flag'; Type = 4; Data = [uint32]7 }
    )

    $null = New-Item -ItemType Directory -Path (Join-Path $Root 'Machine'), (Join-Path $Root 'User') -Force
    ConvertTo-RegistryPolicy -InputObject $_machine -Path (Join-Path $Root 'Machine\registry.pol') -Force
    ConvertTo-RegistryPolicy -InputObject $_user -Path (Join-Path $Root 'User\registry.pol') -Force
  }

  function Write-PolicySource {
    param ([string]$Path, [string[]]$Lines, [switch]$NoByteOrderMark)

    $null = New-Item -ItemType Directory -Path (Split-Path -Path $Path -Parent) -Force
    $_encoding = New-Object System.Text.UTF8Encoding(-not $NoByteOrderMark)
    [System.IO.File]::WriteAllText($Path, (($Lines -join "`r`n") + "`r`n"), $_encoding)
  }

  function Get-RecordSignature {
    param ($Records)
    @($Records | ForEach-Object { "$($_.Key.ToLowerInvariant())|$($_.ValueName)|$($_.Type)|$([Convert]::ToBase64String([byte[]]@($_.Data)))" })
  }

  function Get-BackupDirectory {
    param ([string]$Root)
    @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\{[0-9A-F-]{36}\}$' })
  }
}

Describe 'Policy script conventions' {
  BeforeAll {
    $script:Scripts = @(Get-ChildItem -LiteralPath $script:PolicyScripts -Filter '*.ps1' -File | Sort-Object Name)
  }

  It 'contains the expected script set' {
    @($script:Scripts | ForEach-Object { $_.Name }) | Should -Be @(
      'Build-GroupPolicyBackup.ps1',
      'Export-LocalPolicy.ps1',
      'Import-GroupPolicyBackup.ps1'
    )
  }

  It 'pins PSFoundation 1.4.0, which provides the registry.pol converters' {
    foreach ($_script in $script:Scripts) {
      $_header = Get-Content -LiteralPath $_script.FullName -TotalCount 3
      $_header -match '^#Requires -Version' | Should -Not -BeNullOrEmpty -Because "$($_script.Name) must declare the PowerShell version"
      $_header -match "^#Requires -Modules @\{ ModuleName = 'PSFoundation'; ModuleVersion = '1\.4\.0' \}" | Should -Not -BeNullOrEmpty -Because "$($_script.Name) must pin PSFoundation 1.4.0"
    }
  }

  It 'documents .SYNOPSIS and a complete .NOTES block' {
    foreach ($_script in $script:Scripts) {
      $_content = Get-Content -LiteralPath $_script.FullName -Raw
      $_content -match 'Import-Module PSFoundation -Force' | Should -BeTrue -Because "$($_script.Name) must import PSFoundation"
      (Get-Help $_script.FullName).Synopsis | Should -Not -Match '\.ps1' -Because "$($_script.Name) comment-based help must be discoverable, not auto-generated syntax"
      $_content -match 'Author: MVProwess' | Should -BeTrue
      $_content -match 'License: MIT' | Should -BeTrue
      $_content -match 'Server Core:' | Should -BeTrue -Because "$($_script.Name) .NOTES needs a Server Core note"
      $_content -match 'SYSTEM-account execution:' | Should -BeTrue -Because "$($_script.Name) .NOTES needs a SYSTEM-account note"
    }
  }
}

Describe 'Export-LocalPolicy' {
  BeforeAll {
    $script:PolicyRoot = Join-Path $TestDrive 'export\GroupPolicy'
    Initialize-PolicyFixture -Root $script:PolicyRoot
    $script:ExportPath = Join-Path $TestDrive 'export\out\local.txt'
    $script:ExportResults = & $script:ExportScript -PolicyRoot $script:PolicyRoot -Path $script:ExportPath -PassThru 6>$null
  }

  It 'writes UTF-8 with a byte-order mark, CRLF line endings, and a TODO metadata header' {
    $_bytes = [System.IO.File]::ReadAllBytes($script:ExportPath)
    $_bytes[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
    $_text = [System.Text.Encoding]::UTF8.GetString($_bytes)
    ([regex]::Matches($_text, "(?<!`r)`n")).Count | Should -Be 0
    $_text | Should -Match '; Title: TODO'
    $_text | Should -Match '; Justification: TODO'
  }

  It 'groups entries under their top-most captured key and reports each group' {
    $_captures = @($script:ExportResults | Where-Object { $_.Action -eq 'Capture' -and $_.Status -eq 'Completed' })
    @($_captures | ForEach-Object { $_.Target }) | Should -Contain 'Software\Policies\winkit\Fixture'
    @($_captures | ForEach-Object { $_.Target }) | Should -Not -Contain 'Software\Policies\winkit\Fixture\Sub'
    @($_captures | ForEach-Object { $_.Target }) | Should -Contain 'Software\Policies\winkit\UserSide'
  }

  It 'warns about records that have no LGPO text form and comments them out' {
    @($script:ExportResults | Where-Object { $_.Status -eq 'Warn' }).Count | Should -Be 2
    $_text = Get-Content -LiteralPath $script:ExportPath -Raw
    ([regex]::Matches($_text, '; NOT EXPORTED')).Count | Should -Be 2
    $_text | Should -Not -Match '(?m)^\*\*SecureKey'
  }

  It 'refuses to overwrite an existing file without -Force' {
    $global:LASTEXITCODE = 0
    $_results = & $script:ExportScript -PolicyRoot $script:PolicyRoot -Path $script:ExportPath -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object { $_.Action -eq 'Write' }).Status | Should -Be 'Failed'
  }

  It 'writes nothing when there is no local policy' {
    $_emptyRoot = Join-Path $TestDrive 'export\empty'
    $null = New-Item -ItemType Directory -Path $_emptyRoot -Force
    $_path = Join-Path $TestDrive 'export\out\empty.txt'
    $_results = & $script:ExportScript -PolicyRoot $_emptyRoot -Path $_path -PassThru 6>$null
    ($_results | Where-Object { $_.Action -eq 'Write' }).Status | Should -Be 'Skipped'
    Test-Path -LiteralPath $_path | Should -BeFalse
  }

  It 'does not write the file in a dry run' {
    $_path = Join-Path $TestDrive 'export\out\dry.txt'
    $null = & $script:ExportScript -PolicyRoot $script:PolicyRoot -Path $_path -DryRun 6>$null
    Test-Path -LiteralPath $_path | Should -BeFalse
  }
}

Describe 'Build-GroupPolicyBackup' {
  BeforeAll {
    $script:SourceRoot = Join-Path $TestDrive 'build\GroupPolicy'
    Initialize-PolicyFixture -Root $script:SourceRoot
    $script:SourceDir = Join-Path $TestDrive 'build\src'
    $null = & $script:ExportScript -PolicyRoot $script:SourceRoot -Path (Join-Path $script:SourceDir '10-export.txt') 6>$null
    $script:OutputDir = Join-Path $TestDrive 'build\out'
    $script:BuildResults = & $script:BuildScript -SourcePath $script:SourceDir -OutputPath $script:OutputDir -DisplayName 'winkit Test Baseline' -PassThru 6>$null
    $script:Backup = Get-BackupDirectory -Root $script:OutputDir | Select-Object -First 1
  }

  It 'compiles an exported policy back into identical registry.pol records' {
    $script:Backup | Should -Not -BeNullOrEmpty
    foreach ($_side in @('Machine', 'User')) {
      $_expected = @(ConvertFrom-RegistryPolicy -Path (Join-Path $script:SourceRoot "$_side\registry.pol") -Raw | Where-Object { $_.ValueName -notin $script:Unexportable })
      $_actual = @(ConvertFrom-RegistryPolicy -Path (Join-Path $script:Backup.FullName "DomainSysvol\GPO\$_side\registry.pol") -Raw)
      $_actual.Count | Should -Be $_expected.Count -Because "$_side record count"
      foreach ($_group in ($_expected | Group-Object { $_.Key.ToLowerInvariant() })) {
        $_actualGroup = @($_actual | Where-Object { $_.Key.ToLowerInvariant() -eq $_group.Name })
        (Get-RecordSignature $_actualGroup) | Should -Be (Get-RecordSignature $_group.Group) -Because "records and their order under $($_group.Name)"
      }
    }
  }

  It 'stamps well-formed backup metadata with no placeholders or real domain data left' {
    $_backupText = Get-Content -LiteralPath (Join-Path $script:Backup.FullName 'Backup.xml') -Raw
    $_infoText = Get-Content -LiteralPath (Join-Path $script:Backup.FullName 'bkupInfo.xml') -Raw
    "$_backupText$_infoText" | Should -Not -Match '\{\{\w+\}\}'
    { [xml]$_backupText } | Should -Not -Throw
    ([xml]$_infoText).GetElementsByTagName('ID')[0].InnerText | Should -Be $script:Backup.Name
    ([xml]$_infoText).GetElementsByTagName('GPODisplayName')[0].InnerText | Should -Be 'winkit Test Baseline'

    $_descriptor = ([xml]$_backupText).GroupPolicyBackupScheme.GroupPolicyObject.GroupPolicyCoreSettings.SecurityDescriptor.Trim()
    $_bytes = [byte[]]($_descriptor -split '\s+' | ForEach-Object { [Convert]::ToByte($_, 16) })
    { New-Object System.Security.AccessControl.RawSecurityDescriptor -ArgumentList $_bytes, 0 } | Should -Not -Throw
  }

  It 'registers the Registry client-side extension and file entries for each side with settings' {
    [xml]$_backup = Get-Content -LiteralPath (Join-Path $script:Backup.FullName 'Backup.xml') -Raw
    $_backup.GetElementsByTagName('MachineExtensionGuids')[0].InnerText | Should -Be '[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F72-3407-48AE-BA88-E8213C6761F1}]'
    $_backup.GetElementsByTagName('UserExtensionGuids')[0].InnerText | Should -Be '[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F73-3407-48AE-BA88-E8213C6761F1}]'
    $_locations = @($_backup.GetElementsByTagName('FSObjectFile') | ForEach-Object { $_.GetAttribute('Location', 'http://www.microsoft.com/GroupPolicy/GPOOperations') } | Where-Object { $_ })
    $_locations | Should -Contain 'DomainSysvol\GPO\Machine\registry.pol'
    $_locations | Should -Contain 'DomainSysvol\GPO\User\registry.pol'
  }

  It 'leaves a side without settings unregistered' {
    $_source = Join-Path $TestDrive 'build\machine-only'
    Write-PolicySource -Path (Join-Path $_source 'policy.txt') -Lines @('Computer', 'Software\Policies\winkit\MachineOnly', 'Flag', 'DWORD:1')
    $_output = Join-Path $TestDrive 'build\machine-only-out'
    $null = & $script:BuildScript -SourcePath $_source -OutputPath $_output 6>$null
    $_backup = Get-BackupDirectory -Root $_output | Select-Object -First 1
    [xml]$_xml = Get-Content -LiteralPath (Join-Path $_backup.FullName 'Backup.xml') -Raw
    $_xml.GetElementsByTagName('UserExtensionGuids')[0].InnerText | Should -BeNullOrEmpty
    $_xml.GetElementsByTagName('UserVersionNumber')[0].InnerText | Should -Be '0'
    Test-Path -LiteralPath (Join-Path $_backup.FullName 'DomainSysvol\GPO\User\registry.pol') | Should -BeFalse
  }

  It 'adds each build to manifest.xml' {
    $null = & $script:BuildScript -SourcePath $script:SourceDir -OutputPath $script:OutputDir 6>$null
    [xml]$_manifest = Get-Content -LiteralPath (Join-Path $script:OutputDir 'manifest.xml') -Raw
    @($_manifest.GetElementsByTagName('BackupInst')).Count | Should -Be 2
    (Get-BackupDirectory -Root $script:OutputDir).Count | Should -Be 2
  }

  It 'translates LGPO text actions into registry.pol records' {
    $_source = Join-Path $TestDrive 'build\actions'
    Write-PolicySource -Path (Join-Path $_source 'actions.txt') -Lines @(
      '; comment', '',
      'computer', 'Software\Policies\winkit\A', 'Hex', 'DWORD:0x10', '',
      'Computer', 'Software\Policies\winkit\A', 'Negative', 'DWORD:-1', '',
      'Computer', 'Software\Policies\winkit\A', 'Escaped', 'SZ:C:\\path\\0x', '',
      'Computer', 'Software\Policies\winkit\A', 'List', 'MULTISZ:one\0two', '',
      'Computer', 'Software\Policies\winkit\A', '', 'SZ:default', '',
      'Computer', 'Software\Policies\winkit\A', 'Gone', 'DELETE', '',
      'Computer', 'Software\Policies\winkit\A', '*', 'DELETEALLVALUES', '',
      'Computer', 'Software\Policies\winkit\A\New', '*', 'CREATEKEY', '',
      'Computer', 'Software\Policies\winkit\A', 'K1;K2', 'DELETEKEYS', '',
      'Computer', 'Software\Policies\winkit\A', 'Cleared', 'CLEAR'
    )
    $_output = Join-Path $TestDrive 'build\actions-out'
    $null = & $script:BuildScript -SourcePath $_source -OutputPath $_output 6>$null
    $_backup = Get-BackupDirectory -Root $_output | Select-Object -First 1
    $_records = @(ConvertFrom-RegistryPolicy -Path (Join-Path $_backup.FullName 'DomainSysvol\GPO\Machine\registry.pol'))

    $_records.Count | Should -Be 9
    ($_records | Where-Object ValueName -EQ 'Hex').Data | Should -Be 16
    ($_records | Where-Object ValueName -EQ 'Negative').Data | Should -Be 4294967295
    ($_records | Where-Object ValueName -EQ 'Escaped').Data | Should -Be 'C:\path\0x'
    ($_records | Where-Object ValueName -EQ 'List').Data | Should -Be @('one', 'two')
    ($_records | Where-Object { $_.ValueName -eq '' -and $_.Type -eq 1 }).Data | Should -Be 'default'
    @($_records | ForEach-Object { $_.ValueName }) | Should -Contain '**del.Gone'
    @($_records | ForEach-Object { $_.ValueName }) | Should -Contain '**delvals.'
    ($_records | Where-Object ValueName -EQ '**DeleteKeys').Data | Should -Be 'K1;K2'
    ($_records | Where-Object { $_.Key -like '*\New' }).Type | Should -Be 0
    @($_records | ForEach-Object { $_.ValueName }) | Should -Not -Contain 'Cleared'
  }

  It 'fails without writing a backup when a source <Case>' -TestCases @(
    @{ Case = 'uses a lowercase action (LGPO rejects it)'; Lines = @('Computer', 'Software\Policies\winkit\X', 'V', 'dword:1') }
    @{ Case = 'overflows a DWORD (LGPO wraps it silently)'; Lines = @('Computer', 'Software\Policies\winkit\X', 'V', 'DWORD:4294967296') }
    @{ Case = 'embeds \0 in a string (LGPO stores a NUL)'; Lines = @('Computer', 'Software\Policies\winkit\X', 'V', 'SZ:C:\0data') }
    @{ Case = 'contains a whitespace-only line (LGPO rejects it)'; Lines = @('   ', 'Computer', 'Software\Policies\winkit\X', 'V', 'DWORD:1') }
    @{ Case = 'ends in an incomplete entry'; Lines = @('Computer', 'Software\Policies\winkit\X', 'V') }
  ) {
    $_source = Join-Path $TestDrive "build\invalid-$([guid]::NewGuid().ToString('N'))"
    Write-PolicySource -Path (Join-Path $_source 'bad.txt') -Lines $Lines
    $_output = Join-Path $_source 'out'
    $global:LASTEXITCODE = 0
    $_results = & $script:BuildScript -SourcePath $_source -OutputPath $_output -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    @($_results | Where-Object { $_.Status -eq 'Failed' }).Count | Should -BeGreaterThan 0
    (Get-BackupDirectory -Root $_output).Count | Should -Be 0
  }

  It 'rejects non-ASCII data in a file without a byte-order mark but allows it in comments' {
    $_source = Join-Path $TestDrive 'build\bomless'
    Write-PolicySource -NoByteOrderMark -Path (Join-Path $_source 'comment.txt') -Lines @("; Gr$([char]0xFC)$([char]0xDF)e $([char]0x2014) comment", 'Computer', 'Software\Policies\winkit\X', 'V', 'DWORD:1')
    $null = & $script:BuildScript -SourcePath $_source -OutputPath (Join-Path $_source 'out') 6>$null
    (Get-BackupDirectory -Root (Join-Path $_source 'out')).Count | Should -Be 1

    $_badSource = Join-Path $TestDrive 'build\bomless-data'
    Write-PolicySource -NoByteOrderMark -Path (Join-Path $_badSource 'data.txt') -Lines @('Computer', 'Software\Policies\winkit\X', 'V', "SZ:Gr$([char]0xFC)$([char]0xDF)e")
    $global:LASTEXITCODE = 0
    $null = & $script:BuildScript -SourcePath $_badSource -OutputPath (Join-Path $_badSource 'out') 6>$null
    $global:LASTEXITCODE | Should -Be 1
  }

  It 'skips example sources by default and fails when nothing else is left' {
    $_source = Join-Path $TestDrive 'build\example-only'
    Write-PolicySource -Path (Join-Path $_source '00-example.txt') -Lines @('Computer', 'Software\Policies\winkit\X', 'V', 'DWORD:1')
    $global:LASTEXITCODE = 0
    $_results = & $script:BuildScript -SourcePath $_source -OutputPath (Join-Path $_source 'out') -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object Target -EQ '00-example.txt').Status | Should -Be 'Skipped'
  }

  It 'warns when two sources set the same value differently' {
    $_source = Join-Path $TestDrive 'build\conflict'
    Write-PolicySource -Path (Join-Path $_source '01-a.txt') -Lines @('Computer', 'Software\Policies\winkit\X', 'V', 'DWORD:1')
    Write-PolicySource -Path (Join-Path $_source '02-b.txt') -Lines @('Computer', 'Software\Policies\winkit\X', 'V', 'DWORD:2')
    $_results = & $script:BuildScript -SourcePath $_source -OutputPath (Join-Path $_source 'out') -PassThru 6>$null
    @($_results | Where-Object { $_.Status -eq 'Warn' }).Count | Should -Be 1
  }

  It 'validates without writing in a dry run' {
    $_output = Join-Path $TestDrive 'build\dry-out'
    $_results = & $script:BuildScript -SourcePath $script:SourceDir -OutputPath $_output -DryRun 6>$null
    ($_results | Where-Object Action -EQ 'Build').Status | Should -Be 'Skipped'
    Test-Path -LiteralPath $_output | Should -BeFalse
  }
}

Describe 'Import-GroupPolicyBackup' {
  BeforeAll {
    $script:ImportSource = Join-Path $TestDrive 'import\src'
    Write-PolicySource -Path (Join-Path $script:ImportSource 'policy.txt') -Lines @('Computer', 'Software\Policies\winkit\Import', 'Flag', 'DWORD:1')
    $script:SingleRoot = Join-Path $TestDrive 'import\single'
    $null = & $script:BuildScript -SourcePath $script:ImportSource -OutputPath $script:SingleRoot 6>$null
    $script:MissingLgpo = Join-Path $TestDrive 'import\no-lgpo\LGPO.exe'
  }

  It 'previews a local import of the only backup without elevation or LGPO.exe' {
    $global:LASTEXITCODE = 0
    $_results = & $script:ImportScript -Path $script:SingleRoot -LgpoPath $script:MissingLgpo -DryRun 6>$null
    $global:LASTEXITCODE | Should -Be 0
    ($_results | Where-Object Action -EQ 'Validate').Status | Should -Be 'Completed'
    ($_results | Where-Object Action -EQ 'ResolveLgpo').Status | Should -Be 'Warn'
    ($_results | Where-Object Action -EQ 'Import').Status | Should -Be 'Skipped'
  }

  It 'refuses to choose between several backups without -BackupId' {
    $_root = Join-Path $TestDrive 'import\multiple'
    $null = & $script:BuildScript -SourcePath $script:ImportSource -OutputPath $_root 6>$null
    $null = & $script:BuildScript -SourcePath $script:ImportSource -OutputPath $_root 6>$null
    $global:LASTEXITCODE = 0
    $_results = & $script:ImportScript -Path $_root -DryRun 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object Action -EQ 'Select').Status | Should -Be 'Failed'

    $_id = (Get-BackupDirectory -Root $_root | Select-Object -First 1).Name
    $_selected = & $script:ImportScript -Path $_root -BackupId $_id -LgpoPath $script:MissingLgpo -DryRun 6>$null
    ($_selected | Where-Object Action -EQ 'Import').Status | Should -Be 'Skipped'
  }

  It 'rejects a backup whose settings are not registered with the Registry extension' {
    $_root = Join-Path $TestDrive 'import\unregistered'
    Copy-Item -LiteralPath $script:SingleRoot -Destination $_root -Recurse
    $_backupXml = Join-Path (Get-BackupDirectory -Root $_root | Select-Object -First 1).FullName 'Backup.xml'
    $_text = [System.IO.File]::ReadAllText($_backupXml) -replace '<MachineExtensionGuids>.*?</MachineExtensionGuids>', '<MachineExtensionGuids></MachineExtensionGuids>'
    [System.IO.File]::WriteAllText($_backupXml, $_text)
    $global:LASTEXITCODE = 0
    $_results = & $script:ImportScript -Path $_root -DryRun 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object Action -EQ 'Validate').Status | Should -Be 'Failed'
  }

  It 'rejects domain-only parameters for a local import' {
    $global:LASTEXITCODE = 0
    $_results = & $script:ImportScript -Path $script:SingleRoot -LinkTarget 'OU=Test,DC=example,DC=com' -DryRun 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object Action -EQ 'Validate').Status | Should -Be 'Failed'
  }

  It 'previews a domain import and link without applying anything' {
    $global:LASTEXITCODE = 0
    $_results = & $script:ImportScript -Path $script:SingleRoot -Target Domain -DisplayName 'winkit Test' -CreateIfNeeded -LinkTarget 'OU=Test,DC=example,DC=com' -DryRun 6>$null
    $global:LASTEXITCODE | Should -Be 0
    ($_results | Where-Object Action -EQ 'Import').Status | Should -Contain 'Skipped'
    ($_results | Where-Object Action -EQ 'Link').Status | Should -Be 'Skipped'
  }
}
