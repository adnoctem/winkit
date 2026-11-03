#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Cross-script mock context is removed after each test; mocks run in the invoked script scope.')]
param ()

# Real robocopy runs only against TestDrive. Identity/profile discovery and ACL
# mutations are mocked in workflow tests. No Windows profile is created or changed.
BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.0 -Force
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
    $script:Content | Should -Match "RequiredVersion = '0.3.0'"
    $script:Content | Should -Match 'Server Core: supported'
    $script:Content | Should -Match 'SYSTEM-account execution: supported'
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    (Get-Help $script:ImportPath).Synopsis | Should -Match 'Imports staged user data'
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

  It 'previews without native processes, logs, copying, ACL changes, or cleanup' {
    $run = Invoke-ProfileImport -Extra @{ DryRun = $true; RemoveSource = $true }
    $run.ExitCode | Should -Be 0 -Because $run.Summary.Error
    $run.Summary.Status | Should -Be 'Preview'
    Test-Path -LiteralPath $script:Reports | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents') | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $global:WinkitProfileTestContext.Source 'Documents\report [1].txt') | Should -BeTrue
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly
    Should -Invoke Invoke-ImportPermission -Times 0 -Exactly
  }

  It 'refuses different destination content before making any changes' {
    Write-ProfileFixture -Path (Join-Path $global:WinkitProfileTestContext.Target 'Documents\report [1].txt') -Text 'newer work'
    $run = Invoke-ProfileImport
    $run.ExitCode | Should -Be 1
    $run.Summary.Error | Should -Match 'content conflict'
    Get-Content -LiteralPath (Join-Path $global:WinkitProfileTestContext.Target 'Documents\report [1].txt') | Should -Be 'newer work'
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly
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
      if ($ArgumentList[0] -like '*\Desktop') {
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
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly
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
    Should -Invoke Invoke-SafeProcess -Times 0 -Exactly
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
      Should -Invoke Invoke-SafeProcess -Times 0 -Exactly
    }
    finally {
      [IO.Directory]::Delete($link)
    }
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
