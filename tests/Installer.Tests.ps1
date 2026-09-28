#Requires -Version 5.1

BeforeAll {
  $repositoryRoot = Split-Path -Path $PSScriptRoot -Parent
  $script:InstallerPath = Join-Path -Path $repositoryRoot -ChildPath 'dist/install.ps1'
  $script:BuildPath = Join-Path -Path $repositoryRoot -ChildPath 'tools\build.ps1'
  # Load the actual private implementation to mock only OS boundaries. The HTTP
  # entry point has separate raw-source tests; there is no production test mode.
  $tokens = $null
  $parseErrors = $null
  $installerAst = [Management.Automation.Language.Parser]::ParseFile($script:InstallerPath, [ref]$tokens, [ref]$parseErrors)
  foreach ($definition in $installerAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    . ([scriptblock]::Create($definition.Extent.Text))
  }
  $script:ReadRegistration = ${function:Get-InstallRegistration}
  $script:WriteRegistration = ${function:Set-InstallRegistration}
  $script:DeleteRegistration = ${function:Remove-InstallRegistration}
  $script:NativeTestPath = Get-Command Test-Path -CommandType Cmdlet
  $script:NativeGetItem = Get-Command Get-Item -CommandType Cmdlet

  function Invoke-TestInstaller {
    Set-StrictMode -Version 2.0
    $ErrorActionPreference = 'Stop'
    Invoke-WinkitInstaller -Context (Get-InstallerConfiguration)
  }
  $script:InstallerEnvironmentNames = @(
    'WINKIT_SCOPE', 'WINKIT_INSTALL_PATH', 'WINKIT_REPOSITORY', 'WINKIT_VERSION',
    'WINKIT_NO_PATH', 'WINKIT_FORCE', 'WINKIT_NON_INTERACTIVE', 'WINKIT_DRY_RUN', 'WINKIT_PASS_THRU', 'WINKIT_UNINSTALL', 'WINKIT_CHANGE_SCOPE', 'DRY_RUN'
  )
  $script:OriginalEnvironment = @{}
  foreach ($name in $script:InstallerEnvironmentNames) {
    $script:OriginalEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
  }
  $script:DependencyCommands = @('Install-Module', 'Install-PackageProvider' | Where-Object {
      Get-Command -Name $_ -ErrorAction SilentlyContinue
    })

  function Invoke-TestInstallerPipeline {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingInvokeExpression', '', Justification = 'Regression test of the documented irm | iex entry point with mocked HTTP source, not a constructed command.')]
    [CmdletBinding()]
    param ()

    Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/adnoctem/winkit/main/dist/install.ps1' | Invoke-Expression 6>&1
  }

  function Invoke-TestReleaseArchive {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Version,

      [Parameter(Mandatory = $true)]
      [string]$Destination
    )

    $source = Join-Path -Path $TestDrive -ChildPath "source-$Version"
    $null = New-Item -Path (Join-Path -Path $source -ChildPath 'scripts') -ItemType Directory -Force
    $null = New-Item -Path (Join-Path -Path $source -ChildPath 'bin') -ItemType Directory -Force
    $null = New-Item -Path (Join-Path -Path $source -ChildPath 'resources') -ItemType Directory -Force
    $null = New-Item -Path (Join-Path -Path $source -ChildPath 'dist') -ItemType Directory -Force
    foreach ($name in @('install.ps1', 'uninstall.ps1', 'README.md')) {
      Copy-Item -LiteralPath (Join-Path (Split-Path $script:InstallerPath -Parent) $name) -Destination (Join-Path $source 'dist')
    }
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'scripts\Invoke-Bootstrap.ps1') -Value "# fixture $Version"
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'bin\bootstrap.cmd') -Value "rem fixture $Version"
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'resources\version.txt') -Value $Version
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'requirements.psd1') -Value "@{ PSFoundation = '1.4.0' }"
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'LICENSE') -Value 'MIT fixture'

    $items = @(
      Join-Path -Path $source -ChildPath 'scripts'
      Join-Path -Path $source -ChildPath 'bin'
      Join-Path -Path $source -ChildPath 'resources'
      Join-Path -Path $source -ChildPath 'dist'
      Join-Path -Path $source -ChildPath 'requirements.psd1'
      Join-Path -Path $source -ChildPath 'LICENSE'
    )
    Compress-Archive -Path $items -DestinationPath $Destination -Force
  }

  $env:WINKIT_TEST_RELEASE_VERSION = '1.0.0'
  $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-1.0.0.zip'
  Invoke-TestReleaseArchive -Version $env:WINKIT_TEST_RELEASE_VERSION -Destination $env:WINKIT_TEST_ARCHIVE

  Mock Invoke-RestMethod {
    [pscustomobject]@{
      tag_name     = "v$env:WINKIT_TEST_RELEASE_VERSION"
      published_at = '2026-01-01T00:00:00Z'
      assets       = @(
        [pscustomobject]@{
          name                 = 'winkit.zip'
          browser_download_url = 'https://example.invalid/winkit.zip'
        },
        [pscustomobject]@{
          name                 = 'CHECKSUMS_SHA256.txt'
          browser_download_url = 'https://example.invalid/CHECKSUMS_SHA256.txt'
        }
      )
    }
  }

  Mock Invoke-WebRequest {
    param (
      [string]$Uri,
      [string]$OutFile
    )

    if ($Uri.EndsWith('CHECKSUMS_SHA256.txt')) {
      $hash = (Get-FileHash -LiteralPath $env:WINKIT_TEST_ARCHIVE -Algorithm SHA256).Hash
      if ($env:WINKIT_TEST_BAD_CHECKSUM -eq '1') {
        $hash = '0' * 64
      }
      Set-Content -LiteralPath $OutFile -Value "$hash  winkit.zip"
    }
    else {
      Copy-Item -LiteralPath $env:WINKIT_TEST_ARCHIVE -Destination $OutFile
    }
  }

  Mock Get-Module {
    [pscustomobject]@{ Version = [version]'1.4.0' }
  } -ParameterFilter { $ListAvailable -and $Name -eq 'PSFoundation' }
}

Describe 'install.ps1' {
  BeforeEach {
    foreach ($settingName in $script:InstallerEnvironmentNames) {
      [Environment]::SetEnvironmentVariable($settingName, $null, 'Process')
    }
    $env:WINKIT_PASS_THRU = '1'
    $env:WINKIT_NO_PATH = '1'
    $env:WINKIT_NON_INTERACTIVE = '1'
    $script:Registrations = @{}
    $script:PathValues = @{ User = 'user-tools'; Machine = 'machine-tools'; Process = 'session-tools' }
    Mock Get-InstallRegistration { param($Scope) $script:Registrations[$Scope] }
    Mock Set-InstallRegistration { param($Record) $script:Registrations[$Record.Scope] = $Record }
    Mock Remove-InstallRegistration { param($Record) $script:Registrations.Remove($Record.Scope) }
    Mock Test-Administrator { $false }
    Mock Invoke-DependencyInstall { $false }
    Mock Get-InstallerPathValue { param($Target) $script:PathValues[$Target] }
    Mock Set-InstallerPathValue { param($Target, $Value) $script:PathValues[$Target] = $Value }
    Mock Enter-MaintenanceLock {
      $handle = [pscustomobject]@{}
      $handle | Add-Member -MemberType ScriptMethod -Name ReleaseMutex -Value {}
      $handle | Add-Member -MemberType ScriptMethod -Name Dispose -Value {}
      $handle
    }
    $env:WINKIT_TEST_RELEASE_VERSION = '1.0.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-1.0.0.zip'
    Remove-Item Env:WINKIT_TEST_BAD_CHECKSUM -ErrorAction SilentlyContinue
  }

  AfterAll {
    foreach ($name in $script:InstallerEnvironmentNames) {
      [Environment]::SetEnvironmentVariable($name, $script:OriginalEnvironment[$name], 'Process')
    }
    Remove-Item Env:WINKIT_TEST_RELEASE_VERSION -ErrorAction SilentlyContinue
    Remove-Item Env:WINKIT_TEST_ARCHIVE -ErrorAction SilentlyContinue
    Remove-Item Env:WINKIT_TEST_BAD_CHECKSUM -ErrorAction SilentlyContinue
  }

  It 'remains standalone and avoids dynamic command evaluation' {
    $content = Get-Content -LiteralPath $script:InstallerPath -Raw
    $content | Should -Not -Match '#Requires -Modules'
    $content | Should -Not -Match 'Import-Module PSFoundation'
    $content | Should -Not -Match 'Invoke-Expression'
    $content | Should -Match 'WINKIT_NON_INTERACTIVE'
  }

  It 'plans an install without creating the destination' {
    $env:WINKIT_DRY_RUN = '1'
    $destination = Join-Path -Path $TestDrive -ChildPath 'planned'
    $env:WINKIT_INSTALL_PATH = $destination
    $result = Invoke-TestInstaller

    $result.Status | Should -Be 'Planned'
    $result.Version | Should -Be '1.0.0'
    Test-Path -LiteralPath $destination | Should -BeFalse
  }

  Context 'environment configuration and HTTP entry point' {
    BeforeEach {
      Mock Invoke-WebRequest { throw 'Preview must not download release assets.' }
      foreach ($command in $script:DependencyCommands) {
        Mock $command { throw 'Preview must not install dependencies.' }
      }
      Mock New-Item { throw 'Preview must not create files or directories.' }
      Mock Move-Item { throw 'Preview must not move files or directories.' }
      Mock Read-Host { throw 'Preview must not prompt.' }

      $script:PreviewPath = [Environment]::GetEnvironmentVariable('Path', 'Process')
      $script:PreviewUserPath = [Environment]::GetEnvironmentVariable('Path', 'User')
      $script:PreviewMachinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
      $script:DownloadedInstaller = [Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes($script:InstallerPath))
      $env:WINKIT_DRY_RUN = '1'
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'environment-preview'
    }

    AfterEach {
      Should -Invoke Invoke-WebRequest -Times 0 -Exactly
      foreach ($command in $script:DependencyCommands) {
        Should -Invoke $command -Times 0 -Exactly
      }
      Should -Invoke New-Item -Times 0 -Exactly
      Should -Invoke Move-Item -Times 0 -Exactly
      Should -Invoke Read-Host -Times 0 -Exactly
      [Environment]::GetEnvironmentVariable('Path', 'Process') | Should -Be $script:PreviewPath
      [Environment]::GetEnvironmentVariable('Path', 'User') | Should -Be $script:PreviewUserPath
      [Environment]::GetEnvironmentVariable('Path', 'Machine') | Should -Be $script:PreviewMachinePath
    }

    It 'keeps HTTP source ASCII-only without a BOM and parses its raw decoded bytes' {
      $script:DownloadedInstaller | Should -Not -Match '[^\x00-\x7F]'
      $tokens = $null
      $parseErrors = $null
      $null = [Management.Automation.Language.Parser]::ParseInput($script:DownloadedInstaller, [ref]$tokens, [ref]$parseErrors)
      $parseErrors.Count | Should -Be 0
    }

    It 'previews the normal irm and iex pipeline with default destinations' {
      $env:WINKIT_SCOPE = 'CurrentUser'
      Remove-Item Env:WINKIT_INSTALL_PATH, Env:WINKIT_NO_PATH, Env:WINKIT_PASS_THRU, Env:WINKIT_NON_INTERACTIVE
      Mock Test-Path { $false }
      Mock Invoke-RestMethod { $script:DownloadedInstaller } -ParameterFilter { $Uri -like 'https://raw.githubusercontent.com/*' }

      $messages = Invoke-TestInstallerPipeline
      $text = $messages | Out-String

      $text | Should -Match 'DRY RUN: Install winkit 1.0.0'
      $text | Should -Match ([regex]::Escape((Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Programs\winkit\scripts')))
      $text | Should -Match 'https://example.invalid/winkit.zip'
      $text | Should -Match 'https://example.invalid/CHECKSUMS_SHA256.txt'
      $text | Should -Match 'temporary staging'
      $text | Should -Match 'PSFoundation.*PSGallery.*CurrentUser'
      $text | Should -Match 'user PATH'
    }

    It 'uses a custom path, repository, and version from the environment' {
      $env:WINKIT_REPOSITORY = 'example/winkit'
      $env:WINKIT_VERSION = 'v1.0.0'
      $result = Invoke-TestInstaller

      $result.Status | Should -Be 'Planned'
      $result.InstallPath | Should -Be $env:WINKIT_INSTALL_PATH
      Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://api.github.com/repos/example/winkit/releases/tags/v1.0.0' }
    }

    It 'honors WINKIT_DRY_RUN=<Value> even with force and non-interactive mode' -ForEach @(
      @{ Value = '1' }, @{ Value = 'TRUE' }, @{ Value = ' yes ' }, @{ Value = 'on' }
    ) {
      $env:WINKIT_DRY_RUN = $Value
      $env:WINKIT_FORCE = '1'
      $result = Invoke-TestInstaller

      $result.Status | Should -Be 'Planned'
      Test-Path -LiteralPath $env:WINKIT_INSTALL_PATH | Should -BeFalse
    }

    It 'previews AllUsers and disabled PATH updates without prompting' {
      $env:WINKIT_SCOPE = 'AllUsers'
      $env:WINKIT_NON_INTERACTIVE = '0'
      $output = Invoke-TestInstaller 6>&1
      $result = $output | Where-Object { $_.PSTypeNames -contains 'winkit.InstallationResult' }
      $text = $output | Out-String

      $result.Status | Should -Be 'Planned'
      $result.Scope | Should -Be 'AllUsers'
      $text | Should -Match 'PATH update disabled by WINKIT_NO_PATH'
      $text | Should -Match 'Actual AllUsers installation or scope change requires an elevated'
    }

    It 'accepts false Boolean values and does not return a result when pass-through is disabled' -ForEach @(
      @{ Value = '0' }, @{ Value = 'false' }, @{ Value = 'NO' }, @{ Value = 'off' }, @{ Value = ' ' }
    ) {
      $env:WINKIT_NO_PATH = $Value
      $env:WINKIT_FORCE = $Value
      $env:WINKIT_NON_INTERACTIVE = $Value
      $env:WINKIT_PASS_THRU = $Value
      $result = Invoke-TestInstaller

      $result | Should -BeNullOrEmpty
    }

    It 'rejects invalid <Name> before requesting a release' -ForEach @(
      @{ Name = 'WINKIT_SCOPE'; Value = 'Everywhere' }
      @{ Name = 'WINKIT_REPOSITORY'; Value = 'https://example.invalid/repo' }
      @{ Name = 'WINKIT_VERSION'; Value = 'not-a-version' }
      @{ Name = 'WINKIT_NO_PATH'; Value = 'sometimes' }
      @{ Name = 'WINKIT_FORCE'; Value = 'sometimes' }
      @{ Name = 'WINKIT_NON_INTERACTIVE'; Value = 'sometimes' }
      @{ Name = 'WINKIT_DRY_RUN'; Value = 'sometimes' }
      @{ Name = 'WINKIT_PASS_THRU'; Value = 'sometimes' }
    ) {
      [Environment]::SetEnvironmentVariable($Name, $Value, 'Process')

      { Invoke-TestInstaller } | Should -Throw "*$Name must*"
      Should -Invoke Invoke-RestMethod -Times 0 -Exactly
    }

    It 'rejects legacy script parameters rather than ignoring them' {
      { & $script:InstallerPath -DryRun } | Should -Throw '*environment variables, not command-line parameters*'
      Should -Invoke Invoke-RestMethod -Times 0 -Exactly
    }
  }

  It 'refuses to replace an unrecognized directory even with Force' {
    $env:WINKIT_FORCE = '1'
    $destination = Join-Path -Path $TestDrive -ChildPath 'unrecognized'
    $env:WINKIT_INSTALL_PATH = $destination
    $null = New-Item -Path $destination -ItemType Directory
    Set-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'user-file.txt') -Value 'keep'

    { Invoke-TestInstaller } |
      Should -Throw '*Refusing to overwrite an unrecognized directory*'
    Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'user-file.txt') | Should -Be 'keep'
  }

  It 'installs a verified release and updates it through the same entry point' {
    $env:WINKIT_DRY_RUN = '0'
    $env:DRY_RUN = '1' # Unprefixed settings belong to other applications.
    $destination = Join-Path -Path $TestDrive -ChildPath 'managed'
    $env:WINKIT_INSTALL_PATH = $destination
    $first = Invoke-TestInstaller

    $first.Status | Should -Be 'Installed'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.0.0'

    $env:WINKIT_TEST_RELEASE_VERSION = '1.1.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-1.1.0.zip'
    Invoke-TestReleaseArchive -Version $env:WINKIT_TEST_RELEASE_VERSION -Destination $env:WINKIT_TEST_ARCHIVE
    $second = Invoke-TestInstaller

    $second.Status | Should -Be 'Updated'
    $second.PreviousVersion | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.1.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.1.0'
    @(Get-ChildItem -LiteralPath $TestDrive -Force | Where-Object { $_.Name -match 'install-lock|\.winkit-install-|\.backup-' }).Count | Should -Be 0

    $third = Invoke-TestInstaller
    $third.Status | Should -Be 'Current'
  }

  It 'preserves the installed release when checksum verification fails' {
    $destination = Join-Path -Path $TestDrive -ChildPath 'checksum-rollback'
    $env:WINKIT_INSTALL_PATH = $destination
    $null = Invoke-TestInstaller

    $env:WINKIT_TEST_RELEASE_VERSION = '1.1.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-checksum-1.1.0.zip'
    Invoke-TestReleaseArchive -Version $env:WINKIT_TEST_RELEASE_VERSION -Destination $env:WINKIT_TEST_ARCHIVE
    $env:WINKIT_TEST_BAD_CHECKSUM = '1'

    { Invoke-TestInstaller } |
      Should -Throw '*checksum does not match*'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.0.0'
  }

  It 'rejects an invalid non-interactive environment value' {
    $env:WINKIT_NON_INTERACTIVE = 'sometimes'
    $destination = Join-Path -Path $TestDrive -ChildPath 'invalid-environment'
    $env:WINKIT_INSTALL_PATH = $destination

    { Invoke-TestInstaller } |
      Should -Throw '*WINKIT_NON_INTERACTIVE must be*'
  }

  Context 'registered maintenance and rollback' {
    BeforeEach {
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive ('owned-' + [guid]::NewGuid().ToString('N'))
      $script:Installed = Invoke-TestInstaller
      $script:OwnedPath = $env:WINKIT_INSTALL_PATH
      $script:OwnedRecord = $script:Registrations.CurrentUser
    }

    It 'discovers a custom registered path without environment overrides' {
      Remove-Item Env:WINKIT_INSTALL_PATH, Env:WINKIT_NO_PATH
      $result = Invoke-TestInstaller
      $result.Status | Should -Be 'Current'
      $result.InstallPath | Should -Be $script:OwnedPath
      $script:Registrations.CurrentUser.NoPath | Should -BeTrue
    }

    It 'preserves scope during elevated updates and honors an explicit CurrentUser choice' {
      Mock Test-Administrator { $true }
      $result = Invoke-TestInstaller
      $result.Scope | Should -Be 'CurrentUser'
      $script:Registrations.ContainsKey('AllUsers') | Should -BeFalse
    }

    It 'persists a version pin and allows resetting it to latest' {
      $env:WINKIT_VERSION = 'v1.0.0'
      $null = Invoke-TestInstaller
      Remove-Item Env:WINKIT_VERSION
      $null = Invoke-TestInstaller
      $script:Registrations.CurrentUser.Version | Should -Be 'v1.0.0'
      $env:WINKIT_VERSION = 'latest'
      $null = Invoke-TestInstaller
      $script:Registrations.CurrentUser.Version | Should -Be ''
    }

    It 'rejects an explicit path mismatch without falling back to the registration' {
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'wrong-target'
      $env:WINKIT_UNINSTALL = '1'
      Remove-Item Env:WINKIT_NO_PATH
      { Invoke-TestInstaller } | Should -Throw '*does not match*'
      Test-Path $script:OwnedPath | Should -BeTrue
    }

    It 'requires an explicit scope-change operation' {
      $env:WINKIT_SCOPE = 'AllUsers'
      { Invoke-TestInstaller } | Should -Throw '*WINKIT_CHANGE_SCOPE*'
    }

    It 'blocks removal when files are <Change>' -ForEach @(
      @{ Change = 'added' }, @{ Change = 'modified' }, @{ Change = 'missing' }
    ) {
      $file = Join-Path $script:OwnedPath 'resources\version.txt'
      if ($Change -eq 'added') {
        Set-Content (Join-Path $script:OwnedPath 'personal.txt') 'keep'
      }
      elseif ($Change -eq 'modified') {
        Set-Content $file 'changed'
      }
      else {
        Remove-Item -LiteralPath $file
      }
      $env:WINKIT_UNINSTALL = '1'
      Remove-Item Env:WINKIT_NO_PATH
      { Invoke-TestInstaller } | Should -Throw
      Test-Path $script:OwnedPath | Should -BeTrue
      $script:Registrations.CurrentUser.InstallId | Should -Be $script:OwnedRecord.InstallId
    }

    It 'previews removal without network, mutation, or confirmation' {
      $env:WINKIT_UNINSTALL = '1'
      $env:WINKIT_DRY_RUN = '1'
      $env:WINKIT_NON_INTERACTIVE = '0'
      Remove-Item Env:WINKIT_NO_PATH
      Mock Invoke-RestMethod { throw 'Uninstall must not resolve releases.' }
      Mock Read-Host { throw 'Preview must not prompt.' }
      $result = Invoke-TestInstaller
      $result.Status | Should -Be 'Planned'
      Test-Path $script:OwnedPath | Should -BeTrue
      Should -Invoke Read-Host -Times 0
    }

    It 'requires affirmative interactive removal confirmation, independently of ConfirmPreference' {
      $env:WINKIT_UNINSTALL = '1'
      $env:WINKIT_NON_INTERACTIVE = '0'
      Remove-Item Env:WINKIT_NO_PATH
      $ConfirmPreference = 'None'
      Mock Read-Host { '' }
      $result = Invoke-TestInstaller
      $result.Status | Should -Be 'Skipped'
      Test-Path $script:OwnedPath | Should -BeTrue
      Should -Invoke Read-Host -Times 1 -Exactly
    }

    It 'removes owned files, registration, and only matching PATH entries offline' {
      $env:WINKIT_UNINSTALL = '1'
      Remove-Item Env:WINKIT_NO_PATH
      $bin = Join-Path $script:OwnedPath 'bin'
      $script:PathValues.User = "keep;$bin;$bin-extra"
      $script:PathValues.Process = "session;$bin"
      Mock Invoke-RestMethod { throw 'Uninstall must not resolve releases.' }
      $result = Invoke-TestInstaller
      $result.Status | Should -Be 'Uninstalled'
      Test-Path $script:OwnedPath | Should -BeFalse
      $script:Registrations.Count | Should -Be 0
      $script:PathValues.User | Should -Be "keep;$bin-extra"
      $script:PathValues.Machine | Should -Be 'machine-tools'
      $script:PathValues.Process | Should -Be 'session'
      Should -Invoke Invoke-DependencyInstall -Times 1 -Exactly # Initial fixture installation only.
    }

    It 'restores files, registry, and PATH after uninstall metadata failure' {
      $env:WINKIT_UNINSTALL = '1'
      Remove-Item Env:WINKIT_NO_PATH
      $script:PathValues.User = Join-Path $script:OwnedPath 'bin'
      Mock Remove-InstallRegistration { throw 'registry removal failure' }
      { Invoke-TestInstaller } | Should -Throw '*registry removal failure*'
      Test-Path $script:OwnedPath | Should -BeTrue
      $script:Registrations.CurrentUser.InstallId | Should -Be $script:OwnedRecord.InstallId
      $script:PathValues.User | Should -Be (Join-Path $script:OwnedPath 'bin')
      Assert-OwnedInstallation $script:Registrations.CurrentUser
    }

    It 'retains the original release after update registration failure' {
      $env:WINKIT_TEST_RELEASE_VERSION = '1.1.0'
      $env:WINKIT_TEST_ARCHIVE = Join-Path $TestDrive 'failed-update.zip'
      Invoke-TestReleaseArchive '1.1.0' $env:WINKIT_TEST_ARCHIVE
      Mock Set-InstallRegistration {
        param($Record)
        $script:Registrations[$Record.Scope] = $Record
        if ($Record.InstalledVersion -eq '1.1.0') { throw 'registration write failure' }
      }
      { Invoke-TestInstaller } | Should -Throw '*registration write failure*'
      $script:Registrations.CurrentUser.InstalledVersion | Should -Be '1.0.0'
      Get-Content (Join-Path $script:OwnedPath 'resources\version.txt') | Should -Be '1.0.0'
      Assert-OwnedInstallation $script:Registrations.CurrentUser
    }

    It 'changes scope locally without updating the release or leaving the source registered' {
      Mock Test-Administrator { $true }
      Mock Invoke-RestMethod { throw 'Scope change must not resolve releases.' }
      $env:WINKIT_CHANGE_SCOPE = '1'
      $env:WINKIT_SCOPE = 'AllUsers'
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'machine-destination'
      $result = Invoke-TestInstaller
      $result.Status | Should -Be 'ScopeChanged'
      $result.Version | Should -Be '1.0.0'
      Test-Path $script:OwnedPath | Should -BeFalse
      $script:Registrations.ContainsKey('CurrentUser') | Should -BeFalse
      $script:Registrations.AllUsers.InstallId | Should -Be $script:OwnedRecord.InstallId
      Assert-OwnedInstallation $script:Registrations.AllUsers
      Should -Invoke Invoke-DependencyInstall -Times 1 -Exactly -ParameterFilter { $InstallScope -eq 'AllUsers' }
    }

    It 'rolls back a scope change when source registration removal fails' {
      Mock Test-Administrator { $true }
      Mock Remove-InstallRegistration {
        param($Record)
        if ($Record.Scope -eq 'CurrentUser') { throw 'source registry failure' }
        $script:Registrations.Remove($Record.Scope)
      }
      $env:WINKIT_CHANGE_SCOPE = '1'
      $env:WINKIT_SCOPE = 'AllUsers'
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'rollback-destination'
      { Invoke-TestInstaller } | Should -Throw '*source registry failure*'
      Test-Path $script:OwnedPath | Should -BeTrue
      Test-Path $env:WINKIT_INSTALL_PATH | Should -BeFalse
      $script:Registrations.ContainsKey('AllUsers') | Should -BeFalse
      Assert-OwnedInstallation $script:Registrations.CurrentUser
    }

    It 'supports changing back from AllUsers to CurrentUser' {
      Mock Test-Administrator { $true }
      $env:WINKIT_CHANGE_SCOPE = '1'
      $env:WINKIT_SCOPE = 'AllUsers'
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'reverse-machine'
      $null = Invoke-TestInstaller
      $machinePath = $env:WINKIT_INSTALL_PATH
      $env:WINKIT_SCOPE = 'CurrentUser'
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'reverse-user'
      $result = Invoke-TestInstaller
      $result.Status | Should -Be 'ScopeChanged'
      Test-Path $machinePath | Should -BeFalse
      $script:Registrations.ContainsKey('AllUsers') | Should -BeFalse
      Assert-OwnedInstallation $script:Registrations.CurrentUser
    }

    It 'does not select an arbitrary installation when both scopes are registered' {
      $script:Registrations.AllUsers = $script:OwnedRecord | Select-Object *
      $script:Registrations.AllUsers.Scope = 'AllUsers'
      Remove-Item Env:WINKIT_INSTALL_PATH
      { Invoke-TestInstaller } | Should -Throw '*Multiple installations match*'
      Test-Path $script:OwnedPath | Should -BeTrue
    }

    It 'retains rollback backups when the original path becomes occupied' {
      $env:WINKIT_UNINSTALL = '1'
      Remove-Item Env:WINKIT_NO_PATH
      Mock Remove-InstallRegistration {
        param($Record)
        $null = New-Item -Path $Record.InstallPath -ItemType Directory
        Set-Content (Join-Path $Record.InstallPath 'new-file.txt') 'keep'
        throw 'simulated concurrent recreation'
      }
      { Invoke-TestInstaller } | Should -Throw '*Rollback destination is occupied*'
      Get-Content (Join-Path $script:OwnedPath 'new-file.txt') | Should -Be 'keep'
      $backups = @(Get-ChildItem -LiteralPath $TestDrive -Directory -Filter '.winkit-remove-*')
      $backups.Count | Should -Be 1
      Test-Path (Join-Path $backups[0].FullName 'scripts\Invoke-Bootstrap.ps1') | Should -BeTrue
    }

    It 'reports cleanup remaining after committed removal' {
      $env:WINKIT_UNINSTALL = '1'
      Remove-Item Env:WINKIT_NO_PATH
      Mock Invoke-OwnedDirectoryCleanup { throw 'locked file' } -ParameterFilter { $Path -like '*\.winkit-remove-*' }
      $result = Invoke-TestInstaller -WarningAction SilentlyContinue
      $result.Status | Should -Be 'CleanupPending'
      $result.CleanupPaths.Count | Should -Be 1
      Test-Path $result.CleanupPaths[0] | Should -BeTrue
      $script:Registrations.Count | Should -Be 0
    }
  }

  It 'defaults a fresh elevated installation to AllUsers and announces the choice' {
    Mock Test-Administrator { $true }
    $env:WINKIT_DRY_RUN = '1'
    $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'admin-default'
    $output = Invoke-TestInstaller 6>&1
    ($output | Out-String) | Should -Match 'Elevated session detected; defaulting to AllUsers'
    ($output | Where-Object { $_.PSTypeNames -contains 'winkit.InstallationResult' }).Scope | Should -Be 'AllUsers'
  }

  It 'honors explicit CurrentUser during a fresh elevated installation' {
    Mock Test-Administrator { $true }
    $env:WINKIT_SCOPE = 'CurrentUser'
    $env:WINKIT_DRY_RUN = '1'
    $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'explicit-user'
    (Invoke-TestInstaller).Scope | Should -Be 'CurrentUser'
  }

  It 'retains ordinary installation confirmation behavior without a removal prompt' {
    $env:WINKIT_NON_INTERACTIVE = '0'
    $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'ordinary-confirmation'
    $ConfirmPreference = 'High'
    Mock Read-Host { throw 'Installation must not use the removal prompt.' }
    (Invoke-TestInstaller).Status | Should -Be 'Installed'
    Should -Invoke Read-Host -Times 0
  }

  It 'reads ownership manifests at non-ASCII paths in Windows PowerShell' {
    $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive ('user-' + [char]0x00fc)
    $null = Invoke-TestInstaller
    (Get-InstallState $env:WINKIT_INSTALL_PATH).installPath | Should -Be $env:WINKIT_INSTALL_PATH
    (Invoke-TestInstaller).Status | Should -Be 'Current'
  }

  Context 'registry schema and local uninstaller' {
    BeforeEach {
      $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive ('schema-' + [guid]::NewGuid().ToString('N'))
      $null = Invoke-TestInstaller
      $script:SchemaRecord = $script:Registrations.CurrentUser
      $script:RegistryValues = @{}
      $script:RegistryTypes = @{}
      Mock Test-Path {
        param($LiteralPath, $Path, $PathType)
        $parameters = @{}
        if ($LiteralPath) { $parameters.LiteralPath = $LiteralPath }
        else { $parameters.Path = $Path }
        if ($PathType) { $parameters.PathType = $PathType }
        & $script:NativeTestPath @parameters
      }
      Mock Get-Item { param($LiteralPath) & $script:NativeGetItem -LiteralPath $LiteralPath -Force }
      Mock Test-Path { param($LiteralPath) $script:RegistryValues.ContainsKey([string]$LiteralPath) } -ParameterFilter { $LiteralPath -like 'HK*:\Software\*' }
      Mock New-Item {
        param($Path)
        $registryPath = [string]@($Path)[0]
        if (-not $script:RegistryValues.ContainsKey($registryPath)) { $script:RegistryValues[$registryPath] = @{} }
      } -ParameterFilter { $Path -like 'HK*:\Software\*' }
      Mock New-ItemProperty {
        param($LiteralPath, $Name, $Value, $PropertyType)
        $script:RegistryValues[[string]$LiteralPath][$Name] = $Value
        $script:RegistryTypes["$LiteralPath|$Name"] = $PropertyType
      } -ParameterFilter { $LiteralPath -like 'HK*:\Software\*' }
      Mock Get-ItemProperty { param($LiteralPath) [pscustomobject]$script:RegistryValues[[string]$LiteralPath] } -ParameterFilter { $LiteralPath -like 'HK*:\Software\*' }
      Mock Remove-ItemProperty { param($LiteralPath, $Name) $script:RegistryValues[[string]$LiteralPath].Remove($Name) } -ParameterFilter { $LiteralPath -like 'HK*:\Software\*' }
      Mock Remove-Item { param($LiteralPath) $script:RegistryValues.Remove([string]$LiteralPath) } -ParameterFilter { $LiteralPath -like 'HK*:\Software\*' }
      Mock Get-Item {
        param($LiteralPath)
        [pscustomobject]@{
          ValueCount  = $script:RegistryValues[[string]$LiteralPath].Count
          SubKeyCount = @($script:RegistryValues.Keys | Where-Object { $_.StartsWith("$LiteralPath\") }).Count
        }
      } -ParameterFilter { $LiteralPath -like 'HK*:\Software\*' }
    }

    It 'writes PascalCase persistent options and conventional uninstall registration without Settings or runtime flags' {
      & $script:WriteRegistration $script:SchemaRecord
      $paths = Get-RegistrationPath CurrentUser
      $script:RegistryValues[$paths.Options].Keys | Sort-Object | Should -Be @('InstallPath', 'NoPath', 'Repository', 'Scope', 'Version')
      $script:RegistryTypes["$($paths.Options)|NoPath"] | Should -Be 'DWord'
      $script:RegistryValues[$paths.Uninstall].Publisher | Should -Be 'AdNoctem'
      $script:RegistryValues[$paths.Uninstall].UninstallString | Should -Match '-File ".*\\dist\\uninstall.ps1"$'
      $script:RegistryValues.ContainsKey($paths.Application + '\Settings') | Should -BeFalse
      $record = & $script:ReadRegistration CurrentUser
      $record.InstallId | Should -Be $script:SchemaRecord.InstallId
    }

    It 'rejects stale identity and incomplete registration' {
      & $script:WriteRegistration $script:SchemaRecord
      $paths = Get-RegistrationPath CurrentUser
      $script:RegistryValues[$paths.Application].InstallId = [guid]::NewGuid().ToString()
      { & $script:ReadRegistration CurrentUser } | Should -Throw '*Invalid winkit registration*'
      $script:RegistryValues.Remove($paths.Options)
      { & $script:ReadRegistration CurrentUser } | Should -Throw '*Incomplete*'
    }

    It 'removes only installation registration and preserves unrelated settings' {
      & $script:WriteRegistration $script:SchemaRecord
      $paths = Get-RegistrationPath CurrentUser
      $script:RegistryValues[$paths.Application + '\Settings'] = @{ Example = 'keep' }
      & $script:DeleteRegistration $script:SchemaRecord
      $script:RegistryValues.ContainsKey($paths.Options) | Should -BeFalse
      $script:RegistryValues.ContainsKey($paths.Uninstall) | Should -BeFalse
      $script:RegistryValues[$paths.Application + '\Settings'].Example | Should -Be 'keep'
    }

    It 'does not recreate existing keys when refreshing registration' {
      & $script:WriteRegistration $script:SchemaRecord
      $paths = Get-RegistrationPath CurrentUser
      $script:RegistryValues[$paths.Application + '\Settings'] = @{ Example = 'keep' }
      & $script:WriteRegistration $script:SchemaRecord
      Should -Invoke New-Item -Times 3 -Exactly -ParameterFilter { $Path -like 'HKCU:\Software\*' }
      $script:RegistryValues[$paths.Application + '\Settings'].Example | Should -Be 'keep'
    }

    It 'runs the local uninstaller offline and pins its own target while restoring caller settings' {
      & $script:WriteRegistration $script:SchemaRecord
      $local = Join-Path $script:SchemaRecord.InstallPath 'dist\uninstall.ps1'
      # A fixture engine observes the wrapper's environment. The real engine's
      # removal workflow is exercised above with OS boundaries isolated.
      $engine = Join-Path $script:SchemaRecord.InstallPath 'dist\install.ps1'
      Set-Content -LiteralPath $engine -Value '[pscustomobject]@{ Status = ''Planned''; InstallPath = $env:WINKIT_INSTALL_PATH; Scope = $env:WINKIT_SCOPE; Uninstall = $env:WINKIT_UNINSTALL; ChangeScope = $env:WINKIT_CHANGE_SCOPE }'
      $state = $script:SchemaRecord.State
      ($state.files | Where-Object Path -EQ 'dist\install.ps1').Sha256 = (Get-FileHash $engine -Algorithm SHA256).Hash
      Invoke-StateWrite -Destination $script:SchemaRecord.InstallPath -State $state
      $env:WINKIT_INSTALL_PATH = 'C:\unrelated-path'
      $env:WINKIT_SCOPE = 'AllUsers'
      $env:WINKIT_FORCE = '1'
      $env:WINKIT_CHANGE_SCOPE = '1'
      $env:WINKIT_DRY_RUN = '1'
      Mock Invoke-RestMethod { throw 'Local uninstall must not contact GitHub.' }
      $result = & $local
      $result.Status | Should -Be 'Planned'
      $result.InstallPath | Should -Be $script:SchemaRecord.InstallPath
      $result.Scope | Should -Be 'CurrentUser'
      $result.Uninstall | Should -Be '1'
      $result.ChangeScope | Should -BeNullOrEmpty
      $env:WINKIT_INSTALL_PATH | Should -Be 'C:\unrelated-path'
      $env:WINKIT_SCOPE | Should -Be 'AllUsers'
      $env:WINKIT_FORCE | Should -Be '1'
      $env:WINKIT_CHANGE_SCOPE | Should -Be '1'
    }

    It 'refuses to execute an altered local maintenance engine' {
      $engine = Join-Path $script:SchemaRecord.InstallPath 'dist\install.ps1'
      Add-Content -LiteralPath $engine '# changed'
      { & (Join-Path $script:SchemaRecord.InstallPath 'dist\uninstall.ps1') } | Should -Throw '*does not match the ownership manifest*'
    }
  }
}

Describe 'release bundle' {
  It 'contains the runtime requirements and license with a matching checksum' {
    $output = Join-Path -Path $TestDrive -ChildPath 'dist'
    $null = & $script:BuildPath -OutputDirectory $output -Format Zip
    $zipPath = Join-Path -Path $output -ChildPath 'winkit.zip'
    $checksumPath = Join-Path -Path $output -ChildPath 'CHECKSUMS_SHA256.txt'

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
      $entryNames = @($archive.Entries.FullName.Replace('\', '/'))
    }
    finally {
      $archive.Dispose()
    }

    $entryNames | Should -Contain 'requirements.psd1'
    $entryNames | Should -Contain 'LICENSE'
    $entryNames | Should -Contain 'dist/install.ps1'
    $entryNames | Should -Contain 'dist/uninstall.ps1'
    $entryNames | Should -Contain 'dist/README.md'
    $expected = ((Get-Content -LiteralPath $checksumPath) -split '\s+')[0]
    (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash | Should -Be $expected
  }
}
