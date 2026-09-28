#Requires -Version 5.1

BeforeAll {
  $repositoryRoot = Split-Path -Path $PSScriptRoot -Parent
  $script:InstallerPath = Join-Path -Path $repositoryRoot -ChildPath 'install.ps1'
  $script:BuildPath = Join-Path -Path $repositoryRoot -ChildPath 'tools\build.ps1'
  $script:InstallerEnvironmentNames = @(
    'WINKIT_SCOPE', 'WINKIT_INSTALL_PATH', 'WINKIT_REPOSITORY', 'WINKIT_VERSION',
    'WINKIT_NO_PATH', 'WINKIT_FORCE', 'WINKIT_NON_INTERACTIVE', 'WINKIT_DRY_RUN', 'WINKIT_PASS_THRU', 'DRY_RUN'
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

    Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/adnoctem/winkit/main/install.ps1' | Invoke-Expression 6>&1
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
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'scripts\Invoke-Bootstrap.ps1') -Value "# fixture $Version"
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'bin\bootstrap.cmd') -Value "rem fixture $Version"
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'resources\version.txt') -Value $Version
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'requirements.psd1') -Value "@{ PSFoundation = '1.4.0' }"
    Set-Content -LiteralPath (Join-Path -Path $source -ChildPath 'LICENSE') -Value 'MIT fixture'

    $items = @(
      Join-Path -Path $source -ChildPath 'scripts'
      Join-Path -Path $source -ChildPath 'bin'
      Join-Path -Path $source -ChildPath 'resources'
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
    $result = & $script:InstallerPath

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
      $result = & $script:InstallerPath

      $result.Status | Should -Be 'Planned'
      $result.InstallPath | Should -Be $env:WINKIT_INSTALL_PATH
      Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -eq 'https://api.github.com/repos/example/winkit/releases/tags/v1.0.0' }
    }

    It 'honors WINKIT_DRY_RUN=<Value> even with force and non-interactive mode' -ForEach @(
      @{ Value = '1' }, @{ Value = 'TRUE' }, @{ Value = ' yes ' }, @{ Value = 'on' }
    ) {
      $env:WINKIT_DRY_RUN = $Value
      $env:WINKIT_FORCE = '1'
      $result = & $script:InstallerPath

      $result.Status | Should -Be 'Planned'
      Test-Path -LiteralPath $env:WINKIT_INSTALL_PATH | Should -BeFalse
    }

    It 'previews AllUsers and disabled PATH updates without prompting' {
      $env:WINKIT_SCOPE = 'AllUsers'
      $env:WINKIT_NON_INTERACTIVE = '0'
      $output = & $script:InstallerPath 6>&1
      $result = $output | Where-Object { $_.PSTypeNames -contains 'winkit.InstallationResult' }
      $text = $output | Out-String

      $result.Status | Should -Be 'Planned'
      $result.Scope | Should -Be 'AllUsers'
      $text | Should -Match 'PATH update disabled by WINKIT_NO_PATH'
      $text | Should -Match 'Actual AllUsers installation requires an elevated'
    }

    It 'accepts false Boolean values and does not return a result when pass-through is disabled' -ForEach @(
      @{ Value = '0' }, @{ Value = 'false' }, @{ Value = 'NO' }, @{ Value = 'off' }, @{ Value = ' ' }
    ) {
      $env:WINKIT_NO_PATH = $Value
      $env:WINKIT_FORCE = $Value
      $env:WINKIT_NON_INTERACTIVE = $Value
      $env:WINKIT_PASS_THRU = $Value
      $result = & $script:InstallerPath

      $result | Should -BeNullOrEmpty
    }

    It 'rejects invalid <Name> before requesting a release' -ForEach @(
      @{ Name = 'WINKIT_SCOPE'; Value = 'Everywhere' }
      @{ Name = 'WINKIT_REPOSITORY'; Value = 'https://example.invalid/repo' }
      @{ Name = 'WINKIT_VERSION'; Value = 'latest' }
      @{ Name = 'WINKIT_NO_PATH'; Value = 'sometimes' }
      @{ Name = 'WINKIT_FORCE'; Value = 'sometimes' }
      @{ Name = 'WINKIT_NON_INTERACTIVE'; Value = 'sometimes' }
      @{ Name = 'WINKIT_DRY_RUN'; Value = 'sometimes' }
      @{ Name = 'WINKIT_PASS_THRU'; Value = 'sometimes' }
    ) {
      [Environment]::SetEnvironmentVariable($Name, $Value, 'Process')

      { & $script:InstallerPath } | Should -Throw "*$Name must*"
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

    { & $script:InstallerPath } |
      Should -Throw '*Refusing to overwrite an unrecognized directory*'
    Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'user-file.txt') | Should -Be 'keep'
  }

  It 'installs a verified release and updates it through the same entry point' {
    $env:WINKIT_DRY_RUN = '0'
    $env:DRY_RUN = '1' # Unprefixed settings belong to other applications.
    $destination = Join-Path -Path $TestDrive -ChildPath 'managed'
    $env:WINKIT_INSTALL_PATH = $destination
    $first = & $script:InstallerPath

    $first.Status | Should -Be 'Installed'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.0.0'

    $env:WINKIT_TEST_RELEASE_VERSION = '1.1.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-1.1.0.zip'
    Invoke-TestReleaseArchive -Version $env:WINKIT_TEST_RELEASE_VERSION -Destination $env:WINKIT_TEST_ARCHIVE
    $second = & $script:InstallerPath

    $second.Status | Should -Be 'Updated'
    $second.PreviousVersion | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.1.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.1.0'
    @(Get-ChildItem -LiteralPath $TestDrive -Force | Where-Object { $_.Name -match 'install-lock|\.winkit-install-|\.backup-' }).Count | Should -Be 0

    $third = & $script:InstallerPath
    $third.Status | Should -Be 'Current'
  }

  It 'preserves the installed release when checksum verification fails' {
    $destination = Join-Path -Path $TestDrive -ChildPath 'checksum-rollback'
    $env:WINKIT_INSTALL_PATH = $destination
    $null = & $script:InstallerPath

    $env:WINKIT_TEST_RELEASE_VERSION = '1.1.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-checksum-1.1.0.zip'
    Invoke-TestReleaseArchive -Version $env:WINKIT_TEST_RELEASE_VERSION -Destination $env:WINKIT_TEST_ARCHIVE
    $env:WINKIT_TEST_BAD_CHECKSUM = '1'

    { & $script:InstallerPath } |
      Should -Throw '*checksum does not match*'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.0.0'
  }

  It 'installs through irm and iex using environment configuration' {
    $env:WINKIT_INSTALL_PATH = Join-Path $TestDrive 'pipeline-install'
    $env:WINKIT_NON_INTERACTIVE = '0'
    $ConfirmPreference = 'High'
    $sourceText = [Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes($script:InstallerPath))
    Mock Invoke-RestMethod { $sourceText } -ParameterFilter { $Uri -like 'https://raw.githubusercontent.com/*' }

    $output = Invoke-TestInstallerPipeline
    $result = $output | Where-Object { $_.PSTypeNames -contains 'winkit.InstallationResult' }

    $result.Status | Should -Be 'Installed'
    $result.InstallPath | Should -Be $env:WINKIT_INSTALL_PATH
    Test-Path (Join-Path $env:WINKIT_INSTALL_PATH '.winkit-install.json') | Should -BeTrue
  }

  It 'rejects an invalid non-interactive environment value' {
    $env:WINKIT_NON_INTERACTIVE = 'sometimes'
    $destination = Join-Path -Path $TestDrive -ChildPath 'invalid-environment'
    $env:WINKIT_INSTALL_PATH = $destination

    { & $script:InstallerPath } |
      Should -Throw '*WINKIT_NON_INTERACTIVE must be*'
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
    $expected = ((Get-Content -LiteralPath $checksumPath) -split '\s+')[0]
    (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash | Should -Be $expected
  }
}
