#Requires -Version 5.1

BeforeAll {
  $repositoryRoot = Split-Path -Path $PSScriptRoot -Parent
  $script:InstallerPath = Join-Path -Path $repositoryRoot -ChildPath 'install.ps1'
  $script:BuildPath = Join-Path -Path $repositoryRoot -ChildPath 'tools\build.ps1'

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
    $env:WINKIT_TEST_RELEASE_VERSION = '1.0.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-1.0.0.zip'
    Remove-Item Env:WINKIT_TEST_BAD_CHECKSUM -ErrorAction SilentlyContinue
    Remove-Item Env:WINKIT_NON_INTERACTIVE -ErrorAction SilentlyContinue
  }

  AfterAll {
    Remove-Item Env:WINKIT_NON_INTERACTIVE -ErrorAction SilentlyContinue
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
    $destination = Join-Path -Path $TestDrive -ChildPath 'planned'
    $result = & $script:InstallerPath -InstallPath $destination -DryRun -PassThru

    $result.Status | Should -Be 'Planned'
    $result.Version | Should -Be '1.0.0'
    Test-Path -LiteralPath $destination | Should -BeFalse
  }

  It 'refuses to replace an unrecognized directory even with Force' {
    $destination = Join-Path -Path $TestDrive -ChildPath 'unrecognized'
    $null = New-Item -Path $destination -ItemType Directory
    Set-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'user-file.txt') -Value 'keep'

    { & $script:InstallerPath -InstallPath $destination -Force -NonInteractive } |
      Should -Throw '*Refusing to overwrite an unrecognized directory*'
    Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'user-file.txt') | Should -Be 'keep'
  }

  It 'installs a verified release and updates it through the same entry point' {
    $destination = Join-Path -Path $TestDrive -ChildPath 'managed'
    $first = & $script:InstallerPath -InstallPath $destination -NoPath -NonInteractive -PassThru

    $first.Status | Should -Be 'Installed'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.0.0'

    $env:WINKIT_TEST_RELEASE_VERSION = '1.1.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-1.1.0.zip'
    Invoke-TestReleaseArchive -Version $env:WINKIT_TEST_RELEASE_VERSION -Destination $env:WINKIT_TEST_ARCHIVE
    $second = & $script:InstallerPath -InstallPath $destination -NoPath -NonInteractive -PassThru

    $second.Status | Should -Be 'Updated'
    $second.PreviousVersion | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.1.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.1.0'
    @(Get-ChildItem -LiteralPath $TestDrive -Force | Where-Object { $_.Name -match 'install-lock|\.winkit-install-|\.backup-' }).Count | Should -Be 0

    $third = & $script:InstallerPath -InstallPath $destination -NoPath -NonInteractive -PassThru
    $third.Status | Should -Be 'Current'
  }

  It 'preserves the installed release when checksum verification fails' {
    $destination = Join-Path -Path $TestDrive -ChildPath 'checksum-rollback'
    $null = & $script:InstallerPath -InstallPath $destination -NoPath -NonInteractive -PassThru

    $env:WINKIT_TEST_RELEASE_VERSION = '1.1.0'
    $env:WINKIT_TEST_ARCHIVE = Join-Path -Path $TestDrive -ChildPath 'winkit-checksum-1.1.0.zip'
    Invoke-TestReleaseArchive -Version $env:WINKIT_TEST_RELEASE_VERSION -Destination $env:WINKIT_TEST_ARCHIVE
    $env:WINKIT_TEST_BAD_CHECKSUM = '1'

    { & $script:InstallerPath -InstallPath $destination -NoPath -NonInteractive } |
      Should -Throw '*checksum does not match*'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath 'resources\version.txt')) | Should -Be '1.0.0'
    (Get-Content -LiteralPath (Join-Path -Path $destination -ChildPath '.winkit-install.json') -Raw | ConvertFrom-Json).version | Should -Be '1.0.0'
  }

  It 'rejects an invalid non-interactive environment value' {
    $env:WINKIT_NON_INTERACTIVE = 'sometimes'
    $destination = Join-Path -Path $TestDrive -ChildPath 'invalid-environment'

    { & $script:InstallerPath -InstallPath $destination -DryRun } |
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
