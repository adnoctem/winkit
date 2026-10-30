#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  $script:RepositoryRoot = Split-Path $PSScriptRoot -Parent
  $script:RuntimeName = 'AdNoctem.Substrate.PowerShell'
  $script:RuntimeVersion = (Import-PowerShellDataFile (Join-Path $script:RepositoryRoot 'requirements.psd1'))[$script:RuntimeName]
}

Describe 'Substrate dependency contract' {
  It 'pins every runtime consumer to the manifest version and supported PowerShell host' {
    $consumers = 0
    foreach ($file in Get-ChildItem (Join-Path $script:RepositoryRoot 'scripts') -Filter '*.ps1' -Recurse -File) {
      $parseIssues = $null
      $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$parseIssues)
      $parseIssues | Should -BeNullOrEmpty -Because $file.FullName
      $runtime = @($ast.ScriptRequirements.RequiredModules | Where-Object Name -EQ $script:RuntimeName)
      if ($runtime.Count -eq 0) {
        continue
      }

      $consumers++
      $runtime.Count | Should -Be 1
      $runtime[0].RequiredVersion | Should -Be ([version]$script:RuntimeVersion) -Because $file.FullName
      $ast.ScriptRequirements.RequiredPSVersion | Should -BeGreaterOrEqual ([version]'5.1') -Because $file.FullName
      $content = [IO.File]::ReadAllText($file.FullName)
      $content | Should -Match ([regex]::Escape("Import-Module $script:RuntimeName -RequiredVersion $script:RuntimeVersion -Force"))
      $content | Should -Not -Match 'Import-Module\s+PSFoundation\b'
    }
    $consumers | Should -BeGreaterThan 0
  }

  It 'loads the pinned module identity rather than the frozen PSFoundation module' {
    $module = Import-Module $script:RuntimeName -RequiredVersion $script:RuntimeVersion -Force -PassThru
    $module.Version | Should -Be ([version]$script:RuntimeVersion)
    $module.Guid | Should -Be ([guid]'81c6e99d-6a1d-4fc0-b688-2b36c361df05')
    (Get-Command New-OperationResult).ModuleName | Should -Be $script:RuntimeName
  }
}

Describe 'developer dependency initialization' {
  BeforeEach {
    Mock Import-PowerShellDataFile { @{ 'AdNoctem.Substrate.PowerShell' = '0.2.1' } }
    Mock Install-PackageProvider { }
    Mock Install-Module { }
    # Stands for any installed version ahead of the pin. It must stay out of
    # reach of a real release, or the substitution it guards against vanishes.
    Mock Get-Module {
      [pscustomobject]@{ Version = [version]'9.9.9' }
    } -ParameterFilter { $ListAvailable -and $Name -eq 'AdNoctem.Substrate.PowerShell' }
  }

  It 'does not substitute a newer installed version for the pinned release' {
    & (Join-Path $script:RepositoryRoot 'tools/initialize.ps1')
    Should -Invoke Install-Module -Times 1 -Exactly -ParameterFilter {
      $Name -eq 'AdNoctem.Substrate.PowerShell' -and $RequiredVersion -eq '0.2.1' -and $Repository -eq 'PSGallery'
    }
  }

  It 'propagates installation failures instead of announcing success' {
    Mock Install-Module { throw 'Synthetic installation failure' }
    { & (Join-Path $script:RepositoryRoot 'tools/initialize.ps1') } | Should -Throw '*Synthetic installation failure*'
  }
}
