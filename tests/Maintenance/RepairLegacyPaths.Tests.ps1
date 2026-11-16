#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.4.0 -Force
  $script:ScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../scripts/Maintenance/Repair-LegacyPaths.ps1'))
  $script:Content = Get-Content -LiteralPath $script:ScriptPath -Raw
  $script:HarnessPath = Join-Path $TestDrive 'Repair-LegacyPaths.ps1'
  $harnessText = $script:Content -replace '(?m)^#Requires -RunAsAdministrator\r?\n', ''
  [IO.File]::WriteAllText($script:HarnessPath, $harnessText, (New-Object Text.UTF8Encoding($true)))

  function Invoke-Repair {
    param (
      [hashtable]
      $Extra = @{}
    )

    $parameters = @{}
    foreach ($key in $script:Common.Keys) {
      $parameters[$key] = $script:Common[$key]
    }
    foreach ($key in $Extra.Keys) {
      $parameters[$key] = $Extra[$key]
    }
    $parameters['PassThru'] = $true
    $parameters['Confirm'] = $false

    $global:LASTEXITCODE = 0
    $results = @(& $script:HarnessPath @parameters 6>$null)
    return [PSCustomObject]@{
      ExitCode = $global:LASTEXITCODE
      Results  = $results
      Summary  = $results | Where-Object Action -EQ 'RepairLegacyPaths'
    }
  }
}

Describe 'Repair-LegacyPaths' {
  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $script:ProgramData = Join-Path $script:Root 'ProgramData'
    $script:Users = Join-Path $script:Root 'Users'
    $script:UserProfile = Join-Path $script:Users 'Current'
    $script:Roaming = Join-Path $script:UserProfile 'AppData\Roaming'
    $script:Local = Join-Path $script:UserProfile 'AppData\Local'
    $script:Temp = Join-Path $script:UserProfile 'AppData\Local\Temp'
    $script:Media = Join-Path $script:Root 'Media'
    $script:Reports = Join-Path $script:Root 'Reports'
    $null = New-Item -ItemType Directory -Path $script:ProgramData, $script:UserProfile, $script:Roaming, $script:Local, $script:Temp, $script:Media -Force
    $script:OfficeLegacy = Join-Path $script:ProgramData 'PSFoundation-Office'
    $script:OfficeCurrent = Join-Path $script:ProgramData 'winkit-Office'
    $script:MachineLegacy = Join-Path $script:ProgramData 'PSFoundation'
    $script:MachineCurrent = Join-Path $script:ProgramData 'winkit'
    $script:Common = @{
      ProgramDataRoot = $script:ProgramData
      UserProfile     = $script:UserProfile
      ProfilesRoot    = $script:Users
      RoamingRoot     = $script:Roaming
      LocalRoot       = $script:Local
      TempRoot        = $script:Temp
      MediaPath       = $script:Media
      ReportDirectory = $script:Reports
    }
  }

  It 'requires elevation and exposes preview conventions' {
    $script:Content | Should -Match '#Requires -RunAsAdministrator'
    $script:Content | Should -Match "RequiredVersion = '0.4.0'"
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    (Get-Help $script:ScriptPath).Synopsis | Should -Match 'legacy PSFoundation'
  }

  It 'does nothing on a machine without legacy paths' {
    $run = Invoke-Repair

    $run.ExitCode | Should -Be 0
    $run.Summary.Status | Should -Be 'Skipped'
    $run.Summary.Moved | Should -Be 0
    Test-Path -LiteralPath $script:OfficeLegacy | Should -BeFalse
    Test-Path -LiteralPath $script:MachineLegacy | Should -BeFalse
  }

  It 'moves the legacy Office journal root and writes a report' {
    $null = New-Item -ItemType Directory -Path $script:OfficeLegacy -Force
    [IO.File]::WriteAllText((Join-Path $script:OfficeLegacy 'settings.json'), '{"synthetic":true}')

    $run = Invoke-Repair

    $run.ExitCode | Should -Be 0
    $run.Summary.Status | Should -Be 'Completed'
    $run.Summary.Moved | Should -Be 1
    Test-Path -LiteralPath $script:OfficeLegacy | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $script:OfficeCurrent 'settings.json') | Should -BeTrue
    Test-Path -LiteralPath $run.Summary.ManifestPath | Should -BeTrue

    $manifest = Get-Content -LiteralPath $run.Summary.ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest.Status | Should -Be 'Completed'
    $manifest.Moved | Should -Be 1
  }

  It 'moves machine and current-user PSFoundation paths' {
    $paths = @(
      (Join-Path $script:MachineLegacy 'data.txt'),
      (Join-Path $script:UserProfile 'PSFoundation\home.txt'),
      (Join-Path $script:Roaming 'PSFoundation\config.txt'),
      (Join-Path $script:Local 'PSFoundation\cache.txt'),
      (Join-Path $script:Temp 'PSFoundation\logs.txt')
    )
    foreach ($path in $paths) {
      $null = New-Item -ItemType Directory -Path (Split-Path -Path $path -Parent) -Force
      [IO.File]::WriteAllText($path, 'synthetic')
    }

    $run = Invoke-Repair

    $run.ExitCode | Should -Be 0
    $run.Summary.Moved | Should -Be 5
    Test-Path -LiteralPath (Join-Path $script:MachineCurrent 'data.txt') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:UserProfile 'winkit\home.txt') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:Roaming 'winkit\config.txt') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:Local 'winkit\cache.txt') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:Temp 'winkit\logs.txt') | Should -BeTrue
  }

  It 'moves legacy paths in another local profile' {
    $legacy = Join-Path $script:Users 'Other\AppData\Roaming\PSFoundation\config.txt'
    $null = New-Item -ItemType Directory -Path (Split-Path -Path $legacy -Parent) -Force
    [IO.File]::WriteAllText($legacy, 'synthetic')

    $run = Invoke-Repair

    $run.ExitCode | Should -Be 0
    $run.Summary.Moved | Should -Be 1
    Test-Path -LiteralPath (Join-Path $script:Users 'Other\AppData\Roaming\winkit\config.txt') | Should -BeTrue
  }

  It 'refuses to merge when both Office roots exist' {
    $null = New-Item -ItemType Directory -Path $script:OfficeLegacy, $script:OfficeCurrent -Force
    [IO.File]::WriteAllText((Join-Path $script:OfficeLegacy 'legacy.json'), 'legacy')
    [IO.File]::WriteAllText((Join-Path $script:OfficeCurrent 'new.json'), 'new')

    $run = Invoke-Repair

    $run.ExitCode | Should -Be 1
    $run.Summary.Status | Should -Be 'Failed'
    $run.Summary.Refused | Should -Be 1
    Test-Path -LiteralPath (Join-Path $script:OfficeLegacy 'legacy.json') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:OfficeCurrent 'new.json') | Should -BeTrue
  }

  It 'refuses a pending recovery journal unless Force is supplied' {
    $null = New-Item -ItemType Directory -Path $script:OfficeLegacy -Force
    $journal = Join-Path $script:OfficeLegacy '0123456789abcdef0123456789abcdef.json'
    [IO.File]::WriteAllText($journal, '{"synthetic":true}')

    $blocked = Invoke-Repair
    $blocked.ExitCode | Should -Be 1
    $blocked.Summary.Refused | Should -Be 1
    Test-Path -LiteralPath $journal | Should -BeTrue
    Test-Path -LiteralPath $script:OfficeCurrent | Should -BeFalse

    $forced = Invoke-Repair -Extra @{ Force = $true }
    $forced.ExitCode | Should -Be 0
    $forced.Summary.Status | Should -Be 'Completed'
    Test-Path -LiteralPath $script:OfficeLegacy | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $script:OfficeCurrent '0123456789abcdef0123456789abcdef.json') | Should -BeTrue
  }

  It 'renames a legacy media manifest without changing its content' {
    $folder = Join-Path $script:Media 'Office2019'
    $null = New-Item -ItemType Directory -Path $folder -Force
    $legacyManifest = Join-Path $folder 'psfoundation-office-media.json'
    [IO.File]::WriteAllText($legacyManifest, '{"synthetic":true}')

    $run = Invoke-Repair

    $run.ExitCode | Should -Be 0
    $run.Summary.Status | Should -Be 'Completed'
    $run.Summary.Moved | Should -Be 1
    Test-Path -LiteralPath $legacyManifest | Should -BeFalse
    Get-Content -LiteralPath (Join-Path $folder 'winkit-office-media.json') -Raw | Should -Be '{"synthetic":true}'
  }

  It 'refuses a media folder that has both manifest names' {
    $folder = Join-Path $script:Media 'Office2019'
    $null = New-Item -ItemType Directory -Path $folder -Force
    [IO.File]::WriteAllText((Join-Path $folder 'psfoundation-office-media.json'), 'old')
    [IO.File]::WriteAllText((Join-Path $folder 'winkit-office-media.json'), 'new')

    $run = Invoke-Repair

    $run.ExitCode | Should -Be 1
    $run.Summary.Status | Should -Be 'Failed'
    $run.Summary.Refused | Should -Be 1
    Test-Path -LiteralPath (Join-Path $folder 'psfoundation-office-media.json') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $folder 'winkit-office-media.json') | Should -BeTrue
  }

  It 'previews without moving paths or writing a report' {
    $null = New-Item -ItemType Directory -Path $script:OfficeLegacy -Force
    [IO.File]::WriteAllText((Join-Path $script:OfficeLegacy 'settings.json'), '{"synthetic":true}')

    $run = Invoke-Repair -Extra @{ DryRun = $true }

    $run.ExitCode | Should -Be 0
    $run.Summary.Status | Should -Be 'Preview'
    $run.Summary.Moved | Should -Be 0
    Test-Path -LiteralPath $script:OfficeLegacy | Should -BeTrue
    Test-Path -LiteralPath $script:OfficeCurrent | Should -BeFalse
    Test-Path -LiteralPath $script:Reports | Should -BeFalse
  }

  It 'is idempotent on a second run' {
    $null = New-Item -ItemType Directory -Path $script:OfficeLegacy -Force
    [IO.File]::WriteAllText((Join-Path $script:OfficeLegacy 'settings.json'), '{"synthetic":true}')

    $null = Invoke-Repair
    $run = Invoke-Repair

    $run.ExitCode | Should -Be 0
    $run.Summary.Status | Should -Be 'Skipped'
    $run.Summary.Moved | Should -Be 0
  }
}
