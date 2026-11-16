#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force
  $script:ScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../scripts/Maintenance/Repair-LegacyPaths.ps1'))
  $script:Content = Get-Content -LiteralPath $script:ScriptPath -Raw
  $script:HarnessPath = Join-Path $TestDrive 'Repair-LegacyPaths.ps1'
  $harnessText = $script:Content -replace '(?m)^#Requires -RunAsAdministrator\r?\n', ''
  [IO.File]::WriteAllText($script:HarnessPath, $harnessText, (New-Object Text.UTF8Encoding($true)))
}

Describe 'Repair-LegacyPaths' {
  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Root -Force
    $script:Legacy = Join-Path $script:Root 'PSFoundation-Office'
    $script:New = Join-Path $script:Root 'winkit-Office'
    $script:Media = Join-Path $script:Root 'Media'
    $script:Reports = Join-Path $script:Root 'Reports'
    $null = New-Item -ItemType Directory -Path $script:Media -Force
    $global:LASTEXITCODE = 0
  }

  It 'requires elevation and exposes preview conventions' {
    $script:Content | Should -Match '#Requires -RunAsAdministrator'
    $script:Content | Should -Match "RequiredVersion = '0.3.1'"
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    (Get-Help $script:ScriptPath).Synopsis | Should -Match 'legacy PSFoundation Office paths'
  }

  It 'does nothing on a machine without legacy paths' {
    $results = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -PassThru -Confirm:$false 6>$null
    $summary = $results | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 0
    $summary.Status | Should -Be 'Skipped'
    $summary.Moved | Should -Be 0
    Test-Path -LiteralPath $script:Legacy | Should -BeFalse
    Test-Path -LiteralPath $script:New | Should -BeFalse
  }

  It 'moves the legacy journal root and writes a report' {
    $null = New-Item -ItemType Directory -Path $script:Legacy -Force
    [IO.File]::WriteAllText((Join-Path $script:Legacy 'settings.json'), '{"synthetic":true}')

    $results = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -PassThru -Confirm:$false 6>$null
    $summary = $results | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 0
    $summary.Status | Should -Be 'Completed'
    $summary.Moved | Should -Be 1
    Test-Path -LiteralPath $script:Legacy | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $script:New 'settings.json') | Should -BeTrue
    Test-Path -LiteralPath $summary.ManifestPath | Should -BeTrue

    $manifest = Get-Content -LiteralPath $summary.ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest.Status | Should -Be 'Completed'
    $manifest.Moved | Should -Be 1
  }

  It 'refuses to merge when both roots exist' {
    $null = New-Item -ItemType Directory -Path $script:Legacy -Force
    $null = New-Item -ItemType Directory -Path $script:New -Force
    [IO.File]::WriteAllText((Join-Path $script:Legacy 'legacy.json'), 'legacy')
    [IO.File]::WriteAllText((Join-Path $script:New 'new.json'), 'new')

    $results = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -PassThru -Confirm:$false 6>$null
    $summary = $results | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 1
    $summary.Status | Should -Be 'Failed'
    $summary.Refused | Should -Be 1
    Test-Path -LiteralPath (Join-Path $script:Legacy 'legacy.json') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:New 'new.json') | Should -BeTrue
  }

  It 'refuses a pending recovery journal unless Force is supplied' {
    $null = New-Item -ItemType Directory -Path $script:Legacy -Force
    $journal = Join-Path $script:Legacy '0123456789abcdef0123456789abcdef.json'
    [IO.File]::WriteAllText($journal, '{"synthetic":true}')

    $blocked = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -PassThru -Confirm:$false 6>$null
    $blockedSummary = $blocked | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 1
    $blockedSummary.Refused | Should -Be 1
    Test-Path -LiteralPath $journal | Should -BeTrue
    Test-Path -LiteralPath $script:New | Should -BeFalse

    $forced = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -Force -PassThru -Confirm:$false 6>$null
    $forcedSummary = $forced | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 0
    $forcedSummary.Status | Should -Be 'Completed'
    Test-Path -LiteralPath $script:Legacy | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $script:New '0123456789abcdef0123456789abcdef.json') | Should -BeTrue
  }

  It 'renames a legacy media manifest without changing its content' {
    $folder = Join-Path $script:Media 'Office2019'
    $null = New-Item -ItemType Directory -Path $folder -Force
    $legacyManifest = Join-Path $folder 'psfoundation-office-media.json'
    [IO.File]::WriteAllText($legacyManifest, '{"synthetic":true}')

    $results = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -PassThru -Confirm:$false 6>$null
    $summary = $results | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 0
    $summary.Status | Should -Be 'Completed'
    $summary.Moved | Should -Be 1
    Test-Path -LiteralPath $legacyManifest | Should -BeFalse
    Get-Content -LiteralPath (Join-Path $folder 'winkit-office-media.json') -Raw | Should -Be '{"synthetic":true}'
  }

  It 'refuses a media folder that has both manifest names' {
    $folder = Join-Path $script:Media 'Office2019'
    $null = New-Item -ItemType Directory -Path $folder -Force
    [IO.File]::WriteAllText((Join-Path $folder 'psfoundation-office-media.json'), 'old')
    [IO.File]::WriteAllText((Join-Path $folder 'winkit-office-media.json'), 'new')

    $results = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -PassThru -Confirm:$false 6>$null
    $summary = $results | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 1
    $summary.Status | Should -Be 'Failed'
    $summary.Refused | Should -Be 1
    Test-Path -LiteralPath (Join-Path $folder 'psfoundation-office-media.json') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $folder 'winkit-office-media.json') | Should -BeTrue
  }

  It 'previews without moving paths or writing a report' {
    $null = New-Item -ItemType Directory -Path $script:Legacy -Force
    [IO.File]::WriteAllText((Join-Path $script:Legacy 'settings.json'), '{"synthetic":true}')

    $results = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -DryRun -PassThru -Confirm:$false 6>$null
    $summary = $results | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 0
    $summary.Status | Should -Be 'Preview'
    $summary.Moved | Should -Be 0
    Test-Path -LiteralPath $script:Legacy | Should -BeTrue
    Test-Path -LiteralPath $script:New | Should -BeFalse
    Test-Path -LiteralPath $script:Reports | Should -BeFalse
  }

  It 'is idempotent on a second run' {
    $null = New-Item -ItemType Directory -Path $script:Legacy -Force
    [IO.File]::WriteAllText((Join-Path $script:Legacy 'settings.json'), '{"synthetic":true}')

    $null = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -Confirm:$false 6>$null
    $results = & $script:HarnessPath -LegacyLogRoot $script:Legacy -LogRoot $script:New -MediaPath $script:Media -ReportDirectory $script:Reports -PassThru -Confirm:$false 6>$null
    $summary = $results | Where-Object Action -EQ 'RepairLegacyPaths'

    $LASTEXITCODE | Should -Be 0
    $summary.Status | Should -Be 'Skipped'
    $summary.Moved | Should -Be 0
  }
}
