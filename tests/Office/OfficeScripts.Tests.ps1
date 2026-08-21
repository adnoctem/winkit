#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Tier-0 logic tests: enforce the winkit script conventions on the Office
# scripts without needing Outlook. Pure Outlook parsing logic (transport
# Message-ID extraction) lives in PSFoundation and is tested there.

BeforeAll {
  $script:OfficeScripts = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '..\..\scripts\Office'))
  $script:Scripts = @(Get-ChildItem -LiteralPath $script:OfficeScripts -Filter '*.ps1' -File | Sort-Object Name)
}

Describe 'Office script file conventions' {
  It 'contains the expected script set' {
    @($script:Scripts | ForEach-Object { $_.Name }) | Should -Be @(
      'New-OutlookArchive.ps1',
      'New-TestOutlookMessage.ps1',
      'Optimize-Outlook.ps1',
      'Repair-OutlookDataFile.ps1'
    )
  }

  It 'declares the PowerShell version and PSFoundation module dependency' {
    foreach ($_script in $script:Scripts) {
      $_header = Get-Content -LiteralPath $_script.FullName -TotalCount 3
      $_header -match '^#Requires -Version' | Should -Not -BeNullOrEmpty -Because "$($_script.Name) must declare the PowerShell version"
      $_header -match '^#Requires -Modules @\{ ModuleName = .PSFoundation.' | Should -Not -BeNullOrEmpty -Because "$($_script.Name) must pin PSFoundation"
    }
  }

  It 'imports PSFoundation after the parameter block' {
    foreach ($_script in $script:Scripts) {
      (Get-Content -LiteralPath $_script.FullName -Raw) -match 'Import-Module PSFoundation -Force' | Should -BeTrue -Because "$($_script.Name) must import PSFoundation"
    }
  }

  It 'documents .SYNOPSIS and a complete .NOTES block' {
    foreach ($_script in $script:Scripts) {
      $_content = Get-Content -LiteralPath $_script.FullName -Raw
      $_content -match '\.SYNOPSIS' | Should -BeTrue -Because "$($_script.Name) needs a synopsis"
      $_content -match 'Author: MVProwess' | Should -BeTrue -Because "$($_script.Name) .NOTES needs an Author"
      $_content -match 'License: MIT' | Should -BeTrue -Because "$($_script.Name) .NOTES needs a License"
      $_content -match 'Server Core:' | Should -BeTrue -Because "$($_script.Name) .NOTES needs a Server Core note"
      $_content -match 'SYSTEM-account execution:' | Should -BeTrue -Because "$($_script.Name) .NOTES needs a SYSTEM-account note"
    }
  }
}
