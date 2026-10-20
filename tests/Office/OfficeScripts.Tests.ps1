#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Tier-0 logic tests: enforce the winkit script conventions on the Office
# scripts without needing Outlook. Pure Outlook parsing logic (transport
# Message-ID extraction) lives in AdNoctem.Substrate.PowerShell and is tested there.

BeforeAll {
  $script:OfficeScripts = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '..\..\scripts\Office'))
  $script:Scripts = @(Get-ChildItem -LiteralPath $script:OfficeScripts -Filter '*.ps1' -File | Sort-Object Name)
}

Describe 'Office script file conventions' {
  It 'contains the expected script set' {
    @($script:Scripts | ForEach-Object { $_.Name }) | Should -Be @(
      'Backup-Outlook.ps1',
      'Checkpoint-Outlook.ps1',
      'Install-Office.ps1',
      'New-OutlookArchive.ps1',
      'New-TestOutlookMessage.ps1',
      'Optimize-Outlook.ps1',
      'Remove-Office.ps1',
      'Repair-OutlookDataFile.ps1',
      'Show-OutlookState.ps1',
      'Split-OutlookArchive.ps1',
      'Switch-OfficeVersion.ps1'
    )
  }

  It 'declares the PowerShell version and AdNoctem.Substrate.PowerShell module dependency' {
    foreach ($_script in $script:Scripts) {
      $_header = Get-Content -LiteralPath $_script.FullName -TotalCount 3
      $_header -match '^#Requires -Version' | Should -Not -BeNullOrEmpty -Because "$($_script.Name) must declare the PowerShell version"
      $_header -match '^#Requires -Modules @\{ ModuleName = .AdNoctem\.Substrate\.PowerShell.' | Should -Not -BeNullOrEmpty -Because "$($_script.Name) must pin AdNoctem.Substrate.PowerShell"
    }
  }

  It 'imports AdNoctem.Substrate.PowerShell after the parameter block' {
    foreach ($_script in $script:Scripts) {
      (Get-Content -LiteralPath $_script.FullName -Raw) -match 'Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.1.1 -Force' | Should -BeTrue -Because "$($_script.Name) must import AdNoctem.Substrate.PowerShell"
    }
  }

  It 'requires the released PST helper API in <ScriptName>' -ForEach @(
    @{ ScriptName = 'New-OutlookArchive' }
    @{ ScriptName = 'Split-OutlookArchive' }
    @{ ScriptName = 'Optimize-Outlook' }
  ) {
    $path = Join-Path $script:OfficeScripts "$ScriptName.ps1"
    $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
    $ast.ScriptRequirements.RequiredModules[0].RequiredVersion | Should -Be ([version]'0.1.1')
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
