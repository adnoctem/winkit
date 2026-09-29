#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Wrapper contracts only. Native Office behavior is tested in PSFoundation.
BeforeAll {
  Import-Module PSFoundation -Force

  $script:EntryPoints = @{}
  $script:ScriptAsts = @{}
  $officePath = Join-Path $PSScriptRoot '../../scripts/Office'

  foreach ($name in @('Install-Office', 'Remove-Office', 'Switch-OfficeVersion')) {
    $path = Join-Path $officePath ($name + '.ps1')
    $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
    $script:ScriptAsts[$name] = $ast

    # Exercise the real body without elevation/version directives or exiting Pester.
    # All mutating module entry points are mocked below.
    $statements = @($ast.EndBlock.Statements | Where-Object {
        $_.Extent.Text -notmatch '^Import-Module'
      } | ForEach-Object { $_.Extent.Text })

    $body = ($ast.ParamBlock.Attributes.Extent.Text -join [Environment]::NewLine) + [Environment]::NewLine +
    $ast.ParamBlock.Extent.Text + [Environment]::NewLine + ($statements -join [Environment]::NewLine)
    $body = $body.Replace('exit $_exitCode', 'throw "WrapperExit:$_exitCode"')
    $script:EntryPoints[$name] = [scriptblock]::Create($body)
  }

  function Invoke-OfficeWrapperTest {
    param (
      [string]
      $Entry,

      [hashtable]
      $Arguments
    )

    $collected = New-Object Collections.ArrayList
    $exitCode = 0

    try {
      & $script:EntryPoints[$Entry] @Arguments | ForEach-Object {
        [void]$collected.Add($_)
      }
    }
    catch {
      if ($_.Exception.Message -match '^WrapperExit:(\d+)$') {
        $exitCode = [int]$Matches[1]
      }
      else {
        throw
      }
    }

    [PSCustomObject]@{
      ExitCode = $exitCode
      Results  = @($collected)
      Result   = @($collected)[-1]
    }
  }
}

Describe 'Office wrapper contracts' {
  It 'does not expose the retired migration opt-in' {
    $script:ScriptAsts['Switch-OfficeVersion'].ParamBlock.Parameters.Name.VariablePath.UserPath | Should -Not -Contain 'PilotMigration'
  }

  BeforeEach {
    $script:Inventory = [PSCustomObject]@{
      Products = @()
      Msi      = @()
      Unknowns = @()
    }

    $script:Plan = [PSCustomObject]@{
      Action   = 'Install'
      Eligible = $true
      Blockers = @()
      Warnings = @()
    }

    $script:Outcome = New-OperationResult -Target test -Source Office -Action Install -Status Completed -Property @{
      ReasonCode       = 'AlreadyCompliant'
      WrapperExitCode  = 0
      Changed          = $false
      ChangeKnown      = $true
      AlreadyCompliant = $true
      RebootRequired   = $false
      RecoveryPath     = 'C:\Journal\original.json'
      NativeResults    = @([PSCustomObject]@{ ExitCode = 0 })
    }

    Mock Write-Log {}
    Mock Get-OfficeInventory { $script:Inventory }
    Mock Get-OfficeActivationStatus { [PSCustomObject]@{ Status = 'Licensed' } }
    Mock Get-OfficeDeploymentPlan { $script:Plan }
    Mock Install-Office { $script:Outcome }
    Mock Uninstall-Office { $script:Outcome }
    Mock Switch-OfficeDeployment { $script:Outcome }
    Mock Get-OfficeDeploymentRecovery {
      [PSCustomObject]@{ RunId = $RunId; LogRoot = $LogRoot; Record = @{ Action = 'Install' } }
    }
    Mock Resume-OfficeInstallation { $script:Outcome }
    Mock Resume-OfficeMigration { $script:Outcome }
    Mock Save-OfficeDeploymentMedia {
      [PSCustomObject]@{ Valid = $true; Path = $SourcePath; Manifest = @{ Version = '16.0.17932.20162' } }
    }
    Mock Test-Path { $false }

    # Any accidental direct native operation fails the wrapper test.
    Mock Start-Process { throw 'Wrapper must not launch native processes.' }
    Mock Stop-Process { throw 'Wrapper must not terminate applications.' }
    Mock Remove-Item { throw 'Wrapper must not remove files.' }
  }

  It 'requires PSFoundation 1.8.0 and a mandatory mode in <Entry>' -ForEach @(
    @{ Entry = 'Install-Office' }, @{ Entry = 'Remove-Office' }, @{ Entry = 'Switch-OfficeVersion' }
  ) {
    $ast = $script:ScriptAsts[$Entry]
    $ast.ScriptRequirements.RequiredModules[0].Version | Should -Be ([version]'1.8.0')
    $mode = $ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Mode' }
    $mode.Extent.Text | Should -Match 'Mandatory = \$true'
    $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true).Count | Should -Be 0
  }

  It 'Check only inventories without a target in <Entry>' -ForEach @(
    @{ Entry = 'Install-Office' }, @{ Entry = 'Remove-Office' }, @{ Entry = 'Switch-OfficeVersion' }
  ) {
    $run = Invoke-OfficeWrapperTest $Entry @{ Mode = 'Check'; PassThru = $true }
    $run.ExitCode | Should -Be 0
    $run.Result.Inventory | Should -Be $script:Inventory
    $run.Result.Changed | Should -BeFalse
    Should -Invoke Get-OfficeDeploymentPlan -Times 0
    Should -Invoke Install-Office -Times 0
    Should -Invoke Uninstall-Office -Times 0
    Should -Invoke Switch-OfficeDeployment -Times 0
  }

  It 'preserves blocked Check plans and their language warning' {
    $script:Plan.Eligible = $false
    $script:Plan.Blockers = @('UnsupportedNativeVerification')
    $script:Plan.Warnings = @('Language change: de-de -> en-us')
    Mock Write-Warning {}

    $run = Invoke-OfficeWrapperTest 'Switch-OfficeVersion' @{
      Mode = 'Check'; TargetProductId = 'Standard2024Volume'; PassThru = $true
    }

    $run.ExitCode | Should -Be 1
    $run.Result.Status | Should -Be Blocked
    $run.Result.Plan | Should -Be $script:Plan
    $run.Result.ReasonCode | Should -Be UnsupportedNativeVerification
    Should -Invoke Write-Warning -Times 1 -ParameterFilter { $Message -like '*de-de -> en-us*' }
    Should -Invoke Switch-OfficeDeployment -Times 0
  }

  It 'lets the module default to en-us in <Entry>' -ForEach @(
    @{ Entry = 'Install-Office' }, @{ Entry = 'Switch-OfficeVersion' }
  ) {
    $null = Invoke-OfficeWrapperTest $Entry @{ Mode = 'Check'; TargetProductId = 'Standard2024Volume' }
    Should -Invoke Get-OfficeDeploymentPlan -Times 1 -ParameterFilter {
      $Configuration.Language.Count -eq 1 -and $Configuration.Language[0] -eq 'en-us' -and $Configuration.LocaleSource -eq 'Default'
    }
  }

  It 'preserves requested language order and the Publisher convenience switch' {
    $null = Invoke-OfficeWrapperTest 'Install-Office' @{
      Mode = 'Check'; TargetProductId = 'Standard2024Volume'
      Language = @('de-de', 'en-us'); ExcludeApp = @('Access'); ExcludePublisher = $true
    }

    Should -Invoke Get-OfficeDeploymentPlan -Times 1 -ParameterFilter {
      ($Configuration.Language -join ',') -eq 'de-de,en-us' -and
      $Configuration.PrimaryLanguage -eq 'de-de' -and
      $Configuration.ExcludeApp -contains 'Access' -and $Configuration.ExcludeApp -contains 'Publisher'
    }
  }

  It 'forwards explicit automatic sourcing without adding a default Language argument' {
    Mock New-OfficeDeploymentConfiguration {
      [PSCustomObject]@{ TargetProductId = $TargetProductId; Language = @('de-de') }
    }

    $null = Invoke-OfficeWrapperTest 'Install-Office' @{
      Mode = 'Check'; TargetProductId = 'Standard2024Volume'
      AutoSourceLocales = $true; LocaleSource = 'OperatingSystem'
    }

    Should -Invoke New-OfficeDeploymentConfiguration -Times 1 -ParameterFilter {
      $AutoSourceLocales -and $LocaleSource -eq 'OperatingSystem' -and -not $Language
    }
  }

  It 'reports invalid locale combinations before planning' -ForEach @(
    @{ Locale = @{ Language = @('en-us'); AutoSourceLocales = $true } }
    @{ Locale = @{ LocaleSource = 'OperatingSystem' } }
  ) {
    $arguments = @{ Mode = 'Check'; TargetProductId = 'Standard2024Volume'; PassThru = $true }
    foreach ($key in $Locale.Keys) {
      $arguments[$key] = $Locale[$key]
    }
    $run = Invoke-OfficeWrapperTest 'Install-Office' $arguments
    $run.ExitCode | Should -Be 1
    $run.Result.ReasonCode | Should -Be InvalidConfiguration
    Should -Invoke Get-OfficeDeploymentPlan -Times 0
  }

  It 'dispatches <Entry> with explicit confirmation and preview flags' -ForEach @(
    @{ Entry = 'Install-Office'; Mode = 'Install'; Command = 'Install-Office' }
    @{ Entry = 'Switch-OfficeVersion'; Mode = 'Migrate'; Command = 'Switch-OfficeDeployment' }
    @{ Entry = 'Remove-Office'; Mode = 'Remove'; Command = 'Uninstall-Office' }
  ) {
    $arguments = @{
      Mode = $Mode; OdtPath = 'C:\ODT\setup.exe'; DryRun = $true; Confirm = $false
      ForceCloseApps = $true
    }

    if ($Mode -eq 'Remove') {
      $arguments.RemoveProductId = @('O365ProPlusRetail')
    }
    else {
      $arguments.TargetProductId = 'Standard2024Volume'
      $arguments.SourcePath = 'C:\Media\Office'
    }

    $run = Invoke-OfficeWrapperTest $Entry $arguments
    $run.Result | Should -Be $script:Outcome
    Should -Invoke $Command -Times 1 -Exactly -ParameterFilter {
      $DryRun -and $WhatIf -and -not $Confirm -and $ForceCloseApps -and $Plan -eq $script:Plan
    }
    Should -Invoke Start-Process -Times 0
    Should -Invoke Stop-Process -Times 0
    Should -Invoke Remove-Item -Times 0
  }

  It 'passes WhatIf without requiring DryRun and returns the result' {
    $run = Invoke-OfficeWrapperTest 'Install-Office' @{
      Mode = 'Install'; TargetProductId = 'Standard2024Volume'
      OdtPath = 'C:\ODT\setup.exe'; WhatIf = $true
    }

    $run.Results.Count | Should -Be 1
    Should -Invoke Install-Office -Times 1 -ParameterFilter { $WhatIf -and -not $DryRun }
  }

  It 'resolves relative paths during previews without Set-Variable side effects' {
    $script:ExpectedSource = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\Media\Office')
    $script:ExpectedTool = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath('.\ODT\setup.exe')

    $null = Invoke-OfficeWrapperTest 'Install-Office' @{
      Mode = 'Install'; TargetProductId = 'Standard2024Volume'
      SourcePath = '.\Media\Office'; OdtPath = '.\ODT\setup.exe'; WhatIf = $true
    }

    Should -Invoke Get-OfficeDeploymentPlan -Times 1 -ParameterFilter { $SourcePath -eq $script:ExpectedSource }
    Should -Invoke Install-Office -Times 1 -ParameterFilter { $OdtPath -eq $script:ExpectedTool -and $WhatIf }
  }

  It 'forwards migration removal authority only to the migration plan' {
    $null = Invoke-OfficeWrapperTest 'Switch-OfficeVersion' @{
      Mode = 'Migrate'; TargetProductId = 'Standard2024Volume'; OdtPath = 'C:\ODT\setup.exe'
      RemoveProductId = @('HomeBusiness2019Retail'); RemoveMsi = $true; Confirm = $false
    }

    Should -Invoke Get-OfficeDeploymentPlan -Times 1 -ParameterFilter {
      $Action -eq 'Migrate' -and $RemoveMsi -and ($RemoveProductId -join ',') -eq 'HomeBusiness2019Retail'
    }
    Should -Invoke Install-Office -Times 0
    Should -Invoke Uninstall-Office -Times 0
  }

  It 'has no removal bypass on the installer and no MSI switch on the remover' {
    $script:ScriptAsts['Install-Office'].ParamBlock.Parameters.Name.VariablePath.UserPath | Should -Not -Contain 'RemoveProductId'
    $script:ScriptAsts['Install-Office'].ParamBlock.Parameters.Name.VariablePath.UserPath | Should -Not -Contain 'RemoveMsi'
    $script:ScriptAsts['Remove-Office'].ParamBlock.Parameters.Name.VariablePath.UserPath | Should -Not -Contain 'RemoveMsi'
  }

  It 'preserves module exit code <Code> and unknown changes' -ForEach @(
    @{ Code = 0 }, @{ Code = 1 }, @{ Code = 3010 }
  ) {
    $script:Outcome.WrapperExitCode = $Code
    $script:Outcome.Changed = $null
    $script:Outcome.ChangeKnown = $false
    $script:Outcome.ReasonCode = 'ModuleReason'
    $run = Invoke-OfficeWrapperTest 'Install-Office' @{
      Mode = 'Install'; TargetProductId = 'Standard2024Volume'; OdtPath = 'C:\ODT\setup.exe'
      Confirm = $false; PassThru = $true
    }

    $run.ExitCode | Should -Be $Code
    $run.Result | Should -Be $script:Outcome
    $run.Result.Changed | Should -BeNullOrEmpty
    $run.Result.ChangeKnown | Should -BeFalse
    $run.Result.NativeResults.Count | Should -Be 1
  }

  It 'does not emit results by default' {
    $run = Invoke-OfficeWrapperTest 'Install-Office' @{ Mode = 'Check' }
    $run.Results.Count | Should -Be 0
  }

  It 'fails missing or mode-inappropriate input without invoking mutation' -ForEach @(
    @{ Entry = 'Install-Office'; Arguments = @{ Mode = 'Install' } }
    @{ Entry = 'Install-Office'; Arguments = @{ Mode = 'Prepare'; TargetProductId = 'Standard2024Volume'; OdtPath = 'C:\ODT\setup.exe' } }
    @{ Entry = 'Install-Office'; Arguments = @{ Mode = 'Recover'; OdtPath = 'C:\ODT\setup.exe' } }
    @{ Entry = 'Install-Office'; Arguments = @{ Mode = 'Check'; Language = @('de-de') } }
    @{ Entry = 'Install-Office'; Arguments = @{ Mode = 'Prepare'; ForceCloseApps = $true } }
    @{ Entry = 'Remove-Office'; Arguments = @{ Mode = 'Remove'; OdtPath = 'C:\ODT\setup.exe' } }
    @{ Entry = 'Remove-Office'; Arguments = @{ Mode = 'Check'; ForceCloseApps = $true } }
    @{ Entry = 'Switch-OfficeVersion'; Arguments = @{ Mode = 'Recover'; TargetProductId = 'Standard2024Volume' } }
    @{ Entry = 'Switch-OfficeVersion'; Arguments = @{ Mode = 'Recover'; RemoveMsi = $true } }
    @{ Entry = 'Switch-OfficeVersion'; Arguments = @{ Mode = 'Check'; OdtPath = 'C:\Tools\ODT\setup.exe' } }
  ) {
    $Arguments.PassThru = $true
    $run = Invoke-OfficeWrapperTest $Entry $Arguments
    $run.ExitCode | Should -Be 1
    $run.Result.Phase | Should -Be Validate
    $run.Result.Changed | Should -BeFalse
    Should -Invoke Install-Office -Times 0
    Should -Invoke Switch-OfficeDeployment -Times 0
    Should -Invoke Uninstall-Office -Times 0
    Should -Invoke Save-OfficeDeploymentMedia -Times 0
    Should -Invoke Resume-OfficeInstallation -Times 0
    Should -Invoke Resume-OfficeMigration -Times 0
  }

  It 'routes recovery through the matching command in <Entry>' -ForEach @(
    @{ Entry = 'Install-Office'; Command = 'Resume-OfficeInstallation' }
    @{ Entry = 'Switch-OfficeVersion'; Command = 'Resume-OfficeMigration' }
  ) {
    $runId = '0123456789abcdef0123456789abcdef'
    $key = New-Object Security.SecureString
    $run = Invoke-OfficeWrapperTest $Entry @{
      Mode = 'Recover'; RunId = $runId; OdtPath = 'C:\ODT\setup.exe'
      LogRoot = 'C:\CustomJournal'; ProductKey = $key; DryRun = $true; Confirm = $false
    }

    $run.Result | Should -Be $script:Outcome
    Should -Invoke Get-OfficeDeploymentRecovery -Times 1 -ParameterFilter {
      $RunId -eq '0123456789abcdef0123456789abcdef' -and $LogRoot -eq 'C:\CustomJournal'
    }
    Should -Invoke $Command -Times 1 -ParameterFilter {
      $Recovery.RunId -eq '0123456789abcdef0123456789abcdef' -and
      $ProductKey -is [Security.SecureString] -and $DryRun -and $WhatIf -and -not $Confirm
    }
    Should -Invoke Get-OfficeDeploymentPlan -Times 0
  }

  It 'adapts preparation outcomes without calling a deployment executor' -ForEach @(
    @{ Existing = $false; Preview = $false; Reason = 'MediaPrepared'; Changed = $true }
    @{ Existing = $true; Preview = $false; Reason = 'AlreadyPrepared'; Changed = $false }
    @{ Existing = $false; Preview = $true; Reason = 'NotExecuted'; Changed = $false }
  ) {
    $script:ExistingMedia = $Existing
    $script:PreviewMedia = $Preview
    Mock Test-Path { $script:ExistingMedia }
    Mock Save-OfficeDeploymentMedia {
      if ($script:PreviewMedia) {
        return [PSCustomObject]@{ Status = 'Preview'; Path = $SourcePath }
      }
      [PSCustomObject]@{ Valid = $true; Path = $SourcePath; Manifest = @{ Version = '16.0.17932.20162' } }
    }

    $run = Invoke-OfficeWrapperTest 'Install-Office' @{
      Mode = 'Prepare'; TargetProductId = 'Standard2024Volume'; SourcePath = 'C:\Media\Office'
      OdtPath = 'C:\ODT\setup.exe'; Confirm = $false; DryRun = $Preview; PassThru = $true
    }

    $run.Result.ReasonCode | Should -Be $Reason
    $run.Result.Changed | Should -Be $Changed
    $run.Result.Media | Should -Not -BeNullOrEmpty
    Should -Invoke Save-OfficeDeploymentMedia -Times 1 -ParameterFilter { -not $Confirm }
    Should -Invoke Install-Office -Times 0
  }

  It 'returns structured module exceptions without implying unchanged execution' {
    Mock Install-Office {
      $exception = New-Object InvalidOperationException 'Synthetic deployment failure'
      $exception.Data['OfficeReason'] = 'SyntheticFailure'
      throw $exception
    }

    $run = Invoke-OfficeWrapperTest 'Install-Office' @{
      Mode = 'Install'; TargetProductId = 'Standard2024Volume'
      OdtPath = 'C:\ODT\setup.exe'; PassThru = $true; Confirm = $false
    }

    $run.ExitCode | Should -Be 1
    $run.Result.ReasonCode | Should -Be SyntheticFailure
    $run.Result.ChangeKnown | Should -BeFalse
    $run.Result.Changed | Should -BeNullOrEmpty
  }
}
