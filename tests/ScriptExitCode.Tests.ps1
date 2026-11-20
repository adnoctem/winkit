#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Group Policy and scheduled tasks start winkit scripts with -File and rely on their exit code. Substrate before
# 0.4.1 made every exit N of a script that requires it report 0 under -File, so a real script is run that way.
Describe 'Script exit codes under -File' {
  It 'reports a failure as exit code 1 in <Engine>' -ForEach @(
    @{ Engine = 'pwsh' }
    @{ Engine = 'powershell.exe' }
  ) {
    $script = Join-Path $PSScriptRoot '../scripts/Administration/Show-UserDataInventory.ps1'

    # A UNC -Include is refused before anything is read, so the script fails fast and changes nothing.
    $output = & $Engine -NoProfile -ExecutionPolicy Bypass -File $script -Include '\\server\share' -DryRun 2>&1 | ForEach-Object { [string]$_ }
    $exitCode = $LASTEXITCODE

    ($output -join "`n") | Should -Match '-Include needs a local drive path'
    $exitCode | Should -Be 1
  }
}
