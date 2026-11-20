#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# -DryRun sets $WhatIfPreference. A module that first loads afterwards runs its own setup under WhatIf: it prints
# "What if" lines, leaves its aliases undefined, and under PowerShell 7 can break the proxy for a Windows PowerShell
# module. Every script therefore loads the in-box modules it calls before switching WhatIf on.
BeforeDiscovery {
  $root = Split-Path $PSScriptRoot -Parent
  $script:Scripts = @(
    Get-ChildItem (Join-Path $root 'scripts') -Recurse -Filter *.ps1 |
      Where-Object { [IO.File]::ReadAllText($_.FullName) -match 'WhatIfPreference = \$true' } |
      ForEach-Object { @{ Name = $_.FullName.Substring($root.Length + 1); Path = $_.FullName } }
  )
}

BeforeAll {
  # Modules whose import prints WhatIf lines, measured on Windows PowerShell 5.1 and PowerShell 7.
  $script:Affected = @('CimCmdlets', 'Storage', 'NetTCPIP', 'Defender', 'ConfigDefender', 'Dism', 'Appx', 'SmbShare', 'GroupPolicy')
  $script:CommandModule = @{}
  foreach ($module in $script:Affected) {
    foreach ($command in @(Get-Command -Module $module -ErrorAction SilentlyContinue)) {
      $script:CommandModule[$command.Name] = $module
    }
  }
}

Describe 'Dry-run module preload' {
  It 'loads the affected in-box modules it calls before WhatIf in <Name>' -ForEach $script:Scripts {
    $text = [IO.File]::ReadAllText($Path)
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $used = @(
      $ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] }, $true) |
        ForEach-Object { $_.GetCommandName() } |
        Where-Object { $_ -and $script:CommandModule.ContainsKey($_) } |
        ForEach-Object { $script:CommandModule[$_] } |
        Sort-Object -Unique
    )

    # Defender commands come from ConfigDefender; either name loads them.
    $beforeWhatIf = $text.Substring(0, $text.IndexOf('WhatIfPreference = $true'))
    foreach ($module in $used) {
      $names = @($module)
      if ($module -in @('Defender', 'ConfigDefender')) {
        $names = @('Defender', 'ConfigDefender')
      }

      $loaded = @($names | Where-Object { $beforeWhatIf -match "Import-Module -Name [^\r\n]*\b$([regex]::Escape($_))\b" }).Count -gt 0
      $loaded | Should -BeTrue -Because "$Name calls commands from $module and must load it before setting `$WhatIfPreference"
    }
  }
}
