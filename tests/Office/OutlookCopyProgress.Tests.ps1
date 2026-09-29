#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

Describe 'Large-file progress in <ScriptName>' -ForEach @(
  @{ ScriptName = 'Backup-Outlook' }
  @{ ScriptName = 'Checkpoint-Outlook' }
) {
  BeforeAll {
    $path = Join-Path $PSScriptRoot ("../../scripts/Office/$ScriptName.ps1")
    $ast = [Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
    $assignments = @($ast.FindAll({
          param ($node)
          $node -is [Management.Automation.Language.AssignmentStatementAst] -and
          $node.Left.Extent.Text -eq '$_percent'
        }, $true))
    if ($assignments.Count -ne 1) {
      throw "Expected one copy-progress calculation in $ScriptName."
    }

    # Evaluate the production calculation without allocating or copying multi-GB files.
    $script:CalculateProgress = [scriptblock]::Create('param ($_handle) ' + $assignments[0].Right.Extent.Text)
  }

  It 'reports <Expected>% at byte <Position> of <Length>' -ForEach @(
    @{
      Length   = [long]0
      Position = [long]0
      Expected = 0
    }
    @{
      Length   = [long]1
      Position = [long]1
      Expected = 100
    }
    @{
      Length   = [long]2147483647
      Position = [long]2147483647
      Expected = 100
    }
    @{
      Length   = [long]2147483648
      Position = [long]1073741824
      Expected = 50
    }
    @{
      Length   = [long]8520360960
      Position = [long]0
      Expected = 0
    }
    @{
      Length   = [long]8520360960
      Position = [long]4260180480
      Expected = 50
    }
    @{
      Length   = [long]8520360960
      Position = [long]8520360960
      Expected = 100
    }
    @{
      Length   = [long]::MaxValue
      Position = [long]::MaxValue
      Expected = 100
    }
  ) {
    $handle = [PSCustomObject]@{
      Stream = [PSCustomObject]@{
        Length   = $Length
        Position = $Position
      }
    }

    $percent = & $script:CalculateProgress $handle
    $percent | Should -Be $Expected
    $percent | Should -BeOfType ([int])
  }
}
