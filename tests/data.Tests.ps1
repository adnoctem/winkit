#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }


BeforeAll {
  . $PSScriptRoot/../lib/data.ps1
}

Describe 'Convert-Quote' {
  BeforeAll {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'This is a test file and the variable is used in multiple It blocks.')]

    $testFile = [System.IO.Path]::GetTempFileName()
  }

  AfterAll {
    if (Test-Path $testFile) { Remove-Item $testFile -Force }
  }

  It 'converts single quotes to double quotes by default' {
    Set-Content -Path $testFile -Value "He said 'hello' to the 'world'."
    Convert-Quote -Path $testFile
    $result = (Get-Content -Path $testFile -Raw).TrimEnd()
    $result | Should -Be 'He said "hello" to the "world".'
  }

  It 'converts double quotes to single quotes via -To Single' {
    Set-Content -Path $testFile -Value 'He said "hello" to the "world".'
    Convert-Quote -Path $testFile -To Single
    $result = (Get-Content -Path $testFile -Raw).TrimEnd()
    $result | Should -Be "He said 'hello' to the 'world'."
  }

  It 'leaves file unchanged when no target quotes are present' {
    Set-Content -Path $testFile -Value 'No quotes here.'
    Convert-Quote -Path $testFile
    $result = (Get-Content -Path $testFile -Raw).TrimEnd()
    $result | Should -Be 'No quotes here.'
  }

  It 'handles mixed quote types (converting single to double)' {
    Set-Content -Path $testFile -Value "It's a `"test`" of both 'types'."
    Convert-Quote -Path $testFile
    $result = (Get-Content -Path $testFile -Raw).TrimEnd()
    $result | Should -Be 'It"s a "test" of both "types".'
  }
}

Describe 'Merge-ObjectArrays' {
  # Merge-ObjectArrays is designed for [hashtable] base entries and
  # [PSCustomObject] overrides: .Keys works on hashtables for key
  # iteration, PSObject.Properties resolves custom properties on
  # PSCustomObject for containment checks.

  It 'overrides matching keys by Name' {
    $base = @(@{
        Name = 'A'
        Value = 1
      }, @{
        Name = 'B'
        Value = 2
      })
    $over = @([PSCustomObject]@{
        Name = 'A'
        Value = 99
      })
    Merge-ObjectArrays -Base $base -Overrides $over
    $base[0].Value | Should -Be 99
    $base[1].Value | Should -Be 2
  }

  It 'ignores keys not present on the base object' {
    $base = @(@{
        Name = 'A'
        Value = 1
      })
    $over = @([PSCustomObject]@{
        Name = 'A'
        Value = 99
        Extra = 'dropped'
      })
    Merge-ObjectArrays -Base $base -Overrides $over
    $base[0].Value | Should -Be 99
    $base[0].Keys -contains 'Extra' | Should -BeFalse
  }

  It 'matches by Path and Name when Path is present on override' {
    $base = @(
      @{
        Path = 'HKCU:'
        Name = 'Setting'
        Value = 1
      }
      @{
        Path = 'HKLM:'
        Name = 'Setting'
        Value = 2
      }
    )
    $over = @([PSCustomObject]@{
        Path = 'HKLM:'
        Name = 'Setting'
        Value = 99
      })
    Merge-ObjectArrays -Base $base -Overrides $over
    $base[0].Value | Should -Be 1
    $base[1].Value | Should -Be 99
  }

  It 'falls back to Name-only match when Path is absent' {
    $base = @(@{
        Name = 'A'
        Value = 1
      })
    $over = @([PSCustomObject]@{
        Name = 'A'
        Value = 99
      })
    Merge-ObjectArrays -Base $base -Overrides $over
    $base[0].Value | Should -Be 99
  }

  It 'skips overrides without a Name property' {
    $base = @(@{
        Name = 'A'
        Value = 1
      })
    $over = @([PSCustomObject]@{
        NoName = 'ignored'
        Value = 99
      })
    Merge-ObjectArrays -Base $base -Overrides $over
    $base[0].Value | Should -Be 1
  }

  It 'skips overrides where Name is null' {
    $base = @(@{
        Name = 'A'
        Value = 1
      })
    $over = @([PSCustomObject]@{
        Name = $null
        Value = 99
      })
    Merge-ObjectArrays -Base $base -Overrides $over
    $base[0].Value | Should -Be 1
  }
}
