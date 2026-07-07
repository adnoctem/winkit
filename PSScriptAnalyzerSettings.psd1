@{
  Severity = @('Error', 'Warning')
  IncludeDefaultRules = $true
  # ExcludeRules = @(
  #   # Project convention: Merge-ObjectArrays describes a two-array merge helper.
  #   'PSUseSingularNouns'
  # )

  Rules = @{
    PSPlaceOpenBrace = @{
      Enable = $true
      OnSameLine = $true
      NewLineAfter = $true
      IgnoreOneLineBlock = $true
    }

    PSPlaceCloseBrace = @{
      Enable = $true
      NewLineAfter = $true
      IgnoreOneLineBlock = $true
      NoEmptyLineBefore = $false
    }

    PSUseConsistentIndentation = @{
      Enable = $true
      Kind = 'space'
      IndentationSize = 2
      PipelineIndentation = 'IncreaseIndentationForFirstPipeline'
    }

    PSUseConsistentWhitespace = @{
      Enable = $true
      CheckInnerBrace = $true
      CheckOpenBrace = $true
      CheckOpenParen = $true
      CheckOperator = $true
      CheckPipe = $true
      CheckPipeForRedundantWhitespace = $false
      CheckSeparator = $true
      CheckParameter = $false
      IgnoreAssignmentOperatorInsideHashTable = $false
    }

    PSUseCorrectCasing = @{
      Enable = $true
    }

    PSUseSingularNouns = @{
      Enable = $true
      NounAllowList = @(
        # lib/data.ps1 + scripts/Update-AutoDNSZones.ps1
        'Arrays',
        # lib/system.ps1
        'Paths',
        # scripts/Update-AutoDNSZones.ps1
        'Records',
        # scripts/Find-OffHoursActivity.ps1
        'Profiles',
        # scripts/Remove-Bloatware.ps1
        'Policies'
      )
    }

    PSUseBOMForUnicodeEncodedFiles = @{
      Enable = $true
    }
  }
}
