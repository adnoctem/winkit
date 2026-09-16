@{
  Severity = @('Error', 'Warning')
  IncludeDefaultRules = $true
  # ExcludeRules = @(
  #   # Project convention: Merge-ObjectArrays describes a two-array merge helper.
  #   'PSUseSingularNouns'
  # )

  Rules = @{
    PSUseCompatibleSyntax = @{
      Enable = $true
      TargetVersions = @('5.0', '5.1', '7.0')
    }

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
        # PSFoundation data.ps1 + scripts/Administration/Update-AutoDNSZones.ps1
        'Arrays',
        # PSFoundation system.ps1
        'Paths',
        # scripts/Administration/Update-AutoDNSZones.ps1
        'Records',
        # scripts/Diagnostics/Find-OffHoursActivity.ps1
        'Profiles',
        # scripts/Software/Remove-Bloatware.ps1
        'Policies',
        # scripts/Features/Enable-RemoteDesktopServices.ps1
        'RemoteDesktopServices'
      )
    }

    PSUseBOMForUnicodeEncodedFiles = @{
      Enable = $true
    }
  }
}
