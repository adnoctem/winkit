#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Creates verified, closed-file backups of Outlook PST data files.
.DESCRIPTION
  Discovers PSTs in the current Outlook profile, releases COM references, and
  waits for Outlook to close before copying. PSTPath instead selects existing
  files directly without connecting to Outlook. Each run uses a new directory
  under Destination and writes a JSON manifest with SHA-256 verification.
  Whole PSTs are copied, including non-mail data and deleted or junk items.
  OST caches and server stores are not portable PST backups and are skipped.
  A matching hash verifies the copy, not the logical health of the source PST.
.PARAMETER Destination
  Directory in which to create a uniquely named backup directory.
.PARAMETER PSTPath
  Literal paths to existing PST files. Cannot be combined with any profile
  options: StoreName, AllStores, QuitOutlook, or IgnoreAdministrator.
.PARAMETER StoreName
  Unique display name of one attached store. Defaults to the delivery store.
.PARAMETER AllStores
  Select all attached PST stores in the current profile.
.PARAMETER QuitOutlook
  Request graceful Outlook shutdown after discovery. Otherwise close Outlook
  manually when prompted by the progress message. Never force-kills Outlook.
.PARAMETER IgnoreAdministrator
  Allow elevated profile discovery when Outlook uses the same user and token.
  Not available with PSTPath; direct file backup needs no Outlook identity.
.PARAMETER WaitSeconds
  Maximum seconds to wait for Outlook to exit after profile discovery.
.PARAMETER DryRun
  Preview selected paths without quitting Outlook or writing backup files.
.PARAMETER PassThru
  Return one summary containing backup counts and the manifest ReportPath.
.EXAMPLE
  PS> .\Backup-Outlook.ps1 -Destination E:\OutlookBackups -AllStores -DryRun
.EXAMPLE
  PS> .\Backup-Outlook.ps1 -Destination E:\OutlookBackups -AllStores -QuitOutlook -PassThru
.EXAMPLE
  PS> .\Backup-Outlook.ps1 -PSTPath D:\Archive\mail.pst -Destination E:\OutlookBackups -PassThru
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: PSTPath only; profile discovery requires classic Outlook.
  SYSTEM-account execution: PSTPath only with explicit file permissions.
  Outlook version: 2007 or later for profile discovery; no Outlook needed with PSTPath.
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Profile')]
param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $Destination,

  [Parameter(Mandatory = $true, ParameterSetName = 'File')]
  [ValidateNotNullOrEmpty()]
  [string[]]
  $PSTPath,

  [Parameter(ParameterSetName = 'Profile')]
  [ValidateNotNullOrEmpty()]
  [string]
  $StoreName,

  [Parameter(Mandatory = $true, ParameterSetName = 'All')]
  [switch]
  $AllStores,

  [Parameter(ParameterSetName = 'Profile')]
  [Parameter(ParameterSetName = 'All')]
  [switch]
  $QuitOutlook,

  [Parameter(ParameterSetName = 'Profile')]
  [Parameter(ParameterSetName = 'All')]
  [switch]
  $IgnoreAdministrator,

  [Parameter(ParameterSetName = 'Profile')]
  [Parameter(ParameterSetName = 'All')]
  [ValidateRange(0, 3600)]
  [int]
  $WaitSeconds = 120,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

if ($PSCmdlet.ParameterSetName -ne 'File') {
  $_outlookUser = Get-UserInfo
  if ($_outlookUser.IsAdministrator) {
    if (-not $IgnoreAdministrator) {
      throw "PowerShell is elevated as '$($_outlookUser.UserName)'. Use a non-elevated window as the Outlook user, or explicitly specify IgnoreAdministrator when Outlook intentionally uses that same elevated identity."
    }

    Write-Warning 'IgnoreAdministrator permits elevated discovery. Outlook must use the same Windows user and elevation.'
  }
}

if ($DryRun) {
  $WhatIfPreference = $true
}

$_startedAt = Get-Date
$_results = New-Object Collections.ArrayList
$_sources = New-Object Collections.ArrayList
$_handles = New-Object Collections.ArrayList
$_context = $null
$_reportPath = $null
$_reportStream = $null
$_runDirectory = $null
$_failed = 0
$_copied = 0
$_declined = $false

try {
  $_destination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination)
  if ($PSCmdlet.ParameterSetName -eq 'File') {
    foreach ($_path in $PSTPath) {
      $_resolved = (Resolve-Path -LiteralPath $_path -ErrorAction Stop).ProviderPath
      if ([IO.Path]::GetExtension($_resolved) -ne '.pst' -or -not [IO.File]::Exists($_resolved)) {
        throw "PSTPath must identify an existing .pst file: $_path"
      }

      $null = $_sources.Add([PSCustomObject]@{
          Path      = $_resolved
          StoreName = $null
        })
    }
  }
  else {
    Write-Progress -Id 60 -Activity 'Outlook data-file backup' -Status 'Discovering profile data files'
    $_context = Connect-Outlook
    if ([int](($_context.App.Version -split '\.')[0]) -lt 12) {
      throw 'Outlook 2007 or later is required for profile discovery.'
    }

    $_default = $_context.Namespace.DefaultStore
    try {
      $_defaultId = $_default.StoreID
    }
    finally {
      Remove-ComObject $_default
    }

    $_matches = 0
    $_stores = $_context.Namespace.Stores
    try {
      for ($_index = 1; $_index -le $_stores.Count; $_index++) {
        $_store = $_stores.Item($_index)
        try {
          $_selected = $AllStores -or ($StoreName -and $_store.DisplayName -eq $StoreName) -or
          (-not $StoreName -and -not $AllStores -and $_store.StoreID -eq $_defaultId)
          if (-not $_selected) {
            continue
          }

          $_matches++
          $_path = [string]$_store.FilePath
          if ([IO.Path]::GetExtension($_path) -ne '.pst') {
            Add-OperationResult -Results $_results -Target $_store.DisplayName -Source 'Outlook' -Action 'Backup' -Status 'Skipped' -Detail 'Not a portable PST data file.'
            continue
          }

          $null = $_sources.Add([PSCustomObject]@{
              Path      = $_path
              StoreName = [string]$_store.DisplayName
            })
        }
        finally {
          Remove-ComObject $_store
        }
      }
    }
    finally {
      Remove-ComObject $_stores
    }

    if (-not $AllStores -and $_matches -ne 1) {
      throw "StoreName '$StoreName' matches $_matches stores. Select one unique store."
    }
  }

  $_sources = @($_sources | Sort-Object -Property Path -Unique)
  if ($_sources.Count -eq 0) {
    throw 'No PST files were selected. OST caches and server stores cannot be backed up as portable PST files.'
  }

  $_runName = 'OutlookBackup-{0}-{1}' -f $_startedAt.ToString('yyyyMMdd-HHmmss-fff'), [guid]::NewGuid().ToString('N')
  $_runDirectory = Join-Path $_destination $_runName
  foreach ($_source in $_sources) {
    Write-Log -Message ("Backup source: {0} -> {1}" -f $_source.Path, $_runDirectory) -Color Cyan
  }

  if ($WhatIfPreference) {
    foreach ($_source in $_sources) {
      Add-OperationResult -Results $_results -Target $_source.Path -Source 'Outlook' -Action 'Backup' -Status 'Skipped' -Detail 'DryRun'
    }
  }
  elseif (-not $PSCmdlet.ShouldProcess($_runDirectory, "Back up and verify $($_sources.Count) PST file(s)")) {
    $_declined = $true
  }
  else {
    if ($_context -and $QuitOutlook) {
      if (-not $PSCmdlet.ShouldProcess('Outlook', 'Request graceful shutdown for closed-file backup')) {
        throw 'Outlook shutdown declined. No backup files were copied.'
      }

      $_context.App.Quit()
    }

    if ($_context) {
      Remove-ComObject $_context.Namespace $_context.App
      $_context = $null
      Invoke-ComGarbageCollection

      Write-Log -Message 'Close Outlook now. Waiting for Outlook to exit before copying PST files...' -Color Yellow
      $_timer = [Diagnostics.Stopwatch]::StartNew()
      while (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) {
        if ($_timer.Elapsed.TotalSeconds -ge $WaitSeconds) {
          throw 'Outlook is still running. Close it and retry; no files were copied.'
        }

        Write-Progress -Id 60 -Activity 'Outlook data-file backup' -Status 'Waiting for Outlook to close'
        Start-Sleep -Milliseconds 500
      }
    }

    # Hold every source open exclusively for the complete copy/verification run.
    # This also prevents Outlook from reopening these files after the process check.
    foreach ($_source in $_sources) {
      $_stream = [IO.File]::Open($_source.Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
      $null = $_handles.Add([PSCustomObject]@{
          Source = $_source
          Stream = $_stream
        })
    }

    $null = [IO.Directory]::CreateDirectory($_runDirectory)
    $_reportPath = Join-Path $_runDirectory 'manifest.json'
    $_reportStream = [IO.File]::Open($_reportPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    $_number = 0
    foreach ($_handle in $_handles) {
      $_number++
      $_fileName = '{0:D3}-{1}' -f $_number, [IO.Path]::GetFileName($_handle.Source.Path)
      $_target = Join-Path $_runDirectory $_fileName
      $_partial = $_target + '.partial'
      $_output = $null
      $_sha = $null
      try {
        $_output = [IO.File]::Open($_partial, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $_buffer = New-Object byte[] (1024 * 1024)
        $_timer = [Diagnostics.Stopwatch]::StartNew()
        while (($_read = $_handle.Stream.Read($_buffer, 0, $_buffer.Length)) -gt 0) {
          $_output.Write($_buffer, 0, $_read)
          if ($_timer.ElapsedMilliseconds -ge 200) {
            $_percent = [int](100.0 * $_handle.Stream.Position / [math]::Max(1, $_handle.Stream.Length))
            Write-Progress -Id 60 -Activity 'Outlook data-file backup' -Status "Copying $_number / $($_handles.Count)" -CurrentOperation $_handle.Source.Path -PercentComplete $_percent
            $_timer.Restart()
          }
        }

        $_output.Flush($true)
        Write-Progress -Id 60 -Activity 'Outlook data-file backup' -Status 'Verifying SHA-256 hashes' -CurrentOperation $_handle.Source.Path
        $_sha = [Security.Cryptography.SHA256]::Create()
        $_handle.Stream.Position = 0
        $_sourceHash = [BitConverter]::ToString($_sha.ComputeHash($_handle.Stream)).Replace('-', '')
        $_output.Position = 0
        $_targetHash = [BitConverter]::ToString($_sha.ComputeHash($_output)).Replace('-', '')
        if ($_sourceHash -ne $_targetHash -or $_output.Length -ne $_handle.Stream.Length) {
          throw 'Backup verification failed; the partial file has been retained for inspection.'
        }

        $_output.Dispose()
        $_output = $null
        [IO.File]::Move($_partial, $_target)
        $_copied++
        $_properties = @{
          OriginalPath = $_handle.Source.Path
          StoreName    = $_handle.Source.StoreName
          Bytes        = $_handle.Stream.Length
          SourceSHA256 = $_sourceHash
          BackupSHA256 = $_targetHash
        }

        Add-OperationResult -Results $_results -Target $_target -Source 'Outlook' -Action 'Backup' -Status 'Completed' -Detail 'Closed-file copy verified with SHA-256.' -Property $_properties
      }
      finally {
        if ($_output) {
          $_output.Dispose()
        }
        if ($_sha) {
          $_sha.Dispose()
        }
      }
    }
  }
}
catch {
  $_failed++
  Add-OperationResult -Results $_results -Target $Destination -Source 'Outlook' -Action 'Backup' -Status 'Failed' -Detail $_.Exception.Message
  Write-Warning $_.Exception.Message
}
finally {
  foreach ($_handle in $_handles) {
    $_handle.Stream.Dispose()
  }

  if ($_context) {
    Remove-ComObject $_context.Namespace $_context.App
    Invoke-ComGarbageCollection
  }

  if ($_reportStream) {
    try {
      $_manifest = [ordered]@{
        SchemaVersion = 1
        Script        = 'Backup-Outlook'
        StartedAt     = $_startedAt.ToString('o')
        FinishedAt    = (Get-Date).ToString('o')
        Copied        = $_copied
        Failed        = $_failed
        Results       = @($_results.ToArray())
      }

      $_json = ConvertTo-Json -InputObject $_manifest -Depth 8 -ErrorAction Stop
      $_bytes = (New-Object Text.UTF8Encoding($true)).GetBytes($_json)
      $_reportStream.Write($_bytes, 0, $_bytes.Length)
      $_reportStream.Flush($true)
    }
    catch {
      $_failed++
      Write-Warning "Could not write backup manifest: $($_.Exception.Message)"
      $_reportPath = $null
    }
    finally {
      $_reportStream.Dispose()
    }
  }

  Write-Progress -Id 60 -Activity 'Outlook data-file backup' -Completed
}

$_status = if ($_failed) {
  'Failed'
}
elseif ($_declined) {
  'Skipped'
}
elseif ($WhatIfPreference) {
  'Preview'
}
else {
  'Completed'
}
$_summaryProperty = @{
  Preview         = [bool]$WhatIfPreference
  Copied          = $_copied
  Failed          = $_failed
  ReportPath      = $_reportPath
  BackupDirectory = $_runDirectory
  Results         = @($_results.ToArray())
}

Write-Log -Message "Outlook backup: $_status | Copied: $_copied | Failed: $_failed | Manifest: $_reportPath" -Color $(if ($_failed) { 'Yellow' } else { 'Green' })
if ($PassThru -or $WhatIfPreference) {
  New-OperationResult -Target $Destination -Source 'Outlook' -Action 'Backup' -Status $_status -Detail 'Whole-file PST backup; see Results and ReportPath.' -Property $_summaryProperty
}

if ($_failed) {
  exit 1
}
