#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.8.0' }

<#
.SYNOPSIS
  Captures Outlook data files and Office settings for the current Windows user.
.DESCRIPTION
  Discovers attached data files through classic Outlook, or uses explicit
  StorePaths without opening Outlook. Copies closed files, exports existing
  per-user Office/profile registry keys, and records Office inventory and
  activation status in a JSON manifest. Each checkpoint has a unique directory.
  Includes signatures, templates, dictionaries, Outlook application data,
  AutoComplete cache, and Ribbon/Quick Access Toolbar customization files.
  SHA-256 verifies file copies by default. Missing optional folders are reported;
  inaccessible, missing selected, or locked files fail the checkpoint.
  This is a file/settings checkpoint, not a transactional system snapshot or
  automatic restore. OST copies are caches, not portable mailbox backups.
.PARAMETER Destination
  Local or UNC directory under which a unique checkpoint directory is created.
  Choose storage accessible only to the intended users: it contains personal data.
.PARAMETER StorePaths
  Literal PST/OST paths. Disables Outlook discovery; cannot be combined with
  QuitOutlook or WaitSeconds. Other per-user settings are still captured.
.PARAMETER ExcludeOst
  Omit OST caches throughout the checkpoint, including settings directories.
.PARAMETER QuitOutlook
  Request graceful Outlook shutdown after discovery. Never force-kills a process.
.PARAMETER WaitSeconds
  Maximum seconds to wait for Outlook to close after discovery. Default 120.
.PARAMETER IgnoreAdministrator
  Allow an elevated checkpoint of the current account. Does not select another
  user's registry or files. SYSTEM and service identities are always rejected.
.PARAMETER SkipHash
  Verify copy lengths only. Manifest explicitly records that hashes were skipped.
.PARAMETER DryRun
  Discover and display the capture plan without quitting Outlook or writing files.
  Profile discovery may start Outlook; StorePaths avoids connecting to Outlook.
.PARAMETER PassThru
  Return one operation summary with Results, CheckpointDirectory, and ReportPath.
.EXAMPLE
  PS> .\Checkpoint-Outlook.ps1 -Destination E:\Checkpoints -ExcludeOst -DryRun
.EXAMPLE
  PS> .\Checkpoint-Outlook.ps1 -Destination E:\Checkpoints -QuitOutlook -PassThru
.EXAMPLE
  PS> .\Checkpoint-Outlook.ps1 -Destination \\server\backups -StorePaths D:\Mail\archive.pst -PassThru
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: explicit StorePaths only, for an existing user profile.
  SYSTEM-account execution: not supported; captures the invoking user's HKCU and files.
  Outlook version: 2007 or later for discovery; no Outlook required with StorePaths.
  Close other Office applications before capture. Restore is a separate manual operation.
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
  $StorePaths,

  [switch]
  $ExcludeOst,

  [Parameter(ParameterSetName = 'Profile')]
  [switch]
  $QuitOutlook,

  [Parameter(ParameterSetName = 'Profile')]
  [ValidateRange(0, 3600)]
  [int]
  $WaitSeconds = 120,

  [switch]
  $IgnoreAdministrator,

  [switch]
  $SkipHash,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

$_user = Get-UserInfo
if ($_user.SID -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) {
  throw 'Run the checkpoint as the affected Outlook user, not SYSTEM or a service identity.'
}
if ($_user.IsAdministrator) {
  if (-not $IgnoreAdministrator) {
    throw "PowerShell is elevated as '$($_user.UserName)'. Use a non-elevated window as the Outlook user, or explicitly specify IgnoreAdministrator for that same account."
  }
  Write-Warning 'IgnoreAdministrator captures this elevated account only. Outlook discovery must use the same user and elevation.'
}
if ($DryRun) {
  $WhatIfPreference = $true
}

$_startedAt = Get-Date
$_results = New-Object Collections.ArrayList
$_plan = New-Object Collections.ArrayList
$_registryPlan = New-Object Collections.ArrayList
$_handles = New-Object Collections.ArrayList
$_activation = New-Object Collections.ArrayList
$_warnings = New-Object Collections.ArrayList
$_seen = @{}
$_context = $null
$_inventory = $null
$_runDirectory = $null
$_reportPath = $null
$_manifestStream = $null
$_copied = 0
$_failed = 0
$_status = 'Failed'
$_sessionId = (Get-Process -Id $PID -ErrorAction Stop).SessionId

try {
  $_provider = $null
  $_drive = $null
  $_destination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination, [ref]$_provider, [ref]$_drive)
  if ($_provider.Name -ne 'FileSystem') {
    throw 'Destination must be a filesystem directory.'
  }
  if (Test-Path -LiteralPath $_destination -ErrorAction Stop) {
    if (-not (Test-Path -LiteralPath $_destination -PathType Container -ErrorAction Stop)) {
      throw 'Destination identifies a file rather than a directory.'
    }
    $_destination = Resolve-LongPath -LiteralPath $_destination
  }

  $_runName = 'OutlookCheckpoint-{0}-{1}' -f $_startedAt.ToString('yyyyMMdd-HHmmss-fff'), [guid]::NewGuid().ToString('N')
  $_runDirectory = Join-Path $_destination $_runName
  Write-Log -Message "Planning Outlook checkpoint for $($_user.UserName)..." -Color Cyan
  Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Status 'Discovering data files and user settings'

  $_stores = New-Object Collections.ArrayList
  if ($PSCmdlet.ParameterSetName -eq 'File') {
    foreach ($_path in $StorePaths) {
      $null = $_stores.Add([PSCustomObject]@{
          Path      = $_path
          StoreName = $null
        })
    }
  }
  else {
    $_context = Connect-Outlook
    if ([int](($_context.App.Version -split '\.')[0]) -lt 12) {
      throw 'Outlook 2007 or later is required for profile discovery.'
    }
    $_collection = $_context.Namespace.Stores
    try {
      for ($_index = 1; $_index -le $_collection.Count; $_index++) {
        $_store = $_collection.Item($_index)
        try {
          $_path = [string]$_store.FilePath
          if ([IO.Path]::GetExtension($_path) -notin @('.pst', '.ost')) {
            Add-OperationResult -Results $_results -Target $_store.DisplayName -Source Outlook -Action Checkpoint -Status Skipped -Detail 'Store has no local PST/OST file; server contents are not captured.'
            continue
          }
          $null = $_stores.Add([PSCustomObject]@{
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
      Remove-ComObject $_collection
    }
  }

  $_number = 0
  foreach ($_store in $_stores) {
    $_extension = [IO.Path]::GetExtension($_store.Path)
    if ($_extension -notin @('.pst', '.ost')) {
      throw "StorePaths must contain PST/OST files: $($_store.Path)"
    }
    if ($ExcludeOst -and $_extension -eq '.ost') {
      Add-OperationResult -Results $_results -Target $_store.Path -Source Outlook -Action Checkpoint -Status Skipped -Detail 'OST excluded by request.'
      continue
    }

    $_resolved = Resolve-LongPath -LiteralPath $_store.Path
    $_item = Get-Item -LiteralPath $_resolved -Force -ErrorAction Stop
    if ($_item.PSIsContainer -or ($_item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
      throw "Select a regular data file, not a directory or link: $_resolved"
    }
    if ($_seen.ContainsKey($_resolved)) {
      continue
    }
    $_seen[$_resolved] = $true
    $_number++
    $null = $_plan.Add([PSCustomObject]@{
        OriginalPath = $_resolved
        RelativePath = ('Stores\{0:D3}-{1}' -f $_number, $_item.Name)
        Category     = 'Store'
        StoreName    = $_store.StoreName
        Bytes        = $_item.Length
      })
  }

  # Outlook shutdown can create cache/settings files; enumerate them afterwards.
  if (-not $WhatIfPreference) {
    if ($_context -and $QuitOutlook) {
      if (-not $PSCmdlet.ShouldProcess('Outlook', 'Request graceful shutdown before checkpoint')) {
        throw 'Outlook shutdown declined. No checkpoint files were created.'
      }
      $_context.App.Quit()
    }
    if ($_context) {
      Remove-ComObject $_context.Namespace $_context.App
      $_context = $null
      Invoke-ComGarbageCollection

      Write-Log -Message 'Close Outlook. Waiting for it to exit before capture...' -Color Yellow
      $_timer = [Diagnostics.Stopwatch]::StartNew()
      while (@(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $_sessionId }).Count) {
        if ($_timer.Elapsed.TotalSeconds -ge $WaitSeconds) {
          throw 'Outlook is still running in this session. No checkpoint files were created.'
        }
        Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Status 'Waiting for Outlook to close'
        Start-Sleep -Milliseconds 500
      }
    }

    $_officeApps = @(Get-Process -Name OUTLOOK, WINWORD, EXCEL, POWERPNT, ONENOTE, MSACCESS, MSPUB, VISIO, WINPROJ -ErrorAction SilentlyContinue |
        Where-Object { $_.SessionId -eq $_sessionId })
    if ($_officeApps.Count) {
      throw 'Close Office applications in this session before capturing their files and settings.'
    }

  }

  # Plan settings before creating output. Never follow directory links.
  if (-not $env:APPDATA -or -not $env:LOCALAPPDATA) {
    throw 'The current account has no usable APPDATA/LOCALAPPDATA profile.'
  }
  $_folders = @(
    @{
      Name   = 'Signatures'
      Path   = (Join-Path $env:APPDATA 'Microsoft\Signatures')
      Filter = '*'
    }
    @{
      Name   = 'Templates'
      Path   = (Join-Path $env:APPDATA 'Microsoft\Templates')
      Filter = '*'
    }
    @{
      Name   = 'UProof'
      Path   = (Join-Path $env:APPDATA 'Microsoft\UProof')
      Filter = '*'
    }
    @{
      Name   = 'OutlookAppData'
      Path   = (Join-Path $env:APPDATA 'Microsoft\Outlook')
      Filter = '*'
    }
    @{
      Name   = 'RoamCache'
      Path   = (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\RoamCache')
      Filter = '*'
    }
    @{
      Name   = 'OfficeUI'
      Path   = (Join-Path $env:LOCALAPPDATA 'Microsoft\Office')
      Filter = '*.officeUI'
    }
  )
  foreach ($_folder in $_folders) {
    if (-not (Test-Path -LiteralPath $_folder.Path -ErrorAction Stop)) {
      Add-OperationResult -Results $_results -Target $_folder.Path -Source Outlook -Action Checkpoint -Status Skipped -Detail 'Optional settings directory is absent.'
      continue
    }
    $_root = Resolve-LongPath -LiteralPath $_folder.Path
    if (-not (Test-Path -LiteralPath $_root -PathType Container -ErrorAction Stop)) {
      throw "Expected a settings directory: $_root"
    }
    if ($_destination.TrimEnd('\') -eq $_root.TrimEnd('\') -or $_destination.StartsWith($_root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
      throw "Destination must be outside captured settings directories: $_root"
    }

    $_pending = New-Object Collections.Queue
    $_pending.Enqueue($_root)
    while ($_pending.Count) {
      $_directory = $_pending.Dequeue()
      $_directoryInfo = Get-Item -LiteralPath $_directory -Force -ErrorAction Stop
      if ($_directoryInfo.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        throw "Settings directory is a link; select and back up its real location separately: $_directory"
      }
      Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Status "Reading $($_folder.Name)" -CurrentOperation $_directory
      foreach ($_item in @(Get-ChildItem -LiteralPath $_directory -Force -ErrorAction Stop)) {
        if ($_item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
          throw "Settings contain a link; review its target before capture: $($_item.FullName)"
        }
        if ($_item.PSIsContainer) {
          $_pending.Enqueue($_item.FullName)
          continue
        }
        if ($_item.Name -notlike $_folder.Filter -or $_seen.ContainsKey($_item.FullName)) {
          continue
        }
        if ($ExcludeOst -and $_item.Extension -eq '.ost') {
          Add-OperationResult -Results $_results -Target $_item.FullName -Source Outlook -Action Checkpoint -Status Skipped -Detail 'OST excluded by request.'
          continue
        }
        if ($_item.Extension -in @('.pst', '.ost')) {
          Add-OperationResult -Results $_results -Target $_item.FullName -Source Outlook -Action Checkpoint -Status Skipped -Detail 'Data file was not selected through StorePaths or the Outlook profile.'
          continue
        }
        $_seen[$_item.FullName] = $true
        $_relative = $_item.FullName.Substring($_root.TrimEnd('\').Length).TrimStart('\')
        $null = $_plan.Add([PSCustomObject]@{
            OriginalPath = $_item.FullName
            RelativePath = Join-Path ('Folders\' + $_folder.Name) $_relative
            Category     = $_folder.Name
            StoreName    = $null
            Bytes        = $_item.Length
          })
      }
    }
  }

  foreach ($_version in @('12.0', '14.0', '15.0', '16.0')) {
    $_key = 'HKCU\Software\Microsoft\Office\' + $_version
    if (Test-RegistryPath -Path $_key -ErrorAction Stop) {
      $null = $_registryPlan.Add([PSCustomObject]@{
          Key          = $_key
          RelativePath = "Registry\Office$_version.reg"
        })
    }
  }
  $_legacyProfiles = 'HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
  if (Test-RegistryPath -Path $_legacyProfiles -ErrorAction Stop) {
    $null = $_registryPlan.Add([PSCustomObject]@{
        Key          = $_legacyProfiles
        RelativePath = 'Registry\LegacyOutlookProfiles.reg'
      })
  }
  if (-not $_registryPlan.Count) {
    Add-OperationResult -Results $_results -Target HKCU -Source Outlook -Action Checkpoint -Status Skipped -Detail 'No supported per-user Office/profile registry keys are present.'
  }
  if (-not $_plan.Count -and -not $_registryPlan.Count) {
    throw 'No data files or Office user settings were found for this checkpoint.'
  }

  Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Status 'Reading Office inventory and activation'
  $_inventory = Get-OfficeInventory -ErrorAction Stop
  foreach ($_product in $_inventory.Products) {
    try {
      $null = $_activation.Add((Get-OfficeActivationStatus -TargetProductId $_product.ProductId -ErrorAction Stop))
    }
    catch {
      $null = $_activation.Add([PSCustomObject]@{
          TargetProductId = $_product.ProductId
          Status          = 'Unknown'
        })
      $null = $_warnings.Add("Activation could not be read for $($_product.ProductId): $($_.Exception.Message)")
    }
  }

  Write-Log -Message ("Checkpoint plan: {0} file(s), {1} registry export(s) -> {2}" -f $_plan.Count, $_registryPlan.Count, $_runDirectory) -Color Cyan
  if ($WhatIfPreference) {
    foreach ($_entry in $_plan) {
      Add-OperationResult -Results $_results -Target (Join-Path $_runDirectory $_entry.RelativePath) -Source Outlook -Action Checkpoint -Status Skipped -Detail 'Preview: would copy this file.' -Property @{
        OriginalPath = $_entry.OriginalPath
        RelativePath = $_entry.RelativePath
        Category     = $_entry.Category
        Bytes        = $_entry.Bytes
      }
    }
    foreach ($_entry in $_registryPlan) {
      Add-OperationResult -Results $_results -Target (Join-Path $_runDirectory $_entry.RelativePath) -Source Outlook -Action Checkpoint -Status Skipped -Detail 'Preview: would export this registry key.' -Property @{ RegistryKey = $_entry.Key }
    }
    $_status = 'Preview'
  }
  elseif (-not $PSCmdlet.ShouldProcess($_runDirectory, "Capture $($_plan.Count) files and $($_registryPlan.Count) per-user registry keys")) {
    $_status = 'Skipped'
  }
  else {
    # Lock every source before writing anything; keep locks through verification.
    foreach ($_entry in $_plan) {
      $_stream = [IO.File]::Open($_entry.OriginalPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
      $null = $_handles.Add([PSCustomObject]@{
          Entry  = $_entry
          Stream = $_stream
        })
    }

    $null = [IO.Directory]::CreateDirectory($_runDirectory)
    $_reportPath = Join-Path $_runDirectory 'manifest.json'
    $_manifestStream = [IO.File]::Open($_reportPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    foreach ($_handle in $_handles) {
      $_entry = $_handle.Entry
      $_modifiedAt = [IO.File]::GetLastWriteTimeUtc($_entry.OriginalPath)
      $_target = Join-Path $_runDirectory $_entry.RelativePath
      $null = [IO.Directory]::CreateDirectory((Split-Path -Parent $_target))
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
            $_percent = [int](100.0 * $_handle.Stream.Position / [math]::Max([long]1, $_handle.Stream.Length))
            Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Status "Copying $($_copied + 1) / $($_plan.Count)" -CurrentOperation $_entry.OriginalPath -PercentComplete $_percent
            $_timer.Restart()
          }
        }
        $_output.Flush($true)
        if ($_output.Length -ne $_handle.Stream.Length) {
          throw 'Checkpoint copy length mismatch. The partial file is retained for inspection.'
        }

        $_sourceHash = $null
        $_backupHash = $null
        if (-not $SkipHash) {
          Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Status 'Verifying SHA-256 hashes' -CurrentOperation $_entry.OriginalPath
          $_sha = [Security.Cryptography.SHA256]::Create()
          $_handle.Stream.Position = 0
          $_sourceHash = [BitConverter]::ToString($_sha.ComputeHash($_handle.Stream)).Replace('-', '')
          $_output.Position = 0
          $_backupHash = [BitConverter]::ToString($_sha.ComputeHash($_output)).Replace('-', '')
          if ($_sourceHash -ne $_backupHash) {
            throw 'Checkpoint hash mismatch. The partial file is retained for inspection.'
          }
        }
        $_output.Dispose()
        $_output = $null
        [IO.File]::Move($_partial, $_target)
        [IO.File]::SetLastWriteTimeUtc($_target, $_modifiedAt)
        $_copied++
        Add-OperationResult -Results $_results -Target $_target -Source Outlook -Action Checkpoint -Status Completed -Detail 'Closed-file copy completed.' -Property @{
          OriginalPath     = $_entry.OriginalPath
          RelativePath     = $_entry.RelativePath
          Category         = $_entry.Category
          StoreName        = $_entry.StoreName
          Bytes            = $_handle.Stream.Length
          LastWriteTimeUtc = $_modifiedAt.ToString('o')
          SourceSHA256     = $_sourceHash
          BackupSHA256     = $_backupHash
          Verification     = $(if ($SkipHash) { 'LengthOnly' } else { 'SHA256' })
        }
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

    foreach ($_entry in $_registryPlan) {
      $_target = Join-Path $_runDirectory $_entry.RelativePath
      $null = [IO.Directory]::CreateDirectory((Split-Path -Parent $_target))
      Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Status 'Exporting user registry settings' -CurrentOperation $_entry.Key
      $_exportResult = Invoke-SafeProcess -FilePath (Join-Path $env:SystemRoot 'System32\reg.exe') `
        -ArgumentList @('export', $_entry.Key, $_target, '/y') -AsResult -ErrorAction Stop
      if ($_exportResult.ExitCode -ne 0 -or $_exportResult.TimedOut -or $_exportResult.Cancelled) {
        throw "Registry export failed for $($_entry.Key); exit code $($_exportResult.ExitCode)."
      }
      $_export = Get-Item -LiteralPath $_target -ErrorAction Stop
      if ($_export.Length -eq 0) {
        throw "Registry export is empty: $($_entry.Key)"
      }
      $_hash = $null
      if (-not $SkipHash) {
        $_hash = (Get-FileHash -LiteralPath $_target -Algorithm SHA256 -ErrorAction Stop).Hash
      }
      Add-OperationResult -Results $_results -Target $_target -Source Outlook -Action Checkpoint -Status Completed -Detail 'Registry export captured; hash identifies the exported file, not a transactional registry snapshot.' -Property @{
        RegistryKey  = $_entry.Key
        RelativePath = $_entry.RelativePath
        Category     = 'Registry'
        Bytes        = $_export.Length
        BackupSHA256 = $_hash
      }
    }
    $_status = 'Completed'
  }
}
catch {
  $_failed++
  $_status = 'Failed'
  Add-OperationResult -Results $_results -Target $Destination -Source Outlook -Action Checkpoint -Status Failed -Detail $_.Exception.Message
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
  if ($_manifestStream) {
    try {
      $_manifest = [ordered]@{
        SchemaVersion = 1
        Script        = 'Checkpoint-Outlook'
        Status        = $_status
        StartedAt     = $_startedAt.ToString('o')
        FinishedAt    = (Get-Date).ToString('o')
        Computer      = $env:COMPUTERNAME
        User          = $_user.UserName
        UserSID       = $_user.SID
        Selection     = $PSCmdlet.ParameterSetName
        Hashed        = (-not $SkipHash)
        ExcludeOst    = [bool]$ExcludeOst
        Copied        = $_copied
        Failed        = $_failed
        Office        = $_inventory
        Activation    = @($_activation.ToArray())
        Warnings      = @($_warnings.ToArray())
        Results       = @($_results.ToArray())
      }
      $_json = ConvertTo-Json -InputObject $_manifest -Depth 16 -ErrorAction Stop
      $_bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($_json)
      $_manifestStream.Write($_bytes, 0, $_bytes.Length)
      $_manifestStream.Flush($true)
    }
    catch {
      $_failed++
      $_status = 'Failed'
      $_reportPath = $null
      Add-OperationResult -Results $_results -Target $_runDirectory -Source Outlook -Action Checkpoint -Status Failed -Detail "Manifest write failed: $($_.Exception.Message)"
      Write-Warning "Could not write checkpoint manifest: $($_.Exception.Message)"
    }
    finally {
      $_manifestStream.Dispose()
    }
  }
  Write-Progress -Id 61 -Activity 'Outlook checkpoint' -Completed
}

foreach ($_warning in $_warnings) {
  Write-Warning $_warning
}
Write-Log -Message "Outlook checkpoint: $_status | Copied: $_copied | Failed: $_failed | Manifest: $_reportPath" -Color $(if ($_failed) { 'Yellow' } else { 'Green' })
if ($PassThru -or $WhatIfPreference) {
  New-OperationResult -Target $Destination -Source Outlook -Action Checkpoint -Status $_status -Detail 'User data/settings checkpoint; inspect Results and ReportPath before manual restore.' -Property @{
    Preview             = [bool]$WhatIfPreference
    Copied              = $_copied
    Failed              = $_failed
    CheckpointDirectory = $_runDirectory
    ReportPath          = $_reportPath
    Hashed              = (-not $SkipHash)
    Warnings            = @($_warnings.ToArray())
    Results             = @($_results.ToArray())
  }
}
if ($_failed) {
  exit 1
}
