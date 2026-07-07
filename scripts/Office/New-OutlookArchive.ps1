#Requires -Version 5.1

<#
.SYNOPSIS
  Archives Outlook mail into a standalone Unicode PST.
.DESCRIPTION
  Adds a new PST store to the current Outlook profile, mirrors the source
  store's folder hierarchy, and copies or moves items into it. Optional
  received-date bounds limit which mail items are archived. When finished, the
  PST can be detached so the file is closed and portable.
.PARAMETER ArchivePath
  Full path of the .pst file to create or populate.
.PARAMETER StoreName
  Display name of the source Outlook store. If omitted, the default delivery
  store is used.
.PARAMETER StartDate
  Optional inclusive received-date lower bound.
.PARAMETER EndDate
  Optional inclusive received-date upper bound.
.PARAMETER Mode
  Copy leaves source mail intact. Move removes archived items from the source.
.PARAMETER DisplayName
  Display name for the mounted PST store while it is attached.
.PARAMETER DetachWhenDone
  Remove the PST store from the profile at the end.
.PARAMETER DryRun
  Preview changes without copying or moving messages.
.PARAMETER PassThru
  Return structured operation result objects.
.PARAMETER QuitOutlook
  Quit the Outlook application object on exit. Leave off if Outlook was already
  open interactively.
.EXAMPLE
  PS> .\New-OutlookArchive.ps1 -ArchivePath D:\Backups\user-snapshot.pst -StoreName 'user@example.com' -Mode Copy
.EXAMPLE
  PS> .\New-OutlookArchive.ps1 -ArchivePath D:\Archive\user-2025.pst -StoreName 'user@example.com' -StartDate '2025-01-01' -EndDate '2025-12-31' -Mode Move
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
  [Parameter(Mandatory = $true)]
  [string]
  $ArchivePath,

  [string]
  $StoreName,

  [datetime]
  $StartDate,

  [datetime]
  $EndDate,

  [ValidateSet('Copy', 'Move')]
  [string]
  $Mode = 'Copy',

  [string]
  $DisplayName = 'Archive',

  [bool]
  $DetachWhenDone = $true,

  [switch]
  $DryRun,

  [switch]
  $PassThru,

  [switch]
  $QuitOutlook
)

# ---- Module import -----------------------------------------------------------
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$module = Join-Path $root 'lib/winkit.psm1'
Import-Module $module -Force
# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no Outlook messages will be archived`n" -Color Yellow
}

$_results = New-Object System.Collections.ArrayList
$script:OutlookArchiveCopied = 0
$script:OutlookArchiveMoved = 0

function Test-OutlookItemInRange {
  param (
    [object]
    $Item
  )

  if (-not $StartDate -and -not $EndDate) { return $true }

  $_receivedTime = $null
  try {
    $_receivedTime = $Item.ReceivedTime
  }
  catch {
    return $false
  }

  if ($StartDate -and $_receivedTime -lt $StartDate) { return $false }
  if ($EndDate -and $_receivedTime -gt $EndDate) { return $false }

  return $true
}

function Add-OutlookArchiveResult {
  param (
    [System.Collections.IList]
    $Results,

    [string]
    $Action,

    [string]
    $Status,

    [string]
    $Folder,

    [object]
    $Item,

    [string]
    $Detail
  )

  $_subject = ''
  $_received = $null
  try { $_subject = [string]$Item.Subject } catch { $_subject = '' }
  try { $_received = $Item.ReceivedTime } catch { $_received = $null }

  $_property = @{
    Received = $_received
  }

  Add-OperationResult `
    -Results $Results `
    -Target $_subject `
    -Source 'Outlook' `
    -Scope $Folder `
    -Action $Action `
    -Status $Status `
    -Detail $Detail `
    -Property $_property
}

function Copy-OutlookFolderItem {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $SourceFolder,

    [object]
    $DestinationFolder,

    [string]
    $ArchiveMode,

    [System.Collections.IList]
    $Results
  )

  $_items = $SourceFolder.Items
  try {
    if ($ArchiveMode -eq 'Move') {
      for ($_index = $_items.Count; $_index -ge 1; $_index--) {
        $_item = $_items.Item($_index)
        try {
          if (-not (Test-OutlookItemInRange -Item $_item)) { continue }

          $_destinationPath = if ($DestinationFolder) { $DestinationFolder.FolderPath } else { $_archivePath }
          $_status = if ($WhatIfPreference) { 'Skipped' } else { 'Moved' }
          $_detail = if ($WhatIfPreference) { 'DryRun' } else { "Moved to $_destinationPath" }
          Add-OutlookArchiveResult -Results $Results -Action 'Move' -Status $_status -Folder $SourceFolder.FolderPath -Item $_item -Detail $_detail

          if ($PSCmdlet.ShouldProcess("$($SourceFolder.FolderPath) | $([string]$_item.Subject)", "Move to $_destinationPath")) {
            $null = $_item.Move($DestinationFolder)
            $script:OutlookArchiveMoved++
          }
        }
        catch {
          Add-OperationResult -Results $Results -Target $SourceFolder.FolderPath -Source 'Outlook' -Action 'Move' -Status 'Failed' -Detail $_.Exception.Message
          Write-Warning "Move failed on item $_index in '$($SourceFolder.FolderPath)': $($_.Exception.Message)"
        }
        finally {
          Remove-ComObject $_item
        }
      }

      return
    }

    for ($_index = 1; $_index -le $_items.Count; $_index++) {
      $_item = $_items.Item($_index)
      $_copy = $null
      try {
        if (-not (Test-OutlookItemInRange -Item $_item)) { continue }

        $_destinationPath = if ($DestinationFolder) { $DestinationFolder.FolderPath } else { $_archivePath }
        $_status = if ($WhatIfPreference) { 'Skipped' } else { 'Copied' }
        $_detail = if ($WhatIfPreference) { 'DryRun' } else { "Copied to $_destinationPath" }
        Add-OutlookArchiveResult -Results $Results -Action 'Copy' -Status $_status -Folder $SourceFolder.FolderPath -Item $_item -Detail $_detail

        if ($PSCmdlet.ShouldProcess("$($SourceFolder.FolderPath) | $([string]$_item.Subject)", "Copy to $_destinationPath")) {
          $_copy = $_item.Copy()
          $null = $_copy.Move($DestinationFolder)
          $script:OutlookArchiveCopied++
        }
      }
      catch {
        Add-OperationResult -Results $Results -Target $SourceFolder.FolderPath -Source 'Outlook' -Action 'Copy' -Status 'Failed' -Detail $_.Exception.Message
        Write-Warning "Copy failed on item $_index in '$($SourceFolder.FolderPath)': $($_.Exception.Message)"
      }
      finally {
        Remove-ComObject $_copy $_item
      }
    }
  }
  finally {
    Remove-ComObject $_items
  }
}

function Copy-OutlookFolderTree {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $SourceFolder,

    [object]
    $DestinationParent,

    [string]
    $ArchiveMode,

    [System.Collections.IList]
    $Results
  )

  Write-Verbose "Archiving: $($SourceFolder.FolderPath)"
  $_destinationFolder = if ($WhatIfPreference) {
    $null
  }
  else {
    Get-OutlookSubFolder -ParentFolder $DestinationParent -Name $SourceFolder.Name -Create
  }

  try {
    Copy-OutlookFolderItem -SourceFolder $SourceFolder -DestinationFolder $_destinationFolder -ArchiveMode $ArchiveMode -Results $Results -WhatIf:$WhatIfPreference

    $_folders = $SourceFolder.Folders
    try {
      for ($_index = 1; $_index -le $_folders.Count; $_index++) {
        $_child = $_folders.Item($_index)
        try {
          Copy-OutlookFolderTree -SourceFolder $_child -DestinationParent $_destinationFolder -ArchiveMode $ArchiveMode -Results $Results -WhatIf:$WhatIfPreference
        }
        finally {
          Remove-ComObject $_child
        }
      }
    }
    finally {
      Remove-ComObject $_folders
    }
  }
  finally {
    Remove-ComObject $_destinationFolder
  }
}

$_context = $null
$_sourceRoot = $null
$_archiveRoot = $null
$_archivePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArchivePath)

try {
  if ($StartDate -and $EndDate -and $StartDate -gt $EndDate) {
    throw 'StartDate is after EndDate.'
  }

  $_context = Connect-Outlook
  $_sourceRoot = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $StoreName
  Write-Verbose "Source: $($_sourceRoot.FolderPath)"

  if (-not $WhatIfPreference) {
    $_archiveRoot = Add-OutlookStoreRoot -Namespace $_context.Namespace -Path $_archivePath
    try {
      $_archiveRoot.Name = $DisplayName
    }
    catch {
      Write-Verbose "Could not rename archive PST store to '$DisplayName': $($_.Exception.Message)"
    }
  }

  Write-Verbose "Archive PST: $_archivePath"

  $_folders = $_sourceRoot.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_child = $_folders.Item($_index)
      try {
        Copy-OutlookFolderTree -SourceFolder $_child -DestinationParent $_archiveRoot -ArchiveMode $Mode -Results $_results -WhatIf:$WhatIfPreference
      }
      finally {
        Remove-ComObject $_child
      }
    }
  }
  finally {
    Remove-ComObject $_folders
  }
}
finally {
  if ($DetachWhenDone -and $_context -and $_archiveRoot) {
    try {
      $_context.Namespace.RemoveStore($_archiveRoot)
    }
    catch {
      Add-OperationResult -Results $_results -Target $_archivePath -Source 'Outlook' -Action 'DetachStore' -Status 'Failed' -Detail $_.Exception.Message
      Write-Warning "Could not detach PST: $($_.Exception.Message)"
    }
  }

  Remove-ComObject $_archiveRoot $_sourceRoot

  if ($_context) {
    try {
      if ($QuitOutlook) {
        $_context.App.Quit()
      }
    }
    catch {
      Write-Verbose "Could not quit Outlook: $($_.Exception.Message)"
    }

    Remove-ComObject $_context.Namespace $_context.App
  }

  Invoke-ComGarbageCollection
}

$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count
if ($WhatIfPreference) {
  $_planned = @($_results | Where-Object { $_.Detail -eq 'DryRun' }).Count
  Write-Log -Message "Outlook archive preview complete. Items to process: $_planned | Failed: $_failed | PST: $_archivePath" -Color Yellow
}
else {
  Write-Log -Message "Outlook archive complete. Copied: $script:OutlookArchiveCopied | Moved: $script:OutlookArchiveMoved | Failed: $_failed | PST: $_archivePath" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'New-OutlookArchive'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failed -gt 0) {
  exit 1
}
