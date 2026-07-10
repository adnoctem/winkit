Import-Module PSFoundation -Force

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Deduplicates Outlook mail items per folder using the transport Message-ID.
.DESCRIPTION
  Walks every mail folder in the target Outlook store, keys each received mail
  item by its RFC Message-ID, and moves every occurrence after the first into a
  review folder. It never hard-deletes messages.

  Deduplication scope is per folder. The same message legitimately living in
  two different folders is preserved.
.PARAMETER StoreName
  Display name of the Outlook store to process. If omitted, the default
  delivery store is used. Run with -Verbose to list detected stores.
.PARAMETER ReviewFolderName
  Top-level folder created under the store root for duplicate review.
.PARAMETER ExcludeFolders
  Folder display names to skip entirely.
.PARAMETER ReportPath
  Optional CSV path containing Keep, MoveDuplicate, and SkipNoMessageId results.
.PARAMETER DryRun
  Preview changes without moving duplicate messages.
.PARAMETER PassThru
  Return structured operation result objects.
.PARAMETER QuitOutlook
  Quit the Outlook application object on exit. Leave off if Outlook was already
  open interactively.
.EXAMPLE
  PS> .\Optimize-Outlook.ps1 -StoreName 'user@example.com' -ReportPath .\dedup-preview.csv -DryRun
.EXAMPLE
  PS> .\Optimize-Outlook.ps1 -StoreName 'user@example.com' -ReportPath .\dedup-run.csv
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
  [string]
  $StoreName,

  [string]
  $ReviewFolderName = '_Duplicates_Review',

  [string[]]
  $ExcludeFolders = @(
    'Deleted Items',
    'Junk Email',
    'Junk E-mail',
    'Outbox',
    'Sync Issues',
    'Conflicts',
    'Local Failures',
    'Server Failures'
  ),

  [string]
  $ReportPath,

  [switch]
  $DryRun,

  [switch]
  $PassThru,

  [switch]
  $QuitOutlook
)

# -----------------------------------------------------------------------------

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no Outlook messages will be moved`n" -Color Yellow
}

# olMail object class; PR_TRANSPORT_MESSAGE_HEADERS in Unicode then ANSI form.
$script:OL_MAIL = 43
$script:HDR_TAGS = @(
  'http://schemas.microsoft.com/mapi/proptag/0x007D001F',
  'http://schemas.microsoft.com/mapi/proptag/0x007D001E'
)

$_results = New-Object System.Collections.ArrayList

function Get-MessageId {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'Outlook exposes either Unicode or ANSI headers depending on the item.')]
  param (
    [object]
    $Item
  )

  $_propertyAccessor = $null
  try {
    $_propertyAccessor = $Item.PropertyAccessor
    foreach ($_tag in $script:HDR_TAGS) {
      try {
        $_headers = $_propertyAccessor.GetProperty($_tag)
        if (-not [string]::IsNullOrWhiteSpace($_headers)) {
          $_match = [regex]::Match([string]$_headers, '(?im)^Message-ID:\s*(<[^>]+>)')
          if ($_match.Success) {
            return $_match.Groups[1].Value.Trim()
          }
        }
      }
      catch { }
    }
  }
  catch {
    return $null
  }
  finally {
    Remove-ComObject $_propertyAccessor
  }

  return $null
}

function Add-OutlookItemResult {
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
    $MessageId,

    [string]
    $Detail
  )

  $_subject = ''
  $_received = $null
  try { $_subject = [string]$Item.Subject } catch { $_subject = '' }
  try { $_received = $Item.ReceivedTime } catch { $_received = $null }

  $_property = @{
    Received = $_received
    MessageId = $MessageId
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

function Optimize-OutlookFolder {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $Folder,

    [object]
    $ReviewFolder,

    [System.Collections.IList]
    $Results
  )

  $_folderPath = $Folder.FolderPath
  $_items = $Folder.Items
  $_seen = @{}

  try {
    for ($_index = $_items.Count; $_index -ge 1; $_index--) {
      $_item = $_items.Item($_index)
      try {
        if ($_item.Class -ne $script:OL_MAIL) { continue }

        $_messageId = Get-MessageId -Item $_item
        if (-not $_messageId) {
          Add-OutlookItemResult -Results $Results -Action 'Deduplicate' -Status 'Skipped' -Folder $_folderPath -Item $_item -MessageId '' -Detail 'NoMessageId'
          continue
        }

        if ($_seen.ContainsKey($_messageId)) {
          $_status = if ($WhatIfPreference) { 'Skipped' } else { 'Moved' }
          $_detail = if ($WhatIfPreference) { 'DryRun' } else { 'Duplicate moved to review folder.' }
          Add-OutlookItemResult -Results $Results -Action 'MoveDuplicate' -Status $_status -Folder $_folderPath -Item $_item -MessageId $_messageId -Detail $_detail

          $_reviewPath = if ($ReviewFolder) { $ReviewFolder.FolderPath } else { $ReviewFolderName }
          if ($PSCmdlet.ShouldProcess("$_folderPath | $([string]$_item.Subject)", "Move duplicate to $_reviewPath")) {
            $null = $_item.Move($ReviewFolder)
          }
        }
        else {
          $_seen[$_messageId] = $true
          Add-OutlookItemResult -Results $Results -Action 'Deduplicate' -Status 'Kept' -Folder $_folderPath -Item $_item -MessageId $_messageId -Detail 'First item with Message-ID in folder.'
        }
      }
      catch {
        Add-OperationResult -Results $Results -Target $_folderPath -Source 'Outlook' -Action 'Deduplicate' -Status 'Failed' -Detail $_.Exception.Message
        Write-Warning "Item $_index in '$_folderPath': $($_.Exception.Message)"
      }
      finally {
        Remove-ComObject $_item
      }
    }
  }
  finally {
    Remove-ComObject $_items
  }
}

function Invoke-OutlookFolderTree {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [object]
    $Folder,

    [object]
    $ReviewFolder,

    [string]
    $ReviewName,

    [string[]]
    $Exclude,

    [System.Collections.IList]
    $Results
  )

  if ($Folder.Name -eq $ReviewName) { return }

  if ($Exclude -contains $Folder.Name) {
    Write-Verbose "Skipping excluded folder: $($Folder.Name)"
  }
  else {
    Write-Verbose "Processing: $($Folder.FolderPath)"
    Optimize-OutlookFolder -Folder $Folder -ReviewFolder $ReviewFolder -Results $Results -WhatIf:$WhatIfPreference
  }

  $_folders = $Folder.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_child = $_folders.Item($_index)
      try {
        Invoke-OutlookFolderTree -Folder $_child -ReviewFolder $ReviewFolder -ReviewName $ReviewName -Exclude $Exclude -Results $Results -WhatIf:$WhatIfPreference
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

$_context = $null
$_storeRoot = $null
$_reviewFolder = $null

try {
  $_context = Connect-Outlook
  $_storeRoot = Get-OutlookStoreRoot -Namespace $_context.Namespace -Name $StoreName
  Write-Verbose "Store root: $($_storeRoot.FolderPath)"

  if (-not $WhatIfPreference) {
    $_reviewFolder = Get-OutlookSubFolder -ParentFolder $_storeRoot -Name $ReviewFolderName -Create
  }

  $_folders = $_storeRoot.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_child = $_folders.Item($_index)
      try {
        Invoke-OutlookFolderTree -Folder $_child -ReviewFolder $_reviewFolder -ReviewName $ReviewFolderName -Exclude $ExcludeFolders -Results $_results -WhatIf:$WhatIfPreference
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
  Remove-ComObject $_reviewFolder
  Remove-ComObject $_storeRoot

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

$_moved = @($_results | Where-Object { $_.Action -eq 'MoveDuplicate' -and $_.Status -eq 'Moved' }).Count
$_planned = @($_results | Where-Object { $_.Action -eq 'MoveDuplicate' -and $_.Detail -eq 'DryRun' }).Count
$_kept = @($_results | Where-Object { $_.Status -eq 'Kept' }).Count
$_skipped = @($_results | Where-Object { $_.Detail -eq 'NoMessageId' }).Count
$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count

if ($WhatIfPreference) {
  Write-Log -Message "Outlook deduplication preview complete. Duplicates to move: $_planned | Kept: $_kept | Skipped: $_skipped | Failed: $_failed" -Color Yellow
}
else {
  Write-Log -Message "Outlook deduplication complete. Moved: $_moved | Kept: $_kept | Skipped: $_skipped | Failed: $_failed" -Color $(if ($_failed -gt 0) { 'Yellow' } else { 'Green' })
}

if ($ReportPath) {
  $_reportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
  $_reportRoot = Split-Path -Path $_reportPath -Parent
  if (-not [string]::IsNullOrWhiteSpace($_reportRoot) -and -not (Test-Path -LiteralPath $_reportRoot)) {
    $null = New-Item -Path $_reportRoot -ItemType Directory -Force
  }

  $_results | Export-Csv -Path $_reportPath -NoTypeInformation -Encoding UTF8
  Write-Log -Message "Report: $_reportPath" -Color Gray
}

$_operationLog = Write-OperationResultLog -Results $_results -ScriptName 'Optimize-Outlook'
if ($_operationLog) {
  Write-Log -Message "Operation log: $_operationLog" -Color Gray
}

if ($PassThru -or $DryRun) {
  $_results
}

if ($_failed -gt 0) {
  exit 1
}
