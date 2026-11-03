#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.1' }

<#
.SYNOPSIS
  Builds HTML and DOCX from a saved press-release document and an explicit template pack.
.DESCRIPTION
  Validates the pack and source before writing, stages both outputs, and publishes a
  build.manifest.json last. Language and template version come from source metadata.
  Existing files require Overwrite. Publication uses individual atomic file renames
  with rollback on failure; the set is not a filesystem transaction. Use one output
  directory per build and keep the input and pack unchanged while building.
  PDF conversion is a separate caller/UI operation using Convert-HTMLToPDF.ps1.
  A PDF failure does not invalidate this HTML/DOCX build. No UI is provided.
.PARAMETER InputDocument
  Existing saved .docx source. Authoring .dotx templates are not accepted.
.PARAMETER PackPath
  Explicit directory containing pack.json and its verified template assets.
.PARAMETER OutputDirectory
  Directory for HTML, DOCX and build.manifest.json; created after validation.
.PARAMETER OutputName
  Safe file basename without a directory. Defaults to the input document basename.
.PARAMETER Overwrite
  Allow replacement of existing outputs and the directory's build manifest.
.PARAMETER DryRun
  Set WhatIfPreference: validate and report intended paths without writing files.
.PARAMETER PassThru
  Return status, metadata, output paths/hashes, diagnostics and manifest path.
  Failures also terminate with an error whose TargetObject contains this result.
.EXAMPLE
  PS> .\Build-PressRelease.ps1 -InputDocument C:\Releases\Input\release.docx -PackPath C:\Packs\v1 -OutputDirectory C:\Releases\Output -PassThru
  Build both formats and a manifest without changing the saved source.
.NOTES
  Author: Ad Noctem Collective
  License: MIT
  Server Core: supported with PowerShell 7.4 or later; no Word installation needed.
  SYSTEM-account execution: supported when input, pack and output paths are accessible.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $InputDocument,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $PackPath,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]
  $OutputDirectory,

  [string]
  $OutputName,

  [switch]
  $Overwrite,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force

$ErrorActionPreference = 'Stop'
if ($DryRun) {
  $WhatIfPreference = $true
}

$_stage = $null
$_lock = $null
$_inputLock = $null
$_published = [Collections.Generic.List[string]]::new()
$_backups = @{}
$_result = [pscustomobject]@{
  Target         = $OutputDirectory
  Action         = 'BuildPressRelease'
  Status         = 'Failed'
  Changed        = $false
  Detail         = $null
  ErrorMessage   = $null
  Language       = $null
  PackVersion    = $null
  HtmlPath       = $null
  HtmlHash       = $null
  DocxPath       = $null
  DocxHash       = $null
  ManifestPath   = $null
  Diagnostics    = @()
  PublishedPaths = @()
}

function Resolve-BuildPath {
  param (
    [string]
    $Path
  )

  $provider = $null
  $drive = $null
  $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path, [ref]$provider, [ref]$drive)
  if ($provider.Name -ne 'FileSystem') {
    throw "Only filesystem paths are supported: $Path"
  }

  # Reject aliases through junctions/symlinks, including existing parent directories.
  $current = $resolved
  while ($current) {
    if (Test-Path -LiteralPath $current) {
      $item = Get-Item -LiteralPath $current -Force
      if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        throw "Reparse points are not supported in build paths: $current"
      }
    }
    $current = Split-Path -Path $current -Parent
  }
  return [IO.Path]::GetFullPath($resolved)
}

function Assert-BuildDestination {
  param (
    [string[]]
    $Path
  )

  foreach ($destination in $Path) {
    $null = Resolve-BuildPath -Path $destination
    if (Test-Path -LiteralPath $destination) {
      if (-not (Test-Path -LiteralPath $destination -PathType Leaf)) {
        throw "Output destination must be a file: $destination"
      }
      if (-not $Overwrite) {
        throw "Output already exists; use -Overwrite: $destination"
      }
    }
  }
}

try {
  # Resolve paths and protect the source and pack from output publication.
  $_input = Resolve-BuildPath -Path $InputDocument
  $_packPath = Resolve-BuildPath -Path $PackPath
  $_output = Resolve-BuildPath -Path $OutputDirectory
  if ([IO.Path]::GetExtension($_input) -ne '.docx') {
    throw 'InputDocument must be a saved .docx file, not an authoring .dotx template.'
  }
  if (-not (Test-Path -LiteralPath $_input -PathType Leaf)) {
    throw "InputDocument does not exist: $_input"
  }
  if (-not (Test-Path -LiteralPath $_packPath -PathType Container)) {
    throw "PackPath must be an existing directory: $_packPath"
  }
  if ((Test-Path -LiteralPath $_output) -and -not (Test-Path -LiteralPath $_output -PathType Container)) {
    throw "OutputDirectory must be a directory: $_output"
  }
  if (-not $PSBoundParameters.ContainsKey('OutputName')) {
    $OutputName = [IO.Path]::GetFileNameWithoutExtension($_input)
  }
  if ([string]::IsNullOrWhiteSpace($OutputName) -or
    $OutputName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
    $OutputName -match '[/\\]|[. ]$|^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])($|\.)') {
    throw 'OutputName must be a single safe file basename without separators, reserved names, or trailing dots/spaces.'
  }
  $_packPrefix = $_packPath.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
  if ($_output -ieq $_packPath -or $_output.StartsWith($_packPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputDirectory must be outside the template pack.'
  }

  $_result.HtmlPath = Join-Path $_output ($OutputName + '.html')
  $_result.DocxPath = Join-Path $_output ($OutputName + '.docx')
  $_result.ManifestPath = Join-Path $_output 'build.manifest.json'
  if ($_result.DocxPath -ieq $_input) {
    throw 'The generated DOCX must never be the input path.'
  }
  $_destinations = @($_result.HtmlPath, $_result.DocxPath, $_result.ManifestPath)

  # Keep the source stable through hashing, reading and publication.
  $_inputLock = [IO.File]::Open($_input, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
  $_inputHash = (Get-FileHash -LiteralPath $_input -Algorithm SHA256).Hash
  $_packManifest = Resolve-BuildPath -Path (Join-Path $_packPath 'pack.json')
  $_packHash = $null
  if (Test-Path -LiteralPath $_packManifest -PathType Leaf) {
    $_packHash = (Get-FileHash -LiteralPath $_packManifest -Algorithm SHA256).Hash
  }
  $_pack = [AdNoctem.Substrate.Documents.TemplatePack]::Load([AdNoctem.Substrate.IO.FileSystemPath]::Parse($_packPath))
  if ((Get-FileHash -LiteralPath $_packManifest -Algorithm SHA256).Hash -ne $_packHash) {
    throw 'The pack manifest changed during loading; retry with an immutable pack.'
  }
  $_result.PackVersion = $_pack.Version
  $_reader = [AdNoctem.Substrate.Documents.PressReleaseReader]::new()
  $_read = $_reader.Read([AdNoctem.Substrate.IO.FileSystemPath]::Parse($_input), $_pack.Schema)
  $_result.Diagnostics = @(
    foreach ($diagnostic in $_read.Diagnostics) {
      [pscustomobject]@{
        Severity = [string]$diagnostic.Severity
        Code     = [string]$diagnostic.Code
        Field    = $diagnostic.Field
        Location = $diagnostic.Location
        Message  = $diagnostic.Message
        Remedy   = $diagnostic.Remedy
      }
    }
  )
  foreach ($diagnostic in $_result.Diagnostics) {
    $_message = '{0} | {1} | Field: {2} | {3} | Remedy: {4}' -f
    $diagnostic.Severity, $diagnostic.Code, $diagnostic.Field, $diagnostic.Message, $diagnostic.Remedy
    if ($diagnostic.Severity -eq 'Error') {
      Write-Error -Message $_message -ErrorId $diagnostic.Code -Category InvalidData -ErrorAction Continue
    }
    else {
      Write-Log -Message $_message -Color Yellow
    }
  }
  if ($null -eq $_read.Release -or @($_result.Diagnostics | Where-Object Severity -EQ 'Error').Count -gt 0) {
    throw 'Source validation failed. Correct the reported diagnostics before building.'
  }
  $_result.Language = $_read.Release.Language
  Assert-BuildDestination -Path $_destinations

  if ($PSCmdlet.ShouldProcess($_output, "Publish $OutputName.html, $OutputName.docx and build.manifest.json")) {
    $null = [IO.Directory]::CreateDirectory($_output)
    $_lockPath = Join-Path $_output '.press-release-build.lock'
    $_lock = [IO.FileStream]::new($_lockPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write,
      [IO.FileShare]::None, 4096, [IO.FileOptions]::DeleteOnClose)
    Assert-BuildDestination -Path $_destinations
    $_stage = Join-Path $_output ('.press-release-' + [guid]::NewGuid().ToString('N'))
    $null = [IO.Directory]::CreateDirectory($_stage)
    $_htmlStage = Join-Path $_stage ($OutputName + '.html')
    $_docxStage = Join-Path $_stage ($OutputName + '.docx')
    $_manifestStage = Join-Path $_stage 'build.manifest.json'

    [AdNoctem.Substrate.Documents.PressReleaseHtmlWriter]::new().Write(
      $_read.Release, $_pack, [AdNoctem.Substrate.IO.FileSystemPath]::Parse($_htmlStage), $true)
    [AdNoctem.Substrate.Documents.PressReleaseDocxWriter]::new().Write(
      $_read.Release, $_pack, [AdNoctem.Substrate.IO.FileSystemPath]::Parse($_docxStage), $true)
    $_result.HtmlHash = (Get-FileHash -LiteralPath $_htmlStage -Algorithm SHA256).Hash
    $_result.DocxHash = (Get-FileHash -LiteralPath $_docxStage -Algorithm SHA256).Hash
    $_manifest = [ordered]@{
      schema         = 'press-release-build/1'
      generatedAtUtc = [DateTime]::UtcNow.ToString('o')
      tool           = [ordered]@{
        name             = 'winkit/Build-PressRelease'
        version          = '1.0.0'
        scriptSha256     = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
        substrateVersion = (Get-Module AdNoctem.Substrate.PowerShell).Version.ToString()
      }
      input          = [ordered]@{
        path            = $_input
        sha256          = $_inputHash
        language        = $_read.Release.Language
        templateVersion = $_read.Release.TemplateVersion
        schema          = [AdNoctem.Substrate.Documents.PressReleaseTags]::SchemaId
      }
      pack           = [ordered]@{
        path                = $_packPath
        version             = $_pack.Version
        outputLayoutVersion = $_pack.OutputLayoutVersion
        sha256              = $_packHash
      }
      outputs        = [ordered]@{
        html = [ordered]@{
          path   = $_result.HtmlPath
          sha256 = $_result.HtmlHash
        }
        docx = [ordered]@{
          path   = $_result.DocxPath
          sha256 = $_result.DocxHash
        }
      }
      diagnostics    = $_result.Diagnostics
      status         = 'Completed'
    }
    [IO.File]::WriteAllText($_manifestStage, (ConvertTo-Json -InputObject $_manifest -Depth 10), [Text.UTF8Encoding]::new($false))

    # All generation has succeeded. Back up originals before the first rename.
    Assert-BuildDestination -Path $_destinations
    foreach ($destination in $_destinations) {
      if ([IO.File]::Exists($destination)) {
        $backup = Join-Path $_stage ([guid]::NewGuid().ToString('N') + '.backup')
        [IO.File]::Copy($destination, $backup, $false)
        $_backups[$destination] = $backup
      }
    }
    $_sources = @($_htmlStage, $_docxStage, $_manifestStage)
    for ($index = 0; $index -lt $_destinations.Count; $index++) {
      [IO.File]::Move($_sources[$index], $_destinations[$index], [bool]$Overwrite)
      $_published.Add($_destinations[$index])
    }
    $_result.Status = 'Completed'
    $_result.Changed = $true
    $_result.PublishedPaths = @($_published.ToArray())
    $_result.Detail = 'HTML, DOCX and manifest published.'
  }
  else {
    $_result.Status = 'Skipped'
    $_result.Detail = 'Validation succeeded; publication was not requested by ShouldProcess.'
  }
}
catch {
  $_originalError = $_
  $_exception = $_.Exception
  while ($_exception.InnerException -and $_exception -isnot [AdNoctem.Substrate.Documents.DocumentException]) {
    $_exception = $_exception.InnerException
  }
  $_result.ErrorMessage = $_exception.Message
  if ($_exception -is [AdNoctem.Substrate.Documents.DocumentException]) {
    $_result.ErrorMessage = '{0}: {1}' -f $_exception.Reason, $_exception.Message
  }
  if ($_published.Count -gt 0) {
    $_result.ErrorMessage += ' Failure occurred after publication began.'
    $_rollbackPaths = $_published.ToArray()
    [array]::Reverse($_rollbackPaths)
    foreach ($destination in $_rollbackPaths) {
      try {
        if ($_backups.ContainsKey($destination)) {
          [IO.File]::Move($_backups[$destination], $destination, $true)
        }
        else {
          [IO.File]::Delete($destination)
        }
        $null = $_published.Remove($destination)
      }
      catch {
        $_result.ErrorMessage += " Rollback failed for ${destination}: $($_.Exception.Message). Recovery files: $_stage"
      }
    }
    $_result.ErrorMessage += " Unrestored output count: $($_published.Count)."
  }
  $_result.PublishedPaths = @($_published.ToArray())
  $_result.Changed = $_published.Count -gt 0
  Write-Log -Message "Press-release build failed: $($_result.ErrorMessage)" -Color Red
  if ($PassThru) {
    $_result
  }
  $_errorRecord = [Management.Automation.ErrorRecord]::new(
    [InvalidOperationException]::new($_result.ErrorMessage, $_originalError.Exception),
    'PressReleaseBuildFailed', [Management.Automation.ErrorCategory]::InvalidOperation, $_result)
  $PSCmdlet.ThrowTerminatingError($_errorRecord)
}
finally {
  if ($_inputLock) {
    $_inputLock.Dispose()
  }
  if ($_stage -and ($_result.Status -eq 'Completed' -or $_published.Count -eq 0)) {
    try {
      $_stageFull = [IO.Path]::GetFullPath($_stage)
      if ([IO.Path]::GetDirectoryName($_stageFull) -ne $_output -or
        [IO.Path]::GetFileName($_stageFull) -notmatch '^\.press-release-[a-f0-9]{32}$') {
        throw 'Staging cleanup path is outside the output directory.'
      }
      Remove-Item -LiteralPath $_stageFull -Recurse -Force -ErrorAction Stop
    }
    catch {
      Write-Log -Message "Could not clean staging directory ${_stage}: $($_.Exception.Message)" -Color Yellow
    }
  }
  if ($_lock) {
    $_lock.Dispose()
  }
}

Write-Log -Message "Press-release build: $($_result.Status). $($_result.Detail)" -Color Cyan
if ($PassThru) {
  $_result
}
