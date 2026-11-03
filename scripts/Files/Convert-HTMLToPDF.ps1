#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.1' }

<#
.SYNOPSIS
  Converts HTML to PDF and records a separate PDF audit manifest.
.DESCRIPTION
  Uses EuroPDF (default), DocRaptor, or local PSWritePDF. Cloud modes send HTML to
  the selected service and can incur charges. Test requests are watermarked.
  Saves binary responses directly to staging files, verifies a PDF signature,
  and publishes the PDF and pdf.manifest.json. This is not conformance validation.
  Conversion failures record sanitized errors and null output in the manifest.
  The HTML/DOCX build and build.manifest.json are never modified.
  Use one output directory per conversion and keep it unchanged during the run.
  Preflight errors and previews do not write manifests. Existing PDFs or manifests
  require Overwrite. An unsuccessful retry preserves the old PDF but writes a
  failure manifest for the latest attempt; it does not certify the old PDF.
.PARAMETER InputPath
  Existing .html or .htm file. Path is an alias. UTF-8 is recommended.
.PARAMETER OutputPath
  Output .pdf path. Defaults to the input basename with a .pdf extension.
  The output directory also contains pdf.manifest.json.
.PARAMETER Mode
  EuroPDF, DocRaptor, or PS (local PSWritePDF). Default: EuroPDF.
.PARAMETER ApiKey
  SecureString API key, for example Read-Host -AsSecureString. If omitted, reads
  only EUROPDF_API_KEY for EuroPDF or DOCRAPTOR_API_KEY for DocRaptor.
.PARAMETER Async
  Submit a cloud job and wait for completion. No detached job is returned.
.PARAMETER Profile
  Optional Prince PDF profile for cloud modes. No profile is forced by default.
  Archival/accessibility profiles impose requirements on the HTML and assets.
.PARAMETER Pipeline
  Optional cloud pipeline. EuroPDF uses names such as Prince16.2 or
  Prince16-msfonts; DocRaptor uses numeric versions such as 10.1.
.PARAMETER Test
  Request a watermarked cloud test document. Still sends HTML to the service.
.PARAMETER PollInterval
  Seconds between async status checks. Default: 5.
.PARAMETER Timeout
  Total cloud-operation timeout in seconds, including submission and download.
  Default: 900. A timeout does not cancel a job already submitted to the service.
.PARAMETER NoNetwork
  Disable external resource fetching in Prince. Does not disable API calls.
.PARAMETER Strict
  Request strict HTML validation. Only valid for DocRaptor.
.PARAMETER Overwrite
  Replace an existing PDF or pdf.manifest.json. Retains HTML/DOCX build artifacts.
.PARAMETER DryRun
  Validate local paths and options without credentials, network calls, or writes.
  WhatIf has the same effect; Confirm gates the entire conversion.
.PARAMETER PassThru
  Return an operation result with Status, InputPath,
  InputHash, OutputPath, OutputHash, ManifestPath, ManifestWritten and Engine.
  On failure the terminating error's TargetObject contains the same result.
.EXAMPLE
  PS> .\Convert-HTMLToPDF.ps1 -InputPath C:\Releases\release.html -DryRun -PassThru
  Preview without contacting a service.
.EXAMPLE
  PS> $key = Read-Host 'EuroPDF API key' -AsSecureString
  PS> .\Convert-HTMLToPDF.ps1 -InputPath C:\Releases\release.html -ApiKey $key -Test -PassThru
  Generate a watermarked test PDF with the default PDF profile.
.EXAMPLE
  PS> .\Convert-HTMLToPDF.ps1 -InputPath C:\Releases\release.html -Mode DocRaptor -Async -Profile 'PDF/A-3a+PDF/UA-1' -Overwrite -PassThru
  Use DOCRAPTOR_API_KEY and wait for an asynchronous conversion.
.EXAMPLE
  PS> .\Convert-HTMLToPDF.ps1 -InputPath C:\Releases\release.html -Mode PS -PassThru
  Convert locally with PSWritePDF.
.LINK
  https://www.europdf.eu/docs/api/
  https://docraptor.com/documentation/api
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported when the selected backend is available; no Word needed.
  SYSTEM-account execution: supported with accessible paths and backend credentials.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [Alias('Path')]
  [ValidateNotNullOrEmpty()]
  [string]
  $InputPath,

  [string]
  $OutputPath,

  [ValidateSet('EuroPDF', 'PS', 'DocRaptor')]
  [string]
  $Mode = 'EuroPDF',

  [SecureString]
  $ApiKey,

  [switch]
  $Async,

  [ValidateSet(
    'PDF/A-1a',
    'PDF/A-1b',
    'PDF/A-2a',
    'PDF/A-2b',
    'PDF/A-3a',
    'PDF/A-3b',
    'PDF/UA-1',
    'PDF/X-1a:2001',
    'PDF/X-1a:2003',
    'PDF/X-3:2002',
    'PDF/X-3:2003',
    'PDF/X-4',
    'PDF/A-1a+PDF/UA-1',
    'PDF/A-2a+PDF/UA-1',
    'PDF/A-3a+PDF/UA-1',
    IgnoreCase = $false
  )]
  [string]
  $Profile,

  [string]
  $Pipeline,

  [switch]
  $Test,

  [ValidateRange(1, 300)]
  [int]
  $PollInterval = 5,

  [ValidateRange(10, 3600)]
  [int]
  $Timeout = 900,

  [switch]
  $NoNetwork,

  [switch]
  $Strict,

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
$_manifest = $null
$_manifestWritten = $false
$_failure = $null
$_phase = 'preflight'
$_status = 'Failed'
$_outputHash = $null
$_inputHash = $null
$_manifestPath = $null
$_input = $InputPath
$_output = $OutputPath
$_published = $false
$_rollbackFailed = $false

function New-PdfFailure {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Constructs an exception without changing external state.')]
  [CmdletBinding()]
  param (
    [string]
    $Code,

    [string]
    $Message
  )

  $exception = [InvalidOperationException]::new($Message)
  $exception.Data['PdfErrorCode'] = $Code
  return $exception
}

function Resolve-PdfPath {
  param (
    [string]
    $Path
  )

  $provider = $null
  $drive = $null
  $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path, [ref]$provider, [ref]$drive)
  if ($provider.Name -ne 'FileSystem') {
    throw (New-PdfFailure 'InvalidPath' 'Only filesystem paths are supported.')
  }
  $ancestor = $resolved
  while ($ancestor) {
    if (Test-Path -LiteralPath $ancestor) {
      if ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
        throw (New-PdfFailure 'InvalidPath' 'Junctions and symbolic links are not supported in conversion paths.')
      }
    }
    $ancestor = [IO.Path]::GetDirectoryName($ancestor)
  }
  return $resolved
}

function Get-PdfAuthorization {
  param (
    [SecureString]
    $Key,

    [string]
    $Engine
  )

  $environmentName = if ($Engine -eq 'EuroPDF') { 'EUROPDF_API_KEY' } else { 'DOCRAPTOR_API_KEY' }
  $plainKey = if ($Key) {
    [pscredential]::new('api', $Key).GetNetworkCredential().Password
  }
  else {
    [Environment]::GetEnvironmentVariable($environmentName, 'Process')
  }
  if ([string]::IsNullOrWhiteSpace($plainKey)) {
    throw (New-PdfFailure 'MissingCredential' "Supply a SecureString ApiKey or set $environmentName for this engine.")
  }
  return 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($plainKey + ':'))
}

function Invoke-PdfRequest {
  param (
    [string]
    $Uri,

    [string]
    $Authorization,

    [string]
    $Destination,

    [Diagnostics.Stopwatch]
    $Clock,

    [string]
    $Json
  )

  $remaining = [int][Math]::Ceiling($Timeout - $Clock.Elapsed.TotalSeconds)
  if ($remaining -le 0) {
    throw (New-PdfFailure 'Timeout' 'PDF conversion timed out. A submitted cloud job may still complete; inspect the service dashboard before retrying.')
  }
  $request = @{
    Uri                = $Uri
    Method             = 'Get'
    Headers            = @{ Authorization = $Authorization }
    OutFile            = $Destination
    PassThru           = $true
    SkipHttpErrorCheck = $true
    MaximumRedirection = 0
    TimeoutSec         = $remaining
    ErrorAction        = 'Stop'
  }
  if ($Json) {
    $request.Method = 'Post'
    $request.ContentType = 'application/json; charset=utf-8'
    $request.Body = [Text.Encoding]::UTF8.GetBytes($Json)
  }
  if ((Get-Command Invoke-WebRequest).Parameters.ContainsKey('OperationTimeoutSeconds')) {
    $request.OperationTimeoutSeconds = $remaining
  }
  try {
    $response = Invoke-WebRequest @request
  }
  catch {
    # Original errors can contain credentials, HTML, or PDF bytes; never retain them.
    throw (New-PdfFailure 'RequestFailed' 'The PDF service request failed or timed out. Check connectivity and the service dashboard before retrying; no submission was retried automatically.')
  }
  if ([int]$response.StatusCode -ne 200) {
    $hint = switch ([int]$response.StatusCode) {
      401 { 'Check the API key for the selected engine.' }
      403 { 'Check the API key and account permissions.' }
      422 { 'Check the service dashboard for input/profile validation or account quota errors.' }
      429 { 'The service request limit was reached; retry later.' }
      default { 'Check the service dashboard for details.' }
    }
    throw (New-PdfFailure 'HttpError' ("{0} returned HTTP {1}. {2} Response bodies are omitted from errors and manifests." -f $Mode, [int]$response.StatusCode, $hint))
  }
}

function Read-PdfJob {
  param (
    [string]
    $Path
  )

  try {
    if ((Get-Item -LiteralPath $Path).Length -gt 1MB) {
      throw 'Unexpected job response size.'
    }
    return [IO.File]::ReadAllText($Path) | ConvertFrom-Json -ErrorAction Stop
  }
  catch {
    throw (New-PdfFailure 'InvalidResponse' 'The PDF service returned an invalid job response. Inspect the service dashboard.')
  }
}

function Invoke-CloudPdfConversion {
  param (
    [string]
    $Html,

    [string]
    $Destination,

    [string]
    $JobPath
  )

  $authorization = Get-PdfAuthorization -Key $ApiKey -Engine $Mode
  $baseUrl = if ($Mode -eq 'EuroPDF') { 'https://api.europdf.eu/v1' } else { 'https://api.docraptor.com' }
  $body = @{
    document_content = $Html
  }
  if ($Mode -eq 'DocRaptor') {
    $body.type = 'pdf'
  }
  $prince = @{}
  if ($Profile) {
    $prince.profile = $Profile
  }
  if ($NoNetwork) {
    $prince.no_network = $true
  }
  if ($prince.Count) {
    $body.prince_options = $prince
  }
  if ($Pipeline) {
    $body.pipeline = $Pipeline
  }
  if ($Test) {
    $body.test = $true
  }
  if ($Strict) {
    $body.strict = 'html'
  }
  if ($Async) {
    $body.async = $true
  }

  $clock = [Diagnostics.Stopwatch]::StartNew()
  $request = @{
    Uri           = "$baseUrl/docs"
    Authorization = $authorization
    Destination   = $Destination
    Clock         = $clock
    Json          = ConvertTo-Json -InputObject $body -Depth 5 -Compress
  }
  if (-not $Async) {
    Invoke-PdfRequest @request
    return
  }
  $request.Destination = $JobPath
  Invoke-PdfRequest @request
  $job = Read-PdfJob $JobPath
  if ([string]$job.status_id -notmatch '^[a-zA-Z0-9_-]+$') {
    throw (New-PdfFailure 'InvalidResponse' 'The PDF service did not return a valid status identifier.')
  }
  $request.Remove('Json')
  $request.Uri = "$baseUrl/status/$($job.status_id)"

  while ($true) {
    Invoke-PdfRequest @request
    $job = Read-PdfJob $JobPath
    if ($job.status -eq 'completed') {
      break
    }
    if ($job.status -eq 'failed' -or $job.status -eq 'incinerated') {
      throw (New-PdfFailure 'JobFailed' 'The PDF service reported a failed or expired job. Inspect its dashboard for document/profile errors.')
    }
    if ($job.status -notin @('requested', 'started', 'queued', 'working')) {
      throw (New-PdfFailure 'InvalidResponse' 'The PDF service returned an unknown job status.')
    }
    $remaining = $Timeout - $clock.Elapsed.TotalSeconds
    if ($remaining -le 0) {
      throw (New-PdfFailure 'Timeout' 'PDF conversion timed out. Inspect the service dashboard before retrying.')
    }
    Start-Sleep -Milliseconds ([int](1000 * [Math]::Min($PollInterval, $remaining)))
  }

  if ($Mode -eq 'EuroPDF' -and [string]$job.download_id -match '^[a-zA-Z0-9_-]+$') {
    $download = [uri]"$baseUrl/download/$($job.download_id)"
  }
  else {
    $download = $null
    $validUri = [uri]::TryCreate([string]$job.download_url, [UriKind]::Absolute, [ref]$download)
    $hosts = if ($Mode -eq 'EuroPDF') { @('api.europdf.eu') } else { @('api.docraptor.com', 'docraptor.com') }
    if (-not $validUri -or $download.Scheme -ne 'https' -or $download.Host -notin $hosts -or
      $download.Port -ne 443 -or $download.UserInfo -or $download.Fragment) {
      throw (New-PdfFailure 'InvalidResponse' 'The PDF service returned an untrusted download URL; no credentials were forwarded.')
    }
  }
  $request.Uri = $download.AbsoluteUri
  $request.Destination = $Destination
  Invoke-PdfRequest @request
}

function Assert-PdfFile {
  param (
    [string]
    $Path
  )

  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw (New-PdfFailure 'InvalidPdf' 'The converter did not create a PDF file.')
  }
  $stream = [IO.File]::OpenRead($Path)
  try {
    $header = [byte[]]::new(5)
    if ($stream.Read($header, 0, 5) -ne 5 -or [Text.Encoding]::ASCII.GetString($header) -ne '%PDF-') {
      throw (New-PdfFailure 'InvalidPdf' 'The converter returned non-PDF content. No PDF was published; inspect the service dashboard.')
    }
  }
  finally {
    $stream.Dispose()
  }
}

try {
  # Validate destinations before ShouldProcess or filesystem mutation.
  $_input = Resolve-PdfPath $InputPath
  if (-not (Test-Path -LiteralPath $_input -PathType Leaf) -or [IO.Path]::GetExtension($_input) -notin @('.html', '.htm')) {
    throw (New-PdfFailure 'InvalidInput' 'InputPath must be an existing .html or .htm file.')
  }
  if (-not $OutputPath) {
    $OutputPath = [IO.Path]::ChangeExtension($_input, '.pdf')
  }
  $_output = Resolve-PdfPath $OutputPath
  $_directory = [IO.Path]::GetDirectoryName($_output)
  $_manifestPath = Resolve-PdfPath (Join-Path $_directory 'pdf.manifest.json')
  if ([IO.Path]::GetExtension($_output) -ne '.pdf' -or $_output -eq $_input) {
    throw (New-PdfFailure 'InvalidOutput' 'OutputPath must name a .pdf file distinct from the input.')
  }
  foreach ($_destination in @($_output, $_manifestPath)) {
    if (Test-Path -LiteralPath $_destination) {
      if (-not (Test-Path -LiteralPath $_destination -PathType Leaf) -or -not $Overwrite) {
        throw (New-PdfFailure 'OutputCollision' 'Output PDF or pdf.manifest.json already exists. Use Overwrite to replace files; directories cannot be replaced.')
      }
    }
  }
  if ((Test-Path -LiteralPath $_directory) -and -not (Test-Path -LiteralPath $_directory -PathType Container)) {
    throw (New-PdfFailure 'InvalidOutput' 'The output parent must be a directory.')
  }
  if ($Mode -eq 'PS') {
    foreach ($_option in @('ApiKey', 'Async', 'Profile', 'Pipeline', 'Test', 'NoNetwork', 'Strict', 'PollInterval', 'Timeout')) {
      if ($PSBoundParameters.ContainsKey($_option)) {
        throw (New-PdfFailure 'InvalidOption' "$_option is only supported by cloud engines.")
      }
    }
  }
  if ($Strict -and $Mode -ne 'DocRaptor') {
    throw (New-PdfFailure 'InvalidOption' 'Strict is only supported by DocRaptor.')
  }
  if ($PSBoundParameters.ContainsKey('PollInterval') -and -not $Async) {
    throw (New-PdfFailure 'InvalidOption' 'PollInterval requires Async.')
  }
  if ($Pipeline -and (($Mode -eq 'EuroPDF' -and $Pipeline -notmatch '^Prince\d+(\.\d+)?(-msfonts)?$') -or
      ($Mode -eq 'DocRaptor' -and $Pipeline -notmatch '^\d+(\.\d+)?$'))) {
    throw (New-PdfFailure 'InvalidOption' 'EuroPDF requires a named Prince pipeline; DocRaptor requires a numeric pipeline.')
  }

  $_inputLock = [IO.File]::Open($_input, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
  $_inputHash = (Get-FileHash -LiteralPath $_input -Algorithm SHA256).Hash
  $_manifest = [ordered]@{
    schema         = 'press-release-pdf/1'
    generatedAtUtc = $null
    input          = [ordered]@{
      path   = $_input
      sha256 = $_inputHash
    }
    converter      = [ordered]@{
      path   = $PSCommandPath
      sha256 = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
      engine = $Mode
    }
    options        = [ordered]@{
      profile   = $Profile
      pipeline  = $Pipeline
      async     = [bool]$Async
      test      = [bool]$Test
      noNetwork = [bool]$NoNetwork
      strict    = [bool]$Strict
    }
    output         = $null
    status         = 'Failed'
    error          = $null
  }

  if ($PSCmdlet.ShouldProcess($_output, "Convert HTML with $Mode and write pdf.manifest.json")) {
    $_phase = 'output preparation'
    $null = New-Item -ItemType Directory -Path $_directory -Force -Confirm:$false
    $_lockPath = Join-Path $_directory '.pdf-conversion.lock'
    $_lock = [IO.FileStream]::new($_lockPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write,
      [IO.FileShare]::None, 4096, [IO.FileOptions]::DeleteOnClose)
    $_stage = Join-Path $_directory ('.pdf-conversion-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $_stage -Confirm:$false
    $_pdfStage = Join-Path $_stage 'output.pdf'
    $_manifestStage = Join-Path $_stage 'pdf.manifest.json'
    $_backup = Join-Path $_stage 'previous.pdf'

    $_phase = 'conversion'
    if ($Mode -eq 'PS') {
      Import-Module PSWritePDF -RequiredVersion 0.0.20 -ErrorAction Stop
      $null = PSWritePDF\Convert-HTMLToPDF -FilePath $_input -OutputFilePath $_pdfStage -WarningAction SilentlyContinue -ErrorAction Stop
    }
    else {
      Invoke-CloudPdfConversion -Html ([IO.File]::ReadAllText($_input)) -Destination $_pdfStage -JobPath (Join-Path $_stage 'job.json')
    }
    Assert-PdfFile $_pdfStage
    $_outputHash = (Get-FileHash -LiteralPath $_pdfStage -Algorithm SHA256).Hash
    $_manifest.output = [ordered]@{
      path   = $_output
      sha256 = $_outputHash
    }
    $_manifest.status = 'Completed'
    $_manifest.generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    [IO.File]::WriteAllText($_manifestStage, (ConvertTo-Json $_manifest -Depth 8), [Text.UTF8Encoding]::new($false))

    # Publish verified bytes; restore the old PDF if manifest publication fails.
    $_phase = 'publication'
    if (Test-Path -LiteralPath $_output) {
      if (-not $Overwrite) {
        throw (New-PdfFailure 'OutputCollision' 'The output PDF appeared during conversion. It was not replaced.')
      }
      [IO.File]::Replace($_pdfStage, $_output, $_backup)
    }
    else {
      [IO.File]::Move($_pdfStage, $_output)
    }
    $_published = $true
    [IO.File]::Move($_manifestStage, $_manifestPath, [bool]$Overwrite)
    $_manifestWritten = $true
    $_status = 'Completed'
  }
  else {
    $_status = 'Skipped'
  }
}
catch {
  $_exception = $_.Exception
  while ($_exception.InnerException -and -not $_exception.Data.Contains('PdfErrorCode')) {
    $_exception = $_exception.InnerException
  }
  $_failure = [ordered]@{
    code    = 'ConversionFailed'
    message = "PDF conversion failed during $_phase. Check paths, permissions, backend installation, and the service dashboard. Raw errors are omitted."
  }
  if ($_exception.Data.Contains('PdfErrorCode')) {
    $_failure.code = [string]$_exception.Data['PdfErrorCode']
    $_failure.message = $_exception.Message
  }
  if ($_published) {
    try {
      if (Test-Path -LiteralPath $_backup) {
        [IO.File]::Move($_backup, $_output, $true)
      }
      else {
        [IO.File]::Delete($_output)
      }
      $_published = $false
    }
    catch {
      $_rollbackFailed = $true
      $_failure.message += ' PDF rollback failed; recovery files remain in the output staging directory.'
    }
  }
  $_outputHash = $null
  if ($_stage -and $_lock -and $_manifest) {
    $_manifest.status = 'Failed'
    $_manifest.output = $null
    $_manifest.error = $_failure
    $_manifest.generatedAtUtc = [DateTime]::UtcNow.ToString('o')
    try {
      [IO.File]::WriteAllText($_manifestStage, (ConvertTo-Json $_manifest -Depth 8), [Text.UTF8Encoding]::new($false))
      [IO.File]::Move($_manifestStage, $_manifestPath, [bool]$Overwrite)
      $_manifestWritten = $true
    }
    catch {
      $_failure.message += ' The failure manifest could not be published.'
    }
  }
}
finally {
  if ($_stage -and -not $_rollbackFailed) {
    try {
      $_resolvedStage = [IO.Path]::GetFullPath($_stage)
      if ([IO.Path]::GetDirectoryName($_resolvedStage) -ne $_directory -or
        [IO.Path]::GetFileName($_resolvedStage) -notmatch '^\.pdf-conversion-[a-f0-9]{32}$') {
        throw 'Unsafe staging cleanup path.'
      }
      Remove-Item -LiteralPath $_resolvedStage -Recurse -Force -Confirm:$false -ErrorAction Stop
    }
    catch {
      Write-Log -Message 'PDF staging cleanup failed; inspect the output directory.' -Color Yellow
    }
  }
  if ($_lock) {
    $_lock.Dispose()
  }
  if ($_inputLock) {
    $_inputLock.Dispose()
  }
}

$_resultParameters = @{
  Target   = $_output
  Source   = 'FileSystem'
  Action   = 'Convert'
  Status   = $_status
  Changed  = ($_status -eq 'Completed' -or $_manifestWritten -or $_published)
  Property = @{
    InputPath       = $_input
    InputHash       = $_inputHash
    OutputPath      = $_output
    OutputHash      = $_outputHash
    ManifestPath    = $_manifestPath
    ManifestWritten = $_manifestWritten
    Engine          = $Mode
  }
}
if ($_failure) {
  $_resultParameters.ErrorMessage = $_failure.message
}
$_result = New-OperationResult @_resultParameters
if ($_failure) {
  $_errorRecord = [Management.Automation.ErrorRecord]::new(
    [InvalidOperationException]::new($_failure.message),
    $_failure.code, [Management.Automation.ErrorCategory]::InvalidOperation, $_result)
  $PSCmdlet.ThrowTerminatingError($_errorRecord)
}
Write-Log -Message "PDF conversion: $_status." -Color Cyan
if ($PassThru) {
  $_result
}
