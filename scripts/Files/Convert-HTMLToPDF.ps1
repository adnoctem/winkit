#Requires -Version 7.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Converts an HTML file to PDF using the EuroPDF or DocRaptor cloud APIs, or PSWritePDF locally.

.DESCRIPTION
  Three conversion modes:

  EuroPDF mode (default):
    Sends the HTML content to the EuroPDF cloud API (Prince-based). Supports PDF profiles
    (PDF/UA-1, PDF/A-3a+PDF/UA-1), named Prince pipeline selection, and async processing.

  DocRaptor mode:
    Sends the HTML content to the DocRaptor cloud API (Prince-based, same engine as EuroPDF).
    Supports the same PDF profiles, numeric pipeline selection (9, 10, 10.1), strict HTML
    validation, and async processing with up to 600 seconds of server time.

  PS mode:
    Uses the locally installed PSWritePDF module (iText 7). Works offline with no API key or
    time limits.

  Authentication for cloud APIs uses HTTP Basic Auth (API key as username, empty password).
  Set -ApiKey, or the EUROPDF_API_KEY / DOCRAPTOR_API_KEY environment variable.

.PARAMETER InputPath
  Path to the HTML file to convert. Must be a single .html or .htm file.

.PARAMETER OutputPath
  Path for the output PDF. Defaults to InputPath with .pdf extension.

.PARAMETER Mode
  Conversion engine: 'EuroPDF' (default), 'DocRaptor', or 'PS' (PSWritePDF).

.PARAMETER ApiKey
  API key for cloud services. Can also be set via the EUROPDF_API_KEY or DOCRAPTOR_API_KEY
  environment variables.

.PARAMETER Async
  Use the async endpoint. Recommended for large documents. Async requests get up to
  600 seconds of server time vs 60 for sync. Valid with -Mode EuroPDF and DocRaptor.

.PARAMETER Profile
  PDF profile for the output document. Supported values include:
    PDF/UA-1, PDF/A-1a, PDF/A-2a, PDF/A-3a, PDF/A-3a+PDF/UA-1, PDF/X-4, etc.

.PARAMETER Pipeline
  Prince pipeline. EuroPDF: Prince16.2, Prince16-msfonts, Prince16, Prince15.4.
  DocRaptor: numeric (9, 10, 10.1). Freeform string — see each API's docs.

.PARAMETER Test
  Generate a test document (watermarked, doesn't count against usage quota).

.PARAMETER PollInterval
  Seconds between async status checks (default: 5).

.PARAMETER Timeout
  Maximum wait time in seconds for async conversion (default: 900, max: 3600).

.PARAMETER NoNetwork
  Block external resource fetching during PDF generation (Prince option).

.PARAMETER Strict
  Only valid with -Mode DocRaptor. When set, validates HTML input and fails on markup errors.

.PARAMETER PassThru
  Return the output PDF path to the pipeline.

.EXAMPLE
  PS> Convert-HTMLToPDF.ps1 -InputPath merged/all.html -ApiKey "my-key"

  Converts all.html to all.pdf using EuroPDF (synchronous).

.EXAMPLE
  PS> Convert-HTMLToPDF.ps1 -InputPath merged/all.html -ApiKey "my-key" -Async -Profile "PDF/A-3a+PDF/UA-1" -Mode DocRaptor

  Async conversion with DocRaptor and PDF/UA-1 accessibility profile.

.EXAMPLE
  PS> Convert-HTMLToPDF.ps1 -InputPath page.html -Mode PS

  Converts using the locally installed PSWritePDF module.

.LINK
  https://github.com/adnoctem/winkit
  https://www.europdf.eu/docs/api/
  https://docraptor.com/documentation/api

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true)]
  [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
  [string]
  $InputPath,

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath,

  [Parameter(Mandatory = $false)]
  [ValidateSet('EuroPDF', 'PS', 'DocRaptor')]
  [string]
  $Mode = 'EuroPDF',

  [Parameter(Mandatory = $false)]
  [string]
  $ApiKey,

  [Parameter(Mandatory = $false)]
  [switch]
  $Async,

  [Parameter(Mandatory = $false)]
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

  [Parameter(Mandatory = $false)]
  [string]
  $Pipeline,

  [Parameter(Mandatory = $false)]
  [switch]
  $Test,

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 300)]
  [int]
  $PollInterval = 5,

  [Parameter(Mandatory = $false)]
  [ValidateRange(10, 3600)]
  [int]
  $Timeout = 900,

  [Parameter(Mandatory = $false)]
  [switch]
  $NoNetwork,

  [Parameter(Mandatory = $false)]
  [switch]
  $Strict,

  [Parameter(Mandatory = $false)]
  [switch]
  $PassThru
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------
# Script-wide state
# -----------------------------------------------------------------------------

$_results = New-Object System.Collections.ArrayList

$_euroPdfBase = 'https://api.europdf.eu/v1'
$_docRaptorBase = 'https://api.docraptor.com'

# -----------------------------------------------------------------------------
# Helper functions
# -----------------------------------------------------------------------------

function Resolve-ApiKey {
  param ([string] $Key)

  if ($Key) { return $Key }
  $_envKey = $env:DOCRAPTOR_API_KEY
  if ($_envKey) { return $_envKey }
  $_envKey = $env:EUROPDF_API_KEY
  if ($_envKey) { return $_envKey }
  throw "No API key provided. Set -ApiKey or the DOCRAPTOR_API_KEY / EUROPDF_API_KEY environment variable."
}

function ConvertTo-BasicAuth {
  param ([string] $ApiKey)

  $_bytes = [System.Text.Encoding]::UTF8.GetBytes("${ApiKey}:")
  $_base64 = [Convert]::ToBase64String($_bytes)
  return "Basic $_base64"
}

function Get-DocRaptorError {
  param ([string] $RawBody)

  if ($RawBody -match '<error>([^<]*)</error>') {
    return $matches[1]
  }
  return $RawBody
}

function Invoke-EuroPdfSync {
  param (
    [string] $BaseUrl,
    [string] $ApiKey,
    [string] $HtmlContent,
    [string] $PdfProfile,
    [string] $Pipeline,
    [bool] $Test,
    [bool] $NoNetwork
  )

  $_body = @{ document_content = $HtmlContent }
  if ($PdfProfile) { $_body.profile = $PdfProfile }
  if ($Pipeline) { $_body.pipeline = $Pipeline }
  if ($Test) { $_body.test = $true }
  if ($NoNetwork) { $_body.prince_options = @{ no_network = $true } }

  $_authHeader = ConvertTo-BasicAuth -ApiKey $ApiKey
  $_json = $_body | ConvertTo-Json -Depth 5 -Compress

  try {
    $_response = Invoke-WebRequest -Uri "$BaseUrl/docs" `
      -Method Post `
      -Headers @{ Authorization = $_authHeader } `
      -ContentType 'application/json' `
      -Body $_json `
      -SkipHttpErrorCheck `
      -ErrorAction Stop

    if ($_response.StatusCode -ne 200) {
      $_apiError = try { $_response.Content } catch { '(could not read response)' }
      throw "EuroPDF returned HTTP $($_response.StatusCode): $_apiError"
    }

    return $_response.Content
  }
  catch {
    throw "EuroPDF sync request failed: $($_.Exception.Message)"
  }
}

function Invoke-EuroPdfAsync {
  param (
    [string] $BaseUrl,
    [string] $ApiKey,
    [string] $HtmlContent,
    [string] $PdfProfile,
    [string] $Pipeline,
    [bool] $Test,
    [bool] $NoNetwork,
    [int] $PollInterval,
    [int] $Timeout
  )

  $_body = @{
    document_content = $HtmlContent
    async = $true
  }
  if ($PdfProfile) { $_body.profile = $PdfProfile }
  if ($Pipeline) { $_body.pipeline = $Pipeline }
  if ($Test) { $_body.test = $true }
  if ($NoNetwork) { $_body.prince_options = @{ no_network = $true } }

  $_authHeader = ConvertTo-BasicAuth -ApiKey $ApiKey
  $_json = $_body | ConvertTo-Json -Depth 5 -Compress

  Write-Log -Message '  Submitting async conversion job ...' -Color Cyan
  try {
    $_response = Invoke-WebRequest -Uri "$BaseUrl/docs" `
      -Method Post `
      -Headers @{ Authorization = $_authHeader } `
      -ContentType 'application/json' `
      -Body $_json `
      -SkipHttpErrorCheck `
      -ErrorAction Stop

    if ($_response.StatusCode -ne 200) {
      $_apiError = try { $_response.Content } catch { '(could not read response)' }
      throw "EuroPDF returned HTTP $($_response.StatusCode): $_apiError"
    }

    $_statusResponse = $_response.Content | ConvertFrom-Json
    $_statusId = $_statusResponse.status_id
    if (-not $_statusId) {
      throw "Async response did not contain a status_id: $($_response.Content)"
    }

    Write-Log -Message "  Job submitted: $_statusId" -Color Green
  }
  catch {
    throw "Failed to submit async job: $($_.Exception.Message)"
  }

  $_deadline = (Get-Date).AddSeconds($Timeout)

  while ((Get-Date) -lt $_deadline) {
    Start-Sleep -Seconds $PollInterval

    try {
      $_statusResponse = Invoke-WebRequest -Uri "$BaseUrl/status/$_statusId" `
        -Headers @{ Authorization = $_authHeader } `
        -ErrorAction Stop |
        Select-Object -ExpandProperty Content |
        ConvertFrom-Json

      switch ($_statusResponse.status) {
        'completed' {
          Write-Log -Message "  Status: completed." -Color Green
          Write-Log -Message '  Downloading ...' -Color Cyan

          $_downloadId = $_statusResponse.download_id
          if (-not $_downloadId) {
            $_downloadId = $_statusResponse.download_url -replace '.*/([^/]+)$', '$1'
          }

          $_pdfResponse = Invoke-WebRequest -Uri "$BaseUrl/download/$_downloadId" `
            -Headers @{ Authorization = $_authHeader } `
            -ErrorAction Stop

          return $_pdfResponse.Content
        }
        'failed' {
          throw "EuroPDF conversion failed: $($_statusResponse.error_message)"
        }
        'incinerated' {
          throw 'Document was deleted (incinerated) before download could complete.'
        }
        default {
          Write-Log -Message "  Status: $($_statusResponse.status) - waiting ..." -Color Gray
        }
      }
    }
    catch {
      if ($_.Exception.Message -match 'EuroPDF conversion failed|Document was deleted|Failed to submit') {
        throw
      }
      Write-Log -Message "  Status check error: $($_.Exception.Message) - retrying ..." -Color DarkYellow
    }
  }

  throw "Async conversion timed out after $Timeout seconds."
}

function Invoke-DocRaptorSync {
  param (
    [string] $BaseUrl,
    [string] $ApiKey,
    [string] $HtmlContent,
    [string] $PdfProfile,
    [string] $Pipeline,
    [bool] $Test,
    [bool] $NoNetwork,
    [bool] $Strict
  )

  $_body = @{
    type = 'pdf'
    document_content = $HtmlContent
  }

  if ($PdfProfile) { $_body.prince_options = @{ profile = $PdfProfile } }
  if ($Pipeline) { $_body.pipeline = $Pipeline }
  if ($Test) { $_body.test = $true }
  if ($NoNetwork) {
    if (-not $_body.prince_options) { $_body.prince_options = @{} }
    $_body.prince_options.no_network = $true
  }
  if ($Strict) { $_body.strict = 'html' }

  $_authHeader = ConvertTo-BasicAuth -ApiKey $ApiKey
  $_json = $_body | ConvertTo-Json -Depth 5 -Compress

  try {
    $_response = Invoke-WebRequest -Uri "$BaseUrl/docs" `
      -Method Post `
      -Headers @{ Authorization = $_authHeader } `
      -ContentType 'application/json' `
      -Body $_json `
      -SkipHttpErrorCheck `
      -ErrorAction Stop

    if ($_response.StatusCode -ne 200) {
      $_apiError = try { Get-DocRaptorError -RawBody $_response.Content } catch { '(could not read response)' }
      throw "DocRaptor returned HTTP $($_response.StatusCode): $_apiError"
    }

    return $_response.Content
  }
  catch {
    throw "DocRaptor sync request failed: $($_.Exception.Message)"
  }
}

function Invoke-DocRaptorAsync {
  param (
    [string] $BaseUrl,
    [string] $ApiKey,
    [string] $HtmlContent,
    [string] $PdfProfile,
    [string] $Pipeline,
    [bool] $Test,
    [bool] $NoNetwork,
    [bool] $Strict,
    [int] $PollInterval,
    [int] $Timeout
  )

  $_body = @{
    type = 'pdf'
    document_content = $HtmlContent
    async = $true
  }

  if ($PdfProfile) { $_body.prince_options = @{ profile = $PdfProfile } }
  if ($Pipeline) { $_body.pipeline = $Pipeline }
  if ($Test) { $_body.test = $true }
  if ($NoNetwork) {
    if (-not $_body.prince_options) { $_body.prince_options = @{} }
    $_body.prince_options.no_network = $true
  }
  if ($Strict) { $_body.strict = 'html' }

  $_authHeader = ConvertTo-BasicAuth -ApiKey $ApiKey
  $_json = $_body | ConvertTo-Json -Depth 5 -Compress

  Write-Log -Message '  Submitting async conversion job ...' -Color Cyan
  try {
    $_response = Invoke-WebRequest -Uri "$BaseUrl/docs" `
      -Method Post `
      -Headers @{ Authorization = $_authHeader } `
      -ContentType 'application/json' `
      -Body $_json `
      -SkipHttpErrorCheck `
      -ErrorAction Stop

    if ($_response.StatusCode -ne 200) {
      $_apiError = try { Get-DocRaptorError -RawBody $_response.Content } catch { '(could not read response)' }
      throw "DocRaptor returned HTTP $($_response.StatusCode): $_apiError"
    }

    $_statusResponse = $_response.Content | ConvertFrom-Json
    $_statusId = $_statusResponse.status_id
    if (-not $_statusId) {
      throw "Async response did not contain a status_id: $($_response.Content)"
    }

    Write-Log -Message "  Job submitted: $_statusId" -Color Green
  }
  catch {
    throw "Failed to submit async job: $($_.Exception.Message)"
  }

  $_deadline = (Get-Date).AddSeconds($Timeout)

  while ((Get-Date) -lt $_deadline) {
    Start-Sleep -Seconds $PollInterval

    try {
      $_statusResponse = Invoke-WebRequest -Uri "$BaseUrl/status/$_statusId" `
        -Headers @{ Authorization = $_authHeader } `
        -ErrorAction Stop |
        Select-Object -ExpandProperty Content |
        ConvertFrom-Json

      switch ($_statusResponse.status) {
        'completed' {
          Write-Log -Message "  Status: completed." -Color Green
          Write-Log -Message '  Downloading ...' -Color Cyan

          $_downloadUrl = $_statusResponse.download_url
          if (-not $_downloadUrl) {
            throw 'No download URL in completed status response.'
          }

          $_pdfResponse = Invoke-WebRequest -Uri $_downloadUrl `
            -Headers @{ Authorization = $_authHeader } `
            -ErrorAction Stop

          return $_pdfResponse.Content
        }
        'failed' {
          $_errorMsg = if ($_statusResponse.validation_errors) {
            $_statusResponse.validation_errors
          }
          else {
            $_statusResponse.message
          }
          throw "DocRaptor conversion failed: $_errorMsg"
        }
        default {
          Write-Log -Message "  Status: $($_statusResponse.status) - waiting ..." -Color Gray
        }
      }
    }
    catch {
      if ($_.Exception.Message -match 'DocRaptor conversion failed|No download URL|Failed to submit') {
        throw
      }
      Write-Log -Message "  Status check error: $($_.Exception.Message) - retrying ..." -Color DarkYellow
    }
  }

  throw "Async conversion timed out after $Timeout seconds."
}

function Invoke-PsWritePdfConversion {
  param ([string] $InputPath, [string] $OutputPath)

  $_module = Get-Module -ListAvailable -Name PSWritePDF -ErrorAction SilentlyContinue
  if (-not $_module) {
    throw "PSWritePDF module is not installed. Install it with: Install-Module PSWritePDF -RequiredVersion 0.0.20"
  }

  Import-Module PSWritePDF -Force -ErrorAction Stop

  Write-Log -Message '  Converting with PSWritePDF (Convert-HTMLToPDF) ...' -Color Cyan

  Convert-HTMLToPDF -FilePath $InputPath -OutputFilePath $OutputPath -ErrorAction Stop

  if (Test-Path -LiteralPath $OutputPath) {
    return $OutputPath
  }

  throw 'PSWritePDF completed but no output file was created.'
}

# -----------------------------------------------------------------------------
# Validation
# -----------------------------------------------------------------------------

if ($OutputPath) {
  $_outputPath = $OutputPath
}
else {
  $_outputPath = [System.IO.Path]::ChangeExtension($InputPath, '.pdf')
}

if (-not (Test-Path -LiteralPath ([System.IO.Path]::GetDirectoryName($_outputPath)))) {
  New-Item -Path ([System.IO.Path]::GetDirectoryName($_outputPath)) -ItemType Directory -Force | Out-Null
}

if ($Async -and $Mode -eq 'PS') {
  throw "-Async is only valid with -Mode EuroPDF or DocRaptor."
}

if ($Strict -and $Mode -ne 'DocRaptor') {
  throw "-Strict is only valid with -Mode DocRaptor."
}

if ($WhatIfPreference) {
  $WhatIfPreference = $false
}

# -----------------------------------------------------------------------------
# Conversion
# -----------------------------------------------------------------------------

Write-Log -Message "===== Convert HTML to PDF =====" -Color Cyan
Write-Log -Message "Input:    $InputPath" -Color White
Write-Log -Message "Output:   $_outputPath" -Color White
Write-Log -Message "Mode:     $Mode" -Color White

if ($Mode -eq 'EuroPDF') {
  $_apiKey = Resolve-ApiKey -Key $ApiKey
  Write-Log -Message "API mode: $(if ($Async) { 'async' } else { 'sync' })" -Color White
  if ($Async) {
    Write-Log -Message "Timeout:  $Timeout s (poll every $PollInterval s)" -Color White
  }
  if ($Profile) { Write-Log -Message "Profile:  $Profile" -Color White }
  if ($Pipeline) { Write-Log -Message "Pipeline: $Pipeline" -Color White }

  Write-Log -Message 'Reading HTML content ...' -Color Cyan
  $_htmlContent = [System.IO.File]::ReadAllText($InputPath)
  Write-Log -Message "  Read $($_htmlContent.Length) characters." -Color Gray

  if ($Async) {
    $_pdfBytes = Invoke-EuroPdfAsync -BaseUrl $_euroPdfBase -ApiKey $_apiKey `
      -HtmlContent $_htmlContent -PdfProfile $Profile -Pipeline $Pipeline `
      -Test:$Test.IsPresent -NoNetwork:$NoNetwork.IsPresent `
      -PollInterval $PollInterval -Timeout $Timeout
  }
  else {
    $_pdfBytes = Invoke-EuroPdfSync -BaseUrl $_euroPdfBase -ApiKey $_apiKey `
      -HtmlContent $_htmlContent -PdfProfile $Profile -Pipeline $Pipeline `
      -Test:$Test.IsPresent -NoNetwork:$NoNetwork.IsPresent
  }

  [System.IO.File]::WriteAllBytes($_outputPath, $_pdfBytes)
}
elseif ($Mode -eq 'DocRaptor') {
  $_apiKey = Resolve-ApiKey -Key $ApiKey
  Write-Log -Message "API mode: $(if ($Async) { 'async' } else { 'sync' })" -Color White
  if ($Async) {
    Write-Log -Message "Timeout:  $Timeout s (poll every $PollInterval s)" -Color White
  }
  if ($Profile) { Write-Log -Message "Profile:  $Profile" -Color White }
  if ($Pipeline) { Write-Log -Message "Pipeline: $Pipeline" -Color White }
  if ($Strict) { Write-Log -Message "Strict:   enabled" -Color White }

  Write-Log -Message 'Reading HTML content ...' -Color Cyan
  $_htmlContent = [System.IO.File]::ReadAllText($InputPath)
  Write-Log -Message "  Read $($_htmlContent.Length) characters." -Color Gray

  if ($Async) {
    $_pdfBytes = Invoke-DocRaptorAsync -BaseUrl $_docRaptorBase -ApiKey $_apiKey `
      -HtmlContent $_htmlContent -PdfProfile $Profile -Pipeline $Pipeline `
      -Test:$Test.IsPresent -NoNetwork:$NoNetwork.IsPresent -Strict:$Strict.IsPresent `
      -PollInterval $PollInterval -Timeout $Timeout
  }
  else {
    $_pdfBytes = Invoke-DocRaptorSync -BaseUrl $_docRaptorBase -ApiKey $_apiKey `
      -HtmlContent $_htmlContent -PdfProfile $Profile -Pipeline $Pipeline `
      -Test:$Test.IsPresent -NoNetwork:$NoNetwork.IsPresent -Strict:$Strict.IsPresent
  }

  [System.IO.File]::WriteAllBytes($_outputPath, $_pdfBytes)
}
elseif ($Mode -eq 'PS') {
  Invoke-PsWritePdfConversion -InputPath $InputPath -OutputPath $_outputPath
}

# -----------------------------------------------------------------------------
# Verification
# -----------------------------------------------------------------------------

if (Test-Path -LiteralPath $_outputPath) {
  $_fileInfo = Get-Item -LiteralPath $_outputPath
  $_fileSize = $_fileInfo.Length
  Write-Log -Message "  Wrote: $_outputPath ($([Math]::Round($_fileSize / 1MB, 1)) MB)" -Color Green

  Add-OperationResult -Results $_results -Target $InputPath `
    -Source 'FileSystem' -Action 'Convert' -Status 'Completed' `
    -Detail "PDF: $_outputPath ($([Math]::Round($_fileSize / 1MB, 1)) MB)"
}
else {
  Write-Log -Message '  ERROR: Output file was not created.' -Color Red
  Add-OperationResult -Results $_results -Target $InputPath `
    -Source 'FileSystem' -Action 'Convert' -Status 'Failed' `
    -Detail 'Output file not created'
}

# -----------------------------------------------------------------------------
# Report
# -----------------------------------------------------------------------------

$_completed = @($_results | Where-Object { $_.Status -eq 'Completed' }).Count
$_failed = @($_results | Where-Object { $_.Status -eq 'Failed' }).Count

Write-Log -Message ("`n===== Conversion Summary =====") -Color Cyan
Write-Log -Message ("Completed: $_completed | Failed: $_failed") `
  -Color $(if ($_failed -gt 0) { 'Red' } else { 'Green' })

if ($PassThru) {
  $_outputPath
}
