#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Fixture state must reach mocks inside nested script/module scopes; the previous value is restored in AfterAll.')]
[CmdletBinding()]
param ()

Describe 'Convert-HTMLToPDF' -Skip:($PSVersionTable.PSVersion.Major -lt 7) {
  BeforeAll {
    $script:PreviousFixtureState = Get-Variable WinkitPdfTestState -Scope Global -ErrorAction SilentlyContinue
    $global:WinkitPdfTestState = @{}
    $script:Converter = Join-Path $PSScriptRoot '../../scripts/Files/Convert-HTMLToPDF.ps1'
    $global:WinkitPdfTestState.Utf8 = [Text.UTF8Encoding]::new($false)
    $global:WinkitPdfTestState.PdfBytes = [Text.Encoding]::ASCII.GetBytes("%PDF-1.7`n") + [byte[]]@(0, 128, 255, 13, 10) + [Text.Encoding]::ASCII.GetBytes('%%EOF')
    $script:OldEuroKey = $env:EUROPDF_API_KEY
    $script:OldDocKey = $env:DOCRAPTOR_API_KEY
    $script:FakeModule = New-Module -Name PSWritePDF -Scriptblock {
      function Convert-HTMLToPDF {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Signature-only stub; Pester mocks implement the local backend.')]
        [CmdletBinding()]
        param (
          $FilePath,

          $OutputFilePath
        )

        throw 'Local PDF conversion must be mocked.'
      }
      Export-ModuleMember -Function Convert-HTMLToPDF
    }
    Import-Module $script:FakeModule -Force

    function Invoke-TestPdf {
      param (
        [hashtable]
        $Extra = @{}
      )

      $parameters = $script:Arguments.Clone()
      foreach ($key in $Extra.Keys) {
        $parameters[$key] = $Extra[$key]
      }
      & $script:Converter @parameters -PassThru 6>$null
    }

    function Get-PdfFailure {
      param (
        [hashtable]
        $Extra = @{}
      )

      try {
        $null = Invoke-TestPdf $Extra
      }
      catch {
        return $_
      }
      throw 'Expected PDF conversion to fail.'
    }
  }

  AfterAll {
    $env:EUROPDF_API_KEY = $script:OldEuroKey
    $env:DOCRAPTOR_API_KEY = $script:OldDocKey
    Remove-Module $script:FakeModule -Force
    if ($script:PreviousFixtureState) {
      $global:WinkitPdfTestState = $script:PreviousFixtureState.Value
    }
    else {
      Remove-Variable WinkitPdfTestState -Scope Global
    }
  }

  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Root
    $script:Input = Join-Path $script:Root 'neutral.html'
    $script:OutputDirectory = Join-Path $script:Root 'output'
    $script:Output = Join-Path $script:OutputDirectory 'neutral.pdf'
    $script:ManifestPath = Join-Path $script:OutputDirectory 'pdf.manifest.json'
    [IO.File]::WriteAllText($script:Input, '<html><body>Neutral release: café.</body></html>', $global:WinkitPdfTestState.Utf8)
    $script:Arguments = @{
      InputPath  = $script:Input
      OutputPath = $script:Output
    }
    $env:EUROPDF_API_KEY = 'synthetic-euro-key'
    $env:DOCRAPTOR_API_KEY = 'synthetic-doc-key'
    $global:WinkitPdfTestState.Requests = [Collections.Generic.List[object]]::new()
    $global:WinkitPdfTestState.HttpStatus = 200
    $global:WinkitPdfTestState.AsyncStatus = 'completed'
    $global:WinkitPdfTestState.DownloadUrl = 'https://docraptor.com/download/neutral'
    $global:WinkitPdfTestState.JobCount = 0
    Mock Import-Module { } -ParameterFilter { $Name -eq 'PSWritePDF' }
    Mock 'PSWritePDF\Convert-HTMLToPDF' {
      [IO.File]::WriteAllBytes($OutputFilePath, $global:WinkitPdfTestState.PdfBytes)
    }
    Mock Invoke-WebRequest {
      $json = $null
      if ($Body) {
        $json = [Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json
      }
      $global:WinkitPdfTestState.Requests.Add([pscustomobject]@{
          Uri         = [string]$Uri
          Method      = $Method
          Body        = $json
          Auth        = $Headers.Authorization
          ContentType = $ContentType
          Timeout     = $TimeoutSec
          Redirects   = $MaximumRedirection
        })
      if ($global:WinkitPdfTestState.HttpStatus -ne 200) {
        [IO.File]::WriteAllText($OutFile, 'private-html synthetic-euro-key forbidden-response', $global:WinkitPdfTestState.Utf8)
      }
      elseif ($json.async) {
        [IO.File]::WriteAllText($OutFile, '{"status_id":"neutral-job"}', $global:WinkitPdfTestState.Utf8)
      }
      elseif ([string]$Uri -like '*/status/*') {
        $global:WinkitPdfTestState.JobCount++
        $job = @{
          status        = $global:WinkitPdfTestState.AsyncStatus
          download_id   = 'neutral-download'
          download_url  = $global:WinkitPdfTestState.DownloadUrl
          error_message = 'private-html synthetic-euro-key'
        }
        [IO.File]::WriteAllText($OutFile, (ConvertTo-Json $job), $global:WinkitPdfTestState.Utf8)
      }
      else {
        [IO.File]::WriteAllBytes($OutFile, $global:WinkitPdfTestState.PdfBytes)
      }
      # Reproduces text-decoded PDF Content while OutFile retains the exact bytes.
      [pscustomobject]@{
        StatusCode = $global:WinkitPdfTestState.HttpStatus
        Content    = 'xref trailer %%EOF was not in a correct format'
      }
    }
    Mock Start-Sleep { $global:WinkitPdfTestState.AsyncStatus = 'completed' }
  }

  It 'preserves binary PDF bytes and writes complete successful provenance for <Engine>' -ForEach @(
    @{ Engine = 'EuroPDF'; Key = 'synthetic-euro-key' }
    @{ Engine = 'DocRaptor'; Key = 'synthetic-doc-key' }
  ) {
    $inputHash = (Get-FileHash $script:Input).Hash
    $result = Invoke-TestPdf @{ Mode = $Engine }
    $result.Status | Should -Be 'Completed'
    $result.ManifestWritten | Should -BeTrue
    [IO.File]::ReadAllBytes($script:Output) | Should -Be $global:WinkitPdfTestState.PdfBytes
    $manifest = Get-Content $script:ManifestPath -Raw | ConvertFrom-Json
    $manifest.schema | Should -Be 'press-release-pdf/1'
    Get-Content $script:ManifestPath -Raw | Should -Match '"generatedAtUtc"\s*:\s*"[^"]+Z"'
    { [DateTime]$manifest.generatedAtUtc } | Should -Not -Throw
    $manifest.input.path | Should -Be $script:Input
    $manifest.input.sha256 | Should -Be $inputHash
    $manifest.converter.path | Should -Be ([IO.Path]::GetFullPath($script:Converter))
    $manifest.converter.sha256 | Should -Be (Get-FileHash $script:Converter).Hash
    $manifest.converter.engine | Should -Be $Engine
    $manifest.output.path | Should -Be $script:Output
    $manifest.output.sha256 | Should -Be (Get-FileHash $script:Output).Hash
    $manifest.status | Should -Be 'Completed'
    $manifest.error | Should -BeNullOrEmpty
    [IO.File]::ReadAllBytes($script:ManifestPath)[0] | Should -Be 123
    (Get-FileHash $script:Input).Hash | Should -Be $inputHash
    @(Get-ChildItem $script:OutputDirectory -Force).Count | Should -Be 2
    $global:WinkitPdfTestState.Requests.Count | Should -Be 1
    $global:WinkitPdfTestState.Requests[0].Auth | Should -Be ('Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Key + ':')))
    $global:WinkitPdfTestState.Requests[0].Body.document_content | Should -Match 'café'
    $global:WinkitPdfTestState.Requests[0].Body.prince_options | Should -BeNullOrEmpty
    $global:WinkitPdfTestState.Requests[0].ContentType | Should -Be 'application/json; charset=utf-8'
    $global:WinkitPdfTestState.Requests[0].Redirects | Should -Be 0
  }

  It 'nests profile and network options correctly for <_>' -ForEach @('EuroPDF', 'DocRaptor') {
    $result = Invoke-TestPdf @{ Mode = $_; Profile = 'PDF/A-3a+PDF/UA-1'; NoNetwork = $true; Test = $true }
    $result.Status | Should -Be 'Completed'
    $body = $global:WinkitPdfTestState.Requests[0].Body
    $body.prince_options.profile | Should -Be 'PDF/A-3a+PDF/UA-1'
    $body.prince_options.no_network | Should -BeTrue
    $body.test | Should -BeTrue
    $body.PSObject.Properties.Name | Should -Not -Contain 'profile'
    if ($_ -eq 'DocRaptor') {
      $body.type | Should -Be 'pdf'
    }
    else {
      $body.PSObject.Properties.Name | Should -Not -Contain 'type'
    }
  }

  It 'uses an explicit secure key ahead of engine-specific environment credentials' {
    $key = [Security.SecureString]::new()
    foreach ($character in 'synthetic-explicit-key'.ToCharArray()) {
      $key.AppendChar($character)
    }
    $null = Invoke-TestPdf @{ ApiKey = $key }
    $global:WinkitPdfTestState.Requests[0].Auth | Should -Be ('Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('synthetic-explicit-key:')))
    Get-Content $script:ManifestPath -Raw | Should -Not -Match 'synthetic-.*-key|Authorization'
  }

  It 'never substitutes the other engine key when its own key is missing' {
    $env:EUROPDF_API_KEY = $null
    $failure = Get-PdfFailure
    $failure.FullyQualifiedErrorId | Should -Match 'MissingCredential'
    $failure.TargetObject.ManifestWritten | Should -BeTrue
    $global:WinkitPdfTestState.Requests.Count | Should -Be 0
    (Get-Content $script:ManifestPath -Raw | ConvertFrom-Json).status | Should -Be 'Failed'
  }

  It 'does no network or filesystem mutation for <_>' -ForEach @('DryRun', 'WhatIf') {
    $env:EUROPDF_API_KEY = $null
    $extra = @{}
    $extra[$_] = $true
    $result = Invoke-TestPdf $extra
    $result.Status | Should -Be 'Skipped'
    $result.ManifestWritten | Should -BeFalse
    $result.OutputHash | Should -BeNullOrEmpty
    Test-Path $script:OutputDirectory | Should -BeFalse
    Should -Invoke Invoke-WebRequest -Times 0
  }

  It 'writes sanitized failure manifests for HTTP <_> without leaking response data' -ForEach @(401, 422, 429, 500) {
    $global:WinkitPdfTestState.HttpStatus = $_
    $failure = Get-PdfFailure
    $failure.Exception.Message | Should -Match "HTTP $_"
    $failure.Exception.Message | Should -Not -Match 'private-html|synthetic-euro-key|forbidden-response'
    $failure.Exception.InnerException | Should -BeNullOrEmpty
    $failure.TargetObject.Status | Should -Be 'Failed'
    $manifestText = Get-Content $script:ManifestPath -Raw
    $manifestText | Should -Not -Match 'private-html|synthetic-euro-key|forbidden-response'
    $manifest = $manifestText | ConvertFrom-Json
    $manifest.status | Should -Be 'Failed'
    $manifest.output | Should -BeNullOrEmpty
    $manifest.error.code | Should -Be 'HttpError'
    $manifest.error.message | Should -Be $failure.Exception.Message
    Test-Path $script:Output | Should -BeFalse
    @(Get-ChildItem $script:OutputDirectory -Force).Count | Should -Be 1
    $global:WinkitPdfTestState.Requests.Count | Should -Be 1
  }

  It 'sanitizes transport exceptions and never automatically resubmits a paid request' {
    Mock Invoke-WebRequest { throw 'secret-key private-html xref %%EOF' }
    $failure = Get-PdfFailure
    $failure.Exception.Message | Should -Match 'request failed or timed out'
    $failure.Exception.Message | Should -Not -Match 'secret-key|private-html|xref'
    Get-Content $script:ManifestPath -Raw | Should -Not -Match 'secret-key|private-html|xref'
    Should -Invoke Invoke-WebRequest -Times 1 -Exactly
  }

  It 'rejects non-PDF success responses before publishing a PDF' {
    Mock Invoke-WebRequest {
      [IO.File]::WriteAllText($OutFile, '<html>private response</html>')
      [pscustomobject]@{ StatusCode = 200 }
    }
    $failure = Get-PdfFailure
    $failure.FullyQualifiedErrorId | Should -Match 'InvalidPdf'
    Test-Path $script:Output | Should -BeFalse
    (Get-Content $script:ManifestPath -Raw | ConvertFrom-Json).output | Should -BeNullOrEmpty
  }

  It 'polls and downloads binary async output for <_>' -ForEach @('EuroPDF', 'DocRaptor') {
    $global:WinkitPdfTestState.AsyncStatus = if ($_ -eq 'EuroPDF') { 'started' } else { 'queued' }
    $result = Invoke-TestPdf @{ Mode = $_; Async = $true; PollInterval = 1 }
    $result.Status | Should -Be 'Completed'
    $global:WinkitPdfTestState.JobCount | Should -Be 2
    $global:WinkitPdfTestState.Requests.Count | Should -Be 4
    $global:WinkitPdfTestState.Requests[0].Body.async | Should -BeTrue
    $global:WinkitPdfTestState.Requests[-1].Method | Should -Be 'Get'
    $global:WinkitPdfTestState.Requests[-1].Uri | Should -Match '/download/'
    [IO.File]::ReadAllBytes($script:Output) | Should -Be $global:WinkitPdfTestState.PdfBytes
  }

  It 'records failed and expired async jobs without their raw service errors: <_>' -ForEach @('failed', 'incinerated') {
    $global:WinkitPdfTestState.AsyncStatus = $_
    $failure = Get-PdfFailure @{ Async = $true }
    $failure.FullyQualifiedErrorId | Should -Match 'JobFailed'
    Get-Content $script:ManifestPath -Raw | Should -Not -Match 'private-html|synthetic-euro-key'
    $global:WinkitPdfTestState.Requests.Count | Should -Be 2
  }

  It 'does not send authorization to an untrusted async download host' {
    $global:WinkitPdfTestState.DownloadUrl = 'https://example.invalid/download/neutral'
    $failure = Get-PdfFailure @{ Mode = 'DocRaptor'; Async = $true }
    $failure.Exception.Message | Should -Match 'untrusted download'
    $global:WinkitPdfTestState.Requests.Count | Should -Be 2
  }

  It 'stops async polling when the total timeout expires' {
    $global:WinkitPdfTestState.AsyncStatus = 'started'
    Mock Start-Sleep {
      [Threading.Thread]::Sleep(10100)
    }
    $failure = Get-PdfFailure @{ Async = $true; Timeout = 10; PollInterval = 1 }
    $failure.FullyQualifiedErrorId | Should -Match 'Timeout'
    (Get-Content $script:ManifestPath -Raw | ConvertFrom-Json).error.code | Should -Be 'Timeout'
    $global:WinkitPdfTestState.Requests.Count | Should -Be 2
  }

  It 'uses the local backend with the same manifest contract and no HTTP calls' {
    $result = Invoke-TestPdf @{ Mode = 'PS' }
    $result.Status | Should -Be 'Completed'
    (Get-Content $script:ManifestPath -Raw | ConvertFrom-Json).converter.engine | Should -Be 'PS'
    Should -Invoke 'PSWritePDF\Convert-HTMLToPDF' -Times 1 -Exactly
    Should -Invoke Invoke-WebRequest -Times 0
  }

  It 'records a local converter failure without raw exception contents' {
    Mock 'PSWritePDF\Convert-HTMLToPDF' { throw 'private-html stack trace' }
    $failure = Get-PdfFailure @{ Mode = 'PS' }
    $failure.Exception.Message | Should -Not -Match 'private-html|stack trace'
    (Get-Content $script:ManifestPath -Raw | ConvertFrom-Json).status | Should -Be 'Failed'
  }

  It 'rejects cloud options for PS and invalid engine combinations before writes' -ForEach @(
    @{ Extra = @{ Mode = 'PS'; Async = $true } }
    @{ Extra = @{ Mode = 'PS'; Profile = 'PDF/UA-1' } }
    @{ Extra = @{ Mode = 'PS'; Test = $true } }
    @{ Extra = @{ Mode = 'EuroPDF'; Strict = $true } }
    @{ Extra = @{ Mode = 'EuroPDF'; Pipeline = '10.1' } }
    @{ Extra = @{ Mode = 'DocRaptor'; Pipeline = 'Prince16.2' } }
    @{ Extra = @{ PollInterval = 1 } }
  ) {
    $failure = Get-PdfFailure $Extra
    $failure.FullyQualifiedErrorId | Should -Match 'InvalidOption'
    Test-Path $script:OutputDirectory | Should -BeFalse
    Should -Invoke Invoke-WebRequest -Times 0
  }

  It 'requires Overwrite and preserves old PDF bytes on a failed retry' {
    $first = Invoke-TestPdf
    $hash = (Get-FileHash $script:Output).Hash
    $failure = Get-PdfFailure
    $failure.FullyQualifiedErrorId | Should -Match 'OutputCollision'
    $global:WinkitPdfTestState.Requests.Count | Should -Be 1
    $global:WinkitPdfTestState.HttpStatus = 422
    $failure = Get-PdfFailure @{ Overwrite = $true }
    $failure.TargetObject.ManifestWritten | Should -BeTrue
    (Get-FileHash $script:Output).Hash | Should -Be $hash
    $manifest = Get-Content $first.ManifestPath -Raw | ConvertFrom-Json
    $manifest.status | Should -Be 'Failed'
    $manifest.output | Should -BeNullOrEmpty
  }

  It 'leaves HTML/DOCX build artifacts and their manifest unchanged' {
    $null = New-Item -ItemType Directory $script:OutputDirectory
    $paths = @('release.html', 'release.docx', 'build.manifest.json') | ForEach-Object {
      $path = Join-Path $script:OutputDirectory $_
      [IO.File]::WriteAllText($path, 'synthetic build artifact')
      $path
    }
    $hashes = @($paths | Get-FileHash | Select-Object -ExpandProperty Hash)
    $null = Invoke-TestPdf
    @($paths | Get-FileHash | Select-Object -ExpandProperty Hash) | Should -Be $hashes
  }

  It 'restores the previous PDF if manifest publication fails' {
    $null = Invoke-TestPdf
    $pdfHash = (Get-FileHash $script:Output).Hash
    $manifestHash = (Get-FileHash $script:ManifestPath).Hash
    $lock = [IO.File]::Open($script:ManifestPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
      $failure = Get-PdfFailure @{ Overwrite = $true }
      $failure.TargetObject.Status | Should -Be 'Failed'
      $failure.TargetObject.ManifestWritten | Should -BeFalse
    }
    finally {
      $lock.Dispose()
    }
    (Get-FileHash $script:Output).Hash | Should -Be $pdfHash
    (Get-FileHash $script:ManifestPath).Hash | Should -Be $manifestHash
    @(Get-ChildItem $script:OutputDirectory -Force).Count | Should -Be 2
  }

  It 'does not disturb a competing conversion lock' {
    $null = New-Item -ItemType Directory $script:OutputDirectory
    $lockPath = Join-Path $script:OutputDirectory '.pdf-conversion.lock'
    [IO.File]::WriteAllText($lockPath, 'another conversion')
    $failure = Get-PdfFailure
    $failure.TargetObject.ManifestWritten | Should -BeFalse
    [IO.File]::ReadAllText($lockPath) | Should -Be 'another conversion'
    Should -Invoke Invoke-WebRequest -Times 0
  }
}
