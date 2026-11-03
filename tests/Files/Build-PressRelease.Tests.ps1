#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# The document API ships only for Core. Keep discovery compatible with existing 5.1 suites.
Describe 'Build-PressRelease' -Skip:($PSVersionTable.PSVersion -lt [version]'7.4') {
  BeforeAll {
    Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force
    $script:BuildScript = Join-Path $PSScriptRoot '../../scripts/Files/Build-PressRelease.ps1'
    $script:Utf8 = [Text.UTF8Encoding]::new($false)

    function New-SyntheticDocument {
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates disposable synthetic fixtures only inside Pester TestDrive.')]
      [CmdletBinding()]
      param (
        [string]
        $Path,

        [switch]
        $AsTemplate,

        [switch]
        $Invalid
      )

      $headline = 'Neutral release headline'
      if ($Invalid) {
        $headline = $null
      }
      $body = [AdNoctem.Substrate.Documents.DocumentBlock[]]@(
        [AdNoctem.Substrate.Documents.ParagraphBlock]::new(
          [AdNoctem.Substrate.Documents.DocumentInline[]]@(
            [AdNoctem.Substrate.Documents.TextInline]::new('Neutral release body.')
          )
        )
      )
      $fields = [AdNoctem.Substrate.Documents.AuthoringElement[]]@(
        [AdNoctem.Substrate.Documents.AuthoringField]::Text('pr.headline', 'Headline', 'Enter headline', $headline)
        [AdNoctem.Substrate.Documents.AuthoringField]::Text('pr.dateline', 'Dateline', $null, 'Exampletown, November 2026')
        [AdNoctem.Substrate.Documents.AuthoringField]::Rich('pr.body', 'Body', $null, $body)
        [AdNoctem.Substrate.Documents.AuthoringField]::Group('pr.contact', 'Contact', $true,
          [AdNoctem.Substrate.Documents.AuthoringElement[]]@(
            [AdNoctem.Substrate.Documents.AuthoringField]::Text('pr.contact.heading', 'Heading', $null, 'Contact')
            [AdNoctem.Substrate.Documents.AuthoringField]::Rich('pr.contact.details', 'Details', $null, $body)
          ))
      )
      $properties = [Collections.Generic.Dictionary[string, string]]::new()
      $properties.Add('PressRelease.Schema', 'press-release/1')
      $properties.Add('PressRelease.Language', 'en')
      $properties.Add('PressRelease.TemplateVersion', '1.0')
      $template = [AdNoctem.Substrate.Documents.AuthoringTemplate]::new($fields, $properties, 'en-GB')
      $stream = [IO.File]::Create($Path)
      try {
        [AdNoctem.Substrate.Documents.AuthoringTemplateBuilder]::new().Write($template, $stream, [bool]$AsTemplate)
      }
      finally {
        $stream.Dispose()
      }
    }

    function Edit-SyntheticDocument {
      param (
        [string]
        $Path,

        [scriptblock]
        $Edit
      )

      $archive = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
      try {
        $entry = $archive.GetEntry('word/document.xml')
        $reader = [IO.StreamReader]::new($entry.Open())
        $xml = [xml]$reader.ReadToEnd()
        $reader.Dispose()
        $namespace = [Xml.XmlNamespaceManager]::new($xml.NameTable)
        $namespace.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        & $Edit $xml $namespace
        $entry.Delete()
        $writer = [IO.StreamWriter]::new($archive.CreateEntry('word/document.xml').Open(), $script:Utf8)
        $writer.Write($xml.OuterXml)
        $writer.Dispose()
      }
      finally {
        $archive.Dispose()
      }
    }

    function Invoke-ReleaseBuild {
      param (
        [hashtable]
        $Extra = @{}
      )

      $parameters = $script:Arguments.Clone()
      foreach ($key in $Extra.Keys) {
        $parameters[$key] = $Extra[$key]
      }
      & $script:BuildScript @parameters -PassThru 6>$null
    }
  }

  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $script:Pack = Join-Path $script:Root 'pack'
    $script:Output = Join-Path $script:Root 'Output'
    $null = New-Item -ItemType Directory -Path $script:Pack -Force
    $script:Input = Join-Path $script:Root 'release.docx'
    New-SyntheticDocument -Path $script:Input
    New-SyntheticDocument -Path (Join-Path $script:Pack 'base.docx')
    New-SyntheticDocument -Path (Join-Path $script:Pack 'authoring.dotx') -AsTemplate
    $html = @'
<!doctype html>
<html lang="{{language}}"><head><meta charset="UTF-8"><title>{{title}}</title></head>
<body><header>{{masthead}}</header><h1>{{headline}}</h1>{{body}}
<footer><h2>{{contact.heading}}</h2>{{contact.details}}</footer></body></html>
'@
    [IO.File]::WriteAllText((Join-Path $script:Pack 'release.html'), $html, $script:Utf8)
    $assets = @{}
    foreach ($name in @('base.docx', 'authoring.dotx', 'release.html')) {
      $assets[$name] = @{
        path   = $name
        sha256 = (Get-FileHash -LiteralPath (Join-Path $script:Pack $name) -Algorithm SHA256).Hash
      }
    }
    $manifest = @{
      packVersion         = '1.0'
      schema              = 'press-release/1'
      schemaVersions      = @('1.0')
      outputLayoutVersion = '1.0'
      languages           = @(
        @{
          code                   = 'en'
          masthead               = 'Release'
          datelineSeparator      = ' - '
          specificationSeparator = ' '
          html                   = $assets['release.html']
          docx                   = $assets['base.docx']
          authoring              = $assets['authoring.dotx']
          sections               = @('contact')
        }
      )
    }
    [IO.File]::WriteAllText((Join-Path $script:Pack 'pack.json'), (ConvertTo-Json $manifest -Depth 8), $script:Utf8)
    $script:Arguments = @{
      InputDocument   = $script:Input
      PackPath        = $script:Pack
      OutputDirectory = $script:Output
    }
  }

  It 'publishes HTML and DOCX with matching manifest hashes and metadata, preserving the source' {
    $before = (Get-FileHash -LiteralPath $script:Input).Hash
    $result = Invoke-ReleaseBuild
    $result.Status | Should -Be 'Completed'
    $result.Language | Should -Be 'en'
    $result.PackVersion | Should -Be '1.0'
    $result.HtmlHash | Should -Be (Get-FileHash -LiteralPath $result.HtmlPath).Hash
    $result.DocxHash | Should -Be (Get-FileHash -LiteralPath $result.DocxPath).Hash
    (Get-FileHash -LiteralPath $script:Input).Hash | Should -Be $before
    $manifest = Get-Content -LiteralPath $result.ManifestPath -Raw | ConvertFrom-Json
    $manifest.schema | Should -Be 'press-release-build/1'
    $manifest.status | Should -Be 'Completed'
    $manifest.tool.version | Should -Not -BeNullOrEmpty
    $manifest.tool.substrateVersion | Should -Be (Get-Module AdNoctem.Substrate.PowerShell).Version.ToString()
    $manifest.input.sha256 | Should -Be $before
    $manifest.input.schema | Should -Be 'press-release/1'
    $manifest.input.templateVersion | Should -Be '1.0'
    $manifest.pack.sha256 | Should -Be (Get-FileHash -LiteralPath (Join-Path $script:Pack 'pack.json')).Hash
    $manifest.outputs.html.sha256 | Should -Be $result.HtmlHash
    $manifest.outputs.docx.sha256 | Should -Be $result.DocxHash
    [IO.File]::ReadAllBytes($result.ManifestPath)[0] | Should -Be 123
    @(Get-ChildItem -LiteralPath $script:Output -Force).Count | Should -Be 3
    Get-Content -LiteralPath $result.HtmlPath -Raw | Should -Match 'Neutral release headline'
  }

  It 'turns a plain HTTPS URL in the source body into an HTML hyperlink' {
    Edit-SyntheticDocument -Path $script:Input -Edit {
      param (
        $xml,

        $namespace
      )

      $text = $xml.SelectSingleNode('//w:sdt[w:sdtPr/w:tag/@w:val="pr.body"]//w:t', $namespace)
      $text.InnerText = 'Media: https://example.com/neutral-release'
    }

    $result = Invoke-ReleaseBuild
    Get-Content -LiteralPath $result.HtmlPath -Raw | Should -Match '<a href="https://example\.com/neutral-release">'
  }

  It 'rejects an authoring .dotx path' {
    { Invoke-ReleaseBuild @{ InputDocument = (Join-Path $script:Pack 'authoring.dotx') } } | Should -Throw '*saved .docx*'
    Test-Path -LiteralPath $script:Output | Should -BeFalse
  }

  It 'rejects a template renamed to .docx' {
    Copy-Item -LiteralPath (Join-Path $script:Pack 'authoring.dotx') -Destination $script:Input -Force
    { Invoke-ReleaseBuild 2>$null } | Should -Throw '*validation failed*'
    Test-Path -LiteralPath $script:Output | Should -BeFalse
  }

  It 'rejects missing input or pack paths' -ForEach @('InputDocument', 'PackPath') {
    $extra = @{}
    $extra[$_] = Join-Path $script:Root 'missing.docx'
    { Invoke-ReleaseBuild $extra } | Should -Throw
    Test-Path -LiteralPath $script:Output | Should -BeFalse
  }

  It 'reports the DocumentException reason for invalid packs' {
    [IO.File]::AppendAllText((Join-Path $script:Pack 'release.html'), 'changed')
    { Invoke-ReleaseBuild } | Should -Throw '*TemplateAssetHashMismatch*'
    Test-Path -LiteralPath $script:Output | Should -BeFalse
  }

  It 'rejects malformed input' {
    [IO.File]::WriteAllText($script:Input, 'not a DOCX')
    { Invoke-ReleaseBuild 2>$null } | Should -Throw '*validation failed*'
    Test-Path -LiteralPath $script:Output | Should -BeFalse
  }

  It 'rejects unsafe output names: <_>' -ForEach @('../escape', 'a/b', 'a\b', 'bad:name', 'NUL', 'trailing.', '') {
    { Invoke-ReleaseBuild @{ OutputName = $_ } } | Should -Throw '*safe file basename*'
    Test-Path -LiteralPath $script:Output | Should -BeFalse
  }

  It 'never replaces the input or writes inside the pack even with Overwrite' {
    { Invoke-ReleaseBuild @{ OutputDirectory = $script:Root; Overwrite = $true } } | Should -Throw '*never be the input*'
    { Invoke-ReleaseBuild @{ OutputDirectory = $script:Pack; Overwrite = $true } } | Should -Throw '*outside the template pack*'
  }

  It 'requires Overwrite for collisions and supports explicit output names' {
    $first = Invoke-ReleaseBuild @{ OutputName = 'published' }
    { Invoke-ReleaseBuild @{ OutputName = 'published' } } | Should -Throw '*Overwrite*'
    $second = Invoke-ReleaseBuild @{ OutputName = 'published'; Overwrite = $true }
    $second.Status | Should -Be 'Completed'
    $second.HtmlPath | Should -Be $first.HtmlPath
    $second.DocxPath | Should -BeLike '*published.docx'
  }

  It 'treats the shared manifest as a collision for another output name' {
    $null = Invoke-ReleaseBuild
    { Invoke-ReleaseBuild @{ OutputName = 'another' } } | Should -Throw '*build.manifest.json*'
  }

  It 'validates without creating a directory for <_>' -ForEach @('DryRun', 'WhatIf') {
    $extra = @{}
    $extra[$_] = $true
    $result = Invoke-ReleaseBuild $extra
    $result.Status | Should -Be 'Skipped'
    $result.HtmlHash | Should -BeNullOrEmpty
    Test-Path -LiteralPath $script:Output | Should -BeFalse
  }

  It 'preserves every prior artifact on validation failure and exposes all diagnostic fields' {
    $first = Invoke-ReleaseBuild
    $hashes = @(Get-ChildItem -LiteralPath $script:Output | Get-FileHash | Select-Object -ExpandProperty Hash)
    New-SyntheticDocument -Path $script:Input -Invalid
    try {
      $null = Invoke-ReleaseBuild @{ Overwrite = $true } 2>$null
      throw 'Expected a terminating validation failure'
    }
    catch {
      $_.FullyQualifiedErrorId | Should -Match 'PressReleaseBuildFailed'
      $result = $_.TargetObject
      $result.Status | Should -Be 'Failed'
      @($result.Diagnostics | Where-Object Severity -EQ 'Error').Count | Should -BeGreaterThan 0
      foreach ($diagnostic in $result.Diagnostics) {
        $diagnostic.Code | Should -Not -BeNullOrEmpty
        $diagnostic.PSObject.Properties.Name | Should -Contain 'Field'
        $diagnostic.Message | Should -Not -BeNullOrEmpty
        $diagnostic.Remedy | Should -Not -BeNullOrEmpty
      }
    }
    @(Get-ChildItem -LiteralPath $script:Output | Get-FileHash | Select-Object -ExpandProperty Hash) | Should -Be $hashes
    Test-Path -LiteralPath $first.ManifestPath | Should -BeTrue
  }

  It 'continues with formatting warnings and records them in the result and manifest' {
    Edit-SyntheticDocument -Path $script:Input -Edit {
      param (
        $xml,

        $namespace
      )
      $run = $xml.SelectSingleNode('//w:sdt[w:sdtPr/w:tag/@w:val="pr.body"]//w:r', $namespace)
      $pageBreak = $xml.CreateElement('w', 'br', $namespace.LookupNamespace('w'))
      $pageBreak.SetAttribute('type', $namespace.LookupNamespace('w'), 'page')
      $null = $run.AppendChild($pageBreak)
    }
    $result = Invoke-ReleaseBuild
    $result.Status | Should -Be 'Completed'
    @($result.Diagnostics | Where-Object Severity -EQ 'Warning').Count | Should -BeGreaterThan 0
    $manifest = Get-Content -LiteralPath $result.ManifestPath -Raw | ConvertFrom-Json
    $manifest.diagnostics.Code | Should -Contain 'IgnoredFormatting'
  }

  It 'keeps prior outputs when staging fails before publication' {
    $first = Invoke-ReleaseBuild
    $hashes = @(Get-ChildItem -LiteralPath $script:Output | Get-FileHash | Select-Object -ExpandProperty Hash)
    Mock Get-FileHash {
      throw 'Synthetic post-generation hash failure'
    } -ParameterFilter { $LiteralPath -like '*\.press-release-*\*.html' }
    { Invoke-ReleaseBuild @{ Overwrite = $true } } | Should -Throw '*Synthetic post-generation*'
    @(Get-ChildItem -LiteralPath $script:Output | Get-FileHash | Select-Object -ExpandProperty Hash) | Should -Be $hashes
    Test-Path -LiteralPath $first.ManifestPath | Should -BeTrue
  }

  It 'rolls back published HTML when replacing a locked DOCX fails' {
    $first = Invoke-ReleaseBuild
    $hashes = @(Get-ChildItem -LiteralPath $script:Output | Get-FileHash | Select-Object -ExpandProperty Hash)
    $lock = [IO.File]::Open($first.DocxPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
      { Invoke-ReleaseBuild @{ Overwrite = $true } } | Should -Throw '*after publication began*Unrestored output count: 0*'
    }
    finally {
      $lock.Dispose()
    }
    @(Get-ChildItem -LiteralPath $script:Output -Force).Count | Should -Be 3
    @(Get-ChildItem -LiteralPath $script:Output | Get-FileHash | Select-Object -ExpandProperty Hash) | Should -Be $hashes
  }

  It 'does not disturb an existing build lock' {
    $null = New-Item -ItemType Directory -Path $script:Output
    $lockPath = Join-Path $script:Output '.press-release-build.lock'
    [IO.File]::WriteAllText($lockPath, 'owned by another build')
    { Invoke-ReleaseBuild } | Should -Throw
    [IO.File]::ReadAllText($lockPath) | Should -Be 'owned by another build'
    @(Get-ChildItem -LiteralPath $script:Output -Force).Count | Should -Be 1
  }

  It 'rejects a file used as OutputDirectory' {
    [IO.File]::WriteAllText($script:Output, 'retain')
    { Invoke-ReleaseBuild } | Should -Throw '*must be a directory*'
    [IO.File]::ReadAllText($script:Output) | Should -Be 'retain'
  }

  It 'does not invoke PDF conversion or publish PDF artifacts' {
    $result = Invoke-ReleaseBuild
    @(Get-ChildItem -LiteralPath $script:Output -Filter '*.pdf').Count | Should -Be 0
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($script:BuildScript, [ref]$tokens, [ref]$errors)
    $commands = $ast.FindAll({ param ($node) $node -is [Management.Automation.Language.CommandAst] }, $true)
    @($commands | Where-Object { $_.GetCommandName() -match 'Convert-HTMLToPDF' }).Count | Should -Be 0
    $ast.ScriptRequirements.RequiredPSVersion | Should -Be ([version]'7.4')
    $result.Status | Should -Be 'Completed'
  }
}
