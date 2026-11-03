# Press-release document builds

`Build-PressRelease.ps1` requires **x64 PowerShell 7.4 or later** and exactly **AdNoctem.Substrate.PowerShell 0.3.1**, supplied by the
winkit installer. It needs no Word installation or elevation. This is a non-interactive build engine; a selection UI is future work.

Start in the installed winkit directory, normally `C:\Program Files\winkit`. Open your approved authoring `.dotx` in Word, complete its
fields and save a normal `.docx`. The saved document supplies the language, schema and authoring-template version through its metadata.
Renaming a file does not change its language. A `.dotx` cannot be used as the source.

Supply `-PackPath` explicitly: an external, versioned directory containing `pack.json`, HTML templates, DOCX bases and authoring templates
with matching hashes. winkit does not search for, distribute or modify packs. Keep the source and pack unchanged during a build.

## Build HTML and DOCX

Keep your data outside winkit's managed installation. Use a separate output directory for each build/language because the directory has one
`build.manifest.json`. For example:

```text
C:\ReleaseWork\announcement\
  Input\release-en.docx
  Output\en\release-en.html
  Output\en\release-en.docx
  Output\en\build.manifest.json
C:\ReleasePacks\v1\pack.json
```

Run these commands in PowerShell 7.4+:

```powershell
Set-Location 'C:\Program Files\winkit'
$build = @{
  InputDocument   = 'C:\ReleaseWork\announcement\Input\release-en.docx'
  PackPath        = 'C:\ReleasePacks\v1'
  OutputDirectory = 'C:\ReleaseWork\announcement\Output\en'
}

.\scripts\Files\Build-PressRelease.ps1 @build -DryRun
$result = .\scripts\Files\Build-PressRelease.ps1 @build -PassThru
$result | Format-List Status, Language, PackVersion, HtmlPath, DocxPath, ManifestPath
```

The output basename defaults to the source basename. `-OutputName 'announcement-en'` selects another basename without changing the source.
Names must be single safe filenames, without directory separators, reserved device names or trailing dots/spaces. The generated DOCX can
never be the input path. Output directories cannot be inside the pack. Junctions and symbolic links in build paths are rejected.

Existing HTML, DOCX **or the build manifest** cause a collision. To deliberately replace a previous build:

```powershell
$result = .\scripts\Files\Build-PressRelease.ps1 @build -Overwrite -PassThru
```

`-DryRun` and `-WhatIf` load the pack, validate the input and check collisions without creating files or directories. Their result is
`Skipped`, with intended paths and no output hashes. A successful real build returns `Completed`.

## Diagnostics, publication and audit

Every source diagnostic includes severity, code, field, message and remedy. Errors or a missing release stop the build before output is
written. Warnings allow generation and are included in both the returned result and manifest. Pack failures include the document failure
reason and message. Failures terminate, so a caller can use `try`/`catch`; the error's `TargetObject` carries the failed result and
diagnostics.

```powershell
try {
  $result = .\scripts\Files\Build-PressRelease.ps1 @build -Overwrite -PassThru
} catch {
  $_.TargetObject.Diagnostics | Format-Table Severity, Code, Field, Message, Remedy -Wrap
  throw
}
```

Both writers finish in a temporary directory before publication starts. A directory lock prevents cooperating builds from interleaving.
Publication uses individual atomic file renames and writes the UTF-8, no-BOM manifest last. It is not a transaction across all three files:
readers should wait for the build to finish and verify hashes. On a publication error the script attempts to restore the previous files. The
error explicitly reports publication having begun, any unrestored outputs and retained recovery files if rollback fails. A process or
machine crash may require manual recovery. Do not run another program that edits the output directory during publication.

Validation or generation failure preserves an existing successful build and its manifest. No failure manifest replaces that audit record.
The successful `press-release-build/1` manifest contains UTC generation time, builder version/script hash, Substrate version, input
path/SHA-256/language/template version/schema, pack path/version/layout version/`pack.json` SHA-256, output paths/SHA-256 and diagnostics.
`-PassThru` returns status, language, pack version, HTML/DOCX paths and hashes, manifest path and diagnostics. Without it, the script logs a
concise summary. Manifests contain absolute paths; treat them as part of your release's working data.

## Optional PDF conversion

The caller or future UI invokes `Convert-HTMLToPDF.ps1` separately after `Status` is `Completed`. Select a configured backend; for example,
the local PSWritePDF backend:

```powershell
if ($result.Status -eq 'Completed') {
  $pdf = [IO.Path]::ChangeExtension($result.HtmlPath, '.pdf')
  try {
    .\scripts\Files\Convert-HTMLToPDF.ps1 -InputPath $result.HtmlPath -OutputPath $pdf -Mode PS
  } catch {
    Write-Warning "PDF conversion failed; the completed HTML/DOCX build remains valid: $_"
  }
}
```

Use `Get-Help .\scripts\Files\Convert-HTMLToPDF.ps1 -Full` for cloud modes and backend requirements. PDF rendering quality depends on that
backend. A PDF failure must not change the HTML/DOCX build status or its manifest. The build engine never invokes a PDF converter.
