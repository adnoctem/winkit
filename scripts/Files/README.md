# File tools

Run scripts from the installed winkit directory, normally `C:\Program Files\winkit`. Keep input documents, template packs and generated
files outside that managed directory. See the [file workflow guides](../../docs/www/Files/README.md).

| Script                         | Purpose and requirements                                                                                                                                                                                       |
| ------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Build-PressRelease.ps1`       | Validate a saved DOCX against an explicit template pack; publish HTML, DOCX and a hash manifest. Requires PowerShell 7.4+ and Substrate 0.3.1. See [document builds](../../docs/www/Files/Document-Builds.md). |
| `Convert-HTMLToPDF.ps1`        | Separate HTML-to-PDF conversion using EuroPDF, DocRaptor or local PSWritePDF. Requires PowerShell 7 and the selected backend; cloud modes need credentials.                                                    |
| `Convert-TextFileEncoding.ps1` | Convert text file encoding. Use the script help for encoding and file selection options.                                                                                                                       |
| `Invoke-PrintBatch.ps1`        | Submit a batch of files to a printer; requires an available printer and suitable file handlers.                                                                                                                |
| `Move-FilesRegex.ps1`          | Move files selected by a regular expression. Preview with `-WhatIf` before applying changes.                                                                                                                   |

Use `Get-Help .\scripts\Files\<script>.ps1 -Full` for each command's parameters and examples.
