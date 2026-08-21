# winkit — Agent Guidance

`winkit` is a PowerShell script collection for Windows configuration, hardening, and provisioning. It contains **no module code**: shared functionality lives in the separate `PSFoundation` module (PowerShell Gallery), which every script imports by name. The root `AGENTS.md` is a symlink to this file — edit this one.

## Commands (run from repo root)

- `.\winkit.ps1 init` — install the exact module versions pinned in `requirements.psd1` (aliases: `setup`, `bootstrap`)
- `.\winkit.ps1 format` — in-place PSScriptAnalyzer formatting; `-Check` = read-only mode (used by CI and pre-commit)
- `.\winkit.ps1 lint` — PSScriptAnalyzer rule checks, exit 1 on findings
- `.\winkit.ps1 test` — run the winkit Pester suite under `tests/` (logic tests, no Outlook); `-Outlook` also runs the integration suite (needs Outlook, uses a scratch PST). Module logic tests live in the PSFoundation repo
- `.\winkit.ps1 deps` — compare `requirements.psd1` pins against the PowerShell Gallery (`-Check` reports only)
- `.\winkit.ps1 build` — create `dist/` archives (ZIP + tar.gz) from `scripts/`

## Layout

- `scripts/` — the product: executable admin scripts; subfolders `ADDS/`, `DHCP/`, `Files/`, `Office/`
- `bin/` — `.cmd` launchers for selected scripts
- `tools/` — development tooling behind `winkit.ps1` (initialize, format, lint, test, dependencies, build)
- `docs/` — contributor documentation; TODO.md is currently all-TBA
- `secrets/` — local-only scratch/backlog docs, **gitignored**: `enhancements.md`, `PSFoundation.md`
- `requirements.psd1` — exact module version pins (also a semantic-release commit asset); bump when adopting newer PSFoundation functionality
- `PSScriptAnalyzerSettings.psd1` — formatter/linter rules; contains a singular-noun allowlist

## Script conventions (non-negotiable)

- Every script opens with `#Requires -Version 5.0`, `#Requires -RunAsAdministrator` (when elevation is required), and `#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = 'x.y.z' }` — pin the lowest version that provides every cmdlet used
- `Import-Module PSFoundation -Force` after the param block
- Comment-based help with `.SYNOPSIS`/`.DESCRIPTION`/`.PARAMETER`/`.EXAMPLE`, plus a `.NOTES` block (Author, License, Server Core support, SYSTEM-account suitability)
- `[CmdletBinding(SupportsShouldProcess = $true)]` with `-DryRun` (sets `$WhatIfPreference`) and `-PassThru` switches
- Results use PSFoundation `New-OperationResult`/`Add-OperationResult`; user feedback via `Write-Log -Message -Color`
- PowerShell 5.0 compatibility: no ternary, `??`, or null-conditional operators (PSUseCompatibleSyntax targets 5.0/5.1/7.0)
- Encoding: UTF-8 with BOM, CRLF (enforced by formatter and pre-commit `mixed-line-ending`); 2-space indentation
- New plural nouns in function/script names must be added to the `PSUseSingularNouns` `NounAllowList` in `PSScriptAnalyzerSettings.psd1`

## Gotchas

- `docs/CONTRIBUTING.md` still references the retired `lib/` module and a `tests/` directory — both were split out into PSFoundation; trust the code over that document
- Format/lint exclude `secrets/` and `dist/` by default (`-IncludeSecrets` overrides)
- Commit messages must be conventional (types `build|ci|docs|feat|fix|perf|refactor|test|chore`; scopes `lib|scripts|tools|config|docs` — `lib` is historical) — see `docs/CONTRIBUTING.md`
- Releases are semantic-release driven: superlint gates the push, then dispatches `release`; a `feat` commit on `main` releases a minor, `fix` a patch
- Local pre-commit hooks require `bun` (prettier) — run `pre-commit install` after `init`
- `winkit.ps1` deliberately omits `[CmdletBinding()]` so `$args` pass through to the tool scripts — keep it that way
