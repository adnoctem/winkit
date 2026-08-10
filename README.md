<p align="center">
    <!-- PowerShell -->
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/Powershell_256.png">
      <img src="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/Powershell_256.png" width="225" alt="PowerShell">
    </picture>
    <!-- winkit -->
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/adnoctem/artwork/6da215b35cd1482db4e6b5c19e40e22105576785/projects/winkit/icon/color/winkit-icon-color.png">
      <img src="https://raw.githubusercontent.com/adnoctem/artwork/6da215b35cd1482db4e6b5c19e40e22105576785/projects/winkit/icon/color/winkit-icon-color.png" width="225" alt="winkit">
    </picture>
</p>

[![License](https://img.shields.io/github/license/adnoctem/winkit?label=License)][license]
[![Language](https://img.shields.io/github/languages/top/adnoctem/winkit?label=PowerShell)][powershell]
[![Linting](https://github.com/adnoctem/winkit/actions/workflows/superlint.yaml/badge.svg)][ci_linting_workflow]
[![GitHub Release](https://img.shields.io/github/v/release/adnoctem/winkit?label=Release)][github_releases]
[![GitHub Activity](https://img.shields.io/github/commit-activity/m/adnoctem/winkit?label=Commits)][github_commits]
[![Semantic Release](https://img.shields.io/badge/Semantic_Release-enabled-brightgreen?logo=semanticrelease&logoColor=E5E4E7)][semantic_release]
[![Renovate](https://img.shields.io/badge/Renovate-enabled-brightgreen?logo=renovate&logoColor=1A1F6C)][renovate]
[![PreCommit](https://img.shields.io/badge/PreCommit-enabled-brightgreen?logo=precommit&logoColor=FAB040)][precommit]

`winkit` is an open-source [MIT][license]-licensed collection of [PowerShell][powershell] scripts and modules written and maintained by the [Ad Noctem Collective][org] for the machine-level configuration, hardening, provisioning and maintenance of Windows systems. The repository targets both desktop Windows installations and Windows Server environments and supports [PowerShell][powershell] 5.0 and above, including Windows PowerShell 5.x as well as newer PowerShell releases. It is intended for repeatable system setup, workstation preparation, administrative automation and post-install configuration, especially where Windows machines need to be brought into a known and documented configuration state.

The [`scripts`](scripts) directory contains executable scripts intended to be run directly by an administrator, deployment process, bootstrap routine, scheduled task or other automation system. Scripts depend on the [`PSFoundation`](https://www.powershellgallery.com/packages/PSFoundation) module, which is installed from the PowerShell Gallery; the [`bin`](bin) directory holds convenience `.cmd` launchers for selected entry-point scripts.

Functionally, `winkit` covers several major areas of Windows administration, including system appearance, Explorer behavior, Start menu and taskbar configuration, browser policies, privacy options, telemetry reduction, diagnostic reporting and notification settings. It also includes scripts for disabling or enabling selected Windows components such as OneDrive, Copilot, Game DVR, Remote Assistance, Windows Sandbox and Windows Subsystem for Linux.

For provisioning and maintenance workflows, the repository provides automation around PowerShell Core, WinGet, Visual C++ redistributables, Windows updates, system restore points, hostname changes, mapped drive restoration, file encoding conversion and cleanup operations. Additional scripts cover device and input behavior such as pointer acceleration, compact OS settings, application permission defaults, Python script path configuration and terminal experience defaults.

Network-adjacent and administrative utility scripts are included as well, such as routines for Group Policy backup, DNS zone updates, information retrieval and Wake-on-LAN packet generation. The scripts are intentionally separated by concern so that individual configuration steps can be reviewed, composed and executed independently depending on the target machine, deployment scenario or administrative policy.

For more information on PowerShell itself, refer to Microsoft's official [PowerShell Documentation][powershell_docs]. Existing scripts may also serve as implementation examples for composing new automation routines or reusing the shared functionality provided by the [`PSFoundation`][psfoundation] module.

## ✨ TL;DR

```pwsh
# initialize the project (download dependencies)
.\winkit.ps1 init
# also: .\winkit.ps1 initialize | setup | bootstrap

# format all PowerShell source files
.\winkit.ps1 format
# also: .\winkit.ps1 fmt | fix

# check formatting without modifying (CI / pre-commit)
.\winkit.ps1 format -Check

# run PSScriptAnalyzer lint checks
.\winkit.ps1 lint
# also: .\winkit.ps1 check | analyze

# build distribution archives (ZIP + tar.gz)
.\winkit.ps1 build
# also: .\winkit.ps1 bundle | package

# ─── Example scripts ────────────────────────────────

# show comprehensive system information
.\scripts\Show-SystemInfo.ps1

# apply a system configuration profile
.\scripts\Configure-System.ps1

# remove Windows bloatware
.\scripts\Remove-Bloatware.ps1

# enable the Windows Sandbox feature
.\scripts\Enable-WindowsSandbox.ps1

# 'sudo' usage example, if not using sudo.ps1 (from microsoft/sudo)
sudo pwsh -File .\scripts\Install-WindowsUpdates.ps1 -Profile Recommended -DryRun

# display full parameter documentation
help .\scripts\Configure-Privacy.ps1
```

### 🔃 Contributing

Contributions are welcome via GitHub's Pull Requests. Fork the repository and implement your changes within the forked
repository, after that you may submit a [Pull Request][gh_pr_fork_docs]. Refer to our [documentation for contributors][contributing]
for contributing guidelines, commit message formats and versioning tips.

### 📥 Maintainers

This project is owned and maintained by [Ad Noctem Collective](https://github.com/adnoctem) refer to
the [`AUTHORS`][authors] or [`CODEOWNERS`][owners] for more information. You may also use the linked
contact details to reach out directly.

### ©️ Copyright

_Assets provided by:_ **[Microsoft Corporation][microsoft]**

<!-- File references -->

[license]: LICENSE
[contributing]: docs/CONTRIBUTING.md
[authors]: .github/AUTHORS
[owners]: .github/CODEOWNERS
[ci_linting_workflow]: https://github.com/adnoctem/winkit/actions/workflows/superlint.yaml

<!-- General links -->

[org]: https://github.com/adnoctem
[microsoft]: https://www.microsoft.com/
[powershell]: https://github.com/PowerShell/PowerShell
[powershell_docs]: https://learn.microsoft.com/de-de/powershell/
[gh_pr_fork_docs]: https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/proposing-changes-to-your-work-with-pull-requests/creating-a-pull-request-from-a-fork
[github_releases]: https://github.com/adnoctem/winkit/releases
[github_commits]: https://github.com/adnoctem/winkit/commits/main/
[psfoundation]: https://www.powershellgallery.com/packages/PSFoundation

<!-- Third-party -->

[semantic_release]: https://semantic-release.org/
[renovate]: https://renovatebot.com/
[precommit]: https://pre-commit.com/
