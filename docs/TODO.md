# ✅ TODOs - winkit

## ➕ Additions

- [x] Add [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer)
- [x] Add [Pester](https://pester.dev/)
- [X] Add scripts to mitigate [Public Firewall Profile Bug in Defender 2025](https://jans.cloud/2025/03/unidentified-network-public-firewall-profile-bei-windows-server-2025-domain-controller/)
- [x] ~~Add scripts for [Terminal Server configuration](https://www.server-world.info/en/note?os=Windows_Server_2025&p=remotedesktop&f=1) and the associated [Defender rules](https://gist.github.com/asheroto/530748b3bf0528cc4805d652b612f81f)~~ -> out of scope - do this with Ansible
- [x] Create [WOL-tooling](https://www.pdq.com/blog/wake-on-lan-wol-magic-packet-powershell/)
- [x] Create a dowloader for [VC++ Redistributables](https://learn.microsoft.com/de-de/cpp/windows/latest-supported-vc-redist?view=msvc-170)
- [x] [How to set return type in PowerShell](https://stackoverflow.com/questions/64891846/specify-function-return-type)

## ✏️ Planned Changes

## 💡 Ideas

- [x] [`fleschutz/PowerShell`](https://github.com/fleschutz/PowerShell) -> see [`Test-PendingReboot.ps1`](/scripts/Test-PendingReboot.ps1) and [`Test-SystemFileIntegrity.ps1`](/scripts/Test-SystemFileIntegrity.ps1)
- [x] [`stevencohn/WindowsPowerShell`](https://github.com/stevencohn/WindowsPowerShell) -> see [`Enable-WinRM.ps1`](/scripts/Enable-WinRM.ps1), [`Clear-EventLogs.ps1`](/scripts/Clear-EventLogs.ps1), [`Get-ADUserReport.ps1`](/scripts/ADDS/Get-ADUserReport.ps1) and [`Test-PendingReboot.ps1`](/scripts/Test-PendingReboot.ps1); deferred module changes logged in `secrets/PSFoundation.md`
- [x] [`WinTweakers/WindowsToolbox`](https://github.com/WinTweakers/WindowsToolbox) -> evaluated: archived, unmaintained consumer debloat tool; every real capability is already covered by existing `Configure-*`/`Remove-*`/`Disable-*` scripts with better engineering. Explicitly rejected after review: `DisableWindowsDefender` (AV-disabling has no place in an ISO-oriented toolkit) and `MSDOSMode` (downloads unverified third-party binaries into System32). No implementation taken; independently confirmed the registry/file ownership-takeover gap (logged in `secrets/PSFoundation.md`).
- [X] [`ScoopInstaller/Scoop`](https://github.com/ScoopInstaller/Scoop)
- [X] [`W4RH4WK/Debloat-Windows-10`](https://github.com/W4RH4WK/Debloat-Windows-10)
- [x] [`asheroto/winget-install`](https://github.com/asheroto/winget-install/tree/master)-> feature-diff applied to [`Install-WinGet.ps1`](/scripts/Install-WinGet.ps1): OS compatibility gate, known-error-code translation, dependency skip-if-current provisioning, Server Core portable fallback (beta caveat carried), SYSTEM-account support (documented risk); cross-script generalizations logged in `secrets/enhancements.md`; -GHtoken deferred until rate-limiting is observed.
- [x] [`asheroto/UninstallTeams`](https://github.com/asheroto/UninstallTeams) -> see [`Remove-Teams.ps1`](/scripts/Remove-Teams.ps1)
- [x] [`jrussellfreelance`](https://github.com/jrussellfreelance/powershell-scripts)

-> adapted: [`Get-ADPasswordExpiry.ps1`](/scripts/ADDS/Get-ADPasswordExpiry.ps1) (FGPP-correct expiry via msDS-UserPasswordExpiryTimeComputed) and [`Get-WSUSReport.ps1`](/scripts/Get-WSUSReport.ps1); `Test-RemoteHostReachability` logged for PSFoundation (`secrets/PSFoundation.md`); secret-handling/download anti-patterns logged in `secrets/enhancements.md`. Rejected after review: plaintext-password SSH scripts, inline-AWS-key CloudWatch script, HTTP Chrome installer, MSOnline-based M365 script, role-bundle DC setup, broken send-email.ps1 (winkit already has Send-SMTPMessage.ps1).
- [x] [`nickrod158/PowerShell-Scripts`](https://github.com/nickrod518/PowerShell-Scripts/tree/master)

-> adapted: [`Get-ADGroupAudit.ps1`](/scripts/ADDS/Get-ADGroupAudit.ps1) (privileged-group last-logon audit, max LastLogon across all DCs); five module ports logged in `secrets/PSFoundation.md` (Find-ServiceAccountUsage, Test-ADCredential, Get-CertificateInventory, Convert-RobocopyExitCode, encrypted-credential-file pair); general conventions logged in `secrets/enhancements.md` (robocopy exit codes, cert-expiry standing report, NTP drift concept). Rejected: Disable-UAC (security-control disabling); non-functional references noted: Exchange Online Basic Auth script and MSOnline-based MSO folder. DHCP filter replication skipped (not needed).
- [ ] [`farag2/Utilities`](https://github.com/farag2/Utilities/tree/master)
- [ ] [`nightroman/Invoke-Build`](https://github.com/nightroman/Invoke-Build)
- [X] [`ab14jain/PowerShell`](https://github.com/ab14jain/PowerShell)
- [ ] [`lazywinadmin/PowerShell`](https://github.com/lazywinadmin/PowerShell)
- [ ] [`awesome-windows11/windows11`](https://github.com/awesome-windows11/windows11)
- [ ] [`LeDragoX/Win-Debloat-Tools`](https://github.com/LeDragoX/Win-Debloat-Tools/blob/main/src/scripts/Optimize-Privacy.ps1)

- [X] [Activate RDP via PowerShell](https://nt4admins.de/powershell/rdp-aktivieren-remote-per-powershell/)
- [X] [PowerShell send SMTP notifications](https://nt4admins.de/powershell/e-mail-benachrichtigungen-mit-der-powershell-senden/) -> see [`Send-SMTPMessage.ps1`](/scripts/Send-SMTPMessage.ps1)
- [ ] [Create E-Mail Disclaimer with PowerShell](https://nt4admins.de/powershell/email-disclaimer-mit-der-powershell-erstellen/)
- [x] [Batch print PDF files](https://brndmp.olafritman.com/batch-print-pdf-with-powershell/)

## 🔗 Links

### General

- [ ] [msxfaq.de](https://www.msxfaq.de/code/powershell/psparam.htm)
- [x] [PowerShell special characters](https://www.neolisk.blog/posts/2009-07-23-powershell-special-characters/)
- [x] [PowerShell: What is @{}](https://stackoverflow.com/questions/56965510/what-is-meaning-in-powershell)
- [x] [PowerShell: What is %{}](https://stackoverflow.com/questions/22846596/what-does-percent-do-in-powershell)
- [x] [PSScriptAnalyzer](https://learn.microsoft.com/de-de/powershell/utility-modules/psscriptanalyzer/overview?view=ps-modules)
- [X] [Copy/Move files based on RegEx](https://stackoverflow.com/questions/7893919/powershell-copy-move-files-based-on-a-regex-value-retaining-the-folder-structu) -> handled by `augment`
- [x] [PowerShell: What is CmdLetBinding](https://blog.ironmansoftware.com/powershell-cmdlet-binding/)
- [x] [Explore self-elevating PowerShell scripts](https://www.reddit.com/r/PowerShell/comments/d7y0zp/self_elevating_powershell_script/)
- [x] [PowerShell: `ConvertTo-SecureString`](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/convertto-securestring?view=powershell-7.5)
- [x] [PowerShell: `-join`](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_join?view=powershell-7.5)
- [x] [PowerShell `ValidatePattern`](https://stackoverflow.com/questions/56583435/using-validateset-and-validatepattern-to-allow-new-values)
- [x] [PowerShell Parameters](https://www.mssqltips.com/sqlservertip/4205/powershell-parameters-part-ii-validateset-and-validatepattern/)
- [x] [PowerShell Array Iteration](https://stackoverflow.com/questions/66528639/how-to-iterate-through-an-array-of-objects-in-powershell)
- [x] [PowerShell SecureString Argument](https://stackoverflow.com/questions/40446583/how-to-pass-the-password-as-argument-in-powershell-script-and-convert-to-secure)

### Infrastructure

- [x] [Configure AD Domain Controller on Server 2025](https://www.server-world.info/en/note?os=Windows_Server_2025&p=active_directory&f=2)
- [x] [Microsoft: Install AD Domain Services with PowerShell](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/deploy/install-active-directory-domain-services--level-100-)
- [x] [Set up a Domain Controller on Server 2025](https://www.windowspro.de/wolfgang-sommergut/domaenen-controller-windows-server-2025-installieren-32k-datenbank-nutzen)
- [x] [Set up a Domain Controller on Server 2025](https://it-learner.de/domaenencontroller-mit-windows-server-2025/)
- [x] [Set up Exchange on Server 2025](https://www.frankysweb.de/en/installation-exchange-2019-cu14-on-windows-server-2025/)
- [x] [ADDS - Functional levels](https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/active-directory-functional-levels)
- [x] [32k pages for AD](https://learn.microsoft.com/de-de/windows-server/identity/ad-ds/32k-pages-optional-feature)
      -> use Ansible from a control node instead

- [X] [Microsoft: CSS Exchange - HealthChecker](https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/) -> used in [Initialize-ExchangeServer.ps1](../scripts/ADDS/Initialize-ExchangeServer.ps1)
- [X] [PowerShell Exchange Web Services (EWS) scripts](https://github.com/David-Barrett-MS/PowerShell-EWS-Scripts) -> not currently used
- [X] [PowerShell Remoting](https://learn.microsoft.com/en-us/powershell/scripting/learn/ps101/08-powershell-remoting?view=powershell-7.5) -> [see Initialize-.. scripts](../scripts/ADDS)
- [X] [PowerShell - About Functions with Advanced Methods](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods?view=powershell-7.5)
- [X] [ADDS: Forests/Domains erstellen bzw. konfigurieren](https://blog.andreas-schreiner.de/2017/09/14/adds-forests-domains-erstellen-konfigurieren/)
- [X] [Domain Controller auf Server 2025 migrieren](https://www.windowspro.de/wolfgang-sommergut/domain-controller-windows-server-2025-migrieren)
- [X] [Activate RDP remotely via PowerShell](https://nt4admins.de/powershell/rdp-aktivieren-remote-per-powershell/)
- [X] [Activate RDP remotely via PowerShell](https://sid-500.com/2021/03/22/enable-remote-desktop-remotely-with-powershell-enable-remotedesktop/)
- [X] [Activate RemoteDesktop Firewall Rules](https://www.der-windows-papst.de/2018/03/30/powershell-remotedesktop-aktivieren-und-firewall/)
- [x] [Explore `PSRename`](https://github.com/mgajda83/PSRename)
- [x] [Install `winget` via the Command-Line](https://stackoverflow.com/questions/74166150/install-winget-by-the-command-line-powershell)

### E-Mails

- [x] [Managing an Outlook Mailbox with PowerShell](https://learn.microsoft.com/en-us/archive/msdn-magazine/2013/march/powershell-managing-an-outlook-mailbox-with-powershell)
- [x] [Export Outlook to .pst file](https://gist.github.com/timsonner/d9d9894ca8ee026d8f8285d51a5fdacc)
- [x] [Use PowerShell to move, delete and measure Outlook emails](https://gregcarriger.wordpress.com/2018/12/28/use-powershell-to-move-delete-and-measure-outlook-emails/)
- [x] [Determine Outlook default folder when using different Language](https://stackoverflow.com/questions/32964591/powershell-get-where-outlook-by-default-stores-its-data-files-when-outlook-is-i)
- [x] [Example Script: Move Emails to archive folder](https://stackoverflow.com/questions/28880777/powershell-moving-emails-to-archive-folder-moves-all-but-last-email)

-> see [`Archive-OutlookMails.ps1`](/scripts/Office/Archive-OutlookMails.ps1)

- [x] [Tips for writing cross-platform PowerShell scripts](https://powershell.org/2019/02/tips-for-writing-cross-platform-powershell-code/)
- [x] [PowerShell path separator based on OS](https://rakhesh.com/powershell/powershell-path-separator-based-on-os/)

-> decided against (use Bash on Darwin/Unix..)

### Repository

- [x] [Pester - PowerShell testing framework](https://github.com/pester/Pester/tree/main)
- [x] [`posh-git`](https://github.com/dahlbyk/posh-git) -> out of scope
- [x] [`scoop` CLI](https://github.com/ScoopInstaller/Scoop/tree/master) -> not implementing a CLI
- [x] [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer)
- [x] [Run `pre-commit` PowerShell script](https://scriptingnerd.com/2024/11/01/running-a-powershell-script-on-your-code-before-a-git-commit-using-pre-commit/)
- [PSResourceGet supported repositories](https://learn.microsoft.com/en-us/powershell/gallery/powershellget/supported-repositories?view=powershellget-3.x)
- [Getting started with PowerShell gallery](https://learn.microsoft.com/en-us/powershell/gallery/getting-started?view=powershellget-3.x)

### Windows Updates

- [x] [Automate Windows Updates with PowerShell](https://www.reddit.com/r/sysadmin/comments/1j5c810/how_to_update_windows_store_apps_via_commandline/)
- [x] [Use Windows CIM to automate Updates with PowerShell](https://gist.github.com/wise-io/70687623206a4175ca591bccc8971e7f)
- [x] [Make use of PSWindowsUpdate](https://www.windowspro.de/wolfgang-sommergut/windows-updates-powershell-pswindowsupdate-auflisten-herunterladen-installieren)

-> see [`lib/updates.ps1`](/lib/updates.ps1)

### Windows Registry

- [x] [Guide to understanding ProgIDs and File Type Associations](https://setuserfta.com/guide-to-understanding-progids-and-file-type-associations/)
- [x] [GitHub: DanysysTeam/PS-SFTA](https://github.com/DanysysTeam/PS-SFTA)
- [x] [GitHub: DanysysTeam/PS-TBPin](https://github.com/DanysysTeam/PS-TBPin)
- [x] [Install Windows Updates via PowerShell](https://learn.microsoft.com/en-us/answers/questions/1613848/update-and-restart-from-powershell-or-command-line)

-> see [`lib/registry.ps1`](/lib/registry.ps1) and [`lib/settings.ps1`](/lib/settings.ps1)

### Windows Printing

- [x] [Printing Documents using PowerShell](https://pipe.how/invoke-print/)
- [x] [Streamlining batch PDF printing](https://medium.com/@mayberryjalin/powershell-streamlining-batch-pdf-printing-301f25c1cd03)
- [x] [Trying to send PDF print jobs to a network printer](https://community.spiceworks.com/t/trying-to-send-pdf-print-jobs-to-a-network-printer/957742)

-> see [`Invoke-PrintBatch.ps1`](/scripts/Files/Invoke-PrintBatch.ps1)

## ⚠️ Issues

- [DISM CmdLets not working in PowerShell 7](https://github.com/PowerShell/PowerShell/issues/15428)
