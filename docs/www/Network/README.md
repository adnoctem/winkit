# Network

Prepare firewall configuration files and use the installed network utilities.

| Task                                                                             | Guide                                                  |
| -------------------------------------------------------------------------------- | ------------------------------------------------------ |
| Clone a pfSense backup, change interfaces or UI settings, and restore the result | [Firewall configuration](Firewall-Configuration.md)    |
| Find network scripts and their requirements                                      | [Script catalogue](../../../scripts/Network/README.md) |

## Open the installed toolkit

The [winkit installer](../../../dist/README.md) supplies the pinned module dependency. Use x64 Windows PowerShell 5.1 or PowerShell 7 on x64
Windows. Select the directory matching your installation:

```powershell
# Default CurrentUser installation.
Set-Location -LiteralPath "$env:LOCALAPPDATA\Programs\winkit"
```

```powershell
# Default AllUsers installation.
Set-Location -LiteralPath "$env:ProgramFiles\winkit"
```

For a custom installation, use the location chosen with `WINKIT_INSTALL_PATH`. Run the guide's examples from that directory. The pfSense
builder needs no elevation and supports Server Core and SYSTEM accounts with access to the input and output files.

Store backups, specs, and generated configurations **outside the managed winkit directory**. The guide uses `C:\FirewallBuilds`; substitute
a folder you can write to and whose access is restricted to the appropriate administrators.
