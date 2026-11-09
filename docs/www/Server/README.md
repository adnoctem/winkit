# Server

Prepare and inspect Active Directory, DHCP, Exchange, Group Policy, and WSUS servers.

| Task                                                                        | Guide                                                   |
| --------------------------------------------------------------------------- | ------------------------------------------------------- |
| Export AD users, computers, and groups, or generate the `plead` account map | [Active Directory exports](Active-Directory-Exports.md) |
| Find server scripts and their requirements                                  | [Script catalogue](../../../scripts/Server/README.md)   |

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

For a custom installation, use the location chosen with `WINKIT_INSTALL_PATH`. Run the guides from that directory.

The AD read scripts need the ActiveDirectory PowerShell module and connectivity to AD Web Services. They do not need elevation, but the
account must be able to read the selected objects. Store exports, maps, and reports **outside the managed winkit directory** in a folder
whose access is restricted to the administrators who handle the account data.
