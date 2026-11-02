# Administration

| Task                                                                         | Guide                                                         |
| ---------------------------------------------------------------------------- | ------------------------------------------------------------- |
| Import data folders from an extracted backup into a new Windows user profile | [User data migration](User-Data-Migration.md)                 |
| Find administration scripts and their requirements                           | [Script catalogue](../../../scripts/Administration/README.md) |

## Open the installed toolkit

The [winkit installer](../../../dist/README.md) supplies the exact `AdNoctem.Substrate.PowerShell` dependency. Use x64 Windows PowerShell
5.1 or PowerShell 7 on x64 Windows, and select the directory matching your installation:

```powershell
# Default CurrentUser installation, using the account that installed winkit.
Set-Location -LiteralPath "$env:LOCALAPPDATA\Programs\winkit"
```

```powershell
# Default AllUsers installation.
Set-Location -LiteralPath "$env:ProgramFiles\winkit"
```

For a custom installation, use the directory chosen with `WINKIT_INSTALL_PATH`. Run migration commands in an elevated PowerShell session as
a separate administrator, with the target user signed out. Server Core is supported. SYSTEM execution is supported when account resolution
and filesystem access are available.

Keep extracted backups and reports outside the managed installation. The migration guide uses `C:\Migration\User1` for staged data and the
script defaults to `%ProgramData%\winkit\ProfileImports` for reports.
