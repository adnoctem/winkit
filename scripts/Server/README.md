# Server scripts

Run with x64 Windows PowerShell 5.1 or PowerShell 7 on x64 Windows. The winkit installer provides the pinned `AdNoctem.Substrate.PowerShell`
dependency. The Active Directory scripts also need the `ActiveDirectory` PowerShell module and connectivity to AD Web Services; none of the
AD read scripts require elevation.

| Script                                                                           | Purpose                                                               | Requirements and limits                                                                                    |
| -------------------------------------------------------------------------------- | --------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------- |
| [ADDS/Export-AD.ps1](ADDS/Export-AD.ps1)                                         | Export AD users, computers, and groups to CSV or JSON.                | ActiveDirectory module; read permissions; no elevation. Missing attributes fail unless omitted or allowed. |
| [ADDS/Export-ADAccountMap.ps1](ADDS/Export-ADAccountMap.ps1)                     | Generate the deterministic `ad-account-map.json` consumed by `plead`. | ActiveDirectory module; explicit domain allowlist; override CSV for unreliable `mail` values.              |
| [ADDS/Get-ADUserReport.ps1](ADDS/Get-ADUserReport.ps1)                           | Report one user's attributes, group memberships, and password age.    | Global Catalog reachability; no elevation.                                                                 |
| [ADDS/Get-ADPasswordExpiry.ps1](ADDS/Get-ADPasswordExpiry.ps1)                   | Report users whose passwords expire within a threshold.               | ActiveDirectory module; read permissions.                                                                  |
| [ADDS/Get-ADGroupAudit.ps1](ADDS/Get-ADGroupAudit.ps1)                           | Audit last-logon times for privileged group members.                  | ActiveDirectory module; queries every domain controller.                                                   |
| [ADDS/Get-ADExpiringAccount.ps1](ADDS/Get-ADExpiringAccount.ps1)                 | Report expiring AD accounts.                                          | ActiveDirectory module; read permissions.                                                                  |
| [ADDS/Add-ADController.ps1](ADDS/Add-ADController.ps1)                           | Add a domain controller.                                              | Elevation; AD DS role tooling.                                                                             |
| [ADDS/Initialize-ADController.ps1](ADDS/Initialize-ADController.ps1)             | Prepare a domain controller.                                          | Elevation; AD DS role tooling.                                                                             |
| [ADDS/Test-ADController.ps1](ADDS/Test-ADController.ps1)                         | Validate a domain controller.                                         | Elevation; AD DS role tooling.                                                                             |
| [ADDS/New-OfflineDomainJoin.ps1](ADDS/New-OfflineDomainJoin.ps1)                 | Create an offline domain-join blob.                                   | Elevation; AD DS tooling.                                                                                  |
| [DHCP/Export-DHCP.ps1](DHCP/Export-DHCP.ps1)                                     | Export DHCP server configuration.                                     | DHCP Server module; elevation.                                                                             |
| [DHCP/Import-DHCP.ps1](DHCP/Import-DHCP.ps1)                                     | Import DHCP server configuration.                                     | DHCP Server module; elevation.                                                                             |
| [Exchange/Initialize-ExchangeServer.ps1](Exchange/Initialize-ExchangeServer.ps1) | Prepare an Exchange server.                                           | Exchange tooling; elevation.                                                                               |
| [Exchange/Test-ExchangeServer.ps1](Exchange/Test-ExchangeServer.ps1)             | Validate an Exchange server.                                          | Exchange tooling.                                                                                          |
| [GPO/Get-ADGPOReplication.ps1](GPO/Get-ADGPOReplication.ps1)                     | Report Group Policy replication state.                                | AD and GPO tooling.                                                                                        |
| [WSUS/Get-WSUSReport.ps1](WSUS/Get-WSUSReport.ps1)                               | Report WSUS state.                                                    | WSUS role tooling; elevation.                                                                              |

See the [Server overview](../../docs/www/Server/README.md) and the
[Active Directory exports guide](../../docs/www/Server/Active-Directory-Exports.md) for installed paths, examples, and the account-map
contract.

From the installed winkit directory:

```powershell
.\scripts\Server\ADDS\Export-AD.ps1 -OutputPath C:\ADReports -Format Json -IncludeMissing
.\scripts\Server\ADDS\Export-ADAccountMap.ps1 -AllowedDomain adnoctem.co,example.com -OutputPath C:\ADReports\ad-account-map.json
Get-Help .\scripts\Server\ADDS\Export-ADAccountMap.ps1 -Full
```
