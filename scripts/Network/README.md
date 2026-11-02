# Network scripts

Run with x64 Windows PowerShell 5.1 or PowerShell 7 on x64 Windows. The winkit installer provides the pinned `AdNoctem.Substrate.PowerShell`
dependency.

| Script                                                   | Purpose                                                                                      | Requirements and limits                                                                                                |
| -------------------------------------------------------- | -------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| [Build-PfSenseConfig.ps1](Build-PfSenseConfig.ps1)       | Derive a full pfSense backup using named settings, XPath edits, and repeatable data files.   | Local XML files; no elevation or firewall connection. Rejects encrypted and partial backups.                           |
| [New-WGPeerConfig.ps1](New-WGPeerConfig.ps1)             | Generate WireGuard peer files from a pfSense export or REST API and local private-key files. | Per-user folders; API mode requires the pfSense REST API package and credentials. Review deployment-specific defaults. |
| [Find-MACAddress.ps1](Find-MACAddress.ps1)               | Resolve a device's MAC address through DHCP, neighbor cache, and probing.                    | DHCP lookup needs the DhcpServer module and access to the server.                                                      |
| [Send-WOLPacket.ps1](Send-WOLPacket.ps1)                 | Send a Wake-on-LAN magic packet.                                                             | Wake-on-LAN enabled on the target; a reachable broadcast network.                                                      |
| [Set-SecureDns.ps1](Set-SecureDns.ps1)                   | Configure adapter DNS and supported DNS-over-HTTPS settings.                                 | Elevation; DoH configuration requires Windows 11 22H2 or newer.                                                        |
| [Optimize-TCPParameters.ps1](Optimize-TCPParameters.ps1) | Apply optional TCP latency registry settings per adapter.                                    | Elevation; throughput may decrease.                                                                                    |
| [Send-SMTPMessage.ps1](Send-SMTPMessage.ps1)             | Send plain-text or HTML mail with optional attachments.                                      | A reachable SMTP relay; credentials when required by the relay.                                                        |

See the [Network overview](../../docs/www/Network/README.md) for installed paths and the
[firewall configuration guide](../../docs/www/Network/Firewall-Configuration.md) for previews, reusable specs, interface changes, and
restore instructions.

From the installed winkit directory, with a full backup stored at `C:\FirewallBuilds\config.xml`:

```powershell
.\scripts\Network\Build-PfSenseConfig.ps1 -Path C:\FirewallBuilds\config.xml -Hostname fw-branch -DryRun
Get-Help .\scripts\Network\Build-PfSenseConfig.ps1 -Full
Get-Help .\scripts\Network\New-WGPeerConfig.ps1 -Full
```
