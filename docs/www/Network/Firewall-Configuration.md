# Firewall configuration

`Build-PfSenseConfig.ps1` reads a full pfSense XML backup and writes a separate configuration. It can change system identity, DNS, time,
webConfigurator, SSH, existing interfaces, and IPv4 DHCP settings. XPath edits cover settings without a named parameter. The script works
offline; it does not connect to or restore a firewall.

Use the [installed working directory](README.md#open-the-installed-toolkit) for these examples.

## Export and preview

In pfSense, open **Diagnostics > Backup & Restore**, choose **All** as the backup area, and download an unencrypted configuration. Save it
as `C:\FirewallBuilds\config.xml`. Encrypted, malformed, and partial backups are rejected. Keep the original backup for recovery.

Preview a branch firewall with a new LAN and matching DHCP pool:

```powershell
$build = @{
  Path              = 'C:\FirewallBuilds\config.xml'
  OutputPath        = 'C:\FirewallBuilds\branch.xml'
  Hostname          = 'fw-branch'
  Domain            = 'branch.example.com'
  RemoveSshHostKeys = $true
  Interface         = @{
    lan = @{
      IPAddress = '10.20.0.1/24'
    }
  }
  DhcpServer        = @{
    lan = @{
      RangeFrom = '10.20.0.100'
      RangeTo   = '10.20.0.199'
    }
  }
}
.\scripts\Network\Build-PfSenseConfig.ps1 @build -DryRun
```

`-DryRun` and `-WhatIf` validate the edits and return preview results without writing a file. After reviewing them, build the file:

```powershell
$results = .\scripts\Network\Build-PfSenseConfig.ps1 @build -PassThru
$results | Where-Object Action -EQ Build | Format-List
```

The output folder must exist. Without `-OutputPath`, `config.xml` produces `config.built.xml` beside the source. Existing output requires
`-Force`; the source path cannot be used as the output. A temporary output must parse successfully before it replaces the destination.
Failures return exit code 1 and a `Failed` build result when `-PassThru` is supplied.

## Common settings

```powershell
$settings = @{
  Path                   = 'C:\FirewallBuilds\config.xml'
  OutputPath             = 'C:\FirewallBuilds\admin-settings.xml'
  DnsServer              = @('9.9.9.9', '149.112.112.112')
  DnsAllowOverride       = $false
  Timezone               = 'Europe/Berlin'
  TimeServer             = @('0.de.pool.ntp.org', '1.de.pool.ntp.org')
  Language               = 'en_US'
  WebGuiProtocol         = 'https'
  WebGuiPort             = 8443
  WebGuiTheme            = 'pfSense-dark'
  WebGuiDashboardColumns = 3
  WebGuiTopNavigation    = 'Fixed'
  WebGuiHostnameInMenu   = 'Fqdn'
  WebGuiLoginColor       = '1e3f75'
  WebGuiLoginShowHost    = $true
  SshEnabled             = $true
  SshAuthentication      = 'Key'
}
.\scripts\Network\Build-PfSenseConfig.ps1 @settings -DryRun
```

Replacing the DNS list resets its gateway bindings to `none` and clears associated TLS hostnames; supplying the unchanged list preserves
them. Use XPath for specific gateway or TLS hostname bindings. `-DnsServer @()` clears the list. `-TimeServer @()` and
`-WebGuiAlternateHostname @()` remove those settings. Theme, language, time zone, and gateway availability must be checked on the target.

SSH authentication accepts `PasswordOrKey`, `Key`, or `PasswordAndKey`. Key authentication requires suitable user authorized keys already in
the backup. `-RemoveSshHostKeys` removes only the stored host keys under `sshdata`; it preserves user authorized keys, certificates, and VPN
secrets. The backup UI offers the corresponding host-key omission behavior in
[Netgate's backup options](https://docs.netgate.com/pfsense/en/latest/backup/configuration.html#backup-options).

## Interfaces and DHCP

Use internal interface names (`wan`, `lan`, `opt1`, and so on) or unique existing descriptions. Every selected interface must already exist.
Internal names remain stable if you also change a description.

```powershell
$interfaces = @{
  wan = @{
    IPAddress    = 'dhcp'
    BlockPrivate = $false
  }
  opt1 = @{
    Description = 'Guest & IoT'
    Enable      = $true
    Device      = 'vtnet2'
    IPAddress   = '10.30.0.1/24'
    IPv6Address = '2001:db8:30::1/64'
    MTU         = 1500
  }
}
.\scripts\Network\Build-PfSenseConfig.ps1 -Path C:\FirewallBuilds\config.xml -Interface $interfaces -DryRun
```

IPv4 accepts a static address, CIDR, `dhcp`, or `none`. IPv6 accepts a static address, CIDR, `dhcp6`, `slaac`, `track6`, or `none`. A new
static address needs a prefix supplied by CIDR, `Subnet`/`SubnetV6`, or an existing prefix. Dynamic modes remove the prefix; they cannot be
combined with an explicit prefix. Mode-specific options, such as the interface to track for IPv6, require XPath edits when not already
configured. Interface edits do not automatically update rules, NAT, routes, aliases, VPNs, or DHCP ranges.

Other interface settings are `Gateway`, `GatewayV6`, `MSS`, `SpoofMac`, and `BlockBogons`. An empty string removes optional description,
gateway, MTU, MSS, or spoofed-MAC settings. `DhcpServer` accepts `Enable`, `RangeFrom`, `RangeTo`, `Gateway`, `Domain`, `DnsServer` (up to
four IPv4 addresses), `DefaultLeaseTime`, and `MaxLeaseTime`. An enabled DHCP range outside its static interface subnet produces a warning.
This check does not validate all DHCP constraints or firewall relationships.

## XPath edits

Paths are relative to `pfsense`, or absolute when starting with `/`. Select elements, not attributes or text nodes.

```powershell
$edits = [ordered]@{
  'system/webgui/max_procs'                    = 4
  'system/webgui/althostnames'                 = 'fw-branch.branch.example.com'
  'system/webgui/disablehttpredirect'          = $false
  'aliases/alias[name="Admins"]/address'       = '10.20.0.10 10.20.0.11'
}
$remove = @('aliases/alias[name="Legacy"]')
.\scripts\Network\Build-PfSenseConfig.ps1 -Path C:\FirewallBuilds\config.xml -Set $edits -Remove $remove -DryRun
```

- A scalar sets one leaf element. Multiple matches fail. A missing leaf can be created only when its last path step is a plain element name
  and its parent exists exactly once.
- `$true` writes an empty flag; `$false` or `$null` removes an element. Removing an already absent element through `-Set` is a no-op.
- An array replaces all sibling leaf elements with the selected name; partial selections fail. An empty array clears the list.
- `-Remove` deletes all matching elements, including sections. Each expression must match at least one element.
- Ordered dictionaries preserve edit order. Ordinary hashtables use sorted keys. The root and configuration version cannot be changed, and
  required `system` and `interfaces` sections must remain present.

Named settings run first, followed by `Set`, `Remove`, and host-key removal. The revision timestamp, description, and author are stamped
afterwards. Generic XPath is an advanced editing interface: syntactically valid edits may still produce an unusable firewall setup.

## Repeatable builds

Save this data file as `C:\FirewallBuilds\branch.psd1`. In Windows PowerShell 5.1, save files containing non-ASCII text as UTF-8 with BOM.

```powershell
@{
  Hostname          = 'fw-branch'
  Timezone          = 'Europe/Berlin'
  RemoveSshHostKeys = $true
  Interface         = @{
    lan = @{
      IPAddress = '10.20.0.1/24'
    }
  }
  DhcpServer        = @{
    lan = @{
      RangeFrom = '10.20.0.100'
      RangeTo   = '10.20.0.199'
    }
  }
  Set               = @{
    'system/webgui/max_procs' = 4
  }
}
```

```powershell
$build = @{
  Path       = 'C:\FirewallBuilds\config.xml'
  SpecPath   = 'C:\FirewallBuilds\branch.psd1'
  OutputPath = 'C:\FirewallBuilds\branch-02.xml'
  Hostname   = 'fw-branch-02'
}
.\scripts\Network\Build-PfSenseConfig.ps1 @build -DryRun
```

Specs accept edit parameter names only; file paths, `Force`, `DryRun`, and `PassThru` belong on the command line. Use actual booleans
(`$false`), not strings (`'false'`). The spec's complete edit sequence runs before the command-line sequence, so later edits win where they
can be applied. They do not undo an earlier section deletion or repair an invalid spec. Explicit `-RemoveSshHostKeys:$false` overrides
host-key removal requested by the spec.

## Review and restore

Generated files use tab indentation, LF line endings, and UTF-8 without BOM. Formatting is normalized; byte-for-byte preservation is not
guaranteed. The configuration version is preserved. The builder neither upgrades nor downgrades the schema and does not run pfSense's
complete configuration validation.

Before restoring, review interface assignments, gateway names, access settings, DHCP ranges, NAT, rules, VPN identities, and retained
secrets. Both source and output contain sensitive credentials. Logs and structured changes redact recognized secret paths, but this
name-based detection cannot identify every package-specific secret; protect transcripts and result files too.

Rehearse on an isolated firewall with console access. Follow Netgate's
[restore and compatibility instructions](https://docs.netgate.com/pfsense/en/latest/backup/restore.html): in **Diagnostics > Backup &
Restore**, select the generated file for a full restore. A full restore applies the configuration and reboots. Do not restore a newer
configuration revision to an older firewall. Verify access and connectivity after the reboot; retain the original backup for recovery.
