#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0.3.0' }

<#
.SYNOPSIS
  Builds a new pfSense configuration file from an existing configuration backup.
.DESCRIPTION
  Reads a full pfSense configuration backup (Diagnostics > Backup & Restore), applies
  the requested edits, and writes the result to a new file that can be restored on a
  pfSense firewall. The source file is never modified.

  Edits come from named parameters for common settings (system identity, DNS, time,
  webConfigurator, SSH, interfaces, and DHCP server ranges), from generic XPath -Set
  and -Remove expressions for anything else, and from an optional -SpecPath data file
  holding the same settings for repeatable builds. A spec file is applied first;
  command-line parameters are applied afterwards and win on conflicts.

  Every edit is validated before anything is written: an XPath that matches the wrong
  number of elements, an unknown interface, or an invalid address stops the build.
  Values are written the way pfSense writes them (tab indentation, LF line endings,
  UTF-8 without BOM, CDATA for description-style fields). Formatting is normalized;
  the configuration version is preserved. The <revision> entry records the build
  in the firewall's configuration history. This is not a full pfSense schema validator.

  Encrypted backups and partial (single-area) backups are rejected.
.PARAMETER Path
  Full pfSense configuration backup to read. Relative paths resolve from the current directory.
.PARAMETER OutputPath
  File to write. Defaults to <source name>.built.xml next to the source. Its folder must exist.
  An existing file is only replaced with -Force; the source file can never be the output.
.PARAMETER SpecPath
  PowerShell data file (.psd1) whose keys are this script's edit parameter names, for example
  @{ Hostname = 'fw-branch'; DnsServer = @('9.9.9.9') }. Applied before command-line parameters.
.PARAMETER Hostname
  System hostname without the domain (system/hostname).
.PARAMETER Domain
  System domain (system/domain).
.PARAMETER DnsServer
  System DNS servers, in order (system/dnsserver). Replacing the list resets each server's
  gateway binding to none, as General Setup does for new entries. An empty array removes all.
.PARAMETER DnsAllowOverride
  $true lets DHCP/PPP on WAN override the DNS servers (system/dnsallowoverride).
.PARAMETER Timezone
  Time zone name such as Europe/Berlin or Etc/UTC (system/timezone).
.PARAMETER TimeServer
  NTP servers or pools (system/timeservers). An empty array removes them.
.PARAMETER Language
  webConfigurator language such as en_US or de_DE (system/language).
.PARAMETER WebGuiProtocol
  webConfigurator protocol, http or https (system/webgui/protocol).
.PARAMETER WebGuiPort
  webConfigurator TCP port (system/webgui/port).
.PARAMETER WebGuiTheme
  webConfigurator theme file such as pfSense or pfSense-dark; .css is appended when missing
  (system/webgui/webguicss). The theme must exist on the target firewall.
.PARAMETER WebGuiDashboardColumns
  Number of dashboard columns, 1 to 6 (system/webgui/dashboardcolumns).
.PARAMETER WebGuiTopNavigation
  Scroll (default) or Fixed top navigation bar (system/webgui/webguifixedmenu).
.PARAMETER WebGuiHostnameInMenu
  Default, HostOnly, or Fqdn hostname in the menu bar (system/webgui/webguihostnamemenu).
.PARAMETER WebGuiLoginColor
  Login page color as six hexadecimal digits such as 1e3f75 (system/webgui/logincss).
.PARAMETER WebGuiLoginShowHost
  $true shows the hostname on the login page (system/webgui/loginshowhost).
.PARAMETER WebGuiAntiLockout
  $false disables the webConfigurator anti-lockout rule (system/webgui/noantilockout).
.PARAMETER WebGuiHttpRedirect
  $false disables the HTTP to HTTPS redirect (system/webgui/disablehttpredirect).
.PARAMETER WebGuiAlternateHostname
  Alternate hostnames accepted for DNS rebinding and HTTP_REFERER checks (system/webgui/althostnames).
  An empty array removes them.
.PARAMETER SshEnabled
  $true enables the Secure Shell server (system/ssh/enable).
.PARAMETER SshPort
  Secure Shell TCP port (system/ssh/port).
.PARAMETER SshAuthentication
  PasswordOrKey (default), Key, or PasswordAndKey (system/ssh/sshdkeyonly).
.PARAMETER Interface
  Interface settings keyed by internal name (wan, lan, opt1, ...) or description. Each value is a
  hashtable with any of: Description, Enable, Device, IPAddress, Subnet, Gateway, IPv6Address,
  SubnetV6, GatewayV6, MTU, MSS, SpoofMac, BlockPrivate, BlockBogons.
  IPAddress takes a static address (CIDR form also sets Subnet), dhcp, or none. IPv6Address takes
  a static address (CIDR form also sets SubnetV6), dhcp6, slaac, track6, or none. Gateway names a
  configured gateway. Enable, BlockPrivate, and BlockBogons take $true or $false. An empty string
  removes Description, Gateway, GatewayV6, MTU, MSS, or SpoofMac. The interface must already exist.
.PARAMETER DhcpServer
  DHCP server settings keyed by interface name or description. Each value is a hashtable with any
  of: Enable, RangeFrom, RangeTo, Gateway, Domain, DnsServer, DefaultLeaseTime, MaxLeaseTime.
  An empty string removes Gateway or Domain; DnsServer takes up to four addresses.
.PARAMETER Set
  XPath expressions mapped to values, relative to the <pfsense> root unless they start with /.
  A string or number sets the text of exactly one matched element; $true writes an empty
  (enabled) flag element; $false removes the element; an array replaces all matched sibling
  elements with one element per value. An expression that matches nothing creates its last
  element when that is a plain element name and its parent matches exactly one element.
  Use [ordered]@{} to control the order; plain hashtables are applied in sorted key order.
.PARAMETER Remove
  XPath expressions whose matched elements are removed. Each expression must match at least one element.
.PARAMETER RemoveSshHostKeys
  Remove the stored SSH host keys (<sshdata>), as clearing "Backup SSH keys" does when downloading
  a backup. The restored firewall then generates new host keys. Use this when cloning a firewall.
.PARAMETER Force
  Replace an existing output file.
.PARAMETER DryRun
  Validate and list every change without writing the output file.
.PARAMETER PassThru
  Return structured operation results.
.EXAMPLE
  PS> .\Build-PfSenseConfig.ps1 -Path .\config-fw-hq.example.com.xml -OutputPath .\fw-branch.xml -Hostname fw-branch -RemoveSshHostKeys -Interface @{ lan = @{ IPAddress = '10.20.0.1/24' } } -DhcpServer @{ lan = @{ RangeFrom = '10.20.0.100'; RangeTo = '10.20.0.199' } }
  Clones a firewall for a branch office with its own hostname, LAN subnet, and DHCP range.
.EXAMPLE
  PS> .\Build-PfSenseConfig.ps1 -Path .\config.xml -DnsServer 9.9.9.9, 149.112.112.112 -Timezone Europe/Berlin -WebGuiProtocol https -WebGuiPort 8443 -SshEnabled $true -SshAuthentication Key
  Writes config.built.xml with new DNS servers, time zone, webConfigurator port, and key-only SSH.
.EXAMPLE
  PS> .\Build-PfSenseConfig.ps1 -Path .\config.xml -Set ([ordered]@{ 'system/webgui/max_procs' = 4; 'aliases/alias[name="Admins"]/address' = '10.0.0.10 10.0.0.11' }) -Remove 'aliases/alias[name="Legacy"]'
  Changes arbitrary settings through XPath and removes an alias.
.EXAMPLE
  PS> .\Build-PfSenseConfig.ps1 -Path .\config.xml -SpecPath .\branch.psd1 -OutputPath .\fw-branch.xml -DryRun
  Lists every change a spec file would make without writing anything.
.LINK
  https://github.com/adnoctem/winkit
.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
  Server Core: supported - file operations only.
  SYSTEM-account execution: supported; no interaction or user profile is required.
  The output is as sensitive as the source: it holds password hashes, private keys, and VPN secrets.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Edit parameters are read through PSBoundParameters so that only supplied settings are applied.')]
[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $true, Position = 0)]
  [ValidateNotNullOrEmpty()]
  [string]
  $Path,

  [ValidateNotNullOrEmpty()]
  [string]
  $OutputPath,

  [ValidateNotNullOrEmpty()]
  [string]
  $SpecPath,

  [ValidatePattern('^[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?$')]
  [string]
  $Hostname,

  [ValidatePattern('^([A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)*[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?$')]
  [string]
  $Domain,

  [AllowEmptyCollection()]
  [string[]]
  $DnsServer,

  [Nullable[bool]]
  $DnsAllowOverride,

  [ValidatePattern('^[A-Za-z0-9_+-]+(/[A-Za-z0-9_+-]+)*$')]
  [string]
  $Timezone,

  [AllowEmptyCollection()]
  [ValidatePattern('^[A-Za-z0-9.:-]+$')]
  [string[]]
  $TimeServer,

  [ValidatePattern('^[a-z]{2,3}(_[A-Za-z]{2,4}){0,2}$')]
  [string]
  $Language,

  [ValidateSet('http', 'https')]
  [string]
  $WebGuiProtocol,

  [ValidateRange(1, 65535)]
  [int]
  $WebGuiPort,

  [ValidatePattern('^[A-Za-z0-9_.+-]+$')]
  [string]
  $WebGuiTheme,

  [ValidateRange(1, 6)]
  [int]
  $WebGuiDashboardColumns,

  [ValidateSet('Scroll', 'Fixed')]
  [string]
  $WebGuiTopNavigation,

  [ValidateSet('Default', 'HostOnly', 'Fqdn')]
  [string]
  $WebGuiHostnameInMenu,

  [ValidatePattern('^[0-9A-Fa-f]{6}$')]
  [string]
  $WebGuiLoginColor,

  [Nullable[bool]]
  $WebGuiLoginShowHost,

  [Nullable[bool]]
  $WebGuiAntiLockout,

  [Nullable[bool]]
  $WebGuiHttpRedirect,

  [AllowEmptyCollection()]
  [ValidatePattern('^[A-Za-z0-9.-]+$')]
  [string[]]
  $WebGuiAlternateHostname,

  [Nullable[bool]]
  $SshEnabled,

  [ValidateRange(1, 65535)]
  [int]
  $SshPort,

  [ValidateSet('PasswordOrKey', 'Key', 'PasswordAndKey')]
  [string]
  $SshAuthentication,

  [System.Collections.IDictionary]
  $Interface,

  [System.Collections.IDictionary]
  $DhcpServer,

  [System.Collections.IDictionary]
  $Set,

  [ValidateNotNullOrEmpty()]
  [string[]]
  $Remove,

  [switch]
  $RemoveSshHostKeys,

  [switch]
  $Force,

  [switch]
  $DryRun,

  [switch]
  $PassThru
)

Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.0 -Force

# -----------------------------------------------------------------------------
# Initialization
# -----------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

if ($DryRun) {
  $WhatIfPreference = $true
  Write-Log -Message "DRY RUN - no configuration file will be written`n" -Color Yellow
}

$_scriptParameters = $MyInvocation.MyCommand.Parameters
$_results = New-Object System.Collections.ArrayList
$_changes = New-Object System.Collections.ArrayList
$_warnings = New-Object System.Collections.ArrayList
$_document = $null
$_root = $null

# Settings accepted from parameters and spec files.
$_editParameterNames = @(
  'Hostname',
  'Domain',
  'DnsServer',
  'DnsAllowOverride',
  'Timezone',
  'TimeServer',
  'Language',
  'WebGuiProtocol',
  'WebGuiPort',
  'WebGuiTheme',
  'WebGuiDashboardColumns',
  'WebGuiTopNavigation',
  'WebGuiHostnameInMenu',
  'WebGuiLoginColor',
  'WebGuiLoginShowHost',
  'WebGuiAntiLockout',
  'WebGuiHttpRedirect',
  'WebGuiAlternateHostname',
  'SshEnabled',
  'SshPort',
  'SshAuthentication',
  'Interface',
  'DhcpServer',
  'Set',
  'Remove',
  'RemoveSshHostKeys'
)

$_interfaceSettingNames = @(
  'Description',
  'Enable',
  'Device',
  'IPAddress',
  'Subnet',
  'Gateway',
  'IPv6Address',
  'SubnetV6',
  'GatewayV6',
  'MTU',
  'MSS',
  'SpoofMac',
  'BlockPrivate',
  'BlockBogons'
)

$_dhcpSettingNames = @(
  'Enable',
  'RangeFrom',
  'RangeTo',
  'Gateway',
  'Domain',
  'DnsServer',
  'DefaultLeaseTime',
  'MaxLeaseTime'
)

# Elements pfSense serializes as CDATA (is_cdata_entity in xmlparse.inc).
$_cdataElementNames = @(
  'aclname',
  'auth_pass',
  'auth_prompt',
  'auth_user',
  'certca',
  'certname',
  'city',
  'common_name',
  'descr',
  'detail',
  'email',
  'encryption_password',
  'hint',
  'ldap_attr',
  'ldap_authcn',
  'ldap_basedn',
  'ldap_basedomain',
  'ldap_bind',
  'ldapbinddn',
  'ldapbindpass',
  'ldap_extended_query',
  'ldap_filter',
  'ldap_pam_groupdn',
  'ldap_pass',
  'ldap_user',
  'login_banner',
  'organization',
  'password',
  'proxypass',
  'proxyuser',
  'rangedescr',
  'state',
  'text',
  'username',
  'varusersusername',
  'varuserspassword'
)

# Values below these names are never echoed in logs or results.
$_secretNamePattern = '(?i)(pass|secret|psk|prv|hash|crypt|xmldata|private-?key|shared-?key|authorizedkeys|^key$|[_-]key$)'
$_secretSectionPattern = '^(sshdata|cert|ca|crl)$'
$_ipv4Pattern = '^(25[0-5]|2[0-4]\d|1?\d?\d)(\.(25[0-5]|2[0-4]\d|1?\d?\d)){3}$'

# -----------------------------------------------------------------------------
# Helpers: reading and describing configuration values
# -----------------------------------------------------------------------------

function Get-ConfigChildElement {
  param (
    [System.Xml.XmlNode]
    $Parent,

    [string]
    $Name
  )

  foreach ($_child in $Parent.ChildNodes) {
    if ($_child.NodeType -ne [System.Xml.XmlNodeType]::Element) {
      continue
    }

    if (-not $Name -or $_child.LocalName -ceq $Name) {
      $_child
    }
  }
}

function Test-ConfigSection {
  param (
    [System.Xml.XmlNode]
    $Node
  )

  foreach ($_child in $Node.ChildNodes) {
    if ($_child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
      return $true
    }
  }

  return $false
}

function Test-ConfigAttached {
  param (
    [System.Xml.XmlNode]
    $Node
  )

  $_current = $Node

  while ($null -ne $_current) {
    if ([object]::ReferenceEquals($_current, $_root)) {
      return $true
    }

    $_current = $_current.ParentNode
  }

  return $false
}

function Get-ConfigNodeValue {
  # pfSense decodes HTML entities in every value it parses (cData in xmlparse.inc).
  param (
    [System.Xml.XmlNode]
    $Node
  )

  if ($null -eq $Node) {
    return $null
  }

  if (Test-ConfigSection -Node $Node) {
    return '(section)'
  }

  return [System.Net.WebUtility]::HtmlDecode($Node.InnerText)
}

function Get-ConfigNodePath {
  # Describes an element as a path below <pfsense>, indexing repeated siblings.
  param (
    [System.Xml.XmlNode]
    $Node
  )

  $_segments = New-Object System.Collections.Generic.List[string]
  $_current = $Node

  while ($null -ne $_current -and -not [object]::ReferenceEquals($_current, $_root)) {
    $_parent = $_current.ParentNode

    if ($null -eq $_parent) {
      break
    }

    $_count = 0
    $_position = 0

    foreach ($_sibling in @(Get-ConfigChildElement -Parent $_parent -Name $_current.LocalName)) {
      $_count++

      if ([object]::ReferenceEquals($_sibling, $_current)) {
        $_position = $_count
      }
    }

    $_segment = $_current.LocalName

    if ($_count -gt 1) {
      $_segment = '{0}[{1}]' -f $_segment, $_position
    }

    $_segments.Insert(0, $_segment)
    $_current = $_parent
  }

  return $_segments -join '/'
}

function Test-SecretPath {
  param (
    [string]
    $Path
  )

  $_segments = @($Path -split '/' | ForEach-Object { $_ -replace '\[\d+\]$', '' })

  return @($_segments | Where-Object { $_ -match $_secretNamePattern -or $_ -match $_secretSectionPattern }).Count -gt 0
}

function Format-ConfigValue {
  param (
    [string]
    $Path,

    [object]
    $Value
  )

  if ($null -eq $Value) {
    return '(absent)'
  }

  if (Test-SecretPath -Path $Path) {
    return '(redacted)'
  }

  $_text = @($Value) -join ', '

  if ($_text -eq '') {
    return '(empty)'
  }

  return $_text
}

function Add-ConfigChange {
  param (
    [string]
    $Action,

    [string]
    $Path,

    [object]
    $Before,

    [object]
    $After
  )

  $_change = [PSCustomObject]@{
    Action = $Action
    Path   = $Path
    Before = Format-ConfigValue -Path $Path -Value $Before
    After  = Format-ConfigValue -Path $Path -Value $After
  }

  $null = $_changes.Add($_change)

  $_color = 'Green'

  if ($Action -eq 'Unchanged') {
    $_color = 'Gray'
  }

  Write-Log -Message ('  [{0}] {1}: {2} -> {3}' -f $Action, $Path, $_change.Before, $_change.After) -Color $_color
}

function Add-BuildWarning {
  param (
    [string]
    $Message
  )

  $null = $_warnings.Add($Message)
  Write-Warning $Message
}

# -----------------------------------------------------------------------------
# Helpers: editing the XML document
# -----------------------------------------------------------------------------

function Test-RemovalValue {
  param (
    [object]
    $Value
  )

  return ($null -eq $Value) -or ($Value -is [bool] -and -not $Value)
}

function ConvertTo-ConfigText {
  # $true is an enabled flag, which pfSense writes as an empty element.
  param (
    [object]
    $Value
  )

  if ($Value -is [bool]) {
    return ''
  }

  return [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Write-ConfigNodeText {
  # Mirrors pfSense's serializer: empty flags, escaped CDATA for description-style
  # fields, and plain text everywhere else.
  param (
    [System.Xml.XmlElement]
    $Node,

    [string]
    $Text
  )

  $Node.RemoveAll()
  $Node.IsEmpty = $false

  if ($Text -eq '') {
    return
  }

  $_useCdata = @($_cdataElementNames | Where-Object { $Node.LocalName.StartsWith($_, [System.StringComparison]::Ordinal) }).Count -gt 0

  if ($_useCdata) {
    $_escaped = [System.Security.SecurityElement]::Escape($Text)
    $null = $Node.AppendChild($_document.CreateCDataSection($_escaped))
  }
  else {
    $null = $Node.AppendChild($_document.CreateTextNode($Text))
  }
}

function Add-ConfigElement {
  param (
    [System.Xml.XmlNode]
    $Parent,

    [string]
    $Name,

    [System.Xml.XmlNode]
    $Before
  )

  $_element = $_document.CreateElement($Name)
  $_element.IsEmpty = $false

  if ($null -ne $Before) {
    $null = $Parent.InsertBefore($_element, $Before)
  }
  else {
    $null = $Parent.AppendChild($_element)
  }

  return $_element
}

function Add-ConfigLeaf {
  param (
    [System.Xml.XmlNode]
    $Parent,

    [string]
    $Name,

    [object]
    $Value
  )

  $_node = Add-ConfigElement -Parent $Parent -Name $Name
  $_text = ConvertTo-ConfigText -Value $Value

  Write-ConfigNodeText -Node $_node -Text $_text
  Add-ConfigChange -Action 'Add' -Path (Get-ConfigNodePath -Node $_node) -Before $null -After $_text
}

function Get-ConfigElement {
  # Walks a plain slash-separated path below <pfsense>, optionally creating sections.
  param (
    [string]
    $Path,

    [switch]
    $Create
  )

  $_current = $_root

  foreach ($_name in ($Path -split '/')) {
    $_next = @(Get-ConfigChildElement -Parent $_current -Name $_name) | Select-Object -First 1

    if ($null -eq $_next) {
      if (-not $Create) {
        return $null
      }

      $_next = Add-ConfigElement -Parent $_current -Name $_name
      Add-ConfigChange -Action 'Add' -Path (Get-ConfigNodePath -Node $_next) -Before $null -After '(section)'
    }

    $_current = $_next
  }

  return $_current
}

function Edit-ConfigNode {
  # Sets or removes one existing element and records the change.
  param (
    [System.Xml.XmlElement]
    $Node,

    [object]
    $Value
  )

  $_path = Get-ConfigNodePath -Node $Node
  $_before = Get-ConfigNodeValue -Node $Node

  if (Test-RemovalValue -Value $Value) {
    $null = $Node.ParentNode.RemoveChild($Node)
    Add-ConfigChange -Action 'Remove' -Path $_path -Before $_before -After $null
    return
  }

  if (Test-ConfigSection -Node $Node) {
    throw "$_path contains other elements; set its child elements or remove it instead."
  }

  $_text = ConvertTo-ConfigText -Value $Value

  if ($_before -ceq $_text) {
    Add-ConfigChange -Action 'Unchanged' -Path $_path -Before $_before -After $_text
    return
  }

  Write-ConfigNodeText -Node $Node -Text $_text
  Add-ConfigChange -Action 'Set' -Path $_path -Before $_before -After $_text
}

function Edit-ConfigPath {
  # Sets or removes the element at a plain path, creating missing sections.
  param (
    [string]
    $Path,

    [object]
    $Value
  )

  $_node = Get-ConfigElement -Path $Path

  if (Test-RemovalValue -Value $Value) {
    if ($null -eq $_node) {
      Add-ConfigChange -Action 'Unchanged' -Path $Path -Before $null -After $null
    }
    else {
      Edit-ConfigNode -Node $_node -Value $false
    }

    return
  }

  if ($null -ne $_node) {
    Edit-ConfigNode -Node $_node -Value $Value
    return
  }

  $_separator = $Path.LastIndexOf('/')
  $_parent = $_root

  if ($_separator -gt 0) {
    $_parent = Get-ConfigElement -Path $Path.Substring(0, $_separator) -Create
  }

  Add-ConfigLeaf -Parent $_parent -Name $Path.Substring($_separator + 1) -Value $Value
}

function Edit-ConfigList {
  # Replaces every <Name> child of Parent with one element per value, in place.
  param (
    [System.Xml.XmlNode]
    $Parent,

    [string]
    $Name,

    [object[]]
    $Values
  )

  $_existing = @(Get-ConfigChildElement -Parent $Parent -Name $Name)
  $_parentPath = Get-ConfigNodePath -Node $Parent
  $_path = $Name

  if ($_parentPath) {
    $_path = '{0}/{1}' -f $_parentPath, $Name
  }

  foreach ($_node in $_existing) {
    if (Test-ConfigSection -Node $_node) {
      throw "$_path contains sections; a list can only replace simple values."
    }
  }

  $_before = @($_existing | ForEach-Object { Get-ConfigNodeValue -Node $_ })
  $_after = @($Values | ForEach-Object { ConvertTo-ConfigText -Value $_ })
  $_beforeValue = $null
  $_afterValue = $null

  if ($_before.Count -gt 0) {
    $_beforeValue = $_before
  }

  if ($_after.Count -gt 0) {
    $_afterValue = $_after
  }

  if ($_before.Count -eq $_after.Count -and ($_before -join "`n") -ceq ($_after -join "`n")) {
    Add-ConfigChange -Action 'Unchanged' -Path $_path -Before $_beforeValue -After $_afterValue
    return
  }

  # New values take the position of the first existing element.
  $_anchor = $null

  if ($_existing.Count -gt 0) {
    $_anchor = $_existing[0]
  }

  foreach ($_text in $_after) {
    $_node = Add-ConfigElement -Parent $Parent -Name $Name -Before $_anchor
    Write-ConfigNodeText -Node $_node -Text $_text
  }

  foreach ($_node in $_existing) {
    $null = $Parent.RemoveChild($_node)
  }

  Add-ConfigChange -Action 'Set' -Path $_path -Before $_beforeValue -After $_afterValue
}

# -----------------------------------------------------------------------------
# Helpers: generic XPath edits
# -----------------------------------------------------------------------------

function Get-DictionaryKey {
  # Ordered dictionaries keep their order; others are applied in sorted key order.
  param (
    [System.Collections.IDictionary]
    $Dictionary
  )

  if ($Dictionary -is [System.Collections.Specialized.OrderedDictionary]) {
    return @($Dictionary.Keys)
  }

  return @($Dictionary.Keys | Sort-Object)
}

function Select-ConfigNode {
  param (
    [string]
    $XPath,

    [switch]
    $AllowRoot
  )

  try {
    if ($XPath.StartsWith('/')) {
      $_nodes = @($_document.SelectNodes($XPath))
    }
    else {
      $_nodes = @($_root.SelectNodes($XPath))
    }
  }
  catch {
    throw "Invalid XPath '$XPath': $($_.Exception.GetBaseException().Message)"
  }

  foreach ($_node in $_nodes) {
    if ($_node -isnot [System.Xml.XmlElement]) {
      throw "XPath '$XPath' must select elements, but it selected a $($_node.NodeType) node."
    }

    if (-not $AllowRoot -and [object]::ReferenceEquals($_node, $_root)) {
      throw "XPath '$XPath' selects the <pfsense> root, which cannot be changed."
    }
  }

  foreach ($_node in $_nodes) {
    $_node
  }
}

function Resolve-XPathTarget {
  # For an expression that matched nothing: the single parent and the element name to create.
  param (
    [string]
    $XPath
  )

  $_separator = $XPath.LastIndexOf('/')
  $_name = $XPath.Substring($_separator + 1)

  if ($_name -notmatch '^[A-Za-z_][A-Za-z0-9_.-]*$') {
    throw "XPath '$XPath' matches nothing, and its last step is not a plain element name that could be created."
  }

  if ($_separator -lt 0) {
    $_parent = $_root
  }
  else {
    $_parents = @(Select-ConfigNode -XPath $XPath.Substring(0, $_separator) -AllowRoot)

    if ($_parents.Count -ne 1) {
      throw "XPath '$XPath' matches nothing, and its parent matches $($_parents.Count) elements; the parent must match exactly one to create the element."
    }

    $_parent = $_parents[0]
  }

  return [PSCustomObject]@{
    Parent = $_parent
    Name   = $_name
  }
}

function Invoke-XPathListSetting {
  param (
    [string]
    $XPath,

    [object[]]
    $Nodes,

    [object[]]
    $Values
  )

  foreach ($_value in $Values) {
    if ($_value -is [System.Collections.IEnumerable] -and $_value -isnot [string]) {
      throw "-Set array for '$XPath' must contain simple values."
    }
  }

  if ($Nodes.Count -eq 0) {
    $_target = Resolve-XPathTarget -XPath $XPath
    Edit-ConfigList -Parent $_target.Parent -Name $_target.Name -Values $Values
    return
  }

  $_parent = $Nodes[0].ParentNode
  $_name = $Nodes[0].LocalName

  foreach ($_node in $Nodes) {
    if (-not [object]::ReferenceEquals($_node.ParentNode, $_parent) -or $_node.LocalName -cne $_name) {
      throw "XPath '$XPath' matches elements with different names or parents; an array value replaces sibling elements of one name."
    }
  }

  if (@(Get-ConfigChildElement -Parent $_parent -Name $_name).Count -ne $Nodes.Count) {
    throw "XPath '$XPath' matches only some <$_name> elements; an array value replaces all of them."
  }

  Edit-ConfigList -Parent $_parent -Name $_name -Values $Values
}

function Invoke-XPathSetting {
  param (
    [System.Collections.IDictionary]
    $Set
  )

  foreach ($_xpath in (Get-DictionaryKey -Dictionary $Set)) {
    $_value = $Set[$_xpath]

    if ($_value -is [System.Collections.IDictionary]) {
      throw "-Set value for '$_xpath' is a hashtable; use one XPath per element."
    }

    $_nodes = @(Select-ConfigNode -XPath $_xpath)

    if ($_value -is [System.Collections.IEnumerable] -and $_value -isnot [string]) {
      Invoke-XPathListSetting -XPath $_xpath -Nodes $_nodes -Values @($_value)
      continue
    }

    if ($_nodes.Count -eq 1) {
      Edit-ConfigNode -Node $_nodes[0] -Value $_value
      continue
    }

    if ($_nodes.Count -gt 1) {
      throw "XPath '$_xpath' matches $($_nodes.Count) elements; -Set changes exactly one. Add an index such as [1], or pass an array to replace them all."
    }

    if (Test-RemovalValue -Value $_value) {
      Add-ConfigChange -Action 'Unchanged' -Path $_xpath -Before $null -After $null
      continue
    }

    $_target = Resolve-XPathTarget -XPath $_xpath
    Add-ConfigLeaf -Parent $_target.Parent -Name $_target.Name -Value $_value
  }
}

function Invoke-XPathRemoval {
  param (
    [string[]]
    $Remove
  )

  foreach ($_xpath in $Remove) {
    $_nodes = @(Select-ConfigNode -XPath $_xpath)

    if ($_nodes.Count -eq 0) {
      throw "-Remove XPath '$_xpath' matches nothing."
    }

    # An element inside one removed earlier in this expression is already gone.
    foreach ($_node in $_nodes) {
      if (Test-ConfigAttached -Node $_node) {
        Edit-ConfigNode -Node $_node -Value $false
      }
    }
  }
}

# -----------------------------------------------------------------------------
# Helpers: validating setting values
# -----------------------------------------------------------------------------

function Assert-IPAddress {
  param (
    [string]
    $Value,

    [string]
    $Name,

    [ValidateSet('Any', 'IPv4', 'IPv6')]
    [string]
    $Family = 'Any'
  )

  $_address = $null
  $_valid = [System.Net.IPAddress]::TryParse($Value, [ref]$_address)

  # TryParse also accepts shorthand such as "10.1"; pfSense needs dotted quads.
  if ($_valid -and $_address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
    $_valid = ($Family -ne 'IPv6') -and ($Value -match $_ipv4Pattern)
  }
  elseif ($_valid) {
    $_valid = $Family -ne 'IPv4'
  }

  if (-not $_valid) {
    throw "$Name '$Value' is not a valid $Family address."
  }
}

function ConvertTo-SettingInteger {
  param (
    [object]
    $Value,

    [int]
    $Minimum,

    [int]
    $Maximum,

    [string]
    $Name
  )

  $_number = 0
  $_parsed = [int]::TryParse([string]$Value, [ref]$_number)

  if (-not $_parsed -or $_number -lt $Minimum -or $_number -gt $Maximum) {
    throw "$Name must be a whole number from $Minimum to $Maximum."
  }

  return $_number
}

function ConvertTo-SettingBoolean {
  # Strings are rejected: PowerShell would turn 'false' into $true.
  param (
    [object]
    $Value,

    [string]
    $Name
  )

  if ($Value -isnot [bool]) {
    throw "$Name must be `$true or `$false."
  }

  return $Value
}

function ConvertTo-OptionalText {
  # An empty string removes an optional element.
  param (
    [object]
    $Value
  )

  $_text = [string]$Value

  if ($_text -eq '') {
    return $false
  }

  return $_text
}

function Assert-SettingName {
  param (
    [System.Collections.IDictionary]
    $Settings,

    [string[]]
    $Valid,

    [string]
    $Context
  )

  foreach ($_key in $Settings.Keys) {
    if ($Valid -notcontains $_key) {
      throw "Unknown setting '$_key' for $Context. Valid settings: $($Valid -join ', ')."
    }
  }
}

function ConvertTo-InterfaceAddress {
  param (
    [object]
    $Value,

    [ValidateSet('IPv4', 'IPv6')]
    [string]
    $Family,

    [string]
    $Name
  )

  $_text = [string]$Value
  $_modes = @('dhcp')
  $_maximumPrefix = 32

  if ($Family -eq 'IPv6') {
    $_modes = @('dhcp6', 'slaac', 'track6')
    $_maximumPrefix = 128
  }

  if ($_text -eq '' -or $_text -eq 'none') {
    return [PSCustomObject]@{
      Address = $null
      Prefix  = $null
      Static  = $false
    }
  }

  if ($_modes -contains $_text) {
    return [PSCustomObject]@{
      Address = $_text.ToLowerInvariant()
      Prefix  = $null
      Static  = $false
    }
  }

  $_parts = $_text -split '/', 2
  Assert-IPAddress -Value $_parts[0] -Family $Family -Name $Name
  $_prefix = $null

  if ($_parts.Count -eq 2) {
    $_prefix = ConvertTo-SettingInteger -Value $_parts[1] -Minimum 1 -Maximum $_maximumPrefix -Name "$Name prefix"
  }

  return [PSCustomObject]@{
    Address = $_parts[0]
    Prefix  = $_prefix
    Static  = $true
  }
}

function ConvertTo-SpecValue {
  # Applies the parameter's own type and validation attributes to a spec file value.
  param (
    [string]
    $Name,

    [object]
    $Value
  )

  $_metadata = $_scriptParameters[$Name]
  $_type = $_metadata.ParameterType

  if ($_type -eq [System.Management.Automation.SwitchParameter] -or $_type -eq [Nullable[bool]]) {
    return (ConvertTo-SettingBoolean -Value $Value -Name "Spec setting '$Name'")
  }

  try {
    $_converted = [System.Management.Automation.LanguagePrimitives]::ConvertTo($Value, $_type)
  }
  catch {
    throw "Spec setting '$Name' cannot be converted to $($_type.Name): $($_.Exception.GetBaseException().Message)"
  }

  foreach ($_attribute in $_metadata.Attributes) {
    foreach ($_item in @($_converted)) {
      if ($_attribute -is [System.Management.Automation.ValidateSetAttribute] -and @($_attribute.ValidValues) -notcontains $_item) {
        throw "Spec setting '$Name' value '$_item' is not one of: $($_attribute.ValidValues -join ', ')."
      }

      if ($_attribute -is [System.Management.Automation.ValidateRangeAttribute] -and
        ($_item -lt $_attribute.MinRange -or $_item -gt $_attribute.MaxRange)) {
        throw "Spec setting '$Name' value '$_item' is outside $($_attribute.MinRange) to $($_attribute.MaxRange)."
      }

      if ($_attribute -is [System.Management.Automation.ValidatePatternAttribute] -and $_item -notmatch $_attribute.RegexPattern) {
        throw "Spec setting '$Name' value '$_item' is not in the expected format."
      }

      if ($_attribute -is [System.Management.Automation.ValidateNotNullOrEmptyAttribute] -and [string]::IsNullOrEmpty([string]$_item)) {
        throw "Spec setting '$Name' must not be empty."
      }
    }
  }

  return , $_converted
}

function Read-BuildSpec {
  param (
    [string]
    $Path
  )

  $_specPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

  if (-not (Test-Path -LiteralPath $_specPath -PathType Leaf)) {
    throw "Spec file '$_specPath' does not exist."
  }

  try {
    $_data = Import-PowerShellDataFile -LiteralPath $_specPath
  }
  catch {
    throw "Spec file '$_specPath' could not be read: $($_.Exception.Message)"
  }

  $_settings = @{}

  foreach ($_key in $_data.Keys) {
    $_name = $_editParameterNames | Where-Object { $_ -eq $_key } | Select-Object -First 1

    if (-not $_name) {
      throw "Spec file setting '$_key' is not an edit parameter. Valid settings: $($_editParameterNames -join ', ')."
    }

    $_settings[$_name] = ConvertTo-SpecValue -Name $_name -Value $_data[$_key]
  }

  return $_settings
}

# -----------------------------------------------------------------------------
# Helpers: named settings
# -----------------------------------------------------------------------------

function Edit-SystemDnsServer {
  # Mirrors General Setup: replaced servers get gateway "none", and bindings for
  # servers beyond the new list are removed.
  param (
    [string[]]
    $Server
  )

  foreach ($_address in $Server) {
    Assert-IPAddress -Value $_address -Name 'DnsServer'
  }

  if (@($Server | Select-Object -Unique).Count -ne $Server.Count) {
    throw 'DnsServer contains the same address more than once.'
  }

  $_system = Get-ConfigElement -Path 'system' -Create
  $_current = @(Get-ConfigChildElement -Parent $_system -Name 'dnsserver' | ForEach-Object { Get-ConfigNodeValue -Node $_ })
  Edit-ConfigList -Parent $_system -Name 'dnsserver' -Values $Server

  if (($_current -join "`n") -ceq ($Server -join "`n")) {
    return
  }

  for ($_index = 1; $_index -le $Server.Count; $_index++) {
    Edit-ConfigPath -Path "system/dns${_index}gw" -Value 'none'
    $_hostNode = Get-ConfigElement -Path "system/dns${_index}host"

    if ($null -ne $_hostNode) {
      Edit-ConfigNode -Node $_hostNode -Value $false
    }
  }

  foreach ($_node in @(Get-ConfigChildElement -Parent $_system)) {
    if ($_node.LocalName -match '^dns(\d+)(gw|host)$' -and [int]$Matches[1] -gt $Server.Count) {
      Edit-ConfigNode -Node $_node -Value $false
    }
  }
}

function Invoke-SystemSetting {
  param (
    [System.Collections.IDictionary]
    $Settings
  )

  # General Setup: identity, DNS, time, and language.
  if ($Settings.Contains('Hostname')) {
    Edit-ConfigPath -Path 'system/hostname' -Value $Settings['Hostname']
  }

  if ($Settings.Contains('Domain')) {
    Edit-ConfigPath -Path 'system/domain' -Value $Settings['Domain']
  }

  if ($Settings.Contains('DnsServer')) {
    Edit-SystemDnsServer -Server @($Settings['DnsServer'])
  }

  if ($Settings.Contains('DnsAllowOverride')) {
    Edit-ConfigPath -Path 'system/dnsallowoverride' -Value ([bool]$Settings['DnsAllowOverride'])
  }

  if ($Settings.Contains('Timezone')) {
    Edit-ConfigPath -Path 'system/timezone' -Value $Settings['Timezone']
  }

  if ($Settings.Contains('TimeServer')) {
    $_servers = @($Settings['TimeServer'] | ForEach-Object { $_.ToLowerInvariant() })
    Edit-ConfigPath -Path 'system/timeservers' -Value (ConvertTo-OptionalText -Value ($_servers -join ' '))
  }

  if ($Settings.Contains('Language')) {
    Edit-ConfigPath -Path 'system/language' -Value $Settings['Language']
  }

  # webConfigurator: General Setup and Advanced > Admin Access.
  if ($Settings.Contains('WebGuiProtocol')) {
    Edit-ConfigPath -Path 'system/webgui/protocol' -Value $Settings['WebGuiProtocol'].ToLowerInvariant()
  }

  if ($Settings.Contains('WebGuiPort')) {
    Edit-ConfigPath -Path 'system/webgui/port' -Value ([int]$Settings['WebGuiPort'])
  }

  if ($Settings.Contains('WebGuiTheme')) {
    $_theme = [string]$Settings['WebGuiTheme']

    if (-not $_theme.EndsWith('.css', [System.StringComparison]::OrdinalIgnoreCase)) {
      $_theme = '{0}.css' -f $_theme
    }

    Edit-ConfigPath -Path 'system/webgui/webguicss' -Value $_theme
  }

  if ($Settings.Contains('WebGuiDashboardColumns')) {
    Edit-ConfigPath -Path 'system/webgui/dashboardcolumns' -Value ([int]$Settings['WebGuiDashboardColumns'])
  }

  if ($Settings.Contains('WebGuiTopNavigation')) {
    $_menu = switch ($Settings['WebGuiTopNavigation']) {
      'Fixed' {
        'fixed'
      }
      default {
        $false
      }
    }

    Edit-ConfigPath -Path 'system/webgui/webguifixedmenu' -Value $_menu
  }

  if ($Settings.Contains('WebGuiHostnameInMenu')) {
    $_menu = switch ($Settings['WebGuiHostnameInMenu']) {
      'HostOnly' {
        'hostonly'
      }
      'Fqdn' {
        'fqdn'
      }
      default {
        $false
      }
    }

    Edit-ConfigPath -Path 'system/webgui/webguihostnamemenu' -Value $_menu
  }

  if ($Settings.Contains('WebGuiLoginColor')) {
    Edit-ConfigPath -Path 'system/webgui/logincss' -Value $Settings['WebGuiLoginColor'].ToLowerInvariant()
  }

  if ($Settings.Contains('WebGuiLoginShowHost')) {
    Edit-ConfigPath -Path 'system/webgui/loginshowhost' -Value ([bool]$Settings['WebGuiLoginShowHost'])
  }

  # pfSense stores these two as "disabled" flags.
  if ($Settings.Contains('WebGuiAntiLockout')) {
    Edit-ConfigPath -Path 'system/webgui/noantilockout' -Value (-not [bool]$Settings['WebGuiAntiLockout'])
  }

  if ($Settings.Contains('WebGuiHttpRedirect')) {
    Edit-ConfigPath -Path 'system/webgui/disablehttpredirect' -Value (-not [bool]$Settings['WebGuiHttpRedirect'])
  }

  if ($Settings.Contains('WebGuiAlternateHostname')) {
    $_hostnames = @($Settings['WebGuiAlternateHostname']) -join ' '
    Edit-ConfigPath -Path 'system/webgui/althostnames' -Value (ConvertTo-OptionalText -Value $_hostnames)
  }

  # Secure Shell: Advanced > Admin Access.
  if ($Settings.Contains('SshEnabled')) {
    $_enable = $false

    if ([bool]$Settings['SshEnabled']) {
      $_enable = 'enabled'
    }

    Edit-ConfigPath -Path 'system/ssh/enable' -Value $_enable
  }

  if ($Settings.Contains('SshPort')) {
    Edit-ConfigPath -Path 'system/ssh/port' -Value ([int]$Settings['SshPort'])
  }

  if ($Settings.Contains('SshAuthentication')) {
    $_keyOnly = switch ($Settings['SshAuthentication']) {
      'Key' {
        'enabled'
      }
      'PasswordAndKey' {
        'both'
      }
      default {
        $false
      }
    }

    Edit-ConfigPath -Path 'system/ssh/sshdkeyonly' -Value $_keyOnly
  }
}

function Resolve-InterfaceName {
  # Returns the internal interface name (wan, lan, optN) for a name or description.
  param (
    [string]
    $Name
  )

  $_interfaces = Get-ConfigElement -Path 'interfaces'

  if ($null -eq $_interfaces) {
    throw 'The source configuration has no <interfaces> section.'
  }

  $_candidates = @(Get-ConfigChildElement -Parent $_interfaces)

  foreach ($_candidate in $_candidates) {
    if ($_candidate.LocalName -eq $Name) {
      return $_candidate.LocalName
    }
  }

  $_described = @($_candidates | Where-Object { (Get-ConfigNodeValue -Node $_.SelectSingleNode('descr')) -eq $Name })

  if ($_described.Count -eq 1) {
    return $_described[0].LocalName
  }

  if ($_described.Count -gt 1) {
    throw "Interface description '$Name' is used by more than one interface; use the internal name instead."
  }

  $_available = @(
    $_candidates |
      ForEach-Object {
        $_description = Get-ConfigNodeValue -Node $_.SelectSingleNode('descr')

        if ($_description) {
          '{0} ({1})' -f $_.LocalName, $_description
        }
        else {
          $_.LocalName
        }
      }
  )

  throw "Interface '$Name' does not exist in the source configuration. Available: $($_available -join ', ')."
}

function Edit-InterfaceAddress {
  # Applies one address family's address and prefix settings to an interface.
  param (
    [System.Collections.IDictionary]
    $Settings,

    [string]
    $BasePath,

    [string]
    $Context,

    [ValidateSet('IPv4', 'IPv6')]
    [string]
    $Family
  )

  $_addressName = 'IPAddress'
  $_subnetName = 'Subnet'
  $_addressElement = 'ipaddr'
  $_subnetElement = 'subnet'
  $_maximumPrefix = 32

  if ($Family -eq 'IPv6') {
    $_addressName = 'IPv6Address'
    $_subnetName = 'SubnetV6'
    $_addressElement = 'ipaddrv6'
    $_subnetElement = 'subnetv6'
    $_maximumPrefix = 128
  }

  $_prefix = $null
  $_address = $null

  if ($Settings.Contains($_addressName)) {
    $_address = ConvertTo-InterfaceAddress -Value $Settings[$_addressName] -Family $Family -Name "$Context $_addressName"
    Edit-ConfigPath -Path "$BasePath/$_addressElement" -Value $_address.Address
    $_prefix = $_address.Prefix
  }

  if ($Settings.Contains($_subnetName)) {
    if ($null -ne $_address -and -not $_address.Static) {
      throw "$Context $_subnetName cannot be combined with a dynamic or unassigned $_addressName."
    }

    $_subnet = ConvertTo-SettingInteger -Value $Settings[$_subnetName] -Minimum 1 -Maximum $_maximumPrefix -Name "$Context $_subnetName"

    if ($null -ne $_prefix -and $_prefix -ne $_subnet) {
      throw "$Context $_addressName prefix /$_prefix conflicts with $_subnetName $_subnet."
    }

    $_prefix = $_subnet
  }

  if ($null -ne $_prefix) {
    Edit-ConfigPath -Path "$BasePath/$_subnetElement" -Value $_prefix
  }
  elseif ($null -ne $_address -and -not $_address.Static) {
    # Dynamic and unassigned addresses carry no prefix length.
    $_subnetNode = Get-ConfigElement -Path "$BasePath/$_subnetElement"

    if ($null -ne $_subnetNode) {
      Edit-ConfigNode -Node $_subnetNode -Value $false
    }
  }

  if ($null -ne $_address -and $_address.Static -and $null -eq (Get-ConfigElement -Path "$BasePath/$_subnetElement")) {
    throw "$Context static $_addressName needs a prefix length; use CIDR form or $_subnetName."
  }
}

function Invoke-InterfaceSetting {
  param (
    [System.Collections.IDictionary]
    $Interface
  )

  foreach ($_name in (Get-DictionaryKey -Dictionary $Interface)) {
    $_settings = $Interface[$_name]
    $_context = "interface '$_name'"

    if ($_settings -isnot [System.Collections.IDictionary]) {
      throw "Settings for $_context must be a hashtable, for example @{ IPAddress = '10.0.5.1/24' }."
    }

    Assert-SettingName -Settings $_settings -Valid $_interfaceSettingNames -Context $_context
    $_base = 'interfaces/{0}' -f (Resolve-InterfaceName -Name $_name)

    if ($_settings.Contains('Description')) {
      Edit-ConfigPath -Path "$_base/descr" -Value (ConvertTo-OptionalText -Value $_settings['Description'])
    }

    if ($_settings.Contains('Enable')) {
      Edit-ConfigPath -Path "$_base/enable" -Value (ConvertTo-SettingBoolean -Value $_settings['Enable'] -Name "$_context Enable")
    }

    if ($_settings.Contains('Device')) {
      Edit-ConfigPath -Path "$_base/if" -Value ([string]$_settings['Device'])
    }

    Edit-InterfaceAddress -Settings $_settings -BasePath $_base -Context $_context -Family IPv4
    Edit-InterfaceAddress -Settings $_settings -BasePath $_base -Context $_context -Family IPv6

    if ($_settings.Contains('Gateway')) {
      Edit-ConfigPath -Path "$_base/gateway" -Value (ConvertTo-OptionalText -Value $_settings['Gateway'])
    }

    if ($_settings.Contains('GatewayV6')) {
      Edit-ConfigPath -Path "$_base/gatewayv6" -Value (ConvertTo-OptionalText -Value $_settings['GatewayV6'])
    }

    foreach ($_sizeName in @('MTU', 'MSS')) {
      if (-not $_settings.Contains($_sizeName)) {
        continue
      }

      $_size = ConvertTo-OptionalText -Value $_settings[$_sizeName]

      if ($_size -isnot [bool]) {
        $_size = ConvertTo-SettingInteger -Value $_size -Minimum 1 -Maximum 65535 -Name "$_context $_sizeName"
      }

      Edit-ConfigPath -Path ('{0}/{1}' -f $_base, $_sizeName.ToLowerInvariant()) -Value $_size
    }

    if ($_settings.Contains('SpoofMac')) {
      $_mac = ConvertTo-OptionalText -Value $_settings['SpoofMac']

      if ($_mac -isnot [bool] -and $_mac -notmatch '^([0-9A-Fa-f]{2}:){5}[0-9A-Fa-f]{2}$') {
        throw "$_context SpoofMac '$_mac' must look like 00:11:22:33:44:55."
      }

      Edit-ConfigPath -Path "$_base/spoofmac" -Value $_mac
    }

    if ($_settings.Contains('BlockPrivate')) {
      Edit-ConfigPath -Path "$_base/blockpriv" -Value (ConvertTo-SettingBoolean -Value $_settings['BlockPrivate'] -Name "$_context BlockPrivate")
    }

    if ($_settings.Contains('BlockBogons')) {
      Edit-ConfigPath -Path "$_base/blockbogons" -Value (ConvertTo-SettingBoolean -Value $_settings['BlockBogons'] -Name "$_context BlockBogons")
    }
  }
}

function Invoke-DhcpSetting {
  param (
    [System.Collections.IDictionary]
    $DhcpServer
  )

  foreach ($_name in (Get-DictionaryKey -Dictionary $DhcpServer)) {
    $_settings = $DhcpServer[$_name]
    $_context = "DHCP server '$_name'"

    if ($_settings -isnot [System.Collections.IDictionary]) {
      throw "Settings for $_context must be a hashtable, for example @{ RangeFrom = '10.0.5.100'; RangeTo = '10.0.5.199' }."
    }

    Assert-SettingName -Settings $_settings -Valid $_dhcpSettingNames -Context $_context
    $_base = 'dhcpd/{0}' -f (Resolve-InterfaceName -Name $_name)

    if ($_settings.Contains('Enable')) {
      Edit-ConfigPath -Path "$_base/enable" -Value (ConvertTo-SettingBoolean -Value $_settings['Enable'] -Name "$_context Enable")
    }

    foreach ($_bound in @('From', 'To')) {
      $_settingName = 'Range{0}' -f $_bound

      if ($_settings.Contains($_settingName)) {
        Assert-IPAddress -Value $_settings[$_settingName] -Family IPv4 -Name "$_context $_settingName"
        Edit-ConfigPath -Path ('{0}/range/{1}' -f $_base, $_bound.ToLowerInvariant()) -Value ([string]$_settings[$_settingName])
      }
    }

    if ($_settings.Contains('Gateway')) {
      $_gateway = ConvertTo-OptionalText -Value $_settings['Gateway']

      if ($_gateway -isnot [bool]) {
        Assert-IPAddress -Value $_gateway -Family IPv4 -Name "$_context Gateway"
      }

      Edit-ConfigPath -Path "$_base/gateway" -Value $_gateway
    }

    if ($_settings.Contains('Domain')) {
      Edit-ConfigPath -Path "$_base/domain" -Value (ConvertTo-OptionalText -Value $_settings['Domain'])
    }

    if ($_settings.Contains('DnsServer')) {
      $_servers = @($_settings['DnsServer'] | Where-Object { $_ })

      if ($_servers.Count -gt 4) {
        throw "$_context DnsServer accepts up to four addresses."
      }

      foreach ($_server in $_servers) {
        Assert-IPAddress -Value $_server -Family IPv4 -Name "$_context DnsServer"
      }

      Edit-ConfigList -Parent (Get-ConfigElement -Path $_base -Create) -Name 'dnsserver' -Values $_servers
    }

    foreach ($_leaseName in @('DefaultLeaseTime', 'MaxLeaseTime')) {
      if ($_settings.Contains($_leaseName)) {
        $_seconds = ConvertTo-SettingInteger -Value $_settings[$_leaseName] -Minimum 60 -Maximum ([int]::MaxValue) -Name "$_context $_leaseName"
        Edit-ConfigPath -Path ('{0}/{1}' -f $_base, $_leaseName.ToLowerInvariant()) -Value $_seconds
      }
    }
  }
}

function Invoke-SshHostKeyRemoval {
  $_nodes = @(Get-ConfigChildElement -Parent $_root -Name 'sshdata')

  if ($_nodes.Count -eq 0) {
    Add-ConfigChange -Action 'Unchanged' -Path 'sshdata' -Before $null -After $null
    return
  }

  foreach ($_node in $_nodes) {
    Edit-ConfigNode -Node $_node -Value $false
  }
}

function Invoke-ConfigEdit {
  param (
    [System.Collections.IDictionary]
    $Settings,

    [string]
    $Origin
  )

  Write-Log -Message "Applying settings from $Origin" -Color Cyan

  Invoke-SystemSetting -Settings $Settings

  if ($Settings.Contains('Interface')) {
    Invoke-InterfaceSetting -Interface $Settings['Interface']
  }

  if ($Settings.Contains('DhcpServer')) {
    Invoke-DhcpSetting -DhcpServer $Settings['DhcpServer']
  }

  if ($Settings.Contains('Set')) {
    Invoke-XPathSetting -Set $Settings['Set']
  }

  if ($Settings.Contains('Remove')) {
    Invoke-XPathRemoval -Remove $Settings['Remove']
  }

  if ($Settings.Contains('RemoveSshHostKeys') -and [bool]$Settings['RemoveSshHostKeys']) {
    Invoke-SshHostKeyRemoval
  }
}

# -----------------------------------------------------------------------------
# Helpers: post-edit checks, loading, and writing
# -----------------------------------------------------------------------------

function ConvertTo-IPv4Number {
  param (
    [string]
    $Address
  )

  $_number = [double]0

  foreach ($_octet in $Address.Split('.')) {
    $_number = ($_number * 256) + [int]$_octet
  }

  return $_number
}

function Test-IPv4InSubnet {
  param (
    [string]
    $Address,

    [string]
    $Network,

    [int]
    $Prefix
  )

  $_blockSize = [math]::Pow(2, 32 - $Prefix)
  $_addressBlock = [math]::Floor((ConvertTo-IPv4Number -Address $Address) / $_blockSize)
  $_networkBlock = [math]::Floor((ConvertTo-IPv4Number -Address $Network) / $_blockSize)

  return $_addressBlock -eq $_networkBlock
}

function Test-DhcpRange {
  # Warns when an enabled DHCP range falls outside its interface's static IPv4 subnet.
  $_dhcpd = Get-ConfigElement -Path 'dhcpd'

  if ($null -eq $_dhcpd) {
    return
  }

  foreach ($_server in @(Get-ConfigChildElement -Parent $_dhcpd)) {
    $_id = $_server.LocalName

    if ($null -eq (Get-ConfigElement -Path "dhcpd/$_id/enable")) {
      continue
    }

    $_address = Get-ConfigNodeValue -Node (Get-ConfigElement -Path "interfaces/$_id/ipaddr")
    $_prefix = Get-ConfigNodeValue -Node (Get-ConfigElement -Path "interfaces/$_id/subnet")

    if ($_address -notmatch $_ipv4Pattern -or $_prefix -notmatch '^\d+$') {
      continue
    }

    foreach ($_bound in @('from', 'to')) {
      $_value = Get-ConfigNodeValue -Node (Get-ConfigElement -Path "dhcpd/$_id/range/$_bound")

      if ($_value -match $_ipv4Pattern -and -not (Test-IPv4InSubnet -Address $_value -Network $_address -Prefix ([int]$_prefix))) {
        Add-BuildWarning -Message "DHCP range '$_bound' address $_value on $_id is outside the interface subnet $_address/$_prefix."
      }
    }
  }
}

function Assert-PfSenseDocument {
  param (
    [System.Xml.XmlDocument]
    $Document
  )

  if ($Document.DocumentElement.Name -cne 'pfsense' -or $Document.DocumentElement.NamespaceURI) {
    throw 'Expected a <pfsense> root without a namespace. Partial backups of a single area are not supported.'
  }

  foreach ($_section in @('version', 'system', 'interfaces')) {
    $_nodes = @($Document.DocumentElement.SelectNodes($_section))

    if ($_nodes.Count -ne 1 -or -not $_nodes[0].HasChildNodes) {
      throw "A full pfSense configuration needs exactly one nonempty <$_section> element."
    }
  }

  if ($Document.DocumentElement.SelectSingleNode('version').InnerText -notmatch '^\d+(\.\d+)*$') {
    throw 'The pfSense configuration version must be numeric.'
  }
}

function Read-PfSenseConfig {
  # Loads a full, unencrypted configuration with DTD processing and external resolution disabled.
  param (
    [string]
    $Path
  )

  $_text = [System.IO.File]::ReadAllText($Path)

  if ($_text -match '-{4} BEGIN config\.xml -{4}') {
    throw "'$Path' is an encrypted pfSense backup. Encrypted backups are not supported; download the configuration without encryption."
  }

  if (-not $_text.TrimStart().StartsWith('<')) {
    throw "'$Path' is not an XML pfSense configuration."
  }

  $_readerSettings = New-Object System.Xml.XmlReaderSettings
  $_readerSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
  $_readerSettings.XmlResolver = $null
  $_readerSettings.IgnoreWhitespace = $true

  $_document = New-Object System.Xml.XmlDocument
  $_document.PreserveWhitespace = $false
  $_document.XmlResolver = $null
  $_reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($_text)), $_readerSettings)

  try {
    $_document.Load($_reader)
  }
  catch {
    throw "'$Path' is not well-formed XML: $($_.Exception.GetBaseException().Message)"
  }
  finally {
    $_reader.Dispose()
  }

  Assert-PfSenseDocument -Document $_document

  return $_document
}

function Write-PfSenseConfig {
  # Writes pfSense's own layout through a temporary file that must read back first.
  param (
    [System.Xml.XmlDocument]
    $Document,

    [string]
    $Path,

    [string]
    $ExpectedVersion,

    [switch]
    $Force
  )

  $_encoding = New-Object System.Text.UTF8Encoding($false)
  $_writerSettings = New-Object System.Xml.XmlWriterSettings
  $_writerSettings.Encoding = $_encoding
  $_writerSettings.Indent = $true
  $_writerSettings.IndentChars = "`t"
  $_writerSettings.NewLineChars = "`n"
  $_writerSettings.NewLineHandling = [System.Xml.NewLineHandling]::Replace
  $_writerSettings.OmitXmlDeclaration = $true

  $_stream = New-Object System.IO.MemoryStream
  $_writer = [System.Xml.XmlWriter]::Create($_stream, $_writerSettings)

  try {
    $Document.DocumentElement.WriteTo($_writer)
  }
  finally {
    $_writer.Dispose()
  }

  $_text = "<?xml version=`"1.0`"?>`n{0}`n" -f $_encoding.GetString($_stream.ToArray())
  $_temporary = '{0}.{1}.tmp' -f $Path, [guid]::NewGuid().ToString('N')

  try {
    [System.IO.File]::WriteAllText($_temporary, $_text, $_encoding)

    $_check = Read-PfSenseConfig -Path $_temporary
    $_checkVersion = Get-ConfigNodeValue -Node $_check.DocumentElement.SelectSingleNode('version')

    if ($_checkVersion -cne $ExpectedVersion) {
      throw "The written configuration did not read back with version '$ExpectedVersion'."
    }

    # PowerShell passes $null to a string argument as ''; File.Replace needs a real null.
    if ($Force -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
      [System.IO.File]::Replace($_temporary, $Path, [NullString]::Value)
    }
    else {
      [System.IO.File]::Move($_temporary, $Path)
    }
  }
  finally {
    if (Test-Path -LiteralPath $_temporary) {
      Remove-Item -LiteralPath $_temporary -Force
    }
  }
}

# -----------------------------------------------------------------------------
# Validate the files and collect settings before reading the configuration.
# -----------------------------------------------------------------------------

$_sourcePath = $null
$_outputPath = $null
$_configVersion = $null
$_written = $false
$_failure = $null

try {
  $_sourcePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

  if (-not (Test-Path -LiteralPath $_sourcePath -PathType Leaf)) {
    throw "Source configuration '$_sourcePath' does not exist."
  }

  if ($OutputPath) {
    $_outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
  }
  else {
    $_outputName = '{0}.built.xml' -f [System.IO.Path]::GetFileNameWithoutExtension($_sourcePath)
    $_outputPath = Join-Path -Path (Split-Path -Path $_sourcePath -Parent) -ChildPath $_outputName
  }

  if ([string]::Equals($_outputPath, $_sourcePath, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputPath must differ from the source file; the source is never modified.'
  }

  $_outputFolder = Split-Path -Path $_outputPath -Parent

  if (-not (Test-Path -LiteralPath $_outputFolder -PathType Container)) {
    throw "Output folder '$_outputFolder' does not exist."
  }

  if (Test-Path -LiteralPath $_outputPath -PathType Container) {
    throw "OutputPath '$_outputPath' is a folder; specify a file name."
  }

  if ((Test-Path -LiteralPath $_outputPath) -and -not $Force) {
    throw "Output file '$_outputPath' already exists. Use -Force to replace it."
  }

  # A spec file is applied first, so explicitly supplied parameters win.
  $_specSettings = @{}

  if ($SpecPath) {
    $_specSettings = Read-BuildSpec -Path $SpecPath
  }

  $_parameterSettings = @{}

  foreach ($_name in $_editParameterNames) {
    if ($PSBoundParameters.ContainsKey($_name)) {
      $_parameterSettings[$_name] = $PSBoundParameters[$_name]
    }
  }

  # An explicit false switch must also prevent removal during the spec phase.
  if ($_parameterSettings.Contains('RemoveSshHostKeys')) {
    $_specSettings.Remove('RemoveSshHostKeys')
  }

  # ---------------------------------------------------------------------------
  # Load the source and apply every edit in memory.
  # ---------------------------------------------------------------------------

  Write-Log -Message "Reading pfSense configuration $_sourcePath" -Color Cyan
  $_document = Read-PfSenseConfig -Path $_sourcePath
  $_root = $_document.DocumentElement
  $_configVersion = Get-ConfigNodeValue -Node (Get-ConfigElement -Path 'version')
  $_originalHostname = Get-ConfigNodeValue -Node (Get-ConfigElement -Path 'system/hostname')
  Write-Log -Message "Configuration version $_configVersion, hostname $_originalHostname" -Color Gray

  if ($_specSettings.Count -gt 0) {
    Invoke-ConfigEdit -Settings $_specSettings -Origin "spec file $SpecPath"
  }

  if ($_parameterSettings.Count -gt 0) {
    Invoke-ConfigEdit -Settings $_parameterSettings -Origin 'parameters'
  }

  if ($_specSettings.Count -eq 0 -and $_parameterSettings.Count -eq 0) {
    Write-Log -Message 'No edits requested; only the revision entry will differ from the source.' -Color Yellow
  }

  # ---------------------------------------------------------------------------
  # Check the result, stamp the revision, and write the new file.
  # ---------------------------------------------------------------------------

  Assert-PfSenseDocument -Document $_document

  if ((Get-ConfigNodeValue -Node (Get-ConfigElement -Path 'version')) -cne $_configVersion) {
    throw 'The configuration version cannot be changed; this builder does not migrate pfSense schemas.'
  }

  Test-DhcpRange

  $_finalHostname = Get-ConfigNodeValue -Node (Get-ConfigElement -Path 'system/hostname')
  $_hasHostKeys = @(Get-ConfigChildElement -Parent $_root -Name 'sshdata').Count -gt 0

  if ($_finalHostname -cne $_originalHostname -and $_hasHostKeys) {
    Add-BuildWarning -Message 'The hostname changed but the SSH host keys are kept, so both firewalls would share them. Use -RemoveSshHostKeys when cloning.'
  }

  Write-Log -Message 'Recording the build in the revision entry' -Color Cyan
  Edit-ConfigPath -Path 'revision/time' -Value ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
  Edit-ConfigPath -Path 'revision/description' -Value ('Built by winkit Build-PfSenseConfig from {0}' -f (Split-Path -Path $_sourcePath -Leaf))
  Edit-ConfigPath -Path 'revision/username' -Value ('{0}@{1} (winkit)' -f $env:USERNAME, $env:COMPUTERNAME)

  if ($PSCmdlet.ShouldProcess($_outputPath, 'Write pfSense configuration')) {
    Write-PfSenseConfig -Document $_document -Path $_outputPath -ExpectedVersion $_configVersion -Force:$Force
    $_written = $true
  }
}
catch {
  $_failure = $_.Exception.Message
}

# -----------------------------------------------------------------------------
# Report every edit and the build outcome.
# -----------------------------------------------------------------------------

$_preview = [bool]$WhatIfPreference
$_editStatus = 'Skipped'

if ($_written) {
  $_editStatus = 'Completed'
}
elseif ($_preview -and -not $_failure) {
  $_editStatus = 'Preview'
}

foreach ($_change in $_changes) {
  $_isChange = $_change.Action -ne 'Unchanged'
  $_status = $_editStatus

  if (-not $_isChange) {
    $_status = 'Skipped'
  }

  Add-OperationResult `
    -Results $_results `
    -Target $_change.Path `
    -Source 'PfSenseConfig' `
    -Action $_change.Action `
    -Status $_status `
    -Before $_change.Before `
    -After $_change.After `
    -Changed ($_isChange -and $_written) `
    -AlreadyCompliant (-not $_isChange)
}

$_target = $_outputPath

if (-not $_target) {
  $_target = $Path
}

foreach ($_warning in $_warnings) {
  Add-OperationResult `
    -Results $_results `
    -Target $_target `
    -Source 'PfSenseConfig' `
    -Action 'Check' `
    -Status 'Warn' `
    -Detail $_warning
}

$_changeCount = @($_changes | Where-Object { $_.Action -ne 'Unchanged' }).Count
$_property = @{
  SourcePath    = $_sourcePath
  OutputPath    = $_outputPath
  ConfigVersion = $_configVersion
  ChangeCount   = $_changeCount
}

$_build = @{
  Results  = $_results
  Target   = $_target
  Source   = 'PfSenseConfig'
  Action   = 'Build'
  Changed  = $_written
  Property = $_property
}

if ($_failure) {
  $_build.Status = 'Failed'
  $_build.Detail = 'The configuration was not written.'
  $_build.ErrorMessage = $_failure
  $_color = 'Red'
}
elseif ($_written) {
  $_build.Status = 'Completed'
  $_build.Detail = "Wrote $_changeCount change(s) to $_outputPath."
  $_color = 'Green'
}
elseif ($_preview) {
  $_build.Status = 'Preview'
  $_build.Detail = "Would write $_changeCount change(s) to $_outputPath."
  $_color = 'Yellow'
}
else {
  $_build.Status = 'Skipped'
  $_build.Detail = 'Writing the configuration was declined.'
  $_color = 'Yellow'
}

Add-OperationResult @_build

if ($_failure) {
  Write-Log -Message "Build failed: $_failure" -Color $_color
}
else {
  Write-Log -Message $_build.Detail -Color $_color
}

if ($PassThru -or $_preview) {
  $_results
}

if ($_failure) {
  exit 1
}
