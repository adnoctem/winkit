#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Tier-0 logic tests for Build-PfSenseConfig. Every case runs the script end to end
# against a synthetic config.xml under TestDrive, written in pfSense's own layout
# (tabs, LF, no BOM, CDATA descriptions). No firewall or elevation is needed.

BeforeAll {
  $script:BuildScript = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '..\..\scripts\Network\Build-PfSenseConfig.ps1'))
  $script:Utf8 = New-Object System.Text.UTF8Encoding($false)

  $script:SourceLines = @(
    '<?xml version="1.0"?>',
    '<pfsense>',
    "`t<version>23.3</version>",
    "`t<system>",
    "`t`t<hostname>fw-hq</hostname>",
    "`t`t<domain>example.com</domain>",
    "`t`t<dnsallowoverride></dnsallowoverride>",
    "`t`t<user>",
    "`t`t`t<name>admin</name>",
    "`t`t`t<descr><![CDATA[System Administrator]]></descr>",
    "`t`t`t<bcrypt-hash>`$2y`$10`$abcdefghijklmnopqrstuv</bcrypt-hash>",
    "`t`t`t<authorizedkeys>c3NoLWVkMjU1MTkgQUFBQQ==</authorizedkeys>",
    "`t`t</user>",
    "`t`t<webgui>",
    "`t`t`t<protocol>https</protocol>",
    "`t`t`t<webguicss>pfSense.css</webguicss>",
    "`t`t`t<max_procs>2</max_procs>",
    "`t`t</webgui>",
    "`t`t<ssh>",
    "`t`t`t<enable>enabled</enable>",
    "`t`t</ssh>",
    "`t`t<timezone>Etc/UTC</timezone>",
    "`t`t<timeservers>2.pfsense.pool.ntp.org</timeservers>",
    "`t`t<dnsserver>1.1.1.1</dnsserver>",
    "`t`t<dnsserver>8.8.8.8</dnsserver>",
    "`t`t<dns1gw>none</dns1gw>",
    "`t`t<dns2gw>WAN_DHCP</dns2gw>",
    "`t`t<language>en_US</language>",
    "`t</system>",
    "`t<interfaces>",
    "`t`t<wan>",
    "`t`t`t<enable></enable>",
    "`t`t`t<if>vtnet0</if>",
    "`t`t`t<descr><![CDATA[WAN]]></descr>",
    "`t`t`t<ipaddr>dhcp</ipaddr>",
    "`t`t`t<blockpriv></blockpriv>",
    "`t`t`t<blockbogons></blockbogons>",
    "`t`t</wan>",
    "`t`t<lan>",
    "`t`t`t<enable></enable>",
    "`t`t`t<if>vtnet1</if>",
    "`t`t`t<descr><![CDATA[LAN]]></descr>",
    "`t`t`t<ipaddr>192.168.1.1</ipaddr>",
    "`t`t`t<subnet>24</subnet>",
    "`t`t</lan>",
    "`t`t<opt1>",
    "`t`t`t<if>vtnet2</if>",
    "`t`t`t<descr><![CDATA[DMZ]]></descr>",
    "`t`t`t<ipaddr>172.16.0.1</ipaddr>",
    "`t`t`t<subnet>24</subnet>",
    "`t`t</opt1>",
    "`t</interfaces>",
    "`t<dhcpd>",
    "`t`t<lan>",
    "`t`t`t<enable></enable>",
    "`t`t`t<range>",
    "`t`t`t`t<from>192.168.1.100</from>",
    "`t`t`t`t<to>192.168.1.199</to>",
    "`t`t`t</range>",
    "`t`t</lan>",
    "`t</dhcpd>",
    "`t<aliases>",
    "`t`t<alias>",
    "`t`t`t<name>Admins</name>",
    "`t`t`t<address>10.0.0.10</address>",
    "`t`t</alias>",
    "`t`t<alias>",
    "`t`t`t<name>Legacy</name>",
    "`t`t`t<address>10.0.0.99</address>",
    "`t`t</alias>",
    "`t</aliases>",
    "`t<revision>",
    "`t`t<time>1700000000</time>",
    "`t`t<description><![CDATA[admin@10.0.0.2 (Local Database): Saved]]></description>",
    "`t`t<username><![CDATA[admin@10.0.0.2 (Local Database)]]></username>",
    "`t</revision>",
    "`t<sshdata>",
    "`t`t<sshkeyfile>",
    "`t`t`t<filename>ssh_host_ed25519_key</filename>",
    "`t`t`t<xmldata>S0VZREFUQQ==</xmldata>",
    "`t`t</sshkeyfile>",
    "`t</sshdata>",
    '</pfsense>'
  )
  $script:SourceText = ($script:SourceLines -join "`n") + "`n"

  function Invoke-Build {
    # Runs the script with -PassThru; returns results and the exit code.
    param (
      [hashtable]
      $Parameters
    )

    $global:LASTEXITCODE = 0
    $_results = & $script:BuildScript @Parameters -PassThru 6>$null 3>$null

    return [PSCustomObject]@{
      Results  = @($_results)
      ExitCode = $global:LASTEXITCODE
    }
  }

  function Read-Output {
    param (
      [string]
      $Path
    )

    $_document = New-Object System.Xml.XmlDocument
    $_document.Load($Path)
    return $_document.DocumentElement
  }

  function Get-TextValue {
    param (
      [System.Xml.XmlNode]
      $Root,

      [string]
      $XPath
    )

    return @($Root.SelectNodes($XPath) | ForEach-Object { $_.InnerText })
  }

  function ConvertTo-RevisionFreeText {
    param (
      [string]
      $Text
    )

    return [regex]::Replace($Text, "(?s)`t<revision>.*?</revision>`n", '')
  }
}

Describe 'Build-PfSenseConfig script conventions' {
  It 'pins AdNoctem.Substrate.PowerShell and documents .SYNOPSIS and a complete .NOTES block' {
    $_header = Get-Content -LiteralPath $script:BuildScript -TotalCount 3
    $_header -match "^#Requires -Modules @\{ ModuleName = 'AdNoctem.Substrate.PowerShell'; RequiredVersion = '0\.2\.1' \}" | Should -Not -BeNullOrEmpty
    $_content = Get-Content -LiteralPath $script:BuildScript -Raw
    $_content -match 'Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.2.1 -Force' | Should -BeTrue
    (Get-Help $script:BuildScript).Synopsis | Should -Not -Match 'Build-PfSenseConfig\.ps1'
    $_content -match 'Server Core:' | Should -BeTrue
    $_content -match 'SYSTEM-account execution:' | Should -BeTrue
  }
}

Describe 'Build-PfSenseConfig' {
  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Root -Force
    $script:Source = Join-Path $script:Root 'config-fw-hq.xml'
    $script:Output = Join-Path $script:Root 'out.xml'
    [System.IO.File]::WriteAllText($script:Source, $script:SourceText, $script:Utf8)
    $script:SourceHash = (Get-FileHash -LiteralPath $script:Source).Hash
  }

  AfterEach {
    (Get-FileHash -LiteralPath $script:Source).Hash | Should -Be $script:SourceHash -Because 'the source file is never modified'
  }

  Context 'file handling' {
    It 'writes a .built.xml file in pfSense layout that differs only in the revision when nothing is edited' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source }
      $_run.ExitCode | Should -Be 0

      $_output = Join-Path $script:Root 'config-fw-hq.built.xml'
      $_bytes = [System.IO.File]::ReadAllBytes($_output)
      $_bytes[0] | Should -Be ([byte][char]'<') -Because 'the output has no BOM'
      $_text = $script:Utf8.GetString($_bytes)
      $_text.Contains("`r") | Should -BeFalse
      $_text | Should -BeLike "<?xml version=`"1.0`"?>`n<pfsense>`n*"
      ConvertTo-RevisionFreeText -Text $_text | Should -BeExactly (ConvertTo-RevisionFreeText -Text $script:SourceText)
      $_text | Should -Match '<description><!\[CDATA\[Built by winkit'

      $_root = Read-Output -Path $_output
      Get-TextValue -Root $_root -XPath 'revision/description' | Should -Be 'Built by winkit Build-PfSenseConfig from config-fw-hq.xml'
      [long](@(Get-TextValue -Root $_root -XPath 'revision/time')[0]) | Should -BeGreaterThan 1700000000
      ($_run.Results | Where-Object Action -EQ 'Build').Status | Should -Be 'Completed'
    }

    It 'refuses to replace an existing output without -Force and replaces it with -Force' {
      [System.IO.File]::WriteAllText($script:Output, 'old', $script:Utf8)

      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output }
      $_run.ExitCode | Should -Be 1
      [System.IO.File]::ReadAllText($script:Output) | Should -Be 'old'

      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Force = $true }
      $_run.ExitCode | Should -Be 0
      Get-TextValue -Root (Read-Output -Path $script:Output) -XPath 'system/hostname' | Should -Be 'fw-hq'
    }

    It 'never writes over the source, even with -Force' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Source; Force = $true; Hostname = 'other' }
      $_run.ExitCode | Should -Be 1
    }

    It 'rejects encrypted backups' {
      $_encrypted = Join-Path $script:Root 'encrypted.xml'
      [System.IO.File]::WriteAllText($_encrypted, "---- BEGIN config.xml ----`nU2FsdGVkX19hYmNk`n---- END config.xml ----`n", $script:Utf8)

      $_run = Invoke-Build -Parameters @{ Path = $_encrypted; OutputPath = $script:Output }
      $_run.ExitCode | Should -Be 1
      ($_run.Results | Where-Object Action -EQ 'Build').Error | Should -Match 'encrypted'
      Test-Path -LiteralPath $script:Output | Should -BeFalse
    }

    It 'rejects partial backups of a single area' {
      $_partial = Join-Path $script:Root 'system-config.xml'
      [System.IO.File]::WriteAllText($_partial, "<?xml version=`"1.0`"?>`n<system>`n`t<hostname>x</hostname>`n</system>`n", $script:Utf8)

      $_run = Invoke-Build -Parameters @{ Path = $_partial; OutputPath = $script:Output }
      $_run.ExitCode | Should -Be 1
      ($_run.Results | Where-Object Action -EQ 'Build').Error | Should -Match 'Partial backups'
    }

    It 'writes nothing with -DryRun and reports the planned changes as Preview' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Hostname = 'fw-branch'; DryRun = $true }
      $_run.ExitCode | Should -Be 0
      Test-Path -LiteralPath $script:Output | Should -BeFalse
      ($_run.Results | Where-Object Target -EQ 'system/hostname').Status | Should -Be 'Preview'
      ($_run.Results | Where-Object Action -EQ 'Build').Status | Should -Be 'Preview'
    }

    It 'rejects DTDs and incomplete configurations without touching an existing output' {
      $_invalid = Join-Path $script:Root 'invalid.xml'
      $_inputs = @(
        '<!DOCTYPE pfsense [<!ENTITY value "test">]><pfsense>&value;</pfsense>',
        '<pfsense><system><hostname>x</hostname></system></pfsense>',
        '<pfsense xmlns="urn:other"><version>23.3</version></pfsense>',
        '<pfsense><version>'
      )
      [System.IO.File]::WriteAllText($script:Output, 'keep this output', $script:Utf8)

      foreach ($_input in $_inputs) {
        [System.IO.File]::WriteAllText($_invalid, $_input, $script:Utf8)
        $_parameters = @{
          Path       = $_invalid
          OutputPath = $script:Output
          Force      = $true
        }
        $_run = Invoke-Build -Parameters $_parameters
        $_run.ExitCode | Should -Be 1
        [System.IO.File]::ReadAllText($script:Output) | Should -Be 'keep this output'
      }
    }

    It 'rejects edits to the configuration version during preview as well as writes' {
      foreach ($_preview in @($true, $false)) {
        $_parameters = @{
          Path       = $script:Source
          OutputPath = $script:Output
          Set        = @{ version = '999' }
          DryRun     = $_preview
        }
        $_run = Invoke-Build -Parameters $_parameters
        $_run.ExitCode | Should -Be 1
        Test-Path -LiteralPath $script:Output | Should -BeFalse
      }
    }
  }

  Context 'system and webConfigurator settings' {
    It 'applies identity, time, webConfigurator, and SSH settings' {
      $_parameters = @{
        Path                   = $script:Source
        OutputPath             = $script:Output
        Hostname               = 'fw-branch'
        Domain                 = 'branch.example.com'
        Timezone               = 'Europe/Berlin'
        TimeServer             = @('0.DE.pool.ntp.org', '1.de.pool.ntp.org')
        Language               = 'de_DE'
        WebGuiPort             = 8443
        WebGuiTheme            = 'pfSense-dark'
        WebGuiDashboardColumns = 3
        WebGuiTopNavigation    = 'Fixed'
        WebGuiHostnameInMenu   = 'Fqdn'
        WebGuiLoginColor       = '1E3F75'
        WebGuiAntiLockout      = $false
        WebGuiHttpRedirect     = $true
        SshAuthentication      = 'Key'
        SshPort                = 2222
        DnsAllowOverride       = $false
      }

      $_run = Invoke-Build -Parameters $_parameters
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'system/hostname' | Should -Be 'fw-branch'
      Get-TextValue -Root $_root -XPath 'system/domain' | Should -Be 'branch.example.com'
      Get-TextValue -Root $_root -XPath 'system/timezone' | Should -Be 'Europe/Berlin'
      Get-TextValue -Root $_root -XPath 'system/timeservers' | Should -Be '0.de.pool.ntp.org 1.de.pool.ntp.org'
      Get-TextValue -Root $_root -XPath 'system/language' | Should -Be 'de_DE'
      Get-TextValue -Root $_root -XPath 'system/webgui/port' | Should -Be '8443'
      Get-TextValue -Root $_root -XPath 'system/webgui/webguicss' | Should -Be 'pfSense-dark.css'
      Get-TextValue -Root $_root -XPath 'system/webgui/dashboardcolumns' | Should -Be '3'
      Get-TextValue -Root $_root -XPath 'system/webgui/webguifixedmenu' | Should -Be 'fixed'
      Get-TextValue -Root $_root -XPath 'system/webgui/webguihostnamemenu' | Should -Be 'fqdn'
      Get-TextValue -Root $_root -XPath 'system/webgui/logincss' | Should -Be '1e3f75'
      @($_root.SelectNodes('system/webgui/noantilockout')).Count | Should -Be 1
      @($_root.SelectNodes('system/webgui/disablehttpredirect')).Count | Should -Be 0
      @($_root.SelectNodes('system/dnsallowoverride')).Count | Should -Be 0
      Get-TextValue -Root $_root -XPath 'system/ssh/sshdkeyonly' | Should -Be 'enabled'
      Get-TextValue -Root $_root -XPath 'system/ssh/port' | Should -Be '2222'
      [System.IO.File]::ReadAllText($script:Output) | Should -Match "`t`t`t<noantilockout></noantilockout>`n"
    }

    It 'disables SSH by removing its enable flag' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; SshEnabled = $false }
      $_run.ExitCode | Should -Be 0
      @((Read-Output -Path $script:Output).SelectNodes('system/ssh/enable')).Count | Should -Be 0
    }

    It 'replaces DNS servers in place and resets their gateway bindings' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; DnsServer = @('9.9.9.9') }
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'system/dnsserver' | Should -Be @('9.9.9.9')
      Get-TextValue -Root $_root -XPath 'system/dns1gw' | Should -Be 'none'
      @($_root.SelectNodes('system/dns2gw')).Count | Should -Be 0
      $_root.SelectSingleNode('system/dnsserver').PreviousSibling.LocalName | Should -Be 'timeservers'
    }

    It 'rejects an invalid DNS server address' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; DnsServer = @('10.1') }
      $_run.ExitCode | Should -Be 1
      Test-Path -LiteralPath $script:Output | Should -BeFalse
    }

    It 'clears DNS servers and gateway bindings with an empty list' {
      $_parameters = @{
        Path       = $script:Source
        OutputPath = $script:Output
        DnsServer  = @()
      }
      $_run = Invoke-Build -Parameters $_parameters
      $_run.ExitCode | Should -Be 0
      $_root = Read-Output -Path $script:Output
      @($_root.SelectNodes('system/dnsserver | system/dns1gw | system/dns2gw')).Count | Should -Be 0
    }
  }

  Context 'interfaces and DHCP' {
    It 'sets an interface by description, including CIDR addresses and an escaped CDATA description' {
      $_interface = @{
        DMZ = @{
          IPAddress   = '10.30.0.1/23'
          Description = 'Guest & IoT'
          Enable      = $true
          MTU         = 1400
        }
      }

      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Interface = $_interface }
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'interfaces/opt1/ipaddr' | Should -Be '10.30.0.1'
      Get-TextValue -Root $_root -XPath 'interfaces/opt1/subnet' | Should -Be '23'
      Get-TextValue -Root $_root -XPath 'interfaces/opt1/mtu' | Should -Be '1400'
      @($_root.SelectNodes('interfaces/opt1/enable')).Count | Should -Be 1
      [System.IO.File]::ReadAllText($script:Output) | Should -Match ([regex]::Escape('<descr><![CDATA[Guest &amp; IoT]]></descr>'))
    }

    It 'removes the prefix when an interface switches to DHCP' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Interface = @{ opt1 = @{ IPAddress = 'dhcp' } } }
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'interfaces/opt1/ipaddr' | Should -Be 'dhcp'
      @($_root.SelectNodes('interfaces/opt1/subnet')).Count | Should -Be 0
    }

    It 'fails for an unknown interface and lists the available ones' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Interface = @{ opt9 = @{ Enable = $true } } }
      $_run.ExitCode | Should -Be 1
      ($_run.Results | Where-Object Action -EQ 'Build').Error | Should -Match 'opt1 \(DMZ\)'
      Test-Path -LiteralPath $script:Output | Should -BeFalse
    }

    It 'sets static IPv6 and rejects prefixes with a dynamic address' {
      $_parameters = @{
        Path       = $script:Source
        OutputPath = $script:Output
        Interface  = @{ lan = @{ IPv6Address = '2001:db8:20::1/64' } }
      }
      $_run = Invoke-Build -Parameters $_parameters
      $_run.ExitCode | Should -Be 0
      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'interfaces/lan/ipaddrv6' | Should -Be '2001:db8:20::1'
      Get-TextValue -Root $_root -XPath 'interfaces/lan/subnetv6' | Should -Be '64'

      $_parameters.Interface = @{ lan = @{ IPAddress = 'dhcp'; Subnet = 24 } }
      $_parameters.Force = $true
      $_run = Invoke-Build -Parameters $_parameters
      $_run.ExitCode | Should -Be 1
      Get-TextValue -Root (Read-Output -Path $script:Output) -XPath 'interfaces/lan/ipaddr' | Should -Be '192.168.1.1'
    }

    It 'fails for unknown interface settings and string booleans' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Interface = @{ lan = @{ Address = '10.0.0.1/24' } } }
      $_run.ExitCode | Should -Be 1

      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Interface = @{ lan = @{ Enable = 'false' } } }
      $_run.ExitCode | Should -Be 1
    }

    It 'updates DHCP ranges and warns when a range leaves the interface subnet' {
      $_parameters = @{
        Path       = $script:Source
        OutputPath = $script:Output
        Interface  = @{ lan = @{ IPAddress = '10.20.0.1/24' } }
        DhcpServer = @{ LAN = @{ RangeFrom = '10.20.0.100'; DnsServer = @('10.20.0.1') } }
      }

      $_run = Invoke-Build -Parameters $_parameters
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'dhcpd/lan/range/from' | Should -Be '10.20.0.100'
      Get-TextValue -Root $_root -XPath 'dhcpd/lan/dnsserver' | Should -Be @('10.20.0.1')
      $_warning = @($_run.Results | Where-Object Status -EQ 'Warn')
      $_warning.Count | Should -Be 1
      $_warning[0].Detail | Should -Match '192\.168\.1\.199'
    }
  }

  Context 'generic XPath edits' {
    It 'sets, creates, flags, removes, and replaces lists' {
      $_set = [ordered]@{
        'system/webgui/max_procs'                   = 4
        'system/webgui/althostnames'                = 'fw.example.com'
        'system/webgui/loginautocomplete'           = $true
        'system/dnsallowoverride'                   = $false
        'aliases/alias[name="Admins"]/address'      = '10.0.0.10 10.0.0.11'
        'system/dnsserver'                          = @('9.9.9.9', '149.112.112.112')
        '/pfsense/system/webgui/nohttpreferercheck' = $true
      }

      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Set = $_set }
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'system/webgui/max_procs' | Should -Be '4'
      Get-TextValue -Root $_root -XPath 'system/webgui/althostnames' | Should -Be 'fw.example.com'
      @($_root.SelectNodes('system/webgui/loginautocomplete')).Count | Should -Be 1
      @($_root.SelectNodes('system/webgui/nohttpreferercheck')).Count | Should -Be 1
      @($_root.SelectNodes('system/dnsallowoverride')).Count | Should -Be 0
      Get-TextValue -Root $_root -XPath 'aliases/alias[name="Admins"]/address' | Should -Be '10.0.0.10 10.0.0.11'
      Get-TextValue -Root $_root -XPath 'system/dnsserver' | Should -Be @('9.9.9.9', '149.112.112.112')
    }

    It 'fails when a scalar -Set matches more than one element' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Set = @{ 'system/dnsserver' = '9.9.9.9' } }
      $_run.ExitCode | Should -Be 1
      ($_run.Results | Where-Object Action -EQ 'Build').Error | Should -Match 'matches 2 elements'
    }

    It 'fails when -Set targets a section or a missing parent' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Set = @{ 'system/webgui' = 'x' } }
      $_run.ExitCode | Should -Be 1

      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Set = @{ 'nosuch/section/value' = 'x' } }
      $_run.ExitCode | Should -Be 1
    }

    It 'removes matched elements and fails when an expression matches nothing' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Remove = @('aliases/alias[name="Legacy"]') }
      $_run.ExitCode | Should -Be 0
      Get-TextValue -Root (Read-Output -Path $script:Output) -XPath 'aliases/alias/name' | Should -Be @('Admins')

      Remove-Item -LiteralPath $script:Output
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Remove = @('aliases/alias[name="Missing"]') }
      $_run.ExitCode | Should -Be 1
      Test-Path -LiteralPath $script:Output | Should -BeFalse
    }

    It 'never echoes secret values in results' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Set = @{ 'system/user/bcrypt-hash' = 'new-hash-value' } }
      $_run.ExitCode | Should -Be 0

      $_result = $_run.Results | Where-Object Target -EQ 'system/user/bcrypt-hash'
      $_result.Before | Should -Be '(redacted)'
      $_result.After | Should -Be '(redacted)'
    }
  }

  Context 'SSH host keys' {
    It 'removes the stored host keys and leaves user keys alone' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; RemoveSshHostKeys = $true }
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      @($_root.SelectNodes('sshdata')).Count | Should -Be 0
      @($_root.SelectNodes('system/user/authorizedkeys')).Count | Should -Be 1
    }

    It 'warns when the hostname changes but the host keys are kept' {
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; Hostname = 'fw-branch' }
      $_run.ExitCode | Should -Be 0
      @($_run.Results | Where-Object { $_.Status -eq 'Warn' -and $_.Detail -match 'RemoveSshHostKeys' }).Count | Should -Be 1
    }
  }

  Context 'spec files' {
    It 'allows an explicit false switch to retain host keys requested for removal in the spec' {
      $_spec = Join-Path $script:Root 'keep-keys.psd1'
      Set-Content -LiteralPath $_spec -Value '@{ RemoveSshHostKeys = $true }' -Encoding UTF8
      $_parameters = @{
        Path              = $script:Source
        OutputPath        = $script:Output
        SpecPath          = $_spec
        RemoveSshHostKeys = $false
      }
      $_run = Invoke-Build -Parameters $_parameters
      $_run.ExitCode | Should -Be 0
      @((Read-Output -Path $script:Output).SelectNodes('sshdata')).Count | Should -Be 1
    }

    It 'applies a spec file first and lets parameters override it' {
      $_spec = Join-Path $script:Root 'branch.psd1'
      $_specText = @(
        '@{',
        "  Hostname          = 'from-spec'",
        "  Timezone          = 'Europe/Vienna'",
        '  RemoveSshHostKeys = $true',
        "  Interface         = @{ lan = @{ IPAddress = '10.40.0.1/24' } }",
        "  Set               = @{ 'system/webgui/max_procs' = '3' }",
        '}'
      ) -join "`r`n"
      Set-Content -LiteralPath $_spec -Value $_specText -Encoding UTF8

      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; SpecPath = $_spec; Hostname = 'from-parameter' }
      $_run.ExitCode | Should -Be 0

      $_root = Read-Output -Path $script:Output
      Get-TextValue -Root $_root -XPath 'system/hostname' | Should -Be 'from-parameter'
      Get-TextValue -Root $_root -XPath 'system/timezone' | Should -Be 'Europe/Vienna'
      Get-TextValue -Root $_root -XPath 'interfaces/lan/ipaddr' | Should -Be '10.40.0.1'
      Get-TextValue -Root $_root -XPath 'system/webgui/max_procs' | Should -Be '3'
      @($_root.SelectNodes('sshdata')).Count | Should -Be 0
    }

    It 'rejects unknown spec keys and values that fail parameter validation' {
      $_spec = Join-Path $script:Root 'bad.psd1'

      Set-Content -LiteralPath $_spec -Value "@{ HostName = 'ok'; Colour = 'blue' }" -Encoding UTF8
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; SpecPath = $_spec }
      $_run.ExitCode | Should -Be 1
      ($_run.Results | Where-Object Action -EQ 'Build').Error | Should -Match "'Colour'"

      Set-Content -LiteralPath $_spec -Value '@{ WebGuiPort = 70000 }' -Encoding UTF8
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; SpecPath = $_spec }
      $_run.ExitCode | Should -Be 1

      Set-Content -LiteralPath $_spec -Value "@{ SshEnabled = 'false' }" -Encoding UTF8
      $_run = Invoke-Build -Parameters @{ Path = $script:Source; OutputPath = $script:Output; SpecPath = $_spec }
      $_run.ExitCode | Should -Be 1
      Test-Path -LiteralPath $script:Output | Should -BeFalse
    }
  }
}
