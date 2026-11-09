#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test fixture builders create in-memory objects only.')]
param ()

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force
  $script:LibraryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../../scripts/Server/ADDS/lib/ADDS.ps1'))
  . $script:LibraryPath

  if (-not (Get-Command -Name Get-ADUser -ErrorAction SilentlyContinue)) {
    function Get-ADUser {
      [CmdletBinding()]
      param (
        [string]$Filter,
        [string]$LDAPFilter,
        [string]$SearchBase,
        [string]$Identity,
        [string[]]$Properties,
        [string]$Server,
        [pscredential]$Credential
      )
      $null = $Filter, $LDAPFilter, $SearchBase, $Identity, $Properties, $Server, $Credential
      throw 'Get-ADUser is not available in this test environment.'
    }
  }

  if (-not (Get-Command -Name Get-ADComputer -ErrorAction SilentlyContinue)) {
    function Get-ADComputer {
      [CmdletBinding()]
      param (
        [string]$Filter,
        [string]$LDAPFilter,
        [string]$SearchBase,
        [string]$Identity,
        [string[]]$Properties,
        [string]$Server,
        [pscredential]$Credential
      )
      $null = $Filter, $LDAPFilter, $SearchBase, $Identity, $Properties, $Server, $Credential
      throw 'Get-ADComputer is not available in this test environment.'
    }
  }

  if (-not (Get-Command -Name Get-ADGroup -ErrorAction SilentlyContinue)) {
    function Get-ADGroup {
      [CmdletBinding()]
      param (
        [string]$Filter,
        [string]$LDAPFilter,
        [string]$SearchBase,
        [string]$Identity,
        [string[]]$Properties,
        [string]$Server,
        [pscredential]$Credential
      )
      $null = $Filter, $LDAPFilter, $SearchBase, $Identity, $Properties, $Server, $Credential
      throw 'Get-ADGroup is not available in this test environment.'
    }
  }

  if (-not (Get-Command -Name Get-ADGroupMember -ErrorAction SilentlyContinue)) {
    function Get-ADGroupMember {
      [CmdletBinding()]
      param (
        [string]$Identity,
        [switch]$Recursive,
        [string]$Server,
        [pscredential]$Credential
      )
      $null = $Identity, $Recursive, $Server, $Credential
      throw 'Get-ADGroupMember is not available in this test environment.'
    }
  }

  function New-TestUserRecord {
    param (
      [string]
      $SamAccountName,

      [string]
      $Mail,

      [bool]
      $Enabled = $true,

      [string]
      $DistinguishedName
    )

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) {
      $DistinguishedName = "CN=$SamAccountName"
    }

    $properties = [ordered]@{}
    foreach ($name in (Get-ADResourceProperty -Resource Users)) {
      $properties[$name] = "value-$name"
    }
    $properties['sAMAccountName'] = $SamAccountName
    $properties['mail'] = $Mail
    $properties['displayName'] = $SamAccountName
    $properties['Enabled'] = $Enabled
    $properties['DistinguishedName'] = $DistinguishedName

    $userPrincipalName = $null
    if (-not [string]::IsNullOrWhiteSpace($Mail)) {
      $userPrincipalName = ($Mail -split '@')[0]
    }

    [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.User'
      Resource          = 'Users'
      DistinguishedName = $DistinguishedName
      SamAccountName    = $SamAccountName
      Name              = $SamAccountName
      ObjectGuid        = $null
      Properties        = $properties
      Mail              = $Mail
      Enabled           = $Enabled
      UserPrincipalName = $userPrincipalName
    }
  }

  function New-TestComputerRecord {
    param (
      [string]
      $SamAccountName,

      [bool]
      $Enabled = $true
    )

    $properties = [ordered]@{}
    foreach ($name in (Get-ADResourceProperty -Resource Computers)) {
      $properties[$name] = "value-$name"
    }
    $properties['sAMAccountName'] = $SamAccountName
    $properties['Enabled'] = $Enabled

    [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.Computer'
      Resource          = 'Computers'
      DistinguishedName = "CN=$SamAccountName"
      SamAccountName    = $SamAccountName
      Name              = $SamAccountName
      ObjectGuid        = $null
      Properties        = $properties
      Enabled           = $Enabled
      DNSHostName       = "$SamAccountName.example.com"
    }
  }

  function New-TestGroupRecord {
    param (
      [string]
      $SamAccountName
    )

    $properties = [ordered]@{}
    foreach ($name in (Get-ADResourceProperty -Resource Groups)) {
      $properties[$name] = "value-$name"
    }
    $properties['sAMAccountName'] = $SamAccountName

    [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.Group'
      Resource          = 'Groups'
      DistinguishedName = "CN=$SamAccountName"
      SamAccountName    = $SamAccountName
      Name              = $SamAccountName
      ObjectGuid        = $null
      Properties        = $properties
      GroupScope        = 'Global'
      GroupCategory     = 'Security'
    }
  }

  function New-SparseUserRecord {
    param (
      [string]
      $SamAccountName,

      [string]
      $Mail
    )

    [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.User'
      Resource          = 'Users'
      DistinguishedName = "CN=$SamAccountName"
      SamAccountName    = $SamAccountName
      Name              = $SamAccountName
      ObjectGuid        = $null
      Properties        = [ordered]@{
        sAMAccountName = $SamAccountName
        mail           = $Mail
        Enabled        = $true
      }
      Mail              = $Mail
      Enabled           = $true
    }
  }
}

Describe 'Get-ADResourceProperty' {
  It 'returns the documented stable columns without case duplicates' {
    $users = @(Get-ADResourceProperty -Resource Users)

    $users[0] | Should -Be 'DistinguishedName'
    $users | Should -Contain 'sAMAccountName'
    $users | Should -Contain 'mail'
    @($users | Where-Object { $_ -ieq 'givenName' }).Count | Should -Be 1
  }

  It 'returns resource-specific columns' {
    @(Get-ADResourceProperty -Resource Computers) | Should -Contain 'DNSHostName'
    @(Get-ADResourceProperty -Resource Groups) | Should -Contain 'GroupScope'
  }
}

Describe 'Get-ADDirectoryRecord' {
  BeforeEach {
    Mock Get-ADUser {
      @(
        [PSCustomObject]@{
          SamAccountName    = 'asmith'
          DistinguishedName = 'CN=asmith'
          Name              = 'asmith'
          Enabled           = $true
          mail              = 'a.smith@adnoctem.co'
          displayName       = 'A Smith'
        },
        [PSCustomObject]@{
          SamAccountName    = 'bdis'
          DistinguishedName = 'CN=bdis'
          Name              = 'bdis'
          Enabled           = $false
          mail              = 'b.dis@adnoctem.co'
          displayName       = 'B Dis'
        }
      )
    }
  }

  It 'returns typed records and excludes disabled users by default' {
    $records = @(Get-ADDirectoryRecord -Resource Users -Filter '*')

    $records.Count | Should -Be 1
    $records[0].PSObject.TypeNames[0] | Should -Be 'Winkit.ADDS.User'
    $records[0].SamAccountName | Should -Be 'asmith'
    $records[0].Properties['mail'] | Should -Be 'a.smith@adnoctem.co'
  }

  It 'includes disabled users when requested' {
    $records = @(Get-ADDirectoryRecord -Resource Users -Filter '*' -IncludeDisabled)

    $records.Count | Should -Be 2
  }

  It 'accepts identities from the pipeline' {
    Mock Get-ADUser {
      param (
        $Identity
      )

      [PSCustomObject]@{
        SamAccountName    = $Identity
        DistinguishedName = "CN=$Identity"
        Name              = $Identity
        Enabled           = $true
        mail              = "$Identity@adnoctem.co"
        displayName       = $Identity
      }
    }

    $records = @('asmith', 'bjones' | Get-ADDirectoryRecord -Resource Users)

    $records.Count | Should -Be 2
    $records[0].SamAccountName | Should -Be 'asmith'
    $records[1].SamAccountName | Should -Be 'bjones'
  }

  It 'selects matching group members' {
    Mock Get-ADGroupMember {
      @(
        [PSCustomObject]@{ DistinguishedName = 'CN=asmith'; objectClass = 'user' },
        [PSCustomObject]@{ DistinguishedName = 'CN=pc01'; objectClass = 'computer' }
      )
    }
    Mock Get-ADUser {
      [PSCustomObject]@{
        SamAccountName    = 'asmith'
        DistinguishedName = 'CN=asmith'
        Name              = 'asmith'
        Enabled           = $true
        mail              = 'a.smith@adnoctem.co'
        displayName       = 'A Smith'
      }
    }

    $records = @(Get-ADDirectoryRecord -Resource Users -Group 'GG' -Recursive)

    $records.Count | Should -Be 1
    $records[0].SamAccountName | Should -Be 'asmith'
  }

  It 'rejects Resource All with identity selection' {
    { Get-ADDirectoryRecord -Resource All -Identity 'asmith' } | Should -Throw
  }

  It 'propagates query failures' {
    Mock Get-ADUser { throw 'query failed' }

    { Get-ADDirectoryRecord -Resource Users -Filter '*' } | Should -Throw '*query failed*'
  }
}

Describe 'Test-ADMailAddress' {
  It 'accepts a bare address on an allowed domain' {
    Test-ADMailAddress -Address 'maximilian.gindorfer@adnoctem.co' -AllowedDomain 'adnoctem.co', 'example.com' | Should -BeTrue
  }

  It 'matches domains case-insensitively' {
    Test-ADMailAddress -Address 'a@ADNOCTEM.CO' -AllowedDomain 'adnoctem.co' | Should -BeTrue
  }

  It 'rejects display-name forms, whitespace, and empty values' {
    Test-ADMailAddress -Address 'A Smith <a@adnoctem.co>' | Should -BeFalse
    Test-ADMailAddress -Address ' a@adnoctem.co' | Should -BeFalse
    Test-ADMailAddress -Address '' | Should -BeFalse
  }

  It 'rejects a domain outside the allowlist' {
    Test-ADMailAddress -Address 'a@adnoctem.co' -AllowedDomain 'example.com' | Should -BeFalse
    Test-ADMailAddress -Address 'a@mail.adnoctem.co' -AllowedDomain 'adnoctem.co' | Should -BeFalse
  }

  It 'rejects Unicode domain spellings' {
    $unicodeAddress = [string]::Concat('a@b', [char]0x00FC, 'cher.de')
    Test-ADMailAddress -Address $unicodeAddress | Should -BeFalse
  }
}

Describe 'Import-ADAccountMapOverride' {
  It 'reads a no-BOM override file with an optional note' {
    $path = Join-Path $TestDrive 'overrides.csv'
    $text = "samAccountName,email,note`r`nMGindorfer,maximilian.gindorfer@adnoctem.co,manual`r`n"
    [IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))

    $rows = @(Import-ADAccountMapOverride -Path $path)

    $rows.Count | Should -Be 1
    $rows[0].PSObject.TypeNames[0] | Should -Be 'Winkit.ADDS.AccountMapOverride'
    $rows[0].SamAccountName | Should -Be 'MGindorfer'
    $rows[0].Email | Should -Be 'maximilian.gindorfer@adnoctem.co'
    $rows[0].Note | Should -Be 'manual'
  }

  It 'reads a BOM override file' {
    $path = Join-Path $TestDrive 'overrides-bom.csv'
    $text = "samAccountName,email`r`nMGindorfer,maximilian.gindorfer@adnoctem.co`r`n"
    [IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($true)))

    $rows = @(Import-ADAccountMapOverride -Path $path)

    $rows.Count | Should -Be 1
  }

  It 'treats an empty email as an intentional exclusion' {
    $path = Join-Path $TestDrive 'excluded.csv'
    $text = "samAccountName,email`r`nServiceAccount,,exclude`r`n"
    [IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))

    $rows = @(Import-ADAccountMapOverride -Path $path)

    $rows[0].Email | Should -Be ''
  }

  It 'rejects duplicate account keys' {
    $path = Join-Path $TestDrive 'duplicate.csv'
    $text = "samAccountName,email`r`nMGindorfer,a@adnoctem.co`r`nmgindorfer,b@adnoctem.co`r`n"
    [IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))

    { Import-ADAccountMapOverride -Path $path } | Should -Throw '*repeats*'
  }

  It 'rejects malformed addresses' {
    $path = Join-Path $TestDrive 'invalid.csv'
    $text = "samAccountName,email`r`nMGindorfer,not-an-address`r`n"
    [IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))

    { Import-ADAccountMapOverride -Path $path } | Should -Throw '*invalid email*'
  }

  It 'rejects a file without required headers' {
    $path = Join-Path $TestDrive 'headers.csv'
    $text = "account,address`r`nMGindorfer,a@adnoctem.co`r`n"
    [IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))

    { Import-ADAccountMapOverride -Path $path } | Should -Throw '*samAccountName*'
  }
}

Describe 'New-ADAccountMap' {
  BeforeAll {
    $script:AllowedDomain = @('adnoctem.co', 'example.com')
  }

  It 'resolves the AD mail attribute' {
    $record = New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'

    $result = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain

    $result.Status | Should -Be 'Completed'
    $result.Map['MGindorfer'] | Should -Be 'maximilian.gindorfer@adnoctem.co'
    $result.Entries[0].Source | Should -Be 'ADMail'
  }

  It 'fails when mail is missing by default' {
    $record = New-TestUserRecord -SamAccountName 'cmiss' -Mail ''

    $result = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain

    $result.Status | Should -Be 'Failed'
    $result.Map.Count | Should -Be 0
    $result.Unresolved[0].Reason | Should -Be 'MissingMail'
  }

  It 'omits and reports missing mail with IncludeMissing' {
    $record = New-TestUserRecord -SamAccountName 'cmiss' -Mail ''

    $result = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -IncludeMissing

    $result.Status | Should -Be 'Completed'
    $result.ResolvedCount | Should -Be 0
    $result.UnresolvedCount | Should -Be 1
  }

  It 'prefers an override row over the AD mail attribute' {
    $overridePath = Join-Path $TestDrive 'override.csv'
    $text = "samAccountName,email,note`r`nMGindorfer,maximilian.gindorfer@example.com,manual`r`n"
    [IO.File]::WriteAllText($overridePath, $text, (New-Object System.Text.UTF8Encoding($false)))
    $record = New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'

    $result = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -OverridePath $overridePath

    $result.Map['MGindorfer'] | Should -Be 'maximilian.gindorfer@example.com'
    $result.Entries[0].Source | Should -Be 'Override'
    $result.Entries[0].Note | Should -Be 'manual'
  }

  It 'rejects a mail domain outside the allowlist' {
    $record = New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@example.net'

    $result = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain

    $result.Status | Should -Be 'Failed'
    $result.Unresolved[0].Reason | Should -Be 'MailDomainNotAllowed'
  }

  It 'fails duplicate keys by default' {
    $first = New-TestUserRecord -SamAccountName 'dup' -Mail 'a@adnoctem.co' -DistinguishedName 'CN=z'
    $second = New-TestUserRecord -SamAccountName 'dup' -Mail 'b@adnoctem.co' -DistinguishedName 'CN=a'

    $result = @($first, $second) | New-ADAccountMap -AllowedDomain $script:AllowedDomain

    $result.Status | Should -Be 'Failed'
    $result.Duplicates.Count | Should -Be 2
  }

  It 'chooses the first duplicate deterministically by distinguished name' {
    $first = New-TestUserRecord -SamAccountName 'dup' -Mail 'a@adnoctem.co' -DistinguishedName 'CN=z'
    $second = New-TestUserRecord -SamAccountName 'dup' -Mail 'b@adnoctem.co' -DistinguishedName 'CN=a'

    $result = @($first, $second) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -OnDuplicate First

    $result.Status | Should -Be 'Completed'
    $result.Map['dup'] | Should -Be 'b@adnoctem.co'
  }

  It 'reports and excludes unmatched overrides by default' {
    $overridePath = Join-Path $TestDrive 'unmatched.csv'
    $text = "samAccountName,email`r`nGhost,ghost@adnoctem.co`r`n"
    [IO.File]::WriteAllText($overridePath, $text, (New-Object System.Text.UTF8Encoding($false)))
    $record = New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'

    $result = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -OverridePath $overridePath

    $result.UnmatchedOverride.Count | Should -Be 1
    $result.Map.Contains('Ghost') | Should -BeFalse
  }

  It 'includes unmatched overrides when requested' {
    $overridePath = Join-Path $TestDrive 'unmatched-include.csv'
    $text = "samAccountName,email`r`nGhost,ghost@adnoctem.co`r`n"
    [IO.File]::WriteAllText($overridePath, $text, (New-Object System.Text.UTF8Encoding($false)))
    $record = New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'

    $result = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -OverridePath $overridePath -IncludeUnmatchedOverride

    $result.Map['Ghost'] | Should -Be 'ghost@adnoctem.co'
  }

  It 'requires an explicit switch for an override outside the allowlist' {
    $overridePath = Join-Path $TestDrive 'outside.csv'
    $text = "samAccountName,email`r`nMGindorfer,maximilian.gindorfer@example.net`r`n"
    [IO.File]::WriteAllText($overridePath, $text, (New-Object System.Text.UTF8Encoding($false)))
    $record = New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'

    $blocked = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -OverridePath $overridePath
    $allowed = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -OverridePath $overridePath -AllowOverrideOutsideDomain

    $blocked.Status | Should -Be 'Failed'
    $blocked.Unresolved[0].Reason | Should -Be 'OverrideDomainNotAllowed'
    $allowed.Map['MGindorfer'] | Should -Be 'maximilian.gindorfer@example.net'
  }

  It 'excludes disabled users unless IncludeDisabled is supplied' {
    $record = New-TestUserRecord -SamAccountName 'disabled' -Mail 'disabled@adnoctem.co' -Enabled $false

    $default = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain
    $included = @($record) | New-ADAccountMap -AllowedDomain $script:AllowedDomain -IncludeDisabled

    $default.EntryCount | Should -Be 0
    $included.Map['disabled'] | Should -Be 'disabled@adnoctem.co'
  }

  It 'preserves AD casing and sorts map keys deterministically' {
    $zeta = New-TestUserRecord -SamAccountName 'Zeta' -Mail 'z@adnoctem.co'
    $alpha = New-TestUserRecord -SamAccountName 'Alpha' -Mail 'a@adnoctem.co'

    $result = @($zeta, $alpha) | New-ADAccountMap -AllowedDomain $script:AllowedDomain

    $result.Map.Keys[0] | Should -Be 'Alpha'
    $result.Map.Keys[1] | Should -Be 'Zeta'
  }
}

Describe 'ConvertTo-ADAccountMapJson' {
  It 'writes sorted keys with CRLF and a final newline' {
    $map = [ordered]@{
      Zeta  = 'z@example.com'
      alpha = 'a@adnoctem.co'
      Beta  = 'b@adnoctem.co'
    }

    $json = ConvertTo-ADAccountMapJson -Map $map
    $expected = [string]::Join("`r`n", @(
        '{',
        '  "alpha": "a@adnoctem.co",',
        '  "Beta": "b@adnoctem.co",',
        '  "Zeta": "z@example.com"',
        '}',
        ''
      ))

    $json | Should -Be $expected
  }

  It 'writes an empty object for an empty map' {
    ConvertTo-ADAccountMapJson -Map ([ordered]@{}) | Should -Be "{}`r`n"
  }

  It 'escapes quotes, backslashes, and control characters' {
    $json = ConvertTo-ADAccountMapJson -Map ([ordered]@{ 'a"b' = "c\d`t" })

    $json.Contains('a\"b') | Should -BeTrue
    $json.Contains('c\\d') | Should -BeTrue
    $json.Contains('\u0009') | Should -BeTrue
  }

  It 'rejects duplicate keys after case folding' {
    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    $map.Add('Sam', 'a@adnoctem.co')
    $map.Add('sam', 'b@adnoctem.co')

    { ConvertTo-ADAccountMapJson -Map $map } | Should -Throw '*duplicate*'
  }
}

Describe 'Export-ADDirectoryRecord' {
  BeforeEach {
    $script:OutputDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:OutputDirectory
  }

  It 'writes a semicolon-delimited CSV with the stable schema' {
    $record = New-TestUserRecord -SamAccountName 'asmith' -Mail 'a.smith@adnoctem.co'

    $result = @($record) | Export-ADDirectoryRecord -Resource Users -OutputPath $script:OutputDirectory -Format Csv -PassThru
    $file = Join-Path $script:OutputDirectory 'AD-Users.csv'
    $rows = @(Import-Csv -LiteralPath $file -Delimiter ';')

    $result.Status | Should -Be 'Completed'
    $rows.Count | Should -Be 1
    $rows[0].sAMAccountName | Should -Be 'asmith'
    $rows[0].mail | Should -Be 'a.smith@adnoctem.co'
    (Get-Content -LiteralPath $file -TotalCount 1) | Should -Match '"DistinguishedName";"Enabled";'
  }

  It 'writes a JSON array' {
    $record = New-TestUserRecord -SamAccountName 'asmith' -Mail 'a.smith@adnoctem.co'

    $null = @($record) | Export-ADDirectoryRecord -Resource Users -OutputPath $script:OutputDirectory -Format Json
    $json = Get-Content -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.json') -Raw | ConvertFrom-Json

    $json.Count | Should -Be 1
    $json[0].sAMAccountName | Should -Be 'asmith'
  }

  It 'omits requested properties from the output' {
    $record = New-TestUserRecord -SamAccountName 'asmith' -Mail 'a.smith@adnoctem.co'

    $null = @($record) | Export-ADDirectoryRecord -Resource Users -OutputPath $script:OutputDirectory -Omit 'mail'
    $rows = @(Import-Csv -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') -Delimiter ';')

    @($rows[0].PSObject.Properties.Name) | Should -Not -Contain 'mail'
  }

  It 'fails on missing properties unless IncludeMissing is supplied' {
    $record = New-SparseUserRecord -SamAccountName 'cmiss' -Mail ''

    { @($record) | Export-ADDirectoryRecord -Resource Users -OutputPath $script:OutputDirectory } | Should -Throw '*Missing requested properties*'

    $result = @($record) | Export-ADDirectoryRecord -Resource Users -OutputPath $script:OutputDirectory -IncludeMissing -PassThru

    $result.Status | Should -Be 'Completed'
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') | Should -BeTrue
  }

  It 'writes nothing in WhatIf mode' {
    $record = New-TestUserRecord -SamAccountName 'asmith' -Mail 'a.smith@adnoctem.co'

    $result = @($record) | Export-ADDirectoryRecord -Resource Users -OutputPath $script:OutputDirectory -WhatIf -PassThru

    $result.Status | Should -Be 'DryRun'
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') | Should -BeFalse
  }

  It 'skips an empty input without writing a file' {
    $result = @() | Export-ADDirectoryRecord -Resource Users -OutputPath $script:OutputDirectory -PassThru

    $result.Status | Should -Be 'Skipped'
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') | Should -BeFalse
  }
}

Describe 'Export-ADAccountMap' {
  BeforeEach {
    $script:OutputDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:OutputDirectory
    $script:MapPath = Join-Path $script:OutputDirectory 'ad-account-map.json'
    $script:ReportPath = Join-Path $script:OutputDirectory 'audit.csv'
  }

  It 'writes a BOM-free deterministic map and an audit report' {
    $records = @(
      (New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'),
      (New-TestUserRecord -SamAccountName 'EMustermann' -Mail 'erika.mustermann@example.com')
    )
    $map = @($records) | New-ADAccountMap -AllowedDomain @('adnoctem.co', 'example.com')

    $result = $map | Export-ADAccountMap -OutputPath $script:MapPath -ReportPath $script:ReportPath -PassThru
    $bytes = [IO.File]::ReadAllBytes($script:MapPath)

    $result.Status | Should -Be 'Completed'
    $bytes[0] | Should -Be ([byte][char]'{')
    ($bytes[0..2] | ForEach-Object { $_.ToString('X2') }) -join ' ' | Should -Not -Be 'EF BB BF'
    Test-Path -LiteralPath $script:ReportPath | Should -BeTrue
    (Get-Content -LiteralPath $script:ReportPath -TotalCount 1) | Should -Match 'samAccountName'
  }

  It 'produces identical bytes on a second run' {
    $records = @(New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co')
    $map = @($records) | New-ADAccountMap -AllowedDomain @('adnoctem.co')

    $null = $map | Export-ADAccountMap -OutputPath $script:MapPath
    $firstHash = (Get-FileHash -LiteralPath $script:MapPath -Algorithm SHA256).Hash
    $null = $map | Export-ADAccountMap -OutputPath $script:MapPath -Force
    $secondHash = (Get-FileHash -LiteralPath $script:MapPath -Algorithm SHA256).Hash

    $secondHash | Should -Be $firstHash
  }

  It 'refuses to overwrite without Force' {
    $records = @(New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co')
    $map = @($records) | New-ADAccountMap -AllowedDomain @('adnoctem.co')
    [IO.File]::WriteAllText($script:MapPath, 'existing', (New-Object System.Text.UTF8Encoding($false)))

    { $map | Export-ADAccountMap -OutputPath $script:MapPath } | Should -Throw '*already exists*'
  }

  It 'writes nothing in WhatIf mode' {
    $records = @(New-TestUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co')
    $map = @($records) | New-ADAccountMap -AllowedDomain @('adnoctem.co')

    $result = $map | Export-ADAccountMap -OutputPath $script:MapPath -WhatIf -PassThru

    $result.Status | Should -Be 'DryRun'
    Test-Path -LiteralPath $script:MapPath | Should -BeFalse
  }

  It 'writes only the report for a failed result' {
    $record = New-TestUserRecord -SamAccountName 'cmiss' -Mail ''
    $map = @($record) | New-ADAccountMap -AllowedDomain @('adnoctem.co')

    $result = $map | Export-ADAccountMap -OutputPath $script:MapPath -ReportPath $script:ReportPath -PassThru

    $result.Status | Should -Be 'Failed'
    Test-Path -LiteralPath $script:MapPath | Should -BeFalse
    Test-Path -LiteralPath $script:ReportPath | Should -BeTrue
  }
}
