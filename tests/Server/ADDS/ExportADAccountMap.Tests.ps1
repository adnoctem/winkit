#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test fixture builders create in-memory objects only.')]
param ()

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force
  $script:ScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../../scripts/Server/ADDS/Export-ADAccountMap.ps1'))
  $script:Content = Get-Content -LiteralPath $script:ScriptPath -Raw
  $script:LibraryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../../scripts/Server/ADDS/lib/ADDS.ps1'))
  . $script:LibraryPath

  function New-ScriptUserRecord {
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

    [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.User'
      Resource          = 'Users'
      DistinguishedName = $DistinguishedName
      SamAccountName    = $SamAccountName
      Name              = $SamAccountName
      ObjectGuid        = $null
      Properties        = [ordered]@{
        sAMAccountName    = $SamAccountName
        mail              = $Mail
        displayName       = $SamAccountName
        Enabled           = $Enabled
        DistinguishedName = $DistinguishedName
      }
      Mail              = $Mail
      Enabled           = $Enabled
    }
  }
}

Describe 'Export-ADAccountMap conventions' {
  It 'uses the repository requirements and exposes the account-map surface' {
    $script:Content | Should -Match '#Requires -Version 5.1'
    $script:Content | Should -Not -Match 'RunAsAdministrator'
    $script:Content | Should -Match "RequiredVersion = '0.3.1'"
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    (Get-Help $script:ScriptPath).Synopsis | Should -Match 'account map'

    $command = Get-Command -Name $script:ScriptPath
    $command.Parameters.Keys | Should -Contain 'AllowedDomain'
    $command.Parameters.Keys | Should -Contain 'OverridePath'
    $command.Parameters.Keys | Should -Contain 'IncludeMissing'
    $command.Parameters.Keys | Should -Contain 'IncludeDisabled'
    $command.Parameters.Keys | Should -Contain 'ReportPath'
    $command.Parameters.Keys | Should -Contain 'OnDuplicate'
  }
}

Describe 'Export-ADAccountMap workflow' {
  BeforeEach {
    $script:OutputDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:OutputDirectory
    $script:MapPath = Join-Path $script:OutputDirectory 'ad-account-map.json'
    $script:ReportPath = Join-Path $script:OutputDirectory 'audit.csv'

    Mock Get-ADDirectoryRecord {
      param (
        $Resource,
        $Filter,
        $LDAPFilter,
        $SearchBase,
        $Server,
        [pscredential]$Credential,
        $Properties,
        $IncludeDisabled
      )

      $null = $Resource, $Filter, $LDAPFilter, $SearchBase, $Server, $Credential, $Properties

      if ($IncludeDisabled) {
        @(
          (New-ScriptUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'),
          (New-ScriptUserRecord -SamAccountName 'Disabled' -Mail 'disabled@example.com' -Enabled $false)
        )
      }
      else {
        @(
          New-ScriptUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'
        )
      }
    }
  }

  It 'writes the map and the audit report' {
    $output = @(& $script:ScriptPath -AllowedDomain 'adnoctem.co', 'example.com' -OutputPath $script:MapPath -ReportPath $script:ReportPath -PassThru 6>$null)

    $LASTEXITCODE | Should -Be 0
    $output[0].Status | Should -Be 'Completed'
    Test-Path -LiteralPath $script:MapPath | Should -BeTrue
    Test-Path -LiteralPath $script:ReportPath | Should -BeTrue
    $map = Get-Content -LiteralPath $script:MapPath -Raw | ConvertFrom-Json
    $map.MGindorfer | Should -Be 'maximilian.gindorfer@adnoctem.co'
  }

  It 'fails and writes only the report when mail is missing' {
    Mock Get-ADDirectoryRecord {
      param (
        $Resource,
        $Filter,
        $LDAPFilter,
        $SearchBase,
        $Server,
        [pscredential]$Credential,
        $Properties,
        $IncludeDisabled
      )

      $null = $Resource, $Filter, $LDAPFilter, $SearchBase, $Server, $Credential, $Properties, $IncludeDisabled
      New-ScriptUserRecord -SamAccountName 'cmiss' -Mail ''
    }

    $output = @(& $script:ScriptPath -AllowedDomain 'adnoctem.co' -OutputPath $script:MapPath -ReportPath $script:ReportPath -PassThru 6>$null)

    $LASTEXITCODE | Should -Be 1
    $output[0].Status | Should -Be 'Failed'
    Test-Path -LiteralPath $script:MapPath | Should -BeFalse
    Test-Path -LiteralPath $script:ReportPath | Should -BeTrue
  }

  It 'omits unresolved accounts with IncludeMissing' {
    Mock Get-ADDirectoryRecord {
      param (
        $Resource,
        $Filter,
        $LDAPFilter,
        $SearchBase,
        $Server,
        [pscredential]$Credential,
        $Properties,
        $IncludeDisabled
      )

      $null = $Resource, $Filter, $LDAPFilter, $SearchBase, $Server, $Credential, $Properties, $IncludeDisabled
      @(
        (New-ScriptUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'),
        (New-ScriptUserRecord -SamAccountName 'cmiss' -Mail '')
      )
    }

    $output = @(& $script:ScriptPath -AllowedDomain 'adnoctem.co' -OutputPath $script:MapPath -IncludeMissing -PassThru 6>$null)

    $LASTEXITCODE | Should -Be 0
    $output[0].Status | Should -Be 'Completed'
    $map = Get-Content -LiteralPath $script:MapPath -Raw | ConvertFrom-Json
    $map.MGindorfer | Should -Be 'maximilian.gindorfer@adnoctem.co'
    @($map.PSObject.Properties.Name) | Should -Not -Contain 'cmiss'
  }

  It 'writes nothing in DryRun mode' {
    $output = @(& $script:ScriptPath -AllowedDomain 'adnoctem.co' -OutputPath $script:MapPath -DryRun -PassThru 6>$null)

    $LASTEXITCODE | Should -Be 0
    $output[0].Status | Should -Be 'DryRun'
    Test-Path -LiteralPath $script:MapPath | Should -BeFalse
  }

  It 'requests enabled users by default and includes disabled on request' {
    $null = & $script:ScriptPath -AllowedDomain 'adnoctem.co' -OutputPath $script:MapPath 6>$null
    Should -Invoke Get-ADDirectoryRecord -ParameterFilter { $IncludeDisabled -eq $false }

    $null = & $script:ScriptPath -AllowedDomain 'adnoctem.co' -OutputPath $script:MapPath -IncludeDisabled -Force 6>$null
    Should -Invoke Get-ADDirectoryRecord -ParameterFilter { $IncludeDisabled -eq $true }
  }

  It 'matches the synthetic golden bytes' {
    $records = @(
      (New-ScriptUserRecord -SamAccountName 'AMeyer' -Mail 'andreas.meyer@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'EMustermann' -Mail 'erika.mustermann@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'JSchmidt' -Mail 'johann.schmidt@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'JSchulz' -Mail 'juergen.schulz@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'MGindorfer' -Mail 'maximilian.gindorfer@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'MMustermann' -Mail 'maximilian.mustermann@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'SSchneider' -Mail 'sabine.schneider@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'SWeber' -Mail 'susanne.weber@adnoctem.co'),
      (New-ScriptUserRecord -SamAccountName 'TMueller' -Mail 'thomas.mueller@adnoctem.co')
    )
    $map = @($records) | New-ADAccountMap -AllowedDomain @('adnoctem.co')
    $null = $map | Export-ADAccountMap -OutputPath $script:MapPath

    $expectedPath = Join-Path $PSScriptRoot 'fixtures/ad-account-map.synthetic.json'
    $expected = [Convert]::ToBase64String([IO.File]::ReadAllBytes($expectedPath))
    $actual = [Convert]::ToBase64String([IO.File]::ReadAllBytes($script:MapPath))

    $actual | Should -Be $expected
  }
}

Describe 'Owner-provided AD fixture' {
  It 'generates the committed account map from the owner CSV fixture' -Skip {
    # TODO: Enable after the owner supplies tests/Server/ADDS/fixtures/ad-records.csv.
    #
    # Required CSV columns (header names case-insensitive):
    #   sAMAccountName, Enabled, mail, displayName, distinguishedName
    #
    # Required rows (at least eight):
    #   1. An enabled user with a valid adnoctem.co mail value.
    #   2. An enabled user with a valid example.com mail value.
    #   3. An enabled user with an empty mail value (missing-mail case).
    #   4. An enabled user whose mail is outside the allowlist.
    #   5. A disabled user with a valid mail value.
    #   6. Two rows with the same sAMAccountName and different
    #      distinguishedName values (duplicate case).
    #   7. A case-variant sAMAccountName that must preserve its AD casing.
    #
    # The test will convert each row into a Winkit.ADDS.User record, feed the
    # records to New-ADAccountMap with AllowedDomain adnoctem.co and
    # example.com, and compare the exported map with the owner-provided
    # tests/Server/ADDS/fixtures/ad-account-map.golden.json. An optional
    # tests/Server/ADDS/fixtures/account-map-overrides.csv will cover override
    # precedence, an intentional exclusion, an unmatched override, and an
    # outside-domain override.
  }
}
