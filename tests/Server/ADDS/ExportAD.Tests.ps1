#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test fixture builders create in-memory objects only.')]
param ()

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force
  $script:ScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../../scripts/Server/ADDS/Export-AD.ps1'))
  $script:LibraryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../../scripts/Server/ADDS/lib/ADDS.ps1'))
  $script:Content = Get-Content -LiteralPath $script:ScriptPath -Raw
  . $script:LibraryPath

  function New-ScriptUserRecord {
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
        sAMAccountName    = $SamAccountName
        mail              = $Mail
        displayName       = $SamAccountName
        Enabled           = $true
        DistinguishedName = "CN=$SamAccountName"
      }
      Mail              = $Mail
      Enabled           = $true
    }
  }

  function New-ScriptComputerRecord {
    param (
      [string]
      $SamAccountName
    )

    [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.Computer'
      Resource          = 'Computers'
      DistinguishedName = "CN=$SamAccountName"
      SamAccountName    = $SamAccountName
      Name              = $SamAccountName
      ObjectGuid        = $null
      Properties        = [ordered]@{
        sAMAccountName    = $SamAccountName
        Enabled           = $true
        DistinguishedName = "CN=$SamAccountName"
      }
      Enabled           = $true
      DNSHostName       = "$SamAccountName.example.com"
    }
  }

  function New-ScriptGroupRecord {
    param (
      [string]
      $SamAccountName
    )

    [PSCustomObject]@{
      PSTypeName        = 'Winkit.ADDS.Group'
      Resource          = 'Groups'
      DistinguishedName = "CN=$SamAccountName"
      SamAccountName    = $SamAccountName
      Name              = $SamAccountName
      ObjectGuid        = $null
      Properties        = [ordered]@{
        sAMAccountName    = $SamAccountName
        DistinguishedName = "CN=$SamAccountName"
      }
      GroupScope        = 'Global'
      GroupCategory     = 'Security'
    }
  }
}

Describe 'Export-AD conventions' {
  It 'uses the repository requirements and exposes the new surface' {
    $script:Content | Should -Match '#Requires -Version 5.1'
    $script:Content | Should -Not -Match 'RunAsAdministrator'
    $script:Content | Should -Match "RequiredVersion = '0.3.1'"
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    (Get-Help $script:ScriptPath).Synopsis | Should -Match 'Exports Active Directory'

    $command = Get-Command -Name $script:ScriptPath
    $command.Parameters.Keys | Should -Contain 'Format'
    $command.Parameters.Keys | Should -Contain 'Omit'
    $command.Parameters.Keys | Should -Contain 'IncludeMissing'
    $command.Parameters.Keys | Should -Contain 'IncludeDisabled'
  }
}

Describe 'Export-AD workflow' {
  BeforeEach {
    $script:OutputDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:OutputDirectory

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

      $null = $Filter, $LDAPFilter, $SearchBase, $Server, $Credential, $Properties, $IncludeDisabled

      switch ($Resource) {
        'Users' { New-ScriptUserRecord -SamAccountName 'asmith' -Mail 'a.smith@adnoctem.co' }
        'Computers' { New-ScriptComputerRecord -SamAccountName 'PC01$' }
        'Groups' { New-ScriptGroupRecord -SamAccountName 'GG' }
      }
    }
  }

  It 'writes a CSV and returns operation results' {
    $output = @(& $script:ScriptPath -OutputPath $script:OutputDirectory -Resource Users -IncludeMissing -PassThru 6>$null)

    $LASTEXITCODE | Should -Be 0
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') | Should -BeTrue
    $output.Count | Should -Be 1
    $output[0].Status | Should -Be 'Completed'
    $output[0].Resource | Should -Be 'Users'
  }

  It 'writes JSON when requested' {
    $null = & $script:ScriptPath -OutputPath $script:OutputDirectory -Resource Users -Format Json -IncludeMissing 6>$null

    $LASTEXITCODE | Should -Be 0
    $json = Get-Content -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.json') -Raw | ConvertFrom-Json
    $json[0].sAMAccountName | Should -Be 'asmith'
  }

  It 'writes all three resources' {
    $null = & $script:ScriptPath -OutputPath $script:OutputDirectory -Resource All -IncludeMissing 6>$null

    $LASTEXITCODE | Should -Be 0
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Computers.csv') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Groups.csv') | Should -BeTrue
  }

  It 'writes nothing in DryRun mode' {
    $output = @(& $script:ScriptPath -OutputPath $script:OutputDirectory -Resource Users -IncludeMissing -DryRun -PassThru 6>$null)

    $LASTEXITCODE | Should -Be 0
    $output[0].Status | Should -Be 'DryRun'
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') | Should -BeFalse
  }

  It 'exits 1 when required properties are missing' {
    Mock Get-ADDirectoryRecord {
      param (
        $Resource
      )

      $null = $Resource
      New-ScriptUserRecord -SamAccountName 'asmith' -Mail 'a.smith@adnoctem.co'
    }

    $null = & $script:ScriptPath -OutputPath $script:OutputDirectory -Resource Users 6>$null

    $LASTEXITCODE | Should -Be 1
    Test-Path -LiteralPath (Join-Path $script:OutputDirectory 'AD-Users.csv') | Should -BeFalse
  }

  It 'requests enabled users by default and includes disabled on request' {
    $null = & $script:ScriptPath -OutputPath $script:OutputDirectory -Resource Users -IncludeMissing 6>$null
    Should -Invoke Get-ADDirectoryRecord -ParameterFilter { $IncludeDisabled -eq $false }

    $null = & $script:ScriptPath -OutputPath $script:OutputDirectory -Resource Users -IncludeMissing -IncludeDisabled 6>$null
    Should -Invoke Get-ADDirectoryRecord -ParameterFilter { $IncludeDisabled -eq $true }
  }
}
