#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Mocked ACL state is scoped to the test and removed after each run.')]
param ()

BeforeAll {
  Import-Module AdNoctem.Substrate.PowerShell -RequiredVersion 0.3.1 -Force
  $script:ScriptPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '../../scripts/Maintenance/New-ManagedPaths.ps1'))
  $script:Content = Get-Content -LiteralPath $script:ScriptPath -Raw
  $script:HarnessPath = Join-Path $TestDrive 'New-ManagedPaths.ps1'
  $harnessText = $script:Content -replace '(?m)^#Requires -RunAsAdministrator\r?\n', ''
  $harnessText = $harnessText -replace '(?m)^\s*exit 1\s*$', 'return'
  $harnessText = $harnessText -replace '(?m)^\s*exit 0\s*$', 'return'
  [IO.File]::WriteAllText($script:HarnessPath, $harnessText, (New-Object Text.UTF8Encoding($true)))

  function Invoke-ManagedPaths {
    param (
      [hashtable]
      $Extra = @{}
    )

    $parameters = @{
      Path            = $script:Managed
      ReportDirectory = $script:Reports
      PassThru        = $true
      Confirm         = $false
    }
    foreach ($key in $Extra.Keys) {
      $parameters[$key] = $Extra[$key]
    }

    $global:LASTEXITCODE = 0
    $results = @(. $script:HarnessPath @parameters 6>$null)
    $summary = $results | Where-Object Action -EQ 'NewManagedPaths'
    $exitCode = 0
    if ($summary.Status -eq 'Failed') {
      $exitCode = 1
    }
    return [PSCustomObject]@{
      ExitCode = $exitCode
      Results  = $results
      Summary  = $summary
    }
  }
}

Describe 'New-ManagedPaths' {
  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Root -Force
    $script:Managed = Join-Path $script:Root 'Managed'
    $script:Reports = Join-Path $script:Root 'Reports'
    $global:WinkitManagedAppliedAcls = @{}
    $global:WinkitManagedSetAclCalls = New-Object System.Collections.ArrayList
    $global:LASTEXITCODE = 0

    Mock Set-Acl {
      $target = [string]@($LiteralPath)[0]
      $null = $global:WinkitManagedSetAclCalls.Add([PSCustomObject]@{
          Path = $target
          Acl  = $AclObject
        })
      $global:WinkitManagedAppliedAcls[$target] = $AclObject
    }
  }

  AfterEach {
    Remove-Variable -Name WinkitManagedAppliedAcls, WinkitManagedSetAclCalls -Scope Global -ErrorAction SilentlyContinue
  }

  It 'requires elevation and exposes preview conventions' {
    $script:Content | Should -Match '#Requires -RunAsAdministrator'
    $script:Content | Should -Match "RequiredVersion = '0.3.1'"
    $script:Content | Should -Match 'SupportsShouldProcess = \$true'
    (Get-Help $script:ScriptPath).Synopsis | Should -Match 'managed.*folder tree'
  }

  It 'creates the standard tree with the requested permissions' {
    $run = Invoke-ManagedPaths

    $run.ExitCode | Should -Be 0
    $run.Summary.Status | Should -Be 'Completed'
    $run.Summary.Created | Should -Be 10
    foreach ($name in @('Archives', 'Backups', 'Checkpoints', 'Media', 'Reports', 'Tools', 'Audit', 'Installables', 'Migration')) {
      Test-Path -LiteralPath (Join-Path $script:Managed $name) | Should -BeTrue
    }
    Test-Path -LiteralPath $run.Summary.ManifestPath | Should -BeTrue

    $mediaAcl = $global:WinkitManagedAppliedAcls[(Join-Path $script:Managed 'Media')]
    $mediaUsers = @($mediaAcl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]) | Where-Object {
        $_.IdentityReference.Value -eq 'S-1-5-32-545'
      })
    $mediaUsers.Count | Should -Be 1
    ($mediaUsers[0].FileSystemRights -band [Security.AccessControl.FileSystemRights]::ReadAndExecute) | Should -Be ([Security.AccessControl.FileSystemRights]::ReadAndExecute)

    $archivesAcl = $global:WinkitManagedAppliedAcls[(Join-Path $script:Managed 'Archives')]
    @($archivesAcl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]) | Where-Object {
        $_.IdentityReference.Value -eq 'S-1-5-32-545'
      }).Count | Should -Be 0
  }

  It 'previews without creating folders, changing ACLs, or writing a report' {
    $run = Invoke-ManagedPaths -Extra @{ DryRun = $true }

    $run.ExitCode | Should -Be 0
    $run.Summary.Status | Should -Be 'Preview'
    Test-Path -LiteralPath $script:Managed | Should -BeFalse
    Test-Path -LiteralPath $script:Reports | Should -BeFalse
    $global:WinkitManagedSetAclCalls.Count | Should -Be 0
  }

  It 'reports inherited permissions and changes them only with RepairPermissions' {
    $null = New-Item -ItemType Directory -Path (Join-Path $script:Managed 'Archives') -Force

    $first = Invoke-ManagedPaths
    $first.ExitCode | Should -Be 0
    $first.Summary.NeedsRepair | Should -Be 2
    @($global:WinkitManagedSetAclCalls | Where-Object Path -EQ (Join-Path $script:Managed 'Archives')).Count | Should -Be 0

    $global:WinkitManagedSetAclCalls.Clear()
    $second = Invoke-ManagedPaths -Extra @{ RepairPermissions = $true }
    $second.ExitCode | Should -Be 0
    @($global:WinkitManagedSetAclCalls | Where-Object Path -EQ (Join-Path $script:Managed 'Archives')).Count | Should -Be 1
  }

  It 'does nothing on a second run when the recorded permissions already conform' {
    $first = Invoke-ManagedPaths
    $first.Summary.Created | Should -Be 10

    Mock Get-Acl {
      $target = [string]@($LiteralPath)[0]
      if ($global:WinkitManagedAppliedAcls.ContainsKey($target)) {
        return $global:WinkitManagedAppliedAcls[$target]
      }
      return Microsoft.PowerShell.Security\Get-Acl -LiteralPath $target
    }

    $global:WinkitManagedSetAclCalls.Clear()
    $second = Invoke-ManagedPaths

    $second.ExitCode | Should -Be 0
    $second.Summary.Status | Should -Be 'Skipped'
    $global:WinkitManagedSetAclCalls.Count | Should -Be 0
  }

  It 'creates a per-user folder with Modify access and parent traverse' {
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
    $leaf = ($sid.Translate([Security.Principal.NTAccount])).Value.Split('\')[-1]

    $run = Invoke-ManagedPaths -Extra @{
      GrantUser = $sid.Value
      Subfolder = 'Archives'
    }

    $run.ExitCode | Should -Be 0
    $child = Join-Path (Join-Path $script:Managed 'Archives') $leaf
    Test-Path -LiteralPath $child | Should -BeTrue

    $parentAcl = $global:WinkitManagedAppliedAcls[(Join-Path $script:Managed 'Archives')]
    $parentGrant = @($parentAcl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]) | Where-Object {
        $_.IdentityReference.Value -eq $sid.Value
      })
    $parentGrant.Count | Should -Be 1
    ($parentGrant[0].FileSystemRights -band [Security.AccessControl.FileSystemRights]::Traverse) | Should -Be ([Security.AccessControl.FileSystemRights]::Traverse)

    $childAcl = $global:WinkitManagedAppliedAcls[$child]
    $childGrant = @($childAcl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]) | Where-Object {
        $_.IdentityReference.Value -eq $sid.Value
      })
    $childGrant.Count | Should -Be 1
    ($childGrant[0].FileSystemRights -band [Security.AccessControl.FileSystemRights]::Modify) | Should -Be ([Security.AccessControl.FileSystemRights]::Modify)
  }
}
