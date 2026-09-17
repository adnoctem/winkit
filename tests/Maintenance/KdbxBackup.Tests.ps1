#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Tier-0 logic tests for the KeePass backup scripts (Backup-KdbxDatabase,
# Backup-KdbxDatabaseToDrive, New-DriveMarker). Every script runs end to end
# against fake KDBX files under TestDrive; "drives" are TestDrive folders passed
# through -DriveRoot. No real database, removable drive, or elevation is needed.

BeforeAll {
  $script:MaintenanceScripts = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot '..\..\scripts\Maintenance'))
  $script:BackupScript = Join-Path $script:MaintenanceScripts 'Backup-KdbxDatabase.ps1'
  $script:DriveScript = Join-Path $script:MaintenanceScripts 'Backup-KdbxDatabaseToDrive.ps1'
  $script:MarkerScript = Join-Path $script:MaintenanceScripts 'New-DriveMarker.ps1'

  function Write-FakeKdbx {
    param ([string]$Path, [string]$Content, [switch]$NoSignature)

    $null = New-Item -ItemType Directory -Path (Split-Path -Path $Path -Parent) -Force
    $_signature = if ($NoSignature) { [byte[]]@() } else { [byte[]](0x03, 0xD9, 0xA2, 0x9A, 0x67, 0xFB, 0x4B, 0xB5, 0x00, 0x00, 0x04, 0x00) }
    [System.IO.File]::WriteAllBytes($Path, [byte[]]($_signature + [System.Text.Encoding]::UTF8.GetBytes($Content)))
  }

  function Get-TestHash {
    param ([string]$Path)
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
  }

  function Initialize-KdbxSource {
    param ([string]$Root)

    $_source = [PSCustomObject]@{
      Database = Join-Path $Root 'live\Vault.kdbx'
      Backups  = Join-Path $Root 'backups'
    }
    Write-FakeKdbx -Path $_source.Database -Content 'live version 1'
    Write-FakeKdbx -Path (Join-Path $_source.Backups 'Vault_01_05_2026_10-00.old.kdbx') -Content 'old version A'
    Write-FakeKdbx -Path (Join-Path $_source.Backups 'Vault_02_05_2026_11-30.old.kdbx') -Content 'old version B'
    return $_source
  }
}

Describe 'KeePass backup script conventions' {
  BeforeAll {
    $script:Scripts = @($script:BackupScript, $script:DriveScript, $script:MarkerScript) | ForEach-Object { Get-Item -LiteralPath $_ }
  }

  It 'pins PSFoundation and documents .SYNOPSIS and a complete .NOTES block' {
    foreach ($_script in $script:Scripts) {
      $_header = Get-Content -LiteralPath $_script.FullName -TotalCount 3
      $_header -match "^#Requires -Modules @\{ ModuleName = 'PSFoundation'; ModuleVersion = '1\.4\.0' \}" | Should -Not -BeNullOrEmpty -Because "$($_script.Name) must pin PSFoundation 1.4.0"
      $_content = Get-Content -LiteralPath $_script.FullName -Raw
      $_content -match 'Import-Module PSFoundation -Force' | Should -BeTrue
      (Get-Help $_script.FullName).Synopsis | Should -Not -Match ('^' + [regex]::Escape($_script.Name)) -Because "$($_script.Name) comment-based help must be discoverable, not auto-generated syntax"
      $_content -match 'Author: MVProwess' | Should -BeTrue
      $_content -match 'Server Core:' | Should -BeTrue
      $_content -match 'SYSTEM-account execution:' | Should -BeTrue
    }
  }

  It 'tells the user the drive must be marked with New-DriveMarker.ps1' {
    (Get-Help $script:DriveScript -Full).Description.Text | Should -Match 'New-DriveMarker\.ps1'
  }
}

Describe 'Backup-KdbxDatabase' {
  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $script:Source = Initialize-KdbxSource -Root $script:Root
    $script:Destination = Join-Path $script:Root 'destination'
    $null = New-Item -ItemType Directory -Path $script:Destination -Force
  }

  It 'copies backups and a live snapshot, verified and recorded in SHA256SUMS' {
    $global:LASTEXITCODE = 0
    $_results = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database -BackupPath $script:Source.Backups -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 0
    @($_results | Where-Object Status -EQ 'Completed').Count | Should -Be 3

    $_snapshots = @(Get-ChildItem -LiteralPath (Join-Path $script:Destination 'snapshots') -File)
    $_snapshots.Count | Should -Be 1
    $_snapshots[0].Name | Should -Match '^Vault_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.kdbx$'
    Get-TestHash $_snapshots[0].FullName | Should -Be (Get-TestHash $script:Source.Database)

    $_manifest = Get-Content -LiteralPath (Join-Path $script:Destination 'SHA256SUMS')
    $_manifest.Count | Should -Be 3
    foreach ($_line in $_manifest) {
      $_line | Should -Match '^[0-9a-f]{64} \*(backups|snapshots)/'
      $_hash, $_relative = $_line -split ' \*', 2
      Get-TestHash (Join-Path $script:Destination $_relative) | Should -Be $_hash
    }
    @(Get-ChildItem -LiteralPath $script:Destination -Recurse -Filter '*.tmp').Count | Should -Be 0
  }

  It 'copies nothing on a second run while the database is unchanged' {
    $null = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database -BackupPath $script:Source.Backups 6>$null
    $_results = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database -BackupPath $script:Source.Backups -PassThru 6>$null
    @($_results | Where-Object Status -EQ 'Completed').Count | Should -Be 0
    ($_results | Where-Object Action -EQ 'Snapshot').Detail | Should -Match 'Unchanged since snapshot'
    @(Get-ChildItem -LiteralPath (Join-Path $script:Destination 'snapshots') -File).Count | Should -Be 1
    (Get-Content -LiteralPath (Join-Path $script:Destination 'SHA256SUMS')).Count | Should -Be 3
  }

  It 'takes a new snapshot after the database is saved again' {
    $null = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database 6>$null
    Write-FakeKdbx -Path $script:Source.Database -Content 'live version 2'
    (Get-Item -LiteralPath $script:Source.Database).LastWriteTime = (Get-Date).AddMinutes(5)
    $null = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database 6>$null
    @(Get-ChildItem -LiteralPath (Join-Path $script:Destination 'snapshots') -File).Count | Should -Be 2
  }

  It 'never overwrites a backup when the source file was altered, and exits 1' {
    $null = & $script:BackupScript -Destination $script:Destination -BackupPath $script:Source.Backups 6>$null
    $_backedUp = Join-Path $script:Destination 'backups\Vault_01_05_2026_10-00.old.kdbx'
    $_goodHash = Get-TestHash $_backedUp

    # Simulate the source being encrypted in place: same name, newer, different content.
    Write-FakeKdbx -NoSignature -Path (Join-Path $script:Source.Backups 'Vault_01_05_2026_10-00.old.kdbx') -Content 'ENCRYPTED BY RANSOMWARE'
    Write-FakeKdbx -Path (Join-Path $script:Source.Backups 'Vault_03_05_2026_09-15.old.kdbx') -Content 'old version C'

    $global:LASTEXITCODE = 0
    $_results = & $script:BackupScript -Destination $script:Destination -BackupPath $script:Source.Backups -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    # The encrypted file also fails the KDBX signature check, which is reported first.
    @($_results | Where-Object { $_.Target -eq 'Vault_01_05_2026_10-00.old.kdbx' -and $_.Action -eq 'Backup' } | ForEach-Object Status) | Should -Be @('Warn', 'Conflict')
    Get-TestHash $_backedUp | Should -Be $_goodHash
    ($_results | Where-Object Target -EQ 'Vault_03_05_2026_09-15.old.kdbx').Status | Should -Be 'Completed'
  }

  It 'fails instead of creating the destination when its parent is missing' {
    $_unmounted = Join-Path $script:Root 'not-mounted\Backups\KeePass'
    $global:LASTEXITCODE = 0
    $_results = & $script:BackupScript -Destination $_unmounted -BackupPath $script:Source.Backups -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object Action -EQ 'Validate').Status | Should -Be 'Failed'
    Test-Path -LiteralPath (Join-Path $script:Root 'not-mounted') | Should -BeFalse
  }

  It 'refuses to snapshot a live database without a KDBX signature' {
    Write-FakeKdbx -NoSignature -Path $script:Source.Database -Content 'not a database'
    $global:LASTEXITCODE = 0
    $_results = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object Action -EQ 'Snapshot').Status | Should -Be 'Failed'
    Test-Path -LiteralPath (Join-Path $script:Destination 'snapshots') | Should -BeFalse
  }

  It 'still copies a backup file without a KDBX signature, with a warning' {
    Write-FakeKdbx -NoSignature -Path (Join-Path $script:Source.Backups 'Vault_broken.old.kdbx') -Content 'truncated'
    $_results = & $script:BackupScript -Destination $script:Destination -BackupPath $script:Source.Backups -PassThru 6>$null
    @($_results | Where-Object { $_.Target -eq 'Vault_broken.old.kdbx' } | ForEach-Object Status) | Should -Be @('Warn', 'Completed')
  }

  It 'rejects two databases whose snapshots would collide' {
    $_other = Join-Path $script:Root 'elsewhere\Vault.kdbx'
    Write-FakeKdbx -Path $_other -Content 'another vault'
    $global:LASTEXITCODE = 0
    $_results = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database, $_other -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    @($_results | Where-Object { $_.Action -eq 'Snapshot' -and $_.Status -eq 'Failed' }).Count | Should -Be 2
  }

  It 'requires at least one source' {
    $global:LASTEXITCODE = 0
    $null = & $script:BackupScript -Destination $script:Destination 6>$null
    $global:LASTEXITCODE | Should -Be 1
  }

  It 'writes nothing in a dry run' {
    $_results = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database -BackupPath $script:Source.Backups -DryRun 6>$null
    @($_results | Where-Object Status -EQ 'Skipped').Count | Should -Be 3
    @(Get-ChildItem -LiteralPath $script:Destination -Recurse).Count | Should -Be 0
  }

  It 'verifies the destination and detects a modified backup' {
    $null = & $script:BackupScript -Destination $script:Destination -DatabasePath $script:Source.Database -BackupPath $script:Source.Backups 6>$null

    $global:LASTEXITCODE = 0
    $null = & $script:BackupScript -Destination $script:Destination -Verify 6>$null
    $global:LASTEXITCODE | Should -Be 0

    Write-FakeKdbx -Path (Join-Path $script:Destination 'backups\Vault_02_05_2026_11-30.old.kdbx') -Content 'bit rot'
    Write-FakeKdbx -Path (Join-Path $script:Destination 'backups\Unrecorded.old.kdbx') -Content 'dropped in by hand'
    $global:LASTEXITCODE = 0
    $_results = & $script:BackupScript -Destination $script:Destination -Verify -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 1
    ($_results | Where-Object Target -EQ 'backups/Vault_02_05_2026_11-30.old.kdbx').Status | Should -Be 'Failed'
    ($_results | Where-Object Target -EQ 'backups/Unrecorded.old.kdbx').Status | Should -Be 'Warn'
  }
}

Describe 'New-DriveMarker' {
  BeforeEach {
    $script:Drive = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Drive -Force
    $script:Marker = Join-Path $script:Drive '.backup.marker'
  }

  It 'writes a hidden marker whose first line is the ID it reports' {
    $_results = & $script:MarkerScript -DriveRoot $script:Drive -PassThru 6>$null
    $_id = ($_results | Where-Object Status -EQ 'Completed').MarkerId
    [guid]::Parse($_id) | Should -Not -Be ([guid]::Empty)
    (Get-Content -LiteralPath $script:Marker -TotalCount 1) | Should -Be $_id
    ((Get-Item -LiteralPath $script:Marker -Force).Attributes -band [System.IO.FileAttributes]::Hidden) | Should -Not -Be 0
    @(Get-ChildItem -LiteralPath $script:Drive -File).Count | Should -Be 0 -Because 'the marker stays out of an ordinary directory listing'
  }

  It 'reuses a given ID' {
    $_id = [guid]::NewGuid()
    $null = & $script:MarkerScript -DriveRoot $script:Drive -MarkerId $_id 6>$null
    (Get-Content -LiteralPath $script:Marker -TotalCount 1) | Should -Be "$_id"
  }

  It 'refuses to replace an existing marker without -Force' {
    $_original = [guid]::NewGuid()
    $null = & $script:MarkerScript -DriveRoot $script:Drive -MarkerId $_original 6>$null

    $global:LASTEXITCODE = 0
    $null = & $script:MarkerScript -DriveRoot $script:Drive 6>$null
    $global:LASTEXITCODE | Should -Be 1
    (Get-Content -LiteralPath $script:Marker -TotalCount 1) | Should -Be "$_original"

    # Replacing a hidden marker only works if its attributes are cleared first.
    $_replacement = [guid]::NewGuid()
    $global:LASTEXITCODE = 0
    $null = & $script:MarkerScript -DriveRoot $script:Drive -MarkerId $_replacement -Force 6>$null
    $global:LASTEXITCODE | Should -Be 0
    (Get-Content -LiteralPath $script:Marker -TotalCount 1) | Should -Be "$_replacement"
    ((Get-Item -LiteralPath $script:Marker -Force).Attributes -band [System.IO.FileAttributes]::Hidden) | Should -Not -Be 0
  }
}

Describe 'Backup-KdbxDatabaseToDrive' {
  BeforeEach {
    $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $script:Source = Initialize-KdbxSource -Root $script:Root
    $script:MarkerId = [guid]::NewGuid()
    $script:Drives = @{}
    foreach ($_name in 'matching', 'other', 'unmarked') {
      $script:Drives[$_name] = Join-Path $script:Root "drive-$_name"
      $null = New-Item -ItemType Directory -Path $script:Drives[$_name] -Force
    }
    $null = & $script:MarkerScript -DriveRoot $script:Drives.matching -MarkerId $script:MarkerId 6>$null
    $null = & $script:MarkerScript -DriveRoot $script:Drives.other 6>$null
  }

  It 'backs up only to drives whose marker holds the given ID' {
    $global:LASTEXITCODE = 0
    $_results = & $script:DriveScript -MarkerId $script:MarkerId -DatabasePath $script:Source.Database -BackupPath $script:Source.Backups -DriveRoot $script:Drives.Values -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 0

    @(Get-ChildItem -LiteralPath (Join-Path $script:Drives.matching 'KeePass-Backups\backups') -File).Count | Should -Be 2
    Test-Path -LiteralPath (Join-Path $script:Drives.other 'KeePass-Backups') | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $script:Drives.unmarked 'KeePass-Backups') | Should -BeFalse
    ($_results | Where-Object { $_.Target -eq $script:Drives.other -and $_.Action -eq 'Discover' }).Status | Should -Be 'Skipped'
  }

  It 'skips without error when no marked drive is attached' {
    $global:LASTEXITCODE = 0
    $_results = & $script:DriveScript -MarkerId ([guid]::NewGuid()) -BackupPath $script:Source.Backups -DriveRoot $script:Drives.unmarked -PassThru 6>$null
    $global:LASTEXITCODE | Should -Be 0
    ($_results | Where-Object Target -EQ 'Drives').Detail | Should -Match 'New-DriveMarker\.ps1'
  }

  It 'skips a marked drive that holds one of the sources' {
    $null = & $script:MarkerScript -DriveRoot $script:Root -MarkerId $script:MarkerId 6>$null
    $_results = & $script:DriveScript -MarkerId $script:MarkerId -BackupPath $script:Source.Backups -DriveRoot $script:Root -PassThru 6>$null
    ($_results | Where-Object { $_.Action -eq 'Discover' -and $_.Target -eq $script:Root }).Detail | Should -Match 'holds a source'
    Test-Path -LiteralPath (Join-Path $script:Root 'KeePass-Backups') | Should -BeFalse
  }

  It 'passes -Verify through to each marked drive' {
    $null = & $script:DriveScript -MarkerId $script:MarkerId -BackupPath $script:Source.Backups -DriveRoot $script:Drives.matching 6>$null
    Write-FakeKdbx -Path (Join-Path $script:Drives.matching 'KeePass-Backups\backups\Vault_01_05_2026_10-00.old.kdbx') -Content 'tampered'
    $global:LASTEXITCODE = 0
    $null = & $script:DriveScript -MarkerId $script:MarkerId -DriveRoot $script:Drives.matching -Verify 6>$null
    $global:LASTEXITCODE | Should -Be 1
  }

  It 'rejects a destination outside the drive' {
    { & $script:DriveScript -MarkerId $script:MarkerId -BackupPath $script:Source.Backups -DriveRoot $script:Drives.matching -Destination '..\escape' 6>$null } | Should -Throw
  }

  It 'finds the hidden marker written by New-DriveMarker.ps1' {
    ((Get-Item -LiteralPath (Join-Path $script:Drives.matching '.backup.marker') -Force).Attributes -band [System.IO.FileAttributes]::Hidden) | Should -Not -Be 0
    $_results = & $script:DriveScript -MarkerId $script:MarkerId -BackupPath $script:Source.Backups -DriveRoot $script:Drives.matching -PassThru 6>$null
    ($_results | Where-Object { $_.Target -eq $script:Drives.matching -and $_.Action -eq 'Backup' }).Status | Should -Be 'Completed'
  }
}
