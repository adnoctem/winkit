#Requires -Version 5.0

function Resolve-LGPOSource {
  <#
    .SYNOPSIS
      Returns metadata for a known LGPO source location.
    .DESCRIPTION
      Pure data lookup. No I/O. Centralises the "where do we get LGPO from"
      decision so that every other function in this module can reference it
      without duplicating URLs or hashes.

      When Microsoft moves the file or publishes a new SCT release, update
      the table below and have a human review the diff. This function is the
      single point of change for supply-chain trust.
    .PARAMETER Source
      Source identifier. Currently only 'SCT-LGPO-Standalone' is recognised.
    .EXAMPLE
      PS> Resolve-LGPOSource
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('SCT-LGPO-Standalone')]
    [string]
    $Source = 'SCT-LGPO-Standalone'
  )

  $sources = @{
    'SCT-LGPO-Standalone' = [PSCustomObject]@{
      Name = 'Security Compliance Toolkit - LGPO standalone'
      Url = 'https://download.microsoft.com/download/8/5/C/85C25433-A1B0-4FFA-9429-7E023E7DA8D8/LGPO.zip'
      Sha256 = 'PLACEHOLDER_REPLACE_ON_FIRST_VENDORING'
      ExpectedBinaryPath = 'LGPO_30/LGPO.exe'
      LastVerified = '2026-06-14'
    }
  }

  if (-not $sources.ContainsKey($Source)) {
    throw "Unknown LGPO source '$Source'."
  }
  return $sources[$Source]
}

function Test-LGPOSourceAvailability {
  <#
    .SYNOPSIS
      Verifies the LGPO download URL is still reachable.
    .DESCRIPTION
      Issues a HEAD request to the URL returned by Resolve-LGPOSource.
      Does NOT download or verify the content - that's Install-LGPO's job.

      Intended for weekly CI runs. The output object is suitable for
      serialising to JSON and archiving as ISO supply-chain evidence
      ("we monitor external dependencies weekly").
    .PARAMETER Source
      Source identifier forwarded to Resolve-LGPOSource.
    .EXAMPLE
      PS> Test-LGPOSourceAvailability
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('SCT-LGPO-Standalone')]
    [string]
    $Source = 'SCT-LGPO-Standalone'
  )

  $info = Resolve-LGPOSource -Source $Source
  try {
    $response = Invoke-WebRequest -Uri $info.Url -Method Head -UseBasicParsing -ErrorAction Stop
    [PSCustomObject]@{
      Source = $Source
      Url = $info.Url
      Available = $true
      StatusCode = [int]$response.StatusCode
      ContentLength = $response.Headers['Content-Length']
      CheckedAt = (Get-Date).ToUniversalTime()
    }
  }
  catch {
    [PSCustomObject]@{
      Source = $Source
      Url = $info.Url
      Available = $false
      Error = $_.Exception.Message
      CheckedAt = (Get-Date).ToUniversalTime()
    }
  }
}

function Install-LGPO {
  <#
    .SYNOPSIS
      Downloads, verifies, and installs LGPO.exe to a controllable path.
    .DESCRIPTION
      Hash-verified install. Refuses to proceed if the SHA-256 of the
      downloaded archive does not match Resolve-LGPOSource's recorded hash.

      Idempotent: re-running when LGPO.exe is already present at the
      destination returns the existing path unless -Force is specified.

      Writing to the default destination (%ProgramData%) requires
      administrator elevation. A non-elevated session can specify an
      alternate -Destination within the user's writeable scope.
    .PARAMETER Destination
      Directory where LGPO.exe ends up. Defaults to %ProgramData%\winkit\tools.
      Non-elevated callers should supply a user-writeable path.
    .PARAMETER Source
      Source identifier forwarded to Resolve-LGPOSource.
    .PARAMETER Force
      Re-download and re-install even if LGPO.exe is already present.
    .EXAMPLE
      PS> Install-LGPO
    .EXAMPLE
      PS> Install-LGPO -Force -Verbose
    .EXAMPLE
      PS> Install-LGPO -Destination "$env:LOCALAPPDATA\winkit\tools"
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $Destination = (Join-Path -Path $env:ProgramData -ChildPath 'winkit\tools'),

    [Parameter(Mandatory = $false)]
    [ValidateSet('SCT-LGPO-Standalone')]
    [string]
    $Source = 'SCT-LGPO-Standalone',

    [Parameter(Mandatory = $false)]
    [switch]
    $Force
  )

  $info = Resolve-LGPOSource -Source $Source
  $exePath = Join-Path -Path $Destination -ChildPath 'LGPO.exe'

  if (Test-Path -LiteralPath $exePath -PathType Leaf) {
    if (-not $Force) {
      Write-Verbose "LGPO.exe already present at '$exePath'; skipping download."
      return $exePath
    }
    Write-Verbose "LGPO.exe already present at '$exePath'; -Force supplied, re-downloading."
  }

  if (-not $PSCmdlet.ShouldProcess($Destination, "Install LGPO from $($info.Url)")) {
    return
  }

  if (-not (Read-ProcessElevation)) {
    Write-Error "Writing to '$Destination' requires administrator rights. Run the session elevated or supply a user-writeable -Destination such as '$(Join-Path -Path $env:LOCALAPPDATA -ChildPath 'winkit\tools')'."
    return
  }

  $null = New-Item -Path $Destination -ItemType Directory -Force -ErrorAction SilentlyContinue
  $zipPath = Join-Path -Path $env:TEMP -ChildPath "LGPO_$(New-Guid).zip"
  $extractDir = Join-Path -Path $env:TEMP -ChildPath "LGPO_extract_$(New-Guid)"

  try {
    Write-Verbose "Downloading LGPO from $($info.Url)..."
    Invoke-WebRequest -Uri $info.Url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop

    $actualHash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash
    $expectedHash = $info.Sha256.ToUpperInvariant()
    if ($actualHash -ne $expectedHash) {
      throw "SHA-256 mismatch for $($info.Url). Expected '$expectedHash', got '$actualHash'. Refusing to install."
    }

    Write-Verbose 'SHA-256 hash matches. Extracting archive...'
    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
    $found = Get-ChildItem -Path $extractDir -Recurse -Filter 'LGPO.exe' | Select-Object -First 1
    if (-not $found) {
      throw "LGPO.exe not found in archive at $($info.Url)."
    }
    Copy-Item -Path $found.FullName -Destination $exePath -Force
    Write-Verbose "Installed LGPO.exe to '$exePath'."
  }
  finally {
    if (Test-Path -LiteralPath $zipPath -PathType Leaf) {
      Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $extractDir -PathType Container) {
      Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  return $exePath
}

function Test-LGPOInstalled {
  <#
    .SYNOPSIS
      Returns $true if LGPO.exe is present at the expected path.
    .PARAMETER Path
      Full path to LGPO.exe. Defaults to %ProgramData%\winkit\tools\LGPO.exe.
    .EXAMPLE
      PS> if (Test-LGPOInstalled) { Invoke-LGPO -PolicyPath .\policy.txt }
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $Path = (Join-Path -Path $env:ProgramData -ChildPath 'winkit\tools\LGPO.exe')
  )

  return (Test-Path -LiteralPath $Path -PathType Leaf)
}

function Invoke-LGPO {
  <#
    .SYNOPSIS
      Applies a policy file or GPO backup directory using LGPO.exe.
    .DESCRIPTION
      Accepts a path to either a policy text file (applied via /t) or a
      directory containing a GPO backup (applied via /g). The argument is
      auto-detected based on whether PolicyPath is a file or directory.

      Returns a structured object describing the apply, suitable for
      logging as audit evidence.
    .PARAMETER PolicyPath
      Path to a policy text file or a directory containing a GPO backup.
    .PARAMETER LgpoExe
      Path to LGPO.exe. Defaults to %ProgramData%\winkit\tools\LGPO.exe.
    .EXAMPLE
      PS> Invoke-LGPO -PolicyPath .\resources\policies\01-telemetry.txt
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [string]
    $PolicyPath,

    [Parameter(Mandatory = $false)]
    [string]
    $LgpoExe = (Join-Path -Path $env:ProgramData -ChildPath 'winkit\tools\LGPO.exe')
  )

  if (-not (Test-Path -LiteralPath $LgpoExe -PathType Leaf)) {
    throw "LGPO.exe not found at '$LgpoExe'. Run Install-LGPO first."
  }

  $isDirectory = Test-Path -LiteralPath $PolicyPath -PathType Container
  $arg = if ($isDirectory) { '/g' } else { '/t' }

  if (-not $PSCmdlet.ShouldProcess($PolicyPath, "Apply via LGPO.exe ($arg)")) {
    return
  }

  $stdoutFile = Join-Path -Path $env:TEMP -ChildPath "lgpo_stdout_$(New-Guid).log"
  $stderrFile = Join-Path -Path $env:TEMP -ChildPath "lgpo_stderr_$(New-Guid).log"
  try {
    $proc = Start-Process -FilePath $LgpoExe `
      -ArgumentList @($arg, $PolicyPath) `
      -RedirectStandardOutput $stdoutFile `
      -RedirectStandardError $stderrFile `
      -Wait -NoNewWindow -PassThru

    [PSCustomObject]@{
      PolicyPath = $PolicyPath
      Mode = if ($isDirectory) { 'GpoBackup' } else { 'TextSource' }
      ExitCode = $proc.ExitCode
      StdOut = (Get-Content -LiteralPath $stdoutFile -Raw -ErrorAction SilentlyContinue)
      StdErr = (Get-Content -LiteralPath $stderrFile -Raw -ErrorAction SilentlyContinue)
      AppliedAt = (Get-Date).ToUniversalTime()
      Success = ($proc.ExitCode -eq 0)
    }
  }
  finally {
    if (Test-Path -LiteralPath $stdoutFile -PathType Leaf) {
      Remove-Item -LiteralPath $stdoutFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $stderrFile -PathType Leaf) {
      Remove-Item -LiteralPath $stderrFile -Force -ErrorAction SilentlyContinue
    }
  }
}
