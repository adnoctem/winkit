Set-StrictMode -Version Latest

function Test-Elevation {
  <#
    .SYNOPSIS
      Test-Elevation - Returns whether the current process has administrator privileges.
    .DESCRIPTION
      Checks the Windows principal of the current identity for membership in the
      built-in Administrators role. Returns $true when the process holds an
      elevated (administrator) token, $false otherwise.

      This is a pure predicate: it makes no changes and never elevates. Use
      Request-AdministratorPrivilege to actually elevate.
    .OUTPUTS
      System.Boolean
    .EXAMPLE
      PS> if (-not (Test-Elevation)) { throw 'Run me elevated.' }
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param ()

  # winkit is Windows-only; fail clearly rather than throwing a cryptic type
  # error if this is ever invoked on a non-Windows host (the [Security.Principal]
  # Windows types are present on PS7/Windows but the call throws elsewhere).
  if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
    throw 'Test-Elevation is only supported on Windows.'
  }

  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = [Security.Principal.WindowsPrincipal]::new($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Backwards-compatible alias for the original name. Remove once callers are
# migrated to Test-Elevation.
# CHANGE-NOTE: drop this alias after updating references.
Set-Alias -Name Read-ProcessElevation -Value Test-Elevation

function Request-AdministratorPrivilege {
  <#
    .SYNOPSIS
      Request-AdministratorPrivilege - Ensures the current script runs elevated,
      re-launching it under UAC if necessary.
    .DESCRIPTION
      Intended to be called once near the top of an entry-point script (e.g.
      Invoke-Optimizer.ps1, Invoke-Bootstrap.ps1) as a self-elevation block.

      Behaviour:
        * If already elevated, returns immediately and the script continues.
        * If not elevated, re-launches the SAME PowerShell host with the SAME
          script and arguments under the RunAs verb (triggering UAC), waits for
          the elevated copy to finish, and exits the current (non-elevated)
          process with the elevated copy's exit code.

      Hardening over a naive self-elevation block:
        * Re-launches the actual current host (pwsh vs powershell), not a
          hardcoded powershell.exe.
        * Preserves the original working directory across the RunAs boundary
          (RunAs otherwise starts the child in system32, breaking relative paths).
        * Faithfully reconstructs the original invocation (including bound
          parameters and switches) so -Profile, -WhatIf, paths with spaces, etc.
          survive intact.
        * Propagates the elevated child's exit code to the original caller, so
          .cmd launchers checking %ERRORLEVEL% and CI see the real result.
        * Handles UAC denial gracefully (clear message + non-zero exit) instead
          of an unhandled Win32Exception.
        * Guards against an infinite re-launch loop via an environment marker,
          in case elevation "succeeds" but the token still isn't elevated
          (rare UAC/GPO configurations).

      NOTE: this function calls exit on the non-elevated path. It is designed for
      entry-point scripts, not for dot-sourced library use — calling it from an
      interactive session would terminate that session when not elevated.
    .PARAMETER ScriptPath
      Path to the script to re-launch. Defaults to the caller's own path
      ($PSCommandPath of the calling script). Override only for unusual hosting.
    .PARAMETER BoundParameters
      The calling script's $PSBoundParameters, used to faithfully reconstruct
      named parameters and switches for the elevated re-launch. Strongly
      recommended; pass $PSBoundParameters from the caller.
    .PARAMETER ArgumentList
      Any unbound/positional arguments ($args from the caller) to append after
      the reconstructed bound parameters.
    .PARAMETER IsElevatedRelaunch
      Loop guard. The caller passes this as $true only when its own param block
      received the re-launch marker, indicating THIS process is already the
      elevated child. If set and the process is still not elevated, the function
      aborts instead of spawning again. See the usage example for the pattern.
    .OUTPUTS
      None. Either returns (already elevated) or exits the process (re-launched).
    .EXAMPLE
      PS> # At the top of Invoke-Optimizer.ps1, whose param block includes a
      PS> # hidden [switch]$Elevated used purely as the re-launch marker:
      PS> Request-AdministratorPrivilege `
      PS>     -BoundParameters $PSBoundParameters `
      PS>     -ArgumentList $args `
      PS>     -IsElevatedRelaunch:$Elevated
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>
  [CmdletBinding()]
  [OutputType([void])]
  param (
    [string]$ScriptPath = $MyInvocation.PSCommandPath,

    [System.Collections.IDictionary]$BoundParameters,

    [object[]]$ArgumentList,

    [switch]$IsElevatedRelaunch
  )

  if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
    throw 'Request-AdministratorPrivilege is only supported on Windows.'
  }

  # Already elevated — nothing to do.
  if (Test-Elevation) {
    return
  }

  # Loop guard: if the caller told us this is already the elevated re-launch
  # (via its own -Elevated marker switch) and we're STILL not elevated, stop
  # rather than spawning forever. We pass the marker as an explicit parameter
  # rather than an environment variable (which doesn't survive the RunAs
  # boundary) or a bare argument token (which a CmdletBinding param block would
  # reject) — an explicit switch the caller declares is unambiguous.
  if ($IsElevatedRelaunch) {
    Write-Error ('Elevation was attempted but the process is still not elevated. ' +
      'Check UAC / Group Policy settings (e.g. "Run all administrators in Admin Approval Mode"). ' +
      'Aborting to avoid a re-launch loop.')
    exit 1
  }

  # RunAs minimum: Windows Vista (build 6000). Earlier versions can't elevate.
  $build = [int](Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop |
      Select-Object -ExpandProperty BuildNumber)
  if ($build -lt 6000) {
    throw "Self-elevation requires Windows Vista or later (build >= 6000); detected build $build."
  }

  # Resolve the script to re-launch.
  if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    throw 'Could not determine the script path to re-launch. Pass -ScriptPath explicitly.'
  }
  $ScriptPath = (Resolve-Path -LiteralPath $ScriptPath -ErrorAction Stop).Path

  # Re-launch the SAME host (pwsh.exe or powershell.exe), not a hardcoded one.
  $hostExe = (Get-Process -Id $PID).Path
  if ([string]::IsNullOrWhiteSpace($hostExe)) {
    throw 'Could not determine the current PowerShell host executable path.'
  }

  # Faithfully reconstruct the argument list for the elevated re-launch.
  $reArgs = [System.Collections.Generic.List[string]]::new()
  $reArgs.Add('-NoProfile')
  $reArgs.Add('-ExecutionPolicy'); $reArgs.Add('Bypass')
  $reArgs.Add('-File'); $reArgs.Add($ScriptPath)

  if ($BoundParameters) {
    foreach ($name in $BoundParameters.Keys) {
      $value = $BoundParameters[$name]
      if ($value -is [System.Management.Automation.SwitchParameter]) {
        # Switches: include the flag only when present/true. Use -Name:$true form
        # so the elevated copy receives an explicit value, avoiding ambiguity.
        if ($value.IsPresent) { $reArgs.Add("-$name") }
      }
      elseif ($value -is [bool]) {
        $reArgs.Add("-$name`:$([bool]$value)")
      }
      elseif ($null -ne $value) {
        # Arrays -> repeat the parameter for each element so [string[]] params
        # round-trip correctly (e.g. -VCRedistVersions 14.0,12.0).
        foreach ($item in @($value)) {
          $reArgs.Add("-$name")
          $reArgs.Add([string]$item)
        }
      }
    }
  }

  if ($ArgumentList) {
    foreach ($a in $ArgumentList) { $reArgs.Add([string]$a) }
  }

  # Inject the elevation marker switch so the re-launched child can pass it back
  # into this function as -IsElevatedRelaunch and trip the loop guard if needed.
  # The caller's param block must declare a [switch]$Elevated for this to bind.
  if ($BoundParameters -and -not $BoundParameters.Contains('Elevated')) {
    $reArgs.Add('-Elevated')
  }
  elseif (-not $BoundParameters) {
    $reArgs.Add('-Elevated')
  }

  # Preserve the working directory across the RunAs boundary. Start-Process
  # -Verb RunAs otherwise launches in system32, breaking relative paths the
  # elevated script might use.
  $workingDir = (Get-Location -PSProvider FileSystem).ProviderPath

  $startInfo = @{
    FilePath = $hostExe
    ArgumentList = $reArgs.ToArray()
    Verb = 'RunAs'
    WorkingDirectory = $workingDir
    PassThru = $true
    Wait = $true
  }

  try {
    $proc = Start-Process @startInfo
    # Propagate the elevated child's exit code to our caller.
    exit $proc.ExitCode
  }
  catch [System.ComponentModel.Win32Exception] {
    # 1223 = ERROR_CANCELLED — user clicked "No" on the UAC prompt.
    if ($_.Exception.NativeErrorCode -eq 1223) {
      Write-Error 'Elevation was cancelled by the user. Administrator privileges are required to continue.'
      exit 1223
    }
    throw
  }
}
