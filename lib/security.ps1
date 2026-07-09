function Get-DefenderThreatDetection {
  <#
    .SYNOPSIS
      Retrieves Microsoft Defender threat detections, optionally filtered by date.
    .DESCRIPTION
      Wraps Get-MpThreatDetection.  -Date sets a cutoff -- only detections with an
      InitialDetectionTime on or after that point are returned.  -OutputPath and
      -OutputFormat control whether results are printed to the terminal or written
      to a file (TXT or JSON).
    .PARAMETER Date
      Cutoff date for detections.  Accepts any value that Get-Date can parse
      (string, DateTime, etc.).  Defaults to right now.
    .PARAMETER OutputPath
      File path to write results to.  When omitted, results are printed to the
      terminal.
    .PARAMETER OutputFormat
      Output format: TXT (Formatted-List) or JSON.  Defaults to TXT.
    .PARAMETER IncludeURLs
      Augment each detection with a ThreatDescriptionURL property pointing to the
      official Microsoft threat encyclopedia entry.
    .EXAMPLE
      Get-DefenderThreatDetection
      Prints all threat detections to the terminal.
    .EXAMPLE
      Get-DefenderThreatDetection -Date '2026-04-01' -OutputPath '.\detections.json' -OutputFormat JSON
      Writes detections since April 1st 2026 as JSON.
    .EXAMPLE
      Get-DefenderThreatDetection -IncludeURLs -OutputPath '.\detections.json' -OutputFormat JSON
      Writes detections as JSON, each augmented with a ThreatDescriptionURL.
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [object]
    $Date = (Get-Date),

    [Parameter(Mandatory = $false)]
    [string]
    $OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('TXT', 'JSON')]
    [string]
    $OutputFormat = 'TXT',

    [Parameter(Mandatory = $false)]
    [switch]
    $IncludeURLs
  )

  $cutoffDate = if ($Date -is [datetime]) { $Date } else { Get-Date $Date }
  Write-Log -Message "Filtering Defender threat detections since $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))" -Color Yellow

  $detections = Get-MpThreatDetection | Where-Object { $_.InitialDetectionTime -ge $cutoffDate }

  if ($IncludeURLs -and $detections.Count -gt 0) {
    $detections = $detections | ForEach-Object {
      $_urlName = if ($_.PSObject.Properties.Name -contains 'ThreatName') { $_.ThreatName } else { $_.Name }
      $_ | Add-Member -NotePropertyName 'ThreatDescriptionURL' -NotePropertyValue (Get-DefenderThreatDescriptionURL -ThreatName $_urlName) -PassThru
    }
    Write-Log -Message '  -> ThreatDescriptionURL(s) appended' -Color Gray
  }
  Write-Log -Message "  -> $($detections.Count) detection(s) found" -Color Gray

  if ($PSBoundParameters.ContainsKey('OutputPath') -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $_outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    switch ($OutputFormat) {
      'JSON' {
        $detections | ConvertTo-Json -Depth 3 | Out-File -FilePath $_outPath -Encoding utf8
      }
      'TXT' {
        $detections | Format-List * | Out-String -Width 4096 | Out-File -FilePath $_outPath -Encoding utf8
      }
    }
    Write-Log -Message "  -> Written to: $_outPath" -Color Green
  }
  else {
    $detections | Format-List *
  }
}

function Get-DefenderThreat {
  <#
    .SYNOPSIS
      Retrieves the full Microsoft Defender threat catalog.
    .DESCRIPTION
      Wraps Get-MpThreat.  -OutputPath and -OutputFormat control whether results
      are printed to the terminal or written to a file (TXT or JSON).
    .PARAMETER OutputPath
      File path to write results to.  When omitted, results are printed to the
      terminal.
    .PARAMETER OutputFormat
      Output format: TXT (Formatted-List) or JSON.  Defaults to TXT.
    .PARAMETER IncludeURLs
      Augment each threat with a ThreatDescriptionURL property pointing to the
      official Microsoft threat encyclopedia entry.
    .EXAMPLE
      Get-DefenderThreat
      Prints the threat catalog to the terminal.
    .EXAMPLE
      Get-DefenderThreat -OutputPath '.\threats.json' -OutputFormat JSON
      Writes the threat catalog as JSON.
    .EXAMPLE
      Get-DefenderThreat -IncludeURLs -OutputFormat JSON
      Prints the threat catalog as JSON, each augmented with a ThreatDescriptionURL.
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('TXT', 'JSON')]
    [string]
    $OutputFormat = 'TXT',

    [Parameter(Mandatory = $false)]
    [switch]
    $IncludeURLs
  )

  Write-Log -Message 'Retrieving Microsoft Defender threat catalog' -Color Yellow

  $threats = Get-MpThreat

  if ($IncludeURLs -and $threats.Count -gt 0) {
    $threats = $threats | ForEach-Object {
      $_urlName = if ($_.PSObject.Properties.Name -contains 'ThreatName') { $_.ThreatName } else { $_.Name }
      $_ | Add-Member -NotePropertyName 'ThreatDescriptionURL' -NotePropertyValue (Get-DefenderThreatDescriptionURL -ThreatName $_urlName) -PassThru
    }
    Write-Log -Message '  -> ThreatDescriptionURL(s) appended' -Color Gray
  }
  Write-Log -Message "  -> $($threats.Count) threat(s) found" -Color Gray

  if ($PSBoundParameters.ContainsKey('OutputPath') -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $_outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    switch ($OutputFormat) {
      'JSON' {
        $threats | ConvertTo-Json -Depth 3 | Out-File -FilePath $_outPath -Encoding utf8
      }
      'TXT' {
        $threats | Format-List * | Out-String -Width 4096 | Out-File -FilePath $_outPath -Encoding utf8
      }
    }
    Write-Log -Message "  -> Written to: $_outPath" -Color Green
  }
  else {
    $threats | Format-List *
  }
}

function Get-DefenderThreatDescriptionURL {
  <#
    .SYNOPSIS
      Builds the public Microsoft Defender threat description URL for a given
      threat family name.
    .DESCRIPTION
      Takes a threat name (e.g. 'Trojan:Win32/Emotet'), URL-encodes it, and
      returns the full HTTP/S link to the official WDSI (Windows Defender
      Security Intelligence) threat encyclopedia entry.
    .PARAMETER ThreatName
      The human-readable threat family name, exactly as reported by
      Get-MpThreat (Name property) or Get-MpThreatDetection (ThreatName).
    .EXAMPLE
      Get-DefenderThreatDescriptionURL -ThreatName 'Trojan:Win32/Emotet'
      https://www.microsoft.com/en-us/wdsi/threats/threat/Trojan%3AWin32%2FEmotet
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string]
    $ThreatName
  )

  process {
    $encoded = [System.Web.HttpUtility]::UrlEncode($ThreatName)
    return "https://www.microsoft.com/en-us/wdsi/threats/threat/$encoded"
  }
}

function Add-DefenderExclusion {
  <#
    .SYNOPSIS
      Adds a Microsoft Defender exclusion.
    .DESCRIPTION
      Wraps Add-MpPreference for path, extension, and process exclusions and
      returns a structured operation result instead of writing ad hoc console
      output. The helper supports ShouldProcess so scripts can use -WhatIf and
      -DryRun consistently.
    .PARAMETER Type
      Exclusion kind: Path, Extension, or Process.
    .PARAMETER Value
      Exclusion value to add.
    .EXAMPLE
      PS> Add-DefenderExclusion -Type Process -Value 'wsl.exe'
    .LINK
      https://github.com/adnoctem/winkit/lib/security.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('Path', 'Extension', 'Process')]
    [string]
    $Type,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Value
  )

  $_target = "$Type`: $Value"
  if (-not $PSCmdlet.ShouldProcess($_target, 'Add Microsoft Defender exclusion')) {
    return New-OperationResult -Target $_target -Source 'Defender' -Action 'AddExclusion' -Status 'Skipped' -Detail 'WhatIf'
  }

  try {
    $_parameters = @{
      ErrorAction = 'Stop'
    }

    switch ($Type) {
      'Path' { $_parameters.ExclusionPath = $Value }
      'Extension' { $_parameters.ExclusionExtension = $Value }
      'Process' { $_parameters.ExclusionProcess = $Value }
    }

    Add-MpPreference @_parameters
    New-OperationResult -Target $_target -Source 'Defender' -Action 'AddExclusion' -Status 'Completed' -Detail $Value
  }
  catch {
    New-OperationResult -Target $_target -Source 'Defender' -Action 'AddExclusion' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Enable-WSLFirewallRule {
  <#
    .SYNOPSIS
      Ensures the inbound WSL firewall allow rule exists.
    .DESCRIPTION
      Adds a Windows Firewall rule allowing inbound traffic on the WSL virtual
      Ethernet adapter. Existing rules with the same display name are treated as
      already configured to avoid creating duplicates.
    .PARAMETER DisplayName
      Firewall display name. Defaults to WSL.
    .PARAMETER InterfaceAlias
      Network interface alias for the WSL virtual switch.
    .EXAMPLE
      PS> Enable-WSLFirewallRule
    .LINK
      https://github.com/adnoctem/winkit/lib/security.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [string]
    $DisplayName = 'WSL',

    [string]
    $InterfaceAlias = 'vEthernet (WSL)'
  )

  try {
    $_existing = @(Get-NetFirewallRule -DisplayName $DisplayName -ErrorAction SilentlyContinue)
    if ($_existing.Count -gt 0) {
      return New-OperationResult -Target $DisplayName -Source 'Firewall' -Action 'AllowInbound' -Status 'Skipped' -Detail 'AlreadyExists'
    }

    if (-not $PSCmdlet.ShouldProcess($DisplayName, "Allow inbound traffic on $InterfaceAlias")) {
      return New-OperationResult -Target $DisplayName -Source 'Firewall' -Action 'AllowInbound' -Status 'Skipped' -Detail 'WhatIf'
    }

    $null = New-NetFirewallRule -DisplayName $DisplayName -Direction Inbound -InterfaceAlias $InterfaceAlias -Action Allow -ErrorAction Stop
    New-OperationResult -Target $DisplayName -Source 'Firewall' -Action 'AllowInbound' -Status 'Completed' -Detail $InterfaceAlias
  }
  catch {
    New-OperationResult -Target $DisplayName -Source 'Firewall' -Action 'AllowInbound' -Status 'Failed' -Detail $_.Exception.Message
  }
}

function Disable-JetBrainsFirewallRule {
  <#
    .SYNOPSIS
      Disables public-profile JetBrains IDE firewall rules.
    .DESCRIPTION
      Finds firewall rules attached to the Public profile whose display names
      begin with known JetBrains IDE names and disables them. This supports the
      JetBrains WSL debugging workaround while returning structured operation
      results for every matching rule.
    .PARAMETER Prefix
      Display-name prefixes to match. Defaults to known JetBrains IDE product
      names.
    .EXAMPLE
      PS> Disable-JetBrainsFirewallRule
    .LINK
      https://github.com/adnoctem/winkit/lib/security.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [string[]]
    $Prefix = @('PhpStorm', 'IntelliJ', 'PyCharm', 'RubyMine', 'WebStorm', 'DataGrip', 'GoLand', 'Rider')
  )

  $_results = New-Object System.Collections.ArrayList

  if ($WhatIfPreference) {
    foreach ($_prefix in $Prefix) {
      Add-OperationResult -Results $_results -Target "$_prefix*" -Source 'Firewall' -Action 'Disable' -Status 'Skipped' -Detail 'WhatIf'
    }

    return $_results
  }

  try {
    $_publicProfile = Get-NetFirewallProfile -Name Public -ErrorAction Stop
    $_rules = @($_publicProfile | Get-NetFirewallRule -ErrorAction Stop | Where-Object {
        $_rule = $_
        @($Prefix | Where-Object { $_rule.DisplayName -like "$_*" }).Count -gt 0
      })

    if ($_rules.Count -eq 0) {
      Add-OperationResult -Results $_results -Target 'JetBrainsPublicFirewallRules' -Source 'Firewall' -Action 'Disable' -Status 'Skipped' -Detail 'NoMatch'
      return $_results
    }

    foreach ($_rule in $_rules) {
      if (-not $PSCmdlet.ShouldProcess($_rule.DisplayName, 'Disable public firewall rule')) {
        Add-OperationResult -Results $_results -Target $_rule.DisplayName -Source 'Firewall' -Action 'Disable' -Status 'Skipped' -Detail 'WhatIf'
        continue
      }

      try {
        $null = $_rule | Disable-NetFirewallRule -ErrorAction Stop
        Add-OperationResult -Results $_results -Target $_rule.DisplayName -Source 'Firewall' -Action 'Disable' -Status 'Completed' -Detail 'Public profile rule disabled.'
      }
      catch {
        Add-OperationResult -Results $_results -Target $_rule.DisplayName -Source 'Firewall' -Action 'Disable' -Status 'Failed' -Detail $_.Exception.Message
      }
    }
  }
  catch {
    Add-OperationResult -Results $_results -Target 'JetBrainsPublicFirewallRules' -Source 'Firewall' -Action 'Disable' -Status 'Failed' -Detail $_.Exception.Message
  }

  $_results
}

function Find-NewlyWrittenObject {
  <#
    .SYNOPSIS
      Finds files written near a point in time (e.g. around a Defender detection).
    .DESCRIPTION
      Recursively scans C:\ (or a custom path) for files whose LastWriteTime falls
      within a configurable window around the supplied -Date.  Designed to help
      identify artifacts dropped by malware at the time of a Defender alert.
      Results can be printed to the terminal or exported as TXT / JSON.
    .PARAMETER Date
      Anchor date/time.  Accepts any value that Get-Date can parse (string,
      DateTime, etc.).  Defaults to right now.
    .PARAMETER Before
      Number of hours before the anchor date to include.  Defaults to 2.
    .PARAMETER After
      Number of hours after the anchor date to include.  Defaults to 1.
    .PARAMETER Path
      Root path to search.  Defaults to the system drive (C:\).
    .PARAMETER OutputPath
      File path to write results to.  When omitted, results are printed to the
      terminal.
    .PARAMETER OutputFormat
      Output format: TXT (Formatted custom table) or JSON.  Defaults to TXT.
    .EXAMPLE
      Find-NewlyWrittenObject -Date '2026-04-30 10:15'
      Searches for files written between 08:15 and 11:15 on 2026-04-30.
    .EXAMPLE
      Find-NewlyWrittenObject -Date '2026-04-30' -Before 4 -After 2 -OutputPath '.\artifacts.json' -OutputFormat JSON
      Wider window, exported as JSON.
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [object]
    $Date = (Get-Date),

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 168)]
    [int]
    $Before = 2,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 168)]
    [int]
    $After = 1,

    [Parameter(Mandatory = $false)]
    [string]
    $Path = "$env:SystemDrive\",

    [Parameter(Mandatory = $false)]
    [string]
    $OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('TXT', 'JSON')]
    [string]
    $OutputFormat = 'TXT'
  )

  $anchorDate = if ($Date -is [datetime]) { $Date } else { Get-Date $Date }
  $windowStart = $anchorDate.AddHours(-$Before)
  $windowEnd = $anchorDate.AddHours($After)

  Write-Log -Message "Searching for files written between $($windowStart.ToString('yyyy-MM-dd HH:mm:ss')) and $($windowEnd.ToString('yyyy-MM-dd HH:mm:ss'))" -Color Yellow
  Write-Log -Message "  Root path: $Path" -Color Gray

  $items = Get-ChildItem -LiteralPath $Path -Recurse -ErrorAction SilentlyContinue |
    Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -gt $windowStart -and $_.LastWriteTime -lt $windowEnd } |
    Sort-Object LastWriteTime |
    Select-Object LastWriteTime,
    LastWriteTimeUtc,
    LastAccessTime,
    LastAccessTimeUtc,
    CreationTime,
    CreationTimeUtc,
    Mode,
    IsReadOnly,
    Length,
    Extension,
    FullName

  Write-Log -Message "  -> $($items.Count) file(s) found" -Color Gray

  if ($items.Count -eq 0) { return }

  if ($PSBoundParameters.ContainsKey('OutputPath') -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $_outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    switch ($OutputFormat) {
      'JSON' {
        $items | ConvertTo-Json -Depth 2 | Out-File -FilePath $_outPath -Encoding utf8
      }
      'TXT' {
        $items | Format-Table -AutoSize | Out-String -Width 4096 | Out-File -FilePath $_outPath -Encoding utf8
      }
    }
    Write-Log -Message "  -> Written to: $_outPath" -Color Green
  }
  else {
    $items | Format-Table -AutoSize
  }
}

function Invoke-SafeProcess {
  <#
    .SYNOPSIS
      Runs an external executable and captures stdout/stderr to a file or the pipeline.
    .DESCRIPTION
      Uses System.Diagnostics.Process to invoke an executable with argument list
      and redirects standard output and standard error. The combined output is
      written to -OutputPath. Use -PassThru to return the output as a string
      instead of writing to disk.

      Designed for IR/forensics collection where external tools (reg.exe,
      wevtutil.exe, systeminfo.exe, etc.) need to be called safely and their
      output captured without risking interactive prompts or policy blocks.
    .PARAMETER FilePath
      Executable path (resolved from PATH when a bare name is supplied).
    .PARAMETER ArgumentList
      Array of arguments. Each element is one argument token.
    .PARAMETER OutputPath
      File to write combined stdout + stderr to. When omitted and -PassThru is
      not supplied, output is discarded.
    .PARAMETER PassThru
      Return stdout as a string. When combined with -OutputPath, output is
      both written to disk and returned.
    .EXAMPLE
      PS> Invoke-SafeProcess -FilePath 'whoami.exe' -ArgumentList @('/all') -OutputPath '.\whoami.txt'
    .EXAMPLE
      PS> Invoke-SafeProcess -FilePath 'systeminfo.exe' -PassThru
    .LINK
      https://github.com/adnoctem/winkit/lib/security.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $FilePath,

    [Parameter(Mandatory = $false)]
    [string[]]
    $ArgumentList,

    [Parameter(Mandatory = $false)]
    [string]
    $OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]
    $PassThru
  )

  try {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    if ($ArgumentList) {
      foreach ($arg in $ArgumentList) {
        [void]$psi.ArgumentList.Add($arg)
      }
    }
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    $content = @()
    if (-not [string]::IsNullOrWhiteSpace($stdout)) { $content += $stdout }
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
      $content += "`r`n--- STDERR ---`r`n$stderr"
    }
    $content += "`r`n--- EXITCODE: $($p.ExitCode) ---`r`n"

    $result = $content -join ''

    if ($OutputPath) {
      $result | Out-File -LiteralPath $OutputPath -Encoding UTF8
    }

    if ($PassThru) {
      return $result
    }
  }
  catch {
    $errorMessage = "ERROR running $FilePath $($ArgumentList -join ' '): $($_.Exception.Message)"
    if ($OutputPath) {
      $errorMessage | Out-File -LiteralPath $OutputPath -Encoding UTF8
    }
    if ($PassThru) {
      return $errorMessage
    }
    Write-Error $errorMessage
  }
}

function Export-EventLog {
  <#
    .SYNOPSIS
      Exports a named Windows event log to an .evtx file via wevtutil.
    .DESCRIPTION
      Wraps wevtutil.exe epl. If the log name does not exist, a warning
      is recorded in -MissingLogPath (when supplied) and no error is thrown.
      Designed for bulk log collection during IR triage.
    .PARAMETER LogName
      Full event log name, e.g. 'Security', 'Microsoft-Windows-PowerShell/Operational'.
    .PARAMETER OutputPath
      Path for the exported .evtx file.
    .PARAMETER MissingLogPath
      When supplied and the log is not found, the missing log name is appended
      to this text file so collectors can report what was unavailable.
    .EXAMPLE
      PS> Export-EventLog -LogName 'Security' -OutputPath '.\EVTX\Security.evtx'
    .LINK
      https://github.com/adnoctem/winkit/lib/security.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $LogName,

    [Parameter(Mandatory = $true)]
    [string]
    $OutputPath,

    [Parameter(Mandatory = $false)]
    [string]
    $MissingLogPath
  )

  try {
    $result = Invoke-SafeProcess -FilePath 'wevtutil.exe' -ArgumentList @('el') -PassThru
    $exists = $result -split "`r`n" | Where-Object { $_ -eq $LogName }

    if (-not $exists) {
      if ($MissingLogPath) {
        "Log not present: $LogName" | Out-File -LiteralPath $MissingLogPath -Append -Encoding UTF8
      }
      Write-Verbose "Event log not found: $LogName"
      return $false
    }

    $null = Invoke-SafeProcess -FilePath 'wevtutil.exe' -ArgumentList @('epl', $LogName, $OutputPath)
    return $true
  }
  catch {
    Write-Error "Failed to export event log '$LogName': $_"
    return $false
  }
}

function Get-ScheduledTaskAction {
  <#
    .SYNOPSIS
      Returns structured scheduled task action data for all registered tasks.
    .DESCRIPTION
      Enumerates every scheduled task via Get-ScheduledTask and expands each
      task's Actions collection into a flat list of [PSCustomObject] records
      with TaskName, TaskPath, State, Execute, and Arguments properties.
      Use -SuspiciousOnly to filter to known execution-host paths (powershell,
      cmd, wscript, cscript, mshta, rundll32, regsvr32, InstallUtil).
    .PARAMETER SuspiciousOnly
      Only return actions with an Execute path matching common scripting and
      LOLBin hosts.
    .EXAMPLE
      PS> Get-ScheduledTaskAction | Export-Csv .\ScheduledTasks-Actions.csv -NoTypeInformation
    .EXAMPLE
      PS> Get-ScheduledTaskAction -SuspiciousOnly
    .LINK
      https://github.com/adnoctem/winkit/lib/security.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject[]])]
  param (
    [Parameter(Mandatory = $false)]
    [switch]
    $SuspiciousOnly
  )

  $suspiciousHosts = 'powershell|pwsh|cmd|wscript|cscript|mshta|rundll32|regsvr32|InstallUtil'

  $rows = Get-ScheduledTask -ErrorAction SilentlyContinue |
    ForEach-Object {
      $task = $_
      foreach ($action in $task.Actions) {
        $obj = [PSCustomObject]@{
          TaskName = $task.TaskName
          TaskPath = $task.TaskPath
          State = $task.State
          Execute = $action.Execute
          Arguments = $action.Arguments
        }
        if ($SuspiciousOnly) {
          if ($obj.Execute -match $suspiciousHosts) { $obj }
        }
        else {
          $obj
        }
      }
    }

  $rows | Sort-Object TaskPath, TaskName
}

function Get-WMIPersistence {
  <#
    .SYNOPSIS
      Enumerates WMI subscription-based persistence.
    .DESCRIPTION
      Queries the root\subscription namespace for __EventFilter,
      CommandLineEventConsumer, and __FilterToConsumerBinding instances.
      Returns an object with three properties: EventFilters, CommandLineConsumers,
      and Bindings, each an array of the corresponding WMI objects.
      Returns $null when no subscriptions exist.
    .EXAMPLE
      PS> $wmi = Get-WMIPersistence
      PS> $wmi.EventFilters | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/security.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param()

  [PSCustomObject]@{
    EventFilters = @(Get-CimInstance -Namespace root\subscription -ClassName __EventFilter -ErrorAction SilentlyContinue)
    CommandLineConsumers = @(Get-CimInstance -Namespace root\subscription -ClassName CommandLineEventConsumer -ErrorAction SilentlyContinue)
    Bindings = @(Get-CimInstance -Namespace root\subscription -ClassName __FilterToConsumerBinding -ErrorAction SilentlyContinue)
  }
}

# ==============================================================================
# Windows Security Event API
# ==============================================================================

function Import-SecurityEventConfiguration {
  <#
    .SYNOPSIS
      Loads and caches the security event configuration from security.psd1.
    .DESCRIPTION
      Imports the security event definition data file and caches it in script
      scope. Subsequent calls return the cached configuration unless -Force is
      supplied, avoiding repeated file reads.
    .PARAMETER Path
      Path to the security.psd1 configuration file. Defaults to the file
      alongside this script.
    .PARAMETER Force
      Reload the configuration even if it is already cached.
    .EXAMPLE
      PS> $config = Import-SecurityEventConfiguration
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [string] $Path = (Join-Path -Path $PSScriptRoot -ChildPath 'security.psd1'),

    [Parameter(Mandatory = $false)]
    [switch] $Force
  )

  if ($script:SecurityEventConfiguration -and -not $Force) {
    return $script:SecurityEventConfiguration
  }

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "Security event configuration file not found: $Path"
  }

  $script:SecurityEventConfiguration = Import-PowerShellDataFile -Path $Path
  return $script:SecurityEventConfiguration
}

function Get-SecurityEventGroup {
  <#
    .SYNOPSIS
      Retrieves a single semantic event group from the configuration.
    .DESCRIPTION
      Looks up a named group (e.g. Logon, Service, PowerShell) from the
      cached configuration. Throws if the group is unknown.
    .PARAMETER Name
      Case-sensitive group name to retrieve.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to the cached configuration from
      Import-SecurityEventConfiguration.
    .EXAMPLE
      PS> $group = Get-SecurityEventGroup -Name 'Logon'
      PS> $group.EventIds
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string] $Name,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  if (-not $Configuration.Groups.ContainsKey($Name)) {
    throw "Unknown security event group: $Name"
  }

  return $Configuration.Groups[$Name]
}

function Get-SecurityEventDefinition {
  <#
    .SYNOPSIS
      Retrieves event definition objects from the configuration.
    .DESCRIPTION
      Flattens the nested Events structure into a list of event definition
      hashtables, optionally filtered by Group, LogName, ProviderName, Id,
      or Name.
    .PARAMETER Group
      Limit results to events belonging to this semantic group.
    .PARAMETER LogName
      Limit results to events from this event log channel.
    .PARAMETER ProviderName
      Limit results to events from this provider.
    .PARAMETER Id
      Limit results to events with these event IDs.
    .PARAMETER Name
      Limit results to events with this logical name.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-SecurityEventDefinition -Id 4624, 4625
    .EXAMPLE
      PS> Get-SecurityEventDefinition -Group 'Logon' -LogName 'Security'
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [string] $Group,

    [Parameter(Mandatory = $false)]
    [string] $LogName,

    [Parameter(Mandatory = $false)]
    [string] $ProviderName,

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string] $Name,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  $results = New-Object System.Collections.ArrayList

  foreach ($logKey in $Configuration.Events.Keys) {
    foreach ($subKey in $Configuration.Events[$logKey].Keys) {
      foreach ($def in $Configuration.Events[$logKey][$subKey]) {
        if ($Group -and $def.UtilityGroup -ne $Group) { continue }
        if ($LogName -and $def.LogName -ne $LogName) { continue }
        if ($ProviderName -and $def.ProviderName -ne $ProviderName) { continue }
        if ($Id -and $def.Id -notin $Id) { continue }
        if ($Name -and $def.Name -ne $Name) { continue }
        [void]$results.Add($def)
      }
    }
  }

  return $results
}

function Test-WindowsEventLogChannel {
  <#
    .SYNOPSIS
      Tests whether an event log channel exists on the local machine.
    .DESCRIPTION
      Uses Get-WinEvent -ListLog to verify channel existence. Returns $true
      if the channel is available, $false otherwise. No errors are thrown
      for missing channels.
    .PARAMETER LogName
      Full event log channel name, e.g. 'Security' or
      'Microsoft-Windows-Sysmon/Operational'.
    .EXAMPLE
      PS> if (Test-WindowsEventLogChannel -LogName 'Microsoft-Windows-Sysmon/Operational') { ... }
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string] $LogName
  )

  try {
    $null = Get-WinEvent -ListLog $LogName -ErrorAction Stop
    return $true
  }
  catch {
    return $false
  }
}

function Get-WindowsEventByDefinition {
  <#
    .SYNOPSIS
      Generic event query helper wrapping Get-WinEvent -FilterHashtable.
    .DESCRIPTION
      Queries Windows event logs using filter-hashtable-based queries for
      performance. Accepts event definitions, a group name, or ID lists.
      Definitions are grouped by LogName to minimise individual queries.
      Supports remote computers and optional channel skipping.
    .PARAMETER Definition
      Array of event definition hashtables.
    .PARAMETER Group
      Semantic group name resolved from the configuration.
    .PARAMETER Id
      Event IDs to query (bypasses definition lookup).
    .PARAMETER LogName
      Event log channel(s) to query.
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER SkipMissingChannel
      Skip channels that do not exist rather than throwing.
    .PARAMETER MaxEvents
      Maximum events to return per log/channel query.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsEventByDefinition -Group 'Logon' -StartTime (Get-Date).AddHours(-4)
    .EXAMPLE
      PS> Get-WindowsEventByDefinition -Id 4624, 4625 -LogName 'Security' -MaxEvents 100
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true)]
    [hashtable[]] $Definition,

    [Parameter(Mandatory = $false)]
    [string] $Group,

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string[]] $LogName,

    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [switch] $SkipMissingChannel,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  begin {
    $allDefs = New-Object System.Collections.ArrayList
  }

  process {
    if ($Definition) {
      [void]$allDefs.AddRange($Definition)
    }
  }

  end {
    if ($Group) {
      $groupDefinition = Get-SecurityEventGroup -Name $Group -Configuration $Configuration
      $resolvedDefs = @(Get-SecurityEventDefinition -Group $Group -Configuration $Configuration)

      if ($resolvedDefs.Count -eq 0 -and $groupDefinition.EventIds) {
        foreach ($log in $groupDefinition.DefaultLogs) {
          foreach ($eventId in $groupDefinition.EventIds) {
            [void]$allDefs.Add(@{ LogName = $log; Id = $eventId })
          }
        }
      }
      else {
        [void]$allDefs.AddRange($resolvedDefs)
      }
    }

    if ($Id) {
      if ($LogName) {
        foreach ($log in $LogName) {
          foreach ($eventId in $Id) {
            [void]$allDefs.Add(@{ LogName = $log; Id = $eventId })
          }
        }
      }
      else {
        foreach ($eventId in $Id) {
          [void]$allDefs.Add(@{ Id = $eventId })
        }
      }
    }

    if ($allDefs.Count -eq 0) {
      throw 'No event definitions supplied or resolved. Provide -Definition, -Group, or -Id.'
    }

    $allDefs |
      Group-Object LogName |
      ForEach-Object {
        $log = $_.Name
        $ids = $_.Group.Id | Sort-Object -Unique

        if ($SkipMissingChannel -and -not (Test-WindowsEventLogChannel -LogName $log)) {
          return
        }

        $filter = @{
          LogName = $log
          Id = $ids
          StartTime = $StartTime
          EndTime = $EndTime
        }

        if ($MaxEvents) {
          $filter.MaxEvents = $MaxEvents
        }

        try {
          if ($ComputerName) {
            foreach ($computer in $ComputerName) {
              Get-WinEvent -ComputerName $computer -FilterHashtable $filter -ErrorAction SilentlyContinue
            }
          }
          else {
            Get-WinEvent -FilterHashtable $filter -ErrorAction SilentlyContinue
          }
        }
        catch {
          if (-not $SkipMissingChannel) {
            Write-Error "Failed to query log '$log': $_"
          }
        }
      }
  }
}

function Resolve-WindowsEventMappedField {
  <#
    .SYNOPSIS
      Resolves a mapped field value to its semantic name.
    .DESCRIPTION
      Looks up a value in a named field map (e.g. LogonType, ImpersonationLevel)
      and returns the map entry. Returns $null when the map or value is not
      found.
    .PARAMETER MapName
      Name of the field map (e.g. 'LogonType').
    .PARAMETER Value
      Raw value to resolve. Handles both integer and string keys.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Resolve-WindowsEventMappedField -MapName 'LogonType' -Value 10
      Returns a hashtable with Name = 'RemoteInteractive'.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string] $MapName,

    [Parameter(Mandatory = $true)]
    [object] $Value,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  if (-not $Configuration.FieldMaps.ContainsKey($MapName)) {
    return $null
  }

  $map = $Configuration.FieldMaps[$MapName]

  if ($map.ContainsKey($Value)) {
    return $map[$Value]
  }

  $stringValue = [string] $Value
  if ($map.ContainsKey($stringValue)) {
    return $map[$stringValue]
  }

  return $null
}

function ConvertFrom-WinEvent {
  <#
    .SYNOPSIS
      Converts a raw EventRecord into a structured enriched object.
    .DESCRIPTION
      Parses the XML representation of an event record, extracts named
      EventData fields into a RawData hashtable, and produces a PSCustomObject
      with common top-level properties plus enriched mapped fields such as
      LogonTypeName. The original EventRecord is preserved on the object.
    .PARAMETER Event
      A [System.Diagnostics.Eventing.Reader.EventRecord] to convert.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsEventByDefinition -Id 4624 -MaxEvents 1 | ConvertFrom-WinEvent
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object] $Event,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  process {
    $xml = [xml] $Event.ToXml()

    $rawData = [ordered] @{}
    if ($xml.Event.EventData -and $xml.Event.EventData.Data) {
      foreach ($node in $xml.Event.EventData.Data) {
        if (-not [string]::IsNullOrWhiteSpace($node.Name)) {
          $rawData[$node.Name] = $node.'#text'
        }
      }
    }

    $logonType = $null
    $logonTypeName = $null
    if ($rawData.Contains('LogonType')) {
      try {
        $logonType = [int] $rawData['LogonType']
        $resolved = Resolve-WindowsEventMappedField -MapName 'LogonType' -Value $logonType -Configuration $Configuration
        if ($resolved) {
          $logonTypeName = $resolved.Name
        }
      }
      catch {
        $logonType = $rawData['LogonType']
      }
    }

    $impersonationLevelName = $null
    if ($rawData.Contains('ImpersonationLevel')) {
      $resolved = Resolve-WindowsEventMappedField -MapName 'ImpersonationLevel' -Value $rawData['ImpersonationLevel'] -Configuration $Configuration
      if ($resolved) {
        $impersonationLevelName = $resolved.Name
      }
    }

    [pscustomobject] @{
      TimeCreated = $Event.TimeCreated
      Id = $Event.Id
      ProviderName = $Event.ProviderName
      LogName = $Event.LogName
      MachineName = $Event.MachineName
      RecordId = $Event.RecordId
      LevelDisplayName = $Event.LevelDisplayName
      TargetUserName = $rawData['TargetUserName']
      TargetDomainName = $rawData['TargetDomainName']
      SubjectUserName = $rawData['SubjectUserName']
      SubjectDomainName = $rawData['SubjectDomainName']
      LogonType = $logonType
      LogonTypeName = $logonTypeName
      ImpersonationLevel = $rawData['ImpersonationLevel']
      ImpersonationLevelName = $impersonationLevelName
      IpAddress = $rawData['IpAddress']
      IpPort = $rawData['IpPort']
      WorkstationName = $rawData['WorkstationName']
      ProcessName = $rawData['ProcessName']
      ProcessId = $rawData['ProcessId']
      LogonProcessName = $rawData['LogonProcessName']
      AuthenticationPackageName = $rawData['AuthenticationPackageName']
      Status = $rawData['Status']
      SubStatus = $rawData['SubStatus']
      TargetLogonId = $rawData['TargetLogonId']
      ServiceName = $rawData['ServiceName']
      ImagePath = $rawData['ImagePath']
      ServiceFileName = $rawData['ServiceFileName']
      RawData = $rawData
      EventRecord = $Event
    }
  }
}

function Get-WindowsLogonEvent {
  <#
    .SYNOPSIS
      Queries and normalises logon-related security events.
    .DESCRIPTION
      Wraps Get-WindowsEventByDefinition for the Logon group. Supports
      filtering by event ID and logon type. Suppresses noisy system accounts
      by default unless -IncludeSystem is supplied.
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER Id
      Specific event IDs to return. Defaults to all Logon group IDs.
    .PARAMETER LogonType
      Filter by numeric logon type(s), e.g. 10 for RDP.
    .PARAMETER IncludeSystem
      Include system and machine accounts normally suppressed.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER MaxEvents
      Maximum events to return.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsLogonEvent -Id 4625
      Returns failed logon events from the past 24 hours.
    .EXAMPLE
      PS> Get-WindowsLogonEvent -Id 4624 -LogonType 10
      Returns successful RDP logons.
    .EXAMPLE
      PS> Get-WindowsLogonEvent -Id 4624, 4800, 4801 -LogonType 2, 7
      Returns local console and unlock activity.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [int[]] $LogonType,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeSystem,

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  $group = Get-SecurityEventGroup -Name 'Logon' -Configuration $Configuration

  if (-not $Id) {
    $Id = $group.EventIds
  }

  $queryParams = @{
    Group = 'Logon'
    StartTime = $StartTime
    EndTime = $EndTime
    Configuration = $Configuration
  }
  if ($ComputerName) { $queryParams.ComputerName = $ComputerName }
  if ($MaxEvents) { $queryParams.MaxEvents = $MaxEvents }

  $events = Get-WindowsEventByDefinition @queryParams

  $events |
    ConvertFrom-WinEvent -Configuration $Configuration |
    Where-Object {
      if ($_.Id -notin $Id) {
        return $false
      }

      if ($LogonType -and $_.LogonType -notin $LogonType) {
        return $false
      }

      if (-not $IncludeSystem) {
        if ($group.DefaultSuppressions -and
          $group.DefaultSuppressions.TargetUserName -and
          $_.TargetUserName -in $group.DefaultSuppressions.TargetUserName) {
          return $false
        }

        if ($group.DefaultSuppressions -and
          $group.DefaultSuppressions.TargetUserNameEndsWith) {
          foreach ($suffix in $group.DefaultSuppressions.TargetUserNameEndsWith) {
            if ($_.TargetUserName -like "*$suffix") {
              return $false
            }
          }
        }
      }

      return $true
    }
}

function Get-WindowsAccountChangeEvent {
  <#
    .SYNOPSIS
      Queries account lifecycle and group membership change events.
    .DESCRIPTION
      Wraps Get-WindowsEventByDefinition for the AccountChange group.
      Supports filtering by target user, subject user, and event ID.
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER Id
      Specific event IDs to return. Defaults to all AccountChange group IDs.
    .PARAMETER TargetUserName
      Filter by the target account name of the change.
    .PARAMETER SubjectUserName
      Filter by the account that performed the change.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER MaxEvents
      Maximum events to return.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsAccountChangeEvent -Id 4720
      Returns user account creation events.
    .EXAMPLE
      PS> Get-WindowsAccountChangeEvent -Id 4728, 4732 -StartTime (Get-Date).AddDays(-7)
      Returns group membership additions for the past 7 days.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string] $TargetUserName,

    [Parameter(Mandatory = $false)]
    [string] $SubjectUserName,

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  $group = Get-SecurityEventGroup -Name 'AccountChange' -Configuration $Configuration

  if (-not $Id) {
    $Id = $group.EventIds
  }

  $queryParams = @{
    Group = 'AccountChange'
    StartTime = $StartTime
    EndTime = $EndTime
    Configuration = $Configuration
  }
  if ($ComputerName) { $queryParams.ComputerName = $ComputerName }
  if ($MaxEvents) { $queryParams.MaxEvents = $MaxEvents }

  $events = Get-WindowsEventByDefinition @queryParams

  $events |
    ConvertFrom-WinEvent -Configuration $Configuration |
    Where-Object {
      if ($_.Id -notin $Id) { return $false }
      if ($TargetUserName -and $_.TargetUserName -ne $TargetUserName) { return $false }
      if ($SubjectUserName -and $_.SubjectUserName -ne $SubjectUserName) { return $false }
      return $true
    }
}

function Get-WindowsServiceEvent {
  <#
    .SYNOPSIS
      Queries service lifecycle, failure, and installation events.
    .DESCRIPTION
      Wraps Get-WindowsEventByDefinition for the Service group. Supports
      filtering by service name and event ID.
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER Id
      Specific event IDs to return. Defaults to all Service group IDs.
    .PARAMETER ServiceName
      Filter by the service name involved.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER MaxEvents
      Maximum events to return.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsServiceEvent -Id 7045
      Returns new service installation events (common persistence vector).
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string] $ServiceName,

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  $group = Get-SecurityEventGroup -Name 'Service' -Configuration $Configuration

  if (-not $Id) {
    $Id = $group.EventIds
  }

  $queryParams = @{
    Group = 'Service'
    StartTime = $StartTime
    EndTime = $EndTime
    Configuration = $Configuration
  }
  if ($ComputerName) { $queryParams.ComputerName = $ComputerName }
  if ($MaxEvents) { $queryParams.MaxEvents = $MaxEvents }

  $events = Get-WindowsEventByDefinition @queryParams

  $events |
    ConvertFrom-WinEvent -Configuration $Configuration |
    Where-Object {
      if ($_.Id -notin $Id) { return $false }
      if ($ServiceName -and $_.ServiceName -ne $ServiceName) { return $false }
      return $true
    }
}

function Get-WindowsBootEvent {
  <#
    .SYNOPSIS
      Queries boot, shutdown, and crash events.
    .DESCRIPTION
      Wraps Get-WindowsEventByDefinition for the BootShutdown group.
      Covers unexpected reboots, BSODs, clean shutdowns, and service
      lifecycle transitions.
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER Id
      Specific event IDs to return. Defaults to all BootShutdown group IDs.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER MaxEvents
      Maximum events to return.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsBootEvent -Id 41, 1001
      Returns unexpected shutdowns and BSOD events.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  $group = Get-SecurityEventGroup -Name 'BootShutdown' -Configuration $Configuration

  if (-not $Id) {
    $Id = $group.EventIds
  }

  $queryParams = @{
    Group = 'BootShutdown'
    StartTime = $StartTime
    EndTime = $EndTime
    Configuration = $Configuration
  }
  if ($ComputerName) { $queryParams.ComputerName = $ComputerName }
  if ($MaxEvents) { $queryParams.MaxEvents = $MaxEvents }

  $events = Get-WindowsEventByDefinition @queryParams

  $events |
    ConvertFrom-WinEvent -Configuration $Configuration |
    Where-Object {
      if ($Id -and $_.Id -notin $Id) { return $false }
      return $true
    }
}

function Get-WindowsPowerShellEvent {
  <#
    .SYNOPSIS
      Queries PowerShell operational telemetry events.
    .DESCRIPTION
      Wraps Get-WindowsEventByDefinition for the PowerShell group. Supports
      filtering by event ID, executing user, and script block content pattern.
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER Id
      Specific event IDs to return. Defaults to all PowerShell group IDs.
    .PARAMETER UserName
      Filter by the user account that executed the PowerShell code.
    .PARAMETER ScriptBlockText
      Regex pattern to search within captured script block content.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER MaxEvents
      Maximum events to return.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsPowerShellEvent -Id 4104
      Returns script block logging events (highest-value PowerShell event).
    .EXAMPLE
      PS> Get-WindowsPowerShellEvent -ScriptBlockText 'DownloadString|FromBase64'
      Returns PowerShell events matching suspicious download or encoding patterns.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string] $UserName,

    [Parameter(Mandatory = $false)]
    [string] $ScriptBlockText,

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  $group = Get-SecurityEventGroup -Name 'PowerShell' -Configuration $Configuration

  if (-not $Id) {
    $Id = $group.EventIds
  }

  $queryParams = @{
    Group = 'PowerShell'
    StartTime = $StartTime
    EndTime = $EndTime
    Configuration = $Configuration
  }
  if ($ComputerName) { $queryParams.ComputerName = $ComputerName }
  if ($MaxEvents) { $queryParams.MaxEvents = $MaxEvents }

  $events = Get-WindowsEventByDefinition @queryParams

  $events |
    ConvertFrom-WinEvent -Configuration $Configuration |
    Where-Object {
      if ($_.Id -notin $Id) { return $false }

      if ($UserName) {
        $raw = $_.RawData
        $found = $false
        foreach ($key in @('UserName', 'UserId', 'User', 'SubjectUserName')) {
          if ($raw.ContainsKey($key) -and $raw[$key] -match $UserName) {
            $found = $true
            break
          }
        }
        if (-not $found) { return $false }
      }

      if ($ScriptBlockText) {
        $raw = $_.RawData
        $matched = $false
        foreach ($key in @('ScriptBlockText', 'Message', 'Payload')) {
          if ($raw.ContainsKey($key) -and $raw[$key] -match $ScriptBlockText) {
            $matched = $true
            break
          }
        }
        if (-not $matched) { return $false }
      }

      return $true
    }
}

function Get-WindowsScheduledTaskEvent {
  <#
    .SYNOPSIS
      Queries scheduled task lifecycle telemetry.
    .DESCRIPTION
      Wraps Get-WindowsEventByDefinition for the ScheduledTask group.
      Supports filtering by event ID and task name.
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER Id
      Specific event IDs to return. Defaults to all ScheduledTask group IDs.
    .PARAMETER TaskName
      Filter by the name of the scheduled task.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER MaxEvents
      Maximum events to return.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsScheduledTaskEvent -Id 106, 140, 141
      Returns task registration, update, and deletion events.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string] $TaskName,

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  $group = Get-SecurityEventGroup -Name 'ScheduledTask' -Configuration $Configuration

  if (-not $Id) {
    $Id = $group.EventIds
  }

  $queryParams = @{
    Group = 'ScheduledTask'
    StartTime = $StartTime
    EndTime = $EndTime
    Configuration = $Configuration
  }
  if ($ComputerName) { $queryParams.ComputerName = $ComputerName }
  if ($MaxEvents) { $queryParams.MaxEvents = $MaxEvents }

  $events = Get-WindowsEventByDefinition @queryParams

  $events |
    ConvertFrom-WinEvent -Configuration $Configuration |
    Where-Object {
      if ($_.Id -notin $Id) { return $false }

      if ($TaskName) {
        $raw = $_.RawData
        $found = $false
        foreach ($key in @('TaskName', 'Name')) {
          if ($raw.ContainsKey($key) -and $raw[$key] -match $TaskName) {
            $found = $true
            break
          }
        }
        if (-not $found) { return $false }
      }

      return $true
    }
}

function Get-WindowsSysmonEvent {
  <#
    .SYNOPSIS
      Queries Sysmon telemetry if the channel is available.
    .DESCRIPTION
      Wraps Get-WindowsEventByDefinition for the Sysmon group with
      -SkipMissingChannel enabled by default (Sysmon is optional).
    .PARAMETER StartTime
      Earliest event timestamp. Defaults to 24 hours ago.
    .PARAMETER EndTime
      Latest event timestamp. Defaults to now.
    .PARAMETER Id
      Specific event IDs to return. Defaults to all Sysmon group IDs.
    .PARAMETER ComputerName
      Target remote computer(s).
    .PARAMETER MaxEvents
      Maximum events to return.
    .PARAMETER Configuration
      Configuration hashtable. Defaults to cached.
    .EXAMPLE
      PS> Get-WindowsSysmonEvent -Id 1
      Returns Sysmon process creation events.
    .EXAMPLE
      PS> Get-WindowsSysmonEvent -Id 3, 22
      Returns Sysmon network connection and DNS query events.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [datetime] $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [datetime] $EndTime = (Get-Date),

    [Parameter(Mandatory = $false)]
    [int[]] $Id,

    [Parameter(Mandatory = $false)]
    [string[]] $ComputerName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int] $MaxEvents,

    [Parameter(Mandatory = $false)]
    [hashtable] $Configuration = (Import-SecurityEventConfiguration)
  )

  if (-not (Test-WindowsEventLogChannel -LogName 'Microsoft-Windows-Sysmon/Operational')) {
    Write-Verbose 'Sysmon operational channel not found. Sysmon may not be installed or configured.'
    return
  }

  $group = Get-SecurityEventGroup -Name 'Sysmon' -Configuration $Configuration

  if (-not $Id) {
    $Id = $group.EventIds
  }

  $queryParams = @{
    Group = 'Sysmon'
    StartTime = $StartTime
    EndTime = $EndTime
    SkipMissingChannel = $true
    Configuration = $Configuration
  }
  if ($ComputerName) { $queryParams.ComputerName = $ComputerName }
  if ($MaxEvents) { $queryParams.MaxEvents = $MaxEvents }

  $events = Get-WindowsEventByDefinition @queryParams

  $events |
    ConvertFrom-WinEvent -Configuration $Configuration |
    Where-Object {
      if ($_.Id -notin $Id) { return $false }
      return $true
    }
}
