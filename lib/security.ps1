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
