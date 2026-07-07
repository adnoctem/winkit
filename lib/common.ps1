#Requires -Version 5.0

function New-OperationResult {
  <#
    .SYNOPSIS
      Creates a structured result object for script and library operations.
    .DESCRIPTION
      Builds a small, consistent PSCustomObject that scripts can return when
      they need to report planned, skipped, completed, or failed work. The
      helper keeps the common fields in a predictable order while omitting
      optional fields that were not supplied, which lets each script keep a
      compact output shape.

      Target, Action, and Status are always present. Source, Scope, Detail,
      SkippedReason, and Error are included only when the matching parameter is
      supplied. Additional fields can be appended through -Property for
      script-specific metadata without creating another custom result helper.
    .PARAMETER Target
      Object targeted by the operation, such as a package name, registry value,
      file path, scheduled task, or capability name.
    .PARAMETER Source
      Optional subsystem or provider that handled the operation, such as
      Registry, FileSystem, WinGet, Win32Program, or UPFAppxPackage.
    .PARAMETER Scope
      Optional scope affected by the operation, such as CurrentUser, Machine,
      DefaultUser, or AllExistingUsers.
    .PARAMETER Action
      Operation performed or planned, such as Install, Uninstall, SetValue,
      RemoveShortcut, or Disable.
    .PARAMETER Status
      Result state, such as Completed, Removed, Skipped, Failed, DryRun, or an
      ExitCode:n value.
    .PARAMETER Detail
      Optional human-readable detail that explains the result.
    .PARAMETER SkippedReason
      Optional machine-readable reason when Status is Skipped.
    .PARAMETER ErrorMessage
      Optional failure detail. This is emitted as the Error property to avoid
      assigning to PowerShell's automatic $Error variable.
    .PARAMETER Property
      Optional hashtable of additional properties appended after the common
      fields. Existing common fields are not overwritten.
    .EXAMPLE
      PS> New-OperationResult -Target 'Microsoft.GetHelp' -Source 'UPFAppxPackage' -Action 'Uninstall' -Status 'Skipped' -SkippedReason 'NoMatch'

      Creates a package-style lifecycle result with Target, Source, Action,
      Status, and SkippedReason.
    .EXAMPLE
      PS> New-OperationResult -Target 'HKCU:\Console\%%Startup\DelegationTerminal' -Scope 'CurrentUser' -Action 'SetValue' -Status 'Completed' -Detail 'Terminal delegation default applied.'

      Creates a terminal configuration result with Target, Scope, Action,
      Status, and Detail.
    .EXAMPLE
      PS> New-OperationResult -Target 'OneDrive' -Action 'Uninstall' -Status 'Skipped' -Detail 'No OneDrive uninstaller was found.'

      Creates a compact script result that does not include Source or Scope
      columns because they were not supplied.
    .LINK
      https://github.com/adnoctem/winkit/lib/common.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Factory helper only creates a result object.')]
  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [string]$Target,
    [string]$Source,
    [string]$Scope,
    [string]$Action,
    [string]$Status,
    [string]$Detail,
    [string]$SkippedReason,
    [string]$ErrorMessage,
    [hashtable]$Property
  )

  $_result = [ordered]@{
    Target = $Target
  }

  if ($PSBoundParameters.ContainsKey('Source')) {
    $_result.Source = $Source
  }

  if ($PSBoundParameters.ContainsKey('Scope')) {
    $_result.Scope = $Scope
  }

  $_result.Action = $Action
  $_result.Status = $Status

  if ($PSBoundParameters.ContainsKey('Detail')) {
    $_result.Detail = $Detail
  }

  if ($PSBoundParameters.ContainsKey('SkippedReason')) {
    $_result.SkippedReason = $SkippedReason
  }

  if ($PSBoundParameters.ContainsKey('ErrorMessage')) {
    $_result.Error = $ErrorMessage
  }

  if ($Property) {
    foreach ($_key in $Property.Keys) {
      if (-not $_result.Contains($_key)) {
        $_result[$_key] = $Property[$_key]
      }
    }
  }

  [PSCustomObject]$_result
}

function Add-OperationResult {
  <#
    .SYNOPSIS
      Adds a structured operation result to an existing result collection.
    .DESCRIPTION
      Creates an operation result with New-OperationResult and appends it to a
      mutable result collection, usually a System.Collections.ArrayList used by
      high-impact scripts that support -PassThru. By default the helper writes
      no pipeline output, matching the common pattern of accumulating results
      and returning them once at the end of a script.
    .PARAMETER Results
      Mutable result collection that receives the created result object.
    .PARAMETER Target
      Object targeted by the operation, such as a package name, registry value,
      file path, scheduled task, or capability name.
    .PARAMETER Source
      Optional subsystem or provider that handled the operation.
    .PARAMETER Scope
      Optional scope affected by the operation.
    .PARAMETER Action
      Operation performed or planned.
    .PARAMETER Status
      Result state.
    .PARAMETER Detail
      Optional human-readable detail that explains the result.
    .PARAMETER SkippedReason
      Optional machine-readable reason when Status is Skipped.
    .PARAMETER ErrorMessage
      Optional failure detail. This is emitted as the Error property.
    .PARAMETER Property
      Optional hashtable of additional properties appended after the common
      fields.
    .PARAMETER PassThru
      Return the created result object after adding it to the collection.
    .EXAMPLE
      PS> $results = New-Object System.Collections.ArrayList
      PS> Add-OperationResult -Results $results -Target 'MapsToastTask' -Source 'ScheduledTask' -Action 'Disable' -Status 'Disabled' -Detail 'Scheduled task disabled.'

      Adds a scheduled-task result to the collection without writing output.
    .EXAMPLE
      PS> Add-OperationResult -Results $results -Target 'Microsoft.WindowsTerminal' -Scope 'Machine' -Action 'Install' -Status 'Skipped' -Detail 'AlreadyInstalled' -PassThru

      Adds and returns the result object.
    .LINK
      https://github.com/adnoctem/winkit/lib/common.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.IList]$Results,

    [string]$Target,
    [string]$Source,
    [string]$Scope,
    [string]$Action,
    [string]$Status,
    [string]$Detail,
    [string]$SkippedReason,
    [string]$ErrorMessage,
    [hashtable]$Property,
    [switch]$PassThru
  )

  $_resultParameters = @{
    Target = $Target
    Action = $Action
    Status = $Status
  }

  foreach ($_parameterName in @('Source', 'Scope', 'Detail', 'SkippedReason', 'ErrorMessage', 'Property')) {
    if ($PSBoundParameters.ContainsKey($_parameterName)) {
      $_resultParameters[$_parameterName] = $PSBoundParameters[$_parameterName]
    }
  }

  $_result = New-OperationResult @_resultParameters
  [void]$Results.Add($_result)

  if ($PassThru) {
    $_result
  }
}

function Write-OperationResultLog {
  <#
    .SYNOPSIS
      Writes operation result objects to a JSON Lines log file.
    .DESCRIPTION
      Serializes each operation result as one compact JSON object per line in
      a temp-directory log file. The helper is intended for scripts that keep
      console output concise but still need an auditable record of every
      action, skipped action, and failure.

      Logs are written to %TEMP%\winkit\logs by default. Each line includes a
      timestamp, script name, and the properties already present on the result
      object, such as Target, Source, Scope, Action, Status, Detail,
      SkippedReason, or Error.
    .PARAMETER Results
      Operation result objects to serialize.
    .PARAMETER ScriptName
      Name used in the log entries and default file name. Defaults to the
      calling script name when available.
    .PARAMETER Path
      Optional explicit output file path. When omitted, a timestamped .jsonl
      file is created under %TEMP%\winkit\logs.
    .EXAMPLE
      PS> $path = Write-OperationResultLog -Results $results -ScriptName 'Remove-Bloatware'
      PS> Write-Log -Message "Operation log: $path" -Color Gray

      Writes the accumulated operation results and prints the resulting path.
    .LINK
      https://github.com/adnoctem/winkit/lib/common.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [System.Collections.IEnumerable]$Results,

    [string]$ScriptName,
    [string]$Path
  )

  $_results = @($Results)
  if ($_results.Count -eq 0) {
    return $null
  }

  if ([string]::IsNullOrWhiteSpace($ScriptName)) {
    if ($MyInvocation.ScriptName) {
      $ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.ScriptName)
    }
    else {
      $ScriptName = 'winkit-operation'
    }
  }

  $_safeScriptName = $ScriptName -replace '[^A-Za-z0-9._-]', '-'
  if ([string]::IsNullOrWhiteSpace($Path)) {
    $_tempRoot = if ([string]::IsNullOrWhiteSpace($env:TEMP)) { [System.IO.Path]::GetTempPath() } else { $env:TEMP }
    $_logRoot = Join-Path -Path $_tempRoot -ChildPath 'winkit\logs'
    if (-not (Test-Path -LiteralPath $_logRoot)) {
      $null = [System.IO.Directory]::CreateDirectory($_logRoot)
    }
    $_timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $Path = Join-Path -Path $_logRoot -ChildPath "$_safeScriptName-$_timestamp.jsonl"
  }
  else {
    $_logRoot = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($_logRoot) -and -not (Test-Path -LiteralPath $_logRoot)) {
      $null = [System.IO.Directory]::CreateDirectory($_logRoot)
    }
  }

  $_lines = foreach ($_result in $_results) {
    $_entry = [ordered]@{
      Timestamp = (Get-Date).ToString('o')
      Script = $ScriptName
    }

    foreach ($_property in $_result.PSObject.Properties) {
      $_entry[$_property.Name] = $_property.Value
    }

    [PSCustomObject]$_entry | ConvertTo-Json -Compress -Depth 8
  }

  $_encoding = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllLines($Path, [string[]]$_lines, $_encoding)
  return $Path
}

function Export-RegistrySettingState {
  <#
    .SYNOPSIS
      Exports registry setting definitions with Preferred set to current values.
    .DESCRIPTION
      Reads the current machine value for each registry setting object and
      returns a reusable configuration snapshot. The returned objects preserve
      the original setting metadata, including Path, Name, Default, Type, Group,
      Description, and build gates, but replace Preferred with the value found
      in the registry.

      Missing registry values are exported with Preferred = $null. This keeps
      the output compatible with existing -Config behavior while making absent
      values explicit for later diffs.
    .PARAMETER Settings
      Registry setting objects or hashtables with at least Path and Name
      properties. Additional properties are preserved.
    .EXAMPLE
      PS> Export-RegistrySettingState -Settings $taskbarSettings | ConvertTo-Json -Depth 3

      Exports current taskbar registry values in the same schema consumed by
      Configure-Taskbar.ps1 -Config.
    .LINK
      https://github.com/adnoctem/winkit/lib/common.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Factory helper only creates result objects for registry setting definitions.')]
  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object[]]$Settings
  )

  process {
    foreach ($_setting in $Settings) {
      if ($null -eq $_setting) { continue }

      $_snapshot = [ordered]@{}
      if ($_setting -is [hashtable]) {
        foreach ($_key in $_setting.Keys) {
          $_snapshot[$_key] = $_setting[$_key]
        }
      }
      else {
        foreach ($_property in $_setting.PSObject.Properties) {
          $_snapshot[$_property.Name] = $_property.Value
        }
      }

      if (-not $_snapshot.Contains('Path') -or -not $_snapshot.Contains('Name')) {
        Write-Error 'Registry setting snapshots require Path and Name properties.'
        continue
      }

      $_currentValue = $null
      if (Test-RegistryValue -Path $_snapshot.Path -Name $_snapshot.Name) {
        $_currentValue = Get-RegistryValue -Path $_snapshot.Path -Name $_snapshot.Name
        $_currentKind = Get-RegistryValueKind -Path $_snapshot.Path -Name $_snapshot.Name
        if ($_currentKind -and $_snapshot.Contains('Type')) {
          $_snapshot.Type = $_currentKind.ToString()
        }
      }

      $_snapshot.Preferred = $_currentValue
      [PSCustomObject]$_snapshot
    }
  }
}

function ConvertTo-RegistrySettingResult {
  <#
    .SYNOPSIS
      Creates operation results for registry setting definitions.
    .DESCRIPTION
      Converts registry setting objects into operation-result records suitable
      for Write-OperationResultLog and -PassThru output. In DryRun mode the
      results describe planned SetValue or RemoveValue actions as Skipped with
      Detail = DryRun. In normal mode the helper reads the registry after the
      script has run and marks each setting Completed when the target state is
      present, Removed when an undo removal target is absent, or Failed when the
      current state does not match the expected state.

      This helper is intentionally conservative: it does not attempt to infer
      whether a Completed value was newly changed or already correct. Scripts
      that need exact lifecycle states can still add explicit results during
      their apply loop.
    .PARAMETER Settings
      Registry setting objects or hashtables with Path, Name, Preferred,
      Default, Type, and Description properties.
    .PARAMETER Undo
      Build results for an undo operation, using Default as the target state.
    .PARAMETER DryRun
      Build planned-operation results without reading final state.
    .PARAMETER Source
      Source label for result objects. Defaults to Registry.
    .EXAMPLE
      PS> $results = ConvertTo-RegistrySettingResult -Settings $taskbarSettings -DryRun
      PS> Write-OperationResultLog -Results $results -ScriptName 'Configure-Taskbar'

      Creates dry-run audit records for a registry-backed configuration script.
    .LINK
      https://github.com/adnoctem/winkit/lib/common.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object[]]$Settings,

    [switch]$Undo,
    [switch]$DryRun,
    [string]$Source = 'Registry'
  )

  process {
    foreach ($_setting in $Settings) {
      if ($null -eq $_setting) { continue }

      $_path = $_setting.Path
      $_name = $_setting.Name
      $_target = "$_path\$_name"
      $_targetValue = if ($Undo) { $_setting.Default } else { $_setting.Preferred }
      $_action = if ($Undo -and $null -eq $_setting.Default) { 'RemoveValue' } else { 'SetValue' }
      $_detail = if ($_setting.PSObject.Properties.Name -contains 'Description') { $_setting.Description } else { $null }

      if ($DryRun) {
        New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Skipped' -Detail 'DryRun'
        continue
      }

      if ($_action -eq 'RemoveValue') {
        if ($_path -like 'Registry::HKEY_USERS\DefaultUser\*' -and -not (Test-Path -LiteralPath 'Registry::HKEY_USERS\DefaultUser')) {
          New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Skipped' -Detail 'DefaultUserHiveUnavailable'
          continue
        }

        if (Test-RegistryValue -Path $_path -Name $_name) {
          New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Failed' -Detail 'Value still exists after undo.'
        }
        else {
          New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Removed' -Detail $_detail
        }
        continue
      }

      if ($_path -like 'Registry::HKEY_USERS\DefaultUser\*' -and -not (Test-Path -LiteralPath 'Registry::HKEY_USERS\DefaultUser')) {
        New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Skipped' -Detail 'DefaultUserHiveUnavailable'
        continue
      }

      if (-not (Test-RegistryValue -Path $_path -Name $_name)) {
        New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Failed' -Detail 'Value is missing after apply.'
        continue
      }

      $_currentValue = Get-RegistryValue -Path $_path -Name $_name
      $_matches = $false
      if ($null -eq $_currentValue -and $null -eq $_targetValue) {
        $_matches = $true
      }
      elseif ($null -ne $_currentValue -and $null -ne $_targetValue) {
        if ($_currentValue -is [array] -or $_targetValue -is [array]) {
          $_matches = (($_currentValue -join "`0") -eq ($_targetValue -join "`0"))
        }
        else {
          $_matches = ($_currentValue.ToString() -eq $_targetValue.ToString())
        }
      }

      if ($_matches) {
        New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Completed' -Detail $_detail
      }
      else {
        New-OperationResult -Target $_target -Source $Source -Action $_action -Status 'Failed' -Detail "Expected '$($_targetValue)' but found '$($_currentValue)'."
      }
    }
  }
}
