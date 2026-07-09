# Map of short and long hive names to canonical PowerShell provider paths.
# HKLM: and HKCU: are default PowerShell drives; the remaining hives use the
# Registry:: provider path because their short PS drives are not guaranteed to
# exist in Windows PowerShell 5.1.
$script:HiveMap = @{
  'HKLM' = 'HKLM:'
  'HKCU' = 'HKCU:'
  'HKCR' = 'Registry::HKEY_CLASSES_ROOT'
  'HKU' = 'Registry::HKEY_USERS'
  'HKCC' = 'Registry::HKEY_CURRENT_CONFIG'
  'HKEY_LOCAL_MACHINE' = 'HKLM:'
  'HKEY_CURRENT_USER' = 'HKCU:'
  'HKEY_CLASSES_ROOT' = 'Registry::HKEY_CLASSES_ROOT'
  'HKEY_USERS' = 'Registry::HKEY_USERS'
  'HKEY_CURRENT_CONFIG' = 'Registry::HKEY_CURRENT_CONFIG'
}

# Normalises any registry path format into a canonical PS drive path (HKLM:\...)
function ConvertTo-RegistryProviderPath {
  <#
    .SYNOPSIS
      Converts a registry path to PS drive format suitable for provider cmdlets.
    .DESCRIPTION
      Accepts short hive notation (HKLM\...), PS drive notation (HKLM:\...),
      Registry:: prefix (Registry::HKEY_LOCAL_MACHINE\...), or long .NET names
      (HKEY_LOCAL_MACHINE\...). Returns a provider path that can be passed to
      registry provider cmdlets. HKLM and HKCU use their default PowerShell
      drives; hives such as HKEY_USERS use the Registry:: provider path because
      aliases like HKU: are not present in every shell.
    .EXAMPLE
      PS> ConvertTo-RegistryProviderPath 'HKLM\Software\MyApp'
      HKLM:\Software\MyApp
    .EXAMPLE
      PS> ConvertTo-RegistryProviderPath 'Registry::HKEY_CURRENT_USER\Control Panel'
      HKCU:\Control Panel
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param (
    # Registry path in any supported format
    [Parameter(Mandatory = $true)]
    [string]
    $Path
  )

  # Strip Registry:: provider prefix if present
  $_normalized = $Path -replace '^Registry::', ''

  # Detect the hive portion and replace with canonical PS drive root
  $_matched = $false
  foreach ($_candidate in $script:HiveMap.Keys) {
    # Match exact hive name at start, optionally followed by : or \
    $_candidatePattern = [regex]::Escape($_candidate)
    if ($_normalized -match "^$($_candidatePattern)(?::|\\|$)") {
      $_remainder = $_normalized.Substring($_candidate.Length).TrimStart(':', '\')
      $_normalized = if ([string]::IsNullOrEmpty($_remainder)) {
        $script:HiveMap[$_candidate]
      }
      else {
        "$($script:HiveMap[$_candidate])\$_remainder"
      }
      $_matched = $true
      break
    }
  }

  if (-not $_matched) {
    Write-Error "Unable to resolve registry hive from path: '$Path'"
    return $null
  }

  # Collapse duplicate separators and trim trailing backslash/colon
  $_normalized = $_normalized -replace '\\+', '\' -replace ':+$', '' -replace '\\+$', ''

  return $_normalized
}

# Resolves a registry path to a RegistryKey object
function Resolve-RegistryPath {
  <#
    .SYNOPSIS
      Resolves a registry path string to a Microsoft.Win32.RegistryKey object.
    .DESCRIPTION
      Parses a registry path supplied in PS drive format (HKLM:\... or
      HKCU:\...), short format (HKLM\...), or full .NET provider format
      (Registry::HKEY_LOCAL_MACHINE\...), looks up the corresponding hive from
      the module's HiveMap, and returns the matching RegistryKey. Returns $null
      when the requested key does not exist.
    .EXAMPLE
      PS> Resolve-RegistryPath -Path 'HKLM:\Software\Microsoft'
    .EXAMPLE
      PS> Resolve-RegistryPath -Path 'HKCU:\Control Panel\Desktop' -Writable
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([Microsoft.Win32.RegistryKey])]
  [CmdletBinding()]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp' or 'HKCU:\Control Panel'
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    # Open the key with write access
    [Parameter(Mandatory = $false)]
    [switch]
    $Writable = $false
  )

  # Short hive name -> .NET RegistryHive (kept here for this function only;
  # higher-level functions rely on the provider cmdlets instead.)
  $_hiveEnum = @{
    'HKLM' = [Microsoft.Win32.RegistryHive]::LocalMachine
    'HKCU' = [Microsoft.Win32.RegistryHive]::CurrentUser
    'HKCR' = [Microsoft.Win32.RegistryHive]::ClassesRoot
    'HKU' = [Microsoft.Win32.RegistryHive]::Users
    'HKCC' = [Microsoft.Win32.RegistryHive]::CurrentConfig
    'HKEY_LOCAL_MACHINE' = [Microsoft.Win32.RegistryHive]::LocalMachine
    'HKEY_CURRENT_USER' = [Microsoft.Win32.RegistryHive]::CurrentUser
    'HKEY_CLASSES_ROOT' = [Microsoft.Win32.RegistryHive]::ClassesRoot
    'HKEY_USERS' = [Microsoft.Win32.RegistryHive]::Users
    'HKEY_CURRENT_CONFIG' = [Microsoft.Win32.RegistryHive]::CurrentConfig
  }

  # Normalise path then extract the hive name
  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  $_registryPath = $_providerPath -replace '^Registry::', ''
  $_separator = $_registryPath.IndexOf('\')
  if ($_separator -eq -1) {
    $_hiveName = $_registryPath.TrimEnd(':')
    $_subKey = ''
  }
  else {
    $_hiveName = $_registryPath.Substring(0, $_separator).TrimEnd(':')
    $_subKey = $_registryPath.Substring($_separator + 1)
  }

  if (-not $_hiveEnum.ContainsKey($_hiveName)) {
    Write-Error "Unknown registry hive: '$_hiveName'. Supported: $($_hiveEnum.Keys -join ', ')"
    return $null
  }

  $_hive = $_hiveEnum[$_hiveName]

  try {
    if ([string]::IsNullOrEmpty($_subKey)) {
      return [Microsoft.Win32.RegistryKey]::OpenBaseKey($_hive, [Microsoft.Win32.RegistryView]::Default)
    }
    $_root = [Microsoft.Win32.RegistryKey]::OpenBaseKey($_hive, [Microsoft.Win32.RegistryView]::Default)
    return $_root.OpenSubKey($_subKey, $Writable)
  }
  catch [System.UnauthorizedAccessException] {
    Write-Error "Access denied opening registry key: '$Path'"
    return $null
  }
  catch {
    Write-Error "Failed to resolve registry path '$Path': $_"
    return $null
  }
}

function Get-RegistryKey {
  <#
    .SYNOPSIS
      Retrieves metadata about a registry key.
    .DESCRIPTION
      Opens the registry key at the given path and returns an object describing
      its subkey names and value names. Returns $null when the key does not
      exist.
    .EXAMPLE
      PS> Get-RegistryKey -Path 'HKLM:\Software\Microsoft'
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  if (-not (Test-Path -Path $_providerPath)) {
    Write-Error "Registry key not found: '$Path'"
    return $null
  }

  try {
    # Enumerate subkeys via the PS drive provider
    $_subKeys = @(Get-ChildItem -Path $_providerPath -ErrorAction Stop |
        ForEach-Object { $_.PSChildName })

    # Enumerate value names; '(default)' is normalised to '' to match .NET behaviour
    $_itemProps = Get-ItemProperty -Path $_providerPath -ErrorAction Stop
    $_values = @($_itemProps.PSObject.Properties |
        Where-Object { $_.Name -notin @('PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider') } |
        ForEach-Object { if ($_.Name -eq '(default)') { '' } else { $_.Name } })

    [PSCustomObject]@{
      Path = $Path
      SubKeys = $_subKeys
      Values = $_values
    }
  }
  catch [System.UnauthorizedAccessException] {
    Write-Error "Access denied reading registry key: '$Path'"
    return $null
  }
  catch {
    Write-Error "Failed to read registry key '$Path': $_"
    return $null
  }
}

function Set-RegistryKey {
  <#
    .SYNOPSIS
      Creates a registry key if it does not already exist.
    .DESCRIPTION
      Attempts to create the registry key at the given path. If the key already
      exists the operation is skipped. Returns a status object indicating
      whether the key was created or already present.
    .EXAMPLE
      PS> Set-RegistryKey -Path 'HKLM:\Software\MyApp'
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  # Guard against root-hive creation
  if ($_providerPath -notmatch '\\') {
    Write-Error "Cannot create a root hive key: '$Path'"
    return $null
  }

  # Idempotency: check existence via the provider (cheap, no .NET handle needed)
  if (Test-Path -Path $_providerPath) {
    Write-Verbose "Registry key already exists: '$Path'"
    return [PSCustomObject]@{
      Path = $Path
      Status = 'AlreadyExists'
    }
  }

  if ($PSCmdlet.ShouldProcess($Path, 'Create registry key')) {
    try {
      # New-Item -Force creates the key tree; returns $null when key already exists
      $null = New-Item -Path $_providerPath -Force -ErrorAction Stop
      Write-Verbose "Created registry key: '$Path'"
      return [PSCustomObject]@{
        Path = $Path
        Status = 'Created'
      }
    }
    catch [System.UnauthorizedAccessException] {
      Write-Error "Access denied creating registry key: '$Path'"
      return $null
    }
    catch {
      Write-Error "Failed to create registry key '$Path': $_"
      return $null
    }
  }
}

function Remove-RegistryKey {
  <#
    .SYNOPSIS
      Removes a registry key and optionally its subkeys.
    .DESCRIPTION
      Deletes the registry key at the given path. When -Recurse is specified,
      all descendant subkeys are removed as well. If the key does not exist the
      operation is skipped. Returns a status object indicating whether the key
      was removed or was already absent.
    .EXAMPLE
      PS> Remove-RegistryKey -Path 'HKLM:\Software\MyApp'
    .EXAMPLE
      PS> Remove-RegistryKey -Path 'HKLM:\Software\MyApp' -Recurse
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    # Remove the key and all its descendant subkeys
    [Parameter(Mandatory = $false)]
    [switch]
    $Recurse = $false
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  # Guard against root-hive removal
  if ($_providerPath -notmatch '\\') {
    Write-Error "Cannot remove a root hive key: '$Path'"
    return $null
  }

  # Idempotency: skip when the key is already absent
  if (-not (Test-Path -Path $_providerPath)) {
    Write-Verbose "Registry key does not exist (nothing to remove): '$Path'"
    return [PSCustomObject]@{
      Path = $Path
      Status = 'NotFound'
    }
  }

  if ($PSCmdlet.ShouldProcess($Path, 'Remove registry key')) {
    try {
      # Remove-Item with -Recurse handles the subkey-tree case natively
      $null = Remove-Item -Path $_providerPath -Recurse:$Recurse -Force -ErrorAction Stop
      Write-Verbose "Removed registry key: '$Path'"
      return [PSCustomObject]@{
        Path = $Path
        Status = 'Removed'
      }
    }
    catch [System.UnauthorizedAccessException] {
      Write-Error "Access denied removing registry key: '$Path'"
      return $null
    }
    catch {
      Write-Error "Failed to remove registry key '$Path': $_"
      return $null
    }
  }
}

function Get-RegistryValue {
  <#
    .SYNOPSIS
      Reads a named value from a registry key.
    .DESCRIPTION
      Opens the registry key at the given path and returns the data stored in
      the named value. Returns $null when either the key or the value does not
      exist.
    .EXAMPLE
      PS> Get-RegistryValue -Path 'HKLM:\Software\MyApp' -Name 'Version'
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([object])]
  [CmdletBinding()]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    # Name of the value to read; omit to read the default (unnamed) value
    [Parameter(Mandatory = $false)]
    [string]
    $Name = ''
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  if (-not (Test-Path -Path $_providerPath)) {
    Write-Error "Registry key not found: '$Path'"
    return $null
  }

  try {
    $_resolvedName = if ($Name -eq '') { '(default)' } else { $Name }
    return Get-ItemPropertyValue -Path $_providerPath -Name $_resolvedName -ErrorAction Stop
  }
  catch [System.UnauthorizedAccessException] {
    Write-Error "Access denied reading registry key: '$Path'"
    return $null
  }
  catch {
    Write-Verbose "Value '$Name' not found in registry key '$Path'"
    return $null
  }
}

function Set-RegistryValue {
  <#
    .SYNOPSIS
      Creates or updates a named value in a registry key.
    .DESCRIPTION
      Writes the supplied data to the named value under the given registry key.
      The function is idempotent: if the value already exists and holds the same
      data the operation is skipped. Returns a status object indicating whether
      the value was created, updated, or left unchanged.
    .EXAMPLE
      PS> Set-RegistryValue -Path 'HKLM:\Software\MyApp' -Name 'Version' -Value '2.0'
    .EXAMPLE
      PS> Set-RegistryValue -Path 'HKLM:\Software\MyApp' -Name 'Enabled' -Value 1 -Type DWord
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    # Name of the value to set; omit to set the default (unnamed) value
    [Parameter(Mandatory = $false)]
    [string]
    $Name = '',

    # Data to store in the value
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [object]
    $Value,

    # Registry value kind; defaults to String for strings, DWord for integers,
    # and ExpandString / MultiString / Binary / QWord as appropriate
    [Parameter(Mandatory = $false)]
    [ValidateSet('String', 'ExpandString', 'Binary', 'DWord', 'MultiString', 'QWord', 'Unknown')]
    [Microsoft.Win32.RegistryValueKind]
    $Type
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  # --- Ensure the parent key exists (idempotent) ---
  if (-not (Test-Path -Path $_providerPath)) {
    Write-Verbose "Parent key does not exist, creating: '$Path'"
    $_created = Set-RegistryKey -Path $Path
    if (-not $_created -or $_created.Status -notin @('Created', 'AlreadyExists')) {
      Write-Error "Failed to ensure registry key exists: '$Path'"
      return $null
    }
  }

  try {
    # --- Determine the registry value kind if not explicitly provided ---
    if (-not $PSBoundParameters.ContainsKey('Type')) {
      if ($null -eq $Value) {
        $Type = [Microsoft.Win32.RegistryValueKind]::String
      }
      elseif ($Value -is [int] -or $Value -is [long]) {
        $Type = [Microsoft.Win32.RegistryValueKind]::DWord
      }
      elseif ($Value -is [string[]]) {
        $Type = [Microsoft.Win32.RegistryValueKind]::MultiString
      }
      elseif ($Value -is [byte[]]) {
        $Type = [Microsoft.Win32.RegistryValueKind]::Binary
      }
      else {
        $Type = [Microsoft.Win32.RegistryValueKind]::String
      }
    }

    # --- Idempotency: read current value and compare ---
    # The provider surfaces the default (unnamed) value as '(default)'.
    $_resolvedName = if ($Name -eq '') { '(default)' } else { $Name }

    try {
      $_currentValue = Get-ItemPropertyValue -Path $_providerPath -Name $_resolvedName -ErrorAction Stop
      $_valueExists = $true
    }
    catch {
      $_valueExists = $false
      $_currentValue = $null
    }

    if ($_valueExists) {

      # Infer the RegistryValueKind from the .NET type Get-ItemProperty returned.
      # This is reliable for all kinds except String vs ExpandString (both
      # surface as [string]).  In that rare edge-case we fall back to a
      # lightweight .NET call for an exact match.
      $_inferKind = {
        param($_v)
        if ($null -eq $_v) { return [Microsoft.Win32.RegistryValueKind]::String }
        if ($_v -is [int]) { return [Microsoft.Win32.RegistryValueKind]::DWord }
        if ($_v -is [long]) { return [Microsoft.Win32.RegistryValueKind]::QWord }
        if ($_v -is [string[]]) { return [Microsoft.Win32.RegistryValueKind]::MultiString }
        if ($_v -is [byte[]]) { return [Microsoft.Win32.RegistryValueKind]::Binary }
        return [Microsoft.Win32.RegistryValueKind]::String
      }
      $_currentInferred = & $_inferKind $_currentValue

      # Compare values with proper handling for null and array types
      $_isSameValue = $false
      if ($null -eq $_currentValue -and $null -eq $Value) {
        $_isSameValue = $true
      }
      elseif ($null -ne $_currentValue -and $null -ne $Value) {
        if ($_currentValue -is [string[]] -and $Value -is [string[]]) {
          $_isSameValue = (($_currentValue -join "`0") -eq ($Value -join "`0"))
        }
        elseif ($_currentValue -is [byte[]] -and $Value -is [byte[]]) {
          $_isSameValue = ([System.Convert]::ToBase64String($_currentValue) -eq [System.Convert]::ToBase64String($Value))
        }
        else {
          $_isSameValue = ($_currentValue.ToString() -eq $Value.ToString())
        }
      }

      # Type check: if the inferred kind disagrees with the target type
      # and either is ExpandString, use a precise .NET read to be certain.
      $_isSameKind = ($_currentInferred -eq $Type)
      if (-not $_isSameKind -and
        ($Type -eq [Microsoft.Win32.RegistryValueKind]::ExpandString -or
        $_currentInferred -eq [Microsoft.Win32.RegistryValueKind]::ExpandString -or
        $_currentInferred -eq [Microsoft.Win32.RegistryValueKind]::String -and
        $Type -eq [Microsoft.Win32.RegistryValueKind]::String)) {
        # Ambiguous case - fall back to a precise .NET read for the actual kind
        $_key = Resolve-RegistryPath -Path $Path
        if ($_key) {
          try { $_isSameKind = ($_key.GetValueKind($Name) -eq $Type) }
          catch { $_isSameKind = $false }
          finally { $_key.Dispose() }
        }
      }

      if ($_isSameValue -and $_isSameKind) {
        Write-Verbose "Registry value '$Name' already set to the requested data in '$Path'"
        return [PSCustomObject]@{
          Path = $Path
          Name = $Name
          Status = 'Unchanged'
        }
      }

      # --- Update existing value ---
      if ($PSCmdlet.ShouldProcess("$Path\$Name", "Update registry value to '$Value'")) {
        $null = Set-ItemProperty -Path $_providerPath -Name $_resolvedName -Value $Value -Type $Type -ErrorAction Stop
        Write-Verbose "Updated registry value '$Name' in '$Path'"
        return [PSCustomObject]@{
          Path = $Path
          Name = $Name
          Status = 'Updated'
        }
      }
    }
    else {
      # --- Create new value ---
      if ($PSCmdlet.ShouldProcess("$Path\$Name", "Create registry value with '$Value'")) {
        $null = Set-ItemProperty -Path $_providerPath -Name $_resolvedName -Value $Value -Type $Type -ErrorAction Stop
        Write-Verbose "Created registry value '$Name' in '$Path'"
        return [PSCustomObject]@{
          Path = $Path
          Name = $Name
          Status = 'Created'
        }
      }
    }
  }
  catch [System.UnauthorizedAccessException] {
    Write-Error "Access denied setting registry value '$Name' in '$Path'"
    return $null
  }
  catch {
    Write-Error "Failed to set registry value '$Name' in '$Path': $_"
    return $null
  }
}

function Remove-RegistryValue {
  <#
    .SYNOPSIS
      Removes a named value from a registry key.
    .DESCRIPTION
      Deletes the named value from the registry key at the given path. If the
      value does not exist the operation is skipped. Returns a status object
      indicating whether the value was removed or was already absent.
    .EXAMPLE
      PS> Remove-RegistryValue -Path 'HKLM:\Software\MyApp' -Name 'Version'
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    # Name of the value to remove; omit to remove the default (unnamed) value
    [Parameter(Mandatory = $false)]
    [string]
    $Name = ''
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  # Key existence check (read-only, no handle needed)
  if (-not (Test-Path -Path $_providerPath)) {
    Write-Verbose "Registry key not found (nothing to remove): '$Path'"
    return [PSCustomObject]@{
      Path = $Path
      Name = $Name
      Status = 'KeyNotFound'
    }
  }

  try {
    $_resolvedName = if ($Name -eq '') { '(default)' } else { $Name }

    try {
      $null = Get-ItemPropertyValue -Path $_providerPath -Name $_resolvedName -ErrorAction Stop
    }
    catch {
      Write-Verbose "Registry value '$Name' does not exist in '$Path' (nothing to remove)"
      return [PSCustomObject]@{
        Path = $Path
        Name = $Name
        Status = 'NotFound'
      }
    }

    if ($PSCmdlet.ShouldProcess("$Path\$Name", 'Remove registry value')) {
      $null = Remove-ItemProperty -Path $_providerPath -Name $_resolvedName -ErrorAction Stop
      Write-Verbose "Removed registry value '$Name' from '$Path'"
      return [PSCustomObject]@{
        Path = $Path
        Name = $Name
        Status = 'Removed'
      }
    }
  }
  catch [System.UnauthorizedAccessException] {
    Write-Error "Access denied removing registry value '$Name' from '$Path'"
    return $null
  }
  catch {
    Write-Error "Failed to remove registry value '$Name' from '$Path': $_"
    return $null
  }
}

function Test-RegistryPath {
  <#
    .SYNOPSIS
      Returns $true if the registry path exists, $false otherwise.
    .DESCRIPTION
      A safe, non-terminating existence check for registry keys.  Unlike
      Resolve-RegistryPath, this function never emits errors - it simply
      returns a boolean.  Suitable for use in conditionals and idempotency
      guards.
    .EXAMPLE
      PS> if (Test-RegistryPath 'HKLM:\Software\MyApp') { ... }
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([bool])]
  [CmdletBinding()]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $false }

  return Test-Path -Path $_providerPath
}

function Test-RegistryValue {
  <#
    .SYNOPSIS
      Returns $true if the named registry value exists, $false otherwise.
    .DESCRIPTION
      Checks both that the parent key exists and that the named value is
      present on that key.  Never emits errors, making it ideal for
      conditional guards before reading or removing values.
    .EXAMPLE
      PS> if (Test-RegistryValue 'HKLM:\Software\MyApp' -Name 'Version') { ... }
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([bool])]
  [CmdletBinding()]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    # Name of the value to test; omit to test the default (unnamed) value
    [Parameter(Mandatory = $false)]
    [string]
    $Name = ''
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $false }

  if (-not (Test-Path -Path $_providerPath)) { return $false }

  try {
    $_resolvedName = if ($Name -eq '') { '(default)' } else { $Name }
    $null = Get-ItemPropertyValue -Path $_providerPath -Name $_resolvedName -ErrorAction Stop
    return $true
  }
  catch {
    return $false
  }
}

function Get-RegistryValueKind {
  <#
    .SYNOPSIS
      Gets the exact Microsoft.Win32.RegistryValueKind of a registry value.
    .DESCRIPTION
      Because the PS provider cmdlets (Get-ItemProperty / Get-ItemPropertyValue)
      do not expose the raw RegistryValueKind, this function uses a lightweight
      .NET read solely for type inspection.  It is useful when you need to
      distinguish, e.g., String from ExpandString.
      Returns $null if the key or value does not exist.
    .EXAMPLE
      PS> Get-RegistryValueKind -Path 'HKLM:\Software\MyApp' -Name 'PathVar'
      ExpandString
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([Microsoft.Win32.RegistryValueKind])]
  [CmdletBinding()]
  param (
    # Registry path, e.g. 'HKLM:\Software\MyApp'
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    # Name of the value to inspect; omit to inspect the default (unnamed) value
    [Parameter(Mandatory = $false)]
    [string]
    $Name = ''
  )

  $_providerPath = ConvertTo-RegistryProviderPath -Path $Path
  if (-not $_providerPath) { return $null }

  if (-not (Test-Path -Path $_providerPath)) {
    Write-Verbose "Registry key not found: '$Path'"
    return $null
  }

  # Verify the value is present before opening a .NET handle
  if (-not (Test-RegistryValue -Path $Path -Name $Name)) {
    Write-Verbose "Value '$Name' not found in registry key '$Path'"
    return $null
  }

  $_key = Resolve-RegistryPath -Path $Path
  if (-not $_key) { return $null }

  try {
    return $_key.GetValueKind($Name)
  }
  catch {
    Write-Error "Failed to get registry value kind for '$Name' in '$Path': $_"
    return $null
  }
  finally {
    if ($_key) { $_key.Dispose() }
  }
}

function Mount-DefaultUserHive {
  <#
    .SYNOPSIS
      Loads C:\Users\Default\NTUSER.DAT into the registry under HKU\<MountName>.
    .DESCRIPTION
      Mounts the default user profile hive so it can be modified before
      sealing an image (Sysprep / Audit Mode).  Changes written here propagate
      to every new user account created on the deployed system.
      HKEY_USERS\.DEFAULT is the LocalSystem profile - do NOT write there for
      this purpose.  Use Dismount-DefaultUserHive to unload when done.
      Requires elevation.
    .PARAMETER MountName
      Subkey name under HKEY_USERS to mount at. Defaults to 'DefaultUser'.
    .PARAMETER HivePath
      Path to NTUSER.DAT. Defaults to C:\Users\Default\NTUSER.DAT.
    .EXAMPLE
      Mount-DefaultUserHive
      Set-RegistryValue -Path 'Registry::HKEY_USERS\DefaultUser\Software\...' -Name 'Foo' -Value 1
      Dismount-DefaultUserHive
    .NOTES
      reg.exe is used because hive loading is not exposed through the managed
      registry API without manual privilege elevation (SE_RESTORE_NAME /
      SE_BACKUP_NAME).
  #>
  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
  [OutputType([string])]
  param(
    [ValidatePattern('^[A-Za-z0-9_-]+$')]
    [string]$MountName = 'DefaultUser',

    [ValidateNotNullOrEmpty()]
    [string]$HivePath = (Join-Path $env:SystemDrive 'Users\Default\NTUSER.DAT')
  )

  if (-not (Test-Path -LiteralPath $HivePath)) {
    Write-Log -Message "Default user hive not found at '$HivePath'." -Color Red
    return $null
  }

  # Refuse if already mounted - silent re-use risks writing to the wrong hive
  $mountPath = "Registry::HKEY_USERS\$MountName"
  if (Test-Path -LiteralPath $mountPath) {
    Write-Log -Message "A hive is already mounted at HKEY_USERS\$MountName. Dismount it first or choose a different MountName." -Color Red
    return $null
  }

  if (-not $PSCmdlet.ShouldProcess("HKEY_USERS\$MountName", "Load hive from '$HivePath'")) {
    return $null
  }

  $output = & reg.exe load "HKU\$MountName" $HivePath 2>&1
  if ($LASTEXITCODE -ne 0) {
    Write-Log -Message "reg.exe load failed (exit $LASTEXITCODE): $output" -Color Red
    return $null
  }

  Write-Log -Message "Mounted default user hive: HKEY_USERS\$MountName" -Color Green
  return $mountPath
}

function Dismount-DefaultUserHive {
  <#
    .SYNOPSIS
      Unloads a previously mounted default user hive.
    .DESCRIPTION
      Calls reg.exe unload.  If lingering handles cause the first attempt to
      fail, forces a garbage collection and retries once.  A failed dismount
      leaves the default profile corrupted - investigate before sealing the
      image.
    .PARAMETER MountName
      Subkey name under HKEY_USERS to unload. Must match Mount-DefaultUserHive.
  #>
  [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
  param(
    [ValidatePattern('^[A-Za-z0-9_-]+$')]
    [string]$MountName = 'DefaultUser'
  )

  $mountPath = "Registry::HKEY_USERS\$MountName"
  if (-not (Test-Path -LiteralPath $mountPath)) {
    Write-Log -Message "No hive mounted at HKEY_USERS\$MountName; nothing to unload." -Color Yellow
    return
  }

  if (-not $PSCmdlet.ShouldProcess("HKEY_USERS\$MountName", 'Unload hive')) {
    return
  }

  # First attempt
  $output = & reg.exe unload "HKU\$MountName" 2>&1
  if ($LASTEXITCODE -eq 0) {
    Write-Log -Message "Unloaded HKEY_USERS\$MountName." -Color Green
    return
  }

  # Retry after forcing GC - PowerShell's registry provider may still hold handles
  Write-Log -Message "First unload attempt failed: $output. Forcing GC and retrying." -Color Yellow
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 500

  $output = & reg.exe unload "HKU\$MountName" 2>&1
  if ($LASTEXITCODE -ne 0) {
    Write-Log -Message "reg.exe unload failed after GC retry (exit $LASTEXITCODE): $output. Investigate which process still has the hive open before continuing." -Color Red
    return
  }
  Write-Log -Message "Unloaded HKEY_USERS\$MountName on retry." -Color Green
}

function Export-RegistryKey {
  <#
    .SYNOPSIS
      Exports a registry key to a .reg text file via reg.exe.
    .DESCRIPTION
      Calls reg.exe export /y against the supplied key. Uses Invoke-SafeProcess
      internally so stdout/stderr are captured.
    .PARAMETER Key
      Registry key path, e.g. 'HKLM\Software\Microsoft\Windows\CurrentVersion\Run'.
    .PARAMETER OutputPath
      Path for the exported .reg file.
    .EXAMPLE
      PS> Export-RegistryKey -Key 'HKLM\Software\Microsoft\Windows\CurrentVersion\Run' -OutputPath '.\HKLM-Run.reg'
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Key,

    [Parameter(Mandatory = $true)]
    [string]
    $OutputPath
  )

  try {
    $null = Invoke-SafeProcess -FilePath 'reg.exe' -ArgumentList @('export', $Key, $OutputPath, '/y')
    return $true
  }
  catch {
    Write-Error "Failed to export registry key '$Key': $_"
    return $false
  }
}

function Search-RegistryKey {
  <#
    .SYNOPSIS
      Searches a registry hive for a pattern via reg.exe query.
    .DESCRIPTION
      Calls reg.exe query <Root> /f <Pattern> /s to recursively search for a
      string or pattern across a registry hive. Output is written to -OutputPath.
      Uses Invoke-SafeProcess internally.
    .PARAMETER Root
      Registry hive root, e.g. 'HKLM', 'HKCU'.
    .PARAMETER Pattern
      Search pattern forwarded to /f.
    .PARAMETER OutputPath
      File to write the search results to.
    .EXAMPLE
      PS> Search-RegistryKey -Root 'HKLM' -Pattern 'InstallUtil' -OutputPath '.\Reg-HKLM-InstallUtil.txt'
    .LINK
      https://github.com/adnoctem/winkit/lib/registry.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Root,

    [Parameter(Mandatory = $true)]
    [string]
    $Pattern,

    [Parameter(Mandatory = $true)]
    [string]
    $OutputPath
  )

  try {
    $null = Invoke-SafeProcess -FilePath 'reg.exe' -ArgumentList @('query', $Root, '/f', $Pattern, '/s') -OutputPath $OutputPath
    return $true
  }
  catch {
    Write-Error "Failed to search registry '$Root' for pattern '$Pattern': $_"
    return $false
  }
}
