function Convert-Quote {
  <#
  .SYNOPSIS
    Converts single quotes to double quotes or vice versa in a specified file.
  .DESCRIPTION
    This function reads the content of a file and replaces all single quotes with double quotes or all double quotes with single quotes, based on the specified parameter. It is useful for standardizing quote usage in configuration files, scripts, or any text files.
  .PARAMETER Path
    The full path to the file that needs to be processed. The file must exist and be accessible for reading and writing.
  .PARAMETER To
    Specifies the type of quote conversion to perform. Acceptable values are "Single" for converting double quotes to single quotes and "Double" for converting single quotes to double quotes. The default value is "Double".
  .EXAMPLE
    PS> Convert-Quote -Path 'C:\config.txt' -To 'Single'
    This command converts all double quotes in the file 'C:\config.txt' to single quotes.
  .EXAMPLE
    PS> Convert-Quote -Path 'C:\config.txt' -To 'Double'
    This command converts all single quotes in the file 'C:\config.txt' to double quotes.
  .LINK
    https://github.com/adnoctem/winkit/blob/main/lib/data.ps1
  .NOTES
    Author: Maximilian Gindorfer <info@mvprowess.com>
    License: MIT
    #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Function merges two object arrays; existing public name is intentionally plural.')]
  [OutputType([void])]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Single", "Double")]
    [string]$To = "Double"
  )

  # NOTE: '-Raw' is required to read the entire file as a single string, allowing for proper replacement of quotes
  $content = Get-Content -Path $Path -Raw

  switch ($To) {
    "Double" { $content = $content -replace "'", '"' }
    "Single" { $content = $content -replace '"', "'" }

    default {
      throw "Invalid value for -To parameter. Use 'Single' or 'Double'."
    }
  }

  Set-Content -Path $Path -Value $content
}

function Merge-ObjectArrays {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Merge-ObjectArrays describes a two-array merge helper; existing public name is intentionally plural.')]

  <#
  .SYNOPSIS
    Merges two arrays of objects, overriding base values with those from overrides.
  .DESCRIPTION
    For each object in the Overrides array, finds the matching object in the
    Base array by Name (or by Path + Name when a Path property is present).
    Only keys that already exist on the base object are transferred - unknown
    keys in the override are silently ignored.  The base array is mutated in
    place and no output is returned.
  .PARAMETER Base
    The array of hashtables or objects to merge into (modified in place).
  .PARAMETER Overrides
    The array of hashtables or objects whose values take precedence.
  .EXAMPLE
    PS> $base = @(@{ Name = 'A'; Value = 1 }, @{ Name = 'B'; Value = 2 })
    PS> $over  = @(@{ Name = 'A'; Value = 99; Extra = 'ignored' })
    PS> Merge-ObjectArrays -Base $base -Overrides $over
    # $base[0].Value is now 99; 'Extra' is silently ignored
  .LINK
    https://github.com/adnoctem/winkit/blob/main/lib/data.ps1
  .NOTES
    Author: Maximilian Gindorfer <info@mvprowess.com>
    License: MIT
  #>

  [OutputType([void])]
  param (
    [Parameter(Mandatory = $true)]
    [array]
    $Base,

    [Parameter(Mandatory = $true)]
    [array]
    $Overrides
  )

  foreach ($_override in $Overrides) {
    if ($_override.PSObject.Properties.Name -notcontains 'Name') { continue }
    $_name = $_override.Name
    if ($null -eq $_name) { continue }

    # Match by Path + Name when Path is provided, otherwise by Name alone
    $_match = foreach ($_entry in $Base) {
      if ($_override.PSObject.Properties.Name -contains 'Path' -and $_override.Path) {
        if ($_entry.Path -eq $_override.Path -and $_entry.Name -eq $_name) { $_entry; break }
      }
      else {
        if ($_entry.Name -eq $_name) { $_entry; break }
      }
    }

    if (-not $_match) { continue }

    # Only transfer keys that already exist on the base object.
    # Copy keys to a static array to avoid "collection modified" errors
    # when the assignment below mutates the same hashtable.
    $matchKeys = @($_match.Keys)
    foreach ($_key in $matchKeys) {
      if ($_override.PSObject.Properties.Name -contains $_key) {
        $_match[$_key] = $_override.$_key
      }
    }
  }
}
