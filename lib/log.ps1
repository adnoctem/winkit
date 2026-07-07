#Requires -Version 5.0

function Show-Color {
  <#
    .SYNOPSIS
      Show the available colors to print text in.
    .DESCRIPTION
      Outputs the available colors for use within the -...Color options. These
      are the properties of the System.ConsoleColor enum.
    .EXAMPLE
      PS> Show-LogColor
    .LINK
      https://github.com/adnoctem/winkit/lib/log.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [enum]::GetValues([System.ConsoleColor]) | ForEach-Object {
    Write-Host $_ -ForegroundColor $_
  }
}


function Write-Log {
  <#
    .SYNOPSIS
      Write a message to standard output with color and optional timestamp.
    .DESCRIPTION
      Emits the supplied message to the console using the chosen foreground
      color. When -Timestamps is set, the output is prefixed with the current
      date and time before being written.
    .EXAMPLE
      PS> Write-Log -Message 'Setup completed.' -Color Green -Timestamps
    .LINK
      https://github.com/adnoctem/winkit/lib/log.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Write-Output does not support -ForegroundColor.')]
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'winkit intentionally exposes its established Write-Log helper.')]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Message,

    [Parameter(Mandatory = $false)]
    [ValidateSet(
      'Black',
      'DarkBlue',
      'DarkGreen',
      'DarkCyan',
      'DarkRed',
      'DarkMagenta',
      'DarkYellow',
      'Gray',
      'DarkGray',
      'Blue',
      'Green',
      'Cyan',
      'Red',
      'Magenta',
      'Yellow',
      'White')]
    [System.ConsoleColor]
    $Color = 'White',

    [Parameter(Mandatory = $false)]
    [switch]
    $Timestamps = $false
  )

  process {
    $timestamp = (Get-Date).DateTime
    $Content = "{0}" -f $Message

    if ($Timestamps) {
      $Content = "[{0}]: {1}" -f $timestamp, $Message
    }

    Write-Host $Content -ForegroundColor $Color
  }
}
