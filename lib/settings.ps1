function Get-DefaultApp {
  <#
    .SYNOPSIS
      Get-DefaultApp - Retrieves the default application associated with a given file extension.
    .DESCRIPTION
      This function queries the Windows registry to determine the default application that is set to open files with the specified extension. It first checks the user's file association settings and then retrieves the command used to open files of that type.
    .PARAMETER FileExtension
      The file extension (including the leading dot) for which to retrieve the default application (e.g., ".txt").
    .EXAMPLE
      PS> Get-DefaultApp -FileExtension ".txt"
    .LINK
      https://github.com/adnoctem/winkit/lib/settings.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true)]
    [string]$FileExtension
  )

  try {
    $assoc = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$FileExtension\UserChoice" -ErrorAction Stop
  }
  catch {
    Write-Error "Could not retrieve default application for '$FileExtension'. $_"
  }

  $command = (Get-ItemProperty "HKCR:\$($assoc.ProgId)\shell\open\command" -ErrorAction Stop).'(default)'
  return $command
}
