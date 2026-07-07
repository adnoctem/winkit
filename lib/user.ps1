function Get-UserInfo {
  <#
    .SYNOPSIS
      Get-UserInfo - Retrieves information about the current user, including their username, administrator status, and SID.
    .DESCRIPTION
      This function gathers details about the currently logged-in user by accessing the WindowsIdentity class. It checks if the user has administrator privileges and returns a hashtable containing the username, administrator status, and SID of the user.
    .EXAMPLE
      PS> Get-UserInfo
    .LINK
      https://github.com/adnoctem/winkit/lib/user.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([hashtable])]
  param ()

  $user = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  $isAdmin = (New-Object System.Security.Principal.WindowsPrincipal($user)).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

  return @{
    UserName = $user.Name
    IsAdministrator = $isAdmin
    SID = $user.User.Value
  }
}

function Get-UserSID {
  <#
    .SYNOPSIS
      Get-UserSID - Retrieves the Security Identifier (SID) for a specified user.
    .DESCRIPTION
      This function takes a username as input and attempts to retrieve the corresponding SID by creating an NTAccount object and translating it to a SecurityIdentifier. If the user is not found, it returns null and logs an error message.
    .PARAMETER UserName
      The username for which to retrieve the SID (e.g., "DOMAIN\Username").
    .EXAMPLE
      PS> Get-UserSID -UserName "DOMAIN\Username"
    .LINK
      https://github.com/adnoctem/winkit/lib/user.ps1
    .NOTES
      Author: Maximilian Gindorfer <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true)]
    [string]$UserName
  )

  try {
    $user = New-Object System.Security.Principal.NTAccount($UserName)
    $sid = $user.Translate([System.Security.Principal.SecurityIdentifier])
    return $sid.Value
  }
  catch {
    Write-Error "Could not find SID for user '$UserName'. $_"
    return $null
  }
}
