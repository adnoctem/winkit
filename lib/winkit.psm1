# --------------------------------------------------------------------
# winkit.psm1 - Module wrapper for the 'winkit' function library
# --------------------------------------------------------------------

$_common = Join-Path -Path $PSScriptRoot -ChildPath 'common.ps1'
if (Test-Path -LiteralPath $_common) {
  . $_common
}

$files = Get-ChildItem -Path $PSScriptRoot -Filter *.ps1 -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'common.ps1' }

foreach ($file in $files) {
  . $file.FullName
}

$publicAliases = @(
  'Get-Network',
  'Get-Prefix',
  'Get-NetworkCIDR',
  'Get-PrefixCIDR'
)

$publicFunctions = @(
  # common.ps1
  'New-OperationResult',
  'Add-OperationResult',
  'Write-OperationResultLog',
  'Export-RegistrySettingState',
  'ConvertTo-RegistrySettingResult',

  # data.ps1
  'Convert-Quote',
  'Merge-ObjectArrays',

  # devices.ps1
  'Get-PrintDevice',
  'Get-DefaultPrintDevice',
  'Set-DefaultPrintDevice',
  'Get-ScanDevice',

  # log.ps1
  'Show-Color',
  'Write-Log',

  # interop.ps1
  'Remove-ComObject',
  'Invoke-ComGarbageCollection',
  'Get-OutlookInstallation',
  'Find-OutlookRepairTool',
  'Connect-Outlook',
  'Get-OutlookStoreRoot',
  'Add-OutlookStoreRoot',
  'Get-OutlookSubFolder',

  # networking.ps1
  'Get-DefaultNetworkAdapter',
  'Get-IPAddress',
  'Get-SubnetMask',
  'Get-DefaultGateway',
  'Get-DNSServer',
  'Get-MACAddress',
  'Get-NetworkPrefix',
  'Get-NetworkPrefixCIDR',
  'Get-BroadcastAddress',
  'Get-MulticastAddress',
  'Test-IPv4Address',
  'Test-IPv6Address',

  # packages.ps1
  'New-PackageLifecycleResult',
  'Get-InstalledProgramCount',
  'Get-AppxPackageCount',
  'Get-PackageCount',
  'Get-Win32Program',
  'Find-Win32Program',
  'Install-Win32Program',
  'Uninstall-Win32Program',
  'Get-UPFAppxPackage',
  'Find-UPFAppxPackage',
  'Test-UPFAppxPackageRemovalSafety',
  'Install-UPFAppxPackage',
  'Update-UPFAppxPackage',
  'Repair-UPFAppxPackage',
  'Reset-UPFAppxPackage',
  'Uninstall-UPFAppxPackage',
  'Install-UPFAppxPackageSet',
  'Uninstall-UPFAppxPackageSet',
  'Install-Win32ProgramFromWinGet',
  'Update-Win32ProgramFromWinGet',
  'Uninstall-Win32ProgramFromWinGet',

  # permissions.ps1
  'Request-AdministratorPrivilege',
  'Test-Elevation',

  # registry.ps1
  'ConvertTo-RegistryProviderPath',
  'Resolve-RegistryPath',
  'Get-RegistryKey',
  'Set-RegistryKey',
  'Remove-RegistryKey',
  'Get-RegistryValue',
  'Set-RegistryValue',
  'Remove-RegistryValue',
  'Test-RegistryPath',
  'Test-RegistryValue',
  'Get-RegistryValueKind',
  'Mount-DefaultUserHive',
  'Dismount-DefaultUserHive',
  'Export-RegistryKey',
  'Search-RegistryKey',

  # security.ps1
  'Get-DefenderThreatDetection',
  'Get-DefenderThreat',
  'Get-DefenderThreatDescriptionURL',
  'Add-DefenderExclusion',
  'Enable-WSLFirewallRule',
  'Disable-JetBrainsFirewallRule',
  'Find-NewlyWrittenObject',
  'Invoke-SafeProcess',
  'Export-EventLog',
  'Get-ScheduledTaskAction',
  'Get-WMIPersistence',

  # settings.ps1
  'Get-DefaultApp',

  # system.ps1
  'Get-OSBuildNumber',
  'Get-OSDisplayVersion',
  'Get-OSEdition',
  'Get-OSProductName',
  'Get-OSVersionInfo',
  'Get-SystemMemory',
  'Get-SystemDisk',
  'Get-Hostname',
  'Get-SystemUptime',
  'Get-SystemInfo',
  'Get-SystemPaths',

  # user.ps1
  'Get-UserInfo',
  'Get-UserSID',

  # policies.ps1
  'Resolve-LGPOSource',
  'Test-LGPOSourceAvailability',
  'Install-LGPO',
  'Test-LGPOInstalled',
  'Invoke-LGPO',

  # updates.ps1
  'Test-PSWindowsUpdateAvailable',
  'Get-WindowsUpdate',
  'Install-WindowsUpdate',
  'Hide-WindowsUpdate',
  'Get-WindowsUpdateHistory',
  'Uninstall-WindowsUpdate',
  'Test-WindowsUpdateRebootRequired',
  'Get-WindowsUpdateConfiguration',
  'Get-MSStoreUpdate',
  'Install-MSStoreUpdate'
)

Export-ModuleMember -Function $publicFunctions -Alias $publicAliases
