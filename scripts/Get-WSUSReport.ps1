#Requires -Version 5.0
#Requires -Module UpdateServices
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '1.0.0' }

<#
.SYNOPSIS
  Reports updates from a WSUS server.

.DESCRIPTION
  Queries a WSUS server (Get-WsusServer) for updates matching the given
  classification and status - by default critical updates that failed or are
  needed, the patch-compliance signal commonly used for audit purposes - and
  emits a structured report. No HTML templating: output is structured objects
  with optional CSV export via -OutputPath, per winkit conventions.

  Requires the WSUS management module (UpdateServices), which ships with the
  WSUS console feature.

.PARAMETER ServerName
  WSUS server to query. Defaults to the local machine.

.PARAMETER PortNumber
  WSUS port. Defaults to 8530.

.PARAMETER UseSsl
  Connect over HTTPS.

.PARAMETER Classification
  Update classification filter. Defaults to Critical.

.PARAMETER Status
  Update status filter. Defaults to FailedOrNeeded.

.PARAMETER OutputPath
  Optional CSV file to write the report to.

.EXAMPLE
  PS> ./Get-WSUSReport.ps1
  Reports critical updates that failed or are needed on the local WSUS server.

.EXAMPLE
  PS> ./Get-WSUSReport.ps1 -ServerName wsus01.company.com -OutputPath wsus-critical.csv

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string]
  $ServerName,

  [Parameter(Mandatory = $false)]
  [ValidateRange(1, 65535)]
  [int]
  $PortNumber = 8530,

  [Parameter(Mandatory = $false)]
  [switch]
  $UseSsl,

  [Parameter(Mandatory = $false)]
  [ValidateSet('Critical', 'Definition', 'Drivers', 'FeaturePack', 'Security', 'ServicePack', 'Tool', 'Update', 'UpdateRollup')]
  [string]
  $Classification = 'Critical',

  [Parameter(Mandatory = $false)]
  [ValidateSet('Failed', 'FailedOrNeeded', 'Installed', 'InstalledOrNotApplicable', 'Needed', 'NotApproved', 'NotInstalled', 'Unknown')]
  [string]
  $Status = 'FailedOrNeeded',

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath
)

Import-Module PSFoundation -Force

Write-Log -Message "Querying WSUS server $(if ($ServerName) { $ServerName } else { 'localhost' }) for $Classification updates with status $Status..." -Color Cyan

try {
  $wsusParams = @{ PortNumber = $PortNumber }
  if ($ServerName) {
    $wsusParams.Name = $ServerName
  }
  if ($UseSsl) {
    $wsusParams.UseSsl = $true
  }
  $wsusServer = Get-WsusServer @wsusParams -ErrorAction Stop
}
catch {
  Write-Log -Message "Could not connect to the WSUS server: $($_.Exception.Message)" -Color Red
  exit 1
}

try {
  $updates = @(Get-WsusUpdate -UpdateServer $wsusServer -Classification $Classification -Status $Status -ErrorAction Stop)
}
catch {
  Write-Log -Message "WSUS update query failed: $($_.Exception.Message)" -Color Red
  exit 1
}

$report = @()
foreach ($update in $updates) {
  $report += [pscustomobject]@{
    Title = $update.Title
    Classification = $update.Classification
    Status = $update.Status
    KnowledgeBaseArticle = $update.KnowledgeBaseArticle
    ReleaseDate = $update.ReleaseDate
  }
}

$report = @($report | Sort-Object Title)

if ($report.Count -eq 0) {
  Write-Log -Message "No $Classification updates with status $Status found." -Color Green
  exit 0
}

$report | Write-Output

if ($OutputPath) {
  $report | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
  Write-Log -Message "Report written to $OutputPath" -Color Gray
}

Write-Log -Message "$($report.Count) update(s) found." -Color Yellow
exit 0
