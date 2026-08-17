#Requires -Version 5.0

<#
.SYNOPSIS
  Checks declared PowerShell module dependencies against the PowerShell Gallery
  and opens a pull request with the version updates.

.DESCRIPTION
  Reads the project dependencies declared in requirements.psd1 at the
  repository root (a Name = 'version' hashtable), queries the PowerShell
  Gallery for the latest published stable version of each module, and reports
  which declared versions are out of date.

  In update mode (the default) the declared versions are bumped in place and a
  pull request is opened against the default branch with the version diffs. The
  pull request is picked up by the repository's regular CI on its own. No other
  workflows are dispatched.

  In -Check mode no files are modified and the script exits with code 1 when any
  dependency is outdated, making it suitable for local use and CI checks.

  The gallery is queried through the NuGet v2 OData endpoint. When the
  PSGALLERY_API_KEY environment variable is set it is sent as the X-NuGet-ApiKey
  header to avoid rate limiting. The gh CLI and the git push use the GitHub
  token taken from the conventional environment variables GH_TOKEN,
  GITHUB_TOKEN, GH_ENTERPRISE_TOKEN, or GITHUB_ENTERPRISE_TOKEN, in that order
  of precedence.

  The script is intended to run from the scheduled dependencies.yaml workflow
  on a clean checkout of the default branch.

.PARAMETER Check
  Report outdated dependencies without modifying files. Exits with code 1 when
  any dependency is outdated.

.EXAMPLE
  PS> .\winkit.ps1 dependencies -Check
  Reports outdated dependencies without making any changes.

.EXAMPLE
  PS> .\winkit.ps1 dependencies
  Bumps declared versions in requirements.psd1 and opens a pull request with
  the diffs.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'This script is intended for CI output and Write-Host is appropriate for user feedback.')]
[CmdletBinding()]
param(
  [switch]$Check
)

$ErrorActionPreference = 'Stop'

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$RepoRoot = Split-Path -Path $PSScriptRoot -Parent
$RequirementsPath = Join-Path -Path $RepoRoot -ChildPath 'requirements.psd1'

$PSGalleryQueryUri = "https://www.powershellgallery.com/api/v2/FindPackagesById()?id='{0}'"

function Get-PSGalleryLatestVersion {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $headers = @{ Accept = 'application/atom+xml' }
  if ($env:PSGALLERY_API_KEY) {
    $headers['X-NuGet-ApiKey'] = $env:PSGALLERY_API_KEY
  }

  $uri = $PSGalleryQueryUri -f [uri]::EscapeDataString($Name)
  $response = Invoke-WebRequest -Uri $uri -Headers $headers -UseBasicParsing

  $document = [xml]$response.Content
  $namespace = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
  $namespace.AddNamespace('a', 'http://www.w3.org/2005/Atom')
  $namespace.AddNamespace('m', 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata')
  $namespace.AddNamespace('d', 'http://schemas.microsoft.com/ado/2007/08/dataservices')

  $versions = @()
  foreach ($entry in $document.SelectNodes('//a:entry', $namespace)) {
    $versionNode = $entry.SelectSingleNode('m:properties/d:Version', $namespace)
    $prereleaseNode = $entry.SelectSingleNode('m:properties/d:IsPrerelease', $namespace)
    if (-not $versionNode) {
      continue
    }
    if ($prereleaseNode -and $prereleaseNode.InnerText -eq 'true') {
      continue
    }

    try {
      $null = [version]$versionNode.InnerText
      $versions += $versionNode.InnerText
    }
    catch {
      Write-Host "  Skipping unparsable version '$($versionNode.InnerText)' for $Name" -ForegroundColor DarkGray
    }
  }

  if ($versions.Count -eq 0) {
    return $null
  }

  return ($versions | Sort-Object -Property { [version]$_ } -Descending | Select-Object -First 1)
}

function Get-GitHubToken {
  [CmdletBinding()]
  param()

  foreach ($token in @($env:GH_TOKEN, $env:GITHUB_TOKEN, $env:GH_ENTERPRISE_TOKEN, $env:GITHUB_ENTERPRISE_TOKEN)) {
    if ($token) {
      return $token
    }
  }

  return $null
}

function Update-VersionLine {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory string transform; no system state is modified.')]
  param(
    [Parameter(Mandatory = $true)]
    $Lines,

    [Parameter(Mandatory = $true)]
    [string]$NamePattern,

    [Parameter(Mandatory = $true)]
    [string]$VersionPattern,

    [Parameter(Mandatory = $true)]
    [string]$NewVersion
  )

  $updated = New-Object System.Collections.Generic.List[string]
  foreach ($line in $Lines) {
    $updated.Add($line)
  }

  for ($i = 0; $i -lt $updated.Count; $i++) {
    if ($updated[$i] -match $NamePattern) {
      # requirements.psd1 declares each dependency on a single line
      # (Name = 'version'), so the version lives on the same line as the name.
      for ($j = $i; $j -lt $updated.Count -and $j -le $i + 3; $j++) {
        if ($updated[$j] -match $VersionPattern) {
          $updated[$j] = $updated[$j] -replace $VersionPattern, ('${1}' + $NewVersion + '${2}')
          break
        }
      }
    }
  }

  return $updated.ToArray()
}

Write-Host 'Checking declared PowerShell module dependencies against the PowerShell Gallery...' -ForegroundColor Cyan

$requirements = Import-PowerShellDataFile -LiteralPath $RequirementsPath

$declared = New-Object System.Collections.Generic.List[object]
foreach ($name in $requirements.Keys) {
  $declared.Add([pscustomobject]@{
      Name = $name
      DeclaredVersion = $requirements[$name]
    })
}

$outdated = New-Object System.Collections.Generic.List[object]

Write-Host ('  {0,-24} {1,-12} {2,-12} {3}' -f 'Name', 'Declared', 'Latest', 'Status')

foreach ($dep in $declared) {
  $latest = Get-PSGalleryLatestVersion -Name $dep.Name

  if (-not $latest) {
    Write-Host ('  {0,-24} {1,-12} {2,-12} {3}' -f $dep.Name, '--', '--', 'UNKNOWN') -ForegroundColor Red
    continue
  }

  $isOutdated = $false
  if ($dep.DeclaredVersion) {
    $isOutdated = [version]$latest -gt [version]$dep.DeclaredVersion
  }

  if ($isOutdated) {
    $outdated.Add([pscustomobject]@{
        Name = $dep.Name
        DeclaredVersion = $dep.DeclaredVersion
        LatestVersion = $latest
      })
    Write-Host ('  {0,-24} {1,-12} {2,-12} {3}' -f $dep.Name, $dep.DeclaredVersion, $latest, 'OUTDATED') -ForegroundColor Yellow
  }
  else {
    Write-Host ('  {0,-24} {1,-12} {2,-12} {3}' -f $dep.Name, $dep.DeclaredVersion, $latest, 'OK') -ForegroundColor Green
  }
}

if ($outdated.Count -eq 0) {
  Write-Host 'All declared dependencies are up to date.' -ForegroundColor Green
  exit 0
}

if ($Check) {
  Write-Host ('{0} dependency update(s) required.' -f $outdated.Count) -ForegroundColor Yellow
  exit 1
}

Write-Host ('Bumping {0} dependency version(s) and opening a pull request...' -f $outdated.Count) -ForegroundColor Cyan

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  throw 'git is required to open a dependency update pull request.'
}
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
  throw 'gh (GitHub CLI) is required to open a dependency update pull request.'
}
$githubToken = Get-GitHubToken
if (-not $githubToken) {
  throw 'No GitHub token found; set GH_TOKEN or GITHUB_TOKEN to push the branch and open the pull request.'
}

foreach ($dep in $outdated) {
  $escapedName = [regex]::Escape($dep.Name)
  $namePattern = "^\s*$escapedName\s*=\s*'"
  $versionPattern = "^(\s*$escapedName\s*=\s*')([^']+)(')"
  $encoding = New-Object System.Text.UTF8Encoding($true)

  $lines = @(Get-Content -LiteralPath $RequirementsPath)
  $lines = Update-VersionLine -Lines $lines -NamePattern $namePattern -VersionPattern $versionPattern -NewVersion $dep.LatestVersion
  [System.IO.File]::WriteAllLines($RequirementsPath, $lines, $encoding)
}

$branch = 'chore/psdeps/{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
$commitMessage = 'chore(deps): update PowerShell module dependencies'

& git -C $RepoRoot checkout -b $branch
if ($LASTEXITCODE -ne 0) {
  throw "Failed to create branch '$branch'."
}

& git -C $RepoRoot add -- $RequirementsPath
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to stage dependency manifest changes.'
}

& git -C $RepoRoot -c user.name='github-actions[bot]' -c user.email='41898282+github-actions[bot]@users.noreply.github.com' commit -m $commitMessage
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to commit dependency updates.'
}

$repoName = (& gh repo view --json nameWithOwner --jq '.nameWithOwner' 2>$null).Trim()
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to determine the repository name.'
}

& git -C $RepoRoot push "https://x-access-token:$($githubToken)@github.com/$repoName.git" $branch
if ($LASTEXITCODE -ne 0) {
  throw "Failed to push branch '$branch'."
}

$bodyLines = @(
  '| Dependency | Declared | Latest |',
  '|---|---|---|'
)
foreach ($dep in $outdated) {
  $bodyLines += '| {0} | {1} | {2} |' -f $dep.Name, $dep.DeclaredVersion, $dep.LatestVersion
}
$bodyLines += ''
$bodyLines += 'Generated by the scheduled PSGallery dependency check in `.github/workflows/dependencies.yaml`.'
$body = $bodyLines -join "`n"

$prOutput = & gh pr create --repo $repoName --base main --head $branch --title $commitMessage --body $body
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to open the dependency update pull request.'
}

$prNumber = $null
if ($prOutput -match 'pull/(\d+)') {
  $prNumber = $Matches[1]
}

if ($prNumber) {
  Write-Host ("Opened pull request #{0} with {1} dependency update(s)." -f $prNumber, $outdated.Count) -ForegroundColor Green
}
else {
  Write-Host ("Pushed dependency updates to branch '{0}' and opened a pull request." -f $branch) -ForegroundColor Green
}
