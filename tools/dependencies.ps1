#Requires -Version 5.0

<#
.SYNOPSIS
  Checks declared PowerShell module dependencies against the PowerShell Gallery
  and creates or updates one pull request with the version updates.

.DESCRIPTION
  Reads the project dependencies declared in requirements.psd1 at the
  repository root (a Name = 'version' hashtable), queries the PowerShell
  Gallery for the latest published stable version of each module, and reports
  which declared versions are out of date.

  In update mode (the default) the declared versions are bumped in place and a
  single pull request from the fixed chore/psdeps branch is created or updated
  against the default branch. Superseded timestamped chore/psdeps/* pull
  requests from earlier workflow versions are closed. When nothing is outdated,
  any open automated dependency pull request is closed instead of left behind.
  The pull request is picked up by the repository's regular CI on its own. No
  other workflows are dispatched.

  Substrate updates are reported but require a coordinated manual update of
  the manifest and script RequiredVersion pins while its API is at version 0.x.

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
  Bumps declared versions in requirements.psd1 and creates or updates the
  single dependency update pull request.

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
$PullRequestBranch = 'chore/psdeps'
$PullRequestBase = 'main'
$CommitMessage = 'chore(deps): update PowerShell module dependencies'

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

function Get-GitHubRepositoryName {
  <#
  .SYNOPSIS
    Returns the current repository in owner/name form.
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param()

  $repository = (& gh repo view --json nameWithOwner --jq '.nameWithOwner' 2>$null).Trim()
  if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($repository)) {
    throw 'Failed to determine the repository name.'
  }

  return $repository
}

function Get-DependencyPullRequest {
  <#
  .SYNOPSIS
    Finds the current and superseded automated dependency pull requests.

  .DESCRIPTION
    Returns open pull requests that target the default branch. Current requests
    use the fixed automated branch. Superseded requests use the old timestamped
    branch pattern and the same generated title.

  .PARAMETER Repository
    Repository in owner/name form.

  .PARAMETER Branch
    Fixed automated dependency branch.

  .PARAMETER BaseBranch
    Pull request base branch.

  .PARAMETER Title
    Title used by automated dependency pull requests.

  .OUTPUTS
    PSCustomObject with Current and Superseded properties.
  #>
  [CmdletBinding()]
  [OutputType([pscustomobject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Repository,

    [Parameter(Mandatory = $true)]
    [string]
    $Branch,

    [Parameter(Mandatory = $true)]
    [string]
    $BaseBranch,

    [Parameter(Mandatory = $true)]
    [string]
    $Title
  )

  $json = & gh pr list --repo $Repository --state open --limit 100 --json 'number,headRefName,baseRefName,title'
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to list open pull requests in $Repository."
  }

  $pullRequests = @()
  $jsonText = $json -join "`n"
  if (-not [string]::IsNullOrWhiteSpace($jsonText)) {
    $pullRequests = @($jsonText | ConvertFrom-Json)
  }

  return [PSCustomObject]@{
    Current    = @($pullRequests | Where-Object {
        $_.headRefName -eq $Branch -and $_.baseRefName -eq $BaseBranch
      })
    Superseded = @($pullRequests | Where-Object {
        $_.headRefName -like "$Branch/*" -and $_.baseRefName -eq $BaseBranch -and $_.title -eq $Title
      })
  }
}

function Close-DependencyPullRequest {
  <#
  .SYNOPSIS
    Closes one automated dependency pull request and deletes its branch.

  .PARAMETER Repository
    Repository in owner/name form.

  .PARAMETER Number
    Pull request number.

  .PARAMETER Reason
    Comment added before the pull request is closed.

  .OUTPUTS
    System.Boolean
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Repository,

    [Parameter(Mandatory = $true)]
    [int]
    $Number,

    [Parameter(Mandatory = $true)]
    [string]
    $Reason
  )

  & gh pr close $Number --repo $Repository --comment $Reason --delete-branch
  if ($LASTEXITCODE -ne 0) {
    Write-Warning "Failed to close pull request #$Number in $Repository."
    return $false
  }

  Write-Host "Closed pull request #$Number." -ForegroundColor Green
  return $true
}

function Remove-StaleDependencyPullRequest {
  <#
  .SYNOPSIS
    Closes automated dependency pull requests when nothing should be proposed.

  .DESCRIPTION
    Best-effort cleanup for runs that have no dependency update to publish. The
    cleanup is skipped when gh or a GitHub token is unavailable.

  .PARAMETER Branch
    Fixed automated dependency branch.

  .PARAMETER BaseBranch
    Pull request base branch.

  .PARAMETER Title
    Title used by automated dependency pull requests.

  .PARAMETER Reason
    Comment added to each closed pull request.
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Closes GitHub pull requests through gh; the script is a non-interactive CI helper.')]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Branch,

    [Parameter(Mandatory = $true)]
    [string]
    $BaseBranch,

    [Parameter(Mandatory = $true)]
    [string]
    $Title,

    [Parameter(Mandatory = $true)]
    [string]
    $Reason
  )

  if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    Write-Warning 'gh is not available; skipping stale dependency pull request cleanup.'
    return
  }
  if (-not (Get-GitHubToken)) {
    Write-Warning 'No GitHub token found; skipping stale dependency pull request cleanup.'
    return
  }

  $repository = Get-GitHubRepositoryName
  $pullRequest = Get-DependencyPullRequest -Repository $repository -Branch $Branch -BaseBranch $BaseBranch -Title $Title
  foreach ($request in @($pullRequest.Current) + @($pullRequest.Superseded)) {
    $null = Close-DependencyPullRequest -Repository $repository -Number $request.number -Reason $Reason
  }
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
          $updated[$j] = $updated[$j] -replace $VersionPattern, ('${1}' + $NewVersion + '${3}')
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
      Name            = $name
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
        Name            = $dep.Name
        DeclaredVersion = $dep.DeclaredVersion
        LatestVersion   = $latest
      })
    Write-Host ('  {0,-24} {1,-12} {2,-12} {3}' -f $dep.Name, $dep.DeclaredVersion, $latest, 'OUTDATED') -ForegroundColor Yellow
  }
  else {
    Write-Host ('  {0,-24} {1,-12} {2,-12} {3}' -f $dep.Name, $dep.DeclaredVersion, $latest, 'OK') -ForegroundColor Green
  }
}

if ($outdated.Count -eq 0) {
  Write-Host 'All declared dependencies are up to date.' -ForegroundColor Green

  if (-not $Check) {
    Remove-StaleDependencyPullRequest `
      -Branch $PullRequestBranch `
      -BaseBranch $PullRequestBase `
      -Title $CommitMessage `
      -Reason 'Dependencies are already current; closing the automated dependency update pull request.'
  }

  exit 0
}

if ($Check) {
  Write-Host ('{0} dependency update(s) required.' -f $outdated.Count) -ForegroundColor Yellow
  exit 1
}

$_manualUpdates = @($outdated | Where-Object { $_.Name -eq 'AdNoctem.Substrate.PowerShell' })
foreach ($_dependency in $_manualUpdates) {
  Write-Warning "Substrate $($_dependency.LatestVersion) requires a coordinated adoption: update requirements.psd1 and all script RequiredVersion pins together. Keeping the current pin."
  $null = $outdated.Remove($_dependency)
}
if ($outdated.Count -eq 0) {
  Remove-StaleDependencyPullRequest `
    -Branch $PullRequestBranch `
    -BaseBranch $PullRequestBase `
    -Title $CommitMessage `
    -Reason 'Only the manual Substrate update remains; closing the automated dependency update pull request.'

  exit 1
}

Write-Host ('Bumping {0} dependency version(s) and updating the dependency pull request...' -f $outdated.Count) -ForegroundColor Cyan

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  throw 'git is required to update a dependency pull request.'
}
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
  throw 'gh (GitHub CLI) is required to update a dependency pull request.'
}
$githubToken = Get-GitHubToken
if (-not $githubToken) {
  throw 'No GitHub token found; set GH_TOKEN or GITHUB_TOKEN to push the branch and update the pull request.'
}

$repoName = Get-GitHubRepositoryName
$pullRequest = Get-DependencyPullRequest `
  -Repository $repoName `
  -Branch $PullRequestBranch `
  -BaseBranch $PullRequestBase `
  -Title $CommitMessage

foreach ($dep in $outdated) {
  $escapedName = [regex]::Escape($dep.Name)
  $namePattern = "^\s*$escapedName\s*=\s*'"
  $versionPattern = "^(\s*$escapedName\s*=\s*')([^']+)(')"
  $encoding = New-Object System.Text.UTF8Encoding($true)

  $lines = @(Get-Content -LiteralPath $RequirementsPath)
  $lines = Update-VersionLine -Lines $lines -NamePattern $namePattern -VersionPattern $versionPattern -NewVersion $dep.LatestVersion
  [System.IO.File]::WriteAllLines($RequirementsPath, $lines, $encoding)
}

& git -C $RepoRoot checkout -B $PullRequestBranch
if ($LASTEXITCODE -ne 0) {
  throw "Failed to create branch '$PullRequestBranch'."
}

& git -C $RepoRoot add -- $RequirementsPath
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to stage dependency manifest changes.'
}

& git -C $RepoRoot -c user.name='github-actions[bot]' -c user.email='41898282+github-actions[bot]@users.noreply.github.com' commit -m $CommitMessage
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to commit dependency updates.'
}

& git -C $RepoRoot push --force "https://x-access-token:$($githubToken)@github.com/$repoName.git" $PullRequestBranch
if ($LASTEXITCODE -ne 0) {
  throw "Failed to push branch '$PullRequestBranch'."
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

if (@($pullRequest.Current).Count -gt 0) {
  $current = @($pullRequest.Current)[0]
  & gh pr edit $current.number --repo $repoName --title $CommitMessage --body $body
  if ($LASTEXITCODE -ne 0) {
    throw 'Failed to update the dependency pull request.'
  }

  Write-Host ("Updated pull request #{0} with {1} dependency update(s)." -f $current.number, $outdated.Count) -ForegroundColor Green
}
else {
  $prOutput = & gh pr create --repo $repoName --base $PullRequestBase --head $PullRequestBranch --title $CommitMessage --body $body
  if ($LASTEXITCODE -ne 0) {
    throw 'Failed to open the dependency pull request.'
  }

  $prNumber = $null
  if ($prOutput -match 'pull/(\d+)') {
    $prNumber = $Matches[1]
  }

  if ($prNumber) {
    Write-Host ("Created pull request #{0} with {1} dependency update(s)." -f $prNumber, $outdated.Count) -ForegroundColor Green
  }
  else {
    Write-Host ("Pushed dependency updates to branch '{0}' and created a pull request." -f $PullRequestBranch) -ForegroundColor Green
  }
}

foreach ($request in @($pullRequest.Superseded)) {
  $null = Close-DependencyPullRequest `
    -Repository $repoName `
    -Number $request.number `
    -Reason 'Superseded by the single automated dependency update pull request.'
}
