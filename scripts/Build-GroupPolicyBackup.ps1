#Requires -Version 5.0
#Requires -Modules @{ ModuleName = 'PSFoundation'; ModuleVersion = '0.1.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Metadata is a singular mass noun; there is no other option here.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Build-GroupPolicyBackup is the intended entry verb for this build pipeline.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'StagingRoot', Justification = 'Stub parameter - TODO(1) stubs the function body.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'BackupDir', Justification = 'Stub parameter - TODO(2) stubs the function body.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'DisplayName', Justification = 'Stub parameter - TODO(2) stubs the function body.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Guid', Justification = 'Stub parameter - TODO(2) stubs the function body.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Sources', Justification = 'Stub parameter - TODO(2) stubs the function body.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'WorkDir', Justification = 'Stub parameter - TODO(2) stubs the function body.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = '$pol is assigned as a placeholder for TODO(2) pipeline.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Set-BackupMetadata is a TODO(2) stub; SupportsShouldProcess is declared for future implementation.')]

<#
.SYNOPSIS
  Compiles winkit policy text sources into a domain-importable GPO backup.

.DESCRIPTION
  ============================================================================
  *** WORK IN PROGRESS - NOT YET FUNCTIONAL ***

  This is committed boilerplate. The function shape, parameter contract, and
  overall flow are settled, but two pieces are stubbed and MUST be completed
  before this produces a usable backup:

    [ ] TODO(1): Vendor a reference GPO backup skeleton. Create an empty GPO
                 in a test domain (GPMC -> right-click Group Policy Objects ->
                 New -> name it 'winkit-skeleton' -> Back Up), then commit the
                 resulting backup folder under resources/policies/skeleton/.
                 We wrap the LGPO-built registry.pol into a COPY of this
                 skeleton rather than constructing Backup.xml / bkupInfo.xml /
                 gpreport.xml by hand (their schema is fiddly and brittle to
                 build from scratch). See Initialize-FromSkeleton below.

    [ ] TODO(2): Implement the registry.pol generation via LGPO.exe and the
                 skeleton-rewrite logic (display name, GUID, timestamps).
                 See Build-GroupPolicyBackup body.

  Depends on the LGPO functions in lib/policies.ps1 (Resolve-LGPOSource,
  Install-LGPO, Invoke-LGPO) and on the policy text sources under
  resources/policies/.

  Once complete, the produced backup goes to dist/GP/gpo-backups/ and is
  consumed by the per-company AD repo's Import-GPO / New-GPLink workflow -
  NOT by winkit itself. winkit produces the artifact; it never touches a
  domain.
  ============================================================================

.PARAMETER SourcePath
  Directory containing the policy text source files to compile. Defaults to
  resources/policies/ resolved relative to this script.

.PARAMETER OutputPath
  Directory where the GPO backup folder is written. Defaults to
  dist/GP/gpo-backups/ (top-level dist/, gitignored).

.PARAMETER DisplayName
  The GPO display name baked into the backup metadata. This is what a domain
  admin sees in GPMC after import. Defaults to 'winkit Baseline'.

.PARAMETER SkeletonPath
  Path to the vendored reference GPO backup skeleton (see TODO(1)). Defaults
  to resources/policies/skeleton/.

.EXAMPLE
  .\Build-GroupPolicyBackup.ps1 -DisplayName 'winkit Baseline 2026-06'
  (Once complete) compiles all policy text sources into a named GPO backup.

.LINK
  https://github.com/adnoctem/winkit

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
  WIP - see the description block. Running this today will throw at the first
  stubbed step by design, so it can't produce a misleadingly-empty backup.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [Parameter(Mandatory = $false)]
  [string]
  $SourcePath = (Join-Path $PSScriptRoot '..\resources\policies'),

  [Parameter(Mandatory = $false)]
  [string]
  $OutputPath = (Join-Path $PSScriptRoot '..\dist\GP\gpo-backups'),

  [Parameter(Mandatory = $false)]
  [string]
  $DisplayName = 'winkit Baseline',

  [Parameter(Mandatory = $false)]
  [string]
  $SkeletonPath = (Join-Path $PSScriptRoot '..\resources\policies\skeleton')
)

Import-Module PSFoundation -Force

# -----------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Stage helpers - each is a clear seam. The stubbed ones throw so an incomplete
# run fails loudly instead of emitting a broken backup.
# ---------------------------------------------------------------------------

function Initialize-FromSkeleton {
  <#
    .SYNOPSIS
      Copies the vendored reference GPO backup skeleton into a fresh working
      directory and returns the new backup's path + GUID.
    .DESCRIPTION
      WIP / TODO(1): Requires resources/policies/skeleton/ to exist (a real
      empty-GPO backup produced by GPMC). Logic to implement:
        1. Generate a fresh GUID for this backup instance.
        2. Copy the skeleton folder structure to a {NewGuid} folder under
           a staging area.
        3. Return the staged path and GUID for downstream rewriting.
      The skeleton gives us valid Backup.xml / bkupInfo.xml / gpreport.xml
      we then customise, instead of hand-authoring that XML.
  #>
  [CmdletBinding()]
  param([string]$Skeleton, [string]$StagingRoot)

  if (-not (Test-Path $Skeleton)) {
    throw "GPO skeleton not found at '$Skeleton'. See TODO(1) in the file header: " +
    "create an empty GPO in GPMC, back it up, and commit it under resources/policies/skeleton/."
  }

  throw "NOT IMPLEMENTED (TODO 1): skeleton copy + GUID assignment. " +
  "Skeleton exists at '$Skeleton' - implement the copy/rewrite next."
}

function ConvertTo-RegistryPol {
  <#
    .SYNOPSIS
      Compiles a set of policy text sources into Machine and User registry.pol
      files using LGPO.exe.
    .DESCRIPTION
      WIP / TODO(2): Logic to implement:
        1. Ensure LGPO is installed (Install-LGPO from lib/policies.ps1).
        2. For each policy text source, or for a merged source, invoke
           LGPO.exe /r <text> /w <output.pol> to build the binary .pol.
           (LGPO separates Machine vs User by the 'Computer'/'User' scope
           lines inside the text file - confirm whether one combined pass
           or per-scope passes is cleaner during implementation.)
        3. Return the paths to the built Machine\registry.pol and
           User\registry.pol.
  #>
  [CmdletBinding()]
  param([string]$Sources, [string]$WorkDir)

  throw "NOT IMPLEMENTED (TODO 2): LGPO.exe policy text -> registry.pol compilation."
}

function Set-BackupMetadata {
  <#
    .SYNOPSIS
      Rewrites the skeleton's Backup.xml / bkupInfo.xml to carry this build's
      display name, GUID, and timestamp.
    .DESCRIPTION
      WIP / TODO(2): Logic to implement:
        1. Update bkupInfo.xml: GPODisplayName, backup timestamp.
        2. Update Backup.xml: matching display name / GUID references.
        3. Optionally regenerate gpreport.xml (display-only; a minimal/stale
           report still imports fine, so this is lowest priority).
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  param([string]$BackupDir, [string]$DisplayName, [guid]$Guid)

  throw "NOT IMPLEMENTED (TODO 2): backup metadata rewrite."
}

# ---------------------------------------------------------------------------
# Main flow - orchestration is in place; calls into the stubs above.
# ---------------------------------------------------------------------------

function Build-GroupPolicyBackup {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [string]$Sources,
    [string]$Output,
    [string]$Name,
    [string]$Skeleton
  )

  Write-Log -Message 'WIP: Build-GroupPolicyBackup invoked.' -Color Yellow
  Write-Log -Message "Sources : $Sources" -Color Gray
  Write-Log -Message "Output  : $Output" -Color Gray
  Write-Log -Message "Name    : $Name" -Color Gray

  if (-not (Test-Path $Sources)) {
    throw "Source directory not found: $Sources"
  }
  $policyFiles = @(Get-ChildItem -Path $Sources -Filter '*.txt' -ErrorAction SilentlyContinue)
  if ($policyFiles.Count -eq 0) {
    Write-Log -Message "No policy text source files found in '$Sources'. Nothing to compile yet." -Color Yellow
  }
  else {
    Write-Log -Message "Found $($policyFiles.Count) policy text source file(s) to compile." -Color Gray
  }

  New-Item -ItemType Directory -Path $Output -Force | Out-Null

  # --- Pending pipeline (each call currently throws by design) ---
  $staging = Join-Path $env:TEMP "winkit-gpo-build-$(New-Guid)"

  $backup = Initialize-FromSkeleton -Skeleton $Skeleton -StagingRoot $staging      # TODO(1)
  $pol = ConvertTo-RegistryPol   -Sources  $Sources  -WorkDir     $staging      # TODO(2)
  Set-BackupMetadata -BackupDir $backup.Path -DisplayName $Name -Guid $backup.Guid # TODO(2)

  # Final move of the completed backup into $Output would go here.
  # Copy-Item $backup.Path -Destination $Output -Recurse -Force

  Write-Log -Message "Build complete: $Output" -Color Green
}

# Entry point
Build-GroupPolicyBackup -Sources $SourcePath -Output $OutputPath -Name $DisplayName -Skeleton $SkeletonPath
