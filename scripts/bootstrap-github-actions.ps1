[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Low")]
param(
    [string]$Repo = "",
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "" }),
    [switch]$SetDatabaseSecret,
    [switch]$TriggerRegressionWorkflow,
    [string]$Workflow = "windows-self-hosted-access-regression.yml",
    [string]$Ref = ""
)

$ErrorActionPreference = "Stop"

function Get-RepoRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Resolve-RepoFromRemote {
    param([string]$RepoRoot)

    $remoteUrl = (& git -C $RepoRoot remote get-url origin 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($remoteUrl)) {
        return $null
    }

    $remoteUrl = $remoteUrl.Trim()
    if ($remoteUrl -match "github\.com[:/](?<owner>[^/]+)/(?<name>[^/.]+)(\.git)?$") {
        return ("{0}/{1}" -f $Matches.owner, $Matches.name)
    }

    return $null
}

function Ensure-GhAvailable {
    $gh = Get-Command -Name "gh" -ErrorAction SilentlyContinue
    if ($null -eq $gh) {
        throw "GitHub CLI (gh) is not installed or not in PATH."
    }

    return $gh.Source
}

function Assert-PathExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Path does not exist: $Path"
    }
}

$repoRoot = Get-RepoRoot
$ghExe = Ensure-GhAvailable

$resolvedRepo = $Repo
if ([string]::IsNullOrWhiteSpace($resolvedRepo)) {
    $resolvedRepo = Resolve-RepoFromRemote -RepoRoot $repoRoot
}

if ([string]::IsNullOrWhiteSpace($resolvedRepo)) {
    throw "Unable to resolve GitHub repo. Provide -Repo (example: owner/name)."
}

Write-Host ("Repo root          : {0}" -f $repoRoot)
Write-Host ("GitHub repo        : {0}" -f $resolvedRepo)
Write-Host ("GitHub CLI         : {0}" -f $ghExe)

Write-Host "Checking gh authentication..."
& $ghExe auth status
if ($LASTEXITCODE -ne 0) {
    throw "gh auth status failed. Run 'gh auth login' and retry."
}

Write-Host "Checking repository access..."
$repoIdentity = (& $ghExe repo view $resolvedRepo --json nameWithOwner --jq ".nameWithOwner" 2>$null)
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($repoIdentity)) {
    throw ("Unable to access repo '{0}' with current credentials." -f $resolvedRepo)
}

Write-Host ("Repository access OK: {0}" -f $repoIdentity.Trim())

if ($SetDatabaseSecret) {
    if ([string]::IsNullOrWhiteSpace($DatabasePath)) {
        throw "-SetDatabaseSecret requires -DatabasePath or ACCESS_DATABASE_PATH env var."
    }

    Assert-PathExists -Path $DatabasePath
    $resolvedDatabasePath = [System.IO.Path]::GetFullPath($DatabasePath)

    if ($PSCmdlet.ShouldProcess($resolvedRepo, "Set ACCESS_DATABASE_PATH secret")) {
        & $ghExe secret set ACCESS_DATABASE_PATH -R $resolvedRepo -b $resolvedDatabasePath
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to set ACCESS_DATABASE_PATH secret."
        }
        Write-Host ("Set ACCESS_DATABASE_PATH secret to: {0}" -f $resolvedDatabasePath)
    }
}

if ($TriggerRegressionWorkflow) {
    $workflowArgs = @("workflow", "run", $Workflow, "-R", $resolvedRepo)
    if (-not [string]::IsNullOrWhiteSpace($Ref)) {
        $workflowArgs += @("--ref", $Ref)
    }
    if (-not [string]::IsNullOrWhiteSpace($DatabasePath)) {
        $resolvedDatabasePath = [System.IO.Path]::GetFullPath($DatabasePath)
        $workflowArgs += @("-f", ("database_path={0}" -f $resolvedDatabasePath))
    }

    if ($PSCmdlet.ShouldProcess($resolvedRepo, ("Dispatch workflow {0}" -f $Workflow))) {
        & $ghExe @workflowArgs
        if ($LASTEXITCODE -ne 0) {
            throw ("Failed to trigger workflow '{0}'." -f $Workflow)
        }
        Write-Host ("Workflow dispatched: {0}" -f $Workflow)
    }
}

Write-Host ""
Write-Host "Bootstrap summary:"
Write-Host "- gh auth status: PASS"
Write-Host "- repository access: PASS"
Write-Host ("- ACCESS_DATABASE_PATH secret: {0}" -f $(if ($SetDatabaseSecret) { "attempted" } else { "not requested" }))
Write-Host ("- workflow dispatch: {0}" -f $(if ($TriggerRegressionWorkflow) { "attempted" } else { "not requested" }))
Write-Host ""
Write-Host "Note: org/repo admin policies (runner registration, Actions permissions, branch protection) must be granted outside this script."
