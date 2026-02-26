[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateRange(1, 1000)]
    [int]$KeepNewestCount = 5,
    [switch]$IncludeBackups,
    [ValidateRange(0, 1000)]
    [int]$BackupRetentionCount = 0
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$targetDirectoryName = "mcp-server-official-x64"
$runPrefix = "$targetDirectoryName-run-"
$smokePrefix = "$targetDirectoryName-smoke"
$backupPrefix = "$targetDirectoryName-backup-"

function Remove-StaleDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Label,
        [Parameter(Mandatory = $true)]
        [string]$Prefix,
        [Parameter(Mandatory = $true)]
        [int]$KeepCount
    )

    $matchingDirectories = @(Get-ChildItem -Path $repoRoot -Directory |
            Where-Object { $_.Name.StartsWith($Prefix, [System.StringComparison]::OrdinalIgnoreCase) } |
            Sort-Object LastWriteTime -Descending)

    $directoriesToRemove = @($matchingDirectories | Select-Object -Skip $KeepCount)

    if ($directoriesToRemove.Count -eq 0) {
        Write-Host "No stale $Label directories to prune (keep newest $KeepCount)."
        return [pscustomobject]@{
            Label          = $Label
            MatchedCount   = $matchingDirectories.Count
            CandidateCount = 0
            RemovedCount   = 0
        }
    }

    Write-Host "Found $($directoriesToRemove.Count) stale $Label directories (matched total: $($matchingDirectories.Count))."
    $removedCount = 0
    foreach ($directory in $directoriesToRemove) {
        if ($PSCmdlet.ShouldProcess($directory.FullName, "Remove stale $Label directory")) {
            Write-Host "Pruning stale $Label directory: $($directory.FullName)"
            Remove-Item -LiteralPath $directory.FullName -Recurse -Force
            $removedCount++
        }
    }

    return [pscustomobject]@{
        Label          = $Label
        MatchedCount   = $matchingDirectories.Count
        CandidateCount = $directoriesToRemove.Count
        RemovedCount   = $removedCount
    }
}

Write-Host "Repo root: $repoRoot"
Write-Host "Retention: keep newest $KeepNewestCount run/smoke directories."

$runStats = Remove-StaleDirectories -Label "run" -Prefix $runPrefix -KeepCount $KeepNewestCount
$smokeStats = Remove-StaleDirectories -Label "smoke" -Prefix $smokePrefix -KeepCount $KeepNewestCount

$backupStats = [pscustomobject]@{
    Label          = "backup"
    MatchedCount   = 0
    CandidateCount = 0
    RemovedCount   = 0
}
if ($IncludeBackups) {
    if ($BackupRetentionCount -eq 0) {
        Write-Host "Backup pruning requested but disabled because BackupRetentionCount=0."
    }
    else {
        Write-Host "Backup retention: keep newest $BackupRetentionCount backup directories."
        $backupStats = Remove-StaleDirectories -Label "backup" -Prefix $backupPrefix -KeepCount $BackupRetentionCount
    }
}
else {
    if ($BackupRetentionCount -gt 0) {
        Write-Warning "BackupRetentionCount was provided without IncludeBackups; skipping backup pruning."
    }
    Write-Host "Backup pruning skipped by default. Pass -IncludeBackups to enable it."
}

Write-Host "Cleanup complete."
Write-Host "Run directories - candidates: $($runStats.CandidateCount), removed: $($runStats.RemovedCount)"
Write-Host "Smoke directories - candidates: $($smokeStats.CandidateCount), removed: $($smokeStats.RemovedCount)"
Write-Host "Backup directories - candidates: $($backupStats.CandidateCount), removed: $($backupStats.RemovedCount)"
