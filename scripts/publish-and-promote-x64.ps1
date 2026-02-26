[CmdletBinding()]
param(
    [string]$Project = "MS.Access.MCP.Official/MS.Access.MCP.Official.csproj",
    [string]$Configuration = "Release",
    [string]$RuntimeIdentifier = "win-x64",
    [string]$TargetDirectoryName = "mcp-server-official-x64",
    [bool]$SelfContained = $true,
    [bool]$StopServerProcesses = $true,
    [int]$BackupRetentionCount = 0,
    [switch]$SkipSmokeTest,
    [switch]$RunRegression,
    [string]$RegressionDatabasePath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
$validationManifestName = "release-validation.json"

if ($BackupRetentionCount -lt 0) {
    throw "BackupRetentionCount cannot be negative."
}

function Stop-LockingServerProcesses {
    param(
        [string]$ProcessName,
        [string]$TargetPathPrefix,
        [switch]$StopAllWhenPathUnavailable
    )

    $candidates = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($null -eq $candidates) {
        return 0
    }

    $stoppedCount = 0

    foreach ($process in $candidates) {
        $procPath = $null
        try {
            $procPath = $process.Path
        }
        catch {
            # Accessing Path can fail for short-lived processes.
        }

        $shouldStop = $false
        if (-not [string]::IsNullOrWhiteSpace($procPath) -and
            $procPath.StartsWith($TargetPathPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $shouldStop = $true
        }

        if (-not $shouldStop) {
            if ([string]::IsNullOrWhiteSpace($procPath)) {
                if ($StopAllWhenPathUnavailable) {
                    Write-Warning "Process $($process.Id) path is unavailable; forcing stop by process name."
                    $shouldStop = $true
                }
                else {
                    Write-Host "Skipping process $($process.Id) because its executable path is unavailable."
                }
            }
        }

        if (-not $shouldStop) {
            continue
        }

        Write-Host "Stopping running server process $($process.Id) ($($process.ProcessName))"
        try {
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
            $stoppedCount++
        }
        catch {
            Write-Warning ("Failed to stop process {0}: {1}" -f $process.Id, $_.Exception.Message)
        }
    }

    return $stoppedCount
}

function Restore-PreviousTargetAfterRegressionFailure {
    param(
        [string]$RepoRoot,
        [string]$TargetPath,
        [string]$TargetDirectoryName,
        [string]$BackupPath
    )

    if (-not (Test-Path -LiteralPath $BackupPath)) {
        Write-Warning "Rollback skipped: backup directory not found: $BackupPath"
        return $false
    }

    $archiveBaseName = "$TargetDirectoryName-regression-failed-$(Get-Date -Format "yyyyMMdd-HHmmss")"
    $archiveAttempt = 0
    do {
        $archiveSuffix = if ($archiveAttempt -eq 0) { "" } else { "-$archiveAttempt" }
        $archiveName = "$archiveBaseName$archiveSuffix"
        $archivePath = Join-Path $RepoRoot $archiveName
        $archiveAttempt++
    } while (Test-Path -LiteralPath $archivePath)

    if (Test-Path -LiteralPath $TargetPath) {
        Write-Warning "Regression failed after promotion; archiving current target to: $archivePath"
        try {
            Rename-Item -LiteralPath $TargetPath -NewName $archiveName -ErrorAction Stop
        }
        catch {
            Write-Warning ("Rollback failed while archiving promoted target: {0}" -f $_.Exception.Message)
            Write-Warning "Backup left untouched at: $BackupPath"
            return $false
        }
    }
    else {
        Write-Warning "Regression failed after promotion; active target directory is missing, attempting direct backup restore."
    }

    Write-Warning "Restoring previous target from backup: $BackupPath"
    try {
        Rename-Item -LiteralPath $BackupPath -NewName $TargetDirectoryName -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning ("Rollback failed while restoring backup: {0}" -f $_.Exception.Message)
        if ((Test-Path -LiteralPath $archivePath) -and -not (Test-Path -LiteralPath $TargetPath)) {
            Write-Warning "Attempting to restore archived promoted target to keep a runnable target in place."
            try {
                Rename-Item -LiteralPath $archivePath -NewName $TargetDirectoryName -ErrorAction Stop
                Write-Warning "Archived promoted target restored to: $TargetPath"
            }
            catch {
                Write-Warning ("Failed to restore archived promoted target: {0}" -f $_.Exception.Message)
                Write-Warning "Archived promoted target retained at: $archivePath"
            }
        }

        return $false
    }
}

function Ensure-RegressionEnvironmentClean {
    param(
        [string]$DatabasePath
    )

    $processNames = @("MSACCESS", "MS.Access.MCP.Official")
    @((Get-Process -Name $processNames -ErrorAction SilentlyContinue)) | ForEach-Object {
        try {
            Stop-Process -Id $_.Id -Force -ErrorAction Stop
            Write-Host "Stopped lingering process $($_.ProcessName) (PID $($_.Id))"
        }
        catch {
            Write-Warning ("Cleanup failed to stop process {0}: {1}" -f $_.Id, $_.Exception.Message)
        }
    }

    if ([string]::IsNullOrWhiteSpace($DatabasePath)) {
        return
    }

    $dbDir = Split-Path -Path $DatabasePath -Parent
    $dbName = [System.IO.Path]::GetFileNameWithoutExtension($DatabasePath)
    if ([string]::IsNullOrWhiteSpace($dbDir) -or [string]::IsNullOrWhiteSpace($dbName)) {
        return
    }

    $lockFile = Join-Path $dbDir ("$dbName.laccdb")
    if (Test-Path -LiteralPath $lockFile) {
        try {
            Remove-Item -LiteralPath $lockFile -Force -ErrorAction Stop
            Write-Host "Removed stale lock file: $lockFile"
        }
        catch {
            Write-Warning ("Failed to remove lock file {0}: {1}" -f $lockFile, $_.Exception.Message)
        }
    }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

$projectPath = if ([System.IO.Path]::IsPathRooted($Project)) {
    $Project
}
else {
    Join-Path $repoRoot $Project
}

if (-not (Test-Path -LiteralPath $projectPath)) {
    throw "Project file not found: $projectPath"
}

$targetPath = Join-Path $repoRoot $TargetDirectoryName
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

$attempt = 0
do {
    $suffix = if ($attempt -eq 0) { "" } else { "-$attempt" }
    $stageName = "$TargetDirectoryName-run-$timestamp$suffix"
    $backupName = "$TargetDirectoryName-backup-$timestamp$suffix"
    $stagePath = Join-Path $repoRoot $stageName
    $backupPath = Join-Path $repoRoot $backupName
    $attempt++
} while ((Test-Path -LiteralPath $stagePath) -or (Test-Path -LiteralPath $backupPath))

$publishArgs = @(
    "publish", $projectPath,
    "-c", $Configuration,
    "-r", $RuntimeIdentifier,
    "--self-contained", ($SelfContained.ToString().ToLowerInvariant()),
    "-o", $stagePath
)

Write-Host "Publishing server to staging directory: $stagePath"
& dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE."
}

$stagedExe = Join-Path $stagePath "MS.Access.MCP.Official.exe"
if (-not (Test-Path -LiteralPath $stagedExe)) {
    throw "Published executable was not found at: $stagedExe"
}

if ($StopServerProcesses) {
    $stoppedCount = Stop-LockingServerProcesses -ProcessName "MS.Access.MCP.Official" -TargetPathPrefix $targetPath -StopAllWhenPathUnavailable
    if ($stoppedCount -gt 0) {
        Start-Sleep -Milliseconds 500
    }

    $remainingServerProcesses = @(Get-Process -Name "MS.Access.MCP.Official" -ErrorAction SilentlyContinue | Where-Object {
            $remainingPath = $null
            try {
                $remainingPath = $_.Path
            }
            catch {
                # Path is inaccessible for some processes.
            }

            [string]::IsNullOrWhiteSpace($remainingPath) -or
            $remainingPath.StartsWith($targetPath, [System.StringComparison]::OrdinalIgnoreCase)
        })

    if ($remainingServerProcesses.Count -gt 0) {
        $remainingIds = ($remainingServerProcesses | Select-Object -ExpandProperty Id) -join ", "
        Write-Warning "Potential locking MS.Access.MCP.Official process(es) still running after stop attempt: $remainingIds"
        Write-Warning "If promotion fails with access denied, close those processes from an elevated shell and rerun."
    }
}
else {
    Write-Host "StopServerProcesses=false, skipping running-process shutdown."
}

$backupCreated = $false
$promotionComplete = $false
$promotionAttempt = 0

while (-not $promotionComplete -and $promotionAttempt -lt 2) {
    $promotionAttempt++

    try {
        if (Test-Path -LiteralPath $targetPath) {
            Write-Host "Backing up current target to: $backupPath"
            Rename-Item -LiteralPath $targetPath -NewName $backupName
            $backupCreated = $true
        }

        Write-Host "Promoting staging build to: $targetPath"
        Rename-Item -LiteralPath $stagePath -NewName $TargetDirectoryName
        $promotionComplete = $true
    }
    catch {
        $isAccessDenied = $_.Exception.Message -like "*Access to the path*is denied*"
        $canRetry = $isAccessDenied -and $StopServerProcesses -and $promotionAttempt -lt 2

        if ($canRetry) {
            Write-Warning "Promotion hit access denied; attempting one forced process-stop retry."
            $retryStopped = Stop-LockingServerProcesses -ProcessName "MS.Access.MCP.Official" -TargetPathPrefix $targetPath -StopAllWhenPathUnavailable
            if ($retryStopped -gt 0) {
                Start-Sleep -Milliseconds 750
            }

            continue
        }

        Write-Warning ("Promotion failed: {0}" -f $_.Exception.Message)
        if ($isAccessDenied) {
            Write-Warning "Target directory appears locked. Ensure all MS.Access.MCP.Official processes are stopped (possibly from an elevated shell), then rerun."
            Write-Warning "Process check command: Get-Process MS.Access.MCP.Official -ErrorAction SilentlyContinue"
        }

        if ($backupCreated -and -not (Test-Path -LiteralPath $targetPath) -and (Test-Path -LiteralPath $backupPath)) {
            Write-Warning "Attempting rollback to previous target directory..."
            Rename-Item -LiteralPath $backupPath -NewName $TargetDirectoryName
        }

        if (Test-Path -LiteralPath $stagePath) {
            Write-Warning "Staging directory retained for inspection: $stagePath"
        }

        throw
    }
}

if (-not $promotionComplete) {
    throw "Promotion did not complete after retry."
}

$promotedExe = Join-Path $targetPath "MS.Access.MCP.Official.exe"
if (-not (Test-Path -LiteralPath $promotedExe)) {
    throw "Promoted executable not found after promotion: $promotedExe"
}

if (-not $SkipSmokeTest) {
    Write-Host "Running smoke test (MCP initialize request)..."
    $initializeRequest = '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"release-smoke","version":"1.0"}}}'
    $nativePrefVar = Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue
    if ($nativePrefVar) {
        $previousNativePreference = $nativePrefVar.Value
        $PSNativeCommandUseErrorActionPreference = $false
    }

    try {
        $smokeOutput = @($initializeRequest | & $promotedExe 2>&1)
    }
    finally {
        if ($nativePrefVar) {
            $PSNativeCommandUseErrorActionPreference = $previousNativePreference
        }
    }

    if ($LASTEXITCODE -ne 0) {
        throw "Smoke test failed: process exited with code $LASTEXITCODE."
    }

    $initializeResponse = $null
    foreach ($line in $smokeOutput) {
        if ([string]::IsNullOrWhiteSpace([string]$line)) {
            continue
        }

        try {
            $parsed = ([string]$line) | ConvertFrom-Json
            if ($parsed.id -eq 1) {
                $initializeResponse = $parsed
                break
            }
        }
        catch {
            # Ignore non-JSON lines and continue searching for the initialize response.
        }
    }

    if ($null -eq $initializeResponse) {
        $preview = ($smokeOutput | Select-Object -First 5) -join [Environment]::NewLine
        throw "Smoke test failed: initialize response not found. Output preview: $preview"
    }

    $hasErrorProperty = @($initializeResponse.PSObject.Properties.Name) -contains "error"
    if ($hasErrorProperty -and $null -ne $initializeResponse.error) {
        $errorJson = $initializeResponse.error | ConvertTo-Json -Depth 10 -Compress
        throw "Smoke test failed: initialize response returned an error: $errorJson"
    }

    Write-Host "Smoke test passed."
}
else {
    Write-Host "Smoke test skipped by request."
}

if ($RunRegression) {
    $regressionScript = Join-Path $repoRoot "tests\full_toolset_regression.ps1"
    if (-not (Test-Path -LiteralPath $regressionScript)) {
        throw "Regression script not found: $regressionScript"
    }

    $regressionParameters = @{
        ServerExe = $promotedExe
    }
    $resolvedRegressionDatabase = $null
    if (-not [string]::IsNullOrWhiteSpace($RegressionDatabasePath)) {
        $resolvedDatabasePath = if ([System.IO.Path]::IsPathRooted($RegressionDatabasePath)) {
            $RegressionDatabasePath
        }
        else {
            Join-Path $repoRoot $RegressionDatabasePath
        }

        $regressionParameters["DatabasePath"] = $resolvedDatabasePath
        $resolvedRegressionDatabase = $resolvedDatabasePath
    }

    Ensure-RegressionEnvironmentClean -DatabasePath $resolvedRegressionDatabase
    Write-Host "Running full regression script..."
    try {
        & $regressionScript @regressionParameters
        if ($LASTEXITCODE -ne 0) {
            throw "Regression failed with exit code $LASTEXITCODE."
        }

        Write-Host "Regression passed."
    }
    catch {
        Write-Warning ("Regression phase failed: {0}" -f $_.Exception.Message)

        $canRollbackToBackup = $promotionComplete -and $backupCreated
        if ($canRollbackToBackup -and (Test-Path -LiteralPath $backupPath)) {
            $rollbackSucceeded = Restore-PreviousTargetAfterRegressionFailure -RepoRoot $repoRoot -TargetPath $targetPath -TargetDirectoryName $TargetDirectoryName -BackupPath $backupPath
            if ($rollbackSucceeded) {
                Write-Warning "Rollback to previous target completed."
            }
            else {
                Write-Warning "Rollback attempt did not fully complete; review warnings above."
            }
        }
        elseif ($canRollbackToBackup) {
            Write-Warning "Rollback skipped because backup directory is unavailable: $backupPath"
        }
        else {
            Write-Warning "Rollback skipped because no backup target was created during promotion."
        }

        throw
    }
    finally {
        Ensure-RegressionEnvironmentClean -DatabasePath $resolvedRegressionDatabase
    }
}

$validationManifestPath = Join-Path $targetPath $validationManifestName
if (-not $SkipSmokeTest) {
    $validatedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
    $gitCommit = $null
    try {
        $gitCommit = (& git -C $repoRoot rev-parse HEAD 2>$null)
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($gitCommit)) {
            $gitCommit = $gitCommit.Trim()
        }
        else {
            $gitCommit = $null
        }
    }
    catch {
        $gitCommit = $null
    }

    $validationManifest = [ordered]@{
        validated_at_utc = $validatedAtUtc
        server_exe = $promotedExe
        configuration = $Configuration
        runtime_identifier = $RuntimeIdentifier
        self_contained = $SelfContained
        smoke_test_passed = $true
        regression_run = [bool]$RunRegression
        regression_passed = [bool]$RunRegression
        git_commit = $gitCommit
    }

    $validationManifestJson = $validationManifest | ConvertTo-Json -Depth 10
    Set-Content -LiteralPath $validationManifestPath -Value $validationManifestJson -Encoding utf8

    $writtenManifest = Get-Content -LiteralPath $validationManifestPath -Raw | ConvertFrom-Json -ErrorAction Stop
    foreach ($requiredField in @("git_commit", "regression_run", "regression_passed")) {
        if ($null -eq $writtenManifest.PSObject.Properties[$requiredField]) {
            throw "Validation manifest verification failed: missing '$requiredField' field."
        }
    }

    $writtenRegressionRun = $writtenManifest.PSObject.Properties["regression_run"].Value
    if (($writtenRegressionRun -isnot [bool]) -or ([bool]$writtenRegressionRun -ne [bool]$RunRegression)) {
        throw "Validation manifest verification failed: regression_run does not match -RunRegression."
    }

    $writtenRegressionPassed = $writtenManifest.PSObject.Properties["regression_passed"].Value
    if (($writtenRegressionPassed -isnot [bool]) -or ([bool]$writtenRegressionPassed -ne [bool]$RunRegression)) {
        throw "Validation manifest verification failed: regression_passed does not match -RunRegression."
    }

    Write-Host "Validation manifest written: $validationManifestPath"
    if (-not $RunRegression) {
        Write-Warning "Validation manifest is smoke-backed only (regression_run=false). Strict repair mode (-RequireRegressionBackedManifest) will reject this binary unless explicitly overridden."
    }
}
else {
    if (Test-Path -LiteralPath $validationManifestPath) {
        Remove-Item -LiteralPath $validationManifestPath -Force
    }
    Write-Warning "Smoke test was skipped; validation manifest was not written."
}

if ($BackupRetentionCount -gt 0) {
    $backupPrefix = "$TargetDirectoryName-backup-"
    $backups = Get-ChildItem -Path $repoRoot -Directory |
        Where-Object { $_.Name.StartsWith($backupPrefix, [System.StringComparison]::OrdinalIgnoreCase) } |
        Sort-Object LastWriteTime -Descending

    $backupsToRemove = @($backups | Select-Object -Skip $BackupRetentionCount)
    foreach ($backupDir in $backupsToRemove) {
        Write-Host "Pruning old backup directory: $($backupDir.FullName)"
        Remove-Item -LiteralPath $backupDir.FullName -Recurse -Force
    }
}
else {
    Write-Host "Backup retention pruning disabled (BackupRetentionCount=0)."
}

Write-Host "Release complete."
Write-Host "Active target: $targetPath"
Write-Host "Server executable: $promotedExe"
if (Test-Path -LiteralPath $validationManifestPath) {
    Write-Host "Validation manifest: $validationManifestPath"
}
if (Test-Path -LiteralPath $backupPath) {
    Write-Host "Backup created: $backupPath"
}
