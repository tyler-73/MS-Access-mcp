# Smoke test for _dialog_watcher.ps1 — verifies P/Invoke loads, screenshot works,
# watcher starts/stops, and timeout mechanism functions correctly.

$ErrorActionPreference = "Stop"

$scriptDir = $PSScriptRoot
if (-not $scriptDir) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}

. (Join-Path $scriptDir "_dialog_watcher.ps1")

$passed = 0
$failed = 0
$total = 0

function Assert-True {
    param([string]$Name, [bool]$Value)
    $script:total++
    if ($Value) {
        $script:passed++
        Write-Host "  PASS: $Name"
    }
    else {
        $script:failed++
        Write-Host "  FAIL: $Name"
    }
}

function Assert-NotNull {
    param([string]$Name, [object]$Value)
    $script:total++
    if ($null -ne $Value) {
        $script:passed++
        Write-Host "  PASS: $Name"
    }
    else {
        $script:failed++
        Write-Host "  FAIL: $Name (was null)"
    }
}

# ── Test 1: DialogDetector type loaded ────────────────────────────────────────
Write-Host ""
Write-Host "=== Test 1: P/Invoke type loads ==="
$typeLoaded = $false
try {
    $type = [DialogDetector]
    $typeLoaded = ($null -ne $type)
}
catch {
    $typeLoaded = $false
}
Assert-True "DialogDetector type exists" $typeLoaded

# ── Test 2: FindDialogsForProcess runs without error ──────────────────────────
Write-Host ""
Write-Host "=== Test 2: FindDialogsForProcess on PID 0 (no crash) ==="
$findOk = $false
try {
    $dialogs = [DialogDetector]::FindDialogsForProcess([uint32]0)
    $findOk = ($null -ne $dialogs)
}
catch {
    Write-Host ("  ERROR: {0}" -f $_.Exception.Message)
    $findOk = $false
}
Assert-True "FindDialogsForProcess(0) returned without error" $findOk
Assert-True "FindDialogsForProcess(0) returned empty list" ($dialogs.Count -eq 0)

# ── Test 3: FindDialogsForProcess on explorer.exe (should find no dialogs) ────
Write-Host ""
Write-Host "=== Test 3: FindDialogsForProcess on explorer.exe ==="
$explorerPid = $null
try {
    $explorerProc = Get-Process -Name explorer -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($explorerProc) {
        $explorerPid = [uint32]$explorerProc.Id
    }
}
catch {}

if ($null -ne $explorerPid) {
    $explorerDialogs = [DialogDetector]::FindDialogsForProcess($explorerPid)
    Assert-True "FindDialogsForProcess(explorer) returned list" ($null -ne $explorerDialogs)
    Write-Host ("  INFO: explorer.exe (pid={0}) has {1} dialog(s)" -f $explorerPid, $explorerDialogs.Count)
}
else {
    Write-Host "  SKIP: explorer.exe not found"
}

# ── Test 4: Screenshot capture ────────────────────────────────────────────────
Write-Host ""
Write-Host "=== Test 4: Screenshot capture ==="
$tempScreenshot = Join-Path ([System.IO.Path]::GetTempPath()) "watcher_smoke_test_screenshot.png"
$screenshotOk = $false
try {
    $screenshotOk = Invoke-ScreenshotCapture -OutputPath $tempScreenshot
}
catch {
    Write-Host ("  ERROR: {0}" -f $_.Exception.Message)
    $screenshotOk = $false
}
Assert-True "Invoke-ScreenshotCapture returned true" $screenshotOk
$fileExists = Test-Path $tempScreenshot
Assert-True "Screenshot file exists" $fileExists
if ($fileExists) {
    $fileSize = (Get-Item $tempScreenshot).Length
    Assert-True "Screenshot file is non-empty" ($fileSize -gt 0)
    Write-Host ("  INFO: screenshot size = {0:N0} bytes" -f $fileSize)
    Remove-Item $tempScreenshot -Force -ErrorAction SilentlyContinue
}

# ── Test 5: Watcher start/stop lifecycle ──────────────────────────────────────
Write-Host ""
Write-Host "=== Test 5: Watcher start/stop lifecycle ==="
$tempDiagDir = Join-Path ([System.IO.Path]::GetTempPath()) ("watcher_smoke_test_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
$watcherState = $null
try {
    $watcherState = Start-DialogWatcher -DiagnosticsPath $tempDiagDir -AutoDismiss
    Assert-NotNull "Watcher state returned" $watcherState
    Assert-NotNull "Watcher job exists" $watcherState.Job
    Assert-True "Diagnostics dir created" (Test-Path $tempDiagDir)

    # Let it poll once
    Start-Sleep -Milliseconds 800

    $jobState = $watcherState.Job.State
    Assert-True "Watcher job is running" ($jobState -eq "Running")
    Write-Host ("  INFO: watcher job state = {0}" -f $jobState)
}
catch {
    $failed++
    $total++
    Write-Host ("  FAIL: watcher start error: {0}" -f $_.Exception.Message)
}

if ($null -ne $watcherState) {
    try {
        Stop-DialogWatcher -WatcherState $watcherState
        Start-Sleep -Milliseconds 200
        $jobExists = $false
        try {
            $jobState = $watcherState.Job.State
            $jobExists = $true
        }
        catch {
            $jobExists = $false
        }
        # Job should be stopped/removed
        if ($jobExists) {
            Assert-True "Watcher job stopped" ($jobState -ne "Running")
        }
        else {
            Assert-True "Watcher job removed" $true
        }
    }
    catch {
        $failed++
        $total++
        Write-Host ("  FAIL: watcher stop error: {0}" -f $_.Exception.Message)
    }
}

# ── Test 6: Get-DialogWatcherReport on empty file ─────────────────────────────
Write-Host ""
Write-Host "=== Test 6: Get-DialogWatcherReport on empty/nonexistent JSONL ==="
$emptyEvents = Get-DialogWatcherReport -JsonlPath (Join-Path $tempDiagDir "dialog_events.jsonl")
Assert-True "Empty JSONL returns empty array" ($emptyEvents.Count -eq 0)

# Cleanup temp dir
Remove-Item $tempDiagDir -Recurse -Force -ErrorAction SilentlyContinue

# ── Test 7: Timeout mechanism with a synthetic slow process ───────────────────
Write-Host ""
Write-Host "=== Test 7: Invoke-McpBatchWithTimeout with 3s timeout ==="

# Create a .cmd wrapper that runs a PowerShell script that sleeps forever
$sleepScript = Join-Path ([System.IO.Path]::GetTempPath()) "watcher_smoke_slow_server.ps1"
$sleepCmd = Join-Path ([System.IO.Path]::GetTempPath()) "watcher_smoke_slow_server.cmd"
Set-Content -Path $sleepScript -Value 'while ($true) { Start-Sleep -Seconds 60 }' -Encoding UTF8
Set-Content -Path $sleepCmd -Value "@powershell.exe -NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File `"$sleepScript`"" -Encoding ASCII

$timeoutScreenshotDir = Join-Path ([System.IO.Path]::GetTempPath()) ("watcher_timeout_test_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $timeoutScreenshotDir -Force | Out-Null

$dummyCalls = New-Object 'System.Collections.Generic.List[object]'
$dummyCalls.Add([PSCustomObject]@{
    Id = 2
    Name = "dummy_tool"
    Arguments = @{}
})

$timeoutResult = $null
$startTime = Get-Date
try {
    $timeoutResult = Invoke-McpBatchWithTimeout `
        -ExePath $sleepCmd `
        -Calls $dummyCalls `
        -ClientName "smoke-timeout-test" `
        -TimeoutSeconds 3 `
        -SectionName "smoke-timeout" `
        -ScreenshotDir $timeoutScreenshotDir
}
catch {
    Write-Host ("  ERROR: {0}" -f $_.Exception.Message)
}
$elapsed = ((Get-Date) - $startTime).TotalSeconds

Assert-NotNull "Timeout result returned" $timeoutResult
if ($null -ne $timeoutResult) {
    Assert-True "Result has _timeout=true" ([bool]$timeoutResult._timeout)
    Assert-True "Result has _section" (-not [string]::IsNullOrWhiteSpace($timeoutResult._section))
}
Assert-True "Elapsed time ~3s (between 2s and 10s)" ($elapsed -ge 2 -and $elapsed -le 10)
Write-Host ("  INFO: elapsed = {0:F1}s" -f $elapsed)

# Check if a timeout screenshot was taken
$timeoutScreenshots = @(Get-ChildItem -Path $timeoutScreenshotDir -Filter "timeout_*.png" -ErrorAction SilentlyContinue)
Assert-True "Timeout screenshot captured" ($timeoutScreenshots.Count -gt 0)
if ($timeoutScreenshots.Count -gt 0) {
    Write-Host ("  INFO: timeout screenshot = {0}" -f $timeoutScreenshots[0].Name)
}

# Cleanup
Start-Sleep -Milliseconds 500
Remove-Item $sleepScript -Force -ErrorAction SilentlyContinue
Remove-Item $sleepCmd -Force -ErrorAction SilentlyContinue
Remove-Item $timeoutScreenshotDir -Recurse -Force -ErrorAction SilentlyContinue

# ── Test 8: Invoke-McpBatchWithTimeout with normal exit ───────────────────────
Write-Host ""
Write-Host "=== Test 8: Invoke-McpBatchWithTimeout with fast-exit process ==="

# Create a .cmd wrapper that runs a PowerShell script emitting valid JSON-RPC
$fastScript = Join-Path ([System.IO.Path]::GetTempPath()) "watcher_smoke_fast_server.ps1"
$fastCmd = Join-Path ([System.IO.Path]::GetTempPath()) "watcher_smoke_fast_server.cmd"
Set-Content -Path $fastScript -Value @'
$input | ForEach-Object {
    # just consume stdin
}
# Output valid JSON-RPC responses
Write-Output '{"jsonrpc":"2.0","id":1,"result":{"protocolVersion":"2024-11-05"}}'
Write-Output '{"jsonrpc":"2.0","id":2,"result":{"content":[{"type":"text","text":"{\"success\":true}"}]}}'
'@ -Encoding UTF8
Set-Content -Path $fastCmd -Value "@powershell.exe -NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File `"$fastScript`"" -Encoding ASCII

$fastCalls = New-Object 'System.Collections.Generic.List[object]'
$fastCalls.Add([PSCustomObject]@{
    Id = 2
    Name = "dummy_tool"
    Arguments = @{}
})

$fastResult = $null
try {
    $fastResult = Invoke-McpBatchWithTimeout `
        -ExePath $fastCmd `
        -Calls $fastCalls `
        -ClientName "smoke-fast-test" `
        -TimeoutSeconds 30 `
        -SectionName "smoke-fast"
}
catch {
    Write-Host ("  ERROR: {0}" -f $_.Exception.Message)
}

Assert-NotNull "Fast result returned" $fastResult
if ($null -ne $fastResult) {
    $hasTimeout = $false
    try { $hasTimeout = [bool]$fastResult._timeout } catch {}
    Assert-True "Fast result has no _timeout" (-not $hasTimeout)
    $hasId1 = $fastResult.ContainsKey(1)
    Assert-True "Fast result has initialize response (id=1)" $hasId1
}

Remove-Item $fastScript -Force -ErrorAction SilentlyContinue
Remove-Item $fastCmd -Force -ErrorAction SilentlyContinue

# ── Test 9: Try-AutoDismissDialog with mock dialog info ───────────────────────
Write-Host ""
Write-Host "=== Test 9: Try-AutoDismissDialog pattern matching ==="

# Test pattern matching logic (without real windows - just check return values)
$mockDialog = [PSCustomObject]@{
    Handle = [IntPtr]::Zero
    Title = "Enter Parameter Value"
    ClassName = "#32770"
    ChildTexts = @("Enter Parameter Value", "param1", "OK", "Cancel")
}
$dismissResult = Try-AutoDismissDialog -Dialog $mockDialog
Assert-True "enter_parameter_value pattern detected" ($dismissResult.Pattern -eq "enter_parameter_value")

$mockSecurityDialog = [PSCustomObject]@{
    Handle = [IntPtr]::Zero
    Title = "Microsoft Access Security Notice"
    ClassName = "#32770"
    ChildTexts = @("Security Warning", "This file contains macros", "Enable", "Disable")
}
$securityResult = Try-AutoDismissDialog -Dialog $mockSecurityDialog
Assert-True "security_warning pattern detected" ($securityResult.Pattern -eq "security_warning")
Assert-True "security_warning NOT dismissed" (-not $securityResult.Dismissed)

$mockBusyDialog = [PSCustomObject]@{
    Handle = [IntPtr]::Zero
    Title = "Server is busy"
    ClassName = "#32770"
    ChildTexts = @("The server is busy")
}
$busyResult = Try-AutoDismissDialog -Dialog $mockBusyDialog
Assert-True "server_busy_com_rpc pattern detected" ($busyResult.Pattern -eq "server_busy_com_rpc")
Assert-True "server_busy NOT dismissed" (-not $busyResult.Dismissed)

$mockSaveDialog = [PSCustomObject]@{
    Handle = [IntPtr]::Zero
    Title = "Microsoft Access"
    ClassName = "#32770"
    ChildTexts = @("Do you want to save changes to the design", "Yes", "No", "Cancel")
}
$saveResult = Try-AutoDismissDialog -Dialog $mockSaveDialog
Assert-True "save_prompt pattern detected" ($saveResult.Pattern -eq "save_prompt")

# Test the exact dialog from the screenshot: VBA "File not found"
$mockVbaFileNotFound = [PSCustomObject]@{
    Handle = [IntPtr]::Zero
    Title = "Microsoft Visual Basic for Applications"
    ClassName = "#32770"
    ChildTexts = @("File not found", "OK", "Help")
}
$vbaFnfResult = Try-AutoDismissDialog -Dialog $mockVbaFileNotFound
Assert-True "VBA 'File not found' pattern detected as not_found_path_error" ($vbaFnfResult.Pattern -eq "not_found_path_error")
Assert-True "VBA 'File not found' would be dismissed" ([bool]$vbaFnfResult.Dismissed)

$mockUnknownDialog = [PSCustomObject]@{
    Handle = [IntPtr]::Zero
    Title = "Something Unexpected"
    ClassName = "#32770"
    ChildTexts = @("Totally novel dialog text")
}
$unknownResult = Try-AutoDismissDialog -Dialog $mockUnknownDialog
Assert-True "unknown pattern for unrecognized dialog" ($unknownResult.Pattern -eq "unknown")
Assert-True "unknown NOT dismissed" (-not $unknownResult.Dismissed)

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================================="
Write-Host ("SMOKE_TEST: {0} passed, {1} failed, {2} total" -f $passed, $failed, $total)
Write-Host "================================================================="

if ($failed -gt 0) {
    exit 1
}
exit 0
