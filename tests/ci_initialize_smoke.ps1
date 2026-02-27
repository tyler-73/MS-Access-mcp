param(
    [string]$ServerExe = "$PSScriptRoot\..\mcp-server-official-x64\MS.Access.MCP.Official.exe"
)

$ErrorActionPreference = "Stop"

# Resolve $ServerExe when $PSScriptRoot was empty (MSYS bash / git-bash invocations)
if (-not (Test-Path $ServerExe -ErrorAction SilentlyContinue)) {
    $fallbackRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $fallbackExe  = Join-Path $fallbackRoot "..\mcp-server-official-x64\MS.Access.MCP.Official.exe"
    if (Test-Path $fallbackExe) { $ServerExe = $fallbackExe }
}

function Write-CiMarker {
    param([string]$Message)
    Write-Host "[CI_SMOKE] $Message"
}

function Fail-Smoke {
    param(
        [string]$Reason,
        [int]$ExitCode = 1
    )

    Write-CiMarker "FAIL $Reason"
    exit $ExitCode
}

Write-CiMarker "START"
Write-CiMarker "SERVER_EXE $ServerExe"

if (-not (Test-Path -LiteralPath $ServerExe)) {
    Fail-Smoke -Reason "server executable not found" -ExitCode 2
}

$requests = @(
    (@{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = "ci-initialize-smoke"
                version = "1.0"
            }
        }
    } | ConvertTo-Json -Depth 20 -Compress),
    (@{
        jsonrpc = "2.0"
        method = "notifications/initialized"
        params = @{}
    } | ConvertTo-Json -Depth 20 -Compress),
    (@{
        jsonrpc = "2.0"
        id = 2
        method = "tools/list"
        params = @{}
    } | ConvertTo-Json -Depth 20 -Compress)
)

Write-CiMarker "SEND initialize+tools/list"

try {
    $rawLines = @((($requests -join "`n") | & $ServerExe 2>&1))
    $serverExitCode = [int]$LASTEXITCODE
}
catch {
    Fail-Smoke -Reason ("process execution error: " + $_.Exception.Message) -ExitCode 3
}

Write-CiMarker "SERVER_EXIT_CODE $serverExitCode"

$responsesById = @{}
$nonJsonLineCount = 0

foreach ($raw in $rawLines) {
    $line = [string]$raw
    if ([string]::IsNullOrWhiteSpace($line)) {
        continue
    }

    try {
        $parsed = $line | ConvertFrom-Json
        if ($null -ne $parsed.id) {
            $responsesById[[int]$parsed.id] = $parsed
        }
    }
    catch {
        $nonJsonLineCount++
    }
}

if ($nonJsonLineCount -gt 0) {
    Write-CiMarker "NON_JSON_LINES $nonJsonLineCount"
}

if (-not $responsesById.ContainsKey(1)) {
    Fail-Smoke -Reason "missing initialize response" -ExitCode 4
}

$initializeResponse = $responsesById[1]
if ($initializeResponse.error) {
    $initErrorJson = $initializeResponse.error | ConvertTo-Json -Depth 20 -Compress
    Fail-Smoke -Reason ("initialize error " + $initErrorJson) -ExitCode 5
}

Write-CiMarker "CHECK initialize PASS"

if (-not $responsesById.ContainsKey(2)) {
    Fail-Smoke -Reason "missing tools/list response" -ExitCode 6
}

$toolsListResponse = $responsesById[2]
if ($toolsListResponse.error) {
    $toolsErrorJson = $toolsListResponse.error | ConvertTo-Json -Depth 20 -Compress
    Fail-Smoke -Reason ("tools/list error " + $toolsErrorJson) -ExitCode 7
}

$tools = @()
if ($toolsListResponse.result -and $toolsListResponse.result.tools) {
    $tools = @($toolsListResponse.result.tools)
}

if ($tools.Count -le 0) {
    Fail-Smoke -Reason "tools/list returned zero tools" -ExitCode 8
}

Write-CiMarker ("CHECK tools/list PASS tools=" + $tools.Count)

if ($serverExitCode -ne 0) {
    Fail-Smoke -Reason ("server exited non-zero: " + $serverExitCode) -ExitCode 9
}

Write-CiMarker "PASS"
exit 0
