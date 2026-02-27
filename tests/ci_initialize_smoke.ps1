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
    } | ConvertTo-Json -Depth 20 -Compress),
    (@{
        jsonrpc = "2.0"
        id = 3
        method = "resources/list"
        params = @{}
    } | ConvertTo-Json -Depth 20 -Compress),
    (@{
        jsonrpc = "2.0"
        id = 4
        method = "prompts/list"
        params = @{}
    } | ConvertTo-Json -Depth 20 -Compress)
)

Write-CiMarker "SEND initialize+tools/list+resources/list+prompts/list"

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

# Verify capabilities include resources, prompts, and logging
$caps = $initializeResponse.result.capabilities
if (-not $caps.resources) {
    Fail-Smoke -Reason "initialize missing resources capability" -ExitCode 10
}
if (-not $caps.prompts) {
    Fail-Smoke -Reason "initialize missing prompts capability" -ExitCode 11
}
if (-not $caps.logging) {
    Fail-Smoke -Reason "initialize missing logging capability" -ExitCode 12
}

Write-CiMarker "CHECK initialize PASS (capabilities: tools, resources, prompts, logging)"

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

# Verify resources/list
if (-not $responsesById.ContainsKey(3)) {
    Fail-Smoke -Reason "missing resources/list response" -ExitCode 13
}

$resourcesListResponse = $responsesById[3]
if ($resourcesListResponse.error) {
    $resErrorJson = $resourcesListResponse.error | ConvertTo-Json -Depth 20 -Compress
    Fail-Smoke -Reason ("resources/list error " + $resErrorJson) -ExitCode 14
}

$resources = @()
if ($resourcesListResponse.result -and $resourcesListResponse.result.resources) {
    $resources = @($resourcesListResponse.result.resources)
}

if ($resources.Count -ne 10) {
    Fail-Smoke -Reason ("resources/list expected 10 resources, got " + $resources.Count) -ExitCode 15
}

Write-CiMarker ("CHECK resources/list PASS resources=" + $resources.Count)

# Verify prompts/list
if (-not $responsesById.ContainsKey(4)) {
    Fail-Smoke -Reason "missing prompts/list response" -ExitCode 16
}

$promptsListResponse = $responsesById[4]
if ($promptsListResponse.error) {
    $promptsErrorJson = $promptsListResponse.error | ConvertTo-Json -Depth 20 -Compress
    Fail-Smoke -Reason ("prompts/list error " + $promptsErrorJson) -ExitCode 17
}

$prompts = @()
if ($promptsListResponse.result -and $promptsListResponse.result.prompts) {
    $prompts = @($promptsListResponse.result.prompts)
}

if ($prompts.Count -ne 6) {
    Fail-Smoke -Reason ("prompts/list expected 6 prompts, got " + $prompts.Count) -ExitCode 18
}

Write-CiMarker ("CHECK prompts/list PASS prompts=" + $prompts.Count)

if ($serverExitCode -ne 0) {
    Fail-Smoke -Reason ("server exited non-zero: " + $serverExitCode) -ExitCode 9
}

Write-CiMarker "PASS"
exit 0
