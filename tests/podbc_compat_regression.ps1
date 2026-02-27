param(
    [Alias("ServerExePath")]
    [string]$ServerExe = "$PSScriptRoot\..\mcp-server-official-x64\MS.Access.MCP.Official.exe",
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "$env:USERPROFILE\Documents\MyDatabase.accdb" })
)

$ErrorActionPreference = "Stop"

# Resolve $ServerExe when $PSScriptRoot was empty (MSYS bash / git-bash invocations)
if (-not (Test-Path $ServerExe -ErrorAction SilentlyContinue)) {
    $fallbackRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $fallbackExe  = Join-Path $fallbackRoot "..\mcp-server-official-x64\MS.Access.MCP.Official.exe"
    if (Test-Path $fallbackExe) { $ServerExe = $fallbackExe }
}

function Write-StepMarker {
    param(
        [string]$Step,
        [string]$State,
        [string]$Detail = $null
    )

    if ([string]::IsNullOrWhiteSpace($Detail)) {
        Write-Host ("PODBC_COMPAT_STEP={0} {1}" -f $Step, $State)
        return
    }

    Write-Host ("PODBC_COMPAT_STEP={0} {1} {2}" -f $Step, $State, $Detail)
}

function Fail-Test {
    param(
        [string]$Step,
        [string]$Reason,
        [int]$ExitCode = 1
    )

    Write-StepMarker -Step $Step -State "FAIL"
    Write-Host ("PODBC_COMPAT_ERROR={0}" -f $Reason)
    exit $ExitCode
}

function Decode-McpResult {
    param([object]$Response)

    if ($null -eq $Response) {
        return $null
    }

    if ($Response.result -and $Response.result.structuredContent) {
        return $Response.result.structuredContent
    }

    if ($Response.result -and $Response.result.content) {
        $text = $Response.result.content[0].text
        try {
            return $text | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            return $text
        }
    }

    return $Response.result
}

function Get-SearchText {
    param([object]$Response)

    $decoded = Decode-McpResult -Response $Response
    if ($null -eq $decoded) {
        return ""
    }

    if ($decoded -is [string]) {
        return [string]$decoded
    }

    try {
        return ($decoded | ConvertTo-Json -Depth 100 -Compress)
    }
    catch {
        return [string]$decoded
    }
}

function New-InitializeRequest {
    param(
        [int]$Id = 1,
        [string]$ClientName = "podbc-compat-regression",
        [string]$ClientVersion = "1.0"
    )

    return @{
        jsonrpc = "2.0"
        id = $Id
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = $ClientName
                version = $ClientVersion
            }
        }
    }
}

function New-ToolsListRequest {
    param([int]$Id = 2)

    return @{
        jsonrpc = "2.0"
        id = $Id
        method = "tools/list"
        params = @{}
    }
}

function New-ToolCallRequest {
    param(
        [int]$Id,
        [string]$Name,
        [hashtable]$Arguments = @{}
    )

    return @{
        jsonrpc = "2.0"
        id = $Id
        method = "tools/call"
        params = @{
            name = $Name
            arguments = $Arguments
        }
    }
}

function Invoke-McpRequests {
    param(
        [string]$ExePath,
        [object[]]$Requests
    )

    $jsonLines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($request in $Requests) {
        $jsonLines.Add(($request | ConvertTo-Json -Depth 60 -Compress))
    }

    $rawLines = @()
    try {
        $rawLines = @((($jsonLines -join "`n") | & $ExePath 2>&1))
    }
    catch {
        Fail-Test -Step "SERVER_EXECUTE" -Reason ("process execution error: " + $_.Exception.Message) -ExitCode 90
    }

    $exitCode = [int]$LASTEXITCODE
    $responses = @{}
    $nonJsonLines = 0

    foreach ($raw in $rawLines) {
        $line = [string]$raw
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $parsed = $line | ConvertFrom-Json -ErrorAction Stop
            if ($null -ne $parsed.id) {
                $responses[[int]$parsed.id] = $parsed
            }
        }
        catch {
            $nonJsonLines++
        }
    }

    return [PSCustomObject]@{
        Responses = $responses
        ExitCode = $exitCode
        NonJsonLines = $nonJsonLines
    }
}

function Get-ResponseOrFail {
    param(
        [hashtable]$Responses,
        [int]$Id,
        [string]$Step
    )

    if (-not $Responses.ContainsKey($Id)) {
        Fail-Test -Step $Step -Reason ("missing response id " + $Id) -ExitCode 91
    }

    $response = $Responses[$Id]
    if ($response.error) {
        $errorJson = $response.error | ConvertTo-Json -Depth 60 -Compress
        Fail-Test -Step $Step -Reason ("tool returned error: " + $errorJson) -ExitCode 92
    }

    return $response
}

function Assert-TextContains {
    param(
        [string]$Step,
        [string]$Text,
        [string]$Expected
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        Fail-Test -Step $Step -Reason "empty tool result text" -ExitCode 93
    }

    if ($Text.IndexOf($Expected, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
        Fail-Test -Step $Step -Reason ("expected token not found: " + $Expected) -ExitCode 94
    }
}

Write-StepMarker -Step "PRECHECK" -State "BEGIN"

if (-not (Test-Path -LiteralPath $ServerExe)) {
    Fail-Test -Step "PRECHECK" -Reason ("server executable not found: " + $ServerExe) -ExitCode 2
}

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    Fail-Test -Step "PRECHECK" -Reason ("database file not found: " + $DatabasePath) -ExitCode 3
}

Write-StepMarker -Step "PRECHECK" -State "PASS"

Write-StepMarker -Step "TOOLS_LIST" -State "BEGIN"
$toolsListRun = Invoke-McpRequests -ExePath $ServerExe -Requests @(
    (New-InitializeRequest -Id 1 -ClientName "podbc-compat-tools-list"),
    (New-ToolsListRequest -Id 2)
)

if ($toolsListRun.ExitCode -ne 0) {
    Fail-Test -Step "TOOLS_LIST" -Reason ("server exited non-zero: " + $toolsListRun.ExitCode) -ExitCode 4
}

$null = Get-ResponseOrFail -Responses $toolsListRun.Responses -Id 1 -Step "TOOLS_LIST_INITIALIZE"
$toolsListResponse = Get-ResponseOrFail -Responses $toolsListRun.Responses -Id 2 -Step "TOOLS_LIST"

$allTools = @()
if ($toolsListResponse.result -and $toolsListResponse.result.tools) {
    $allTools = @($toolsListResponse.result.tools)
}

if ($allTools.Count -le 0) {
    Fail-Test -Step "TOOLS_LIST" -Reason "tools/list returned zero tools" -ExitCode 5
}

$toolNameSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($tool in $allTools) {
    if ($tool.name) {
        $null = $toolNameSet.Add([string]$tool.name)
    }
}

$requiredPodbcTools = @(
    "podbc_get_schemas",
    "podbc_get_tables",
    "podbc_describe_table",
    "podbc_filter_table_names",
    "podbc_execute_query",
    "podbc_execute_query_md",
    "podbc_query_database"
)

$missingTools = @()
foreach ($requiredTool in $requiredPodbcTools) {
    if (-not $toolNameSet.Contains($requiredTool)) {
        $missingTools += $requiredTool
    }
}

if ($missingTools.Count -gt 0) {
    Fail-Test -Step "TOOLS_LIST" -Reason ("missing podbc tools: " + ($missingTools -join ", ")) -ExitCode 6
}

Write-StepMarker -Step "TOOLS_LIST" -State "PASS" -Detail ("tool_count=" + $allTools.Count)

$suffix = [System.Guid]::NewGuid().ToString("N").Substring(0, 8)
$tempTableName = "MCP_PODBC_COMPAT_{0}" -f $suffix
$rowValue = "podbc_row_{0}" -f $suffix
$selectQuery = "SELECT [CompatId], [CompatText] FROM [{0}] WHERE [CompatId] = 1" -f $tempTableName

$flowRequests = @(
    (New-InitializeRequest -Id 1 -ClientName "podbc-compat-flow"),
    (New-ToolCallRequest -Id 100 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }),
    (New-ToolCallRequest -Id 110 -Name "execute_sql" -Arguments @{ sql = ("CREATE TABLE [{0}] ([CompatId] INTEGER, [CompatText] TEXT(255))" -f $tempTableName) }),
    (New-ToolCallRequest -Id 120 -Name "execute_sql" -Arguments @{ sql = ("INSERT INTO [{0}] ([CompatId], [CompatText]) VALUES (1, '{1}')" -f $tempTableName, $rowValue) }),
    (New-ToolCallRequest -Id 130 -Name "podbc_get_schemas" -Arguments @{}),
    (New-ToolCallRequest -Id 140 -Name "podbc_get_tables" -Arguments @{}),
    (New-ToolCallRequest -Id 150 -Name "podbc_describe_table" -Arguments @{ table = $tempTableName }),
    (New-ToolCallRequest -Id 160 -Name "podbc_filter_table_names" -Arguments @{ q = $tempTableName }),
    (New-ToolCallRequest -Id 170 -Name "podbc_execute_query" -Arguments @{ query = $selectQuery }),
    (New-ToolCallRequest -Id 180 -Name "podbc_execute_query_md" -Arguments @{ query = $selectQuery }),
    (New-ToolCallRequest -Id 190 -Name "podbc_query_database" -Arguments @{ query = $selectQuery }),
    (New-ToolCallRequest -Id 900 -Name "execute_sql" -Arguments @{ sql = ("DROP TABLE [{0}]" -f $tempTableName) }),
    (New-ToolCallRequest -Id 910 -Name "disconnect_access" -Arguments @{}),
    (New-ToolCallRequest -Id 920 -Name "close_access" -Arguments @{})
)

Write-StepMarker -Step "FLOW" -State "BEGIN"
$flowRun = Invoke-McpRequests -ExePath $ServerExe -Requests $flowRequests

if ($flowRun.ExitCode -ne 0) {
    Write-StepMarker -Step "FLOW" -State "WARN" -Detail ("non_zero_exit=" + $flowRun.ExitCode)
}

$null = Get-ResponseOrFail -Responses $flowRun.Responses -Id 1 -Step "FLOW_INITIALIZE"
Write-StepMarker -Step "FLOW" -State "PASS" -Detail ("non_json_lines=" + $flowRun.NonJsonLines)

Write-StepMarker -Step "CONNECT_ACCESS" -State "BEGIN"
$null = Get-ResponseOrFail -Responses $flowRun.Responses -Id 100 -Step "CONNECT_ACCESS"
Write-StepMarker -Step "CONNECT_ACCESS" -State "PASS"

Write-StepMarker -Step "CREATE_TEMP_TABLE" -State "BEGIN"
$null = Get-ResponseOrFail -Responses $flowRun.Responses -Id 110 -Step "CREATE_TEMP_TABLE"
Write-StepMarker -Step "CREATE_TEMP_TABLE" -State "PASS"

Write-StepMarker -Step "INSERT_ROW" -State "BEGIN"
$null = Get-ResponseOrFail -Responses $flowRun.Responses -Id 120 -Step "INSERT_ROW"
Write-StepMarker -Step "INSERT_ROW" -State "PASS"

Write-StepMarker -Step "PODBC_GET_SCHEMAS" -State "BEGIN"
$schemasResponse = Get-ResponseOrFail -Responses $flowRun.Responses -Id 130 -Step "PODBC_GET_SCHEMAS"
$schemasText = Get-SearchText -Response $schemasResponse
if ([string]::IsNullOrWhiteSpace($schemasText)) {
    Fail-Test -Step "PODBC_GET_SCHEMAS" -Reason "returned no schema data" -ExitCode 8
}
Write-StepMarker -Step "PODBC_GET_SCHEMAS" -State "PASS"

Write-StepMarker -Step "PODBC_GET_TABLES" -State "BEGIN"
$tablesResponse = Get-ResponseOrFail -Responses $flowRun.Responses -Id 140 -Step "PODBC_GET_TABLES"
$tablesText = Get-SearchText -Response $tablesResponse
Assert-TextContains -Step "PODBC_GET_TABLES" -Text $tablesText -Expected $tempTableName
Write-StepMarker -Step "PODBC_GET_TABLES" -State "PASS"

Write-StepMarker -Step "PODBC_DESCRIBE_TABLE" -State "BEGIN"
$describeResponse = Get-ResponseOrFail -Responses $flowRun.Responses -Id 150 -Step "PODBC_DESCRIBE_TABLE"
$describeText = Get-SearchText -Response $describeResponse
Assert-TextContains -Step "PODBC_DESCRIBE_TABLE" -Text $describeText -Expected "CompatId"
Assert-TextContains -Step "PODBC_DESCRIBE_TABLE" -Text $describeText -Expected "CompatText"
Write-StepMarker -Step "PODBC_DESCRIBE_TABLE" -State "PASS"

Write-StepMarker -Step "PODBC_FILTER_TABLE_NAMES" -State "BEGIN"
$filterResponse = Get-ResponseOrFail -Responses $flowRun.Responses -Id 160 -Step "PODBC_FILTER_TABLE_NAMES"
$filterText = Get-SearchText -Response $filterResponse
Assert-TextContains -Step "PODBC_FILTER_TABLE_NAMES" -Text $filterText -Expected $tempTableName
Write-StepMarker -Step "PODBC_FILTER_TABLE_NAMES" -State "PASS"

Write-StepMarker -Step "PODBC_EXECUTE_QUERY" -State "BEGIN"
$executeQueryResponse = Get-ResponseOrFail -Responses $flowRun.Responses -Id 170 -Step "PODBC_EXECUTE_QUERY"
$executeQueryText = Get-SearchText -Response $executeQueryResponse
Assert-TextContains -Step "PODBC_EXECUTE_QUERY" -Text $executeQueryText -Expected $rowValue
Write-StepMarker -Step "PODBC_EXECUTE_QUERY" -State "PASS"

Write-StepMarker -Step "PODBC_EXECUTE_QUERY_MD" -State "BEGIN"
$executeQueryMdResponse = Get-ResponseOrFail -Responses $flowRun.Responses -Id 180 -Step "PODBC_EXECUTE_QUERY_MD"
$executeQueryMdText = Get-SearchText -Response $executeQueryMdResponse
Assert-TextContains -Step "PODBC_EXECUTE_QUERY_MD" -Text $executeQueryMdText -Expected $rowValue
Write-StepMarker -Step "PODBC_EXECUTE_QUERY_MD" -State "PASS"

Write-StepMarker -Step "PODBC_QUERY_DATABASE" -State "BEGIN"
$queryDatabaseResponse = Get-ResponseOrFail -Responses $flowRun.Responses -Id 190 -Step "PODBC_QUERY_DATABASE"
$queryDatabaseText = Get-SearchText -Response $queryDatabaseResponse
Assert-TextContains -Step "PODBC_QUERY_DATABASE" -Text $queryDatabaseText -Expected $rowValue
Write-StepMarker -Step "PODBC_QUERY_DATABASE" -State "PASS"

Write-StepMarker -Step "CLEANUP_DROP_TABLE" -State "BEGIN"
$null = Get-ResponseOrFail -Responses $flowRun.Responses -Id 900 -Step "CLEANUP_DROP_TABLE"
Write-StepMarker -Step "CLEANUP_DROP_TABLE" -State "PASS"

Write-StepMarker -Step "CLEANUP_DISCONNECT" -State "BEGIN"
$null = Get-ResponseOrFail -Responses $flowRun.Responses -Id 910 -Step "CLEANUP_DISCONNECT"
Write-StepMarker -Step "CLEANUP_DISCONNECT" -State "PASS"

Write-StepMarker -Step "CLEANUP_CLOSE" -State "BEGIN"
$null = Get-ResponseOrFail -Responses $flowRun.Responses -Id 920 -Step "CLEANUP_CLOSE"
Write-StepMarker -Step "CLEANUP_CLOSE" -State "PASS"

Write-Host "PODBC_COMPAT_PASS=1"
exit 0
