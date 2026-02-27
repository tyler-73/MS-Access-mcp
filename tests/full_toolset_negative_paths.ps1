[CmdletBinding()]
param(
    [Alias("ServerExePath")]
    [string]$ServerExe = "$PSScriptRoot\..\mcp-server-official-x64\MS.Access.MCP.Official.exe",
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "$env:USERPROFILE\Documents\MyDatabase.accdb" }),
    [switch]$NoCleanup
)

$ErrorActionPreference = "Stop"

# Resolve $ServerExe when $PSScriptRoot was empty (MSYS bash / git-bash invocations)
if (-not (Test-Path $ServerExe -ErrorAction SilentlyContinue)) {
    $fallbackRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $fallbackExe  = Join-Path $fallbackRoot "..\mcp-server-official-x64\MS.Access.MCP.Official.exe"
    if (Test-Path $fallbackExe) { $ServerExe = $fallbackExe }
}

$script:TrackedMsAccessPids = New-Object 'System.Collections.Generic.HashSet[int]'

function Resolve-NormalizedPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    try {
        return [System.IO.Path]::GetFullPath($Path).TrimEnd('\').ToLowerInvariant()
    }
    catch {
        return $Path.Trim().TrimEnd('\').ToLowerInvariant()
    }
}

function Get-ProcessIdsByName {
    param([string]$Name)

    return @((Get-Process -Name $Name -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id }))
}

function Get-ProcessMetadataById {
    param([string]$ImageName)

    $metadata = @{}
    foreach ($entry in @(Get-CimInstance -ClassName Win32_Process -Filter ("Name='{0}'" -f $ImageName) -ErrorAction SilentlyContinue)) {
        $metadata[[int]$entry.ProcessId] = [PSCustomObject]@{
            ExecutablePath = [string]$entry.ExecutablePath
            CommandLine = [string]$entry.CommandLine
        }
    }

    return $metadata
}

function Get-ProcessExecutablePath {
    param(
        [object]$Process,
        [hashtable]$MetadataById
    )

    $path = $null
    try {
        $path = [string]$Process.Path
    }
    catch {
        $path = $null
    }

    if ([string]::IsNullOrWhiteSpace($path) -and $MetadataById.ContainsKey([int]$Process.Id)) {
        $path = [string]$MetadataById[[int]$Process.Id].ExecutablePath
    }

    return $path
}

function Register-NewMsAccessPids {
    param([int[]]$BeforeIds)

    $beforeSet = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($id in @($BeforeIds)) {
        [void]$beforeSet.Add([int]$id)
    }

    foreach ($id in (Get-ProcessIdsByName -Name "MSACCESS")) {
        if (-not $beforeSet.Contains([int]$id)) {
            [void]$script:TrackedMsAccessPids.Add([int]$id)
        }
    }
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
            return $text | ConvertFrom-Json
        }
        catch {
            return $text
        }
    }

    return $Response.result
}

function Add-ToolCall {
    param(
        [System.Collections.Generic.List[object]]$Calls,
        [int]$Id,
        [string]$Name,
        [hashtable]$Arguments = @{}
    )

    $Calls.Add([PSCustomObject]@{
        Id        = $Id
        Name      = $Name
        Arguments = $Arguments
    })
}

function Invoke-McpBatch {
    param(
        [string]$ExePath,
        [System.Collections.Generic.List[object]]$Calls,
        [string]$ClientName = "full-negative-regression",
        [string]$ClientVersion = "1.0"
    )

    $jsonLines = New-Object 'System.Collections.Generic.List[string]'
    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = $ClientName
                version = $ClientVersion
            }
        }
    } | ConvertTo-Json -Depth 40 -Compress))

    foreach ($call in $Calls) {
        $jsonLines.Add((@{
            jsonrpc = "2.0"
            id = $call.Id
            method = "tools/call"
            params = @{
                name = $call.Name
                arguments = $call.Arguments
            }
        } | ConvertTo-Json -Depth 50 -Compress))
    }

    $msAccessBeforeInvoke = Get-ProcessIdsByName -Name "MSACCESS"
    $rawLines = @((($jsonLines -join "`n") | & $ExePath))
    Register-NewMsAccessPids -BeforeIds $msAccessBeforeInvoke
    $responses = @{}
    foreach ($line in $rawLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $parsed = $line | ConvertFrom-Json
            if ($null -ne $parsed.id) {
                $responses[[int]$parsed.id] = $parsed
            }
        }
        catch {
            Write-Host "WARN: Could not parse response line: $line"
        }
    }

    return $responses
}

function Get-McpToolsList {
    param([string]$ExePath)

    $jsonLines = New-Object 'System.Collections.Generic.List[string]'
    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 1
        method = "initialize"
        params = @{
            protocolVersion = "2024-11-05"
            capabilities = @{}
            clientInfo = @{
                name = "negative-regression-tools-list"
                version = "1.0"
            }
        }
    } | ConvertTo-Json -Depth 40 -Compress))

    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 2
        method = "tools/list"
        params = @{}
    } | ConvertTo-Json -Depth 40 -Compress))

    $msAccessBeforeInvoke = Get-ProcessIdsByName -Name "MSACCESS"
    $rawLines = @((($jsonLines -join "`n") | & $ExePath))
    Register-NewMsAccessPids -BeforeIds $msAccessBeforeInvoke
    foreach ($line in $rawLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        try {
            $parsed = $line | ConvertFrom-Json
            if ($parsed.id -eq 2 -and $parsed.result -and $parsed.result.tools) {
                return @($parsed.result.tools)
            }
        }
        catch {
            Write-Host "WARN: Could not parse tools/list response line: $line"
        }
    }

    return @()
}

function Resolve-ToolName {
    param(
        [System.Collections.Generic.Dictionary[string, object]]$ToolByName,
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if ($ToolByName.ContainsKey($candidate)) {
            return $candidate
        }
    }

    return $null
}

function Get-ToolPropertyNames {
    param([object]$ToolDefinition)

    if ($null -eq $ToolDefinition -or $null -eq $ToolDefinition.inputSchema -or $null -eq $ToolDefinition.inputSchema.properties) {
        return @()
    }

    return @($ToolDefinition.inputSchema.properties.PSObject.Properties.Name)
}

function Get-DatabaseLockPath {
    param([string]$DbPath)

    if ([string]::IsNullOrWhiteSpace($DbPath)) {
        return $null
    }

    $dbDir = Split-Path -Path $DbPath -Parent
    if ([string]::IsNullOrWhiteSpace($dbDir)) {
        return $null
    }

    $dbName = [System.IO.Path]::GetFileNameWithoutExtension($DbPath)
    return (Join-Path $dbDir ($dbName + ".laccdb"))
}

function Stop-StaleProcesses {
    param([string]$DbPath)

    $normalizedServerExe = Resolve-NormalizedPath -Path $ServerExe

    $serverMetadataById = Get-ProcessMetadataById -ImageName "MS.Access.MCP.Official.exe"
    foreach ($proc in @(Get-Process -Name "MS.Access.MCP.Official" -ErrorAction SilentlyContinue)) {
        $procPath = Get-ProcessExecutablePath -Process $proc -MetadataById $serverMetadataById
        if ([string]::IsNullOrWhiteSpace($procPath)) {
            continue
        }

        $normalizedProcPath = Resolve-NormalizedPath -Path $procPath
        if ($normalizedProcPath -eq $normalizedServerExe) {
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }
    }

    $msAccessMetadataById = Get-ProcessMetadataById -ImageName "MSACCESS.EXE"
    foreach ($proc in @(Get-Process -Name "MSACCESS" -ErrorAction SilentlyContinue)) {
        $procId = [int]$proc.Id
        $isTracked = $script:TrackedMsAccessPids.Contains($procId)

        $commandLine = $null
        if ($msAccessMetadataById.ContainsKey($procId)) {
            $commandLine = [string]$msAccessMetadataById[$procId].CommandLine
        }

        $isEmbedding = (-not [string]::IsNullOrWhiteSpace($commandLine)) -and ($commandLine -match '(?i)(^|\s)/embedding(\s|$)')
        $hasEmptyWindowTitle = [string]::IsNullOrWhiteSpace([string]$proc.MainWindowTitle)

        if ($isTracked -or $isEmbedding -or $hasEmptyWindowTitle) {
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        }
    }
}

function Remove-LockFile {
    param([string]$DbPath)

    $lockFile = Get-DatabaseLockPath -DbPath $DbPath
    if ([string]::IsNullOrWhiteSpace($lockFile)) {
        return
    }

    Remove-Item -Path $lockFile -ErrorAction SilentlyContinue
}

function Cleanup-AccessArtifacts {
    param([string]$DbPath)

    Stop-StaleProcesses -DbPath $DbPath
    Remove-LockFile -DbPath $DbPath
}

function Acquire-RegressionLock {
    param([string]$LockName = "ms-access-mcp-regression")

    $lockRoot = [System.IO.Path]::GetTempPath()
    if ([string]::IsNullOrWhiteSpace($lockRoot)) {
        $lockRoot = $env:TEMP
    }
    if ([string]::IsNullOrWhiteSpace($lockRoot)) {
        throw "Unable to resolve a temporary directory for regression lock file."
    }

    $lockPath = Join-Path $lockRoot ($LockName + ".lock")
    try {
        $stream = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        return [pscustomobject]@{
            Path = $lockPath
            Stream = $stream
        }
    }
    catch {
        throw ("Another regression run is already active (lock file: {0}). Wait for it to finish or remove stale lock after confirming no run is active." -f $lockPath)
    }
}

function Release-RegressionLock {
    param([object]$LockState)

    if ($null -eq $LockState) {
        return
    }

    try {
        if ($LockState.Stream) {
            $LockState.Stream.Dispose()
        }
    }
    catch {
        # Ignore lock cleanup failures.
    }
}

function Assert-FailureResponse {
    param(
        [hashtable]$Responses,
        [int]$Id,
        [string]$Name,
        [switch]$RequirePreflight
    )

    if (-not $Responses.ContainsKey($Id)) {
        throw "Missing response for $Name (id=$Id)."
    }

    $response = $Responses[$Id]
    if ($response.error) {
        return
    }

    $decoded = Decode-McpResult -Response $response
    $isFailure = $false
    $hasPreflight = $false

    if ($decoded -and $decoded.PSObject.Properties["success"]) {
        $isFailure = (-not [bool]$decoded.success)
    }
    if ($decoded -and $decoded.PSObject.Properties["preflight"] -and $decoded.preflight) {
        $hasPreflight = $true
    }

    if (-not $isFailure) {
        throw "$Name (id=$Id) unexpectedly succeeded."
    }

    if ($RequirePreflight -and (-not $hasPreflight)) {
        throw "$Name (id=$Id) failed but did not include preflight diagnostics."
    }
}

if (-not (Test-Path -LiteralPath $ServerExe)) {
    throw "Server executable not found: $ServerExe"
}

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database file not found: $DatabasePath"
}

$regressionLock = Acquire-RegressionLock
Write-Host ("Regression lock acquired: {0}" -f $regressionLock.Path)

try {
    if (-not $NoCleanup) {
        Write-Host "Pre-run cleanup: clearing stale Access/MCP processes and locks."
        Cleanup-AccessArtifacts -DbPath $DatabasePath
    }
    else {
        Write-Warning "Skipping pre-run cleanup per -NoCleanup; final cleanup will still execute."
    }
}
catch {
    Release-RegressionLock -LockState $regressionLock
    throw
}

$exitCode = 1
try {
    $toolList = Get-McpToolsList -ExePath $ServerExe
    if ($toolList.Count -eq 0) {
        throw "tools/list returned no tools; cannot execute negative-path coverage."
    }

    $toolByName = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($tool in $toolList) {
        $name = [string]$tool.name
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $toolByName[$name] = $tool
        }
    }

    $createLinkedTableToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("create_linked_table", "link_table")
    if ([string]::IsNullOrWhiteSpace($createLinkedTableToolName)) {
        throw "Linked-table create tool not found (expected create_linked_table or link_table)."
    }
    $createDatabaseToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("create_database")
    $backupDatabaseToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("backup_database")
    $compactRepairDatabaseToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("compact_repair_database")
    if ([string]::IsNullOrWhiteSpace($createDatabaseToolName) -or
        [string]::IsNullOrWhiteSpace($backupDatabaseToolName) -or
        [string]::IsNullOrWhiteSpace($compactRepairDatabaseToolName)) {
        $missingDatabaseLifecycleTools = @()
        if ([string]::IsNullOrWhiteSpace($createDatabaseToolName)) { $missingDatabaseLifecycleTools += "create_database" }
        if ([string]::IsNullOrWhiteSpace($backupDatabaseToolName)) { $missingDatabaseLifecycleTools += "backup_database" }
        if ([string]::IsNullOrWhiteSpace($compactRepairDatabaseToolName)) { $missingDatabaseLifecycleTools += "compact_repair_database" }
        throw ("Database lifecycle tools required for negative-path coverage are missing: {0}" -f ($missingDatabaseLifecycleTools -join ", "))
    }

    if (-not $toolByName.ContainsKey("connect_access")) {
        throw "connect_access tool definition missing from tools/list."
    }
    $connectAccessToolDefinition = $toolByName["connect_access"]
    $connectAccessPropertyNames = Get-ToolPropertyNames -ToolDefinition $connectAccessToolDefinition
    $secureConnectArgNames = @($connectAccessPropertyNames | Where-Object { $_ -imatch "password|pwd|secret|secure|credential|system_database_path|workgroup" })
    $secureConnectArgNames = @($secureConnectArgNames | Where-Object { [string]$_ -ine "database_path" })
    if ($secureConnectArgNames.Count -eq 0) {
        throw "connect_access secure argument coverage failed: no secure/password/system-db related input schema properties detected."
    }
    $connectSecureArgName = [string]$secureConnectArgNames[0]
    Write-Host ("connect_access_secure_arg_detected={0}" -f $connectSecureArgName)

    $hasSystemDatabasePathArg = @($connectAccessPropertyNames | Where-Object { [string]$_ -ieq "system_database_path" }).Count -gt 0
    $secureConnectPathProbeArgs = $null

    $suffix = [Guid]::NewGuid().ToString("N").Substring(0, 8)
    $invalidDatabasePath = Join-Path (Split-Path -Path $DatabasePath -Parent) ("MCP_DoesNotExist_{0}.accdb" -f $suffix)
    $invalidLinkedSourcePath = Join-Path (Split-Path -Path $DatabasePath -Parent) ("MCP_LinkSourceMissing_{0}.accdb" -f $suffix)
    $invalidCreateDatabasePath = Join-Path (Split-Path -Path $DatabasePath -Parent) ("MCP_InvalidCreate_{0}.txt" -f $suffix)
    $invalidBackupDestinationPath = Join-Path (Split-Path -Path $DatabasePath -Parent) ("MCP_BackupInvalid_{0}.accdb" -f $suffix)
    $invalidCompactDestinationPath = Join-Path (Split-Path -Path $DatabasePath -Parent) ("MCP_CompactInvalid_{0}.accdb" -f $suffix)
    $linkedTableName = "MCP_Linked_Invalid_$suffix"
    if ($hasSystemDatabasePathArg) {
        $secureConnectPathProbeArgs = @{
            database_path = $DatabasePath
            system_database_path = $invalidDatabasePath
        }
    }
    foreach ($dbPath in @($invalidBackupDestinationPath, $invalidCompactDestinationPath)) {
        Remove-Item -Path $dbPath -Force -ErrorAction SilentlyContinue
    }

    $calls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $invalidDatabasePath }
    Add-ToolCall -Calls $calls -Id 3 -Name "get_tables" -Arguments @{}
    Add-ToolCall -Calls $calls -Id 4 -Name "get_queries" -Arguments @{}
    Add-ToolCall -Calls $calls -Id 5 -Name "get_relationships" -Arguments @{}
    Add-ToolCall -Calls $calls -Id 6 -Name "execute_sql" -Arguments @{ sql = "SELECT * FROM NonExistentTable" }
    Add-ToolCall -Calls $calls -Id 7 -Name "execute_query_md" -Arguments @{ sql = "SELECT * FROM NonExistentTable" }
    Add-ToolCall -Calls $calls -Id 8 -Name "describe_table" -Arguments @{ table_name = "DefinitelyMissingTable" }
    Add-ToolCall -Calls $calls -Id 9 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    Add-ToolCall -Calls $calls -Id 10 -Name "commit_transaction" -Arguments @{}
    Add-ToolCall -Calls $calls -Id 11 -Name "rollback_transaction" -Arguments @{}
    Add-ToolCall -Calls $calls -Id 12 -Name "execute_sql" -Arguments @{ sql = "SELEKT * FRM DefinitelyMissingTable" }
    Add-ToolCall -Calls $calls -Id 13 -Name $createLinkedTableToolName -Arguments @{
        source_database_path = $invalidLinkedSourcePath
        source_table_name    = "AnyTable"
        linked_table_name    = $linkedTableName
    }
    Add-ToolCall -Calls $calls -Id 14 -Name "import_form_from_text" -Arguments @{ form_data = "{invalid-json" }
    Add-ToolCall -Calls $calls -Id 15 -Name "import_report_from_text" -Arguments @{ report_data = "{invalid-json" }
    Add-ToolCall -Calls $calls -Id 16 -Name "disconnect_access" -Arguments @{}
    Add-ToolCall -Calls $calls -Id 17 -Name "get_tables" -Arguments @{}
    Add-ToolCall -Calls $calls -Id 18 -Name $createDatabaseToolName -Arguments @{ database_path = $invalidCreateDatabasePath }
    Add-ToolCall -Calls $calls -Id 19 -Name $backupDatabaseToolName -Arguments @{
        database_path = $invalidDatabasePath
        source_database_path = $invalidDatabasePath
        backup_path = $invalidBackupDestinationPath
        backup_database_path = $invalidBackupDestinationPath
        destination_path = $invalidBackupDestinationPath
    }
    Add-ToolCall -Calls $calls -Id 20 -Name $compactRepairDatabaseToolName -Arguments @{
        database_path = $invalidDatabasePath
        source_database_path = $invalidDatabasePath
        output_database_path = $invalidCompactDestinationPath
        compacted_database_path = $invalidCompactDestinationPath
        destination_database_path = $invalidCompactDestinationPath
    }
    if ($hasSystemDatabasePathArg) {
        Add-ToolCall -Calls $calls -Id 21 -Name "connect_access" -Arguments $secureConnectPathProbeArgs
    }

    $responses = Invoke-McpBatch -ExePath $ServerExe -Calls $calls

    # Positive checkpoints used to make sure negative assertions run in the intended state.
    $connectValid = Decode-McpResult -Response $responses[9]
    if (-not ($connectValid -and $connectValid.PSObject.Properties["success"] -and [bool]$connectValid.success)) {
        throw "connect_access valid-path checkpoint failed; negative coverage cannot proceed."
    }

    $disconnectValid = Decode-McpResult -Response $responses[16]
    if (-not ($disconnectValid -and $disconnectValid.PSObject.Properties["success"] -and [bool]$disconnectValid.success)) {
        throw "disconnect_access checkpoint failed."
    }

    Assert-FailureResponse -Responses $responses -Id 2 -Name "connect_access_invalid_path"
    Assert-FailureResponse -Responses $responses -Id 3 -Name "get_tables_disconnected" -RequirePreflight
    Assert-FailureResponse -Responses $responses -Id 4 -Name "get_queries_disconnected" -RequirePreflight
    Assert-FailureResponse -Responses $responses -Id 5 -Name "get_relationships_disconnected" -RequirePreflight
    Assert-FailureResponse -Responses $responses -Id 6 -Name "execute_sql_disconnected" -RequirePreflight
    Assert-FailureResponse -Responses $responses -Id 7 -Name "execute_query_md_disconnected" -RequirePreflight
    Assert-FailureResponse -Responses $responses -Id 8 -Name "describe_table_disconnected" -RequirePreflight
    Assert-FailureResponse -Responses $responses -Id 10 -Name "commit_transaction_without_begin"
    Assert-FailureResponse -Responses $responses -Id 11 -Name "rollback_transaction_without_begin"
    Assert-FailureResponse -Responses $responses -Id 12 -Name "execute_sql_invalid_syntax"
    Assert-FailureResponse -Responses $responses -Id 13 -Name "create_linked_table_missing_source_path"
    Assert-FailureResponse -Responses $responses -Id 14 -Name "import_form_from_text_invalid_payload"
    Assert-FailureResponse -Responses $responses -Id 15 -Name "import_report_from_text_invalid_payload"
    Assert-FailureResponse -Responses $responses -Id 17 -Name "get_tables_post_disconnect" -RequirePreflight
    Assert-FailureResponse -Responses $responses -Id 18 -Name "create_database_invalid_path"
    Assert-FailureResponse -Responses $responses -Id 19 -Name "backup_database_missing_source"
    Assert-FailureResponse -Responses $responses -Id 20 -Name "compact_repair_database_missing_source"
    if ($hasSystemDatabasePathArg) {
        Assert-FailureResponse -Responses $responses -Id 21 -Name "connect_access_system_database_path_missing_file"
    }
    else {
        Write-Host "connect_access_secure_arg_negative_path: SKIP system_database_path not exposed by tool schema"
    }

    Write-Host "NEGATIVE_PATHS_PASS=1"
    $exitCode = 0
}
finally {
    Write-Host "Final cleanup: clearing stale Access/MCP processes and locks."
    Cleanup-AccessArtifacts -DbPath $DatabasePath
    Release-RegressionLock -LockState $regressionLock
}

exit $exitCode
