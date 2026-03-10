[CmdletBinding()]
param(
    [Alias("ServerExePath")]
    [string]$ServerExe = "$PSScriptRoot\..\mcp-server-official-x64\MS.Access.MCP.Official.exe",
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "$env:USERPROFILE\Documents\MyDatabase.accdb" }),
    [switch]$NoCleanup,
    [int]$BatchTimeoutSeconds = 300,
    [switch]$NoDialogWatcher
)

$ErrorActionPreference = "Stop"

# ── Dialog watcher and timeout-aware batch support ─────────────────────────────
$script:DialogWatcherAvailable = $false
$script:DialogWatcherState = $null
$script:DiagnosticsDir = $null
$script:TimeoutCount = 0
$script:TimeoutSections = @{}
$script:BatchTimeoutSeconds = $BatchTimeoutSeconds

$dialogWatcherPath = Join-Path $PSScriptRoot "_dialog_watcher.ps1"
if (-not $PSScriptRoot) {
    $dialogWatcherPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "_dialog_watcher.ps1"
}
if (Test-Path $dialogWatcherPath) {
    . $dialogWatcherPath
    $script:DialogWatcherAvailable = $true
}

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

    $msAccessBeforeInvoke = Get-ProcessIdsByName -Name "MSACCESS"
    try {
        if ($script:DialogWatcherAvailable) {
            $responses = Invoke-McpBatchWithTimeout -ExePath $ExePath -Calls $Calls `
                -ClientName $ClientName -ClientVersion $ClientVersion `
                -TimeoutSeconds $script:BatchTimeoutSeconds `
                -SectionName $ClientName `
                -ScreenshotDir $script:DiagnosticsDir
            if ($responses._timeout) {
                $script:TimeoutCount++
                $script:TimeoutSections[$ClientName] = $true
                Write-Host ("SECTION_TIMEOUT: {0} after {1}s" -f $ClientName, $script:BatchTimeoutSeconds)
                Stop-StaleProcesses -DbPath $DatabasePath
            }
            return $responses
        }

        # Legacy fallback
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

        $rawLines = @((($jsonLines -join "`n") | & $ExePath))

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
    finally {
        Register-NewMsAccessPids -BeforeIds $msAccessBeforeInvoke
    }
}

function Get-McpToolsList {
    param([string]$ExePath)

    $msAccessBeforeInvoke = Get-ProcessIdsByName -Name "MSACCESS"
    try {
        if ($script:DialogWatcherAvailable) {
            return (Get-McpToolsListWithTimeout -ExePath $ExePath `
                -ClientName "negative-regression-tools-list" -ClientVersion "1.0" `
                -TimeoutSeconds 60 `
                -ScreenshotDir $script:DiagnosticsDir)
        }

        # Legacy fallback
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

        $rawLines = @((($jsonLines -join "`n") | & $ExePath))

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
    finally {
        Register-NewMsAccessPids -BeforeIds $msAccessBeforeInvoke
    }
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

# ── Diagnostics directory and dialog watcher setup ────────────────────────────
$runTimestamp = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmss") + "Z"
$script:DiagnosticsDir = Join-Path (Join-Path $PSScriptRoot "_diagnostics") ("neg_run_" + $runTimestamp)
if (-not $PSScriptRoot) {
    $script:DiagnosticsDir = Join-Path (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "_diagnostics") ("neg_run_" + $runTimestamp)
}
if (-not (Test-Path $script:DiagnosticsDir)) {
    New-Item -ItemType Directory -Path $script:DiagnosticsDir -Force | Out-Null
}

if ($script:DialogWatcherAvailable -and (-not $NoDialogWatcher)) {
    $script:DialogWatcherState = Start-DialogWatcher -DiagnosticsPath $script:DiagnosticsDir -AutoDismiss
    Write-Host ("Dialog watcher started: diagnostics={0}" -f $script:DiagnosticsDir)
}

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
    if ($null -ne $script:DialogWatcherState) { Stop-DialogWatcher -WatcherState $script:DialogWatcherState }
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

    # ── Phase 9-12 Negative Path Tests ──
    Write-Host ""
    Write-Host "=== Phase 9-12 Negative Path Tests ==="
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np2Calls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np2Calls -Id 100 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    # Recordset tools with invalid handles
    Add-ToolCall -Calls $np2Calls -Id 101 -Name "recordset_get_string" -Arguments @{ recordset_id = "rs_invalid" }
    Add-ToolCall -Calls $np2Calls -Id 102 -Name "recordset_seek" -Arguments @{ recordset_id = "rs_invalid"; index_name = "PK"; key_values = @(1) }
    Add-ToolCall -Calls $np2Calls -Id 103 -Name "recordset_clone" -Arguments @{ recordset_id = "rs_invalid" }
    # Control tools with invalid form/report names
    Add-ToolCall -Calls $np2Calls -Id 104 -Name "create_control" -Arguments @{ form_name = "nonexistent_form_np"; control_type = "TextBox" }
    Add-ToolCall -Calls $np2Calls -Id 105 -Name "delete_control" -Arguments @{ form_name = "nonexistent_form_np"; control_name = "ctrl1" }
    Add-ToolCall -Calls $np2Calls -Id 106 -Name "create_report_control" -Arguments @{ report_name = "nonexistent_report_np"; control_type = "TextBox" }
    Add-ToolCall -Calls $np2Calls -Id 107 -Name "delete_report_control" -Arguments @{ report_name = "nonexistent_report_np"; control_name = "ctrl1" }
    Add-ToolCall -Calls $np2Calls -Id 108 -Name "control_set_zorder" -Arguments @{ object_type = "form"; object_name = "nonexistent_form_np"; control_name = "ctrl1"; position = "front" }
    Add-ToolCall -Calls $np2Calls -Id 109 -Name "get_tab_control_pages" -Arguments @{ form_name = "nonexistent_form_np"; control_name = "tabCtrl" }
    # Phase 10-11 tools with invalid args
    Add-ToolCall -Calls $np2Calls -Id 110 -Name "reset_subdatasheet_properties" -Arguments @{ table_name = "nonexistent_table_np" }
    Add-ToolCall -Calls $np2Calls -Id 111 -Name "get_object_dependencies" -Arguments @{ object_type = "table"; object_name = "nonexistent_table_np" }
    Add-ToolCall -Calls $np2Calls -Id 112 -Name "get_record_source_fields" -Arguments @{ source = "nonexistent_table_np" }
    Add-ToolCall -Calls $np2Calls -Id 113 -Name "disconnect_access" -Arguments @{}

    $np2Responses = Invoke-McpBatch -ExePath $ServerExe -Calls $np2Calls

    # Checkpoint: connect must succeed
    $np2Connect = Decode-McpResult -Response $np2Responses[100]
    if (-not ($np2Connect -and $np2Connect.PSObject.Properties["success"] -and [bool]$np2Connect.success)) {
        throw "Phase 9-12 negative paths: connect_access checkpoint failed."
    }

    Assert-FailureResponse -Responses $np2Responses -Id 101 -Name "recordset_get_string_invalid_handle"
    Assert-FailureResponse -Responses $np2Responses -Id 102 -Name "recordset_seek_invalid_handle"
    Assert-FailureResponse -Responses $np2Responses -Id 103 -Name "recordset_clone_invalid_handle"
    Assert-FailureResponse -Responses $np2Responses -Id 104 -Name "create_control_invalid_form"
    Assert-FailureResponse -Responses $np2Responses -Id 105 -Name "delete_control_invalid_form"
    Assert-FailureResponse -Responses $np2Responses -Id 106 -Name "create_report_control_invalid_report"
    Assert-FailureResponse -Responses $np2Responses -Id 107 -Name "delete_report_control_invalid_report"
    Assert-FailureResponse -Responses $np2Responses -Id 108 -Name "control_set_zorder_invalid_form"
    Assert-FailureResponse -Responses $np2Responses -Id 109 -Name "get_tab_control_pages_invalid_form"
    Assert-FailureResponse -Responses $np2Responses -Id 110 -Name "reset_subdatasheet_properties_invalid_table"
    Assert-FailureResponse -Responses $np2Responses -Id 111 -Name "get_object_dependencies_invalid_object"
    Assert-FailureResponse -Responses $np2Responses -Id 112 -Name "get_record_source_fields_invalid_source"

    $np2Disconnect = Decode-McpResult -Response $np2Responses[113]
    if (-not ($np2Disconnect -and $np2Disconnect.PSObject.Properties["success"] -and [bool]$np2Disconnect.success)) {
        Write-Host "Phase 9-12 negative paths: disconnect checkpoint failed (non-fatal)."
    }

    # ── Batch 3: Non-Existent Object Operations (IDs 200-249) ──
    Write-Host ""
    Write-Host "=== Batch 3: Non-Existent Object Operations ==="
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np3Calls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np3Calls -Id 200 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    # Delete non-existent objects (IDs 201-213)
    Add-ToolCall -Calls $np3Calls -Id 201 -Name "delete_table" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 202 -Name "delete_query" -Arguments @{ query_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 203 -Name "delete_form" -Arguments @{ form_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 204 -Name "delete_report" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 205 -Name "delete_macro" -Arguments @{ macro_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 206 -Name "delete_module" -Arguments @{ module_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 207 -Name "delete_index" -Arguments @{ table_name = "np_nonexistent"; index_name = "idx1" }
    Add-ToolCall -Calls $np3Calls -Id 208 -Name "delete_relationship" -Arguments @{ relationship_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 209 -Name "delete_linked_table" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 210 -Name "delete_navigation_group" -Arguments @{ group_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 211 -Name "delete_conditional_formatting" -Arguments @{ object_type = "form"; object_name = "np_nonexistent"; control_name = "c1"; rule_index = 0 }
    Add-ToolCall -Calls $np3Calls -Id 212 -Name "delete_data_macro" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 213 -Name "delete_import_export_spec" -Arguments @{ spec_name = "np_nonexistent" }
    # Get/describe non-existent objects (IDs 214-248)
    Add-ToolCall -Calls $np3Calls -Id 214 -Name "describe_table" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 215 -Name "get_indexes" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 216 -Name "get_form_controls" -Arguments @{ form_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 217 -Name "get_report_controls" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 218 -Name "get_control_properties" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np3Calls -Id 219 -Name "get_report_control_properties" -Arguments @{ report_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np3Calls -Id 220 -Name "get_form_properties" -Arguments @{ form_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 221 -Name "get_report_properties" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 222 -Name "get_form_runtime_state" -Arguments @{ form_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 223 -Name "get_module_info" -Arguments @{ module_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 224 -Name "get_module_declarations" -Arguments @{ module_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 225 -Name "list_procedures" -Arguments @{ module_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 226 -Name "get_procedure_code" -Arguments @{ module_name = "np_nonexistent"; procedure_name = "Sub1" }
    Add-ToolCall -Calls $np3Calls -Id 227 -Name "find_text_in_module" -Arguments @{ module_name = "np_nonexistent"; search_text = "x" }
    Add-ToolCall -Calls $np3Calls -Id 228 -Name "get_vba_code" -Arguments @{ module_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 229 -Name "get_report_grouping" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 230 -Name "get_report_sorting" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 231 -Name "get_subdatasheet_properties" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 232 -Name "get_table_validation" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 233 -Name "get_table_description" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 234 -Name "get_table_properties" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 235 -Name "get_table_custom_property" -Arguments @{ table_name = "np_nonexistent"; property_name = "x" }
    Add-ToolCall -Calls $np3Calls -Id 236 -Name "get_conditional_formatting" -Arguments @{ object_type = "form"; object_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np3Calls -Id 237 -Name "get_attachment_files" -Arguments @{ table_name = "np_nonexistent"; attachment_field = "f1"; record_id = 1 }
    Add-ToolCall -Calls $np3Calls -Id 238 -Name "get_object_metadata" -Arguments @{ object_type = "table"; object_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 239 -Name "get_object_dates" -Arguments @{ object_type = "table"; object_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 240 -Name "get_object_events" -Arguments @{ object_type = "form"; object_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 241 -Name "get_navigation_group_objects" -Arguments @{ group_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 242 -Name "get_field_properties" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1" }
    Add-ToolCall -Calls $np3Calls -Id 243 -Name "get_field_attributes" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1" }
    Add-ToolCall -Calls $np3Calls -Id 244 -Name "get_query_parameters" -Arguments @{ query_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 245 -Name "get_query_properties" -Arguments @{ query_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 246 -Name "get_table_data_macros" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 247 -Name "get_import_export_spec" -Arguments @{ spec_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 248 -Name "get_table_dependencies" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np3Calls -Id 249 -Name "disconnect_access" -Arguments @{}

    $np3Responses = Invoke-McpBatch -ExePath $ServerExe -Calls $np3Calls -ClientName "neg-batch3-nonexistent"

    $np3Connect = Decode-McpResult -Response $np3Responses[200]
    if (-not ($np3Connect -and $np3Connect.PSObject.Properties["success"] -and [bool]$np3Connect.success)) {
        throw "Batch 3: connect_access checkpoint failed."
    }

    # Tools that return success with empty results for non-existent objects (schema queries return no rows)
    $np3SkipIds = @(215, 238, 239, 248)  # get_indexes, get_object_metadata, get_object_dates, get_table_dependencies
    for ($i = 201; $i -le 248; $i++) {
        if ($np3SkipIds -contains $i) {
            Write-Host "batch3_id${i}: SKIP (returns success with empty/default results)"
            continue
        }
        Assert-FailureResponse -Responses $np3Responses -Id $i -Name "batch3_id$i"
    }

    $np3Disconnect = Decode-McpResult -Response $np3Responses[249]
    if (-not ($np3Disconnect -and $np3Disconnect.PSObject.Properties["success"] -and [bool]$np3Disconnect.success)) {
        Write-Host "Batch 3: disconnect checkpoint failed (non-fatal)."
    }

    # ── Batch 4: Property Setters on Non-Existent Objects (IDs 250-289) ──
    Write-Host ""
    Write-Host "=== Batch 4: Property Setters on Non-Existent Objects ==="
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np4Calls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np4Calls -Id 250 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    Add-ToolCall -Calls $np4Calls -Id 251 -Name "set_control_property" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1"; property_name = "Caption"; property_value = "x" }
    Add-ToolCall -Calls $np4Calls -Id 252 -Name "set_report_control_property" -Arguments @{ report_name = "np_nonexistent"; control_name = "c1"; property_name = "Caption"; property_value = "x" }
    Add-ToolCall -Calls $np4Calls -Id 253 -Name "set_form_property" -Arguments @{ form_name = "np_nonexistent"; property_name = "Caption"; property_value = "x" }
    Add-ToolCall -Calls $np4Calls -Id 254 -Name "set_report_property" -Arguments @{ report_name = "np_nonexistent"; property_name = "Caption"; property_value = "x" }
    Add-ToolCall -Calls $np4Calls -Id 255 -Name "set_form_record_source" -Arguments @{ form_name = "np_nonexistent"; record_source = "tbl" }
    Add-ToolCall -Calls $np4Calls -Id 256 -Name "set_report_record_source" -Arguments @{ report_name = "np_nonexistent"; record_source = "tbl" }
    Add-ToolCall -Calls $np4Calls -Id 257 -Name "set_field_required" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; required = $true }
    Add-ToolCall -Calls $np4Calls -Id 258 -Name "set_field_format" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; format = "General" }
    Add-ToolCall -Calls $np4Calls -Id 259 -Name "set_field_description" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; description = "test" }
    Add-ToolCall -Calls $np4Calls -Id 260 -Name "set_field_validation" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; validation_rule = ">0" }
    Add-ToolCall -Calls $np4Calls -Id 261 -Name "set_field_input_mask" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; input_mask = "000" }
    Add-ToolCall -Calls $np4Calls -Id 262 -Name "set_field_caption" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; caption = "test" }
    Add-ToolCall -Calls $np4Calls -Id 263 -Name "set_field_decimal_places" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; decimal_places = 2 }
    Add-ToolCall -Calls $np4Calls -Id 264 -Name "set_field_default" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; default_value = "0" }
    Add-ToolCall -Calls $np4Calls -Id 265 -Name "set_field_allow_zero_length" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; allow_zero_length = $true }
    Add-ToolCall -Calls $np4Calls -Id 266 -Name "set_field_append_only" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; append_only = $true }
    Add-ToolCall -Calls $np4Calls -Id 267 -Name "set_table_custom_property" -Arguments @{ table_name = "np_nonexistent"; property_name = "x"; property_value = "y" }
    Add-ToolCall -Calls $np4Calls -Id 268 -Name "set_table_description" -Arguments @{ table_name = "np_nonexistent"; description = "test" }
    Add-ToolCall -Calls $np4Calls -Id 269 -Name "set_subdatasheet_properties" -Arguments @{ table_name = "np_nonexistent"; subdatasheet_name = "[None]" }
    Add-ToolCall -Calls $np4Calls -Id 270 -Name "set_report_grouping" -Arguments @{ report_name = "np_nonexistent"; expression = "f1"; group_on = 0 }
    Add-ToolCall -Calls $np4Calls -Id 271 -Name "set_report_sorting" -Arguments @{ report_name = "np_nonexistent"; order_by = "f1" }
    Add-ToolCall -Calls $np4Calls -Id 272 -Name "set_object_event" -Arguments @{ object_type = "form"; object_name = "np_nonexistent"; event_name = "OnClick"; procedure_name = "Sub1" }
    Add-ToolCall -Calls $np4Calls -Id 273 -Name "set_hidden_attribute" -Arguments @{ object_type = 0; object_name = "np_nonexistent"; hidden = $true }
    Add-ToolCall -Calls $np4Calls -Id 274 -Name "set_query_properties" -Arguments @{ query_name = "np_nonexistent"; description = "test" }
    Add-ToolCall -Calls $np4Calls -Id 275 -Name "set_query_advanced_properties" -Arguments @{ query_name = "np_nonexistent"; max_records = 100 }
    Add-ToolCall -Calls $np4Calls -Id 276 -Name "add_conditional_formatting" -Arguments @{ object_type = "form"; object_name = "np_nonexistent"; control_name = "c1"; expression = "1" }
    Add-ToolCall -Calls $np4Calls -Id 277 -Name "set_document_property" -Arguments @{ container_name = "Tables"; document_name = "np_nonexistent"; property_name = "x"; property_value = "y" }
    Add-ToolCall -Calls $np4Calls -Id 2779 -Name "disconnect_access" -Arguments @{}

    $np4Responses = Invoke-McpBatch -ExePath $ServerExe -Calls $np4Calls -ClientName "neg-batch4a-setters"

    $np4Connect = Decode-McpResult -Response $np4Responses[250]
    if (-not ($np4Connect -and $np4Connect.PSObject.Properties["success"] -and [bool]$np4Connect.success)) {
        throw "Batch 4a: connect_access checkpoint failed."
    }

    for ($i = 251; $i -le 277; $i++) {
        Assert-FailureResponse -Responses $np4Responses -Id $i -Name "batch4a_id$i"
    }

    $np4Disconnect = Decode-McpResult -Response $np4Responses[2779]
    if (-not ($np4Disconnect -and $np4Disconnect.PSObject.Properties["success"] -and [bool]$np4Disconnect.success)) {
        Write-Host "Batch 4a: disconnect checkpoint failed (non-fatal)."
    }

    # ── Batch 4b: Rename + Exclusive-Mode DDL on Non-Existent Objects ──
    # VBE tools (278-280, 283-286) SKIPPED: hang under Office build 19725 (broken .NET COM → VBE)
    Write-Host ""
    Write-Host "=== Batch 4b: Rename + DDL on Non-Existent Objects ==="
    Write-Host "SKIP: IDs 278-280,283-286 (VBE tools hang under Office build 19725)"
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np4bCalls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np4bCalls -Id 2780 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    Add-ToolCall -Calls $np4bCalls -Id 281 -Name "rename_table" -Arguments @{ table_name = "np_nonexistent"; new_table_name = "np_new" }
    Add-ToolCall -Calls $np4bCalls -Id 282 -Name "rename_object" -Arguments @{ object_name = "np_nonexistent"; new_name = "np_new"; object_type = "table" }
    Add-ToolCall -Calls $np4bCalls -Id 287 -Name "add_field" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; type = "TEXT" }
    Add-ToolCall -Calls $np4bCalls -Id 288 -Name "alter_field" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; new_type = "LONG" }
    Add-ToolCall -Calls $np4bCalls -Id 289 -Name "disconnect_access" -Arguments @{}

    $np4bResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $np4bCalls -ClientName "neg-batch4b-rename-ddl"

    $np4bConnect = Decode-McpResult -Response $np4bResponses[2780]
    if (-not ($np4bConnect -and $np4bConnect.PSObject.Properties["success"] -and [bool]$np4bConnect.success)) {
        throw "Batch 4b: connect_access checkpoint failed."
    }

    foreach ($id in @(281, 282, 287, 288)) {
        Assert-FailureResponse -Responses $np4bResponses -Id $id -Name "batch4b_id$id"
    }

    $np4bDisconnect = Decode-McpResult -Response $np4bResponses[289]
    if (-not ($np4bDisconnect -and $np4bDisconnect.PSObject.Properties["success"] -and [bool]$np4bDisconnect.success)) {
        Write-Host "Batch 4b: disconnect checkpoint failed (non-fatal)."
    }

    # ── Batch 5: Remaining Recordset + Form Runtime Invalid Handles (IDs 290-319) ──
    Write-Host ""
    Write-Host "=== Batch 5: Remaining Recordset + Form Runtime Invalid Handles ==="
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np5Calls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np5Calls -Id 290 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    # Recordset tools with invalid handles (IDs 291-301)
    Add-ToolCall -Calls $np5Calls -Id 291 -Name "recordset_add_record" -Arguments @{ recordset_id = "rs_invalid"; fields = @{id=1} }
    Add-ToolCall -Calls $np5Calls -Id 292 -Name "recordset_delete_record" -Arguments @{ recordset_id = "rs_invalid" }
    Add-ToolCall -Calls $np5Calls -Id 293 -Name "recordset_edit_record" -Arguments @{ recordset_id = "rs_invalid"; fields = @{id=2} }
    Add-ToolCall -Calls $np5Calls -Id 294 -Name "recordset_move" -Arguments @{ recordset_id = "rs_invalid"; direction = "next" }
    Add-ToolCall -Calls $np5Calls -Id 295 -Name "recordset_find" -Arguments @{ recordset_id = "rs_invalid"; criteria = "id=1" }
    Add-ToolCall -Calls $np5Calls -Id 296 -Name "recordset_filter_sort" -Arguments @{ recordset_id = "rs_invalid"; sort = "id" }
    Add-ToolCall -Calls $np5Calls -Id 297 -Name "recordset_bookmark" -Arguments @{ recordset_id = "rs_invalid" }
    Add-ToolCall -Calls $np5Calls -Id 298 -Name "recordset_count" -Arguments @{ recordset_id = "rs_invalid" }
    Add-ToolCall -Calls $np5Calls -Id 299 -Name "recordset_get_record" -Arguments @{ recordset_id = "rs_invalid" }
    Add-ToolCall -Calls $np5Calls -Id 300 -Name "recordset_get_rows" -Arguments @{ recordset_id = "rs_invalid"; num_rows = 10 }
    Add-ToolCall -Calls $np5Calls -Id 301 -Name "close_recordset" -Arguments @{ recordset_id = "rs_invalid" }
    # Open/close non-existent objects (IDs 302-318)
    Add-ToolCall -Calls $np5Calls -Id 302 -Name "open_recordset" -Arguments @{ source = "np_nonexistent_table" }
    Add-ToolCall -Calls $np5Calls -Id 303 -Name "open_form" -Arguments @{ form_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 304 -Name "close_form" -Arguments @{ form_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 305 -Name "open_report" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 306 -Name "close_report" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 307 -Name "open_table" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 308 -Name "open_query" -Arguments @{ query_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 309 -Name "close_object" -Arguments @{ object_type = 0; object_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 310 -Name "save_object" -Arguments @{ object_type = 0; object_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 311 -Name "open_module" -Arguments @{ module_name = "np_nonexistent" }
    Add-ToolCall -Calls $np5Calls -Id 312 -Name "control_set_focus" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np5Calls -Id 313 -Name "control_requery" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np5Calls -Id 314 -Name "control_undo" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np5Calls -Id 315 -Name "combobox_dropdown" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np5Calls -Id 316 -Name "listbox_add_item" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1"; item = "x" }
    Add-ToolCall -Calls $np5Calls -Id 317 -Name "listbox_remove_item" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1"; index = 0 }
    Add-ToolCall -Calls $np5Calls -Id 318 -Name "listbox_get_items" -Arguments @{ form_name = "np_nonexistent"; control_name = "c1" }
    Add-ToolCall -Calls $np5Calls -Id 319 -Name "disconnect_access" -Arguments @{}

    $np5Responses = Invoke-McpBatch -ExePath $ServerExe -Calls $np5Calls -ClientName "neg-batch5-handles"

    $np5Connect = Decode-McpResult -Response $np5Responses[290]
    if (-not ($np5Connect -and $np5Connect.PSObject.Properties["success"] -and [bool]$np5Connect.success)) {
        throw "Batch 5: connect_access checkpoint failed."
    }

    # DoCmd.Close on non-open objects is idempotent (no-op, not error); listbox_get_items returns empty
    $np5SkipIds = @(304, 306, 309, 318)  # close_form, close_report, close_object, listbox_get_items
    for ($i = 291; $i -le 318; $i++) {
        if ($np5SkipIds -contains $i) {
            Write-Host "batch5_id${i}: SKIP (idempotent close or empty-result behavior)"
            continue
        }
        Assert-FailureResponse -Responses $np5Responses -Id $i -Name "batch5_id$i"
    }

    $np5Disconnect = Decode-McpResult -Response $np5Responses[319]
    if (-not ($np5Disconnect -and $np5Disconnect.PSObject.Properties["success"] -and [bool]$np5Disconnect.success)) {
        Write-Host "Batch 5: disconnect checkpoint failed (non-fatal)."
    }

    # ── Batch 6: Create Duplicates + Invalid SQL + Export Non-Existent (IDs 320-352) ──
    Write-Host ""
    Write-Host "=== Batch 6: Create Duplicates + Invalid Operations ==="
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np6Calls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np6Calls -Id 320 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    # Pre-cleanup: remove leftover temp objects from prior interrupted runs
    Add-ToolCall -Calls $np6Calls -Id 3200 -Name "delete_table" -Arguments @{ table_name = "np_temp" }
    Add-ToolCall -Calls $np6Calls -Id 3201 -Name "delete_query" -Arguments @{ query_name = "np_temp_q" }
    # Setup: create temp table and query (should succeed)
    Add-ToolCall -Calls $np6Calls -Id 321 -Name "create_table" -Arguments @{ table_name = "np_temp"; fields = @(@{name="id";type="LONG"}) }
    Add-ToolCall -Calls $np6Calls -Id 322 -Name "create_table" -Arguments @{ table_name = "np_temp"; fields = @(@{name="id";type="LONG"}) }
    Add-ToolCall -Calls $np6Calls -Id 323 -Name "create_query" -Arguments @{ query_name = "np_temp_q"; sql = "SELECT 1" }
    Add-ToolCall -Calls $np6Calls -Id 324 -Name "create_query" -Arguments @{ query_name = "np_temp_q"; sql = "SELECT 1" }
    # Invalid creates and operations (IDs 325-349)
    Add-ToolCall -Calls $np6Calls -Id 325 -Name "create_index" -Arguments @{ table_name = "np_nonexistent"; index_name = "idx1"; columns = @("f1") }
    Add-ToolCall -Calls $np6Calls -Id 326 -Name "create_relationship" -Arguments @{ table_name = "np_nonexistent"; field_name = "f1"; foreign_table_name = "np_nonexistent2"; foreign_field_name = "f1" }
    Add-ToolCall -Calls $np6Calls -Id 327 -Name "execute_action_query" -Arguments @{ query_name = "np_nonexistent" }
    Add-ToolCall -Calls $np6Calls -Id 328 -Name "execute_sql" -Arguments @{ sql = "SELECT * FROM np_nonexistent_table" }
    Add-ToolCall -Calls $np6Calls -Id 329 -Name "execute_sql_timed" -Arguments @{ sql = "SELECT * FROM np_nonexistent_table" }
    Add-ToolCall -Calls $np6Calls -Id 330 -Name "run_macro" -Arguments @{ macro_name = "np_nonexistent" }
    Add-ToolCall -Calls $np6Calls -Id 331 -Name "run_vba_procedure" -Arguments @{ procedure_name = "np_nonexistent_proc" }
    Add-ToolCall -Calls $np6Calls -Id 332 -Name "run_data_macro" -Arguments @{ table_name = "np_nonexistent"; macro_name = "np_macro" }
    Add-ToolCall -Calls $np6Calls -Id 333 -Name "export_form_to_text" -Arguments @{ form_name = "np_nonexistent" }
    Add-ToolCall -Calls $np6Calls -Id 334 -Name "export_report_to_text" -Arguments @{ report_name = "np_nonexistent" }
    Add-ToolCall -Calls $np6Calls -Id 335 -Name "export_macro_to_text" -Arguments @{ macro_name = "np_nonexistent" }
    Add-ToolCall -Calls $np6Calls -Id 336 -Name "export_data_macro_axl" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np6Calls -Id 337 -Name "export_xml" -Arguments @{ object_type = 0; data_source = "np_nonexistent"; data_target = "C:\temp\np_out.xml" }
    Add-ToolCall -Calls $np6Calls -Id 338 -Name "import_macro_from_text" -Arguments @{ macro_data = "{invalid-json" }
    Add-ToolCall -Calls $np6Calls -Id 339 -Name "import_data_macro_axl" -Arguments @{ table_name = "np_nonexistent"; axl_xml = "<invalid>" }
    Add-ToolCall -Calls $np6Calls -Id 340 -Name "import_xml" -Arguments @{ data_source = "C:\np_nonexistent_file.xml" }
    # ID 341 (output_to) SKIPPED: pops undismissable modal "Output To" dialog
    # Cleanup: delete temp objects
    Add-ToolCall -Calls $np6Calls -Id 350 -Name "delete_table" -Arguments @{ table_name = "np_temp" }
    Add-ToolCall -Calls $np6Calls -Id 351 -Name "delete_query" -Arguments @{ query_name = "np_temp_q" }
    Add-ToolCall -Calls $np6Calls -Id 352 -Name "disconnect_access" -Arguments @{}

    $np6Responses = Invoke-McpBatch -ExePath $ServerExe -Calls $np6Calls -ClientName "neg-batch6a-duplicates"

    $np6Connect = Decode-McpResult -Response $np6Responses[320]
    if (-not ($np6Connect -and $np6Connect.PSObject.Properties["success"] -and [bool]$np6Connect.success)) {
        throw "Batch 6a: connect_access checkpoint failed."
    }

    $np6CreateTable = Decode-McpResult -Response $np6Responses[321]
    if (-not ($np6CreateTable -and $np6CreateTable.PSObject.Properties["success"] -and [bool]$np6CreateTable.success)) {
        throw "Batch 6a: create_table setup checkpoint failed."
    }

    $np6CreateQuery = Decode-McpResult -Response $np6Responses[323]
    if (-not ($np6CreateQuery -and $np6CreateQuery.PSObject.Properties["success"] -and [bool]$np6CreateQuery.success)) {
        throw "Batch 6a: create_query setup checkpoint failed."
    }

    Assert-FailureResponse -Responses $np6Responses -Id 322 -Name "batch6a_create_table_duplicate"
    Assert-FailureResponse -Responses $np6Responses -Id 324 -Name "batch6a_create_query_duplicate"
    for ($i = 325; $i -le 340; $i++) {
        Assert-FailureResponse -Responses $np6Responses -Id $i -Name "batch6a_id$i"
    }

    $np6DeleteTable = Decode-McpResult -Response $np6Responses[350]
    if (-not ($np6DeleteTable -and $np6DeleteTable.PSObject.Properties["success"] -and [bool]$np6DeleteTable.success)) {
        Write-Host "Batch 6a: delete_table cleanup failed (non-fatal)."
    }

    $np6DeleteQuery = Decode-McpResult -Response $np6Responses[351]
    if (-not ($np6DeleteQuery -and $np6DeleteQuery.PSObject.Properties["success"] -and [bool]$np6DeleteQuery.success)) {
        Write-Host "Batch 6a: delete_query cleanup failed (non-fatal)."
    }

    $np6Disconnect = Decode-McpResult -Response $np6Responses[352]
    if (-not ($np6Disconnect -and $np6Disconnect.PSObject.Properties["success"] -and [bool]$np6Disconnect.success)) {
        Write-Host "Batch 6a: disconnect checkpoint failed (non-fatal)."
    }

    # ── Batch 6b: VBA/Encryption/Attachment Operations (IDs 342-349) ──
    Write-Host ""
    Write-Host "=== Batch 6b: VBA + Encryption + Attachment Operations ==="
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np6bCalls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np6bCalls -Id 3420 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    Add-ToolCall -Calls $np6bCalls -Id 342 -Name "find_and_replace_in_vba" -Arguments @{ find_text = "" }
    Add-ToolCall -Calls $np6bCalls -Id 343 -Name "convert_database" -Arguments @{ source_database_path = "C:\np_nonexistent.accdb"; destination_database_path = "C:\temp\np_out.accdb" }
    Add-ToolCall -Calls $np6bCalls -Id 344 -Name "encrypt_database" -Arguments @{ password = "test" }
    Add-ToolCall -Calls $np6bCalls -Id 345 -Name "set_database_password" -Arguments @{ new_password = "test" }
    Add-ToolCall -Calls $np6bCalls -Id 346 -Name "remove_database_password" -Arguments @{}
    Add-ToolCall -Calls $np6bCalls -Id 347 -Name "reset_autonumber" -Arguments @{ table_name = "np_nonexistent" }
    Add-ToolCall -Calls $np6bCalls -Id 348 -Name "add_attachment_file" -Arguments @{ table_name = "np_nonexistent"; attachment_field = "f1"; record_id = 1; file_path = "C:\np.txt" }
    Add-ToolCall -Calls $np6bCalls -Id 349 -Name "remove_attachment_file" -Arguments @{ table_name = "np_nonexistent"; attachment_field = "f1"; record_id = 1; file_name = "np.txt" }
    Add-ToolCall -Calls $np6bCalls -Id 3499 -Name "disconnect_access" -Arguments @{}

    $np6bResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $np6bCalls -ClientName "neg-batch6b-vba-encrypt"

    $np6bConnect = Decode-McpResult -Response $np6bResponses[3420]
    if (-not ($np6bConnect -and $np6bConnect.PSObject.Properties["success"] -and [bool]$np6bConnect.success)) {
        throw "Batch 6b: connect_access checkpoint failed."
    }

    # ID 346 (remove_database_password) SKIPPED: succeeds as no-op when no password is set
    foreach ($id in @(342, 343, 344, 345, 347, 348, 349)) {
        Assert-FailureResponse -Responses $np6bResponses -Id $id -Name "batch6b_id$id"
    }

    $np6bDisconnect = Decode-McpResult -Response $np6bResponses[3499]
    if (-not ($np6bDisconnect -and $np6bDisconnect.PSObject.Properties["success"] -and [bool]$np6bDisconnect.success)) {
        Write-Host "Batch 6b: disconnect checkpoint failed (non-fatal)."
    }

    # ── Batch 7: Disconnected State - Extended (IDs 360-379) ──
    Write-Host ""
    Write-Host "=== Batch 7: Disconnected State - Extended ==="
    Cleanup-AccessArtifacts -DbPath $DatabasePath

    $np7Calls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $np7Calls -Id 360 -Name "create_table" -Arguments @{ table_name = "x"; fields = @(@{name="id";type="LONG"}) }
    Add-ToolCall -Calls $np7Calls -Id 361 -Name "create_query" -Arguments @{ query_name = "x"; sql = "SELECT 1" }
    Add-ToolCall -Calls $np7Calls -Id 362 -Name "create_form" -Arguments @{ form_name = "x" }
    Add-ToolCall -Calls $np7Calls -Id 363 -Name "create_report" -Arguments @{ report_name = "x" }
    Add-ToolCall -Calls $np7Calls -Id 364 -Name "create_macro" -Arguments @{ macro_name = "x" }
    Add-ToolCall -Calls $np7Calls -Id 365 -Name "add_field" -Arguments @{ table_name = "x"; field_name = "f1"; type = "TEXT" }
    Add-ToolCall -Calls $np7Calls -Id 366 -Name "open_form" -Arguments @{ form_name = "x" }
    Add-ToolCall -Calls $np7Calls -Id 367 -Name "open_report" -Arguments @{ report_name = "x" }
    Add-ToolCall -Calls $np7Calls -Id 368 -Name "open_recordset" -Arguments @{ source = "x" }
    Add-ToolCall -Calls $np7Calls -Id 369 -Name "begin_transaction" -Arguments @{}
    Add-ToolCall -Calls $np7Calls -Id 370 -Name "set_database_property" -Arguments @{ property_name = "x"; property_value = "y" }
    Add-ToolCall -Calls $np7Calls -Id 371 -Name "get_database_properties" -Arguments @{}
    Add-ToolCall -Calls $np7Calls -Id 372 -Name "get_form_controls" -Arguments @{ form_name = "x" }
    Add-ToolCall -Calls $np7Calls -Id 373 -Name "get_modules" -Arguments @{}
    Add-ToolCall -Calls $np7Calls -Id 374 -Name "get_macros" -Arguments @{}
    Add-ToolCall -Calls $np7Calls -Id 375 -Name "export_form_to_text" -Arguments @{ form_name = "x" }
    Add-ToolCall -Calls $np7Calls -Id 376 -Name "import_form_from_text" -Arguments @{ form_data = "{}" }
    Add-ToolCall -Calls $np7Calls -Id 377 -Name "domain_aggregate" -Arguments @{ function = "DCount"; expression = "*"; domain = "x" }
    Add-ToolCall -Calls $np7Calls -Id 378 -Name "find_and_replace_in_vba" -Arguments @{ find_text = "x" }
    Add-ToolCall -Calls $np7Calls -Id 379 -Name "is_connected" -Arguments @{}

    $np7Responses = Invoke-McpBatch -ExePath $ServerExe -Calls $np7Calls -ClientName "neg-batch7-disconnected"

    for ($i = 360; $i -le 378; $i++) {
        Assert-FailureResponse -Responses $np7Responses -Id $i -Name "batch7_disconnected_id$i"
    }

    # is_connected (ID 379) may return success=true with connected=false; both outcomes are acceptable
    if ($np7Responses.ContainsKey(379)) {
        $np7IsConnected = Decode-McpResult -Response $np7Responses[379]
        if ($np7IsConnected -and $np7IsConnected.PSObject.Properties["success"] -and [bool]$np7IsConnected.success) {
            if ($np7IsConnected.PSObject.Properties["connected"] -and [bool]$np7IsConnected.connected) {
                throw "Batch 7: is_connected unexpectedly shows connected=true in disconnected state."
            }
        }
        # success=false is also acceptable
    }

    Write-Host "NEGATIVE_PATHS_PASS=1"
    $exitCode = 0
}
finally {
    Write-Host "Final cleanup: clearing stale Access/MCP processes and locks."

    # Stop dialog watcher and write diagnostics summary
    if ($null -ne $script:DialogWatcherState) {
        Stop-DialogWatcher -WatcherState $script:DialogWatcherState
        if (-not [string]::IsNullOrWhiteSpace($script:DiagnosticsDir)) {
            Write-DialogWatcherSummary -JsonlPath $script:DialogWatcherState.JsonlPath
            Write-DiagnosticsSummary -DiagnosticsPath $script:DiagnosticsDir `
                -JsonlPath $script:DialogWatcherState.JsonlPath `
                -TimeoutCount $script:TimeoutCount `
                -TimeoutSections $script:TimeoutSections
        }
    }

    if ($script:TimeoutCount -gt 0) {
        Write-Host ("TIMEOUT_SECTIONS={0} ({1})" -f $script:TimeoutCount, (($script:TimeoutSections.Keys | Sort-Object) -join ", "))
    }

    Cleanup-AccessArtifacts -DbPath $DatabasePath
    Release-RegressionLock -LockState $regressionLock
}

exit $exitCode
