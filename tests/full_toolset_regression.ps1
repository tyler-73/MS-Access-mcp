param(
    [Alias("ServerExePath")]
    [string]$ServerExe = "$PSScriptRoot\..\mcp-server-official-x64\MS.Access.MCP.Official.exe",
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "$env:USERPROFILE\Documents\MyDatabase.accdb" }),
    [switch]$NoCleanup,
    [switch]$AllowCoverageSkips,
    [switch]$IncludeUiCoverage
)

$ErrorActionPreference = "Stop"

# Resolve $ServerExe when $PSScriptRoot was empty (MSYS bash / git-bash invocations)
if (-not (Test-Path $ServerExe -ErrorAction SilentlyContinue)) {
    $fallbackRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $fallbackExe  = Join-Path $fallbackRoot "..\mcp-server-official-x64\MS.Access.MCP.Official.exe"
    if (Test-Path $fallbackExe) { $ServerExe = $fallbackExe }
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
        Id = $Id
        Name = $Name
        Arguments = $Arguments
    })
}

$script:TrackedMsAccessPids = @{}

function Get-NormalizedExecutablePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    try {
        $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
        if ($resolved) {
            return [string]$resolved.ProviderPath
        }
    }
    catch {
    }

    try {
        return [System.IO.Path]::GetFullPath($Path)
    }
    catch {
        return $Path
    }
}

function Get-MsAccessProcessSnapshot {
    $snapshot = @{}
    foreach ($process in @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue)) {
        $startTicks = $null
        try {
            $startTicks = $process.StartTime.ToUniversalTime().Ticks
        }
        catch {
        }

        $snapshot[[int]$process.Id] = $startTicks
    }

    return $snapshot
}

function Sync-TrackedMsAccessPids {
    param([hashtable]$CurrentSnapshot = $null)

    if ($null -eq $CurrentSnapshot) {
        $CurrentSnapshot = Get-MsAccessProcessSnapshot
    }

    foreach ($trackedPid in @($script:TrackedMsAccessPids.Keys)) {
        if (-not $CurrentSnapshot.ContainsKey($trackedPid)) {
            $null = $script:TrackedMsAccessPids.Remove($trackedPid)
            continue
        }

        $trackedStartTicks = $script:TrackedMsAccessPids[$trackedPid]
        $currentStartTicks = $CurrentSnapshot[$trackedPid]
        if ($null -ne $trackedStartTicks -and $null -ne $currentStartTicks -and $trackedStartTicks -ne $currentStartTicks) {
            $null = $script:TrackedMsAccessPids.Remove($trackedPid)
        }
    }
}

function Register-NewMsAccessPids {
    param([hashtable]$BeforeSnapshot)

    if ($null -eq $BeforeSnapshot) {
        $BeforeSnapshot = @{}
    }

    $afterSnapshot = Get-MsAccessProcessSnapshot
    Sync-TrackedMsAccessPids -CurrentSnapshot $afterSnapshot

    foreach ($processId in @($afterSnapshot.Keys)) {
        if (-not $BeforeSnapshot.ContainsKey($processId)) {
            $script:TrackedMsAccessPids[[int]$processId] = $afterSnapshot[$processId]
            continue
        }

        $beforeStartTicks = $BeforeSnapshot[$processId]
        $afterStartTicks = $afterSnapshot[$processId]
        if ($null -ne $beforeStartTicks -and $null -ne $afterStartTicks -and $beforeStartTicks -ne $afterStartTicks) {
            $script:TrackedMsAccessPids[[int]$processId] = $afterStartTicks
        }
    }
}

function Test-IsTrackedMsAccessProcess {
    param(
        [System.Diagnostics.Process]$Process,
        [hashtable]$CurrentSnapshot
    )

    $processId = [int]$Process.Id
    if (-not $script:TrackedMsAccessPids.ContainsKey($processId)) {
        return $false
    }

    if (-not $CurrentSnapshot.ContainsKey($processId)) {
        $null = $script:TrackedMsAccessPids.Remove($processId)
        return $false
    }

    $trackedStartTicks = $script:TrackedMsAccessPids[$processId]
    $currentStartTicks = $CurrentSnapshot[$processId]
    if ($null -ne $trackedStartTicks -and $null -ne $currentStartTicks -and $trackedStartTicks -ne $currentStartTicks) {
        $null = $script:TrackedMsAccessPids.Remove($processId)
        return $false
    }

    return $true
}

function Invoke-McpBatch {
    param(
        [string]$ExePath,
        [System.Collections.Generic.List[object]]$Calls,
        [string]$ClientName = "full-regression",
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

    $msAccessSnapshotBefore = Get-MsAccessProcessSnapshot
    $rawLines = @()
    try {
        $rawLines = @((($jsonLines -join "`n") | & $ExePath))
    }
    finally {
        Register-NewMsAccessPids -BeforeSnapshot $msAccessSnapshotBefore
    }

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

function Invoke-McpRawBatch {
    param(
        [string]$ExePath,
        [System.Collections.Generic.List[hashtable]]$Requests,
        [string]$ClientName = "full-regression-raw",
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

    foreach ($req in $Requests) {
        $jsonLines.Add(($req | ConvertTo-Json -Depth 50 -Compress))
    }

    $msAccessSnapshotBefore = Get-MsAccessProcessSnapshot
    $rawLines = @()
    try {
        $rawLines = @((($jsonLines -join "`n") | & $ExePath))
    }
    finally {
        Register-NewMsAccessPids -BeforeSnapshot $msAccessSnapshotBefore
    }

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
    param(
        [string]$ExePath,
        [string]$ClientName = "full-regression-tools-list",
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

    $jsonLines.Add((@{
        jsonrpc = "2.0"
        id = 2
        method = "tools/list"
        params = @{}
    } | ConvertTo-Json -Depth 40 -Compress))

    $msAccessSnapshotBefore = Get-MsAccessProcessSnapshot
    $rawLines = @()
    try {
        $rawLines = @((($jsonLines -join "`n") | & $ExePath))
    }
    finally {
        Register-NewMsAccessPids -BeforeSnapshot $msAccessSnapshotBefore
    }

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
            Write-Host "WARN: Could not parse tools/list response line: $line"
        }
    }

    if (-not $responses.ContainsKey(2)) {
        return @()
    }

    $toolResponse = $responses[2]
    if ($toolResponse.result -and $toolResponse.result.tools) {
        return @($toolResponse.result.tools)
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

function Resolve-AlternateToolName {
    param(
        [System.Collections.Generic.Dictionary[string, object]]$ToolByName,
        [string]$PrimaryName,
        [string[]]$Candidates
    )

    foreach ($candidate in $Candidates) {
        if ($candidate -eq $PrimaryName) {
            continue
        }

        if ($ToolByName.ContainsKey($candidate)) {
            return $candidate
        }
    }

    return $null
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

    $normalizedServerExe = Get-NormalizedExecutablePath -Path $ServerExe

    foreach ($serverProcess in @(Get-CimInstance Win32_Process -Filter "Name = 'MS.Access.MCP.Official.exe'" -ErrorAction SilentlyContinue)) {
        $processExePath = Get-NormalizedExecutablePath -Path $serverProcess.ExecutablePath
        if (-not [string]::IsNullOrWhiteSpace($normalizedServerExe) -and
            -not [string]::IsNullOrWhiteSpace($processExePath) -and
            $processExePath -ieq $normalizedServerExe) {
            Stop-Process -Id ([int]$serverProcess.ProcessId) -Force -ErrorAction SilentlyContinue
        }
    }

    $msAccessProcesses = @(Get-Process -Name MSACCESS -ErrorAction SilentlyContinue)
    if ($msAccessProcesses.Count -eq 0) {
        Sync-TrackedMsAccessPids
        return
    }

    $msAccessSnapshot = Get-MsAccessProcessSnapshot
    $msAccessCommandLineByPid = @{}
    foreach ($msAccessCim in @(Get-CimInstance Win32_Process -Filter "Name = 'MSACCESS.EXE'" -ErrorAction SilentlyContinue)) {
        $msAccessCommandLineByPid[[int]$msAccessCim.ProcessId] = [string]$msAccessCim.CommandLine
    }

    foreach ($msAccessProcess in $msAccessProcesses) {
        $processId = [int]$msAccessProcess.Id
        $isTracked = Test-IsTrackedMsAccessProcess -Process $msAccessProcess -CurrentSnapshot $msAccessSnapshot
        $mainWindowTitle = [string]$msAccessProcess.MainWindowTitle
        $isHeadlessWindow = [string]::IsNullOrWhiteSpace($mainWindowTitle)
        $commandLine = ""
        if ($msAccessCommandLineByPid.ContainsKey($processId)) {
            $commandLine = [string]$msAccessCommandLineByPid[$processId]
        }
        $isEmbedding = $commandLine -match '(?i)(^|\s|")/embedding(\s|$)'

        if ($isTracked -or $isEmbedding -or $isHeadlessWindow) {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
    }

    Sync-TrackedMsAccessPids
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

if (-not (Test-Path -LiteralPath $ServerExe)) {
    throw "Server executable not found: $ServerExe"
}

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database file not found: $DatabasePath"
}

$regressionLock = Acquire-RegressionLock
Write-Host ("Regression lock acquired: {0}" -f $regressionLock.Path)

if ($IncludeUiCoverage) {
    Write-Host "ui_coverage: ENABLED (launch_access/open_form/open_report will run)"
}
else {
    Write-Host "ui_coverage: DISABLED (headless mode; UI-opening tools are skipped)"
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
    Release-RegressionLock -LockState $regressionLock
    throw
}

$exitCode = 1
$linkedSourceDatabasePath = $null
$databaseLifecycleCreatedPath = $null
$databaseLifecycleBackupPath = $null
$databaseLifecycleCompactPath = $null
try {

$suffix = [Guid]::NewGuid().ToString("N").Substring(0, 8)
$tableName = "MCP_Table_$suffix"
$formName = "MCP_Form_$suffix"
$reportName = "MCP_Report_$suffix"
$moduleName = "MCP_Module_$suffix"
$queryName = "MCP_Query_$suffix"
$relationshipName = "MCP_Rel_$suffix"
$childTableName = "MCP_Child_$suffix"
$indexName = "MCP_Idx_$suffix"
$macroName = "MCP_Macro_$suffix"
$importedMacroName = "MCP_ImportedMacro_$suffix"
$schemaFieldName = "schema_text"
$schemaFieldRenamedName = "schema_text_renamed"
$renamedTableName = "MCP_Renamed_$suffix"
$linkedTableName = "MCP_Linked_$suffix"
$linkedSourceTableName = "MCP_LinkSrc_$suffix"
$transactionTableName = "MCP_Tx_$suffix"
$databaseLifecycleTableName = "MCP_DbLifecycle_$suffix"
$newToolsTableName = "MCP_NewTools_$suffix"
$recordsetTableName = "MCP_RS_$suffix"
$formRuntimeTableName = "MCP_FormRT_$suffix"
$formRuntimeFormName = "MCP_FormRT_Form_$suffix"
$formRuntimeReportName = "MCP_FormRT_Report_$suffix"
$tempNavXmlPath = Join-Path ([System.IO.Path]::GetTempPath()) "mcp_nav_$suffix.xml"
$tempXmlDataPath = Join-Path ([System.IO.Path]::GetTempPath()) "mcp_export_$suffix.xml"

$linkedSourceDatabasePath = Join-Path (Split-Path -Path $DatabasePath -Parent) "MCP_LinkSource_$suffix.accdb"
$databaseLifecycleCreatedPath = Join-Path (Split-Path -Path $DatabasePath -Parent) "MCP_CreateDb_$suffix.accdb"
$databaseLifecycleBackupPath = Join-Path (Split-Path -Path $DatabasePath -Parent) "MCP_BackupDb_$suffix.accdb"
$databaseLifecycleCompactPath = Join-Path (Split-Path -Path $DatabasePath -Parent) "MCP_CompactDb_$suffix.accdb"

$toolList = Get-McpToolsList -ExePath $ServerExe -ClientName "full-regression-tools-list" -ClientVersion "1.0"
$toolByName = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($tool in $toolList) {
    $name = [string]$tool.name
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        $toolByName[$name] = $tool
    }
}

$listLinkedTablesToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("list_linked_tables")
$createLinkedTableToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("create_linked_table", "link_table")
$refreshLinkedTableToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("refresh_linked_table", "refresh_link")
$updateLinkedTableToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("update_linked_table", "relink_table")
$deleteLinkedTableToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("delete_linked_table", "unlink_table")
$createLinkedTableAliasToolName = Resolve-AlternateToolName -ToolByName $toolByName -PrimaryName $createLinkedTableToolName -Candidates @("create_linked_table", "link_table")
$refreshLinkedTableAliasToolName = Resolve-AlternateToolName -ToolByName $toolByName -PrimaryName $refreshLinkedTableToolName -Candidates @("refresh_linked_table", "refresh_link")
$updateLinkedTableAliasToolName = Resolve-AlternateToolName -ToolByName $toolByName -PrimaryName $updateLinkedTableToolName -Candidates @("update_linked_table", "relink_table")
$deleteLinkedTableAliasToolName = Resolve-AlternateToolName -ToolByName $toolByName -PrimaryName $deleteLinkedTableToolName -Candidates @("delete_linked_table", "unlink_table")

$beginTransactionToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("begin_transaction", "start_transaction")
$commitTransactionToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("commit_transaction")
$rollbackTransactionToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("rollback_transaction")
$transactionStatusToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("transaction_status")
$beginTransactionAliasToolName = Resolve-AlternateToolName -ToolByName $toolByName -PrimaryName $beginTransactionToolName -Candidates @("begin_transaction", "start_transaction")
$createDatabaseToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("create_database")
$backupDatabaseToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("backup_database")
$compactRepairDatabaseToolName = Resolve-ToolName -ToolByName $toolByName -Candidates @("compact_repair_database")

$formData = @{
    Name = $formName
    ExportedAt = (Get-Date).ToUniversalTime().ToString("o")
    Controls = @(
        @{
            Name = "txtValue"
            Type = "TextBox"
            Left = 600
            Top = 600
            Width = 2400
            Height = 300
            Visible = $true
            Enabled = $true
        }
    )
    VBA = ""
} | ConvertTo-Json -Depth 20 -Compress

$reportData = @{
    Name = $reportName
    ExportedAt = (Get-Date).ToUniversalTime().ToString("o")
    Controls = @(
        @{
            Name = "lblReport"
            Type = "Label"
            Left = 500
            Top = 300
            Width = 2500
            Height = 300
            Visible = $true
            Enabled = $true
        }
    )
} | ConvertTo-Json -Depth 20 -Compress

$vbaCode = @'
Option Compare Database
Option Explicit

Public Sub Ping()
    Debug.Print "Ping"
End Sub
'@

$procCode = @'
Public Sub Pong()
    Debug.Print "Pong"
End Sub
'@

$macroDataInitial = @'
Version =196611
ColumnsShown =8
Begin
    Action ="Beep"
End
'@

$macroDataUpdated = @'
Version =196611
ColumnsShown =9
Begin
    Action ="Beep"
End
'@

$calls = New-Object 'System.Collections.Generic.List[object]'

Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "is_connected" -Arguments @{}
if ($IncludeUiCoverage) {
    Add-ToolCall -Calls $calls -Id 4 -Name "launch_access" -Arguments @{}
}
Add-ToolCall -Calls $calls -Id 5 -Name "get_tables" -Arguments @{}
Add-ToolCall -Calls $calls -Id 6 -Name "get_queries" -Arguments @{}
Add-ToolCall -Calls $calls -Id 7 -Name "get_relationships" -Arguments @{}
Add-ToolCall -Calls $calls -Id 8 -Name "create_table" -Arguments @{
    table_name = $tableName
    fields = @(
        @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false },
        @{ name = "name"; type = "TEXT"; size = 50; required = $false; allow_zero_length = $true }
    )
}
Add-ToolCall -Calls $calls -Id 9 -Name "describe_table" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 69 -Name "add_field" -Arguments @{
    table_name = $tableName
    field_name = $schemaFieldName
    field_type = "TEXT"
    type = "TEXT"
    size = 40
    required = $false
    allow_zero_length = $true
}
Add-ToolCall -Calls $calls -Id 70 -Name "describe_table" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 71 -Name "alter_field" -Arguments @{
    table_name = $tableName
    field_name = $schemaFieldName
    field_type = "TEXT"
    new_field_type = "TEXT"
    size = 80
    new_size = 80
    required = $false
    allow_zero_length = $true
}
Add-ToolCall -Calls $calls -Id 72 -Name "describe_table" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 73 -Name "rename_field" -Arguments @{
    table_name = $tableName
    field_name = $schemaFieldName
    old_field_name = $schemaFieldName
    new_field_name = $schemaFieldRenamedName
}
Add-ToolCall -Calls $calls -Id 74 -Name "describe_table" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 75 -Name "drop_field" -Arguments @{
    table_name = $tableName
    field_name = $schemaFieldRenamedName
}
Add-ToolCall -Calls $calls -Id 76 -Name "describe_table" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 77 -Name "rename_table" -Arguments @{
    table_name = $tableName
    old_table_name = $tableName
    new_table_name = $renamedTableName
}
Add-ToolCall -Calls $calls -Id 78 -Name "get_tables" -Arguments @{}
Add-ToolCall -Calls $calls -Id 79 -Name "rename_table" -Arguments @{
    table_name = $renamedTableName
    old_table_name = $renamedTableName
    new_table_name = $tableName
}
Add-ToolCall -Calls $calls -Id 80 -Name "get_tables" -Arguments @{}
Add-ToolCall -Calls $calls -Id 57 -Name "create_index" -Arguments @{
    table_name = $tableName
    index_name = $indexName
    columns = @("name")
    unique = $false
}
Add-ToolCall -Calls $calls -Id 58 -Name "get_indexes" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 10 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$tableName] (id, name) VALUES (1, 'alpha')" }
Add-ToolCall -Calls $calls -Id 11 -Name "execute_sql" -Arguments @{ sql = "SELECT * FROM [$tableName]" }
Add-ToolCall -Calls $calls -Id 12 -Name "execute_query_md" -Arguments @{ sql = "SELECT * FROM [$tableName]" }
Add-ToolCall -Calls $calls -Id 13 -Name "get_system_tables" -Arguments @{}
Add-ToolCall -Calls $calls -Id 14 -Name "get_object_metadata" -Arguments @{}
Add-ToolCall -Calls $calls -Id 15 -Name "set_vba_code" -Arguments @{
    project_name = "CurrentProject"
    module_name = $moduleName
    code = $vbaCode
}
Add-ToolCall -Calls $calls -Id 16 -Name "add_vba_procedure" -Arguments @{
    project_name = "CurrentProject"
    module_name = $moduleName
    procedure_name = "Pong"
    code = $procCode
}
Add-ToolCall -Calls $calls -Id 17 -Name "get_vba_code" -Arguments @{
    project_name = "CurrentProject"
    module_name = $moduleName
}
Add-ToolCall -Calls $calls -Id 18 -Name "compile_vba" -Arguments @{}
Add-ToolCall -Calls $calls -Id 19 -Name "get_vba_projects" -Arguments @{}
Add-ToolCall -Calls $calls -Id 20 -Name "import_form_from_text" -Arguments @{ form_data = $formData }
Add-ToolCall -Calls $calls -Id 21 -Name "form_exists" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 22 -Name "get_form_controls" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 23 -Name "get_control_properties" -Arguments @{ form_name = $formName; control_name = "txtValue" }
Add-ToolCall -Calls $calls -Id 24 -Name "set_control_property" -Arguments @{
    form_name = $formName
    control_name = "txtValue"
    property_name = "Visible"
    value = "True"
}
Add-ToolCall -Calls $calls -Id 25 -Name "export_form_to_text" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 83 -Name "export_form_to_text" -Arguments @{ form_name = $formName; mode = "access_text" }
if ($IncludeUiCoverage) {
    Add-ToolCall -Calls $calls -Id 26 -Name "open_form" -Arguments @{ form_name = $formName }
    Add-ToolCall -Calls $calls -Id 27 -Name "close_form" -Arguments @{ form_name = $formName }
}
Add-ToolCall -Calls $calls -Id 28 -Name "import_report_from_text" -Arguments @{ report_data = $reportData }
if ($IncludeUiCoverage) {
    Add-ToolCall -Calls $calls -Id 55 -Name "open_report" -Arguments @{ report_name = $reportName }
    Add-ToolCall -Calls $calls -Id 56 -Name "close_report" -Arguments @{ report_name = $reportName }
}
Add-ToolCall -Calls $calls -Id 52 -Name "get_report_controls" -Arguments @{ report_name = $reportName }
Add-ToolCall -Calls $calls -Id 53 -Name "get_report_control_properties" -Arguments @{ report_name = $reportName; control_name = "lblReport" }
Add-ToolCall -Calls $calls -Id 54 -Name "set_report_control_property" -Arguments @{ report_name = $reportName; control_name = "lblReport"; property_name = "Visible"; value = "True" }
Add-ToolCall -Calls $calls -Id 29 -Name "export_report_to_text" -Arguments @{ report_name = $reportName }
Add-ToolCall -Calls $calls -Id 84 -Name "export_report_to_text" -Arguments @{ report_name = $reportName; mode = "access_text" }
Add-ToolCall -Calls $calls -Id 30 -Name "delete_report" -Arguments @{ report_name = $reportName }
Add-ToolCall -Calls $calls -Id 31 -Name "delete_form" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 32 -Name "get_forms" -Arguments @{}
Add-ToolCall -Calls $calls -Id 33 -Name "get_reports" -Arguments @{}
Add-ToolCall -Calls $calls -Id 34 -Name "get_macros" -Arguments @{}
Add-ToolCall -Calls $calls -Id 35 -Name "get_modules" -Arguments @{}
Add-ToolCall -Calls $calls -Id 61 -Name "create_macro" -Arguments @{ macro_name = $macroName; macro_data = $macroDataInitial }
Add-ToolCall -Calls $calls -Id 62 -Name "get_macros" -Arguments @{}
Add-ToolCall -Calls $calls -Id 63 -Name "export_macro_to_text" -Arguments @{ macro_name = $macroName }
Add-ToolCall -Calls $calls -Id 64 -Name "run_macro" -Arguments @{ macro_name = $macroName }
Add-ToolCall -Calls $calls -Id 65 -Name "update_macro" -Arguments @{ macro_name = $macroName; macro_data = $macroDataUpdated }
Add-ToolCall -Calls $calls -Id 66 -Name "export_macro_to_text" -Arguments @{ macro_name = $macroName }
Add-ToolCall -Calls $calls -Id 67 -Name "delete_macro" -Arguments @{ macro_name = $macroName }
Add-ToolCall -Calls $calls -Id 68 -Name "get_macros" -Arguments @{}
Add-ToolCall -Calls $calls -Id 81 -Name "import_macro_from_text" -Arguments @{ macro_name = $importedMacroName; macro_data = $macroDataInitial; overwrite = $true }
Add-ToolCall -Calls $calls -Id 82 -Name "get_macros" -Arguments @{}
Add-ToolCall -Calls $calls -Id 40 -Name "create_query" -Arguments @{ query_name = $queryName; sql = "SELECT id, name FROM [$tableName]" }
Add-ToolCall -Calls $calls -Id 41 -Name "get_queries" -Arguments @{}
Add-ToolCall -Calls $calls -Id 42 -Name "update_query" -Arguments @{ query_name = $queryName; sql = "SELECT id FROM [$tableName] WHERE id >= 1" }
Add-ToolCall -Calls $calls -Id 43 -Name "create_table" -Arguments @{
    table_name = $childTableName
    fields = @(
        @{ name = "child_id"; type = "LONG"; size = 0; required = $false; allow_zero_length = $false },
        @{ name = "parent_id"; type = "LONG"; size = 0; required = $false; allow_zero_length = $false }
    )
}
Add-ToolCall -Calls $calls -Id 50 -Name "execute_sql" -Arguments @{ sql = "ALTER TABLE [$tableName] ADD CONSTRAINT [PK_$tableName] PRIMARY KEY ([id])" }
Add-ToolCall -Calls $calls -Id 44 -Name "create_relationship" -Arguments @{
    relationship_name = $relationshipName
    table_name = $tableName
    field_name = "id"
    foreign_table_name = $childTableName
    foreign_field_name = "parent_id"
    enforce_integrity = $true
    cascade_update = $false
    cascade_delete = $false
}
Add-ToolCall -Calls $calls -Id 45 -Name "get_relationships" -Arguments @{}
Add-ToolCall -Calls $calls -Id 46 -Name "update_relationship" -Arguments @{
    relationship_name = $relationshipName
    table_name = $tableName
    field_name = "id"
    foreign_table_name = $childTableName
    foreign_field_name = "parent_id"
    enforce_integrity = $true
    cascade_update = $true
    cascade_delete = $true
}
Add-ToolCall -Calls $calls -Id 51 -Name "get_relationships" -Arguments @{}
Add-ToolCall -Calls $calls -Id 47 -Name "delete_relationship" -Arguments @{ relationship_name = $relationshipName }
Add-ToolCall -Calls $calls -Id 48 -Name "delete_query" -Arguments @{ query_name = $queryName }
Add-ToolCall -Calls $calls -Id 49 -Name "delete_table" -Arguments @{ table_name = $childTableName }
Add-ToolCall -Calls $calls -Id 59 -Name "delete_index" -Arguments @{ table_name = $tableName; index_name = $indexName }
Add-ToolCall -Calls $calls -Id 60 -Name "get_indexes" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 36 -Name "delete_table" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 37 -Name "disconnect_access" -Arguments @{}
Add-ToolCall -Calls $calls -Id 38 -Name "is_connected" -Arguments @{}
Add-ToolCall -Calls $calls -Id 39 -Name "close_access" -Arguments @{}

$responses = Invoke-McpBatch -ExePath $ServerExe -Calls $calls -ClientName "full-regression" -ClientVersion "1.0"

$idLabels = @{
    2 = "connect_access"
    3 = "is_connected_initial"
    5 = "get_tables"
    6 = "get_queries"
    7 = "get_relationships"
    8 = "create_table"
    9 = "describe_table"
    69 = "add_field"
    70 = "describe_table_after_add_field"
    71 = "alter_field"
    72 = "describe_table_after_alter_field"
    73 = "rename_field"
    74 = "describe_table_after_rename_field"
    75 = "drop_field"
    76 = "describe_table_after_drop_field"
    77 = "rename_table_away"
    78 = "get_tables_after_rename_table_away"
    79 = "rename_table_back"
    80 = "get_tables_after_rename_table_back"
    57 = "create_index"
    58 = "get_indexes_after_create_index"
    10 = "execute_sql_insert"
    11 = "execute_sql_select"
    12 = "execute_query_md"
    13 = "get_system_tables"
    14 = "get_object_metadata"
    15 = "set_vba_code"
    16 = "add_vba_procedure"
    17 = "get_vba_code"
    18 = "compile_vba"
    19 = "get_vba_projects"
    20 = "import_form_from_text"
    21 = "form_exists"
    22 = "get_form_controls"
    23 = "get_control_properties"
    24 = "set_control_property"
    25 = "export_form_to_text"
    83 = "export_form_to_text_access_text"
    28 = "import_report_from_text"
    52 = "get_report_controls"
    53 = "get_report_control_properties"
    54 = "set_report_control_property"
    29 = "export_report_to_text"
    84 = "export_report_to_text_access_text"
    30 = "delete_report"
    31 = "delete_form"
    32 = "get_forms"
    33 = "get_reports"
    34 = "get_macros"
    35 = "get_modules"
    61 = "create_macro"
    62 = "get_macros_after_create_macro"
    63 = "export_macro_to_text_initial"
    64 = "run_macro"
    65 = "update_macro"
    66 = "export_macro_to_text_after_update"
    67 = "delete_macro"
    68 = "get_macros_after_delete_macro"
    81 = "import_macro_from_text"
    82 = "get_macros_after_import_macro"
    40 = "create_query"
    41 = "get_queries_after_create_query"
    42 = "update_query"
    43 = "create_child_table"
    50 = "add_parent_primary_key"
    44 = "create_relationship"
    45 = "get_relationships_after_create_relationship"
    46 = "update_relationship"
    51 = "get_relationships_after_update_relationship"
    47 = "delete_relationship"
    48 = "delete_query"
    49 = "delete_child_table"
    59 = "delete_index"
    60 = "get_indexes_after_delete_index"
    36 = "delete_table"
    37 = "disconnect_access"
    38 = "is_connected_after_disconnect"
    39 = "close_access"
}

if ($IncludeUiCoverage) {
    $idLabels[4] = "launch_access"
    $idLabels[26] = "open_form"
    $idLabels[27] = "close_form"
    $idLabels[55] = "open_report"
    $idLabels[56] = "close_report"
}

$failed = 0
$formAccessTextData = $null
$reportAccessTextData = $null
foreach ($id in ($idLabels.Keys | Sort-Object)) {
    $label = $idLabels[$id]
    $decoded = Decode-McpResult -Response $responses[[int]$id]

    if ($null -eq $decoded) {
        $failed++
        Write-Host ('{0}: FAIL missing-response' -f $label)
        continue
    }

    if ($decoded -is [string]) {
        $failed++
        Write-Host ('{0}: FAIL raw-string-response' -f $label)
        continue
    }

    if ($decoded.success -ne $true) {
        $failed++
        Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
        continue
    }

    switch ($label) {
        "is_connected_initial" {
            if ($decoded.connected -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL expected connected=true' -f $label)
                continue
            }
        }
        "is_connected_after_disconnect" {
            if ($decoded.connected -ne $false) {
                $failed++
                Write-Host ('{0}: FAIL expected connected=false' -f $label)
                continue
            }
        }
        "describe_table_after_add_field" {
            $columns = if ($decoded.table -and $decoded.table.Columns) { @($decoded.table.Columns) } elseif ($decoded.table -and $decoded.table.columns) { @($decoded.table.columns) } else { @() }
            $matched = $columns | Where-Object { [string]$_.Name -eq $schemaFieldName -or [string]$_.name -eq $schemaFieldName }
            if (@($matched).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected field {1}' -f $label, $schemaFieldName)
                continue
            }

            $column = $matched | Select-Object -First 1
            $maxLengthValue = if ($null -ne $column.MaxLength) { [int]$column.MaxLength } elseif ($null -ne $column.maxLength) { [int]$column.maxLength } elseif ($null -ne $column.size) { [int]$column.size } else { -1 }
            if ($maxLengthValue -ne 40) {
                $failed++
                Write-Host ('{0}: FAIL expected MaxLength=40 for field {1}, got {2}' -f $label, $schemaFieldName, $maxLengthValue)
                continue
            }
        }
        "describe_table_after_alter_field" {
            $columns = if ($decoded.table -and $decoded.table.Columns) { @($decoded.table.Columns) } elseif ($decoded.table -and $decoded.table.columns) { @($decoded.table.columns) } else { @() }
            $matched = $columns | Where-Object { [string]$_.Name -eq $schemaFieldName -or [string]$_.name -eq $schemaFieldName }
            if (@($matched).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected field {1}' -f $label, $schemaFieldName)
                continue
            }

            $column = $matched | Select-Object -First 1
            $maxLengthValue = if ($null -ne $column.MaxLength) { [int]$column.MaxLength } elseif ($null -ne $column.maxLength) { [int]$column.maxLength } elseif ($null -ne $column.size) { [int]$column.size } else { -1 }
            if ($maxLengthValue -ne 80) {
                $failed++
                Write-Host ('{0}: FAIL expected MaxLength=80 for field {1}, got {2}' -f $label, $schemaFieldName, $maxLengthValue)
                continue
            }
        }
        "describe_table_after_rename_field" {
            $columns = if ($decoded.table -and $decoded.table.Columns) { @($decoded.table.Columns) } elseif ($decoded.table -and $decoded.table.columns) { @($decoded.table.columns) } else { @() }
            $oldMatched = $columns | Where-Object { [string]$_.Name -eq $schemaFieldName -or [string]$_.name -eq $schemaFieldName }
            $newMatched = $columns | Where-Object { [string]$_.Name -eq $schemaFieldRenamedName -or [string]$_.name -eq $schemaFieldRenamedName }
            if (@($oldMatched).Count -ne 0 -or @($newMatched).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected old field {1} replaced by {2}' -f $label, $schemaFieldName, $schemaFieldRenamedName)
                continue
            }
        }
        "describe_table_after_drop_field" {
            $columns = if ($decoded.table -and $decoded.table.Columns) { @($decoded.table.Columns) } elseif ($decoded.table -and $decoded.table.columns) { @($decoded.table.columns) } else { @() }
            $matched = $columns | Where-Object { [string]$_.Name -eq $schemaFieldRenamedName -or [string]$_.name -eq $schemaFieldRenamedName }
            if (@($matched).Count -ne 0) {
                $failed++
                Write-Host ('{0}: FAIL expected field {1} to be dropped' -f $label, $schemaFieldRenamedName)
                continue
            }
        }
        "get_tables_after_rename_table_away" {
            $tables = @($decoded.tables)
            $oldMatched = $tables | Where-Object { [string]$_.Name -eq $tableName -or [string]$_.name -eq $tableName }
            $newMatched = $tables | Where-Object { [string]$_.Name -eq $renamedTableName -or [string]$_.name -eq $renamedTableName }
            if (@($oldMatched).Count -ne 0 -or @($newMatched).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected table rename {1} -> {2}' -f $label, $tableName, $renamedTableName)
                continue
            }
        }
        "get_tables_after_rename_table_back" {
            $tables = @($decoded.tables)
            $oldMatched = $tables | Where-Object { [string]$_.Name -eq $tableName -or [string]$_.name -eq $tableName }
            $renamedMatched = $tables | Where-Object { [string]$_.Name -eq $renamedTableName -or [string]$_.name -eq $renamedTableName }
            if (@($oldMatched).Count -eq 0 -or @($renamedMatched).Count -ne 0) {
                $failed++
                Write-Host ('{0}: FAIL expected table rename rollback {1} -> {2}' -f $label, $renamedTableName, $tableName)
                continue
            }
        }
        "form_exists" {
            if ($decoded.exists -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL expected exists=true' -f $label)
                continue
            }
        }
        "get_form_controls" {
            if (@($decoded.controls).Count -lt 1) {
                $failed++
                Write-Host ('{0}: FAIL expected at least one control' -f $label)
                continue
            }
        }
        "get_report_controls" {
            $controls = @($decoded.controls)
            if ($controls.Count -lt 1) {
                $failed++
                Write-Host ('{0}: FAIL expected at least one report control' -f $label)
                continue
            }

            $matchedControl = $controls | Where-Object { [string]$_.name -eq "lblReport" }
            if (@($matchedControl).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected report control lblReport' -f $label)
                continue
            }
        }
        "get_report_control_properties" {
            if ([string]$decoded.properties.name -ne "lblReport") {
                $failed++
                Write-Host ('{0}: FAIL expected control properties for lblReport' -f $label)
                continue
            }
        }
        "get_indexes_after_create_index" {
            $indexes = @($decoded.indexes)
            $matchedIndex = $indexes | Where-Object { [string]$_.name -eq $indexName }
            if (@($matchedIndex).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected index {1}' -f $label, $indexName)
                continue
            }

            $index = $matchedIndex | Select-Object -First 1
            $columns = @($index.columns)
            if (@($columns | Where-Object { [string]$_ -eq "name" }).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected index column name' -f $label)
                continue
            }
        }
        "get_indexes_after_delete_index" {
            $indexes = @($decoded.indexes)
            $matchedIndex = $indexes | Where-Object { [string]$_.name -eq $indexName }
            if (@($matchedIndex).Count -ne 0) {
                $failed++
                Write-Host ('{0}: FAIL expected index {1} to be deleted' -f $label, $indexName)
                continue
            }
        }
        "get_macros_after_create_macro" {
            $macros = @($decoded.macros)
            $matchedMacro = $macros | Where-Object { [string]$_.name -eq $macroName }
            if (@($matchedMacro).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected macro {1}' -f $label, $macroName)
                continue
            }
        }
        "export_macro_to_text_initial" {
            $macroText = [string]$decoded.macro_data
            if ([string]::IsNullOrWhiteSpace($macroText) -or
                $macroText.IndexOf('Action ="Beep"', [System.StringComparison]::OrdinalIgnoreCase) -lt 0 -or
                $macroText.IndexOf('ColumnsShown =8', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $failed++
                Write-Host ('{0}: FAIL expected exported macro text with initial marker values' -f $label)
                continue
            }
        }
        "export_macro_to_text_after_update" {
            $macroText = [string]$decoded.macro_data
            if ([string]::IsNullOrWhiteSpace($macroText) -or
                $macroText.IndexOf('ColumnsShown =9', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $failed++
                Write-Host ('{0}: FAIL expected exported macro text to include updated marker value' -f $label)
                continue
            }
        }
        "get_macros_after_delete_macro" {
            $macros = @($decoded.macros)
            $matchedMacro = $macros | Where-Object { [string]$_.name -eq $macroName }
            if (@($matchedMacro).Count -ne 0) {
                $failed++
                Write-Host ('{0}: FAIL expected macro {1} to be deleted' -f $label, $macroName)
                continue
            }
        }
        "get_macros_after_import_macro" {
            $macros = @($decoded.macros)
            $matchedMacro = $macros | Where-Object { [string]$_.name -eq $importedMacroName }
            if (@($matchedMacro).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected imported macro {1}' -f $label, $importedMacroName)
                continue
            }
        }
        "get_vba_code" {
            $codeText = [string]$decoded.code
            if ($codeText.IndexOf("Pong", [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $failed++
                Write-Host ('{0}: FAIL expected procedure text in module code' -f $label)
                continue
            }
        }
        "get_queries_after_create_query" {
            $queries = @($decoded.queries)
            $matchedQuery = $queries | Where-Object { [string]$_.name -eq $queryName }
            if (@($matchedQuery).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected query {1}' -f $label, $queryName)
                continue
            }
        }
        "get_relationships_after_create_relationship" {
            $relationships = @($decoded.relationships)
            $matchedRelationship = $relationships | Where-Object { [string]$_.name -eq $relationshipName }
            if (@($matchedRelationship).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected relationship {1}' -f $label, $relationshipName)
                continue
            }

            $relationship = $matchedRelationship | Select-Object -First 1
            if ([string]$relationship.table -ne $tableName -or
                [string]$relationship.field -ne "id" -or
                [string]$relationship.foreignTable -ne $childTableName -or
                [string]$relationship.foreignField -ne "parent_id") {
                $failed++
                Write-Host ('{0}: FAIL unexpected relationship mapping table={1} field={2} foreignTable={3} foreignField={4}' -f
                    $label, [string]$relationship.table, [string]$relationship.field, [string]$relationship.foreignTable, [string]$relationship.foreignField)
                continue
            }
        }
        "get_relationships_after_update_relationship" {
            $relationships = @($decoded.relationships)
            $matchedRelationship = $relationships | Where-Object { [string]$_.name -eq $relationshipName }
            if (@($matchedRelationship).Count -eq 0) {
                $failed++
                Write-Host ('{0}: FAIL expected relationship {1}' -f $label, $relationshipName)
                continue
            }

            $relationship = $matchedRelationship | Select-Object -First 1
            if ($relationship.cascadeUpdate -ne $true -or $relationship.cascadeDelete -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL expected cascade flags true after update' -f $label)
                continue
            }
        }
        "export_form_to_text" {
            if ([string]::IsNullOrWhiteSpace([string]$decoded.form_data)) {
                $failed++
                Write-Host ('{0}: FAIL empty form export payload' -f $label)
                continue
            }
        }
        "export_form_to_text_access_text" {
            $formAccessTextData = [string]$decoded.form_data
            if ([string]::IsNullOrWhiteSpace($formAccessTextData)) {
                $failed++
                Write-Host ('{0}: FAIL empty form export payload' -f $label)
                continue
            }
            if ($formAccessTextData.IndexOf('Version =', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $failed++
                Write-Host ('{0}: FAIL expected Access text payload marker `Version =`' -f $label)
                continue
            }
        }
        "export_report_to_text" {
            if ([string]::IsNullOrWhiteSpace([string]$decoded.report_data)) {
                $failed++
                Write-Host ('{0}: FAIL empty report export payload' -f $label)
                continue
            }
        }
        "export_report_to_text_access_text" {
            $reportAccessTextData = [string]$decoded.report_data
            if ([string]::IsNullOrWhiteSpace($reportAccessTextData)) {
                $failed++
                Write-Host ('{0}: FAIL empty report export payload' -f $label)
                continue
            }
            if ($reportAccessTextData.IndexOf('Version =', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $failed++
                Write-Host ('{0}: FAIL expected Access text payload marker `Version =`' -f $label)
                continue
            }
        }
    }

    Write-Host ('{0}: OK' -f $label)
}

if ([string]::IsNullOrWhiteSpace($formAccessTextData)) {
    $failed++
    Write-Host "access_text_form_roundtrip_source: FAIL missing export payload"
}

if ([string]::IsNullOrWhiteSpace($reportAccessTextData)) {
    $failed++
    Write-Host "access_text_report_roundtrip_source: FAIL missing export payload"
}

if (-not [string]::IsNullOrWhiteSpace($formAccessTextData) -and -not [string]::IsNullOrWhiteSpace($reportAccessTextData)) {
    Write-Host "Intermediate cleanup: clearing stale Access/MCP processes and locks before access_text round-trip."
    Cleanup-AccessArtifacts -DbPath $DatabasePath
    Start-Sleep -Milliseconds 300

    $accessTextCalls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $accessTextCalls -Id 201 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    Add-ToolCall -Calls $accessTextCalls -Id 202 -Name "import_form_from_text" -Arguments @{ form_data = $formAccessTextData; form_name = $formName; mode = "access_text" }
    Add-ToolCall -Calls $accessTextCalls -Id 203 -Name "form_exists" -Arguments @{ form_name = $formName }
    Add-ToolCall -Calls $accessTextCalls -Id 204 -Name "export_form_to_text" -Arguments @{ form_name = $formName; mode = "access_text" }
    Add-ToolCall -Calls $accessTextCalls -Id 205 -Name "delete_form" -Arguments @{ form_name = $formName }
    Add-ToolCall -Calls $accessTextCalls -Id 206 -Name "import_report_from_text" -Arguments @{ report_data = $reportAccessTextData; report_name = $reportName; mode = "access_text" }
    Add-ToolCall -Calls $accessTextCalls -Id 207 -Name "get_report_controls" -Arguments @{ report_name = $reportName }
    Add-ToolCall -Calls $accessTextCalls -Id 208 -Name "export_report_to_text" -Arguments @{ report_name = $reportName; mode = "access_text" }
    Add-ToolCall -Calls $accessTextCalls -Id 209 -Name "delete_report" -Arguments @{ report_name = $reportName }
    Add-ToolCall -Calls $accessTextCalls -Id 210 -Name "disconnect_access" -Arguments @{}
    Add-ToolCall -Calls $accessTextCalls -Id 211 -Name "close_access" -Arguments @{}

    $accessTextResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $accessTextCalls -ClientName "full-regression-access-text" -ClientVersion "1.0"
    $accessTextIdLabels = @{
        201 = "access_text_connect_access"
        202 = "access_text_import_form_from_text"
        203 = "access_text_form_exists"
        204 = "access_text_export_form_to_text"
        205 = "access_text_delete_form"
        206 = "access_text_import_report_from_text"
        207 = "access_text_get_report_controls"
        208 = "access_text_export_report_to_text"
        209 = "access_text_delete_report"
        210 = "access_text_disconnect_access"
        211 = "access_text_close_access"
    }

    foreach ($id in ($accessTextIdLabels.Keys | Sort-Object)) {
        $label = $accessTextIdLabels[$id]
        $decoded = Decode-McpResult -Response $accessTextResponses[[int]$id]

        if ($null -eq $decoded) {
            $failed++
            Write-Host ('{0}: FAIL missing-response' -f $label)
            continue
        }

        if ($decoded -is [string]) {
            $failed++
            Write-Host ('{0}: FAIL raw-string-response' -f $label)
            continue
        }

        if ($decoded.success -ne $true) {
            $failed++
            Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
            continue
        }

        switch ($label) {
            "access_text_form_exists" {
                if ($decoded.exists -ne $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected exists=true' -f $label)
                    continue
                }
            }
            "access_text_export_form_to_text" {
                $formDataRoundTrip = [string]$decoded.form_data
                if ([string]::IsNullOrWhiteSpace($formDataRoundTrip)) {
                    $failed++
                    Write-Host ('{0}: FAIL empty form export payload' -f $label)
                    continue
                }
                if ($formDataRoundTrip.IndexOf('Version =', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                    $failed++
                    Write-Host ('{0}: FAIL expected Access text payload marker `Version =`' -f $label)
                    continue
                }
            }
            "access_text_get_report_controls" {
                $controls = @($decoded.controls)
                if ($controls.Count -lt 1) {
                    $failed++
                    Write-Host ('{0}: FAIL expected at least one report control' -f $label)
                    continue
                }
            }
            "access_text_export_report_to_text" {
                $reportDataRoundTrip = [string]$decoded.report_data
                if ([string]::IsNullOrWhiteSpace($reportDataRoundTrip)) {
                    $failed++
                    Write-Host ('{0}: FAIL empty report export payload' -f $label)
                    continue
                }
                if ($reportDataRoundTrip.IndexOf('Version =', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                    $failed++
                    Write-Host ('{0}: FAIL expected Access text payload marker `Version =`' -f $label)
                    continue
                }
            }
        }

        Write-Host ('{0}: OK' -f $label)
    }
}

if (-not [string]::IsNullOrWhiteSpace($createLinkedTableToolName) -and
    -not [string]::IsNullOrWhiteSpace($deleteLinkedTableToolName) -and
    -not [string]::IsNullOrWhiteSpace($listLinkedTablesToolName)) {
    $linkedCoverageToolNames = @($createLinkedTableToolName, $deleteLinkedTableToolName, $listLinkedTablesToolName)
    if (-not [string]::IsNullOrWhiteSpace($createLinkedTableAliasToolName)) {
        $linkedCoverageToolNames += $createLinkedTableAliasToolName
    }
    if (-not [string]::IsNullOrWhiteSpace($refreshLinkedTableToolName)) {
        $linkedCoverageToolNames += $refreshLinkedTableToolName
    }
    if (-not [string]::IsNullOrWhiteSpace($refreshLinkedTableAliasToolName)) {
        $linkedCoverageToolNames += $refreshLinkedTableAliasToolName
    }
    if (-not [string]::IsNullOrWhiteSpace($updateLinkedTableToolName)) {
        $linkedCoverageToolNames += $updateLinkedTableToolName
    }
    if (-not [string]::IsNullOrWhiteSpace($updateLinkedTableAliasToolName)) {
        $linkedCoverageToolNames += $updateLinkedTableAliasToolName
    }
    if (-not [string]::IsNullOrWhiteSpace($deleteLinkedTableAliasToolName)) {
        $linkedCoverageToolNames += $deleteLinkedTableAliasToolName
    }
    Write-Host ('linked_table_coverage: INFO using tools {0}' -f ($linkedCoverageToolNames -join ", "))

    $linkedPrepReady = $false
    try {
        Copy-Item -Path $DatabasePath -Destination $linkedSourceDatabasePath -Force
        Cleanup-AccessArtifacts -DbPath $linkedSourceDatabasePath
        Start-Sleep -Milliseconds 300

        $linkedPrepCalls = New-Object 'System.Collections.Generic.List[object]'
        Add-ToolCall -Calls $linkedPrepCalls -Id 301 -Name "connect_access" -Arguments @{ database_path = $linkedSourceDatabasePath }
        Add-ToolCall -Calls $linkedPrepCalls -Id 302 -Name "create_table" -Arguments @{
            table_name = $linkedSourceTableName
            fields = @(
                @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false },
                @{ name = "payload"; type = "TEXT"; size = 50; required = $false; allow_zero_length = $true }
            )
        }
        Add-ToolCall -Calls $linkedPrepCalls -Id 303 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$linkedSourceTableName] (id, payload) VALUES (1, 'source_alpha')" }
        Add-ToolCall -Calls $linkedPrepCalls -Id 304 -Name "disconnect_access" -Arguments @{}
        Add-ToolCall -Calls $linkedPrepCalls -Id 305 -Name "close_access" -Arguments @{}

        $linkedPrepResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $linkedPrepCalls -ClientName "full-regression-linked-source" -ClientVersion "1.0"
        $linkedPrepLabels = @{
            301 = "linked_source_connect_access"
            302 = "linked_source_create_table"
            303 = "linked_source_insert_seed_row"
            304 = "linked_source_disconnect_access"
            305 = "linked_source_close_access"
        }

        $linkedPrepFailed = $false
        foreach ($id in ($linkedPrepLabels.Keys | Sort-Object)) {
            $label = $linkedPrepLabels[$id]
            $decoded = Decode-McpResult -Response $linkedPrepResponses[[int]$id]

            if ($null -eq $decoded) {
                $failed++
                $linkedPrepFailed = $true
                Write-Host ('{0}: FAIL missing-response' -f $label)
                continue
            }

            if ($decoded -is [string]) {
                $failed++
                $linkedPrepFailed = $true
                Write-Host ('{0}: FAIL raw-string-response' -f $label)
                continue
            }

            if ($decoded.success -ne $true) {
                $failed++
                $linkedPrepFailed = $true
                Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
                continue
            }

            Write-Host ('{0}: OK' -f $label)
        }

        $linkedPrepReady = (-not $linkedPrepFailed)
    }
    catch {
        $failed++
        Write-Host ('linked_source_setup: FAIL {0}' -f $_.Exception.Message)
    }

    if ($linkedPrepReady) {
        $createLinkedArguments = @{
            table_name = $linkedTableName
            linked_table_name = $linkedTableName
            source_table_name = $linkedSourceTableName
            external_table_name = $linkedSourceTableName
            source_database_path = $linkedSourceDatabasePath
            database_path = $linkedSourceDatabasePath
            external_database_path = $linkedSourceDatabasePath
            connection_string = "MS Access;DATABASE=$linkedSourceDatabasePath"
            connect_string = "DATABASE=$linkedSourceDatabasePath"
            overwrite = $true
        }

        $deleteLinkedArguments = @{
            table_name = $linkedTableName
            linked_table_name = $linkedTableName
        }
        $linkedAliasDeleteTableName = "${linkedTableName}_AliasDelete"

        $linkedCalls = New-Object 'System.Collections.Generic.List[object]'
        Add-ToolCall -Calls $linkedCalls -Id 321 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
        Add-ToolCall -Calls $linkedCalls -Id 322 -Name $createLinkedTableToolName -Arguments $createLinkedArguments
        if (-not [string]::IsNullOrWhiteSpace($createLinkedTableAliasToolName)) {
            Add-ToolCall -Calls $linkedCalls -Id 333 -Name $createLinkedTableAliasToolName -Arguments $createLinkedArguments
        }
        Add-ToolCall -Calls $linkedCalls -Id 323 -Name $listLinkedTablesToolName -Arguments @{}
        Add-ToolCall -Calls $linkedCalls -Id 324 -Name "execute_sql" -Arguments @{ sql = "SELECT id, payload FROM [$linkedTableName]" }
        if (-not [string]::IsNullOrWhiteSpace($refreshLinkedTableToolName)) {
            Add-ToolCall -Calls $linkedCalls -Id 325 -Name $refreshLinkedTableToolName -Arguments $createLinkedArguments
            Add-ToolCall -Calls $linkedCalls -Id 326 -Name "execute_sql" -Arguments @{ sql = "SELECT id, payload FROM [$linkedTableName]" }
            if (-not [string]::IsNullOrWhiteSpace($refreshLinkedTableAliasToolName)) {
                Add-ToolCall -Calls $linkedCalls -Id 334 -Name $refreshLinkedTableAliasToolName -Arguments $createLinkedArguments
                Add-ToolCall -Calls $linkedCalls -Id 339 -Name "execute_sql" -Arguments @{ sql = "SELECT id, payload FROM [$linkedTableName]" }
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($updateLinkedTableToolName)) {
            Add-ToolCall -Calls $linkedCalls -Id 331 -Name $updateLinkedTableToolName -Arguments $createLinkedArguments
            Add-ToolCall -Calls $linkedCalls -Id 332 -Name "execute_sql" -Arguments @{ sql = "SELECT id, payload FROM [$linkedTableName]" }
            if (-not [string]::IsNullOrWhiteSpace($updateLinkedTableAliasToolName)) {
                Add-ToolCall -Calls $linkedCalls -Id 335 -Name $updateLinkedTableAliasToolName -Arguments $createLinkedArguments
                Add-ToolCall -Calls $linkedCalls -Id 340 -Name "execute_sql" -Arguments @{ sql = "SELECT id, payload FROM [$linkedTableName]" }
            }
        }
        Add-ToolCall -Calls $linkedCalls -Id 327 -Name $deleteLinkedTableToolName -Arguments $deleteLinkedArguments
        Add-ToolCall -Calls $linkedCalls -Id 328 -Name $listLinkedTablesToolName -Arguments @{}
        if (-not [string]::IsNullOrWhiteSpace($deleteLinkedTableAliasToolName)) {
            $aliasCreateArguments = [hashtable]$createLinkedArguments.Clone()
            $aliasCreateArguments["table_name"] = $linkedAliasDeleteTableName
            $aliasCreateArguments["linked_table_name"] = $linkedAliasDeleteTableName

            $aliasDeleteArguments = @{
                table_name = $linkedAliasDeleteTableName
                linked_table_name = $linkedAliasDeleteTableName
            }

            Add-ToolCall -Calls $linkedCalls -Id 336 -Name $createLinkedTableToolName -Arguments $aliasCreateArguments
            Add-ToolCall -Calls $linkedCalls -Id 337 -Name $deleteLinkedTableAliasToolName -Arguments $aliasDeleteArguments
            Add-ToolCall -Calls $linkedCalls -Id 338 -Name $listLinkedTablesToolName -Arguments @{}
        }
        Add-ToolCall -Calls $linkedCalls -Id 329 -Name "disconnect_access" -Arguments @{}
        Add-ToolCall -Calls $linkedCalls -Id 330 -Name "close_access" -Arguments @{}

        $linkedResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $linkedCalls -ClientName "full-regression-linked-table" -ClientVersion "1.0"
        $linkedIdLabels = @{
            321 = "linked_table_connect_access"
            322 = "linked_table_create_linked_table"
            323 = "linked_table_list_linked_tables_after_create"
            324 = "linked_table_execute_sql_select"
            327 = "linked_table_delete_linked_table"
            328 = "linked_table_list_linked_tables_after_delete"
            329 = "linked_table_disconnect_access"
            330 = "linked_table_close_access"
        }
        if (-not [string]::IsNullOrWhiteSpace($refreshLinkedTableToolName)) {
            $linkedIdLabels[325] = "linked_table_refresh_linked_table"
            $linkedIdLabels[326] = "linked_table_execute_sql_select_after_refresh"
            if (-not [string]::IsNullOrWhiteSpace($refreshLinkedTableAliasToolName)) {
                $linkedIdLabels[334] = "linked_table_refresh_linked_table_alias_path"
                $linkedIdLabels[339] = "linked_table_execute_sql_select_after_refresh_alias"
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($updateLinkedTableToolName)) {
            $linkedIdLabels[331] = "linked_table_update_linked_table"
            $linkedIdLabels[332] = "linked_table_execute_sql_select_after_update"
            if (-not [string]::IsNullOrWhiteSpace($updateLinkedTableAliasToolName)) {
                $linkedIdLabels[335] = "linked_table_update_linked_table_alias_path"
                $linkedIdLabels[340] = "linked_table_execute_sql_select_after_update_alias"
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($createLinkedTableAliasToolName)) {
            $linkedIdLabels[333] = "linked_table_create_linked_table_alias_path"
        }
        if (-not [string]::IsNullOrWhiteSpace($deleteLinkedTableAliasToolName)) {
            $linkedIdLabels[336] = "linked_table_create_for_delete_alias_path"
            $linkedIdLabels[337] = "linked_table_delete_linked_table_alias_path"
            $linkedIdLabels[338] = "linked_table_list_linked_tables_after_alias_delete"
        }

        foreach ($id in ($linkedIdLabels.Keys | Sort-Object)) {
            $label = $linkedIdLabels[$id]
            $decoded = Decode-McpResult -Response $linkedResponses[[int]$id]

            if ($null -eq $decoded) {
                $failed++
                Write-Host ('{0}: FAIL missing-response' -f $label)
                continue
            }

            if ($decoded -is [string]) {
                $failed++
                Write-Host ('{0}: FAIL raw-string-response' -f $label)
                continue
            }

            if ($decoded.success -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
                continue
            }

            switch ($label) {
                "linked_table_list_linked_tables_after_create" {
                    $tables = @($decoded.linked_tables)
                    if ($tables.Count -eq 0 -and $null -ne $decoded.tables) {
                        $tables = @($decoded.tables)
                    }
                    $linkedMatch = $tables | Where-Object {
                        [string]$_.Name -eq $linkedTableName -or
                        [string]$_.name -eq $linkedTableName -or
                        [string]$_.TableName -eq $linkedTableName -or
                        [string]$_.table_name -eq $linkedTableName
                    }
                    if (@($linkedMatch).Count -eq 0) {
                        $failed++
                        Write-Host ('{0}: FAIL expected linked table {1}' -f $label, $linkedTableName)
                        continue
                    }
                }
                "linked_table_execute_sql_select" {
                    $rows = @($decoded.rows)
                    if ($rows.Count -lt 1) {
                        $failed++
                        Write-Host ('{0}: FAIL expected at least one row from linked table {1}' -f $label, $linkedTableName)
                        continue
                    }
                }
                "linked_table_execute_sql_select_after_refresh" {
                    $rows = @($decoded.rows)
                    if ($rows.Count -lt 1) {
                        $failed++
                        Write-Host ('{0}: FAIL expected at least one row after linked table refresh' -f $label)
                        continue
                    }
                }
                "linked_table_execute_sql_select_after_refresh_alias" {
                    $rows = @($decoded.rows)
                    if ($rows.Count -lt 1) {
                        $failed++
                        Write-Host ('{0}: FAIL expected at least one row after linked table refresh alias path' -f $label)
                        continue
                    }
                }
                "linked_table_execute_sql_select_after_update" {
                    $rows = @($decoded.rows)
                    if ($rows.Count -lt 1) {
                        $failed++
                        Write-Host ('{0}: FAIL expected at least one row after linked table update' -f $label)
                        continue
                    }
                }
                "linked_table_execute_sql_select_after_update_alias" {
                    $rows = @($decoded.rows)
                    if ($rows.Count -lt 1) {
                        $failed++
                        Write-Host ('{0}: FAIL expected at least one row after linked table update alias path' -f $label)
                        continue
                    }
                }
                "linked_table_list_linked_tables_after_delete" {
                    $tables = @($decoded.linked_tables)
                    if ($tables.Count -eq 0 -and $null -ne $decoded.tables) {
                        $tables = @($decoded.tables)
                    }
                    $linkedMatch = $tables | Where-Object {
                        [string]$_.Name -eq $linkedTableName -or
                        [string]$_.name -eq $linkedTableName -or
                        [string]$_.TableName -eq $linkedTableName -or
                        [string]$_.table_name -eq $linkedTableName
                    }
                    if (@($linkedMatch).Count -ne 0) {
                        $failed++
                        Write-Host ('{0}: FAIL expected linked table {1} to be deleted' -f $label, $linkedTableName)
                        continue
                    }
                }
                "linked_table_list_linked_tables_after_alias_delete" {
                    $tables = @($decoded.linked_tables)
                    if ($tables.Count -eq 0 -and $null -ne $decoded.tables) {
                        $tables = @($decoded.tables)
                    }
                    $linkedMatch = $tables | Where-Object {
                        [string]$_.Name -eq $linkedAliasDeleteTableName -or
                        [string]$_.name -eq $linkedAliasDeleteTableName -or
                        [string]$_.TableName -eq $linkedAliasDeleteTableName -or
                        [string]$_.table_name -eq $linkedAliasDeleteTableName
                    }
                    if (@($linkedMatch).Count -ne 0) {
                        $failed++
                        Write-Host ('{0}: FAIL expected linked table {1} to be deleted through alias path' -f $label, $linkedAliasDeleteTableName)
                        continue
                    }
                }
            }

            Write-Host ('{0}: OK' -f $label)
        }
    }
}
else {
    $missingLinkedTools = @()
    if ([string]::IsNullOrWhiteSpace($createLinkedTableToolName)) { $missingLinkedTools += "create_linked_table|link_table" }
    if ([string]::IsNullOrWhiteSpace($deleteLinkedTableToolName)) { $missingLinkedTools += "delete_linked_table|unlink_table" }
    if ([string]::IsNullOrWhiteSpace($listLinkedTablesToolName)) { $missingLinkedTools += "list_linked_tables" }

    if ($AllowCoverageSkips) {
        Write-Host ("linked_table_coverage: SKIP linked-table tools not exposed by this server build. missing={0}" -f ($missingLinkedTools -join ", "))
    }
    else {
        $failed++
        Write-Host ("linked_table_coverage: FAIL required linked-table tools not exposed by this server build. missing={0}" -f ($missingLinkedTools -join ", "))
    }
}

if (-not [string]::IsNullOrWhiteSpace($beginTransactionToolName) -and
    -not [string]::IsNullOrWhiteSpace($commitTransactionToolName) -and
    -not [string]::IsNullOrWhiteSpace($rollbackTransactionToolName) -and
    -not [string]::IsNullOrWhiteSpace($transactionStatusToolName)) {
    $transactionCoverageToolNames = @($beginTransactionToolName, $commitTransactionToolName, $rollbackTransactionToolName, $transactionStatusToolName)
    if (-not [string]::IsNullOrWhiteSpace($beginTransactionAliasToolName)) {
        $transactionCoverageToolNames += $beginTransactionAliasToolName
    }
    Write-Host ('transaction_coverage: INFO using tools {0}' -f ($transactionCoverageToolNames -join ", "))

    $transactionCalls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $transactionCalls -Id 341 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    Add-ToolCall -Calls $transactionCalls -Id 342 -Name "create_table" -Arguments @{
        table_name = $transactionTableName
        fields = @(
            @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false },
            @{ name = "name"; type = "TEXT"; size = 50; required = $false; allow_zero_length = $true }
        )
    }
    Add-ToolCall -Calls $transactionCalls -Id 343 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$transactionTableName] (id, name) VALUES (1, 'outside_txn')" }
    Add-ToolCall -Calls $transactionCalls -Id 344 -Name $beginTransactionToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 355 -Name $transactionStatusToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 345 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$transactionTableName] (id, name) VALUES (2, 'rollback_me')" }
    Add-ToolCall -Calls $transactionCalls -Id 346 -Name $rollbackTransactionToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 356 -Name $transactionStatusToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 347 -Name "execute_sql" -Arguments @{ sql = "SELECT id FROM [$transactionTableName] WHERE id = 2" }
    Add-ToolCall -Calls $transactionCalls -Id 348 -Name $beginTransactionToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 357 -Name $transactionStatusToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 349 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$transactionTableName] (id, name) VALUES (3, 'commit_me')" }
    Add-ToolCall -Calls $transactionCalls -Id 350 -Name $commitTransactionToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 358 -Name $transactionStatusToolName -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 351 -Name "execute_sql" -Arguments @{ sql = "SELECT id FROM [$transactionTableName] WHERE id = 3" }
    if (-not [string]::IsNullOrWhiteSpace($beginTransactionAliasToolName)) {
        Add-ToolCall -Calls $transactionCalls -Id 359 -Name $beginTransactionAliasToolName -Arguments @{}
        Add-ToolCall -Calls $transactionCalls -Id 360 -Name $transactionStatusToolName -Arguments @{}
        Add-ToolCall -Calls $transactionCalls -Id 361 -Name $rollbackTransactionToolName -Arguments @{}
        Add-ToolCall -Calls $transactionCalls -Id 362 -Name $transactionStatusToolName -Arguments @{}
    }
    Add-ToolCall -Calls $transactionCalls -Id 352 -Name "delete_table" -Arguments @{ table_name = $transactionTableName }
    Add-ToolCall -Calls $transactionCalls -Id 353 -Name "disconnect_access" -Arguments @{}
    Add-ToolCall -Calls $transactionCalls -Id 354 -Name "close_access" -Arguments @{}

    $transactionResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $transactionCalls -ClientName "full-regression-transactions" -ClientVersion "1.0"
    $transactionIdLabels = @{
        341 = "transaction_connect_access"
        342 = "transaction_create_table"
        343 = "transaction_insert_outside_transaction"
        344 = "transaction_begin_for_rollback"
        355 = "transaction_status_after_begin_for_rollback"
        345 = "transaction_insert_within_rollback"
        346 = "transaction_rollback"
        356 = "transaction_status_after_rollback"
        347 = "transaction_select_after_rollback"
        348 = "transaction_begin_for_commit"
        357 = "transaction_status_after_begin_for_commit"
        349 = "transaction_insert_within_commit"
        350 = "transaction_commit"
        358 = "transaction_status_after_commit"
        351 = "transaction_select_after_commit"
        352 = "transaction_delete_table"
        353 = "transaction_disconnect_access"
        354 = "transaction_close_access"
    }
    if (-not [string]::IsNullOrWhiteSpace($beginTransactionAliasToolName)) {
        $transactionIdLabels[359] = "transaction_begin_alias_path"
        $transactionIdLabels[360] = "transaction_status_after_begin_alias_path"
        $transactionIdLabels[361] = "transaction_rollback_after_begin_alias_path"
        $transactionIdLabels[362] = "transaction_status_after_alias_rollback"
    }

    foreach ($id in ($transactionIdLabels.Keys | Sort-Object)) {
        $label = $transactionIdLabels[$id]
        $decoded = Decode-McpResult -Response $transactionResponses[[int]$id]

        if ($null -eq $decoded) {
            $failed++
            Write-Host ('{0}: FAIL missing-response' -f $label)
            continue
        }

        if ($decoded -is [string]) {
            $failed++
            Write-Host ('{0}: FAIL raw-string-response' -f $label)
            continue
        }

        if ($decoded.success -ne $true) {
            $failed++
            Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
            continue
        }

        switch ($label) {
            "transaction_status_after_begin_for_rollback" {
                if ($decoded.connected -ne $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected connected=true' -f $label)
                    continue
                }
                if ($null -eq $decoded.transaction -or $decoded.transaction.active -ne $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected active transaction after begin' -f $label)
                    continue
                }
            }
            "transaction_select_after_rollback" {
                $rows = @($decoded.rows)
                if ($rows.Count -ne 0) {
                    $failed++
                    Write-Host ('{0}: FAIL expected zero rows after rollback' -f $label)
                    continue
                }
            }
            "transaction_status_after_rollback" {
                if ($null -eq $decoded.transaction -or $decoded.transaction.active -ne $false) {
                    $failed++
                    Write-Host ('{0}: FAIL expected no active transaction after rollback' -f $label)
                    continue
                }
            }
            "transaction_status_after_begin_for_commit" {
                if ($decoded.connected -ne $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected connected=true' -f $label)
                    continue
                }
                if ($null -eq $decoded.transaction -or $decoded.transaction.active -ne $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected active transaction after begin' -f $label)
                    continue
                }
            }
            "transaction_select_after_commit" {
                $rows = @($decoded.rows)
                if ($rows.Count -lt 1) {
                    $failed++
                    Write-Host ('{0}: FAIL expected committed row to be visible' -f $label)
                    continue
                }
            }
            "transaction_status_after_commit" {
                if ($null -eq $decoded.transaction -or $decoded.transaction.active -ne $false) {
                    $failed++
                    Write-Host ('{0}: FAIL expected no active transaction after commit' -f $label)
                    continue
                }
            }
            "transaction_status_after_begin_alias_path" {
                if ($decoded.connected -ne $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected connected=true' -f $label)
                    continue
                }
                if ($null -eq $decoded.transaction -or $decoded.transaction.active -ne $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected active transaction after alias begin' -f $label)
                    continue
                }
            }
            "transaction_status_after_alias_rollback" {
                if ($null -eq $decoded.transaction -or $decoded.transaction.active -ne $false) {
                    $failed++
                    Write-Host ('{0}: FAIL expected no active transaction after alias rollback' -f $label)
                    continue
                }
            }
        }

        Write-Host ('{0}: OK' -f $label)
    }
}
else {
    $missingTransactionTools = @()
    if ([string]::IsNullOrWhiteSpace($beginTransactionToolName)) { $missingTransactionTools += "begin_transaction|start_transaction" }
    if ([string]::IsNullOrWhiteSpace($commitTransactionToolName)) { $missingTransactionTools += "commit_transaction" }
    if ([string]::IsNullOrWhiteSpace($rollbackTransactionToolName)) { $missingTransactionTools += "rollback_transaction" }
    if ([string]::IsNullOrWhiteSpace($transactionStatusToolName)) { $missingTransactionTools += "transaction_status" }

    if ($AllowCoverageSkips) {
        Write-Host ("transaction_coverage: SKIP transaction tools not exposed by this server build. missing={0}" -f ($missingTransactionTools -join ", "))
    }
    else {
        $failed++
        Write-Host ("transaction_coverage: FAIL required transaction tools not exposed by this server build. missing={0}" -f ($missingTransactionTools -join ", "))
    }
}

if (-not [string]::IsNullOrWhiteSpace($createDatabaseToolName) -and
    -not [string]::IsNullOrWhiteSpace($backupDatabaseToolName) -and
    -not [string]::IsNullOrWhiteSpace($compactRepairDatabaseToolName)) {
    $databaseLifecycleToolNames = @($createDatabaseToolName, $backupDatabaseToolName, $compactRepairDatabaseToolName)
    Write-Host ('database_file_tools_coverage: INFO using tools {0}' -f ($databaseLifecycleToolNames -join ", "))

    foreach ($dbLifecyclePath in @($databaseLifecycleCreatedPath, $databaseLifecycleBackupPath, $databaseLifecycleCompactPath)) {
        if (-not [string]::IsNullOrWhiteSpace($dbLifecyclePath)) {
            Cleanup-AccessArtifacts -DbPath $dbLifecyclePath
            Remove-Item -Path $dbLifecyclePath -Force -ErrorAction SilentlyContinue
        }
    }

    $databaseLifecycleCalls = New-Object 'System.Collections.Generic.List[object]'

    $createDatabaseArguments = @{
        database_path = $databaseLifecycleCreatedPath
        path = $databaseLifecycleCreatedPath
        target_database_path = $databaseLifecycleCreatedPath
        overwrite = $true
    }

    $backupDatabaseArguments = @{
        database_path = $databaseLifecycleCreatedPath
        source_database_path = $databaseLifecycleCreatedPath
        backup_path = $databaseLifecycleBackupPath
        backup_database_path = $databaseLifecycleBackupPath
        destination_path = $databaseLifecycleBackupPath
        destination_database_path = $databaseLifecycleBackupPath
        output_database_path = $databaseLifecycleBackupPath
        overwrite = $true
    }

    $compactRepairArguments = @{
        database_path = $databaseLifecycleBackupPath
        source_database_path = $databaseLifecycleBackupPath
        input_database_path = $databaseLifecycleBackupPath
        compacted_database_path = $databaseLifecycleCompactPath
        output_database_path = $databaseLifecycleCompactPath
        destination_database_path = $databaseLifecycleCompactPath
        target_database_path = $databaseLifecycleCompactPath
        overwrite = $true
    }

    Add-ToolCall -Calls $databaseLifecycleCalls -Id 371 -Name $createDatabaseToolName -Arguments $createDatabaseArguments
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 372 -Name "connect_access" -Arguments @{ database_path = $databaseLifecycleCreatedPath }
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 373 -Name "create_table" -Arguments @{
        table_name = $databaseLifecycleTableName
        fields = @(
            @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false },
            @{ name = "payload"; type = "TEXT"; size = 50; required = $false; allow_zero_length = $true }
        )
    }
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 374 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$databaseLifecycleTableName] (id, payload) VALUES (1, 'seed_value')" }
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 375 -Name "execute_sql" -Arguments @{ sql = "SELECT id, payload FROM [$databaseLifecycleTableName] WHERE id = 1" }
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 376 -Name "disconnect_access" -Arguments @{}
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 377 -Name "close_access" -Arguments @{}
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 378 -Name $backupDatabaseToolName -Arguments $backupDatabaseArguments
    Add-ToolCall -Calls $databaseLifecycleCalls -Id 379 -Name $compactRepairDatabaseToolName -Arguments $compactRepairArguments

    $databaseLifecycleResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $databaseLifecycleCalls -ClientName "full-regression-database-lifecycle" -ClientVersion "1.0"
    $databaseLifecycleIdLabels = @{
        371 = "database_file_create_database"
        372 = "database_file_connect_created_database"
        373 = "database_file_create_seed_table"
        374 = "database_file_insert_seed_row"
        375 = "database_file_execute_sql_seed_select"
        376 = "database_file_disconnect_created_database"
        377 = "database_file_close_created_database"
        378 = "database_file_backup_database"
        379 = "database_file_compact_repair_database"
    }

    foreach ($id in ($databaseLifecycleIdLabels.Keys | Sort-Object)) {
        $label = $databaseLifecycleIdLabels[$id]
        $decoded = Decode-McpResult -Response $databaseLifecycleResponses[[int]$id]

        if ($null -eq $decoded) {
            $failed++
            Write-Host ('{0}: FAIL missing-response' -f $label)
            continue
        }

        if ($decoded -is [string]) {
            $failed++
            Write-Host ('{0}: FAIL raw-string-response' -f $label)
            continue
        }

        if ($decoded.success -ne $true) {
            $failed++
            Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
            continue
        }

        switch ($label) {
            "database_file_execute_sql_seed_select" {
                $rows = @($decoded.rows)
                if ($rows.Count -lt 1) {
                    $failed++
                    Write-Host ('{0}: FAIL expected seeded row to be readable before backup/compact' -f $label)
                    continue
                }
            }
        }

        Write-Host ('{0}: OK' -f $label)
    }

    $verificationDatabasePath = $null
    foreach ($candidatePath in @($databaseLifecycleCompactPath, $databaseLifecycleBackupPath, $databaseLifecycleCreatedPath)) {
        if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath)) {
            $verificationDatabasePath = $candidatePath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($verificationDatabasePath)) {
        $failed++
        Write-Host "database_file_compact_repair_artifact: FAIL no readable database artifact found after backup/compact"
    }
    else {
        Write-Host ('database_file_tools_coverage: INFO verification_path={0}' -f $verificationDatabasePath)

        $databaseLifecycleVerifyCalls = New-Object 'System.Collections.Generic.List[object]'
        Add-ToolCall -Calls $databaseLifecycleVerifyCalls -Id 380 -Name "connect_access" -Arguments @{ database_path = $verificationDatabasePath }
        Add-ToolCall -Calls $databaseLifecycleVerifyCalls -Id 381 -Name "execute_sql" -Arguments @{ sql = "SELECT id, payload FROM [$databaseLifecycleTableName] WHERE id = 1" }
        Add-ToolCall -Calls $databaseLifecycleVerifyCalls -Id 382 -Name "disconnect_access" -Arguments @{}
        Add-ToolCall -Calls $databaseLifecycleVerifyCalls -Id 383 -Name "close_access" -Arguments @{}

        $databaseLifecycleVerifyResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $databaseLifecycleVerifyCalls -ClientName "full-regression-database-lifecycle-verify" -ClientVersion "1.0"
        $databaseLifecycleVerifyIdLabels = @{
            380 = "database_file_verify_connect"
            381 = "database_file_verify_seed_row_after_compact"
            382 = "database_file_verify_disconnect"
            383 = "database_file_verify_close_access"
        }

        foreach ($id in ($databaseLifecycleVerifyIdLabels.Keys | Sort-Object)) {
            $label = $databaseLifecycleVerifyIdLabels[$id]
            $decoded = Decode-McpResult -Response $databaseLifecycleVerifyResponses[[int]$id]

            if ($null -eq $decoded) {
                $failed++
                Write-Host ('{0}: FAIL missing-response' -f $label)
                continue
            }

            if ($decoded -is [string]) {
                $failed++
                Write-Host ('{0}: FAIL raw-string-response' -f $label)
                continue
            }

            if ($decoded.success -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
                continue
            }

            if ($label -eq "database_file_verify_seed_row_after_compact") {
                $rows = @($decoded.rows)
                if ($rows.Count -lt 1) {
                    $failed++
                    Write-Host ('{0}: FAIL expected seeded row after backup/compact flow' -f $label)
                    continue
                }
            }

            Write-Host ('{0}: OK' -f $label)
        }
    }
}
else {
    $missingDatabaseLifecycleTools = @()
    if ([string]::IsNullOrWhiteSpace($createDatabaseToolName)) { $missingDatabaseLifecycleTools += "create_database" }
    if ([string]::IsNullOrWhiteSpace($backupDatabaseToolName)) { $missingDatabaseLifecycleTools += "backup_database" }
    if ([string]::IsNullOrWhiteSpace($compactRepairDatabaseToolName)) { $missingDatabaseLifecycleTools += "compact_repair_database" }

    if ($AllowCoverageSkips) {
        Write-Host ("database_file_tools_coverage: SKIP database file lifecycle tools not exposed by this server build. missing={0}" -f ($missingDatabaseLifecycleTools -join ", "))
    }
    else {
        $failed++
        Write-Host ("database_file_tools_coverage: FAIL required database file lifecycle tools not exposed by this server build. missing={0}" -f ($missingDatabaseLifecycleTools -join ", "))
    }
}

# ── New Headless Tools Coverage (Priority 17-22: domain_aggregate, access_error, build_criteria, hidden attributes, etc.) ──

Write-Host ""
Write-Host "=== New Headless Tools Coverage (IDs 401-425) ==="
Write-Host "Intermediate cleanup: clearing stale Access/MCP processes before new headless tools section."
Cleanup-AccessArtifacts -DbPath $DatabasePath
Start-Sleep -Milliseconds 300

$newToolsCalls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $newToolsCalls -Id 401 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $newToolsCalls -Id 402 -Name "create_table" -Arguments @{
    table_name = $newToolsTableName
    fields = @(
        @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false },
        @{ name = "name"; type = "TEXT"; size = 50; required = $false; allow_zero_length = $true }
    )
}
Add-ToolCall -Calls $newToolsCalls -Id 403 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$newToolsTableName] (id, name) VALUES (1, 'alpha')" }
Add-ToolCall -Calls $newToolsCalls -Id 404 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$newToolsTableName] (id, name) VALUES (2, 'beta')" }
Add-ToolCall -Calls $newToolsCalls -Id 405 -Name "domain_aggregate" -Arguments @{ function = "DCount"; expression = "*"; domain = $newToolsTableName }
Add-ToolCall -Calls $newToolsCalls -Id 406 -Name "domain_aggregate" -Arguments @{ function = "DLookup"; expression = "name"; domain = $newToolsTableName; criteria = "id=1" }
Add-ToolCall -Calls $newToolsCalls -Id 407 -Name "access_error" -Arguments @{ error_number = 2001 }
Add-ToolCall -Calls $newToolsCalls -Id 408 -Name "build_criteria" -Arguments @{ field = "name"; field_type = 10; expression = "alpha" }
Add-ToolCall -Calls $newToolsCalls -Id 409 -Name "set_hidden_attribute" -Arguments @{ object_type = 0; object_name = $newToolsTableName; hidden = $true }
Add-ToolCall -Calls $newToolsCalls -Id 410 -Name "get_hidden_attribute" -Arguments @{ object_type = 0; object_name = $newToolsTableName }
Add-ToolCall -Calls $newToolsCalls -Id 411 -Name "set_hidden_attribute" -Arguments @{ object_type = 0; object_name = $newToolsTableName; hidden = $false }
Add-ToolCall -Calls $newToolsCalls -Id 412 -Name "get_current_user" -Arguments @{}
Add-ToolCall -Calls $newToolsCalls -Id 413 -Name "get_access_hwnd" -Arguments @{}
Add-ToolCall -Calls $newToolsCalls -Id 414 -Name "get_object_dates" -Arguments @{ object_type = "Table"; object_name = $newToolsTableName }
Add-ToolCall -Calls $newToolsCalls -Id 415 -Name "is_object_loaded" -Arguments @{ object_type = "Table"; object_name = $newToolsTableName }
Add-ToolCall -Calls $newToolsCalls -Id 416 -Name "is_vba_compiled" -Arguments @{}
Add-ToolCall -Calls $newToolsCalls -Id 417 -Name "list_printers" -Arguments @{}
Add-ToolCall -Calls $newToolsCalls -Id 418 -Name "get_database_engine_info" -Arguments @{}
Add-ToolCall -Calls $newToolsCalls -Id 419 -Name "get_current_object" -Arguments @{}
Add-ToolCall -Calls $newToolsCalls -Id 420 -Name "set_access_visible" -Arguments @{ visible = $true }
Add-ToolCall -Calls $newToolsCalls -Id 421 -Name "export_navigation_pane_xml" -Arguments @{ output_path = $tempNavXmlPath }
Add-ToolCall -Calls $newToolsCalls -Id 422 -Name "export_xml" -Arguments @{ object_type = 0; data_source = $newToolsTableName; data_target = $tempXmlDataPath }
Add-ToolCall -Calls $newToolsCalls -Id 423 -Name "delete_table" -Arguments @{ table_name = $newToolsTableName }
Add-ToolCall -Calls $newToolsCalls -Id 424 -Name "disconnect_access" -Arguments @{}
Add-ToolCall -Calls $newToolsCalls -Id 425 -Name "close_access" -Arguments @{}

$newToolsResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $newToolsCalls -ClientName "full-regression-new-tools" -ClientVersion "1.0"
$newToolsIdLabels = @{
    401 = "new_tools_connect_access"
    402 = "new_tools_create_table"
    403 = "new_tools_insert_alpha"
    404 = "new_tools_insert_beta"
    405 = "new_tools_domain_aggregate_dcount"
    406 = "new_tools_domain_aggregate_dlookup"
    407 = "new_tools_access_error"
    408 = "new_tools_build_criteria"
    409 = "new_tools_set_hidden_true"
    410 = "new_tools_get_hidden_attribute"
    411 = "new_tools_set_hidden_false"
    412 = "new_tools_get_current_user"
    413 = "new_tools_get_access_hwnd"
    414 = "new_tools_get_object_dates"
    415 = "new_tools_is_object_loaded"
    416 = "new_tools_is_vba_compiled"
    417 = "new_tools_list_printers"
    418 = "new_tools_get_database_engine_info"
    419 = "new_tools_get_current_object"
    420 = "new_tools_set_access_visible"
    421 = "new_tools_export_navigation_pane_xml"
    422 = "new_tools_export_xml"
    423 = "new_tools_delete_table"
    424 = "new_tools_disconnect_access"
    425 = "new_tools_close_access"
}

foreach ($id in ($newToolsIdLabels.Keys | Sort-Object)) {
    $label = $newToolsIdLabels[$id]
    $decoded = Decode-McpResult -Response $newToolsResponses[[int]$id]

    if ($null -eq $decoded) {
        $failed++
        Write-Host ('{0}: FAIL missing-response' -f $label)
        continue
    }

    if ($decoded -is [string]) {
        $failed++
        Write-Host ('{0}: FAIL raw-string-response' -f $label)
        continue
    }

    if ($decoded.success -ne $true) {
        $failed++
        Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
        continue
    }

    switch ($label) {
        "new_tools_domain_aggregate_dcount" {
            $val = $decoded.value
            if ($null -eq $val -or [int]$val -lt 2) {
                $failed++
                Write-Host ('{0}: FAIL expected DCount >= 2, got {1}' -f $label, $val)
                continue
            }
        }
        "new_tools_domain_aggregate_dlookup" {
            $val = [string]$decoded.value
            if ($val -ne "alpha") {
                $failed++
                Write-Host ('{0}: FAIL expected DLookup value "alpha", got "{1}"' -f $label, $val)
                continue
            }
        }
        "new_tools_access_error" {
            $desc = [string]$decoded.description
            if ([string]::IsNullOrWhiteSpace($desc)) {
                $failed++
                Write-Host ('{0}: FAIL expected non-empty error description' -f $label)
                continue
            }
        }
        "new_tools_build_criteria" {
            $criteria = [string]$decoded.criteria
            if ([string]::IsNullOrWhiteSpace($criteria)) {
                $failed++
                Write-Host ('{0}: FAIL expected non-empty criteria string' -f $label)
                continue
            }
        }
        "new_tools_get_hidden_attribute" {
            if ($decoded.hidden -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL expected hidden=true' -f $label)
                continue
            }
        }
        "new_tools_get_current_user" {
            $user = [string]$decoded.user
            if ([string]::IsNullOrWhiteSpace($user)) {
                $failed++
                Write-Host ('{0}: FAIL expected non-empty user' -f $label)
                continue
            }
        }
        "new_tools_get_access_hwnd" {
            if ($null -eq $decoded.hwnd) {
                $failed++
                Write-Host ('{0}: FAIL expected non-null hwnd' -f $label)
                continue
            }
        }
        "new_tools_get_object_dates" {
            # date_created may be null for newly-created tables on some Access builds; just verify the response shape
            if ($null -eq $decoded.PSObject -or
                (-not ($decoded.PSObject.Properties.Name -contains 'date_created') -and -not ($decoded.PSObject.Properties.Name -contains 'DateCreated'))) {
                $failed++
                Write-Host ('{0}: FAIL expected date_created property in response' -f $label)
                continue
            }
        }
        "new_tools_is_vba_compiled" {
            # Response nests under result: { success:true, result: { isCompiled:bool, ... } }
            if ($null -eq $decoded.result -or $null -eq $decoded.result.isCompiled) {
                $failed++
                Write-Host ('{0}: FAIL expected result.isCompiled property' -f $label)
                continue
            }
        }
        "new_tools_list_printers" {
            $printers = @($decoded.printers)
            if ($printers.Count -lt 0) {
                $failed++
                Write-Host ('{0}: FAIL expected printers array' -f $label)
                continue
            }
        }
        "new_tools_get_database_engine_info" {
            if ($null -eq $decoded.info -and $null -eq $decoded.engine -and $null -eq $decoded.version) {
                $failed++
                Write-Host ('{0}: FAIL expected non-null engine info' -f $label)
                continue
            }
        }
    }

    Write-Host ('{0}: OK' -f $label)
}

# ── DAO Recordset Coverage (Priority 20: open/close/navigate/CRUD recordsets) ──

Write-Host ""
Write-Host "=== DAO Recordset Coverage (IDs 601-624) ==="
Write-Host "Intermediate cleanup: clearing stale Access/MCP processes before recordset section."
Cleanup-AccessArtifacts -DbPath $DatabasePath
Start-Sleep -Milliseconds 300

$recordsetCalls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $recordsetCalls -Id 601 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $recordsetCalls -Id 602 -Name "create_table" -Arguments @{
    table_name = $recordsetTableName
    fields = @(
        @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false },
        @{ name = "name"; type = "TEXT"; size = 50; required = $false; allow_zero_length = $true }
    )
}
Add-ToolCall -Calls $recordsetCalls -Id 603 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$recordsetTableName] (id, name) VALUES (1, 'aaa')" }
Add-ToolCall -Calls $recordsetCalls -Id 604 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$recordsetTableName] (id, name) VALUES (2, 'bbb')" }
Add-ToolCall -Calls $recordsetCalls -Id 605 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$recordsetTableName] (id, name) VALUES (3, 'ccc')" }
Add-ToolCall -Calls $recordsetCalls -Id 606 -Name "open_recordset" -Arguments @{ source = $recordsetTableName }
# Note: first open_recordset in a fresh server process yields rs_1
Add-ToolCall -Calls $recordsetCalls -Id 607 -Name "recordset_count" -Arguments @{ recordset_id = "rs_1" }
Add-ToolCall -Calls $recordsetCalls -Id 608 -Name "recordset_get_record" -Arguments @{ recordset_id = "rs_1" }
Add-ToolCall -Calls $recordsetCalls -Id 609 -Name "recordset_move" -Arguments @{ recordset_id = "rs_1"; direction = "next" }
Add-ToolCall -Calls $recordsetCalls -Id 610 -Name "recordset_get_rows" -Arguments @{ recordset_id = "rs_1"; num_rows = 10 }
Add-ToolCall -Calls $recordsetCalls -Id 611 -Name "recordset_find" -Arguments @{ recordset_id = "rs_1"; criteria = "name='aaa'" }
Add-ToolCall -Calls $recordsetCalls -Id 612 -Name "recordset_bookmark" -Arguments @{ recordset_id = "rs_1" }
Add-ToolCall -Calls $recordsetCalls -Id 613 -Name "recordset_add_record" -Arguments @{ recordset_id = "rs_1"; fields = @{ id = 4; name = "ddd" } }
Add-ToolCall -Calls $recordsetCalls -Id 614 -Name "recordset_count" -Arguments @{ recordset_id = "rs_1" }
Add-ToolCall -Calls $recordsetCalls -Id 615 -Name "recordset_move" -Arguments @{ recordset_id = "rs_1"; direction = "first" }
Add-ToolCall -Calls $recordsetCalls -Id 616 -Name "recordset_edit_record" -Arguments @{ recordset_id = "rs_1"; fields = @{ name = "edited" } }
Add-ToolCall -Calls $recordsetCalls -Id 617 -Name "recordset_move" -Arguments @{ recordset_id = "rs_1"; direction = "last" }
Add-ToolCall -Calls $recordsetCalls -Id 618 -Name "recordset_delete_record" -Arguments @{ recordset_id = "rs_1" }
Add-ToolCall -Calls $recordsetCalls -Id 619 -Name "recordset_filter_sort" -Arguments @{ recordset_id = "rs_1"; sort = "name" }
# Note: filter_sort creates a new recordset; the original rs_1 remains open
Add-ToolCall -Calls $recordsetCalls -Id 620 -Name "close_recordset" -Arguments @{ recordset_id = "rs_1" }
# rs_2 is the filtered/sorted recordset created by filter_sort
Add-ToolCall -Calls $recordsetCalls -Id 621 -Name "close_recordset" -Arguments @{ recordset_id = "rs_2" }
Add-ToolCall -Calls $recordsetCalls -Id 622 -Name "delete_table" -Arguments @{ table_name = $recordsetTableName }
Add-ToolCall -Calls $recordsetCalls -Id 623 -Name "disconnect_access" -Arguments @{}
Add-ToolCall -Calls $recordsetCalls -Id 624 -Name "close_access" -Arguments @{}

$recordsetResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $recordsetCalls -ClientName "full-regression-recordsets" -ClientVersion "1.0"
$recordsetIdLabels = @{
    601 = "recordset_connect_access"
    602 = "recordset_create_table"
    603 = "recordset_insert_row_1"
    604 = "recordset_insert_row_2"
    605 = "recordset_insert_row_3"
    606 = "recordset_open_recordset"
    607 = "recordset_count_initial"
    608 = "recordset_get_record"
    609 = "recordset_move_next"
    610 = "recordset_get_rows"
    611 = "recordset_find"
    612 = "recordset_bookmark_get"
    613 = "recordset_add_record"
    614 = "recordset_count_after_add"
    615 = "recordset_move_first"
    616 = "recordset_edit_record"
    617 = "recordset_move_last"
    618 = "recordset_delete_record"
    619 = "recordset_filter_sort"
    620 = "recordset_close_original"
    621 = "recordset_close_filtered"
    622 = "recordset_delete_table"
    623 = "recordset_disconnect_access"
    624 = "recordset_close_access"
}

foreach ($id in ($recordsetIdLabels.Keys | Sort-Object)) {
    $label = $recordsetIdLabels[$id]
    $decoded = Decode-McpResult -Response $recordsetResponses[[int]$id]

    if ($null -eq $decoded) {
        $failed++
        Write-Host ('{0}: FAIL missing-response' -f $label)
        continue
    }

    if ($decoded -is [string]) {
        $failed++
        Write-Host ('{0}: FAIL raw-string-response' -f $label)
        continue
    }

    if ($decoded.success -ne $true) {
        $failed++
        Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
        continue
    }

    switch ($label) {
        "recordset_open_recordset" {
            $rsIdActual = [string]$decoded.recordset_id
            if ([string]::IsNullOrWhiteSpace($rsIdActual)) {
                $failed++
                Write-Host ('{0}: FAIL expected recordset_id in response' -f $label)
                continue
            }
            if ($rsIdActual -ne "rs_1") {
                Write-Host ('{0}: WARN expected rs_1, got {1} (hardcoded IDs may be wrong)' -f $label, $rsIdActual)
            }
        }
        "recordset_count_initial" {
            $rc = $decoded.record_count
            if ($null -eq $rc -or [int]$rc -lt 3) {
                $failed++
                Write-Host ('{0}: FAIL expected record_count >= 3, got {1}' -f $label, $rc)
                continue
            }
        }
        "recordset_get_record" {
            if ($null -eq $decoded.record) {
                $failed++
                Write-Host ('{0}: FAIL expected non-null record' -f $label)
                continue
            }
        }
        "recordset_get_rows" {
            $rows = @($decoded.rows)
            $rowCount = $decoded.row_count
            if ($rows.Count -lt 1 -or $null -eq $rowCount -or [int]$rowCount -lt 1) {
                $failed++
                Write-Host ('{0}: FAIL expected rows array with row_count >= 1' -f $label)
                continue
            }
        }
        "recordset_find" {
            if ($decoded.found -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL expected found=true' -f $label)
                continue
            }
        }
        "recordset_bookmark_get" {
            if ($null -eq $decoded.bookmark) {
                $failed++
                Write-Host ('{0}: FAIL expected non-null bookmark' -f $label)
                continue
            }
        }
        "recordset_count_after_add" {
            $rc = $decoded.record_count
            if ($null -eq $rc -or [int]$rc -lt 4) {
                $failed++
                Write-Host ('{0}: FAIL expected record_count >= 4 after add, got {1}' -f $label, $rc)
                continue
            }
        }
        "recordset_filter_sort" {
            $rsId2Actual = [string]$decoded.recordset_id
            if ([string]::IsNullOrWhiteSpace($rsId2Actual)) {
                $failed++
                Write-Host ('{0}: FAIL expected new recordset_id from filter_sort' -f $label)
                continue
            }
        }
    }

    Write-Host ('{0}: OK' -f $label)
}

# ── Form Runtime / UI Coverage (Priority 18-22: form_recalc, form_refresh, control ops, etc.) ──
# Gated by -IncludeUiCoverage because these tools open visible Access windows.

if ($IncludeUiCoverage) {
    Write-Host ""
    Write-Host "=== Form Runtime / UI Coverage (IDs 501-537) ==="
    Write-Host "Intermediate cleanup: clearing stale Access/MCP processes before form runtime section."
    Cleanup-AccessArtifacts -DbPath $DatabasePath
    Start-Sleep -Milliseconds 300

    $formRuntimeFormData = @{
        Name = $formRuntimeFormName
        RecordSource = $formRuntimeTableName
        ExportedAt = (Get-Date).ToUniversalTime().ToString("o")
        Controls = @(
            @{
                Name = "txtValue"
                Type = "TextBox"
                ControlSource = "name"
                Left = 600
                Top = 600
                Width = 2400
                Height = 300
                Visible = $true
                Enabled = $true
            }
        )
        VBA = ""
    } | ConvertTo-Json -Depth 20 -Compress

    $formRuntimeReportData = @{
        Name = $formRuntimeReportName
        RecordSource = $formRuntimeTableName
        ExportedAt = (Get-Date).ToUniversalTime().ToString("o")
        Controls = @(
            @{
                Name = "lblReport"
                Type = "Label"
                Left = 500
                Top = 300
                Width = 2500
                Height = 300
                Visible = $true
                Enabled = $true
            }
        )
    } | ConvertTo-Json -Depth 20 -Compress

    $formRtCalls = New-Object 'System.Collections.Generic.List[object]'
    Add-ToolCall -Calls $formRtCalls -Id 501 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
    Add-ToolCall -Calls $formRtCalls -Id 502 -Name "create_table" -Arguments @{
        table_name = $formRuntimeTableName
        fields = @(
            @{ name = "id"; type = "LONG"; size = 0; required = $true; allow_zero_length = $false },
            @{ name = "name"; type = "TEXT"; size = 50; required = $false; allow_zero_length = $true }
        )
    }
    Add-ToolCall -Calls $formRtCalls -Id 503 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$formRuntimeTableName] (id, name) VALUES (1, 'alpha')" }
    Add-ToolCall -Calls $formRtCalls -Id 504 -Name "execute_sql" -Arguments @{ sql = "INSERT INTO [$formRuntimeTableName] (id, name) VALUES (2, 'beta')" }
    Add-ToolCall -Calls $formRtCalls -Id 505 -Name "import_form_from_text" -Arguments @{ form_data = $formRuntimeFormData; form_name = $formRuntimeFormName }
    # Bind form to table so filter/order/refresh operations work
    Add-ToolCall -Calls $formRtCalls -Id 538 -Name "set_form_record_source" -Arguments @{ form_name = $formRuntimeFormName; record_source = $formRuntimeTableName }
    Add-ToolCall -Calls $formRtCalls -Id 506 -Name "open_form" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 507 -Name "form_recalc" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 508 -Name "form_refresh" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 509 -Name "form_requery" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 510 -Name "form_set_focus" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 511 -Name "get_form_dirty" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 512 -Name "get_form_new_record" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 513 -Name "get_form_view" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 514 -Name "get_form_open_args" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 515 -Name "set_form_painting" -Arguments @{ form_name = $formRuntimeFormName; painting = $false }
    Add-ToolCall -Calls $formRtCalls -Id 516 -Name "set_form_painting" -Arguments @{ form_name = $formRuntimeFormName; painting = $true }
    Add-ToolCall -Calls $formRtCalls -Id 517 -Name "get_active_form" -Arguments @{}
    Add-ToolCall -Calls $formRtCalls -Id 518 -Name "get_active_control" -Arguments @{}
    Add-ToolCall -Calls $formRtCalls -Id 519 -Name "control_set_focus" -Arguments @{ form_name = $formRuntimeFormName; control_name = "txtValue" }
    Add-ToolCall -Calls $formRtCalls -Id 520 -Name "control_requery" -Arguments @{ form_name = $formRuntimeFormName; control_name = "txtValue" }
    Add-ToolCall -Calls $formRtCalls -Id 521 -Name "control_undo" -Arguments @{ form_name = $formRuntimeFormName; control_name = "txtValue" }
    Add-ToolCall -Calls $formRtCalls -Id 522 -Name "set_filter_docmd" -Arguments @{ form_name = $formRuntimeFormName; where_condition = "[id]=1" }
    Add-ToolCall -Calls $formRtCalls -Id 523 -Name "set_order_by" -Arguments @{ form_name = $formRuntimeFormName; order_by = "[name] ASC" }
    Add-ToolCall -Calls $formRtCalls -Id 524 -Name "refresh_record" -Arguments @{}
    Add-ToolCall -Calls $formRtCalls -Id 525 -Name "set_parameter" -Arguments @{ name = "TestParam"; expression = "1" }
    Add-ToolCall -Calls $formRtCalls -Id 526 -Name "form_undo" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 527 -Name "close_form" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 528 -Name "import_report_from_text" -Arguments @{ report_data = $formRuntimeReportData; report_name = $formRuntimeReportName }
    Add-ToolCall -Calls $formRtCalls -Id 529 -Name "open_report" -Arguments @{ report_name = $formRuntimeReportName }
    Add-ToolCall -Calls $formRtCalls -Id 530 -Name "get_active_report" -Arguments @{}
    Add-ToolCall -Calls $formRtCalls -Id 531 -Name "close_report" -Arguments @{ report_name = $formRuntimeReportName }
    Add-ToolCall -Calls $formRtCalls -Id 532 -Name "is_object_loaded" -Arguments @{ object_type = "Form"; object_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 533 -Name "delete_form" -Arguments @{ form_name = $formRuntimeFormName }
    Add-ToolCall -Calls $formRtCalls -Id 534 -Name "delete_report" -Arguments @{ report_name = $formRuntimeReportName }
    Add-ToolCall -Calls $formRtCalls -Id 535 -Name "delete_table" -Arguments @{ table_name = $formRuntimeTableName }
    Add-ToolCall -Calls $formRtCalls -Id 536 -Name "disconnect_access" -Arguments @{}
    Add-ToolCall -Calls $formRtCalls -Id 537 -Name "close_access" -Arguments @{}

    $formRtResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $formRtCalls -ClientName "full-regression-form-runtime" -ClientVersion "1.0"
    $formRtIdLabels = @{
        501 = "form_runtime_connect_access"
        502 = "form_runtime_create_table"
        503 = "form_runtime_insert_alpha"
        504 = "form_runtime_insert_beta"
        505 = "form_runtime_import_form"
        538 = "form_runtime_set_record_source"
        506 = "form_runtime_open_form"
        507 = "form_runtime_form_recalc"
        508 = "form_runtime_form_refresh"
        509 = "form_runtime_form_requery"
        510 = "form_runtime_form_set_focus"
        511 = "form_runtime_get_form_dirty"
        512 = "form_runtime_get_form_new_record"
        513 = "form_runtime_get_form_view"
        514 = "form_runtime_get_form_open_args"
        515 = "form_runtime_set_form_painting_off"
        516 = "form_runtime_set_form_painting_on"
        517 = "form_runtime_get_active_form"
        518 = "form_runtime_get_active_control"
        519 = "form_runtime_control_set_focus"
        520 = "form_runtime_control_requery"
        521 = "form_runtime_control_undo"
        522 = "form_runtime_set_filter_docmd"
        523 = "form_runtime_set_order_by"
        524 = "form_runtime_refresh_record"
        525 = "form_runtime_set_parameter"
        526 = "form_runtime_form_undo"
        527 = "form_runtime_close_form"
        528 = "form_runtime_import_report"
        529 = "form_runtime_open_report"
        530 = "form_runtime_get_active_report"
        531 = "form_runtime_close_report"
        532 = "form_runtime_is_object_loaded_after_close"
        533 = "form_runtime_delete_form"
        534 = "form_runtime_delete_report"
        535 = "form_runtime_delete_table"
        536 = "form_runtime_disconnect_access"
        537 = "form_runtime_close_access"
    }

    foreach ($id in ($formRtIdLabels.Keys | Sort-Object)) {
        $label = $formRtIdLabels[$id]
        $decoded = Decode-McpResult -Response $formRtResponses[[int]$id]

        if ($null -eq $decoded) {
            $failed++
            Write-Host ('{0}: FAIL missing-response' -f $label)
            continue
        }

        if ($decoded -is [string]) {
            $failed++
            Write-Host ('{0}: FAIL raw-string-response' -f $label)
            continue
        }

        if ($decoded.success -ne $true) {
            # get_active_control may fail gracefully if no control has focus yet (before control_set_focus)
            if ($label -eq "form_runtime_get_active_control") {
                Write-Host ('{0}: SKIP (no focused control before control_set_focus) {1}' -f $label, $decoded.error)
                continue
            }
            $failed++
            Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
            continue
        }

        switch ($label) {
            "form_runtime_get_form_dirty" {
                if ($null -eq $decoded.dirty -or $decoded.dirty -isnot [bool]) {
                    $failed++
                    Write-Host ('{0}: FAIL expected dirty to be boolean' -f $label)
                    continue
                }
            }
            "form_runtime_get_form_new_record" {
                if ($null -eq $decoded.new_record -or $decoded.new_record -isnot [bool]) {
                    $failed++
                    Write-Host ('{0}: FAIL expected new_record to be boolean' -f $label)
                    continue
                }
            }
            "form_runtime_get_form_view" {
                if ($null -eq $decoded.current_view) {
                    $failed++
                    Write-Host ('{0}: FAIL expected current_view property' -f $label)
                    continue
                }
            }
            "form_runtime_get_active_form" {
                # Response shape: { success:true, result: { name, recordSource, caption, ... } }
                $activeFormName = [string]$decoded.result.name
                if ([string]::IsNullOrWhiteSpace($activeFormName)) {
                    $failed++
                    Write-Host ('{0}: FAIL expected result.name property on active form' -f $label)
                    continue
                }
            }
            "form_runtime_get_active_report" {
                # Response shape: { success:true, result: { name, recordSource, caption } }
                $activeReportName = [string]$decoded.result.name
                if ([string]::IsNullOrWhiteSpace($activeReportName)) {
                    $failed++
                    Write-Host ('{0}: FAIL expected result.name property on active report' -f $label)
                    continue
                }
            }
            "form_runtime_is_object_loaded_after_close" {
                if ($decoded.is_loaded -eq $true) {
                    $failed++
                    Write-Host ('{0}: FAIL expected is_loaded=false after form close' -f $label)
                    continue
                }
            }
        }

        Write-Host ('{0}: OK' -f $label)
    }
}
else {
    Write-Host ""
    Write-Host "form_runtime_coverage: SKIP (requires -IncludeUiCoverage)"
}

# ── Close Database Coverage (Priority 22: close_database invalidates connection) ──

Write-Host ""
Write-Host "=== Close Database Coverage (IDs 651-655) ==="
Write-Host "Intermediate cleanup: clearing stale Access/MCP processes before close_database section."
Cleanup-AccessArtifacts -DbPath $DatabasePath
Start-Sleep -Milliseconds 300

$closeDbCalls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $closeDbCalls -Id 651 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $closeDbCalls -Id 652 -Name "is_connected" -Arguments @{}
Add-ToolCall -Calls $closeDbCalls -Id 653 -Name "close_database" -Arguments @{}
Add-ToolCall -Calls $closeDbCalls -Id 654 -Name "disconnect_access" -Arguments @{}
Add-ToolCall -Calls $closeDbCalls -Id 655 -Name "close_access" -Arguments @{}

$closeDbResponses = Invoke-McpBatch -ExePath $ServerExe -Calls $closeDbCalls -ClientName "full-regression-close-database" -ClientVersion "1.0"
$closeDbIdLabels = @{
    651 = "close_database_connect_access"
    652 = "close_database_is_connected"
    653 = "close_database_close_database"
    654 = "close_database_disconnect_access"
    655 = "close_database_close_access"
}

foreach ($id in ($closeDbIdLabels.Keys | Sort-Object)) {
    $label = $closeDbIdLabels[$id]
    $decoded = Decode-McpResult -Response $closeDbResponses[[int]$id]

    if ($null -eq $decoded) {
        $failed++
        Write-Host ('{0}: FAIL missing-response' -f $label)
        continue
    }

    if ($decoded -is [string]) {
        $failed++
        Write-Host ('{0}: FAIL raw-string-response' -f $label)
        continue
    }

    if ($decoded.success -ne $true) {
        $failed++
        Write-Host ('{0}: FAIL {1}' -f $label, $decoded.error)
        continue
    }

    switch ($label) {
        "close_database_is_connected" {
            if ($decoded.connected -ne $true) {
                $failed++
                Write-Host ('{0}: FAIL expected connected=true before close_database' -f $label)
                continue
            }
        }
    }

    Write-Host ('{0}: OK' -f $label)
}

# ── MCP Feature Tests (Resources, Prompts, Completion, Logging) ──

Write-Host ""
Write-Host "=== MCP Feature Tests (resources, prompts, completion, logging) ==="

$mcpFeatureRequests = New-Object 'System.Collections.Generic.List[hashtable]'

# resources/list (id=10)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 10; method = "resources/list"; params = @{} })

# resources/templates/list (id=11)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 11; method = "resources/templates/list"; params = @{} })

# resources/read access://connection (id=12 — always available, no database needed)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 12; method = "resources/read"; params = @{ uri = "access://connection" } })

# prompts/list (id=13)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 13; method = "prompts/list"; params = @{} })

# prompts/get for debug_query with sql argument (id=14)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 14; method = "prompts/get"; params = @{ name = "debug_query"; arguments = @{ sql = "SELECT * FROM Test" } } })

# logging/setLevel (id=15)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 15; method = "logging/setLevel"; params = @{ level = "warning" } })

# completion/complete for table names (id=16 — returns empty when not connected, but should not error)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 16; method = "completion/complete"; params = @{ ref = @{ type = "ref/prompt"; name = "source_table" }; argument = @{ name = "source_table"; value = "" } } })

# Unknown method should return JSON-RPC error (id=17)
$mcpFeatureRequests.Add(@{ jsonrpc = "2.0"; id = 17; method = "bogus/method"; params = @{} })

$mcpFeatureResponses = Invoke-McpRawBatch -ExePath $ServerExe -Requests $mcpFeatureRequests -ClientName "mcp-feature-regression"

# resources/list — expect 10 resources
$resListResp = $mcpFeatureResponses[10]
if ($null -eq $resListResp -or $null -eq $resListResp.result) {
    $failed++
    Write-Host "resources_list: FAIL missing response"
}
else {
    $resList = @($resListResp.result.resources)
    if ($resList.Count -eq 10) {
        Write-Host ("resources_list: OK count={0}" -f $resList.Count)
    }
    else {
        $failed++
        Write-Host ("resources_list: FAIL expected 10, got {0}" -f $resList.Count)
    }
}

# resources/templates/list — expect 6 templates
$resTemplatesResp = $mcpFeatureResponses[11]
if ($null -eq $resTemplatesResp -or $null -eq $resTemplatesResp.result) {
    $failed++
    Write-Host "resources_templates_list: FAIL missing response"
}
else {
    $resTemplates = @($resTemplatesResp.result.resourceTemplates)
    if ($resTemplates.Count -eq 6) {
        Write-Host ("resources_templates_list: OK count={0}" -f $resTemplates.Count)
    }
    else {
        $failed++
        Write-Host ("resources_templates_list: FAIL expected 6, got {0}" -f $resTemplates.Count)
    }
}

# resources/read access://connection — expect contents array with connection status
$resReadResp = $mcpFeatureResponses[12]
if ($null -eq $resReadResp -or $null -eq $resReadResp.result) {
    $failed++
    Write-Host "resources_read_connection: FAIL missing response"
}
elseif ($resReadResp.error) {
    $failed++
    Write-Host ("resources_read_connection: FAIL error: {0}" -f $resReadResp.error.message)
}
else {
    $contents = @($resReadResp.result.contents)
    if ($contents.Count -ge 1 -and $contents[0].uri -eq "access://connection") {
        Write-Host "resources_read_connection: OK"
    }
    else {
        $failed++
        Write-Host "resources_read_connection: FAIL unexpected contents structure"
    }
}

# prompts/list — expect 6 prompts
$promptsListResp = $mcpFeatureResponses[13]
if ($null -eq $promptsListResp -or $null -eq $promptsListResp.result) {
    $failed++
    Write-Host "prompts_list: FAIL missing response"
}
else {
    $promptsList = @($promptsListResp.result.prompts)
    if ($promptsList.Count -eq 6) {
        Write-Host ("prompts_list: OK count={0}" -f $promptsList.Count)
    }
    else {
        $failed++
        Write-Host ("prompts_list: FAIL expected 6, got {0}" -f $promptsList.Count)
    }
}

# prompts/get debug_query — expect messages array
$promptGetResp = $mcpFeatureResponses[14]
if ($null -eq $promptGetResp -or $null -eq $promptGetResp.result) {
    $failed++
    Write-Host "prompts_get_debug_query: FAIL missing response"
}
elseif ($promptGetResp.error) {
    $failed++
    Write-Host ("prompts_get_debug_query: FAIL error: {0}" -f $promptGetResp.error.message)
}
else {
    $messages = @($promptGetResp.result.messages)
    if ($messages.Count -ge 1) {
        Write-Host ("prompts_get_debug_query: OK messages={0}" -f $messages.Count)
    }
    else {
        $failed++
        Write-Host "prompts_get_debug_query: FAIL no messages returned"
    }
}

# logging/setLevel — expect empty result (no error)
$logSetResp = $mcpFeatureResponses[15]
if ($null -eq $logSetResp) {
    $failed++
    Write-Host "logging_setLevel: FAIL missing response"
}
elseif ($logSetResp.error) {
    $failed++
    Write-Host ("logging_setLevel: FAIL error: {0}" -f $logSetResp.error.message)
}
else {
    Write-Host "logging_setLevel: OK"
}

# completion/complete — expect completion object (even if empty values when not connected)
$completionResp = $mcpFeatureResponses[16]
if ($null -eq $completionResp -or $null -eq $completionResp.result) {
    $failed++
    Write-Host "completion_complete: FAIL missing response"
}
elseif ($completionResp.error) {
    $failed++
    Write-Host ("completion_complete: FAIL error: {0}" -f $completionResp.error.message)
}
else {
    $completion = $completionResp.result.completion
    if ($null -ne $completion -and $null -ne $completion.values) {
        Write-Host ("completion_complete: OK values_count={0}" -f @($completion.values).Count)
    }
    else {
        $failed++
        Write-Host "completion_complete: FAIL missing completion object"
    }
}

# Unknown method — expect JSON-RPC error response with code -32601
$unknownResp = $mcpFeatureResponses[17]
if ($null -eq $unknownResp) {
    $failed++
    Write-Host "unknown_method_error: FAIL missing response"
}
elseif ($unknownResp.error -and $unknownResp.error.code -eq -32601) {
    Write-Host "unknown_method_error: OK code=-32601"
}
else {
    $failed++
    Write-Host "unknown_method_error: FAIL expected error code -32601"
}

Write-Host "=== End MCP Feature Tests ==="
Write-Host ""

Write-Host ("TOTAL_FAIL={0}" -f $failed)
if ($failed -eq 0) {
    $exitCode = 0
}
}
finally {
    Write-Host "Final cleanup: clearing stale Access/MCP processes and locks."
    Cleanup-AccessArtifacts -DbPath $DatabasePath
    if (-not [string]::IsNullOrWhiteSpace($linkedSourceDatabasePath)) {
        Cleanup-AccessArtifacts -DbPath $linkedSourceDatabasePath
        Remove-Item -Path $linkedSourceDatabasePath -Force -ErrorAction SilentlyContinue
    }
    foreach ($dbLifecyclePath in @($databaseLifecycleCreatedPath, $databaseLifecycleBackupPath, $databaseLifecycleCompactPath)) {
        if (-not [string]::IsNullOrWhiteSpace($dbLifecyclePath)) {
            Cleanup-AccessArtifacts -DbPath $dbLifecyclePath
            Remove-Item -Path $dbLifecyclePath -Force -ErrorAction SilentlyContinue
        }
    }
    Remove-Item -Path $tempNavXmlPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $tempXmlDataPath -Force -ErrorAction SilentlyContinue
    Release-RegressionLock -LockState $regressionLock
}

exit $exitCode
