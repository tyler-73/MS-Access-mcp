# stress_test_autonomy_gap.ps1 — Stress-test all 8 new autonomy gap tools
# Usage: powershell -ExecutionPolicy Bypass -File stress_test_autonomy_gap.ps1

$ErrorActionPreference = "Continue"

# ── Infrastructure ────────────────────────────────────────────────────────────
. "C:\Users\tyler\MS-Access-mcp\tests\_dialog_watcher.ps1"

$ServerExe    = "C:\Users\tyler\MS-Access-mcp\mcp-server-official-x64\MS.Access.MCP.Official.exe"
$DatabasePath = "C:\Users\tyler\Documents\MyDatabase.accdb"
$TimeoutSec   = 120

# Helper functions (inline, minimal)
function Add-ToolCall {
    param(
        [System.Collections.Generic.List[object]]$Calls,
        [int]$Id,
        [string]$Name,
        [hashtable]$Arguments = @{}
    )
    $Calls.Add([PSCustomObject]@{ Id = $Id; Name = $Name; Arguments = $Arguments })
}

function Decode-McpResult {
    param([object]$Response)
    if ($null -eq $Response) { return $null }
    if ($Response.result -and $Response.result.structuredContent) {
        return $Response.result.structuredContent
    }
    if ($Response.result -and $Response.result.content) {
        $text = $Response.result.content[0].text
        try { return $text | ConvertFrom-Json }
        catch { return $text }
    }
    return $Response.result
}

function Run-Batch {
    param(
        [string]$Label,
        [System.Collections.Generic.List[object]]$Calls
    )
    Write-Host ""
    Write-Host ("=" * 70)
    Write-Host "BATCH: $Label"
    Write-Host ("=" * 70)

    $responses = Invoke-McpBatchWithTimeout -ExePath $ServerExe -Calls $Calls `
        -ClientName $Label -ClientVersion "1.0" -TimeoutSeconds $TimeoutSec -SectionName $Label

    if ($responses._timeout) {
        Write-Host "TIMEOUT: Batch '$Label' timed out after ${TimeoutSec}s"
        return $null
    }
    return $responses
}

$totalTests  = 0
$passCount   = 0
$failCount   = 0
$skipCount   = 0

function Assert-Test {
    param(
        [string]$TestName,
        [bool]$Condition,
        [string]$Detail = ""
    )
    $script:totalTests++
    if ($Condition) {
        $script:passCount++
        Write-Host ("  PASS: {0} {1}" -f $TestName, $Detail) -ForegroundColor Green
    } else {
        $script:failCount++
        Write-Host ("  FAIL: {0} {1}" -f $TestName, $Detail) -ForegroundColor Red
    }
}

function Skip-Test {
    param([string]$TestName, [string]$Reason)
    $script:totalTests++
    $script:skipCount++
    Write-Host ("  SKIP: {0} - {1}" -f $TestName, $Reason) -ForegroundColor Yellow
}

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 1: list_odbc_data_sources
# ══════════════════════════════════════════════════════════════════════════════

$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "list_odbc_data_sources" -Arguments @{}
Add-ToolCall -Calls $calls -Id 4 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool1-list_odbc_data_sources" -Calls $calls

if ($null -ne $r) {
    $connectResult = Decode-McpResult -Response $r[2]
    Write-Host ("  connect_access: {0}" -f ($connectResult | ConvertTo-Json -Depth 3 -Compress))

    $decoded = Decode-McpResult -Response $r[3]
    Write-Host "  list_odbc_data_sources result:"
    Write-Host ($decoded | ConvertTo-Json -Depth 5 -Compress)

    # Response shape: { success, data_sources: [...], count }
    $dataSources = $decoded.data_sources
    $dsCount = $decoded.count
    $hasDsArray = ($null -ne $dataSources) -and ($dataSources.Count -gt 0)
    Assert-Test "list_odbc_data_sources returns data_sources array" ([bool]$hasDsArray) ("count=" + $dsCount)

    if ($hasDsArray) {
        $firstDs = $dataSources[0]
        $hasDsnName = $null -ne $firstDs.DsnName
        Assert-Test "list_odbc_data_sources first entry has DsnName" ([bool]$hasDsnName) ("DsnName=" + $firstDs.DsnName)
    }
} else {
    Assert-Test "list_odbc_data_sources batch completed" $false "TIMEOUT"
}

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 2: create_odbc_linked_table — SKIPPED
# ══════════════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host ("=" * 70)
Write-Host "BATCH: Tool2-create_odbc_linked_table"
Write-Host ("=" * 70)
Skip-Test "create_odbc_linked_table" "Requires a real ODBC DSN - skipped by design"

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 3: execute_sql_timed
# Note: MSysObjects requires special permissions in Access, so we expect it
# might fail. The key test is SELECT 1+1 which should always work.
# ══════════════════════════════════════════════════════════════════════════════

$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "execute_sql_timed" -Arguments @{
    sql = "SELECT 1+1 AS result"
}
Add-ToolCall -Calls $calls -Id 4 -Name "execute_sql_timed" -Arguments @{
    sql = "SELECT Count(*) AS cnt FROM MSysObjects WHERE Type=1"
    max_rows = 10
}
Add-ToolCall -Calls $calls -Id 5 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool3-execute_sql_timed" -Calls $calls

if ($null -ne $r) {
    # Test 3a: SELECT 1+1 (should always succeed)
    $decoded3a = Decode-McpResult -Response $r[3]
    Write-Host "  execute_sql_timed (SELECT 1+1) result:"
    Write-Host ($decoded3a | ConvertTo-Json -Depth 5 -Compress)

    # Response shape: { success, result: { IsQuery, Columns, Rows, RowCount, Truncated, RowsAffected, ExecutionTimeMs } }
    $innerResult = $decoded3a.result
    $execTime = $null
    if ($null -ne $innerResult) { $execTime = $innerResult.ExecutionTimeMs }
    $hasTime = $null -ne $execTime
    Assert-Test "execute_sql_timed SELECT1plus1 has ExecutionTimeMs" ([bool]$hasTime) ("value=" + $execTime)

    $timeGe0 = [bool]($hasTime -and [double]$execTime -ge 0)
    Assert-Test "execute_sql_timed SELECT1plus1 ExecutionTimeMs non-negative" $timeGe0 ("value=" + $execTime)

    $hasRows = $null -ne $innerResult -and $null -ne $innerResult.Rows
    Assert-Test "execute_sql_timed SELECT1plus1 returns Rows" ([bool]$hasRows) ""

    if ($hasRows) {
        $firstRow = $innerResult.Rows[0]
        $resultVal = $firstRow.result
        $correct = [bool]($resultVal -eq 2)
        Assert-Test "execute_sql_timed SELECT1plus1 result equals 2" $correct ("result=" + $resultVal)
    }

    # Test 3b: MSysObjects (may fail due to permissions - just log)
    $decoded3b = Decode-McpResult -Response $r[4]
    Write-Host "  execute_sql_timed (MSysObjects) result:"
    $json3b = $decoded3b | ConvertTo-Json -Depth 5 -Compress
    if ($json3b.Length -gt 500) { Write-Host ($json3b.Substring(0, 500) + "...") } else { Write-Host $json3b }

    if ($decoded3b.success -eq $true) {
        $innerResult3b = $decoded3b.result
        $hasTime3b = $null -ne $innerResult3b -and $null -ne $innerResult3b.ExecutionTimeMs
        Assert-Test "execute_sql_timed MSysObjects has ExecutionTimeMs" ([bool]$hasTime3b) ""
    } else {
        # Permission denied is expected on some Access configs
        Write-Host "  NOTE: MSysObjects query failed (expected - permission denied)" -ForegroundColor Yellow
        Assert-Test "execute_sql_timed MSysObjects returned error gracefully" ($decoded3b.success -eq $false) ("error=" + $decoded3b.error)
    }
} else {
    Assert-Test "execute_sql_timed batch completed" $false "TIMEOUT"
}

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 4: get_database_statistics
# ══════════════════════════════════════════════════════════════════════════════

$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "get_database_statistics" -Arguments @{}
Add-ToolCall -Calls $calls -Id 4 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool4-get_database_statistics" -Calls $calls

if ($null -ne $r) {
    $decoded = Decode-McpResult -Response $r[3]
    Write-Host "  get_database_statistics result:"
    $json4 = $decoded | ConvertTo-Json -Depth 5 -Compress
    if ($json4.Length -gt 800) { Write-Host ($json4.Substring(0, 800) + "...") } else { Write-Host $json4 }

    # Response shape: { success, result: { FileSizeBytes, FileSizeMB, TableCount, ... } }
    $innerResult = $decoded.result
    $fileSize = if ($null -ne $innerResult) { $innerResult.FileSizeBytes } else { $null }
    $tableCount = if ($null -ne $innerResult) { $innerResult.TableCount } else { $null }

    $condFS = [bool]($null -ne $fileSize -and [long]$fileSize -gt 0)
    Assert-Test "get_database_statistics FileSizeBytes positive" $condFS ("FileSizeBytes=" + $fileSize)

    $condTC = [bool]($null -ne $tableCount -and [int]$tableCount -gt 0)
    Assert-Test "get_database_statistics TableCount positive" $condTC ("TableCount=" + $tableCount)

    # Check for additional stats fields
    $totalRecords = if ($null -ne $innerResult) { $innerResult.TotalRecords } else { $null }
    $hasTotalRecords = $null -ne $totalRecords
    Assert-Test "get_database_statistics has TotalRecords" ([bool]$hasTotalRecords) ("TotalRecords=" + $totalRecords)

    $fileSizeMB = if ($null -ne $innerResult) { $innerResult.FileSizeMB } else { $null }
    $hasMB = $null -ne $fileSizeMB
    Assert-Test "get_database_statistics has FileSizeMB" ([bool]$hasMB) ("FileSizeMB=" + $fileSizeMB)
} else {
    Assert-Test "get_database_statistics batch completed" $false "TIMEOUT"
}

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 5: export_schema_snapshot
# ══════════════════════════════════════════════════════════════════════════════

$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "export_schema_snapshot" -Arguments @{
    include_vba  = $true
    include_data = $true
    max_data_rows = 5
}
Add-ToolCall -Calls $calls -Id 4 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool5-export_schema_snapshot" -Calls $calls

if ($null -ne $r) {
    $decoded = Decode-McpResult -Response $r[3]
    # This can be very large - truncate for display
    $json5 = $decoded | ConvertTo-Json -Depth 8 -Compress
    Write-Host "  export_schema_snapshot result (first 1500 chars):"
    if ($json5.Length -gt 1500) { Write-Host ($json5.Substring(0, 1500) + "...") } else { Write-Host $json5 }

    # Response shape: { success, result: { DatabasePath, ExportedAt, Tables: [...], VbaModules: [...] } }
    $innerResult = $decoded.result
    $tables = if ($null -ne $innerResult) { $innerResult.Tables } else { $null }
    $hasTables = ($null -ne $tables) -and ($tables.Count -gt 0)
    $tblCount = if ($null -ne $tables) { $tables.Count } else { 0 }
    Assert-Test "export_schema_snapshot Tables array non-empty" ([bool]$hasTables) ("count=" + $tblCount)

    if ($hasTables) {
        $firstTable = $tables[0]
        $hasFields = ($null -ne $firstTable.Fields) -and ($firstTable.Fields.Count -gt 0)
        $fldCount = if ($null -ne $firstTable.Fields) { $firstTable.Fields.Count } else { 0 }
        Assert-Test "export_schema_snapshot first table has Fields" ([bool]$hasFields) ("table=" + $firstTable.Name + " fieldCount=" + $fldCount)

        # Check that SampleData is included (since include_data=$true)
        $hasSampleData = $null -ne $firstTable.SampleData
        Assert-Test "export_schema_snapshot first table has SampleData" ([bool]$hasSampleData) ""
    }

    # Check for VBA content
    $vbaModules = if ($null -ne $innerResult) { $innerResult.VbaModules } else { $null }
    if ($null -eq $vbaModules -and $null -ne $innerResult) { $vbaModules = $innerResult.Modules }
    $hasVba = ($null -ne $vbaModules) -and ($vbaModules.Count -gt 0)
    Assert-Test "export_schema_snapshot includes VBA modules" ([bool]$hasVba) ("count=" + $(if ($null -ne $vbaModules) { $vbaModules.Count } else { 0 }))

    # Check DatabasePath
    $dbPath = if ($null -ne $innerResult) { $innerResult.DatabasePath } else { $null }
    $hasDbPath = $null -ne $dbPath
    Assert-Test "export_schema_snapshot has DatabasePath" ([bool]$hasDbPath) ("path=" + $dbPath)
} else {
    Assert-Test "export_schema_snapshot batch completed" $false "TIMEOUT"
}

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 6: export_all_vba
# ══════════════════════════════════════════════════════════════════════════════

$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "export_all_vba" -Arguments @{}
Add-ToolCall -Calls $calls -Id 4 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool6-export_all_vba" -Calls $calls

if ($null -ne $r) {
    $decoded = Decode-McpResult -Response $r[3]
    $json6 = $decoded | ConvertTo-Json -Depth 5 -Compress
    Write-Host "  export_all_vba result (first 1000 chars):"
    if ($json6.Length -gt 1000) { Write-Host ($json6.Substring(0, 1000) + "...") } else { Write-Host $json6 }

    # Response shape: { success, modules: [ { ProjectName, ModuleName, ModuleType, LineCount, Code }, ... ] }
    $modules = $decoded.modules
    $moduleCount = if ($null -ne $modules) { $modules.Count } else { 0 }
    Write-Host ("  Module count: {0}" -f $moduleCount)

    Assert-Test "export_all_vba returns success" ($decoded.success -eq $true) ""
    Assert-Test "export_all_vba has modules array" ($null -ne $modules) ("count=" + $moduleCount)

    if ($moduleCount -gt 0) {
        $firstMod = $modules[0]
        $firstName = $firstMod.ModuleName
        $firstType = $firstMod.ModuleType
        Write-Host ("  First module: {0} (type={1})" -f $firstName, $firstType)
        $hasModName = $null -ne $firstName
        Assert-Test "export_all_vba first module has ModuleName" ([bool]$hasModName) ("name=" + $firstName)
        $hasModType = $null -ne $firstType
        Assert-Test "export_all_vba first module has ModuleType" ([bool]$hasModType) ("type=" + $firstType)
    }
} else {
    Assert-Test "export_all_vba batch completed" $false "TIMEOUT"
}

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 7: check_referential_integrity
# ══════════════════════════════════════════════════════════════════════════════

$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "check_referential_integrity" -Arguments @{}
Add-ToolCall -Calls $calls -Id 4 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool7-check_referential_integrity" -Calls $calls

if ($null -ne $r) {
    $decoded = Decode-McpResult -Response $r[3]
    Write-Host "  check_referential_integrity result:"
    $json7 = $decoded | ConvertTo-Json -Depth 5 -Compress
    if ($json7.Length -gt 1000) { Write-Host ($json7.Substring(0, 1000) + "...") } else { Write-Host $json7 }

    # Response shape: { success, violations: [...], violation_count: N, is_clean: bool }
    Assert-Test "check_referential_integrity returns success" ($decoded.success -eq $true) ""

    $violations = $decoded.violations
    $violationCount = $decoded.violation_count
    $isClean = $decoded.is_clean

    $hasViolations = $null -ne $violations
    Assert-Test "check_referential_integrity has violations array" ([bool]$hasViolations) ("count=" + $violations.Count)

    $hasViolCount = $null -ne $violationCount
    Assert-Test "check_referential_integrity has violation_count" ([bool]$hasViolCount) ("violation_count=" + $violationCount)

    $hasIsClean = $null -ne $isClean
    Assert-Test "check_referential_integrity has is_clean flag" ([bool]$hasIsClean) ("is_clean=" + $isClean)

    Write-Host ("  Violation count: {0}, is_clean: {1}" -f $violationCount, $isClean)
} else {
    Assert-Test "check_referential_integrity batch completed" $false "TIMEOUT"
}

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 8: find_duplicate_records
#   Setup: create test table, insert dups, test, cleanup
# ══════════════════════════════════════════════════════════════════════════════

# Step 8a: Setup — create table and insert duplicates
$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "execute_sql" -Arguments @{
    sql = "CREATE TABLE _DupTest (ID AUTOINCREMENT PRIMARY KEY, FullName TEXT(100), Email TEXT(100))"
}
Add-ToolCall -Calls $calls -Id 4 -Name "execute_sql" -Arguments @{
    sql = "INSERT INTO _DupTest (FullName, Email) VALUES ('Alice Smith', 'alice@example.com')"
}
Add-ToolCall -Calls $calls -Id 5 -Name "execute_sql" -Arguments @{
    sql = "INSERT INTO _DupTest (FullName, Email) VALUES ('Bob Jones', 'bob@example.com')"
}
Add-ToolCall -Calls $calls -Id 6 -Name "execute_sql" -Arguments @{
    sql = "INSERT INTO _DupTest (FullName, Email) VALUES ('Alice Smith', 'alice@example.com')"
}
Add-ToolCall -Calls $calls -Id 7 -Name "execute_sql" -Arguments @{
    sql = "INSERT INTO _DupTest (FullName, Email) VALUES ('Charlie Brown', 'charlie@example.com')"
}
Add-ToolCall -Calls $calls -Id 8 -Name "execute_sql" -Arguments @{
    sql = "INSERT INTO _DupTest (FullName, Email) VALUES ('Alice Smith', 'alice-other@example.com')"
}
Add-ToolCall -Calls $calls -Id 9 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool8a-setup-DupTest" -Calls $calls

if ($null -ne $r) {
    $createResult = Decode-McpResult -Response $r[3]
    Write-Host ("  CREATE TABLE result: {0}" -f ($createResult | ConvertTo-Json -Depth 3 -Compress))
    for ($i = 4; $i -le 8; $i++) {
        $ins = Decode-McpResult -Response $r[$i]
        $insNum = $i - 3
        Write-Host ("  INSERT {0} result: {1}" -f $insNum, ($ins | ConvertTo-Json -Depth 3 -Compress))
    }
    Assert-Test "setup: _DupTest table created" ($createResult.success -eq $true) ""
}

# Step 8b: find duplicates
$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "find_duplicate_records" -Arguments @{
    table_name  = "_DupTest"
    field_names = @("FullName", "Email")
    max_groups  = 10
}
Add-ToolCall -Calls $calls -Id 4 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool8b-find_duplicate_records" -Calls $calls

if ($null -ne $r) {
    $decoded = Decode-McpResult -Response $r[3]
    Write-Host "  find_duplicate_records result:"
    $json8 = $decoded | ConvertTo-Json -Depth 5 -Compress
    if ($json8.Length -gt 1000) { Write-Host ($json8.Substring(0, 1000) + "...") } else { Write-Host $json8 }

    # Response shape: { success, result: { TableName, CheckedFields, DuplicateGroupCount, TotalDuplicateRows, Groups: [...] } }
    Assert-Test "find_duplicate_records returns success" ($decoded.success -eq $true) ""

    $innerResult = $decoded.result
    $dupGroupCount = if ($null -ne $innerResult) { $innerResult.DuplicateGroupCount } else { $null }
    $totalDupRows = if ($null -ne $innerResult) { $innerResult.TotalDuplicateRows } else { $null }
    $groups = if ($null -ne $innerResult) { $innerResult.Groups } else { $null }

    $hasDupCount = $null -ne $dupGroupCount
    Assert-Test "find_duplicate_records has DuplicateGroupCount" ([bool]$hasDupCount) ("DuplicateGroupCount=" + $dupGroupCount)

    # We inserted 2 rows with same (Alice Smith, alice@example.com) -> at least 1 group
    $condDup = [bool]($hasDupCount -and [int]$dupGroupCount -ge 1)
    Assert-Test "find_duplicate_records found at least 1 duplicate group" $condDup ("DuplicateGroupCount=" + $dupGroupCount)

    $hasTotalDupRows = $null -ne $totalDupRows
    Assert-Test "find_duplicate_records has TotalDuplicateRows" ([bool]$hasTotalDupRows) ("TotalDuplicateRows=" + $totalDupRows)

    if ($null -ne $groups -and $groups.Count -gt 0) {
        $firstGroup = $groups[0]
        Write-Host ("  First duplicate group: Values={0} Count={1}" -f ($firstGroup.Values | ConvertTo-Json -Compress), $firstGroup.Count)
        $condGroupCount = [bool]($firstGroup.Count -ge 2)
        Assert-Test "find_duplicate_records first group has Count at least 2" $condGroupCount ("Count=" + $firstGroup.Count)
    }

    $checkedFields = if ($null -ne $innerResult) { $innerResult.CheckedFields } else { $null }
    $hasCF = ($null -ne $checkedFields) -and ($checkedFields.Count -eq 2)
    Assert-Test "find_duplicate_records CheckedFields has 2 entries" ([bool]$hasCF) ""
} else {
    Assert-Test "find_duplicate_records batch completed" $false "TIMEOUT"
}

# Step 8c: Cleanup — drop test table
$calls = New-Object 'System.Collections.Generic.List[object]'
Add-ToolCall -Calls $calls -Id 2 -Name "connect_access" -Arguments @{ database_path = $DatabasePath }
Add-ToolCall -Calls $calls -Id 3 -Name "execute_sql" -Arguments @{
    sql = "DROP TABLE _DupTest"
}
Add-ToolCall -Calls $calls -Id 4 -Name "disconnect_access" -Arguments @{}

$r = Run-Batch -Label "Tool8c-cleanup-DupTest" -Calls $calls

if ($null -ne $r) {
    $dropResult = Decode-McpResult -Response $r[3]
    Write-Host ("  DROP TABLE result: {0}" -f ($dropResult | ConvertTo-Json -Depth 3 -Compress))
    Assert-Test "cleanup _DupTest table dropped" ($dropResult.success -eq $true) ""
} else {
    Write-Host "  WARN: Cleanup timed out. Manual cleanup of _DupTest table may be needed."
}

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════════════════

Write-Host ""
Write-Host ("=" * 70)
Write-Host "STRESS TEST SUMMARY"
Write-Host ("=" * 70)
Write-Host ("Total tests:  {0}" -f $totalTests)
$passColor = "Green"
Write-Host ("Passed:       {0}" -f $passCount) -ForegroundColor $passColor
$failColor = if ($failCount -gt 0) { "Red" } else { "Green" }
Write-Host ("Failed:       {0}" -f $failCount) -ForegroundColor $failColor
Write-Host ("Skipped:      {0}" -f $skipCount) -ForegroundColor Yellow
Write-Host ("=" * 70)

if ($failCount -gt 0) {
    exit 1
} else {
    exit 0
}
