param(
    [string]$ServerExe = "$PSScriptRoot\..\mcp-server-official-x64\MS.Access.MCP.Official.exe",
    [string]$DatabasePath = $(if ($env:ACCESS_DATABASE_PATH) { $env:ACCESS_DATABASE_PATH } else { "$env:USERPROFILE\Documents\MyDatabase.accdb" }),
    [switch]$NoCleanup
)

$ErrorActionPreference = "Stop"

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

function Stop-StaleProcesses {
    Get-Process MSACCESS, MS.Access.MCP.Official -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
}

function Remove-LockFile {
    param([string]$DbPath)

    $dbDir = Split-Path -Path $DbPath -Parent
    $dbName = [System.IO.Path]::GetFileNameWithoutExtension($DbPath)
    $lockFile = Join-Path $dbDir ($dbName + ".laccdb")
    Remove-Item -Path $lockFile -ErrorAction SilentlyContinue
}

function Cleanup-AccessArtifacts {
    param([string]$DbPath)

    Stop-StaleProcesses
    Remove-LockFile -DbPath $DbPath
}

if (-not (Test-Path -LiteralPath $ServerExe)) {
    throw "Server executable not found: $ServerExe"
}

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database file not found: $DatabasePath"
}

if (-not $NoCleanup) {
    Write-Host "Pre-run cleanup: clearing stale Access/MCP processes and locks."
    Cleanup-AccessArtifacts -DbPath $DatabasePath
}
else {
    Write-Warning "Skipping pre-run cleanup per -NoCleanup; final cleanup will still execute."
}

$exitCode = 1
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
Add-ToolCall -Calls $calls -Id 4 -Name "launch_access" -Arguments @{}
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
Add-ToolCall -Calls $calls -Id 26 -Name "open_form" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 27 -Name "close_form" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 28 -Name "import_report_from_text" -Arguments @{ report_data = $reportData }
Add-ToolCall -Calls $calls -Id 55 -Name "open_report" -Arguments @{ report_name = $reportName }
Add-ToolCall -Calls $calls -Id 56 -Name "close_report" -Arguments @{ report_name = $reportName }
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
    4 = "launch_access"
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
    26 = "open_form"
    27 = "close_form"
    28 = "import_report_from_text"
    55 = "open_report"
    56 = "close_report"
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

Write-Host ("TOTAL_FAIL={0}" -f $failed)
if ($failed -eq 0) {
    $exitCode = 0
}
}
finally {
    Write-Host "Final cleanup: clearing stale Access/MCP processes and locks."
    Cleanup-AccessArtifacts -DbPath $DatabasePath
}

exit $exitCode
