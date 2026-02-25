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

if (-not (Test-Path -LiteralPath $ServerExe)) {
    throw "Server executable not found: $ServerExe"
}

if (-not (Test-Path -LiteralPath $DatabasePath)) {
    throw "Database file not found: $DatabasePath"
}

if (-not $NoCleanup) {
    Stop-StaleProcesses
    Remove-LockFile -DbPath $DatabasePath
}

$suffix = [Guid]::NewGuid().ToString("N").Substring(0, 8)
$tableName = "MCP_Table_$suffix"
$formName = "MCP_Form_$suffix"
$reportName = "MCP_Report_$suffix"
$moduleName = "MCP_Module_$suffix"
$queryName = "MCP_Query_$suffix"
$relationshipName = "MCP_Rel_$suffix"
$childTableName = "MCP_Child_$suffix"

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
Add-ToolCall -Calls $calls -Id 26 -Name "open_form" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 27 -Name "close_form" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 28 -Name "import_report_from_text" -Arguments @{ report_data = $reportData }
Add-ToolCall -Calls $calls -Id 29 -Name "export_report_to_text" -Arguments @{ report_name = $reportName }
Add-ToolCall -Calls $calls -Id 30 -Name "delete_report" -Arguments @{ report_name = $reportName }
Add-ToolCall -Calls $calls -Id 31 -Name "delete_form" -Arguments @{ form_name = $formName }
Add-ToolCall -Calls $calls -Id 32 -Name "get_forms" -Arguments @{}
Add-ToolCall -Calls $calls -Id 33 -Name "get_reports" -Arguments @{}
Add-ToolCall -Calls $calls -Id 34 -Name "get_macros" -Arguments @{}
Add-ToolCall -Calls $calls -Id 35 -Name "get_modules" -Arguments @{}
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
Add-ToolCall -Calls $calls -Id 36 -Name "delete_table" -Arguments @{ table_name = $tableName }
Add-ToolCall -Calls $calls -Id 37 -Name "disconnect_access" -Arguments @{}
Add-ToolCall -Calls $calls -Id 38 -Name "is_connected" -Arguments @{}
Add-ToolCall -Calls $calls -Id 39 -Name "close_access" -Arguments @{}

$jsonLines = New-Object 'System.Collections.Generic.List[string]'
$jsonLines.Add((@{
    jsonrpc = "2.0"
    id = 1
    method = "initialize"
    params = @{
        protocolVersion = "2024-11-05"
        capabilities = @{}
        clientInfo = @{
            name = "full-regression"
            version = "1.0"
        }
    }
} | ConvertTo-Json -Depth 40 -Compress))

foreach ($call in $calls) {
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

$rawLines = @((($jsonLines -join "`n") | & $ServerExe))
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

$idLabels = @{
    2 = "connect_access"
    3 = "is_connected_initial"
    4 = "launch_access"
    5 = "get_tables"
    6 = "get_queries"
    7 = "get_relationships"
    8 = "create_table"
    9 = "describe_table"
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
    26 = "open_form"
    27 = "close_form"
    28 = "import_report_from_text"
    29 = "export_report_to_text"
    30 = "delete_report"
    31 = "delete_form"
    32 = "get_forms"
    33 = "get_reports"
    34 = "get_macros"
    35 = "get_modules"
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
    36 = "delete_table"
    37 = "disconnect_access"
    38 = "is_connected_after_disconnect"
    39 = "close_access"
}

$failed = 0
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
        "export_report_to_text" {
            if ([string]::IsNullOrWhiteSpace([string]$decoded.report_data)) {
                $failed++
                Write-Host ('{0}: FAIL empty report export payload' -f $label)
                continue
            }
        }
    }

    Write-Host ('{0}: OK' -f $label)
}

Write-Host ("TOTAL_FAIL={0}" -f $failed)
if ($failed -gt 0) {
    exit 1
}

exit 0
