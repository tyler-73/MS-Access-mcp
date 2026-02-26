# Microsoft Access MCP Server

A comprehensive Model Context Protocol (MCP) server that provides full Microsoft Access automation capabilities for Claude Desktop and other MCP clients.

## Overview

This MCP server enables AI assistants to interact with Microsoft Access databases through a complete set of tools that implement all seven required capabilities:

1. **Connection Management** - Establish and manage database connections
2. **Data Access Object Models** - Discover and manipulate tables, queries, and relationships
3. **COM Automation** - Launch Access and manage forms, reports, macros, and modules
4. **VBA Extensibility** - Read, write, and compile VBA code
5. **System Table Metadata Access** - Access hidden system tables and metadata
6. **Form & Control Discovery & Editing APIs** - Discover and modify form controls
7. **Persistence & Versioning** - Export/import database objects for version control

## Architecture

The solution consists of two main components:

1. **MS.Access.MCP.Interop** - A .NET 8.0 interop library that handles COM interactions with Microsoft Access
2. **MS.Access.MCP.Official** - An MCP server using the official ModelContextProtocol package that exposes Access functionality as MCP tools

## Prerequisites

- Windows operating system
- Microsoft Access installed (for COM interop)
- .NET 8.0 runtime
- Access database file (.accdb or .mdb)

## Installation

### Building from Source

1. Clone or download this repository
2. Build the solution:
   ```bash
   dotnet build
   ```
3. Publish the official server:
   ```bash
   dotnet publish MS.Access.MCP.Official/MS.Access.MCP.Official.csproj -c Release -r win-x64 --self-contained true -o ./mcp-server-official-x64
   ```

### Repeatable x64 Publish + Promotion

Use this from repo root for repeatable release promotion with backup safety:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1
```

What this script does:
1. Publishes to a timestamped staging folder (`mcp-server-official-x64-run-*`)
2. Stops running `MS.Access.MCP.Official` processes (default behavior) to unlock the target folder
3. Renames the active target to `mcp-server-official-x64-backup-*`
4. Promotes staging to `mcp-server-official-x64`
5. Runs an MCP `initialize` smoke test against the promoted exe

If promotion fails with `Access denied`, one or more server processes are still running under a context this shell cannot terminate. Rerun from an elevated PowerShell session after stopping `MS.Access.MCP.Official`.

Rollback behavior:
1. If promotion fails after a backup was created, the script attempts to restore the previous target from `mcp-server-official-x64-backup-*`.
2. If `-RunRegression` is enabled and regression fails after promotion, the script archives the promoted target as `mcp-server-official-x64-regression-failed-*` and restores the backup target.
3. If backup restore fails in step 2, the script attempts to restore the archived promoted target so an active target remains.

Backup retention options:

```powershell
# Keep all backups (default behavior)
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1 -BackupRetentionCount 0

# Keep only the newest 5 backup directories
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1 -BackupRetentionCount 5
```

Cleanup stale release folders:

```powershell
# Prune stale run/smoke folders, keep newest 5 in each set (default)
powershell -ExecutionPolicy Bypass -File .\scripts\prune-release-artifacts.ps1

# Preview cleanup without deleting anything
powershell -ExecutionPolicy Bypass -File .\scripts\prune-release-artifacts.ps1 -WhatIf

# Also prune backups explicitly, keeping newest 3 backups
powershell -ExecutionPolicy Bypass -File .\scripts\prune-release-artifacts.ps1 -IncludeBackups -BackupRetentionCount 3
```

If you need to skip automatic server process shutdown:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1 -StopServerProcesses $false
```

If you need a framework-dependent publish instead of self-contained output:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1 -SelfContained $false
```

Optional full regression invocation as part of release:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1 `
  -RunRegression `
  -RegressionDatabasePath "C:\path\to\database.accdb"
```

### Repair + Verify Hardening

Use this from repo root to clean stale state, enforce trusted location, probe candidate binaries with `connect_access`, and run full regression.

```powershell
# Full hardening run (includes regression and both config updates)
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateConfigs

# Dry-run / WhatIf
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateConfigs `
  -WhatIf

# Config update toggle: update only Codex config
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateCodexConfig

# Config update toggle: update only Claude config
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateClaudeConfig

# x86 fallback when x86 binary is missing (no extra flag required)
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb"
```

## Configuration

### Claude Desktop Configuration

Add the following to your Claude Desktop configuration file (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "access-mcp-server": {
      "command": "C:\\Users\\brickly\\Desktop\\MS-Access-MCP\\mcp-server-official\\MS.Access.MCP.Official.exe",
      "args": []
    }
  }
}
```

**Note**: Use the absolute path to your MCP server executable and double backslashes.

## Available Tools

The MCP server provides comprehensive tools across all seven capability areas:

### 1. Connection Management
- **connect_access**: Connect to an Access database (requires `database_path` parameter)
- **disconnect_access**: Disconnect from the current Access database
- **is_connected**: Check if connected to an Access database

### 2. Data Access Object Models
- **get_tables**: Get all tables from the connected database
- **get_queries**: Get all queries from the connected database
- **get_relationships**: Get all relationships from the connected database
- **create_query**: Create a saved Access query definition
- **update_query**: Update SQL for a saved query definition
- **delete_query**: Delete a saved query definition
- **create_relationship**: Create a relationship (`table_name`/`field_name` = referenced primary side, `foreign_table_name`/`foreign_field_name` = dependent side)
- **update_relationship**: Replace an existing relationship definition using the same parameter mapping
- **delete_relationship**: Delete a relationship by name
- **execute_sql**: Execute SQL directly (query or action)
- **execute_query_md**: Execute SQL and return markdown table output
- **describe_table**: Describe table schema, key columns, and defaults
- **create_table**: Create a new table in the database
- **delete_table**: Delete a table from the database
- **add_field**: Add a new field to an existing table
- **alter_field**: Alter an existing field definition on a table
- **rename_field**: Rename a field on an existing table
- **drop_field**: Drop a field from an existing table
- **rename_table**: Rename an existing table
- **get_indexes**: Get index metadata for a table
- **create_index**: Create an index on one or more columns
- **delete_index**: Delete an index from a table
- **list_linked_tables**: List linked tables in the current database
- **create_linked_table** (alias: `link_table`): Create a linked table in the current database from another Access database file
- **refresh_linked_table** (alias: `refresh_link`): Refresh an existing linked table definition
- **update_linked_table** (alias: `relink_table`): Repoint a linked table to a new source database/table
- **delete_linked_table** (alias: `unlink_table`): Remove a linked table from the current database
- **begin_transaction**: Begin a database transaction on the current connection
- **commit_transaction**: Commit the active transaction
- **rollback_transaction**: Roll back the active transaction
- **transaction_status**: Return current transaction status (active state, isolation level, and start time)

### 3. COM Automation
- **launch_access**: Launch Microsoft Access application
- **close_access**: Close Microsoft Access application
- **get_forms**: Get all forms in the database
- **get_reports**: Get all reports in the database
- **get_macros**: Get all macros in the database
- **get_modules**: Get all modules in the database
- **open_form**: Open a form in Access
- **close_form**: Close a form in Access
- **open_report**: Open a report in Access
- **close_report**: Close a report in Access
- **run_macro**: Run a macro by name
- **create_macro**: Create a macro from text representation
- **update_macro**: Update an existing macro from text representation

### 4. VBA Extensibility
- **get_vba_projects**: Get all VBA projects in the database
- **get_vba_code**: Get VBA code from a module
- **set_vba_code**: Set VBA code in a module
- **add_vba_procedure**: Add a VBA procedure to a module
- **compile_vba**: Compile VBA code

### 5. System Table Metadata Access
- **get_system_tables**: Get system tables from the database
- **get_object_metadata**: Get object metadata from system tables

### 6. Form & Control Discovery & Editing APIs
- **form_exists**: Check if a form exists
- **get_form_controls**: Get all controls in a form
- **get_control_properties**: Get properties of a control
- **set_control_property**: Set a property of a control
- **get_report_controls**: Get all controls in a report
- **get_report_control_properties**: Get properties of a report control
- **set_report_control_property**: Set a property of a report control

### 7. Persistence & Versioning
- **export_form_to_text**: Export a form to text representation
- **import_form_from_text**: Import a form from text representation
- **delete_form**: Delete a form from the database
- **export_report_to_text**: Export a report to text representation
- **import_report_from_text**: Import a report from text representation
- **delete_report**: Delete a report from the database
- **export_macro_to_text**: Export a macro to text representation
- **import_macro_from_text**: Import a macro from text representation
- **delete_macro**: Delete a macro from the database

`access_text` note:
- For `import_form_from_text` with `mode="access_text"`, pass `form_name` (object name to import).
- For `import_report_from_text` with `mode="access_text"`, pass `report_name` (object name to import).

## Usage Examples

### Basic Connection and Discovery
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "connect_access",
    "arguments": {
      "database_path": "C:\\path\\to\\database.accdb"
    }
  }
}
```

### Table Operations
```json
{
  "jsonrpc": "2.0",
  "id": 2,
  "method": "tools/call",
  "params": {
    "name": "get_tables",
    "arguments": {}
  }
}
```

### VBA Code Management
```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "method": "tools/call",
  "params": {
    "name": "set_vba_code",
    "arguments": {
      "project_name": "CurrentProject",
      "module_name": "Module1",
      "code": "Sub TestProcedure()\n    MsgBox \"Hello World\"\nEnd Sub"
    }
  }
}
```

### Form Export/Import
```json
{
  "jsonrpc": "2.0",
  "id": 4,
  "method": "tools/call",
  "params": {
    "name": "export_form_to_text",
    "arguments": {
      "form_name": "MyForm"
    }
  }
}
```

## CI

This repository includes two GitHub Actions workflows with different coverage goals:

- `windows-hosted-build-smoke.yml` runs on GitHub-hosted `windows-latest` for `push` and `pull_request`. It validates publish/build health and runs an MCP `initialize` smoke test that does not require Microsoft Access.
- `windows-self-hosted-access-regression.yml` runs on self-hosted Windows (`workflow_dispatch` plus weekly schedule) and executes `tests\full_toolset_regression.ps1`. It requires Microsoft Access on the runner and an `ACCESS_DATABASE_PATH` value provided by dispatch input (`access_database_path`) or secret (`ACCESS_DATABASE_PATH`).

## Testing

### Running the Full Regression Test

Use the committed PowerShell harness to validate every currently exposed MCP tool in one run:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_regression.ps1
```

Optional arguments:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_regression.ps1 `
  -ServerExe "C:\path\to\MS.Access.MCP.Official.exe" `
  -DatabasePath "C:\path\to\database.accdb"
```

The script verifies:
1. Connection lifecycle and status checks
2. Table creation/query/description/deletion
3. Deterministic schema evolution coverage (`add_field`, `alter_field`, `rename_field`, `drop_field`, `rename_table`)
4. SQL execution and markdown query output
5. VBA set/add/get/compile flows
6. Form import/export/control discovery/edit/delete in JSON mode, plus `mode="access_text"` export/import round-trip persistence checks
7. Report import/export/delete in JSON mode, plus `mode="access_text"` export/import round-trip persistence checks
8. Macro create/export/run/update/delete plus `import_macro_from_text` round-trip verification
9. Metadata discovery and Access COM automation calls
10. Linked-table tranche-1 coverage using a local copied `.accdb` source (no external database server dependency)
11. Transaction tranche-1 coverage validating rollback/commit visibility through deterministic SQL checks

When linked-table and transaction tranche-1 tools are not exposed by `tools/list`, the harness records `SKIP` lines for those sections and preserves the existing pass criterion.

Pass criterion: `TOTAL_FAIL=0` and process exit code `0`.

### Manual Protocol Probe

To test the server manually:

```bash
echo {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}} | dotnet run --project MS.Access.MCP.Official/MS.Access.MCP.Official.csproj
```

This should return a proper MCP initialize response with server information.

## Development

The server is built using the official `ModelContextProtocol` package (version 0.3.0-preview.3) and implements a JSON-RPC message handler that:

1. Reads JSON-RPC messages from stdin
2. Processes MCP protocol methods (initialize, tools/list, tools/call)
3. Executes Access database operations through the Interop library
4. Returns results as JSON-RPC responses on stdout

### Key Features

- **Full COM Interop**: Direct access to Microsoft Access COM objects
- **Comprehensive Coverage**: All seven required capabilities implemented
- **Error Handling**: Robust error handling and resource cleanup
- **Type Safety**: Strongly typed C# interfaces for all operations
- **Extensibility**: Easy to add new tools and capabilities

### Data Models

The Interop library includes comprehensive data models for:
- Tables, fields, and relationships
- Forms, reports, macros, and modules
- VBA projects and code
- System tables and metadata
- Form controls and properties
- Export/import data structures

## Troubleshooting

### Common Issues

1. **Access not installed**: Ensure Microsoft Access is installed on the system
2. **Database not found**: Verify the database path is correct and accessible
3. **Permission errors**: Ensure the application has permission to access the database file
4. **COM errors**: May indicate Access is not properly installed or registered
5. **Stale lock state (`.laccdb`)**: Close leftover `MSACCESS` processes and remove the sibling `.laccdb` file before rerunning tests

### MCP Preflight Diagnostics (Error Responses)

`connect_access`, `get_tables`, `get_queries`, `get_relationships`, `execute_sql`, `execute_query_md`, and `describe_table` now include a `preflight` object when `success=false`.

Preflight fields:
- `process_bitness`: Current MCP process bitness (`x86` or `x64`)
- `ace_oledb_provider_registered`: Whether `Microsoft.ACE.OLEDB.12.0` is registered for this process bitness
- `ace_oledb_issue_detected`: True when ACE provider availability/bitness mismatch is likely blocking operations
- `trust_center_active_content_indicator`: True when the failure message suggests Access Trust Center active-content blocking
- `remediation_hints`: Suggested next steps

Use these indicators for fast remediation:
1. If `ace_oledb_issue_detected=true`, install the Access Database Engine with matching bitness or run a matching MCP build.
2. If `trust_center_active_content_indicator=true`, add the database folder to Access Trusted Locations and unblock downloaded database files.

PowerShell cleanup snippet:

```powershell
Get-Process MSACCESS,MS.Access.MCP.Official -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Remove-Item C:\path\to\database.laccdb -ErrorAction SilentlyContinue
```

### Debug Mode

To run with detailed logging, modify the Program.cs to include console output for debugging:

```csharp
Console.WriteLine($"Processing method: {methodName}");
```

## License

This project is provided as-is for educational and development purposes.

## Contributing

Contributions are welcome! Please ensure all tests pass before submitting changes. 
