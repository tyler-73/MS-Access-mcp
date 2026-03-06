# Microsoft Access MCP Server

A comprehensive Model Context Protocol (MCP) server providing **350 tools** that cover 100% of the practical Microsoft Access COM/DAO automation surface. Works with Claude Desktop, Claude Code, VS Code (GitHub Copilot), and any MCP-compatible client.

## Overview

This MCP server enables AI assistants to interact with Microsoft Access databases through a complete toolset spanning:

- **Connection & database lifecycle** - Connect, create, backup, compact/repair, close
- **Tables, fields & schema** - Full DDL: create/alter/rename/drop tables and fields, indexes, calculated fields, multi-value fields
- **Queries & SQL** - Create/update/delete queries, execute SQL, passthrough queries, action queries, parameterized queries
- **Relationships** - Create, update, delete relationships with referential integrity
- **Forms & controls** - Create forms, open/close, get/set control properties, runtime state, conditional formatting
- **Reports & grouping** - Create reports, grouping/sorting, output/print, control properties
- **VBA & modules** - Full VBA IDE automation: get/set code, compile, run procedures, manage references, module analysis
- **Macros** - Create/update/run/delete macros, data macros, AutoExec
- **Linked tables** - Create/refresh/relink/delete linked tables, ODBC links, refresh all
- **Transactions** - Begin/commit/rollback with status tracking
- **DAO recordsets** - Open/navigate/CRUD recordset operations with bookmarks and filtering
- **Database properties** - Application options, startup properties, custom properties, DAO documents
- **TempVars** - Set/get/remove/clear TempVars
- **Attachments** - Add/remove/save attachment fields
- **Import/export** - XML exchange, navigation pane XML, import/export specs, schema snapshots, VBA export
- **Navigation groups** - Create/delete groups, add/remove objects
- **Security & encryption** - Database passwords, encryption
- **Printing** - PrintOut, OutputTo, printer management
- **DoCmd operations** - Comprehensive DoCmd surface (beep, echo, hourglass, send object, follow hyperlink, etc.)
- **pyodbc compatibility** - Drop-in replacement for 7 core pyodbc-access data tools

## Architecture

The solution consists of two components:

1. **MS.Access.MCP.Interop** - A .NET 8.0 interop library providing COM/DAO late-binding access to Microsoft Access. Handles exclusive mode, dialog dismissal, process lifecycle, and OleDb connectivity.
2. **MS.Access.MCP.Official** - A self-contained MCP server implementing the JSON-RPC stdio transport. Supports MCP protocol versions `2024-11-05` and `2025-11-25` with full capability negotiation (tools, resources, prompts, logging, completions).

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
6. Writes `mcp-server-official-x64\release-validation.json` after smoke success (includes `git_commit`, `regression_run`, and `regression_passed`)

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

Use framework-dependent output only for targeted diagnostics; normal MCP client/runtime config should keep using the validated `mcp-server-official-x64` promoted binary.

Optional full regression invocation as part of release:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-promote-x64.ps1 `
  -RunRegression `
  -RegressionDatabasePath "C:\path\to\database.accdb"
```

Use `-RunRegression` when you intend to consume the promoted binary with strict repair selection (`-RequireRegressionBackedManifest`).

### Repair + Verify Hardening

Use this from repo root to clean stale state, enforce trusted location, validate candidate binaries with `initialize` smoke + `connect_access`, and run full regression. By default, candidate manifests must include `git_commit` matching repo `HEAD`.

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

# Optional: override Claude config path (default resolves %APPDATA%\Claude\claude_desktop_config.json,
# then falls back to %USERPROFILE%\.claude.json if present)
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -UpdateClaudeConfig `
  -ClaudeConfigPath "C:\path\to\claude_desktop_config.json"

# Default: require validated x64 promoted binary and manifest git_commit matching HEAD
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb"

# Optional: bypass git_commit vs HEAD enforcement (diagnostics only)
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -AllowManifestHeadMismatch

# Strict mode: require regression-backed validation manifest for candidate selection
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -RequireRegressionBackedManifest

# Strict mode override: allow non-regression manifest while keeping strict flag visible
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -RequireRegressionBackedManifest `
  -AllowNonRegressionManifest

# Optional: allow unvalidated binaries (diagnostics only)
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -AllowUnvalidatedBinary

# Optional: include x86 candidate
powershell -ExecutionPolicy Bypass -File .\scripts\repair-and-verify-access-mcp.ps1 `
  -DatabasePath "C:\path\to\database.accdb" `
  -AllowX86Fallback
```

## Configuration

### Claude Desktop Configuration

Add the following to your Claude Desktop configuration file (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "access-mcp": {
      "command": "C:\\path\\to\\mcp-server-official-x64\\MS.Access.MCP.Official.exe",
      "args": []
    }
  }
}
```

### Claude Code Configuration

```bash
claude mcp add access-mcp -- "C:\path\to\mcp-server-official-x64\MS.Access.MCP.Official.exe"
```

### VS Code Configuration (GitHub Copilot)

Add to `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "access-mcp": {
      "type": "stdio",
      "command": "C:\\path\\to\\mcp-server-official-x64\\MS.Access.MCP.Official.exe",
      "args": []
    }
  }
}
```

**Notes**:
- Use the absolute path to your MCP server executable with double backslashes in JSON.
- The server auto-discovers databases: pass `database_path` to `connect_access`, set the `ACCESS_DATABASE_PATH` env var, or place a `.accdb` file in your Documents folder.
- `repair-and-verify-access-mcp.ps1 -UpdateClaudeConfig` updates either `mcpServers.access-mcp-server.command` or `mcpServers.access-mcp.command` (whichever exists) and fails fast if neither key is present.

## Available Tools (350)

The server exposes 350 tools covering the full practical Access COM/DAO surface. Below is a summary by category — run `tools/list` for the complete schema with argument definitions.

### Connection & Database Lifecycle
`connect_access`, `disconnect_access`, `is_connected`, `close_access`, `close_database`, `launch_access`, `create_database`, `backup_database`, `compact_repair_database`

### Tables & Schema
`get_tables`, `create_table`, `delete_table`, `describe_table`, `rename_table`, `get_system_tables`, `get_table_properties`, `get_table_custom_property`, `set_table_custom_property`, `get_table_description`, `set_table_description`, `get_table_validation`

### Fields
`add_field` (with format/decimal_places/description/default_value), `alter_field`, `rename_field`, `drop_field`, `get_field_properties`, `get_field_attributes`, `set_field_required`, `set_field_format`, `set_field_description`, `set_field_validation`, `set_field_input_mask`, `set_field_caption`, `set_field_decimal_places`, `set_field_default`, `set_field_allow_zero_length`, `set_field_append_only`, `add_calculated_field`, `detect_multi_value_fields`, `get_multi_value_field_values`, `set_multi_value_field_values`

### Queries & SQL
`create_query`, `update_query`, `delete_query`, `get_queries`, `execute_sql`, `execute_sql_timed`, `execute_query_md`, `execute_action_query`, `set_parameter`, `get_query_parameters`, `get_query_properties`, `set_query_properties`, `set_query_advanced_properties`, `create_passthrough_query`

### Relationships & Indexes
`create_relationship`, `update_relationship`, `delete_relationship`, `get_relationships`, `create_index` (with ignore_nulls), `delete_index`, `get_indexes`

### Linked Tables
`create_linked_table` / `link_table`, `refresh_linked_table` / `refresh_link`, `update_linked_table` / `relink_table`, `delete_linked_table` / `unlink_table`, `list_linked_tables`, `refresh_all_linked_tables`, `create_odbc_linked_table`

### Transactions
`begin_transaction` / `start_transaction`, `commit_transaction`, `rollback_transaction`, `transaction_status`

### Forms
`create_form`, `delete_form`, `form_exists`, `get_forms`, `open_form` (full params: view, filter, where, data_mode, window_mode), `close_form`, `get_form_controls`, `get_control_properties`, `set_control_property`, `get_form_runtime_state`, `get_active_form`

### Reports
`create_report`, `delete_report`, `get_reports`, `open_report` (full params), `close_report`, `get_report_controls`, `get_report_control_properties`, `set_report_control_property`, `get_report_grouping`, `set_report_grouping`, `delete_report_grouping`, `get_report_sorting`, `set_report_sorting`, `output_to`, `print_out`, `get_active_report`

### DAO Recordsets
`open_recordset`, `close_recordset`, `recordset_get_rows`, `recordset_get_record`, `recordset_add_record`, `recordset_delete_record`, `recordset_edit_record`, `recordset_count`, `recordset_find`, `recordset_move`, `recordset_bookmark`, `recordset_filter_sort`

### VBA & Modules
`get_vba_code`, `set_vba_code`, `compile_vba`, `is_vba_compiled`, `add_vba_procedure`, `run_vba_procedure`, `execute_vba`, `get_vba_projects`, `get_vba_references`, `add_vba_reference`, `remove_vba_reference`, `create_module`, `delete_module`, `rename_module`, `get_modules`, `get_module_info`, `get_module_declarations`, `list_procedures`, `list_all_procedures`, `get_procedure_code`, `find_text_in_module`, `insert_lines`, `delete_lines`, `replace_line`, `get_compilation_errors`, `export_all_vba`, `get_vba_project_properties`, `set_vba_project_properties`

### Macros
`create_macro`, `update_macro`, `delete_macro`, `get_macros`, `run_macro` (with repeat count), `export_macro_to_text`, `import_macro_from_text`

### Persistence & Versioning
`export_form_to_text`, `import_form_from_text`, `export_report_to_text`, `import_report_from_text` (both support `mode="access_text"`)

### Database Properties & Metadata
`get_database_properties`, `get_database_property`, `set_database_property`, `get_database_summary_properties`, `set_database_summary_properties`, `get_application_info`, `get_application_option`, `set_application_option`, `get_current_project_data`, `get_database_engine_info`, `get_database_statistics`, `get_object_metadata`, `get_object_dates`, `get_current_user`, `get_autoexec_info`, `get_containers`, `get_container_documents`, `get_document_properties`, `set_document_property`, `get_open_objects`, `is_object_loaded`, `get_current_object`

### Startup Properties
`get_startup_properties`, `set_startup_properties`, `get_startup_properties_extended`, `set_startup_properties_extended`

### TempVars
`set_temp_var`, `get_temp_vars`, `remove_temp_var`, `clear_temp_vars`

### Attachments
`add_attachment_file`, `remove_attachment_file`, `get_attachment_files`, `get_attachment_metadata`, `save_attachment_to_disk`

### Import/Export & XML
`export_xml`, `import_xml`, `export_navigation_pane_xml`, `export_schema_snapshot`, `create_import_export_spec`, `delete_import_export_spec`, `get_import_export_spec`, `list_import_export_specs`

### Data Macros
`run_data_macro`, `delete_data_macro`, `export_data_macro_axl`, `import_data_macro_axl`, `get_table_data_macros`

### Conditional Formatting
`add_conditional_formatting`, `update_conditional_formatting`, `delete_conditional_formatting`, `clear_conditional_formatting`, `get_conditional_formatting`, `list_all_conditional_formats`

### Navigation Groups
`create_navigation_group`, `delete_navigation_group`, `get_navigation_groups`, `add_navigation_group_object`, `remove_navigation_group_object`, `get_navigation_group_objects`

### DoCmd Operations
`run_sql`, `run_command`, `beep`, `echo`, `hourglass`, `set_warnings`, `refresh_database_window`, `requery`, `show_all_records`, `apply_filter`, `open_table`, `open_query`, `save_object`, `close_object`, `rename_object`, `copy_object`, `delete_object`, `select_object`, `goto_record`, `goto_page`, `goto_control`, `find_record`, `find_next`, `maximize_window`, `minimize_window`, `restore_window`, `move_size`, `send_object`, `output_to`, `print_out`, `browse_to`, `navigate_to`, `follow_hyperlink`, `run_autoexec`, `sys_cmd`, `search_for_record`

### Security & Encryption
`set_database_password`, `remove_database_password`, `get_database_security`, `encrypt_database`

### Printing
`list_printers`, `get_printer_info`, `set_default_printer`

### Subdatasheet Properties
`get_subdatasheet_properties`, `set_subdatasheet_properties`, `reset_subdatasheet_properties`

### Miscellaneous
`access_error`, `build_criteria`, `set_hidden_attribute`, `get_hidden_attribute`, `get_access_hwnd`, `set_access_visible`, `reset_autonumber`, `get_object_events`, `set_object_event`, `find_duplicate_records`, `check_referential_integrity`, `list_odbc_data_sources`, `domain_aggregate`

### pyodbc Compatibility Layer (7 tools)

Drop-in replacement for the core data tools from `pyodbc-access` (OpenLink `mcp-pyodbc-server`):

`podbc_get_schemas`, `podbc_get_tables`, `podbc_filter_table_names`, `podbc_describe_table`, `podbc_query_database`, `podbc_execute_query`, `podbc_execute_query_md`

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
- `windows-self-hosted-access-regression.yml` runs on self-hosted Windows (`workflow_dispatch` plus weekly schedule) and executes `tests\full_toolset_regression.ps1`, `tests\full_toolset_negative_paths.ps1`, and `tests\podbc_compat_regression.ps1`. It requires Microsoft Access on the runner and an `ACCESS_DATABASE_PATH` value provided by dispatch input (`access_database_path`) or secret (`ACCESS_DATABASE_PATH`), and asserts that database-lifecycle, secure-connect, and podbc compatibility coverage markers are present in logs.

Bootstrap GitHub auth/secret/workflow setup from terminal:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\bootstrap-github-actions.ps1 `
  -SetDatabaseSecret `
  -DatabasePath "C:\path\to\database.accdb" `
  -TriggerRegressionWorkflow
```

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

UI coverage mode:

```powershell
# Default behavior is headless (prevents Access UI pop-ups during automation)
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_regression.ps1

# Opt-in only when you intentionally want UI-opening tool coverage
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_regression.ps1 -IncludeUiCoverage
```

The script verifies:
1. Connection lifecycle and status checks
2. Table creation/query/description/deletion
3. Database lifecycle coverage for `create_database`, `backup_database`, and `compact_repair_database` using ephemeral local `.accdb` files
4. Deterministic schema evolution coverage (`add_field`, `alter_field`, `rename_field`, `drop_field`, `rename_table`)
5. SQL execution and markdown query output
6. VBA set/add/get/compile flows
7. Form import/export/control discovery/edit/delete in JSON mode, plus `mode="access_text"` export/import round-trip persistence checks
8. Report import/export/delete in JSON mode, plus `mode="access_text"` export/import round-trip persistence checks
9. Macro create/export/run/update/delete plus `import_macro_from_text` round-trip verification
10. Metadata discovery and Access COM automation calls
11. Linked-table tranche-1 coverage using a local copied `.accdb` source (no external database server dependency), including explicit alias-path calls beyond candidate resolution (`link_table`, `refresh_link`, `relink_table`, `unlink_table` when exposed)
12. Transaction tranche-1 coverage validating rollback/commit visibility through deterministic SQL checks, including explicit alias-path begin coverage (`start_transaction` when exposed)

By default, linked-table, transaction, and database lifecycle coverage are required: if those tool families are missing from `tools/list`, the harness records `FAIL` and increments `TOTAL_FAIL`.

If you are intentionally validating a reduced server surface, use:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_regression.ps1 -AllowCoverageSkips
```

Pass criterion: `TOTAL_FAIL=0` and process exit code `0`.

### Running the Negative-Path Regression Test

Use this harness to validate failure contracts (expected `success=false` + preflight coverage on disconnected operations), including invalid-path checks for `create_database`/`backup_database`/`compact_repair_database` and secure-argument validation coverage for `connect_access`:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_negative_paths.ps1
```

Optional arguments:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\full_toolset_negative_paths.ps1 `
  -ServerExe "C:\path\to\MS.Access.MCP.Official.exe" `
  -DatabasePath "C:\path\to\database.accdb"
```

Pass criterion: output contains `NEGATIVE_PATHS_PASS=1` and process exit code `0`.

### Manual Protocol Probe

To test the server manually:

```bash
echo {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-11-25","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}} | .\mcp-server-official-x64\MS.Access.MCP.Official.exe
```

This should return a proper MCP initialize response with server info and capabilities.

## Development

The server implements the MCP JSON-RPC stdio transport directly (no external MCP SDK dependency). It:

1. Reads JSON-RPC messages line-by-line from stdin
2. Negotiates protocol version (supports `2024-11-05` and `2025-11-25`)
3. Dispatches to handlers for `initialize`, `ping`, `tools/list`, `tools/call`, `resources/*`, `prompts/*`, `completion/complete`, `logging/setLevel`
4. Executes Access operations via the Interop library (COM late-binding + OleDb)
5. Returns JSON-RPC responses on stdout

### Key Features

- **Full COM/DAO Interop**: Late-binding access to the complete Access object model
- **350 Tools**: Covers connection, DDL, DML, forms, reports, VBA, macros, recordsets, properties, attachments, navigation groups, conditional formatting, data macros, security, printing, and more
- **Exclusive Mode**: Automatic exclusive DB access for DDL operations with process lifecycle management
- **Dialog Dismisser**: Background thread auto-dismisses modal Access/VBA dialogs during batch operations
- **MCP Resources**: 10 read-only resources (connection status, tables, queries, relationships, etc.)
- **MCP Prompts**: 6 prompt templates for common database operations
- **Preflight Diagnostics**: Error responses include bitness, ACE provider, and Trust Center indicators
- **pyodbc Compatibility**: Drop-in replacement for 7 core pyodbc-access tools

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
