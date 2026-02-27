# MS Access MCP Server — Agent Instructions

## Architecture

This is a .NET 8 (x64-only, Windows) MCP server using custom JSON-RPC over stdio.
- `MS.Access.MCP.Official/Program.cs` — JSON-RPC dispatch, handler methods, MCP protocol
- `MS.Access.MCP.Interop/AccessInteropService.cs` — All COM/DAO/OLEDB interop with Access
- No SDK builder pattern — dispatch is a `method switch` in the main loop (line ~60)
- Handler pattern: `static object HandleXxx(AccessInteropService svc, JsonElement args)`
- Always try-catch, return `{ success = true, ... }` or `BuildOperationErrorResponse(...)`
- Tools registered in `HandleToolsList()` as anonymous objects with `name`, `description`, `inputSchema`

## Build & Test

- Build: `dotnet build MS.Access.MCP.Official/MS.Access.MCP.Official.csproj -c Release`
- Publish: `dotnet publish MS.Access.MCP.Official/MS.Access.MCP.Official.csproj -c Release -r win-x64 --self-contained true -o mcp-server-official-x64`
- Smoke test: `powershell -ExecutionPolicy Bypass -File tests/ci_initialize_smoke.ps1`
- Full test (requires DB): `powershell -ExecutionPolicy Bypass -File tests/full_toolset_regression.ps1`
- Build after EVERY batch of changes. Fix errors before continuing.

## Conventions

- All new interop methods go in AccessInteropService.cs as public methods
- All new tools go in Program.cs: add to HandleToolsList (tool definition) AND HandleToolsCall (dispatch) AND a new HandleXxx method
- Use existing helpers: TryGetRequiredString, TryGetOptionalString, GetOptionalBool, BuildOperationErrorResponse
- Return types: create new classes at the bottom of AccessInteropService.cs (namespace MS.Access.MCP.Interop)
- Log notable operations via SendLogNotification(level, logger, data)
- Do NOT modify existing tool handlers — only add new ones
- Do NOT change the JsonRpcResponse/JsonRpcRequest/JsonRpcErrorResponse classes
- Keep the same coding style — no extra usings, no external packages, no async interop

## Current Coverage (85 tools, ~35-40% of Access surface)

Covered: connection, table CRUD, indexes, queries, SQL exec, relationships, linked tables,
transactions, forms, reports, macros, VBA, metadata, resources, prompts, completion, logging.

## What To Implement (Priority Order)

### Priority 1: Import/Export (DoCmd transfers)
- TransferSpreadsheet (Excel import/export)
- TransferText (CSV/delimited import/export)
- OutputTo (export to PDF, XLS, RTF, TXT, HTML)
- These use Access COM: accessApp.DoCmd.TransferSpreadsheet(...) etc.

### Priority 2: Remaining DoCmd Operations
- SetWarnings, Echo, Hourglass (session control)
- GoToRecord, FindRecord (navigation)
- ApplyFilter, ShowAllRecords (filtering)
- Maximize, Minimize, Restore (window control)
- PrintOut (printing)
- OpenQuery (open query in datasheet)
- RunSQL (direct SQL execution via DoCmd)

### Priority 3: Database & Object Properties
- Database summary properties (Title, Author, Subject, Keywords, Comments)
- Custom database properties (read/write via DAO Properties collection)
- Table description, validation rules, validation text
- Field validation rules, input masks, default values
- Query description, parameters

### Priority 4: Field Property Enhancements
- Set field ValidationRule, ValidationText
- Set field DefaultValue
- Set field InputMask
- Set field Caption
- Lookup field properties (RowSource, BoundColumn, ColumnCount, etc.)

### Priority 5: VBA & Application Enhancements
- Enumerate/manage VBA project references
- Application startup properties (StartupForm, AppTitle, AppIcon)
- Ribbon XML read/write
- CurrentProject/CurrentData properties

### Priority 6: Advanced Features
- Data macros (table-level triggers via SaveAsAXL/LoadFromAXL)
- Attachment field read/write
- Navigation pane groups
- Conditional formatting rules
- Database encryption/password management

## Research Phase

Before implementing, search GitHub for these repos and examine their implementations:
- github.com/brickly26/MS-Access-mcp
- github.com/sub-arjun/OMNI-MS-Access-MCP
- github.com/ayamnash/MCP_server_ms_access_control
- github.com/scanzy/mcp-server-access-mdb

Extract any useful patterns for DoCmd, import/export, properties, or features we lack.
These are Python/Node servers — adapt the logic to our C# COM interop pattern.

## Workflow

1. Research: fetch competitor repos, identify useful implementations
2. Plan: list the exact methods and tools to add for each priority batch
3. Implement interop methods in AccessInteropService.cs
4. Implement tool definitions + handlers in Program.cs
5. Build & fix any compilation errors
6. Run smoke test to verify nothing broke
7. Git commit with message: "feat: Priority N — [description of what was added]"
8. Move to next priority batch
9. Repeat until all 6 priorities are done
