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

## Current Coverage (310 tools, 100% of practical Access COM/DAO surface)

All COM-accessible, non-deprecated features that a developer working with Access through
an MCP server would reasonably use. Excluded: deprecated replication, user-level security
(.mdw), SharePoint sync, DoCmd.Quit (kills COM server), Data Access Pages (removed in 2013),
ADP projects, report drawing (Circle/Line/PSet), FileDialog, CommandBars, COMAddIns, AutoCorrect.

Covered areas: connection, table CRUD, indexes, queries, SQL exec, relationships, linked tables,
transactions, forms, reports, macros, VBA, metadata, resources, prompts, completion, logging,
import/export (TransferSpreadsheet, TransferText, OutputTo), comprehensive DoCmd operations,
database/object/field properties, VBA references, startup properties, ribbon XML, data macros,
attachments, navigation groups, conditional formatting, database encryption/password,
form/report design-time APIs (sections, controls, properties, tab order, page setup),
module analysis (info, procedures, declarations, insert/delete/replace lines, find),
import/export specs, app options, DAO containers/documents, AutoExec, query parameters,
report grouping/sorting, page setup, printer info, DAO document properties, table validation,
field attributes, multi-value fields, table/field descriptions, advanced VBA execution
(Eval, Run, module CRUD, compilation errors, project properties), TempVars, open objects,
form runtime (record count, current record, filter), RefreshDatabaseWindow, SysCmd,
DoCmd remaining (FindNext, SearchForRecord, SetFilter, SetOrderBy, SetParameter, SetProperty,
RefreshRecord, CloseDatabase), domain aggregates (DLookup/DCount/DSum/DAvg/DMin/DMax/DFirst/DLast),
AccessError, BuildCriteria, Screen object (ActiveForm/ActiveReport/ActiveControl/ActiveDatasheet),
object visibility (SetHiddenAttribute/GetHiddenAttribute), CurrentObjectName/Type, CurrentUser,
Application.Visible, hWndAccessApp, form runtime methods (Recalc/Refresh/Requery/Undo/SetFocus,
Dirty/NewRecord/Bookmark/CurrentView/OpenArgs/Painting), DAO Recordset operations (open/close,
move/find, get record/rows, add/edit/delete, count/bookmark/filter+sort with handle tracking),
XML exchange (ExportXML/ImportXML/TransformXML, NavigationPane XML), printer management
(set default/form/report printer, list printers), database engine info, control methods
(SetFocus/Requery/Undo, ComboBox.Dropdown, ListBox AddItem/RemoveItem/GetItems),
AccessObject metadata (DateCreated/DateModified, IsLoaded), IsCompiled + broken references.
