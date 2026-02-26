using System.Text.Json;
using System.Text.Json.Serialization;
using System.Linq;
using System.Runtime.InteropServices;
using MS.Access.MCP.Interop;

class Program
{
    static readonly JsonElement EmptyJsonObject = JsonSerializer.Deserialize<JsonElement>("{}");

    static async Task Main(string[] args)
    {
        // JSON-RPC mode — no output until we receive a request
        
        var accessService = new AccessInteropService();
        
        try
        {
            string? line;
            while ((line = await Console.In.ReadLineAsync()) != null)
            {
                try
                {
                    var normalizedLine = line.TrimStart('\uFEFF', '\u00EF', '\u00BB', '\u00BF');
                    if (string.IsNullOrWhiteSpace(normalizedLine))
                        continue;

                    var trimmed = normalizedLine.TrimStart();
                    if (!(trimmed.StartsWith("{") || trimmed.StartsWith("[")))
                        continue;

                    var document = JsonDocument.Parse(trimmed);
                    var root = document.RootElement;
                    
                    if (!root.TryGetProperty("method", out var methodElement))
                        continue;
                        
                    var method = methodElement.GetString();
                    if (string.IsNullOrEmpty(method))
                        continue;
                        
                    // Skip notifications (no response needed)
                    if (method.StartsWith("notifications/"))
                        continue;

                    JsonElement? id = null;
                    if (root.TryGetProperty("id", out var idElement))
                        id = idElement.Clone();

                    var hasParams = root.TryGetProperty("params", out var paramsElement);
                    var safeParams = hasParams ? paramsElement : EmptyJsonObject;

                    object result = method switch
                    {
                        "initialize" => HandleInitialize(),
                        "tools/list" => HandleToolsList(),
                        "tools/call" => WrapCallToolResult(HandleToolsCall(accessService, safeParams)),
                        _ => new { error = $"Unknown method: {method}" }
                    };

                    var response = new JsonRpcResponse
                    {
                        Id = id,
                        Result = result
                    };

                    var jsonResponse = JsonSerializer.Serialize(response);
                    Console.WriteLine(jsonResponse);
                }
                catch (JsonException ex)
                {
                    // Log JSON parsing errors to stderr
                    Console.Error.WriteLine($"JSON parsing error: {ex.Message}");
                    continue;
                }
                catch (Exception ex)
                {
                    // Log other errors to stderr
                    Console.Error.WriteLine($"Error processing request: {ex.Message}");
                    continue;
                }
            }
        }
        catch (Exception ex)
        {
            // Log fatal errors to stderr
            Console.Error.WriteLine($"Fatal error: {ex.Message}");
            Environment.Exit(1);
        }
    }

    static object HandleInitialize()
    {
        return new
        {
            protocolVersion = "2024-11-05",
            capabilities = new { tools = new { } },
            serverInfo = new
            {
                name = "Access MCP Server",
                version = "1.0.1"
            }
        };
    }

    static object HandleToolsList()
    {
        return new
        {
            tools = new object[]
            {
                new { name = "connect_access", description = "Connect to an Access database. Uses database_path argument, ACCESS_DATABASE_PATH env var, or first database found in Documents.", inputSchema = new { type = "object", properties = new { database_path = new { type = "string" } } } },
                new { name = "disconnect_access", description = "Disconnect from the current Access database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "is_connected", description = "Check if connected to an Access database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_tables", description = "Get list of all tables in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_queries", description = "Get list of all queries in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_relationships", description = "Get list of all relationships in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "list_linked_tables", description = "List all linked tables in the current Access database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "link_table", description = "Create a linked table to an external Access database table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, source_database_path = new { type = "string" }, source_table_name = new { type = "string" }, connect_string = new { type = "string" }, overwrite = new { type = "boolean" } }, required = new string[] { "table_name", "source_database_path", "source_table_name" } } },
                new { name = "create_linked_table", description = "Alias for link_table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, source_database_path = new { type = "string" }, source_table_name = new { type = "string" }, connect_string = new { type = "string" }, overwrite = new { type = "boolean" } }, required = new string[] { "table_name", "source_database_path", "source_table_name" } } },
                new { name = "refresh_link", description = "Refresh a linked table connection", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "refresh_linked_table", description = "Alias for refresh_link", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "relink_table", description = "Change a linked table to point to a new source database/table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, source_database_path = new { type = "string" }, source_table_name = new { type = "string" }, connect_string = new { type = "string" } }, required = new string[] { "table_name", "source_database_path" } } },
                new { name = "update_linked_table", description = "Alias for relink_table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, source_database_path = new { type = "string" }, source_table_name = new { type = "string" }, connect_string = new { type = "string" } }, required = new string[] { "table_name", "source_database_path" } } },
                new { name = "unlink_table", description = "Remove a linked table from the database", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "delete_linked_table", description = "Alias for unlink_table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "create_query", description = "Create a saved Access query definition", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" }, sql = new { type = "string" } }, required = new string[] { "query_name", "sql" } } },
                new { name = "update_query", description = "Update SQL text for an existing saved query", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" }, sql = new { type = "string" } }, required = new string[] { "query_name", "sql" } } },
                new { name = "delete_query", description = "Delete a saved Access query definition", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" } }, required = new string[] { "query_name" } } },
                new { name = "create_relationship", description = "Create a table relationship", inputSchema = new { type = "object", properties = new { relationship_name = new { type = "string" }, table_name = new { type = "string" }, field_name = new { type = "string" }, foreign_table_name = new { type = "string" }, foreign_field_name = new { type = "string" }, enforce_integrity = new { type = "boolean" }, cascade_update = new { type = "boolean" }, cascade_delete = new { type = "boolean" } }, required = new string[] { "table_name", "field_name", "foreign_table_name", "foreign_field_name" } } },
                new { name = "update_relationship", description = "Replace an existing relationship definition", inputSchema = new { type = "object", properties = new { relationship_name = new { type = "string" }, table_name = new { type = "string" }, field_name = new { type = "string" }, foreign_table_name = new { type = "string" }, foreign_field_name = new { type = "string" }, enforce_integrity = new { type = "boolean" }, cascade_update = new { type = "boolean" }, cascade_delete = new { type = "boolean" } }, required = new string[] { "relationship_name", "table_name", "field_name", "foreign_table_name", "foreign_field_name" } } },
                new { name = "delete_relationship", description = "Delete an existing relationship by name", inputSchema = new { type = "object", properties = new { relationship_name = new { type = "string" } }, required = new string[] { "relationship_name" } } },
                new { name = "execute_sql", description = "Execute a SQL statement against the connected Access database. For SELECT queries, returns columns and rows. For action queries, returns rows_affected.", inputSchema = new { type = "object", properties = new { sql = new { type = "string" }, max_rows = new { type = "integer" } }, required = new string[] { "sql" } } },
                new { name = "execute_query_md", description = "Execute a SQL statement and return result as a markdown table (or action-query summary).", inputSchema = new { type = "object", properties = new { sql = new { type = "string" }, max_rows = new { type = "integer" } }, required = new string[] { "sql" } } },
                new { name = "begin_transaction", description = "Begin an explicit database transaction", inputSchema = new { type = "object", properties = new { isolation_level = new { type = "string", description = "Optional isolation level: read_committed, read_uncommitted, repeatable_read, serializable, chaos, unspecified" } } } },
                new { name = "start_transaction", description = "Alias for begin_transaction", inputSchema = new { type = "object", properties = new { isolation_level = new { type = "string", description = "Optional isolation level: read_committed, read_uncommitted, repeatable_read, serializable, chaos, unspecified" } } } },
                new { name = "commit_transaction", description = "Commit the active transaction", inputSchema = new { type = "object", properties = new { } } },
                new { name = "rollback_transaction", description = "Rollback the active transaction", inputSchema = new { type = "object", properties = new { } } },
                new { name = "transaction_status", description = "Get status for the current transaction", inputSchema = new { type = "object", properties = new { } } },
                new { name = "describe_table", description = "Describe a table schema including columns, nullability, defaults, and primary key columns.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "create_table", description = "Create a new table in the database", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, fields = new { type = "array", items = new { type = "object", properties = new { name = new { type = "string" }, type = new { type = "string" }, size = new { type = "integer" }, required = new { type = "boolean" }, allow_zero_length = new { type = "boolean" } } } } }, required = new string[] { "table_name", "fields" } } },
                new { name = "delete_table", description = "Delete a table from the database", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "add_field", description = "Add a field to an existing table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, field_type = new { type = "string" }, size = new { type = "integer" }, required = new { type = "boolean" }, allow_zero_length = new { type = "boolean" } }, required = new string[] { "table_name", "field_name", "field_type" } } },
                new { name = "alter_field", description = "Alter an existing field definition on a table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, new_field_type = new { type = "string" }, new_size = new { type = "integer" } }, required = new string[] { "table_name", "field_name", "new_field_type" } } },
                new { name = "drop_field", description = "Drop a field from a table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "rename_table", description = "Rename an existing table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, new_table_name = new { type = "string" } }, required = new string[] { "table_name", "new_table_name" } } },
                new { name = "rename_field", description = "Rename an existing field on a table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, new_field_name = new { type = "string" } }, required = new string[] { "table_name", "field_name", "new_field_name" } } },
                new { name = "get_indexes", description = "Get indexes for a table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "create_index", description = "Create an index on one or more columns", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, index_name = new { type = "string" }, columns = new { type = "array", items = new { type = "string" } }, unique = new { type = "boolean" } }, required = new string[] { "table_name", "index_name", "columns" } } },
                new { name = "delete_index", description = "Delete an index from a table", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, index_name = new { type = "string" } }, required = new string[] { "table_name", "index_name" } } },
                new { name = "launch_access", description = "Launch Microsoft Access application", inputSchema = new { type = "object", properties = new { } } },
                new { name = "close_access", description = "Close Microsoft Access application", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_forms", description = "Get list of all forms in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_reports", description = "Get list of all reports in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_macros", description = "Get list of all macros in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_modules", description = "Get list of all modules in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "open_form", description = "Open a form in Access", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "close_form", description = "Close a form in Access", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "open_report", description = "Open a report in Access", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "close_report", description = "Close a report in Access", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "run_macro", description = "Run an Access macro", inputSchema = new { type = "object", properties = new { macro_name = new { type = "string" } }, required = new string[] { "macro_name" } } },
                new { name = "create_macro", description = "Create a macro from text representation", inputSchema = new { type = "object", properties = new { macro_name = new { type = "string" }, macro_data = new { type = "string" } }, required = new string[] { "macro_name", "macro_data" } } },
                new { name = "update_macro", description = "Update an existing macro from text representation", inputSchema = new { type = "object", properties = new { macro_name = new { type = "string" }, macro_data = new { type = "string" } }, required = new string[] { "macro_name", "macro_data" } } },
                new { name = "export_macro_to_text", description = "Export a macro to text format", inputSchema = new { type = "object", properties = new { macro_name = new { type = "string" } }, required = new string[] { "macro_name" } } },
                new { name = "import_macro_from_text", description = "Import or replace a macro from text format", inputSchema = new { type = "object", properties = new { macro_name = new { type = "string" }, macro_data = new { type = "string" }, overwrite = new { type = "boolean" } }, required = new string[] { "macro_name", "macro_data" } } },
                new { name = "delete_macro", description = "Delete a macro from the database", inputSchema = new { type = "object", properties = new { macro_name = new { type = "string" } }, required = new string[] { "macro_name" } } },
                new { name = "get_vba_projects", description = "Get list of VBA projects", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_vba_code", description = "Get VBA code from a module", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" } }, required = new string[] { "project_name", "module_name" } } },
                new { name = "set_vba_code", description = "Set VBA code in a module", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, code = new { type = "string" } }, required = new string[] { "project_name", "module_name", "code" } } },
                new { name = "add_vba_procedure", description = "Add a VBA procedure to a module", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, procedure_name = new { type = "string" }, code = new { type = "string" } }, required = new string[] { "project_name", "module_name", "procedure_name", "code" } } },
                new { name = "compile_vba", description = "Compile VBA code", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_system_tables", description = "Get list of system tables", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_object_metadata", description = "Get metadata for database objects", inputSchema = new { type = "object", properties = new { } } },
                new { name = "form_exists", description = "Check if a form exists", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_form_controls", description = "Get list of controls in a form", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_control_properties", description = "Get properties of a control", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "form_name", "control_name" } } },
                new { name = "set_control_property", description = "Set a property of a control", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "form_name", "control_name", "property_name", "value" } } },
                new { name = "get_report_controls", description = "Get list of controls in a report", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "get_report_control_properties", description = "Get properties of a report control", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "report_name", "control_name" } } },
                new { name = "set_report_control_property", description = "Set a property of a report control", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, control_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "report_name", "control_name", "property_name", "value" } } },
                new { name = "export_form_to_text", description = "Export a form to text format", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "form_name" } } },
                new { name = "import_form_from_text", description = "Import a form from text format", inputSchema = new { type = "object", properties = new { form_data = new { type = "string" }, form_name = new { type = "string", description = "Optional form name override. Required for some access_text payloads." }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "form_data" } } },
                new { name = "delete_form", description = "Delete a form from the database", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "export_report_to_text", description = "Export a report to text format", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "report_name" } } },
                new { name = "import_report_from_text", description = "Import a report from text format", inputSchema = new { type = "object", properties = new { report_data = new { type = "string" }, report_name = new { type = "string", description = "Optional report name override. Required for some access_text payloads." }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "report_data" } } },
                new { name = "delete_report", description = "Delete a report from the database", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } }
            }
        };
    }

    static object HandleToolsCall(AccessInteropService accessService, JsonElement arguments)
    {
        if (arguments.ValueKind != JsonValueKind.Object || !arguments.TryGetProperty("name", out var toolNameElement))
            return new { success = false, error = "Missing required tools/call parameter: name" };

        var toolName = toolNameElement.GetString();
        if (string.IsNullOrWhiteSpace(toolName))
            return new { success = false, error = "Tool name is empty" };

        var toolArguments = GetToolArguments(arguments);

        return toolName switch
        {
            "connect_access" => HandleConnectAccess(accessService, toolArguments),
            "disconnect_access" => HandleDisconnectAccess(accessService, toolArguments),
            "is_connected" => HandleIsConnected(accessService, toolArguments),
            "get_tables" => HandleGetTables(accessService, toolArguments),
            "get_queries" => HandleGetQueries(accessService, toolArguments),
            "get_relationships" => HandleGetRelationships(accessService, toolArguments),
            "list_linked_tables" => HandleListLinkedTables(accessService, toolArguments),
            "link_table" => HandleLinkTable(accessService, toolArguments),
            "create_linked_table" => HandleLinkTable(accessService, toolArguments),
            "refresh_link" => HandleRefreshLink(accessService, toolArguments),
            "refresh_linked_table" => HandleRefreshLink(accessService, toolArguments),
            "relink_table" => HandleRelinkTable(accessService, toolArguments),
            "update_linked_table" => HandleRelinkTable(accessService, toolArguments),
            "unlink_table" => HandleUnlinkTable(accessService, toolArguments),
            "delete_linked_table" => HandleUnlinkTable(accessService, toolArguments),
            "create_query" => HandleCreateQuery(accessService, toolArguments),
            "update_query" => HandleUpdateQuery(accessService, toolArguments),
            "delete_query" => HandleDeleteQuery(accessService, toolArguments),
            "create_relationship" => HandleCreateRelationship(accessService, toolArguments),
            "update_relationship" => HandleUpdateRelationship(accessService, toolArguments),
            "delete_relationship" => HandleDeleteRelationship(accessService, toolArguments),
            "execute_sql" => HandleExecuteSql(accessService, toolArguments),
            "execute_query_md" => HandleExecuteQueryMd(accessService, toolArguments),
            "begin_transaction" => HandleBeginTransaction(accessService, toolArguments),
            "start_transaction" => HandleBeginTransaction(accessService, toolArguments),
            "commit_transaction" => HandleCommitTransaction(accessService, toolArguments),
            "rollback_transaction" => HandleRollbackTransaction(accessService, toolArguments),
            "transaction_status" => HandleTransactionStatus(accessService, toolArguments),
            "describe_table" => HandleDescribeTable(accessService, toolArguments),
            "create_table" => HandleCreateTable(accessService, toolArguments),
            "delete_table" => HandleDeleteTable(accessService, toolArguments),
            "add_field" => HandleAddField(accessService, toolArguments),
            "alter_field" => HandleAlterField(accessService, toolArguments),
            "drop_field" => HandleDropField(accessService, toolArguments),
            "rename_table" => HandleRenameTable(accessService, toolArguments),
            "rename_field" => HandleRenameField(accessService, toolArguments),
            "get_indexes" => HandleGetIndexes(accessService, toolArguments),
            "create_index" => HandleCreateIndex(accessService, toolArguments),
            "delete_index" => HandleDeleteIndex(accessService, toolArguments),
            "launch_access" => HandleLaunchAccess(accessService, toolArguments),
            "close_access" => HandleCloseAccess(accessService, toolArguments),
            "get_forms" => HandleGetForms(accessService, toolArguments),
            "get_reports" => HandleGetReports(accessService, toolArguments),
            "get_macros" => HandleGetMacros(accessService, toolArguments),
            "get_modules" => HandleGetModules(accessService, toolArguments),
            "open_form" => HandleOpenForm(accessService, toolArguments),
            "close_form" => HandleCloseForm(accessService, toolArguments),
            "open_report" => HandleOpenReport(accessService, toolArguments),
            "close_report" => HandleCloseReport(accessService, toolArguments),
            "run_macro" => HandleRunMacro(accessService, toolArguments),
            "create_macro" => HandleCreateMacro(accessService, toolArguments),
            "update_macro" => HandleUpdateMacro(accessService, toolArguments),
            "export_macro_to_text" => HandleExportMacroToText(accessService, toolArguments),
            "import_macro_from_text" => HandleImportMacroFromText(accessService, toolArguments),
            "delete_macro" => HandleDeleteMacro(accessService, toolArguments),
            "get_vba_projects" => HandleGetVBAProjects(accessService, toolArguments),
            "get_vba_code" => HandleGetVBACode(accessService, toolArguments),
            "set_vba_code" => HandleSetVBACode(accessService, toolArguments),
            "add_vba_procedure" => HandleAddVBAProcedure(accessService, toolArguments),
            "compile_vba" => HandleCompileVBA(accessService, toolArguments),
            "get_system_tables" => HandleGetSystemTables(accessService, toolArguments),
            "get_object_metadata" => HandleGetObjectMetadata(accessService, toolArguments),
            "form_exists" => HandleFormExists(accessService, toolArguments),
            "get_form_controls" => HandleGetFormControls(accessService, toolArguments),
            "get_control_properties" => HandleGetControlProperties(accessService, toolArguments),
            "set_control_property" => HandleSetControlProperty(accessService, toolArguments),
            "get_report_controls" => HandleGetReportControls(accessService, toolArguments),
            "get_report_control_properties" => HandleGetReportControlProperties(accessService, toolArguments),
            "set_report_control_property" => HandleSetReportControlProperty(accessService, toolArguments),
            "export_form_to_text" => HandleExportFormToText(accessService, toolArguments),
            "import_form_from_text" => HandleImportFormFromText(accessService, toolArguments),
            "delete_form" => HandleDeleteForm(accessService, toolArguments),
            "export_report_to_text" => HandleExportReportToText(accessService, toolArguments),
            "import_report_from_text" => HandleImportReportFromText(accessService, toolArguments),
            "delete_report" => HandleDeleteReport(accessService, toolArguments),
            _ => new { success = false, error = $"Unknown tool: {toolName}" }
        };
    }

    static object HandleConnectAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            string? databasePath = null;
            if (arguments.ValueKind == JsonValueKind.Object &&
                arguments.TryGetProperty("database_path", out var pathElement) &&
                pathElement.ValueKind == JsonValueKind.String)
            {
                databasePath = pathElement.GetString();
            }

            databasePath ??= ResolveDatabasePath();
            if (string.IsNullOrWhiteSpace(databasePath))
            {
                return new
                {
                    success = false,
                    error = "No database path was provided or discoverable. Set ACCESS_DATABASE_PATH or place a .accdb/.mdb file in Documents."
                };
            }
            
            // Check if database file exists
            if (!File.Exists(databasePath))
                return new { success = false, error = $"Database file not found: {databasePath}" };

            accessService.Connect(databasePath);
            
            // Verify connection was successful
            if (!accessService.IsConnected)
                return new { success = false, error = "Failed to establish database connection" };
                
            return new { success = true, message = $"Connected to {databasePath}", connected = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("connect_access", ex);
        }
    }

    static object HandleDisconnectAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.Disconnect();
            return new { success = true, message = "Disconnected from database" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleIsConnected(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var isConnected = accessService.IsConnected;
            return new { success = true, connected = isConnected };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetTables(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var tables = accessService.GetTables();
            return new { success = true, tables = tables.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_tables", ex);
        }
    }

    static object HandleGetQueries(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var queries = accessService.GetQueries();
            return new { success = true, queries = queries.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_queries", ex);
        }
    }

    static object HandleGetRelationships(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var relationships = accessService.GetRelationships();
            return new { success = true, relationships = relationships.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_relationships", ex);
        }
    }

    static object HandleListLinkedTables(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var linkedTables = accessService.GetLinkedTables();
            return new { success = true, linked_tables = linkedTables.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("list_linked_tables", ex);
        }
    }

    static object HandleLinkTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "source_database_path", out var sourceDatabasePath, out var sourceDatabasePathError))
                return sourceDatabasePathError;
            if (!TryGetRequiredString(arguments, "source_table_name", out var sourceTableName, out var sourceTableNameError))
                return sourceTableNameError;

            _ = TryGetOptionalString(arguments, "connect_string", out var connectString);
            var overwrite = GetOptionalBool(arguments, "overwrite", false);

            var linkedTable = accessService.LinkTable(
                tableName,
                sourceDatabasePath,
                sourceTableName,
                string.IsNullOrWhiteSpace(connectString) ? null : connectString,
                overwrite);

            return new
            {
                success = true,
                message = $"Linked table {linkedTable.Name}",
                table = linkedTable
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("link_table", ex);
        }
    }

    static object HandleRefreshLink(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var linkedTable = accessService.RefreshLink(tableName);
            return new
            {
                success = true,
                message = $"Refreshed link for {tableName}",
                table = linkedTable
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("refresh_link", ex);
        }
    }

    static object HandleRelinkTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "source_database_path", out var sourceDatabasePath, out var sourceDatabasePathError))
                return sourceDatabasePathError;

            _ = TryGetOptionalString(arguments, "source_table_name", out var sourceTableName);
            _ = TryGetOptionalString(arguments, "connect_string", out var connectString);

            var linkedTable = accessService.RelinkTable(
                tableName,
                sourceDatabasePath,
                string.IsNullOrWhiteSpace(sourceTableName) ? null : sourceTableName,
                string.IsNullOrWhiteSpace(connectString) ? null : connectString);

            return new
            {
                success = true,
                message = $"Relinked table {linkedTable.Name}",
                table = linkedTable
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("relink_table", ex);
        }
    }

    static object HandleUnlinkTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            accessService.UnlinkTable(tableName);
            return new
            {
                success = true,
                message = $"Unlinked table {tableName}",
                table_name = tableName
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("unlink_table", ex);
        }
    }

    static object HandleCreateQuery(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "query_name", out var queryName, out var queryNameError))
                return queryNameError;
            if (!TryGetRequiredString(arguments, "sql", out var sql, out var sqlError))
                return sqlError;

            accessService.CreateQuery(queryName, sql);
            return new { success = true, message = $"Created query {queryName}", query_name = queryName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleUpdateQuery(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "query_name", out var queryName, out var queryNameError))
                return queryNameError;
            if (!TryGetRequiredString(arguments, "sql", out var sql, out var sqlError))
                return sqlError;

            accessService.UpdateQuery(queryName, sql);
            return new { success = true, message = $"Updated query {queryName}", query_name = queryName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteQuery(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "query_name", out var queryName, out var queryNameError))
                return queryNameError;

            accessService.DeleteQuery(queryName);
            return new { success = true, message = $"Deleted query {queryName}", query_name = queryName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCreateRelationship(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "foreign_table_name", out var foreignTableName, out var foreignTableNameError))
                return foreignTableNameError;
            if (!TryGetRequiredString(arguments, "foreign_field_name", out var foreignFieldName, out var foreignFieldNameError))
                return foreignFieldNameError;

            _ = TryGetOptionalString(arguments, "relationship_name", out var relationshipName);
            var enforceIntegrity = GetOptionalBool(arguments, "enforce_integrity", true);
            var cascadeUpdate = GetOptionalBool(arguments, "cascade_update", false);
            var cascadeDelete = GetOptionalBool(arguments, "cascade_delete", false);

            var createdRelationshipName = accessService.CreateRelationship(
                tableName,
                fieldName,
                foreignTableName,
                foreignFieldName,
                relationshipName,
                enforceIntegrity,
                cascadeUpdate,
                cascadeDelete);

            return new
            {
                success = true,
                message = $"Created relationship {createdRelationshipName}",
                relationship_name = createdRelationshipName
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleUpdateRelationship(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "relationship_name", out var relationshipName, out var relationshipNameError))
                return relationshipNameError;
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "foreign_table_name", out var foreignTableName, out var foreignTableNameError))
                return foreignTableNameError;
            if (!TryGetRequiredString(arguments, "foreign_field_name", out var foreignFieldName, out var foreignFieldNameError))
                return foreignFieldNameError;

            var enforceIntegrity = GetOptionalBool(arguments, "enforce_integrity", true);
            var cascadeUpdate = GetOptionalBool(arguments, "cascade_update", false);
            var cascadeDelete = GetOptionalBool(arguments, "cascade_delete", false);

            var updatedRelationshipName = accessService.UpdateRelationship(
                relationshipName,
                tableName,
                fieldName,
                foreignTableName,
                foreignFieldName,
                enforceIntegrity,
                cascadeUpdate,
                cascadeDelete);

            return new
            {
                success = true,
                message = $"Updated relationship {updatedRelationshipName}",
                relationship_name = updatedRelationshipName
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteRelationship(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "relationship_name", out var relationshipName, out var relationshipNameError))
                return relationshipNameError;

            accessService.DeleteRelationship(relationshipName);
            return new { success = true, message = $"Deleted relationship {relationshipName}", relationship_name = relationshipName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleExecuteSql(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!arguments.TryGetProperty("sql", out var sqlElement) || sqlElement.ValueKind != JsonValueKind.String)
                return new { success = false, error = "SQL is required" };

            var sql = sqlElement.GetString();
            if (string.IsNullOrWhiteSpace(sql))
                return new { success = false, error = "SQL is required" };

            var maxRows = 200;
            if (arguments.TryGetProperty("max_rows", out var maxRowsElement) &&
                maxRowsElement.ValueKind == JsonValueKind.Number &&
                maxRowsElement.TryGetInt32(out var parsedMaxRows))
            {
                maxRows = parsedMaxRows;
            }

            if (maxRows <= 0)
                return new { success = false, error = "max_rows must be greater than 0" };

            var result = accessService.ExecuteSql(sql, maxRows);

            if (result.IsQuery)
            {
                return new
                {
                    success = true,
                    is_query = true,
                    columns = result.Columns,
                    rows = result.Rows,
                    row_count = result.RowCount,
                    truncated = result.Truncated,
                    max_rows = maxRows
                };
            }

            return new
            {
                success = true,
                is_query = false,
                rows_affected = result.RowsAffected
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("execute_sql", ex);
        }
    }

    static object HandleExecuteQueryMd(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!arguments.TryGetProperty("sql", out var sqlElement) || sqlElement.ValueKind != JsonValueKind.String)
                return new { success = false, error = "SQL is required" };

            var sql = sqlElement.GetString();
            if (string.IsNullOrWhiteSpace(sql))
                return new { success = false, error = "SQL is required" };

            var maxRows = 100;
            if (arguments.TryGetProperty("max_rows", out var maxRowsElement) &&
                maxRowsElement.ValueKind == JsonValueKind.Number &&
                maxRowsElement.TryGetInt32(out var parsedMaxRows))
            {
                maxRows = parsedMaxRows;
            }

            if (maxRows <= 0)
                return new { success = false, error = "max_rows must be greater than 0" };

            var markdown = accessService.ExecuteQueryMarkdown(sql, maxRows);
            return new { success = true, markdown };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("execute_query_md", ex);
        }
    }

    static object HandleBeginTransaction(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            string? isolationLevel = null;
            if (arguments.TryGetProperty("isolation_level", out var isolationLevelElement))
            {
                if (isolationLevelElement.ValueKind != JsonValueKind.String)
                    return new { success = false, error = "isolation_level must be a string when provided" };

                var candidate = isolationLevelElement.GetString();
                if (!string.IsNullOrWhiteSpace(candidate))
                    isolationLevel = candidate;
            }

            var transaction = accessService.BeginTransaction(isolationLevel);
            return new
            {
                success = true,
                message = "Transaction started",
                transaction
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("begin_transaction", ex);
        }
    }

    static object HandleCommitTransaction(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var transaction = accessService.CommitTransaction();
            return new
            {
                success = true,
                message = "Transaction committed",
                transaction
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("commit_transaction", ex);
        }
    }

    static object HandleRollbackTransaction(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var transaction = accessService.RollbackTransaction();
            return new
            {
                success = true,
                message = "Transaction rolled back",
                transaction
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("rollback_transaction", ex);
        }
    }

    static object HandleTransactionStatus(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var transaction = accessService.GetTransactionStatus();
            return new
            {
                success = true,
                connected = accessService.IsConnected,
                transaction
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("transaction_status", ex);
        }
    }

    static object HandleDescribeTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!arguments.TryGetProperty("table_name", out var tableNameElement) || tableNameElement.ValueKind != JsonValueKind.String)
                return new { success = false, error = "table_name is required" };

            var tableName = tableNameElement.GetString();
            if (string.IsNullOrWhiteSpace(tableName))
                return new { success = false, error = "table_name is required" };

            var description = accessService.DescribeTable(tableName);
            return new { success = true, table = description };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("describe_table", ex);
        }
    }

    static object HandleCreateTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var tableName = arguments.GetProperty("table_name").GetString();
            if (string.IsNullOrEmpty(tableName))
                return new { success = false, error = "Table name is required" };
                
            var fieldsArray = arguments.GetProperty("fields");
            var fields = new List<FieldInfo>();

            foreach (var fieldElement in fieldsArray.EnumerateArray())
            {
                fields.Add(new FieldInfo
                {
                    Name = fieldElement.GetProperty("name").GetString() ?? "",
                    Type = fieldElement.GetProperty("type").GetString() ?? "",
                    Size = fieldElement.GetProperty("size").GetInt32(),
                    Required = fieldElement.GetProperty("required").GetBoolean(),
                    AllowZeroLength = fieldElement.GetProperty("allow_zero_length").GetBoolean()
                });
            }

            accessService.CreateTable(tableName, fields);
            return new { success = true, message = $"Created table {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var tableName = arguments.GetProperty("table_name").GetString();
            if (string.IsNullOrEmpty(tableName))
                return new { success = false, error = "Table name is required" };
                
            accessService.DeleteTable(tableName);
            return new { success = true, message = $"Deleted table {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleAddField(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "field_name", "name" }, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "field_type", "type" }, "field_type", out var fieldType, out var fieldTypeError))
                return fieldTypeError;

            var field = new FieldInfo
            {
                Name = fieldName,
                Type = fieldType,
                Size = GetOptionalIntFromAliases(arguments, new[] { "size" }, 0),
                Required = GetOptionalBool(arguments, "required", false),
                AllowZeroLength = GetOptionalBool(arguments, "allow_zero_length", false)
            };

            accessService.AddField(tableName, field);
            return new { success = true, message = $"Added field {fieldName} to {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleAlterField(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "field_name" }, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "new_field_type", "field_type", "type" }, "new_field_type", out var newFieldType, out var newFieldTypeError))
                return newFieldTypeError;

            var newSize = GetOptionalIntFromAliases(arguments, new[] { "new_size", "size" }, 0);
            accessService.AlterField(tableName, fieldName, newFieldType, newSize);
            return new { success = true, message = $"Altered field {fieldName} on {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDropField(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "field_name" }, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;

            accessService.DropField(tableName, fieldName);
            return new { success = true, message = $"Dropped field {fieldName} from {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleRenameTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "table_name", "old_table_name" }, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "new_table_name", out var newTableName, out var newTableNameError))
                return newTableNameError;

            accessService.RenameTable(tableName, newTableName);
            return new { success = true, message = $"Renamed table {tableName} to {newTableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleRenameField(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "field_name", "old_field_name" }, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "new_field_name", out var newFieldName, out var newFieldNameError))
                return newFieldNameError;

            accessService.RenameField(tableName, fieldName, newFieldName);
            return new { success = true, message = $"Renamed field {fieldName} to {newFieldName} on {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetIndexes(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var indexes = accessService.GetIndexes(tableName);
            return new { success = true, indexes = indexes.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCreateIndex(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "index_name", out var indexName, out var indexNameError))
                return indexNameError;
            if (!TryGetRequiredStringArray(arguments, "columns", out var columns, out var columnsError))
                return columnsError;

            var unique = GetOptionalBool(arguments, "unique", false);
            accessService.CreateIndex(tableName, indexName, columns, unique);
            return new { success = true, message = $"Created index {indexName}", index_name = indexName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteIndex(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "index_name", out var indexName, out var indexNameError))
                return indexNameError;

            accessService.DeleteIndex(tableName, indexName);
            return new { success = true, message = $"Deleted index {indexName}", index_name = indexName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleLaunchAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.LaunchAccess();
            return new { success = true, message = "Access launched successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCloseAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.CloseAccess();
            return new { success = true, message = "Access closed successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetForms(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var forms = accessService.GetForms();
            return new { success = true, forms = forms.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetReports(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reports = accessService.GetReports();
            return new { success = true, reports = reports.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetMacros(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var macros = accessService.GetMacros();
            return new { success = true, macros = macros.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetModules(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var modules = accessService.GetModules();
            return new { success = true, modules = modules.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleOpenForm(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            accessService.OpenForm(formName);
            return new { success = true, message = $"Opened form {formName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCloseForm(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            accessService.CloseForm(formName);
            return new { success = true, message = $"Closed form {formName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetVBAProjects(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projects = accessService.GetVBAProjects();
            return new { success = true, projects = projects.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetVBACode(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projectName = arguments.GetProperty("project_name").GetString();
            var moduleName = arguments.GetProperty("module_name").GetString();
            
            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName))
                return new { success = false, error = "Project name and module name are required" };
                
            var code = accessService.GetVBACode(projectName, moduleName);
            return new { success = true, code = code };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleSetVBACode(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projectName = arguments.GetProperty("project_name").GetString();
            var moduleName = arguments.GetProperty("module_name").GetString();
            var code = arguments.GetProperty("code").GetString();
            
            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(code))
                return new { success = false, error = "Project name, module name, and code are required" };
                
            accessService.SetVBACode(projectName, moduleName, code);
            return new { success = true, message = $"Updated VBA code in {moduleName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleAddVBAProcedure(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projectName = arguments.GetProperty("project_name").GetString();
            var moduleName = arguments.GetProperty("module_name").GetString();
            var procedureName = arguments.GetProperty("procedure_name").GetString();
            var code = arguments.GetProperty("code").GetString();
            
            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName) || 
                string.IsNullOrEmpty(procedureName) || string.IsNullOrEmpty(code))
                return new { success = false, error = "All parameters are required" };
                
            accessService.AddVBAProcedure(projectName, moduleName, procedureName, code);
            return new { success = true, message = $"Added VBA procedure {procedureName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCompileVBA(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.CompileVBA();
            return new { success = true, message = "VBA compiled successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetSystemTables(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var systemTables = accessService.GetSystemTables();
            return new { success = true, system_tables = systemTables.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetObjectMetadata(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var metadata = accessService.GetObjectMetadata();
            return new { success = true, metadata = metadata };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleFormExists(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            var exists = accessService.FormExists(formName);
            return new { success = true, exists = exists };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetFormControls(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            var controls = accessService.GetFormControls(formName);
            return new { success = true, controls = controls.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetControlProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            var controlName = arguments.GetProperty("control_name").GetString();
            
            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName))
                return new { success = false, error = "Form name and control name are required" };
                
            var properties = accessService.GetControlProperties(formName, controlName);
            return new { success = true, properties = properties };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleSetControlProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            var controlName = arguments.GetProperty("control_name").GetString();
            var propertyName = arguments.GetProperty("property_name").GetString();
            var value = arguments.GetProperty("value").GetString();
            
            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName) || 
                string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(value))
                return new { success = false, error = "All parameters are required" };
                
            accessService.SetControlProperty(formName, controlName, propertyName, value);
            return new { success = true, message = $"Updated property {propertyName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleOpenReport(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            if (string.IsNullOrEmpty(reportName))
                return new { success = false, error = "Report name is required" };

            accessService.OpenReport(reportName);
            return new { success = true, message = $"Opened report {reportName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCloseReport(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            if (string.IsNullOrEmpty(reportName))
                return new { success = false, error = "Report name is required" };

            accessService.CloseReport(reportName);
            return new { success = true, message = $"Closed report {reportName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleRunMacro(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;

            accessService.RunMacro(macroName);
            return new { success = true, message = $"Ran macro {macroName}", macro_name = macroName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCreateMacro(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;
            if (!TryGetRequiredString(arguments, "macro_data", out var macroData, out var macroDataError))
                return macroDataError;

            accessService.CreateMacro(macroName, macroData);
            return new { success = true, message = $"Created macro {macroName}", macro_name = macroName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleUpdateMacro(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;
            if (!TryGetRequiredString(arguments, "macro_data", out var macroData, out var macroDataError))
                return macroDataError;

            accessService.UpdateMacro(macroName, macroData);
            return new { success = true, message = $"Updated macro {macroName}", macro_name = macroName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleExportMacroToText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;

            var macroData = accessService.ExportMacroToText(macroName);
            return new { success = true, macro_name = macroName, macro_data = macroData };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleImportMacroFromText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;
            if (!TryGetRequiredString(arguments, "macro_data", out var macroData, out var macroDataError))
                return macroDataError;

            var overwrite = GetOptionalBool(arguments, "overwrite", true);
            accessService.ImportMacroFromText(macroName, macroData, overwrite);
            return new { success = true, message = $"Imported macro {macroName}", macro_name = macroName, overwrite };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteMacro(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;

            accessService.DeleteMacro(macroName);
            return new { success = true, message = $"Deleted macro {macroName}", macro_name = macroName };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetReportControls(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            if (string.IsNullOrEmpty(reportName))
                return new { success = false, error = "Report name is required" };

            var controls = accessService.GetReportControls(reportName);
            return new { success = true, controls = controls.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetReportControlProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            var controlName = arguments.GetProperty("control_name").GetString();

            if (string.IsNullOrEmpty(reportName) || string.IsNullOrEmpty(controlName))
                return new { success = false, error = "Report name and control name are required" };

            var properties = accessService.GetReportControlProperties(reportName, controlName);
            return new { success = true, properties = properties };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleSetReportControlProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            var controlName = arguments.GetProperty("control_name").GetString();
            var propertyName = arguments.GetProperty("property_name").GetString();
            var value = arguments.GetProperty("value").GetString();

            if (string.IsNullOrEmpty(reportName) || string.IsNullOrEmpty(controlName) ||
                string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(value))
                return new { success = false, error = "All parameters are required" };

            accessService.SetReportControlProperty(reportName, controlName, propertyName, value);
            return new { success = true, message = $"Updated report property {propertyName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleExportFormToText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };

            if (!TryGetOptionalMode(arguments, out var mode, out var modeError))
                return modeError;

            var formData = accessService.ExportFormToText(formName, mode);
            return new { success = true, form_data = formData };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleImportFormFromText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formData = arguments.GetProperty("form_data").GetString();
            if (string.IsNullOrEmpty(formData))
                return new { success = false, error = "Form data is required" };

            if (!TryGetOptionalMode(arguments, out var mode, out var modeError))
                return modeError;

            _ = TryGetOptionalString(arguments, "form_name", out var formName);
            if (string.Equals(mode, "access_text", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(formName))
                return new { success = false, error = "form_name is required when mode is access_text" };

            accessService.ImportFormFromText(formData, mode, formName);
            return new { success = true, message = "Form imported successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteForm(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            accessService.DeleteForm(formName);
            return new { success = true, message = $"Deleted form {formName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleExportReportToText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            if (string.IsNullOrEmpty(reportName))
                return new { success = false, error = "Report name is required" };

            if (!TryGetOptionalMode(arguments, out var mode, out var modeError))
                return modeError;

            var reportData = accessService.ExportReportToText(reportName, mode);
            return new { success = true, report_data = reportData };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleImportReportFromText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportData = arguments.GetProperty("report_data").GetString();
            if (string.IsNullOrEmpty(reportData))
                return new { success = false, error = "Report data is required" };

            if (!TryGetOptionalMode(arguments, out var mode, out var modeError))
                return modeError;

            _ = TryGetOptionalString(arguments, "report_name", out var reportName);
            if (string.Equals(mode, "access_text", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(reportName))
                return new { success = false, error = "report_name is required when mode is access_text" };

            accessService.ImportReportFromText(reportData, mode, reportName);
            return new { success = true, message = "Report imported successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteReport(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            if (string.IsNullOrEmpty(reportName))
                return new { success = false, error = "Report name is required" };
                
            accessService.DeleteReport(reportName);
            return new { success = true, message = $"Deleted report {reportName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static bool TryGetRequiredString(JsonElement arguments, string propertyName, out string value, out object error)
    {
        value = string.Empty;
        if (!TryGetOptionalString(arguments, propertyName, out value) || string.IsNullOrWhiteSpace(value))
        {
            error = new { success = false, error = $"{propertyName} is required" };
            return false;
        }

        error = new { success = true };
        return true;
    }

    static bool TryGetOptionalString(JsonElement arguments, string propertyName, out string value)
    {
        value = string.Empty;
        if (!arguments.TryGetProperty(propertyName, out var element))
            return false;

        if (element.ValueKind != JsonValueKind.String)
            return false;

        value = element.GetString() ?? string.Empty;
        return true;
    }

    static bool TryGetOptionalMode(JsonElement arguments, out string? mode, out object error)
    {
        mode = null;

        if (!arguments.TryGetProperty("mode", out var modeElement))
        {
            error = new { success = true };
            return true;
        }

        if (modeElement.ValueKind != JsonValueKind.String)
        {
            error = new { success = false, error = "mode must be a string when provided" };
            return false;
        }

        var rawMode = modeElement.GetString();
        if (string.IsNullOrWhiteSpace(rawMode))
        {
            error = new { success = true };
            return true;
        }

        var normalizedMode = rawMode.Trim().ToLowerInvariant();
        if (normalizedMode != "json" && normalizedMode != "access_text")
        {
            error = new { success = false, error = "mode must be either 'json' or 'access_text'" };
            return false;
        }

        mode = normalizedMode;
        error = new { success = true };
        return true;
    }

    static bool TryGetRequiredStringFromAliases(JsonElement arguments, string[] aliases, string propertyNameForError, out string value, out object error)
    {
        value = string.Empty;

        foreach (var alias in aliases)
        {
            if (TryGetOptionalString(arguments, alias, out var candidate) && !string.IsNullOrWhiteSpace(candidate))
            {
                value = candidate;
                error = new { success = true };
                return true;
            }
        }

        error = new { success = false, error = $"{propertyNameForError} is required" };
        return false;
    }

    static int GetOptionalIntFromAliases(JsonElement arguments, string[] aliases, int defaultValue)
    {
        foreach (var alias in aliases)
        {
            if (!arguments.TryGetProperty(alias, out var element))
                continue;

            if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out var numeric))
                return numeric;

            if (element.ValueKind == JsonValueKind.String && int.TryParse(element.GetString(), out var parsed))
                return parsed;
        }

        return defaultValue;
    }

    static bool TryGetRequiredStringArray(JsonElement arguments, string propertyName, out List<string> values, out object error)
    {
        values = new List<string>();

        if (!arguments.TryGetProperty(propertyName, out var element) || element.ValueKind != JsonValueKind.Array)
        {
            error = new { success = false, error = $"{propertyName} is required" };
            return false;
        }

        foreach (var item in element.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String)
            {
                error = new { success = false, error = $"{propertyName} must be an array of strings" };
                return false;
            }

            var value = item.GetString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                values.Add(value.Trim());
            }
        }

        values = values
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (values.Count == 0)
        {
            error = new { success = false, error = $"{propertyName} must contain at least one non-empty value" };
            return false;
        }

        error = new { success = true };
        return true;
    }

    static bool GetOptionalBool(JsonElement arguments, string propertyName, bool defaultValue)
    {
        if (!arguments.TryGetProperty(propertyName, out var element))
            return defaultValue;

        return element.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number when element.TryGetInt32(out var numeric) => numeric != 0,
            JsonValueKind.String when bool.TryParse(element.GetString(), out var parsed) => parsed,
            _ => defaultValue
        };
    }

    static object BuildOperationErrorResponse(string operationName, Exception ex)
    {
        var preflight = BuildPreflightDiagnostics(ex);
        var error = BuildRemediatedErrorMessage(operationName, ex, preflight);

        return new
        {
            success = false,
            error,
            preflight = ToPreflightPayload(preflight)
        };
    }

    static PreflightDiagnostics BuildPreflightDiagnostics(Exception? ex = null)
    {
        var processBitness = Environment.Is64BitProcess ? "x64" : "x86";
        var aceProviderRegistered = IsAceOleDbProviderRegistered();
        var aceProviderIssueDetected = ex != null && HasAceProviderRegistrationIndicator(ex);
        var trustCenterActiveContentIndicator = ex != null && HasTrustCenterActiveContentIndicator(ex);

        var remediationHints = new List<string>();
        if (aceProviderIssueDetected)
        {
            remediationHints.Add($"Install Microsoft Access Database Engine (ACE) with {processBitness} bitness.");
            remediationHints.Add("If Office and MCP server bitness differ, run the server build that matches the installed ACE provider.");
        }

        if (trustCenterActiveContentIndicator)
        {
            remediationHints.Add("In Access: File > Options > Trust Center > Trust Center Settings > Trusted Locations, add the database folder.");
            remediationHints.Add("If the database was downloaded, open file properties and click Unblock before retrying.");
        }

        return new PreflightDiagnostics
        {
            ProcessBitness = processBitness,
            AceOleDbProviderRegistered = aceProviderRegistered,
            AceOleDbIssueDetected = aceProviderIssueDetected,
            TrustCenterActiveContentIndicator = trustCenterActiveContentIndicator,
            RemediationHints = remediationHints
        };
    }

    static object ToPreflightPayload(PreflightDiagnostics preflight)
    {
        return new
        {
            process_bitness = preflight.ProcessBitness,
            ace_oledb_provider_registered = preflight.AceOleDbProviderRegistered,
            ace_oledb_issue_detected = preflight.AceOleDbIssueDetected,
            trust_center_active_content_indicator = preflight.TrustCenterActiveContentIndicator,
            remediation_hints = preflight.RemediationHints.ToArray()
        };
    }

    static string BuildRemediatedErrorMessage(string operationName, Exception ex, PreflightDiagnostics preflight)
    {
        if (preflight.AceOleDbIssueDetected)
        {
            return $"{operationName} failed: ACE OLEDB provider (Microsoft.ACE.OLEDB.12.0) is unavailable for this {preflight.ProcessBitness} process. Install matching Access Database Engine bitness or run a matching MCP server build. Original error: {ex.Message}";
        }

        if (preflight.TrustCenterActiveContentIndicator)
        {
            return $"{operationName} failed: Access Trust Center appears to be blocking active content for this database. Add the database folder to Trusted Locations and unblock the file if needed. Original error: {ex.Message}";
        }

        return ex.Message;
    }

    static bool IsAceOleDbProviderRegistered()
    {
        try
        {
            return Type.GetTypeFromProgID("Microsoft.ACE.OLEDB.12.0", throwOnError: false) != null;
        }
        catch
        {
            return false;
        }
    }

    static bool HasAceProviderRegistrationIndicator(Exception ex)
    {
        if (ex is COMException comException && (uint)comException.ErrorCode == 0x80040154)
            return true;

        var message = ex.Message ?? string.Empty;
        if ((message.IndexOf("Microsoft.ACE.OLEDB.12.0", StringComparison.OrdinalIgnoreCase) >= 0 &&
             message.IndexOf("not registered", StringComparison.OrdinalIgnoreCase) >= 0) ||
            message.IndexOf("provider cannot be found", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("class not registered", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return true;
        }

        return ex.InnerException != null && HasAceProviderRegistrationIndicator(ex.InnerException);
    }

    static bool HasTrustCenterActiveContentIndicator(Exception ex)
    {
        var message = ex.Message ?? string.Empty;
        if (message.IndexOf("active content", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("disabled mode", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("trusted location", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("security warning", StringComparison.OrdinalIgnoreCase) >= 0 ||
            message.IndexOf("has blocked", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return true;
        }

        return ex.InnerException != null && HasTrustCenterActiveContentIndicator(ex.InnerException);
    }

    sealed class PreflightDiagnostics
    {
        public string ProcessBitness { get; init; } = string.Empty;
        public bool AceOleDbProviderRegistered { get; init; }
        public bool AceOleDbIssueDetected { get; init; }
        public bool TrustCenterActiveContentIndicator { get; init; }
        public List<string> RemediationHints { get; init; } = new();
    }

    static object WrapCallToolResult(object payload)
    {
        var structuredContent = JsonSerializer.SerializeToElement(payload);
        var isError = structuredContent.TryGetProperty("success", out var successElement) &&
                      successElement.ValueKind == JsonValueKind.False;

        return new
        {
            content = new object[]
            {
                new
                {
                    type = "text",
                    text = structuredContent.GetRawText()
                }
            },
            structuredContent,
            isError
        };
    }

    static JsonElement GetToolArguments(JsonElement callParams)
    {
        if (callParams.ValueKind == JsonValueKind.Object &&
            callParams.TryGetProperty("arguments", out var args) &&
            args.ValueKind == JsonValueKind.Object)
        {
            return args;
        }

        return EmptyJsonObject;
    }

    static string? ResolveDatabasePath()
    {
        var fromEnv = Environment.GetEnvironmentVariable("ACCESS_DATABASE_PATH");
        if (!string.IsNullOrWhiteSpace(fromEnv))
            return fromEnv;

        var searchFolders = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddFolder(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;
            if (!Directory.Exists(path))
                return;
            if (!seen.Add(path))
                return;

            searchFolders.Add(path);
        }

        AddFolder(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        AddFolder(Environment.GetFolderPath(Environment.SpecialFolder.Personal));

        var userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
        if (!string.IsNullOrWhiteSpace(userProfile))
            AddFolder(Path.Combine(userProfile, "Documents"));

        var oneDrive = Environment.GetEnvironmentVariable("OneDrive");
        if (!string.IsNullOrWhiteSpace(oneDrive))
            AddFolder(Path.Combine(oneDrive, "Documents"));

        if (searchFolders.Count == 0)
            return null;

        foreach (var folder in searchFolders)
        {
            var defaultPath = Path.Combine(folder, "Database1.accdb");
            if (File.Exists(defaultPath))
                return defaultPath;
        }

        foreach (var folder in searchFolders)
        {
            var found = Directory.EnumerateFiles(folder, "*.accdb", SearchOption.TopDirectoryOnly)
                .Concat(Directory.EnumerateFiles(folder, "*.mdb", SearchOption.TopDirectoryOnly))
                .FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(found))
                return found;
        }

        return null;
    }
}

public class JsonRpcRequest
{
    [JsonPropertyName("jsonrpc")]
    public string Jsonrpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public int Id { get; set; }

    [JsonPropertyName("method")]
    public string Method { get; set; } = string.Empty;

    [JsonPropertyName("params")]
    public JsonElement Params { get; set; }
}

public class JsonRpcResponse
{
    [JsonPropertyName("jsonrpc")]
    public string Jsonrpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public JsonElement? Id { get; set; }

    [JsonPropertyName("result")]
    public object Result { get; set; } = new { };
}
