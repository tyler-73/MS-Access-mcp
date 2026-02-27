using System.Text.Json;
using System.Text.Json.Serialization;
using System.Linq;
using System.Runtime.InteropServices;
using MS.Access.MCP.Interop;

class Program
{
    static readonly JsonElement EmptyJsonObject = JsonSerializer.Deserialize<JsonElement>("{}");
    const string PodbcDefaultSchema = "AccessCatalog";

    // Logging level for MCP logging capability
    static string _minimumLogLevel = "debug";
    static readonly string[] LogLevelOrder = { "debug", "info", "notice", "warning", "error", "critical", "alert", "emergency" };
    static int LogLevelSeverity(string level) => Array.IndexOf(LogLevelOrder, level) is int i and >= 0 ? i : 0;

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
                        "resources/list" => HandleResourcesList(accessService),
                        "resources/read" => HandleResourcesRead(accessService, safeParams),
                        "resources/templates/list" => HandleResourceTemplatesList(),
                        "prompts/list" => HandlePromptsList(),
                        "prompts/get" => HandlePromptsGet(accessService, safeParams),
                        "completion/complete" => HandleCompletionComplete(accessService, safeParams),
                        "logging/setLevel" => HandleLoggingSetLevel(safeParams),
                        _ => new JsonRpcErrorSentinel { Code = -32601, Message = $"Method not found: {method}" }
                    };

                    string jsonResponse;
                    if (result is JsonRpcErrorSentinel errSentinel)
                    {
                        jsonResponse = JsonSerializer.Serialize(new JsonRpcErrorResponse
                        {
                            Id = id,
                            Error = new JsonRpcError { Code = errSentinel.Code, Message = errSentinel.Message }
                        });
                    }
                    else
                    {
                        jsonResponse = JsonSerializer.Serialize(new JsonRpcResponse
                        {
                            Id = id,
                            Result = result
                        });
                    }
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
            capabilities = new
            {
                tools = new { },
                resources = new { listChanged = true },
                prompts = new { },
                logging = new { }
            },
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
                new { name = "connect_access", description = "Connect to an Access database. Uses database_path argument, ACCESS_DATABASE_PATH env var, or first database found in Documents.", inputSchema = new { type = "object", properties = new { database_path = new { type = "string" }, database_password = new { type = "string" }, system_database_path = new { type = "string" } } } },
                new { name = "create_database", description = "Create a new Access database file (.accdb or .mdb).", inputSchema = new { type = "object", properties = new { database_path = new { type = "string" }, overwrite = new { type = "boolean" } }, required = new string[] { "database_path" } } },
                new { name = "backup_database", description = "Back up an Access database by copying it to a destination path.", inputSchema = new { type = "object", properties = new { source_database_path = new { type = "string" }, destination_database_path = new { type = "string" }, overwrite = new { type = "boolean" } }, required = new string[] { "destination_database_path" } } },
                new { name = "compact_repair_database", description = "Compact and repair an Access database. Supports in-place replacement when destination_database_path is omitted.", inputSchema = new { type = "object", properties = new { source_database_path = new { type = "string" }, destination_database_path = new { type = "string" }, overwrite = new { type = "boolean" } } } },
                new { name = "transfer_spreadsheet", description = "Import, export, or link spreadsheet data using Access DoCmd.TransferSpreadsheet.", inputSchema = new { type = "object", properties = new { transfer_type = new { type = "string", description = "import, export, link, or Access enum integer value as string" }, spreadsheet_type = new { type = "string", description = "Optional spreadsheet type name or Access enum integer value as string" }, table_name = new { type = "string" }, file_name = new { type = "string" }, has_field_names = new { type = "boolean" }, range = new { type = "string" }, use_oa = new { type = "boolean" } }, required = new string[] { "transfer_type", "table_name", "file_name" } } },
                new { name = "transfer_text", description = "Import, export, or link text/CSV data using Access DoCmd.TransferText.", inputSchema = new { type = "object", properties = new { transfer_type = new { type = "string", description = "import, export, link, or Access enum integer value as string" }, specification_name = new { type = "string" }, table_name = new { type = "string" }, file_name = new { type = "string" }, has_field_names = new { type = "boolean" }, html_table_name = new { type = "string" }, code_page = new { type = "integer" } }, required = new string[] { "transfer_type", "table_name", "file_name" } } },
                new { name = "output_to", description = "Export Access objects using DoCmd.OutputTo.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "table, query, form, report, module, etc. or Access enum integer value as string" }, object_name = new { type = "string" }, output_format = new { type = "string", description = "pdf/xlsx/etc. or Access format value" }, output_file = new { type = "string" }, auto_start = new { type = "boolean" }, template_file = new { type = "string" }, encoding = new { type = "string" }, output_quality = new { type = "string", description = "print, screen, or Access enum integer value as string" } }, required = new string[] { "object_type", "output_format" } } },
                new { name = "set_warnings", description = "Enable or disable Access action query warning dialogs (DoCmd.SetWarnings).", inputSchema = new { type = "object", properties = new { warnings_on = new { type = "boolean" } } } },
                new { name = "echo", description = "Enable or disable Access screen repainting (DoCmd.Echo).", inputSchema = new { type = "object", properties = new { echo_on = new { type = "boolean" }, status_bar_text = new { type = "string" } } } },
                new { name = "hourglass", description = "Turn the Access hourglass cursor on or off (DoCmd.Hourglass).", inputSchema = new { type = "object", properties = new { hourglass_on = new { type = "boolean" } } } },
                new { name = "goto_record", description = "Navigate to a record using DoCmd.GoToRecord.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" }, record = new { type = "string" }, offset = new { type = "string" } } } },
                new { name = "find_record", description = "Find a record using DoCmd.FindRecord.", inputSchema = new { type = "object", properties = new { find_what = new { type = "string" }, match = new { type = "string" }, match_case = new { type = "string" }, search = new { type = "string" }, search_as_formatted = new { type = "string" }, only_current_field = new { type = "string" }, find_first = new { type = "string" } }, required = new string[] { "find_what" } } },
                new { name = "apply_filter", description = "Apply a filter using DoCmd.ApplyFilter.", inputSchema = new { type = "object", properties = new { filter_name = new { type = "string" }, where_condition = new { type = "string" }, control_name = new { type = "string" } } } },
                new { name = "show_all_records", description = "Clear active filters and show all records (DoCmd.ShowAllRecords).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "maximize_window", description = "Maximize the active Access window (DoCmd.Maximize).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "minimize_window", description = "Minimize the active Access window (DoCmd.Minimize).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "restore_window", description = "Restore the active Access window (DoCmd.Restore).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "print_out", description = "Print the active object using DoCmd.PrintOut.", inputSchema = new { type = "object", properties = new { print_range = new { type = "string" }, page_from = new { type = "integer" }, page_to = new { type = "integer" }, print_quality = new { type = "string" }, copies = new { type = "integer" }, collate_copies = new { type = "boolean" } } } },
                new { name = "open_query", description = "Open a saved query in Access using DoCmd.OpenQuery.", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" }, view = new { type = "string" }, data_mode = new { type = "string" } }, required = new string[] { "query_name" } } },
                new { name = "run_sql", description = "Execute SQL with Access DoCmd.RunSQL.", inputSchema = new { type = "object", properties = new { sql = new { type = "string" }, use_transaction = new { type = "boolean" } }, required = new string[] { "sql" } } },
                new { name = "get_database_summary_properties", description = "Get Access database summary properties (Title, Author, Subject, Keywords, Comments).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "set_database_summary_properties", description = "Set Access database summary properties.", inputSchema = new { type = "object", properties = new { title = new { type = "string" }, author = new { type = "string" }, subject = new { type = "string" }, keywords = new { type = "string" }, comments = new { type = "string" } } } },
                new { name = "get_database_properties", description = "List database properties, including custom properties.", inputSchema = new { type = "object", properties = new { include_system = new { type = "boolean" } } } },
                new { name = "get_database_property", description = "Get a single database property by name.", inputSchema = new { type = "object", properties = new { property_name = new { type = "string" } }, required = new string[] { "property_name" } } },
                new { name = "set_database_property", description = "Set or create a database property.", inputSchema = new { type = "object", properties = new { property_name = new { type = "string" }, value = new { type = "string" }, property_type = new { type = "string" }, create_if_missing = new { type = "boolean" } }, required = new string[] { "property_name", "value" } } },
                new { name = "get_table_properties", description = "Get table-level properties such as description and validation settings.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "set_table_properties", description = "Set table-level properties such as description and validation settings.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, description = new { type = "string" }, validation_rule = new { type = "string" }, validation_text = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "get_query_properties", description = "Get query properties including description, SQL text, and parameters.", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" } }, required = new string[] { "query_name" } } },
                new { name = "set_query_properties", description = "Set query properties such as description and SQL text.", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" }, description = new { type = "string" }, sql = new { type = "string" } }, required = new string[] { "query_name" } } },
                new { name = "set_field_validation", description = "Set field validation rule and validation text.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, validation_rule = new { type = "string" }, validation_text = new { type = "string" } }, required = new string[] { "table_name", "field_name", "validation_rule" } } },
                new { name = "set_field_default", description = "Set field default value.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, default_value = new { type = "string" } }, required = new string[] { "table_name", "field_name", "default_value" } } },
                new { name = "set_field_input_mask", description = "Set field input mask.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, input_mask = new { type = "string" } }, required = new string[] { "table_name", "field_name", "input_mask" } } },
                new { name = "set_field_caption", description = "Set field caption.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, caption = new { type = "string" } }, required = new string[] { "table_name", "field_name", "caption" } } },
                new { name = "get_field_properties", description = "Get field properties including validation/default/input mask/caption and lookup settings.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "set_lookup_properties", description = "Set lookup properties for a field (RowSource, BoundColumn, ColumnCount, ColumnWidths, etc.).", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, row_source = new { type = "string" }, bound_column = new { type = "integer" }, column_count = new { type = "integer" }, column_widths = new { type = "string" }, limit_to_list = new { type = "boolean" }, allow_multiple_values = new { type = "boolean" }, display_control = new { type = "integer" } }, required = new string[] { "table_name", "field_name" } } },
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
                new { name = "podbc_get_schemas", description = "PyODBC-compat: retrieve schema names using Access fallback behavior.", inputSchema = new { type = "object", properties = new { Schema = new { type = "string" }, user = new { type = "string" }, password = new { type = "string" }, dsn = new { type = "string" } } } },
                new { name = "podbc_get_tables", description = "PyODBC-compat: retrieve tables with optional schema filter argument (ignored for Access).", inputSchema = new { type = "object", properties = new { Schema = new { type = "string" }, user = new { type = "string" }, password = new { type = "string" }, dsn = new { type = "string" } } } },
                new { name = "podbc_describe_table", description = "PyODBC-compat: describe a table schema using Access table metadata.", inputSchema = new { type = "object", properties = new { table = new { type = "string" }, table_name = new { type = "string" }, Schema = new { type = "string" }, user = new { type = "string" }, password = new { type = "string" }, dsn = new { type = "string" } }, required = new string[] { "table" } } },
                new { name = "podbc_filter_table_names", description = "PyODBC-compat: list tables whose names contain substring q.", inputSchema = new { type = "object", properties = new { q = new { type = "string" }, Schema = new { type = "string" }, user = new { type = "string" }, password = new { type = "string" }, dsn = new { type = "string" } }, required = new string[] { "q" } } },
                new { name = "podbc_execute_query", description = "PyODBC-compat: execute SQL and return structured results.", inputSchema = new { type = "object", properties = new { query = new { type = "string" }, sql = new { type = "string" }, max_rows = new { type = "integer" }, @params = new { type = "array", description = "Unsupported in Access fallback when non-empty." }, Schema = new { type = "string" }, user = new { type = "string" }, password = new { type = "string" }, dsn = new { type = "string" } }, required = new string[] { "query" } } },
                new { name = "podbc_execute_query_md", description = "PyODBC-compat: execute SQL and return markdown output.", inputSchema = new { type = "object", properties = new { query = new { type = "string" }, sql = new { type = "string" }, max_rows = new { type = "integer" }, @params = new { type = "array", description = "Unsupported in Access fallback when non-empty." }, Schema = new { type = "string" }, user = new { type = "string" }, password = new { type = "string" }, dsn = new { type = "string" } }, required = new string[] { "query" } } },
                new { name = "podbc_query_database", description = "PyODBC-compat alias of podbc_execute_query.", inputSchema = new { type = "object", properties = new { query = new { type = "string" }, sql = new { type = "string" }, max_rows = new { type = "integer" }, @params = new { type = "array", description = "Unsupported in Access fallback when non-empty." }, Schema = new { type = "string" }, user = new { type = "string" }, password = new { type = "string" }, dsn = new { type = "string" } }, required = new string[] { "query" } } },
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
            "create_database" => HandleCreateDatabase(accessService, toolArguments),
            "backup_database" => HandleBackupDatabase(accessService, toolArguments),
            "compact_repair_database" => HandleCompactRepairDatabase(accessService, toolArguments),
            "transfer_spreadsheet" => HandleTransferSpreadsheet(accessService, toolArguments),
            "transfer_text" => HandleTransferText(accessService, toolArguments),
            "output_to" => HandleOutputTo(accessService, toolArguments),
            "set_warnings" => HandleSetWarnings(accessService, toolArguments),
            "echo" => HandleEcho(accessService, toolArguments),
            "hourglass" => HandleHourglass(accessService, toolArguments),
            "goto_record" => HandleGoToRecord(accessService, toolArguments),
            "find_record" => HandleFindRecord(accessService, toolArguments),
            "apply_filter" => HandleApplyFilter(accessService, toolArguments),
            "show_all_records" => HandleShowAllRecords(accessService, toolArguments),
            "maximize_window" => HandleMaximizeWindow(accessService, toolArguments),
            "minimize_window" => HandleMinimizeWindow(accessService, toolArguments),
            "restore_window" => HandleRestoreWindow(accessService, toolArguments),
            "print_out" => HandlePrintOut(accessService, toolArguments),
            "open_query" => HandleOpenQuery(accessService, toolArguments),
            "run_sql" => HandleRunSqlDocmd(accessService, toolArguments),
            "get_database_summary_properties" => HandleGetDatabaseSummaryProperties(accessService, toolArguments),
            "set_database_summary_properties" => HandleSetDatabaseSummaryProperties(accessService, toolArguments),
            "get_database_properties" => HandleGetDatabaseProperties(accessService, toolArguments),
            "get_database_property" => HandleGetDatabaseProperty(accessService, toolArguments),
            "set_database_property" => HandleSetDatabaseProperty(accessService, toolArguments),
            "get_table_properties" => HandleGetTableProperties(accessService, toolArguments),
            "set_table_properties" => HandleSetTableProperties(accessService, toolArguments),
            "get_query_properties" => HandleGetQueryProperties(accessService, toolArguments),
            "set_query_properties" => HandleSetQueryProperties(accessService, toolArguments),
            "set_field_validation" => HandleSetFieldValidation(accessService, toolArguments),
            "set_field_default" => HandleSetFieldDefault(accessService, toolArguments),
            "set_field_input_mask" => HandleSetFieldInputMask(accessService, toolArguments),
            "set_field_caption" => HandleSetFieldCaption(accessService, toolArguments),
            "get_field_properties" => HandleGetFieldProperties(accessService, toolArguments),
            "set_lookup_properties" => HandleSetLookupProperties(accessService, toolArguments),
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
            "podbc_get_schemas" => HandlePodbcGetSchemas(accessService, toolArguments),
            "podbc_get_tables" => HandlePodbcGetTables(accessService, toolArguments),
            "podbc_describe_table" => HandlePodbcDescribeTable(accessService, toolArguments),
            "podbc_filter_table_names" => HandlePodbcFilterTableNames(accessService, toolArguments),
            "podbc_execute_query" => HandlePodbcExecuteQuery(accessService, toolArguments),
            "podbc_execute_query_md" => HandlePodbcExecuteQueryMd(accessService, toolArguments),
            "podbc_query_database" => HandlePodbcQueryDatabase(accessService, toolArguments),
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
            _ = TryGetOptionalString(arguments, "database_path", out var databasePath);
            _ = TryGetOptionalString(arguments, "database_password", out var databasePassword);
            _ = TryGetOptionalString(arguments, "system_database_path", out var systemDatabasePath);

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

            accessService.Connect(
                databasePath,
                string.IsNullOrWhiteSpace(databasePassword) ? null : databasePassword,
                string.IsNullOrWhiteSpace(systemDatabasePath) ? null : systemDatabasePath);

            // Verify connection was successful
            if (!accessService.IsConnected)
            {
                SendLogNotification("error", "connection", new { databasePath, message = "Failed to establish connection" });
                return new { success = false, error = "Failed to establish database connection" };
            }

            SendLogNotification("info", "connection", new { databasePath = accessService.CurrentDatabasePath ?? databasePath, message = "Connected successfully" });
            return new
            {
                success = true,
                message = $"Connected to {databasePath}",
                connected = true,
                database_path = accessService.CurrentDatabasePath ?? databasePath,
                secured = !string.IsNullOrWhiteSpace(databasePassword),
                system_database_path = string.IsNullOrWhiteSpace(systemDatabasePath) ? null : systemDatabasePath
            };
        }
        catch (Exception ex)
        {
            SendLogNotification("error", "connection", new { error = ex.Message });
            return BuildOperationErrorResponse("connect_access", ex);
        }
    }

    static object HandleCreateDatabase(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "database_path", out var databasePath, out var databasePathError))
                return databasePathError;

            var overwrite = GetOptionalBool(arguments, "overwrite", false);
            var result = accessService.CreateDatabase(databasePath, overwrite);

            return new
            {
                success = true,
                message = $"Created database at {result.DatabasePath}",
                database_path = result.DatabasePath,
                existed_before = result.ExistedBefore,
                size_bytes = result.SizeBytes,
                last_write_time_utc = result.LastWriteTimeUtc
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("create_database", ex);
        }
    }

    static object HandleBackupDatabase(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "source_database_path", out var sourceDatabasePath);
            if (string.IsNullOrWhiteSpace(sourceDatabasePath))
                sourceDatabasePath = accessService.CurrentDatabasePath ?? string.Empty;

            if (string.IsNullOrWhiteSpace(sourceDatabasePath))
            {
                return new
                {
                    success = false,
                    error = "source_database_path is required when there is no active database connection"
                };
            }

            if (!TryGetRequiredString(arguments, "destination_database_path", out var destinationDatabasePath, out var destinationDatabasePathError))
                return destinationDatabasePathError;

            var overwrite = GetOptionalBool(arguments, "overwrite", false);
            var result = accessService.BackupDatabase(sourceDatabasePath, destinationDatabasePath, overwrite);

            return new
            {
                success = true,
                message = $"Backed up {result.SourceDatabasePath} to {result.DestinationDatabasePath}",
                source_database_path = result.SourceDatabasePath,
                destination_database_path = result.DestinationDatabasePath,
                bytes_copied = result.BytesCopied,
                source_last_write_time_utc = result.SourceLastWriteTimeUtc,
                destination_last_write_time_utc = result.DestinationLastWriteTimeUtc,
                operated_on_connected_database = result.OperatedOnConnectedDatabase
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("backup_database", ex);
        }
    }

    static object HandleCompactRepairDatabase(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "source_database_path", out var sourceDatabasePath);
            if (string.IsNullOrWhiteSpace(sourceDatabasePath))
                sourceDatabasePath = accessService.CurrentDatabasePath ?? string.Empty;

            if (string.IsNullOrWhiteSpace(sourceDatabasePath))
            {
                return new
                {
                    success = false,
                    error = "source_database_path is required when there is no active database connection"
                };
            }

            _ = TryGetOptionalString(arguments, "destination_database_path", out var destinationDatabasePath);
            var overwrite = GetOptionalBool(arguments, "overwrite", false);

            var result = accessService.CompactRepairDatabase(
                sourceDatabasePath,
                string.IsNullOrWhiteSpace(destinationDatabasePath) ? null : destinationDatabasePath,
                overwrite);

            return new
            {
                success = true,
                message = result.InPlace
                    ? $"Compacted and repaired {result.SourceDatabasePath} in place"
                    : $"Compacted and repaired {result.SourceDatabasePath} to {result.DestinationDatabasePath}",
                source_database_path = result.SourceDatabasePath,
                destination_database_path = result.DestinationDatabasePath,
                in_place = result.InPlace,
                source_size_bytes = result.SourceSizeBytes,
                destination_size_bytes = result.DestinationSizeBytes,
                destination_last_write_time_utc = result.DestinationLastWriteTimeUtc,
                operated_on_connected_database = result.OperatedOnConnectedDatabase
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("compact_repair_database", ex);
        }
    }

    static object HandleTransferSpreadsheet(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "transfer_type", out var transferType, out var transferTypeError))
                return transferTypeError;
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "file_name", out var fileName, out var fileNameError))
                return fileNameError;

            _ = TryGetOptionalString(arguments, "spreadsheet_type", out var spreadsheetType);
            _ = TryGetOptionalString(arguments, "range", out var range);
            var hasFieldNames = GetOptionalBool(arguments, "has_field_names", true);
            var useOA = GetOptionalBool(arguments, "use_oa", false);

            var result = accessService.TransferSpreadsheet(
                transferType,
                tableName,
                fileName,
                string.IsNullOrWhiteSpace(spreadsheetType) ? null : spreadsheetType,
                hasFieldNames,
                string.IsNullOrWhiteSpace(range) ? null : range,
                useOA);

            return new
            {
                success = true,
                transfer_type = result.TransferType,
                spreadsheet_type = result.SpreadsheetType,
                table_name = result.TableName,
                file_name = result.FileName,
                has_field_names = result.HasFieldNames,
                range = result.Range,
                use_oa = result.UseOA
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("transfer_spreadsheet", ex);
        }
    }

    static object HandleTransferText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "transfer_type", out var transferType, out var transferTypeError))
                return transferTypeError;
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "file_name", out var fileName, out var fileNameError))
                return fileNameError;

            _ = TryGetOptionalString(arguments, "specification_name", out var specificationName);
            _ = TryGetOptionalString(arguments, "html_table_name", out var htmlTableName);
            var hasFieldNames = GetOptionalBool(arguments, "has_field_names", true);

            int? codePage = null;
            if (arguments.TryGetProperty("code_page", out var codePageElement))
            {
                if (codePageElement.ValueKind == JsonValueKind.Number && codePageElement.TryGetInt32(out var numericCodePage))
                {
                    codePage = numericCodePage;
                }
                else if (codePageElement.ValueKind == JsonValueKind.String && int.TryParse(codePageElement.GetString(), out var parsedCodePage))
                {
                    codePage = parsedCodePage;
                }
                else if (codePageElement.ValueKind is not (JsonValueKind.Null or JsonValueKind.Undefined))
                {
                    return new { success = false, error = "code_page must be an integer when provided" };
                }
            }

            var result = accessService.TransferText(
                transferType,
                tableName,
                fileName,
                string.IsNullOrWhiteSpace(specificationName) ? null : specificationName,
                hasFieldNames,
                string.IsNullOrWhiteSpace(htmlTableName) ? null : htmlTableName,
                codePage);

            return new
            {
                success = true,
                transfer_type = result.TransferType,
                specification_name = result.SpecificationName,
                table_name = result.TableName,
                file_name = result.FileName,
                has_field_names = result.HasFieldNames,
                html_table_name = result.HtmlTableName,
                code_page = result.CodePage
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("transfer_text", ex);
        }
    }

    static object HandleOutputTo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "output_format", out var outputFormat, out var outputFormatError))
                return outputFormatError;

            _ = TryGetOptionalString(arguments, "object_name", out var objectName);
            _ = TryGetOptionalString(arguments, "output_file", out var outputFile);
            _ = TryGetOptionalString(arguments, "template_file", out var templateFile);
            _ = TryGetOptionalString(arguments, "encoding", out var encoding);
            _ = TryGetOptionalString(arguments, "output_quality", out var outputQuality);
            var autoStart = GetOptionalBool(arguments, "auto_start", false);

            var result = accessService.OutputTo(
                objectType,
                string.IsNullOrWhiteSpace(objectName) ? null : objectName,
                outputFormat,
                string.IsNullOrWhiteSpace(outputFile) ? null : outputFile,
                autoStart,
                string.IsNullOrWhiteSpace(templateFile) ? null : templateFile,
                string.IsNullOrWhiteSpace(encoding) ? null : encoding,
                string.IsNullOrWhiteSpace(outputQuality) ? null : outputQuality);

            return new
            {
                success = true,
                object_type = result.ObjectType,
                object_name = result.ObjectName,
                output_format = result.OutputFormat,
                output_file = result.OutputFile,
                auto_start = result.AutoStart,
                template_file = result.TemplateFile,
                encoding = result.Encoding,
                output_quality = result.OutputQuality
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("output_to", ex);
        }
    }

    static object HandleSetWarnings(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var warningsOn = GetOptionalBool(arguments, "warnings_on", true);
            accessService.SetWarnings(warningsOn);
            return new { success = true, warnings_on = warningsOn };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_warnings", ex);
        }
    }

    static object HandleEcho(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var echoOn = GetOptionalBool(arguments, "echo_on", true);
            _ = TryGetOptionalString(arguments, "status_bar_text", out var statusBarText);
            accessService.Echo(echoOn, string.IsNullOrWhiteSpace(statusBarText) ? null : statusBarText);
            return new { success = true, echo_on = echoOn, status_bar_text = string.IsNullOrWhiteSpace(statusBarText) ? null : statusBarText };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("echo", ex);
        }
    }

    static object HandleHourglass(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var hourglassOn = GetOptionalBool(arguments, "hourglass_on", true);
            accessService.Hourglass(hourglassOn);
            return new { success = true, hourglass_on = hourglassOn };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("hourglass", ex);
        }
    }

    static object HandleGoToRecord(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            _ = TryGetOptionalString(arguments, "object_name", out var objectName);
            _ = TryGetOptionalString(arguments, "record", out var record);
            _ = TryGetOptionalString(arguments, "offset", out var offset);

            accessService.GoToRecord(
                string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                string.IsNullOrWhiteSpace(objectName) ? null : objectName,
                string.IsNullOrWhiteSpace(record) ? null : record,
                string.IsNullOrWhiteSpace(offset) ? null : offset);

            return new
            {
                success = true,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                object_name = string.IsNullOrWhiteSpace(objectName) ? null : objectName,
                record = string.IsNullOrWhiteSpace(record) ? null : record,
                offset = string.IsNullOrWhiteSpace(offset) ? null : offset
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("goto_record", ex);
        }
    }

    static object HandleFindRecord(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "find_what", out var findWhat, out var findWhatError))
                return findWhatError;

            _ = TryGetOptionalString(arguments, "match", out var match);
            _ = TryGetOptionalString(arguments, "match_case", out var matchCase);
            _ = TryGetOptionalString(arguments, "search", out var search);
            _ = TryGetOptionalString(arguments, "search_as_formatted", out var searchAsFormatted);
            _ = TryGetOptionalString(arguments, "only_current_field", out var onlyCurrentField);
            _ = TryGetOptionalString(arguments, "find_first", out var findFirst);

            accessService.FindRecord(
                findWhat,
                string.IsNullOrWhiteSpace(match) ? null : match,
                string.IsNullOrWhiteSpace(matchCase) ? null : matchCase,
                string.IsNullOrWhiteSpace(search) ? null : search,
                string.IsNullOrWhiteSpace(searchAsFormatted) ? null : searchAsFormatted,
                string.IsNullOrWhiteSpace(onlyCurrentField) ? null : onlyCurrentField,
                string.IsNullOrWhiteSpace(findFirst) ? null : findFirst);

            return new
            {
                success = true,
                find_what = findWhat
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("find_record", ex);
        }
    }

    static object HandleApplyFilter(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "filter_name", out var filterName);
            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            _ = TryGetOptionalString(arguments, "control_name", out var controlName);

            accessService.ApplyFilter(
                string.IsNullOrWhiteSpace(filterName) ? null : filterName,
                string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition,
                string.IsNullOrWhiteSpace(controlName) ? null : controlName);

            return new
            {
                success = true,
                filter_name = string.IsNullOrWhiteSpace(filterName) ? null : filterName,
                where_condition = string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition,
                control_name = string.IsNullOrWhiteSpace(controlName) ? null : controlName
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("apply_filter", ex);
        }
    }

    static object HandleShowAllRecords(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.ShowAllRecords();
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("show_all_records", ex);
        }
    }

    static object HandleMaximizeWindow(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.MaximizeWindow();
            return new { success = true, state = "maximized" };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("maximize_window", ex);
        }
    }

    static object HandleMinimizeWindow(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.MinimizeWindow();
            return new { success = true, state = "minimized" };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("minimize_window", ex);
        }
    }

    static object HandleRestoreWindow(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.RestoreWindow();
            return new { success = true, state = "restored" };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("restore_window", ex);
        }
    }

    static object HandlePrintOut(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "print_range", out var printRange);
            _ = TryGetOptionalString(arguments, "print_quality", out var printQuality);
            var collateCopies = GetOptionalBool(arguments, "collate_copies", false);

            if (!TryGetOptionalInt(arguments, "page_from", out var pageFrom, out var pageFromError))
                return pageFromError;
            if (!TryGetOptionalInt(arguments, "page_to", out var pageTo, out var pageToError))
                return pageToError;
            if (!TryGetOptionalInt(arguments, "copies", out var copies, out var copiesError))
                return copiesError;

            accessService.PrintOut(
                string.IsNullOrWhiteSpace(printRange) ? null : printRange,
                pageFrom,
                pageTo,
                string.IsNullOrWhiteSpace(printQuality) ? null : printQuality,
                copies,
                collateCopies);

            return new
            {
                success = true,
                print_range = string.IsNullOrWhiteSpace(printRange) ? null : printRange,
                page_from = pageFrom,
                page_to = pageTo,
                print_quality = string.IsNullOrWhiteSpace(printQuality) ? null : printQuality,
                copies,
                collate_copies = collateCopies
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("print_out", ex);
        }
    }

    static object HandleOpenQuery(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "query_name", out var queryName, out var queryNameError))
                return queryNameError;

            _ = TryGetOptionalString(arguments, "view", out var view);
            _ = TryGetOptionalString(arguments, "data_mode", out var dataMode);

            accessService.OpenQuery(
                queryName,
                string.IsNullOrWhiteSpace(view) ? null : view,
                string.IsNullOrWhiteSpace(dataMode) ? null : dataMode);

            return new
            {
                success = true,
                query_name = queryName,
                view = string.IsNullOrWhiteSpace(view) ? null : view,
                data_mode = string.IsNullOrWhiteSpace(dataMode) ? null : dataMode
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("open_query", ex);
        }
    }

    static object HandleRunSqlDocmd(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "sql", out var sql, out var sqlError))
                return sqlError;

            var useTransaction = GetOptionalBool(arguments, "use_transaction", true);
            accessService.RunSqlDoCmd(sql, useTransaction);
            return new { success = true, sql, use_transaction = useTransaction };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("run_sql", ex);
        }
    }

    static object HandleGetDatabaseSummaryProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var properties = accessService.GetDatabaseSummaryProperties();
            return new { success = true, properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_database_summary_properties", ex);
        }
    }

    static object HandleSetDatabaseSummaryProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var hasTitle = TryGetOptionalString(arguments, "title", out var title);
            var hasAuthor = TryGetOptionalString(arguments, "author", out var author);
            var hasSubject = TryGetOptionalString(arguments, "subject", out var subject);
            var hasKeywords = TryGetOptionalString(arguments, "keywords", out var keywords);
            var hasComments = TryGetOptionalString(arguments, "comments", out var comments);

            if (!hasTitle && !hasAuthor && !hasSubject && !hasKeywords && !hasComments)
            {
                return new
                {
                    success = false,
                    error = "At least one of title, author, subject, keywords, or comments is required"
                };
            }

            accessService.SetDatabaseSummaryProperties(
                hasTitle ? title : null,
                hasAuthor ? author : null,
                hasSubject ? subject : null,
                hasKeywords ? keywords : null,
                hasComments ? comments : null);

            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_database_summary_properties", ex);
        }
    }

    static object HandleGetDatabaseProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var includeSystem = GetOptionalBool(arguments, "include_system", false);
            var properties = accessService.GetDatabaseProperties(includeSystem);
            return new { success = true, properties = properties.ToArray(), include_system = includeSystem };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_database_properties", ex);
        }
    }

    static object HandleGetDatabaseProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "property_name", out var propertyName, out var propertyNameError))
                return propertyNameError;

            var property = accessService.GetDatabaseProperty(propertyName);
            return new { success = true, property };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_database_property", ex);
        }
    }

    static object HandleSetDatabaseProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "property_name", out var propertyName, out var propertyNameError))
                return propertyNameError;
            if (!TryGetRequiredString(arguments, "value", out var value, out var valueError))
                return valueError;

            _ = TryGetOptionalString(arguments, "property_type", out var propertyType);
            var createIfMissing = GetOptionalBool(arguments, "create_if_missing", true);

            accessService.SetDatabaseProperty(
                propertyName,
                value,
                string.IsNullOrWhiteSpace(propertyType) ? null : propertyType,
                createIfMissing);

            return new
            {
                success = true,
                property_name = propertyName,
                value,
                property_type = string.IsNullOrWhiteSpace(propertyType) ? null : propertyType,
                create_if_missing = createIfMissing
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_database_property", ex);
        }
    }

    static object HandleGetTableProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var properties = accessService.GetTableProperties(tableName);
            return new { success = true, properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_table_properties", ex);
        }
    }

    static object HandleSetTableProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var hasDescription = TryGetOptionalString(arguments, "description", out var description);
            var hasValidationRule = TryGetOptionalString(arguments, "validation_rule", out var validationRule);
            var hasValidationText = TryGetOptionalString(arguments, "validation_text", out var validationText);

            if (!hasDescription && !hasValidationRule && !hasValidationText)
            {
                return new
                {
                    success = false,
                    error = "At least one of description, validation_rule, or validation_text is required"
                };
            }

            accessService.SetTableProperties(
                tableName,
                hasDescription ? description : null,
                hasValidationRule ? validationRule : null,
                hasValidationText ? validationText : null);

            return new { success = true, table_name = tableName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_table_properties", ex);
        }
    }

    static object HandleGetQueryProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "query_name", out var queryName, out var queryNameError))
                return queryNameError;

            var properties = accessService.GetQueryProperties(queryName);
            return new { success = true, properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_query_properties", ex);
        }
    }

    static object HandleSetQueryProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "query_name", out var queryName, out var queryNameError))
                return queryNameError;

            var hasDescription = TryGetOptionalString(arguments, "description", out var description);
            var hasSql = TryGetOptionalString(arguments, "sql", out var sql);

            if (!hasDescription && !hasSql)
                return new { success = false, error = "At least one of description or sql is required" };

            accessService.SetQueryProperties(
                queryName,
                hasDescription ? description : null,
                hasSql ? sql : null);

            return new { success = true, query_name = queryName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_query_properties", ex);
        }
    }

    static object HandleSetFieldValidation(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "validation_rule", out var validationRule, out var validationRuleError))
                return validationRuleError;

            _ = TryGetOptionalString(arguments, "validation_text", out var validationText);
            accessService.SetFieldValidation(
                tableName,
                fieldName,
                validationRule,
                string.IsNullOrWhiteSpace(validationText) ? null : validationText);

            return new { success = true, table_name = tableName, field_name = fieldName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_field_validation", ex);
        }
    }

    static object HandleSetFieldDefault(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "default_value", out var defaultValue, out var defaultValueError))
                return defaultValueError;

            accessService.SetFieldDefault(tableName, fieldName, defaultValue);
            return new { success = true, table_name = tableName, field_name = fieldName, default_value = defaultValue };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_field_default", ex);
        }
    }

    static object HandleSetFieldInputMask(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "input_mask", out var inputMask, out var inputMaskError))
                return inputMaskError;

            accessService.SetFieldInputMask(tableName, fieldName, inputMask);
            return new { success = true, table_name = tableName, field_name = fieldName, input_mask = inputMask };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_field_input_mask", ex);
        }
    }

    static object HandleSetFieldCaption(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "caption", out var caption, out var captionError))
                return captionError;

            accessService.SetFieldCaption(tableName, fieldName, caption);
            return new { success = true, table_name = tableName, field_name = fieldName, caption };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_field_caption", ex);
        }
    }

    static object HandleGetFieldProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;

            var properties = accessService.GetFieldProperties(tableName, fieldName);
            return new { success = true, properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_field_properties", ex);
        }
    }

    static object HandleSetLookupProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;

            var hasRowSource = TryGetOptionalString(arguments, "row_source", out var rowSource);
            var hasColumnWidths = TryGetOptionalString(arguments, "column_widths", out var columnWidths);
            if (!TryGetOptionalInt(arguments, "bound_column", out var boundColumn, out var boundColumnError))
                return boundColumnError;
            if (!TryGetOptionalInt(arguments, "column_count", out var columnCount, out var columnCountError))
                return columnCountError;
            if (!TryGetOptionalInt(arguments, "display_control", out var displayControl, out var displayControlError))
                return displayControlError;
            if (!TryGetOptionalBoolNullable(arguments, "limit_to_list", out var limitToList, out var limitToListError))
                return limitToListError;
            if (!TryGetOptionalBoolNullable(arguments, "allow_multiple_values", out var allowMultipleValues, out var allowMultipleValuesError))
                return allowMultipleValuesError;

            if (!hasRowSource &&
                !hasColumnWidths &&
                !boundColumn.HasValue &&
                !columnCount.HasValue &&
                !displayControl.HasValue &&
                !limitToList.HasValue &&
                !allowMultipleValues.HasValue)
            {
                return new
                {
                    success = false,
                    error = "At least one lookup property is required"
                };
            }

            accessService.SetLookupProperties(
                tableName,
                fieldName,
                hasRowSource ? rowSource : null,
                boundColumn,
                columnCount,
                hasColumnWidths ? columnWidths : null,
                limitToList,
                allowMultipleValues,
                displayControl);

            return new { success = true, table_name = tableName, field_name = fieldName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_lookup_properties", ex);
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

            SendLogNotification("debug", "sql", new { sql, maxRows });
            var result = accessService.ExecuteSql(sql, maxRows);

            if (result.IsQuery)
            {
                SendLogNotification("info", "sql", new { sql, rowCount = result.RowCount, truncated = result.Truncated });
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

            SendLogNotification("info", "sql", new { sql, rowsAffected = result.RowsAffected });
            return new
            {
                success = true,
                is_query = false,
                rows_affected = result.RowsAffected
            };
        }
        catch (Exception ex)
        {
            SendLogNotification("error", "sql", new { error = ex.Message });
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

    static object HandlePodbcGetSchemas(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "Schema", out var requestedSchema);

            var inferredSchemas = accessService.GetTables()
                .Select(GetTableNameValue)
                .Select(ParsePodbcTableIdentifier)
                .Select(parts => parts.SchemaName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!string.IsNullOrWhiteSpace(requestedSchema) &&
                !inferredSchemas.Contains(requestedSchema, StringComparer.OrdinalIgnoreCase))
            {
                inferredSchemas.Insert(0, requestedSchema.Trim());
            }

            if (inferredSchemas.Count == 0)
                inferredSchemas.Add(PodbcDefaultSchema);

            var schemaRows = inferredSchemas
                .Select(name => new
                {
                    TABLE_CAT = name,
                    TABLE_SCHEM = name,
                    schema_name = name
                })
                .ToArray();

            var metadata = BuildPyodbcCompatMetadata(arguments);

            return new
            {
                success = true,
                schemas = inferredSchemas.ToArray(),
                results = schemaRows,
                metadata
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("podbc_get_schemas", ex);
        }
    }

    static object HandlePodbcGetTables(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "Schema", out var requestedSchema);
            var tableDetails = accessService.GetTables().ToArray();
            var tableRows = BuildPodbcTableRows(tableDetails.Cast<object>(), requestedSchema);
            var metadata = BuildPyodbcCompatMetadata(arguments);

            return new
            {
                success = true,
                tables = tableRows,
                table_names = tableRows.Select(row => row.TABLE_NAME).ToArray(),
                table_details = tableDetails,
                results = tableRows,
                metadata
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("podbc_get_tables", ex);
        }
    }

    static object HandlePodbcDescribeTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "table", "table_name" }, "table", out var tableName, out var tableError))
                return tableError;

            var description = accessService.DescribeTable(tableName);
            _ = TryGetOptionalString(arguments, "Schema", out var requestedSchema);
            var parsedIdentifier = ParsePodbcTableIdentifier(description.TableName);
            var effectiveSchema = !string.IsNullOrWhiteSpace(requestedSchema)
                ? requestedSchema.Trim()
                : string.IsNullOrWhiteSpace(parsedIdentifier.SchemaName)
                    ? PodbcDefaultSchema
                    : parsedIdentifier.SchemaName;
            var effectiveTableName = string.IsNullOrWhiteSpace(parsedIdentifier.TableName)
                ? description.TableName
                : parsedIdentifier.TableName;
            var primaryKeyNames = new HashSet<string>(description.PrimaryKeyColumns, StringComparer.OrdinalIgnoreCase);
            var columns = description.Columns.Select(column => new
            {
                name = column.Name,
                type = column.DataType,
                column_size = column.MaxLength,
                num_prec_radix = column.NumericPrecision,
                nullable = column.IsNullable,
                @default = column.DefaultValue,
                primary_key = primaryKeyNames.Contains(column.Name),
                ordinal_position = column.OrdinalPosition,
                data_type_code = column.DataTypeCode,
                numeric_scale = column.NumericScale
            }).ToArray();
            var table = new
            {
                TABLE_CAT = effectiveSchema,
                TABLE_SCHEM = effectiveSchema,
                TABLE_NAME = effectiveTableName,
                columns,
                primary_keys = description.PrimaryKeyColumns.ToArray()
            };
            var metadata = BuildPyodbcCompatMetadata(arguments);

            return new
            {
                success = true,
                table,
                table_definition = description,
                table_name = effectiveTableName,
                results = table,
                metadata
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("podbc_describe_table", ex);
        }
    }

    static object HandlePodbcFilterTableNames(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "q", out var queryText, out var queryTextError))
                return queryTextError;

            _ = TryGetOptionalString(arguments, "Schema", out var requestedSchema);

            var filteredDetails = accessService.GetTables()
                .Where(table =>
                {
                    var name = GetTableNameValue(table);
                    return name.IndexOf(queryText, StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .ToArray();
            var filteredRows = BuildPodbcTableRows(filteredDetails.Cast<object>(), requestedSchema);

            var metadata = BuildPyodbcCompatMetadata(arguments);

            return new
            {
                success = true,
                q = queryText,
                tables = filteredRows,
                table_names = filteredRows.Select(row => row.TABLE_NAME).ToArray(),
                table_details = filteredDetails,
                results = filteredRows,
                metadata
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("podbc_filter_table_names", ex);
        }
    }

    static object HandlePodbcExecuteQuery(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            return HandlePodbcExecuteQueryCore(accessService, arguments, "podbc_execute_query", 200);
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("podbc_execute_query", ex);
        }
    }

    static object HandlePodbcExecuteQueryMd(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredStringFromAliases(arguments, new[] { "query", "sql" }, "query", out var query, out var queryError))
                return queryError;

            if (!TryValidatePodbcParams(arguments, "podbc_execute_query_md", out var paramsError))
                return paramsError;

            var maxRows = GetOptionalIntFromAliases(arguments, new[] { "max_rows" }, 100);
            if (maxRows <= 0)
                return new { success = false, error = "max_rows must be greater than 0" };

            var markdown = accessService.ExecuteQueryMarkdown(query, maxRows);
            var metadata = BuildPyodbcCompatMetadata(arguments);

            return new
            {
                success = true,
                query,
                markdown,
                max_rows = maxRows,
                results = markdown,
                metadata
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("podbc_execute_query_md", ex);
        }
    }

    static object HandlePodbcQueryDatabase(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            return HandlePodbcExecuteQueryCore(accessService, arguments, "podbc_query_database", 200);
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("podbc_query_database", ex);
        }
    }

    static object HandlePodbcExecuteQueryCore(AccessInteropService accessService, JsonElement arguments, string operationName, int defaultMaxRows)
    {
        if (!TryGetRequiredStringFromAliases(arguments, new[] { "query", "sql" }, "query", out var query, out var queryError))
            return queryError;

        if (!TryValidatePodbcParams(arguments, operationName, out var paramsError))
            return paramsError;

        var maxRows = GetOptionalIntFromAliases(arguments, new[] { "max_rows" }, defaultMaxRows);
        if (maxRows <= 0)
            return new { success = false, error = "max_rows must be greater than 0" };

        var result = accessService.ExecuteSql(query, maxRows);
        var metadata = BuildPyodbcCompatMetadata(arguments);

        if (result.IsQuery)
        {
            return new
            {
                success = true,
                query,
                is_query = true,
                columns = result.Columns,
                rows = result.Rows,
                row_count = result.RowCount,
                truncated = result.Truncated,
                max_rows = maxRows,
                results = result.Rows,
                metadata
            };
        }

        return new
        {
            success = true,
            query,
            is_query = false,
            rows_affected = result.RowsAffected,
            results = Array.Empty<object>(),
            metadata
        };
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

    static bool TryValidatePodbcParams(JsonElement arguments, string toolName, out object error)
    {
        error = new { success = true };

        if (!arguments.TryGetProperty("params", out var paramsElement))
            return true;

        var hasUnsupportedParams = paramsElement.ValueKind switch
        {
            JsonValueKind.Null => false,
            JsonValueKind.Undefined => false,
            JsonValueKind.String => !string.IsNullOrWhiteSpace(paramsElement.GetString()),
            JsonValueKind.Array => paramsElement.EnumerateArray().Any(),
            JsonValueKind.Object => paramsElement.EnumerateObject().Any(),
            _ => true
        };

        if (!hasUnsupportedParams)
            return true;

        error = new
        {
            success = false,
            error = $"{toolName} does not support non-empty params in Access fallback mode",
            metadata = BuildPyodbcCompatMetadata(arguments),
            unsupported_params = true
        };
        return false;
    }

    static string GetTableNameValue(object tableEntry)
    {
        if (tableEntry is null)
            return string.Empty;

        if (tableEntry is string tableName)
            return tableName;

        var runtimeType = tableEntry.GetType();
        var nameProperty = runtimeType.GetProperty("Name") ?? runtimeType.GetProperty("name");
        if (nameProperty?.GetValue(tableEntry) is string namedTable && !string.IsNullOrWhiteSpace(namedTable))
            return namedTable;

        return tableEntry.ToString() ?? string.Empty;
    }

    static (string? SchemaName, string TableName) ParsePodbcTableIdentifier(string tableIdentifier)
    {
        if (string.IsNullOrWhiteSpace(tableIdentifier))
            return (null, string.Empty);

        var trimmed = tableIdentifier.Trim();
        var separatorIndex = trimmed.IndexOf('.');
        if (separatorIndex <= 0 || separatorIndex == trimmed.Length - 1)
            return (null, trimmed);

        var schemaName = trimmed[..separatorIndex].Trim();
        var tableName = trimmed[(separatorIndex + 1)..].Trim();
        if (string.IsNullOrWhiteSpace(schemaName) || string.IsNullOrWhiteSpace(tableName))
            return (null, trimmed);

        return (schemaName, tableName);
    }

    static PodbcTableRow[] BuildPodbcTableRows(IEnumerable<object> tableEntries, string? requestedSchema)
    {
        var schemaOverride = string.IsNullOrWhiteSpace(requestedSchema) ? null : requestedSchema.Trim();

        return tableEntries
            .Select(GetTableNameValue)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(name =>
            {
                var parsed = ParsePodbcTableIdentifier(name);
                var tableSchema = !string.IsNullOrWhiteSpace(schemaOverride)
                    ? schemaOverride
                    : string.IsNullOrWhiteSpace(parsed.SchemaName)
                        ? PodbcDefaultSchema
                        : parsed.SchemaName;
                var tableName = string.IsNullOrWhiteSpace(parsed.TableName) ? name : parsed.TableName;

                return new PodbcTableRow
                {
                    TABLE_CAT = tableSchema ?? PodbcDefaultSchema,
                    TABLE_SCHEM = tableSchema ?? PodbcDefaultSchema,
                    TABLE_NAME = tableName
                };
            })
            .ToArray();
    }

    static object BuildPyodbcCompatMetadata(JsonElement arguments)
    {
        _ = TryGetOptionalString(arguments, "Schema", out var schema);
        _ = TryGetOptionalString(arguments, "user", out var user);
        _ = TryGetOptionalString(arguments, "password", out var password);
        _ = TryGetOptionalString(arguments, "dsn", out var dsn);

        return new
        {
            schema = string.IsNullOrWhiteSpace(schema) ? null : schema,
            user = string.IsNullOrWhiteSpace(user) ? null : user,
            password_provided = !string.IsNullOrWhiteSpace(password),
            dsn = string.IsNullOrWhiteSpace(dsn) ? null : dsn,
            ignored_arguments = new[] { "Schema", "user", "password", "dsn" }
        };
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

    static bool TryGetOptionalInt(JsonElement arguments, string propertyName, out int? value, out object error)
    {
        value = null;

        if (!arguments.TryGetProperty(propertyName, out var element))
        {
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined)
        {
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out var numericValue))
        {
            value = numericValue;
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.String && int.TryParse(element.GetString(), out var parsedValue))
        {
            value = parsedValue;
            error = new { success = true };
            return true;
        }

        error = new { success = false, error = $"{propertyName} must be an integer when provided" };
        return false;
    }

    static bool TryGetOptionalBoolNullable(JsonElement arguments, string propertyName, out bool? value, out object error)
    {
        value = null;

        if (!arguments.TryGetProperty(propertyName, out var element))
        {
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined)
        {
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.True)
        {
            value = true;
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.False)
        {
            value = false;
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out var numericValue))
        {
            value = numericValue != 0;
            error = new { success = true };
            return true;
        }

        if (element.ValueKind == JsonValueKind.String)
        {
            var text = element.GetString();
            if (bool.TryParse(text, out var parsedBool))
            {
                value = parsedBool;
                error = new { success = true };
                return true;
            }

            if (int.TryParse(text, out var parsedInt))
            {
                value = parsedInt != 0;
                error = new { success = true };
                return true;
            }
        }

        error = new { success = false, error = $"{propertyName} must be a boolean when provided" };
        return false;
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

        var isCom = ex is COMException || ex.InnerException is COMException;
        SendLogNotification(
            "error",
            isCom ? "com" : "operation",
            new { operation = operationName, error = ex.Message, exceptionType = ex.GetType().Name });

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

    sealed class PodbcTableRow
    {
        public string TABLE_CAT { get; init; } = string.Empty;
        public string TABLE_SCHEM { get; init; } = string.Empty;
        public string TABLE_NAME { get; init; } = string.Empty;
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

    // ── JSON-RPC error sentinel (used by dispatch to signal error responses) ──

    sealed class JsonRpcErrorSentinel
    {
        public int Code { get; init; }
        public string Message { get; init; } = "";
    }

    // ── Notification helpers ──

    static void SendNotification(string method, object? @params = null)
    {
        var notification = @params != null
            ? new { jsonrpc = "2.0", method, @params }
            : (object)new { jsonrpc = "2.0", method };
        Console.WriteLine(JsonSerializer.Serialize(notification));
    }

    static void SendLogNotification(string level, string logger, object data)
    {
        if (LogLevelSeverity(level) < LogLevelSeverity(_minimumLogLevel))
            return;
        SendNotification("notifications/message", new { level, logger, data });
    }

    // ── logging/setLevel handler ──

    static object HandleLoggingSetLevel(JsonElement @params)
    {
        if (@params.TryGetProperty("level", out var levelElement) &&
            levelElement.ValueKind == JsonValueKind.String)
        {
            var level = levelElement.GetString()?.ToLowerInvariant() ?? "debug";
            if (Array.IndexOf(LogLevelOrder, level) >= 0)
                _minimumLogLevel = level;
        }
        return new { };
    }

    // ── resources/list handler ──

    static object HandleResourcesList(AccessInteropService accessService)
    {
        var resources = new List<object>
        {
            new { uri = "access://connection", name = "Connection Status", description = "Current database connection status and path", mimeType = "application/json" },
            new { uri = "access://tables", name = "Tables", description = "All user tables with fields and record counts", mimeType = "application/json" },
            new { uri = "access://queries", name = "Queries", description = "All saved queries with SQL definitions", mimeType = "application/json" },
            new { uri = "access://relationships", name = "Relationships", description = "Table relationships and referential integrity rules", mimeType = "application/json" },
            new { uri = "access://forms", name = "Forms", description = "All form objects in the database", mimeType = "application/json" },
            new { uri = "access://reports", name = "Reports", description = "All report objects in the database", mimeType = "application/json" },
            new { uri = "access://macros", name = "Macros", description = "All macro objects in the database", mimeType = "application/json" },
            new { uri = "access://modules", name = "Modules", description = "All VBA module objects in the database", mimeType = "application/json" },
            new { uri = "access://linked-tables", name = "Linked Tables", description = "All linked/attached table definitions", mimeType = "application/json" },
            new { uri = "access://metadata", name = "Object Metadata", description = "MSysObjects metadata for all database objects", mimeType = "application/json" }
        };

        return new { resources };
    }

    // ── resources/templates/list handler ──

    static object HandleResourceTemplatesList()
    {
        var resourceTemplates = new List<object>
        {
            new { uriTemplate = "access://schema/{tableName}", name = "Table Schema", description = "Column definitions, types, and constraints for a specific table", mimeType = "application/json" },
            new { uriTemplate = "access://indexes/{tableName}", name = "Table Indexes", description = "Index definitions for a specific table", mimeType = "application/json" },
            new { uriTemplate = "access://query/{queryName}", name = "Query Definition", description = "SQL definition of a specific saved query", mimeType = "application/json" },
            new { uriTemplate = "access://vba/{moduleName}", name = "VBA Module Code", description = "Source code of a specific VBA module", mimeType = "text/plain" },
            new { uriTemplate = "access://form/{formName}", name = "Form Definition", description = "Exported text representation of a specific form", mimeType = "application/json" },
            new { uriTemplate = "access://report/{reportName}", name = "Report Definition", description = "Exported text representation of a specific report", mimeType = "application/json" }
        };

        return new { resourceTemplates };
    }

    // ── resources/read handler ──

    static object HandleResourcesRead(AccessInteropService accessService, JsonElement @params)
    {
        if (!@params.TryGetProperty("uri", out var uriElement) || uriElement.ValueKind != JsonValueKind.String)
            return new JsonRpcErrorSentinel { Code = -32602, Message = "Missing required parameter: uri" };

        var uri = uriElement.GetString() ?? "";
        if (!uri.StartsWith("access://"))
            return new JsonRpcErrorSentinel { Code = -32602, Message = $"Invalid resource URI scheme: {uri}" };

        var path = uri.Substring("access://".Length);

        try
        {
            // Static resources
            object? data = path switch
            {
                "connection" => new { connected = accessService.IsConnected, databasePath = accessService.CurrentDatabasePath },
                "tables" => accessService.IsConnected ? accessService.GetTables() : (object)new List<object>(),
                "queries" => accessService.IsConnected ? accessService.GetQueries() : (object)new List<object>(),
                "relationships" => accessService.IsConnected ? accessService.GetRelationships() : (object)new List<object>(),
                "forms" => accessService.IsConnected ? accessService.GetForms() : (object)new List<object>(),
                "reports" => accessService.IsConnected ? accessService.GetReports() : (object)new List<object>(),
                "macros" => accessService.IsConnected ? accessService.GetMacros() : (object)new List<object>(),
                "modules" => accessService.IsConnected ? accessService.GetModules() : (object)new List<object>(),
                "linked-tables" => accessService.IsConnected ? accessService.GetLinkedTables() : (object)new List<object>(),
                "metadata" => accessService.IsConnected ? accessService.GetObjectMetadata() : (object)new List<object>(),
                _ => null
            };

            // Template resources (parameterized)
            if (data == null)
            {
                if (!accessService.IsConnected)
                    return new JsonRpcErrorSentinel { Code = -32002, Message = "Not connected to a database" };

                if (path.StartsWith("schema/"))
                {
                    var tableName = Uri.UnescapeDataString(path.Substring("schema/".Length));
                    data = accessService.DescribeTable(tableName);
                }
                else if (path.StartsWith("indexes/"))
                {
                    var tableName = Uri.UnescapeDataString(path.Substring("indexes/".Length));
                    data = accessService.GetIndexes(tableName);
                }
                else if (path.StartsWith("query/"))
                {
                    var queryName = Uri.UnescapeDataString(path.Substring("query/".Length));
                    var allQueries = accessService.GetQueries();
                    data = allQueries.FirstOrDefault(q => string.Equals(q.Name, queryName, StringComparison.OrdinalIgnoreCase));
                    if (data == null)
                        return new JsonRpcErrorSentinel { Code = -32002, Message = $"Query not found: {queryName}" };
                }
                else if (path.StartsWith("vba/"))
                {
                    var moduleName = Uri.UnescapeDataString(path.Substring("vba/".Length));
                    var projects = accessService.GetVBAProjects();
                    var projectName = projects.FirstOrDefault()?.Name ?? "";
                    if (string.IsNullOrEmpty(projectName))
                        return new JsonRpcErrorSentinel { Code = -32002, Message = "No VBA projects found in the database" };
                    var code = accessService.GetVBACode(projectName, moduleName);
                    return new
                    {
                        contents = new object[]
                        {
                            new { uri, mimeType = "text/plain", text = code }
                        }
                    };
                }
                else if (path.StartsWith("form/"))
                {
                    var formName = Uri.UnescapeDataString(path.Substring("form/".Length));
                    var formData = accessService.ExportFormToText(formName);
                    data = new { name = formName, definition = formData };
                }
                else if (path.StartsWith("report/"))
                {
                    var reportName = Uri.UnescapeDataString(path.Substring("report/".Length));
                    var reportData = accessService.ExportReportToText(reportName);
                    data = new { name = reportName, definition = reportData };
                }
                else
                {
                    return new JsonRpcErrorSentinel { Code = -32002, Message = $"Resource not found: {uri}" };
                }
            }

            var jsonText = JsonSerializer.Serialize(data);
            return new
            {
                contents = new object[]
                {
                    new { uri, mimeType = "application/json", text = jsonText }
                }
            };
        }
        catch (Exception ex)
        {
            SendLogNotification("error", "resources", new { uri, error = ex.Message });
            return new JsonRpcErrorSentinel { Code = -32002, Message = ex.Message };
        }
    }

    // ── prompts/list handler ──

    static object HandlePromptsList()
    {
        var prompts = new object[]
        {
            new
            {
                name = "analyze_schema",
                description = "Analyze the database schema for normalization issues, missing indexes, and design improvements",
                arguments = Array.Empty<object>()
            },
            new
            {
                name = "debug_query",
                description = "Analyze a SQL query for Jet/ACE compatibility issues, performance problems, and errors",
                arguments = new object[]
                {
                    new { name = "sql", description = "The SQL query to analyze", required = true }
                }
            },
            new
            {
                name = "create_normalized_schema",
                description = "Design a normalized Access database schema from a natural-language description",
                arguments = new object[]
                {
                    new { name = "description", description = "Natural-language description of the data to be stored", required = true }
                }
            },
            new
            {
                name = "migrate_data",
                description = "Plan a data migration from a source table to a new structure",
                arguments = new object[]
                {
                    new { name = "source_table", description = "Name of the source table", required = true },
                    new { name = "target_description", description = "Description of the desired target structure", required = true }
                }
            },
            new
            {
                name = "document_database",
                description = "Generate comprehensive documentation of all database objects",
                arguments = Array.Empty<object>()
            },
            new
            {
                name = "vba_review",
                description = "Review a VBA module for errors, best practices, and improvements",
                arguments = new object[]
                {
                    new { name = "module_name", description = "Name of the VBA module to review", required = true }
                }
            }
        };

        return new { prompts };
    }

    // ── prompts/get handler ──

    static object HandlePromptsGet(AccessInteropService accessService, JsonElement @params)
    {
        if (!@params.TryGetProperty("name", out var nameElement) || nameElement.ValueKind != JsonValueKind.String)
            return new JsonRpcErrorSentinel { Code = -32602, Message = "Missing required parameter: name" };

        var promptName = nameElement.GetString() ?? "";

        // Extract arguments if provided
        var arguments = EmptyJsonObject;
        if (@params.TryGetProperty("arguments", out var argsElement) && argsElement.ValueKind == JsonValueKind.Object)
            arguments = argsElement;

        return promptName switch
        {
            "analyze_schema" => BuildAnalyzeSchemaPrompt(accessService),
            "debug_query" => BuildDebugQueryPrompt(arguments),
            "create_normalized_schema" => BuildCreateNormalizedSchemaPrompt(arguments),
            "migrate_data" => BuildMigrateDataPrompt(arguments),
            "document_database" => BuildDocumentDatabasePrompt(accessService),
            "vba_review" => BuildVbaReviewPrompt(accessService, arguments),
            _ => new JsonRpcErrorSentinel { Code = -32602, Message = $"Unknown prompt: {promptName}" }
        };
    }

    static object BuildAnalyzeSchemaPrompt(AccessInteropService accessService)
    {
        var messages = new List<object>
        {
            new
            {
                role = "user",
                content = new object[]
                {
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://tables", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetTables()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://relationships", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetRelationships()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://queries", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetQueries()) : "[]" }
                    },
                    new
                    {
                        type = "text",
                        text = "Analyze this Microsoft Access database schema. Examine the tables, fields, relationships, and queries. Report on:\n\n1. **Normalization issues** — identify any tables that violate 1NF, 2NF, or 3NF and suggest fixes\n2. **Missing relationships** — fields that look like foreign keys but have no defined relationship\n3. **Missing indexes** — fields used in relationships or likely query filters that lack indexes\n4. **Data type concerns** — fields using Text where Number/Date would be better, oversized field lengths, etc.\n5. **Naming conventions** — inconsistent naming patterns across tables and fields\n6. **Query optimization** — any saved queries that could be improved\n\nProvide specific, actionable recommendations for each issue found."
                    }
                }
            }
        };

        return new { messages };
    }

    static object BuildDebugQueryPrompt(JsonElement arguments)
    {
        var sql = "";
        if (arguments.TryGetProperty("sql", out var sqlElement) && sqlElement.ValueKind == JsonValueKind.String)
            sql = sqlElement.GetString() ?? "";

        if (string.IsNullOrWhiteSpace(sql))
            return new JsonRpcErrorSentinel { Code = -32602, Message = "Missing required argument: sql" };

        var messages = new List<object>
        {
            new
            {
                role = "user",
                content = new object[]
                {
                    new
                    {
                        type = "text",
                        text = $"Debug this SQL query for Microsoft Access (Jet/ACE SQL dialect):\n\n```sql\n{sql}\n```\n\nAnalyze for:\n1. **Syntax errors** — Jet/ACE SQL differs from standard SQL (no LIMIT, use TOP; no BOOLEAN, use Yes/No; date literals use #...#)\n2. **Compatibility issues** — features not supported by Access (CTEs, window functions, FULL OUTER JOIN, etc.)\n3. **Performance concerns** — missing indexes, cartesian products, unnecessary subqueries\n4. **Correctness** — logic errors, NULL handling issues, ambiguous column references\n5. **Suggested fix** — provide a corrected version of the query if issues are found"
                    }
                }
            }
        };

        return new { messages };
    }

    static object BuildCreateNormalizedSchemaPrompt(JsonElement arguments)
    {
        var description = "";
        if (arguments.TryGetProperty("description", out var descElement) && descElement.ValueKind == JsonValueKind.String)
            description = descElement.GetString() ?? "";

        if (string.IsNullOrWhiteSpace(description))
            return new JsonRpcErrorSentinel { Code = -32602, Message = "Missing required argument: description" };

        var messages = new List<object>
        {
            new
            {
                role = "user",
                content = new object[]
                {
                    new
                    {
                        type = "text",
                        text = $"Design a normalized Microsoft Access database schema for the following requirements:\n\n{description}\n\nProvide:\n1. **Table definitions** — table names, field names, data types (use Access/Jet types: Text, Long Integer, Currency, Date/Time, Yes/No, Memo, AutoNumber, etc.)\n2. **Primary keys** — every table should have a primary key (prefer AutoNumber for surrogate keys)\n3. **Relationships** — define all foreign keys with referential integrity settings (cascade update/delete where appropriate)\n4. **Indexes** — suggest indexes for frequently queried fields and foreign keys\n5. **Sample CREATE TABLE SQL** — in Jet SQL dialect\n6. **Normalization justification** — explain how the design satisfies 3NF\n\nEnsure the design follows Access best practices (e.g., avoid reserved words for field names, use appropriate field sizes)."
                    }
                }
            }
        };

        return new { messages };
    }

    static object BuildMigrateDataPrompt(JsonElement arguments)
    {
        var sourceTable = "";
        var targetDescription = "";
        if (arguments.TryGetProperty("source_table", out var srcElement) && srcElement.ValueKind == JsonValueKind.String)
            sourceTable = srcElement.GetString() ?? "";
        if (arguments.TryGetProperty("target_description", out var tgtElement) && tgtElement.ValueKind == JsonValueKind.String)
            targetDescription = tgtElement.GetString() ?? "";

        if (string.IsNullOrWhiteSpace(sourceTable))
            return new JsonRpcErrorSentinel { Code = -32602, Message = "Missing required argument: source_table" };
        if (string.IsNullOrWhiteSpace(targetDescription))
            return new JsonRpcErrorSentinel { Code = -32602, Message = "Missing required argument: target_description" };

        var messages = new List<object>
        {
            new
            {
                role = "user",
                content = new object[]
                {
                    new
                    {
                        type = "text",
                        text = $"Plan a data migration in Microsoft Access.\n\n**Source table:** {sourceTable}\n**Target structure:** {targetDescription}\n\nProvide:\n1. **Target table DDL** — CREATE TABLE statements in Jet SQL\n2. **Migration SQL** — INSERT INTO ... SELECT queries to transform and move data\n3. **Data transformation rules** — how each source field maps to target fields\n4. **Validation queries** — SQL to verify row counts and data integrity after migration\n5. **Rollback plan** — steps to undo the migration if issues are found\n6. **Risks and considerations** — data truncation, NULL handling, type conversion issues\n\nUse the `describe_table` and `execute_sql` tools to examine the source table structure and data before generating the plan."
                    }
                }
            }
        };

        return new { messages };
    }

    static object BuildDocumentDatabasePrompt(AccessInteropService accessService)
    {
        var messages = new List<object>
        {
            new
            {
                role = "user",
                content = new object[]
                {
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://tables", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetTables()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://queries", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetQueries()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://relationships", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetRelationships()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://forms", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetForms()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://reports", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetReports()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://macros", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetMacros()) : "[]" }
                    },
                    new
                    {
                        type = "resource",
                        resource = new { uri = "access://modules", mimeType = "application/json", text = accessService.IsConnected ? JsonSerializer.Serialize(accessService.GetModules()) : "[]" }
                    },
                    new
                    {
                        type = "text",
                        text = "Generate comprehensive documentation for this Microsoft Access database. Include:\n\n1. **Overview** — database purpose (infer from object names), file path, object counts\n2. **Table documentation** — for each table: purpose, field descriptions, data types, primary key, record count\n3. **Relationship diagram** — text-based ERD showing all relationships and cardinality\n4. **Query documentation** — for each query: purpose, SQL, input parameters if any\n5. **Form inventory** — list all forms with likely purpose\n6. **Report inventory** — list all reports with likely purpose\n7. **Macro inventory** — list all macros with likely purpose\n8. **VBA modules** — list all modules\n9. **Recommendations** — any issues or improvements noted during documentation\n\nFormat the output as a well-structured Markdown document."
                    }
                }
            }
        };

        return new { messages };
    }

    static object BuildVbaReviewPrompt(AccessInteropService accessService, JsonElement arguments)
    {
        var moduleName = "";
        if (arguments.TryGetProperty("module_name", out var modElement) && modElement.ValueKind == JsonValueKind.String)
            moduleName = modElement.GetString() ?? "";

        if (string.IsNullOrWhiteSpace(moduleName))
            return new JsonRpcErrorSentinel { Code = -32602, Message = "Missing required argument: module_name" };

        // Try to fetch actual VBA code if connected
        string codeContent = "";
        if (accessService.IsConnected)
        {
            try
            {
                var projects = accessService.GetVBAProjects();
                var projectName = projects.FirstOrDefault()?.Name ?? "";
                if (!string.IsNullOrEmpty(projectName))
                    codeContent = accessService.GetVBACode(projectName, moduleName);
            }
            catch
            {
                // Will be empty, prompt will instruct to fetch via tool
            }
        }

        var contentParts = new List<object>();
        if (!string.IsNullOrEmpty(codeContent))
        {
            contentParts.Add(new
            {
                type = "resource",
                resource = new { uri = $"access://vba/{Uri.EscapeDataString(moduleName)}", mimeType = "text/plain", text = codeContent }
            });
        }
        contentParts.Add(new
        {
            type = "text",
            text = string.IsNullOrEmpty(codeContent)
                ? $"Review the VBA module named \"{moduleName}\" in this Access database. Use the `get_vba_code` tool to retrieve the source code, then analyze it."
                : $"Review this VBA module \"{moduleName}\" from the Access database."
                + "\n\nAnalyze for:\n1. **Errors and bugs** — unhandled errors, incorrect logic, missing error handlers\n2. **Error handling** — ensure all procedures have proper On Error statements\n3. **Best practices** — Option Explicit, meaningful variable names, proper scoping (Dim vs Public)\n4. **Performance** — unnecessary loops, repeated COM calls that could be cached, missing DoEvents in long operations\n5. **Security** — SQL injection in string-concatenated queries (use parameterized queries instead)\n6. **Maintainability** — dead code, overly complex procedures that should be split, missing comments on non-obvious logic\n7. **Access-specific issues** — proper use of CurrentDb vs CodeDb, DAO vs ADO consistency, proper form/report references\n\nProvide specific line-level feedback and corrected code snippets."
        });

        var messages = new List<object>
        {
            new { role = "user", content = contentParts }
        };

        return new { messages };
    }

    // ── completion/complete handler ──

    static object HandleCompletionComplete(AccessInteropService accessService, JsonElement @params)
    {
        if (!@params.TryGetProperty("ref", out var refElement) || refElement.ValueKind != JsonValueKind.Object)
            return new { completion = new { values = Array.Empty<string>(), total = 0, hasMore = false } };

        var refType = "";
        if (refElement.TryGetProperty("type", out var typeEl) && typeEl.ValueKind == JsonValueKind.String)
            refType = typeEl.GetString() ?? "";

        var argumentName = "";
        if (refElement.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String)
            argumentName = nameEl.GetString() ?? "";

        // Get the partial input to filter on
        var partial = "";
        if (@params.TryGetProperty("argument", out var argElement) && argElement.TryGetProperty("value", out var valEl) && valEl.ValueKind == JsonValueKind.String)
            partial = valEl.GetString() ?? "";

        if (!accessService.IsConnected)
            return new { completion = new { values = Array.Empty<string>(), total = 0, hasMore = false } };

        try
        {
            IEnumerable<string> candidates = Array.Empty<string>();

            if (refType == "ref/resource")
            {
                // Resource template URI — determine parameter from the URI pattern
                var uri = argumentName; // For resource refs, name holds the URI template
                if (uri.Contains("{tableName}"))
                    candidates = accessService.GetTables().Select(t => t.Name);
                else if (uri.Contains("{queryName}"))
                    candidates = accessService.GetQueries().Select(q => q.Name);
                else if (uri.Contains("{formName}"))
                    candidates = accessService.GetForms().Select(f => f.Name);
                else if (uri.Contains("{reportName}"))
                    candidates = accessService.GetReports().Select(r => r.Name);
                else if (uri.Contains("{moduleName}"))
                    candidates = accessService.GetModules().Select(m => m.Name);
            }
            else if (refType == "ref/prompt")
            {
                // Prompt argument completion
                if (argumentName == "module_name")
                    candidates = accessService.GetModules().Select(m => m.Name);
                else if (argumentName == "source_table")
                    candidates = accessService.GetTables().Select(t => t.Name);
            }

            var filtered = string.IsNullOrEmpty(partial)
                ? candidates.Take(100).ToArray()
                : candidates.Where(c => c.StartsWith(partial, StringComparison.OrdinalIgnoreCase)).Take(100).ToArray();

            return new { completion = new { values = filtered, total = filtered.Length, hasMore = false } };
        }
        catch (Exception ex)
        {
            SendLogNotification("warning", "completion", new { error = ex.Message });
            return new { completion = new { values = Array.Empty<string>(), total = 0, hasMore = false } };
        }
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

public class JsonRpcErrorResponse
{
    [JsonPropertyName("jsonrpc")]
    public string Jsonrpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public JsonElement? Id { get; set; }

    [JsonPropertyName("error")]
    public JsonRpcError Error { get; set; } = new();
}

public class JsonRpcError
{
    [JsonPropertyName("code")]
    public int Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
}
