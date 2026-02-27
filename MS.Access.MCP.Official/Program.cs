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
                new { name = "open_table", description = "Open a table in Access using DoCmd.OpenTable.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, view = new { type = "string", description = "datasheet, design, print_preview, pivot_table, pivot_chart, or Access enum integer value as string" }, data_mode = new { type = "string", description = "add, edit, read_only, or Access enum integer value as string" } }, required = new string[] { "table_name" } } },
                new { name = "open_module", description = "Open a VBA module using DoCmd.OpenModule.", inputSchema = new { type = "object", properties = new { module_name = new { type = "string" }, procedure_name = new { type = "string" } }, required = new string[] { "module_name" } } },
                new { name = "copy_object", description = "Copy an Access object using DoCmd.CopyObject.", inputSchema = new { type = "object", properties = new { source_object_name = new { type = "string" }, source_object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, destination_database_path = new { type = "string" }, new_name = new { type = "string" } }, required = new string[] { "source_object_name" } } },
                new { name = "delete_object", description = "Delete an Access object using DoCmd.DeleteObject.", inputSchema = new { type = "object", properties = new { object_name = new { type = "string" }, object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" } }, required = new string[] { "object_name" } } },
                new { name = "rename_object", description = "Rename an Access object using DoCmd.Rename.", inputSchema = new { type = "object", properties = new { new_name = new { type = "string" }, object_name = new { type = "string" }, object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" } }, required = new string[] { "new_name", "object_name" } } },
                new { name = "select_object", description = "Select an Access object using DoCmd.SelectObject.", inputSchema = new { type = "object", properties = new { object_name = new { type = "string" }, object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, in_database_window = new { type = "boolean" } }, required = new string[] { "object_name" } } },
                new { name = "save_object", description = "Save an Access object using DoCmd.Save.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, object_name = new { type = "string" } } } },
                new { name = "close_object", description = "Close an Access object using DoCmd.Close.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, object_name = new { type = "string" }, save = new { type = "string", description = "prompt, yes, no, or Access enum integer value as string" } } } },
                new { name = "transfer_database", description = "Transfer database objects using DoCmd.TransferDatabase.", inputSchema = new { type = "object", properties = new { transfer_type = new { type = "string", description = "import, export, link, or Access enum integer value as string" }, database_type = new { type = "string", description = "Database type such as Microsoft Access" }, database_name = new { type = "string", description = "External database path or DSN" }, object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, source = new { type = "string", description = "Source object name" }, destination = new { type = "string", description = "Destination object name when importing/exporting" }, structure_only = new { type = "boolean" }, store_login = new { type = "boolean" } }, required = new string[] { "transfer_type", "database_type", "database_name", "object_type", "source" } } },
                new { name = "run_command", description = "Run an Access command using DoCmd.RunCommand.", inputSchema = new { type = "object", properties = new { command = new { type = "string", description = "acCommand integer value as string (or supported acCmd constant name)" } }, required = new string[] { "command" } } },
                new { name = "sys_cmd", description = "Execute Access SysCmd for status/progress operations.", inputSchema = new { type = "object", properties = new { command = new { type = "string" }, arg1 = new { }, arg2 = new { }, arg3 = new { } }, required = new string[] { "command" } } },
                new { name = "goto_page", description = "Navigate to a form page using DoCmd.GoToPage.", inputSchema = new { type = "object", properties = new { page_number = new { type = "string" }, right = new { type = "string" }, down = new { type = "string" } }, required = new string[] { "page_number" } } },
                new { name = "goto_control", description = "Move focus to a control using DoCmd.GoToControl.", inputSchema = new { type = "object", properties = new { control_name = new { type = "string" } }, required = new string[] { "control_name" } } },
                new { name = "move_size", description = "Move or resize the active window using DoCmd.MoveSize.", inputSchema = new { type = "object", properties = new { right = new { type = "integer" }, down = new { type = "integer" }, width = new { type = "integer" }, height = new { type = "integer" } } } },
                new { name = "requery", description = "Requery data using DoCmd.Requery.", inputSchema = new { type = "object", properties = new { control_name = new { type = "string" } } } },
                new { name = "repaint_object", description = "Repaint an object using DoCmd.RepaintObject.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, object_name = new { type = "string" } } } },
                new { name = "send_object", description = "Send an Access object as email attachment using DoCmd.SendObject.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, object_name = new { type = "string" }, output_format = new { type = "string" }, to = new { type = "string" }, cc = new { type = "string" }, bcc = new { type = "string" }, subject = new { type = "string" }, message_text = new { type = "string" }, edit_message = new { type = "boolean" }, template_file = new { type = "string" } } } },
                new { name = "browse_to", description = "Browse to an Access object path using DoCmd.BrowseTo.", inputSchema = new { type = "object", properties = new { object_name = new { type = "string" }, object_type = new { type = "string", description = "table, query, form, report, macro, module, or Access enum integer value as string" }, path_to_subform_control = new { type = "string" }, where_condition = new { type = "string" }, page = new { type = "string" } }, required = new string[] { "object_name" } } },
                new { name = "lock_navigation_pane", description = "Lock or unlock the Access navigation pane using DoCmd.LockNavigationPane.", inputSchema = new { type = "object", properties = new { lock_navigation_pane = new { type = "boolean" } } } },
                new { name = "navigate_to", description = "Navigate to an Access navigation category using DoCmd.NavigateTo.", inputSchema = new { type = "object", properties = new { navigation_category = new { type = "string" } }, required = new string[] { "navigation_category" } } },
                new { name = "beep", description = "Play the system beep using DoCmd.Beep.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_database_summary_properties", description = "Get Access database summary properties (Title, Author, Subject, Keywords, Comments).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "set_database_summary_properties", description = "Set Access database summary properties.", inputSchema = new { type = "object", properties = new { title = new { type = "string" }, author = new { type = "string" }, subject = new { type = "string" }, keywords = new { type = "string" }, comments = new { type = "string" } } } },
                new { name = "get_database_properties", description = "List database properties, including custom properties.", inputSchema = new { type = "object", properties = new { include_system = new { type = "boolean" } } } },
                new { name = "get_database_property", description = "Get a single database property by name.", inputSchema = new { type = "object", properties = new { property_name = new { type = "string" } }, required = new string[] { "property_name" } } },
                new { name = "set_database_property", description = "Set or create a database property.", inputSchema = new { type = "object", properties = new { property_name = new { type = "string" }, value = new { type = "string" }, property_type = new { type = "string" }, create_if_missing = new { type = "boolean" } }, required = new string[] { "property_name", "value" } } },
                new { name = "get_table_properties", description = "Get table-level properties such as description and validation settings.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "set_table_properties", description = "Set table-level properties such as description and validation settings.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, description = new { type = "string" }, validation_rule = new { type = "string" }, validation_text = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "get_table_validation", description = "Get table-level ValidationRule and ValidationText.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "get_table_description", description = "Get table Description property.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "set_table_description", description = "Set table Description property.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, description = new { type = "string" } }, required = new string[] { "table_name", "description" } } },
                new { name = "get_all_field_descriptions", description = "Get descriptions for all fields in a table.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "get_query_properties", description = "Get query properties including description, SQL text, and parameters.", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" } }, required = new string[] { "query_name" } } },
                new { name = "get_query_parameters", description = "Get query parameter metadata for a saved query.", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" } }, required = new string[] { "query_name" } } },
                new { name = "set_query_properties", description = "Set query properties such as description and SQL text.", inputSchema = new { type = "object", properties = new { query_name = new { type = "string" }, description = new { type = "string" }, sql = new { type = "string" } }, required = new string[] { "query_name" } } },
                new { name = "set_field_validation", description = "Set field validation rule and validation text.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, validation_rule = new { type = "string" }, validation_text = new { type = "string" } }, required = new string[] { "table_name", "field_name", "validation_rule" } } },
                new { name = "set_field_default", description = "Set field default value.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, default_value = new { type = "string" } }, required = new string[] { "table_name", "field_name", "default_value" } } },
                new { name = "set_field_input_mask", description = "Set field input mask.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, input_mask = new { type = "string" } }, required = new string[] { "table_name", "field_name", "input_mask" } } },
                new { name = "set_field_caption", description = "Set field caption.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, caption = new { type = "string" } }, required = new string[] { "table_name", "field_name", "caption" } } },
                new { name = "get_field_properties", description = "Get field properties including validation/default/input mask/caption and lookup settings.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "get_field_attributes", description = "Get detailed field attributes.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "detect_multi_value_fields", description = "Identify multi-value lookup fields in a table.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "get_multi_value_field_values", description = "Read multi-value field values.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, where_condition = new { type = "string" }, max_rows = new { type = "integer" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "set_multi_value_field_values", description = "Write multi-value field values.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, values = new { type = "array" }, where_condition = new { type = "string" } }, required = new string[] { "table_name", "field_name", "values" } } },
                new { name = "set_lookup_properties", description = "Set lookup properties for a field (RowSource, BoundColumn, ColumnCount, ColumnWidths, etc.).", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, row_source = new { type = "string" }, bound_column = new { type = "integer" }, column_count = new { type = "integer" }, column_widths = new { type = "string" }, limit_to_list = new { type = "boolean" }, allow_multiple_values = new { type = "boolean" }, display_control = new { type = "integer" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "get_vba_references", description = "List VBA references for a VBA project.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" } } } },
                new { name = "add_vba_reference", description = "Add a VBA reference by file path or GUID.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, reference_path = new { type = "string" }, reference_guid = new { type = "string" }, major = new { type = "integer" }, minor = new { type = "integer" } } } },
                new { name = "remove_vba_reference", description = "Remove a VBA reference by name, GUID, or path.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, reference_identifier = new { type = "string" } }, required = new string[] { "reference_identifier" } } },
                new { name = "get_startup_properties", description = "Get application startup properties (StartupForm, AppTitle, AppIcon).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "set_startup_properties", description = "Set application startup properties (StartupForm, AppTitle, AppIcon).", inputSchema = new { type = "object", properties = new { startup_form = new { type = "string" }, app_title = new { type = "string" }, app_icon = new { type = "string" } } } },
                new { name = "get_ribbon_xml", description = "Get ribbon XML by ribbon name or by default database ribbon property.", inputSchema = new { type = "object", properties = new { ribbon_name = new { type = "string" } } } },
                new { name = "set_ribbon_xml", description = "Create or replace ribbon XML in USysRibbons and optionally set as default.", inputSchema = new { type = "object", properties = new { ribbon_name = new { type = "string" }, ribbon_xml = new { type = "string" }, apply_as_default = new { type = "boolean" } }, required = new string[] { "ribbon_name", "ribbon_xml" } } },
                new { name = "get_application_info", description = "Get Access application metadata and current project/data info.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_current_project_data", description = "Get CurrentProject and CurrentData properties.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_application_option", description = "Get an Access application option value using Application.GetOption.", inputSchema = new { type = "object", properties = new { option_name = new { type = "string" } }, required = new string[] { "option_name" } } },
                new { name = "set_application_option", description = "Set an Access application option value using Application.SetOption.", inputSchema = new { type = "object", properties = new { option_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "option_name", "value" } } },
                new { name = "get_temp_vars", description = "List all Access TempVars.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "set_temp_var", description = "Set or create an Access TempVar.", inputSchema = new { type = "object", properties = new { name = new { type = "string" }, value = new { } }, required = new string[] { "name" } } },
                new { name = "remove_temp_var", description = "Remove an Access TempVar.", inputSchema = new { type = "object", properties = new { name = new { type = "string" } }, required = new string[] { "name" } } },
                new { name = "clear_temp_vars", description = "Remove all Access TempVars.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "export_data_macro_axl", description = "Export table data macro definition as AXL.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "import_data_macro_axl", description = "Import table data macro definition from AXL.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, axl_xml = new { type = "string" } }, required = new string[] { "table_name", "axl_xml" } } },
                new { name = "run_data_macro", description = "Run a data macro by name.", inputSchema = new { type = "object", properties = new { macro_name = new { type = "string" } }, required = new string[] { "macro_name" } } },
                new { name = "get_table_data_macros", description = "List data macros defined on a table.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "delete_data_macro", description = "Delete a data macro from a table.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, macro_name = new { type = "string" } }, required = new string[] { "table_name", "macro_name" } } },
                new { name = "get_autoexec_info", description = "Check if the AutoExec macro exists.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "run_autoexec", description = "Run the AutoExec macro.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_database_security", description = "Get current database security state (password/encryption indicators).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "set_database_password", description = "Set or replace the database password (compact/repair based).", inputSchema = new { type = "object", properties = new { new_password = new { type = "string" } }, required = new string[] { "new_password" } } },
                new { name = "remove_database_password", description = "Remove the current database password (compact/repair based).", inputSchema = new { type = "object", properties = new { } } },
                new { name = "encrypt_database", description = "Compact/encrypt the current database with a password.", inputSchema = new { type = "object", properties = new { password = new { type = "string" } } } },
                new { name = "get_navigation_groups", description = "List Access navigation pane groups.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "set_display_categories", description = "Show or hide navigation pane display categories.", inputSchema = new { type = "object", properties = new { show_categories = new { type = "boolean" } } } },
                new { name = "refresh_database_window", description = "Refresh the Access navigation/database window.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "create_navigation_group", description = "Create a navigation pane group.", inputSchema = new { type = "object", properties = new { group_name = new { type = "string" } }, required = new string[] { "group_name" } } },
                new { name = "add_navigation_group_object", description = "Add an object to a navigation pane group.", inputSchema = new { type = "object", properties = new { group_name = new { type = "string" }, object_name = new { type = "string" }, object_type = new { type = "string" } }, required = new string[] { "group_name", "object_name" } } },
                new { name = "delete_navigation_group", description = "Delete a navigation pane group.", inputSchema = new { type = "object", properties = new { group_name = new { type = "string" } }, required = new string[] { "group_name" } } },
                new { name = "remove_navigation_group_object", description = "Remove an object from a navigation pane group.", inputSchema = new { type = "object", properties = new { group_name = new { type = "string" }, object_name = new { type = "string" } }, required = new string[] { "group_name", "object_name" } } },
                new { name = "set_navigation_pane_visibility", description = "Show or hide the Access navigation pane.", inputSchema = new { type = "object", properties = new { visible = new { type = "boolean" } } } },
                new { name = "get_navigation_group_objects", description = "List objects in a navigation pane group.", inputSchema = new { type = "object", properties = new { group_name = new { type = "string" } }, required = new string[] { "group_name" } } },
                new { name = "get_conditional_formatting", description = "Get conditional formatting rules for a form/report control.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "object_type", "object_name", "control_name" } } },
                new { name = "add_conditional_formatting", description = "Add a conditional formatting rule to a form/report control.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" }, control_name = new { type = "string" }, expression = new { type = "string" }, fore_color = new { type = "integer" }, back_color = new { type = "integer" } }, required = new string[] { "object_type", "object_name", "control_name", "expression" } } },
                new { name = "delete_conditional_formatting", description = "Delete a conditional formatting rule by index.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" }, control_name = new { type = "string" }, rule_index = new { type = "integer" } }, required = new string[] { "object_type", "object_name", "control_name", "rule_index" } } },
                new { name = "update_conditional_formatting", description = "Update properties of a conditional formatting rule.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" }, control_name = new { type = "string" }, rule_index = new { type = "integer" }, expression = new { type = "string" }, fore_color = new { type = "integer" }, back_color = new { type = "integer" }, enabled = new { type = "boolean" } }, required = new string[] { "object_type", "object_name", "control_name", "rule_index" } } },
                new { name = "clear_conditional_formatting", description = "Clear all conditional formatting rules from a control.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "object_type", "object_name", "control_name" } } },
                new { name = "list_all_conditional_formats", description = "List all controls with conditional formatting on a form/report.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" } }, required = new string[] { "object_type", "object_name" } } },
                new { name = "get_attachment_files", description = "List files from an attachment field.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, where_condition = new { type = "string" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "add_attachment_file", description = "Add a file into an attachment field.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, file_path = new { type = "string" }, where_condition = new { type = "string" } }, required = new string[] { "table_name", "field_name", "file_path" } } },
                new { name = "remove_attachment_file", description = "Remove a file from an attachment field.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, file_name = new { type = "string" }, where_condition = new { type = "string" } }, required = new string[] { "table_name", "field_name", "file_name" } } },
                new { name = "save_attachment_to_disk", description = "Save an attachment file to a destination path.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, file_path = new { type = "string" }, file_name = new { type = "string" }, where_condition = new { type = "string" } }, required = new string[] { "table_name", "field_name", "file_path" } } },
                new { name = "get_attachment_metadata", description = "Get detailed metadata for attachment field files.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, field_name = new { type = "string" }, where_condition = new { type = "string" } }, required = new string[] { "table_name", "field_name" } } },
                new { name = "get_object_events", description = "Get object event bindings from Access form/report objects.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" } }, required = new string[] { "object_type", "object_name" } } },
                new { name = "set_object_event", description = "Set object event binding on Access form/report objects.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string" }, object_name = new { type = "string" }, event_name = new { type = "string" }, event_value = new { type = "string" } }, required = new string[] { "object_type", "object_name", "event_name", "event_value" } } },
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
                new { name = "get_form_record_count", description = "Get record count from an open form.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_form_current_record", description = "Get current record data from an open form.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "set_form_filter", description = "Set Filter and FilterOn on a form at runtime.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, filter = new { type = "string" }, filter_on = new { type = "boolean" } }, required = new string[] { "form_name" } } },
                new { name = "get_open_objects", description = "List currently open Access objects.", inputSchema = new { type = "object", properties = new { } } },
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
                new { name = "execute_vba", description = "Evaluate a VBA expression using Application.Eval. WARNING: This can execute arbitrary VBA code and should only be used with trusted input.", inputSchema = new { type = "object", properties = new { expression = new { type = "string" } }, required = new string[] { "expression" } } },
                new { name = "run_vba_procedure", description = "Run a named VBA Sub/Function using Application.Run.", inputSchema = new { type = "object", properties = new { procedure_name = new { type = "string" }, args = new { type = "array" } }, required = new string[] { "procedure_name" } } },
                new { name = "create_module", description = "Create a new standard VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" } }, required = new string[] { "module_name" } } },
                new { name = "delete_module", description = "Delete a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" } }, required = new string[] { "module_name" } } },
                new { name = "rename_module", description = "Rename a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, new_module_name = new { type = "string" } }, required = new string[] { "module_name", "new_module_name" } } },
                new { name = "get_compilation_errors", description = "Compile VBA and return any compilation errors.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "list_all_procedures", description = "List procedures across all modules in a VBA project.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" } } } },
                new { name = "get_vba_project_properties", description = "Get VBA project properties.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" } } } },
                new { name = "set_vba_project_properties", description = "Set VBA project properties.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, name = new { type = "string" }, description = new { type = "string" }, help_file = new { type = "string" }, help_context_id = new { type = "integer" } } } },
                new { name = "get_module_info", description = "Get metadata for a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" } }, required = new string[] { "module_name" } } },
                new { name = "list_procedures", description = "List procedures in a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" } }, required = new string[] { "module_name" } } },
                new { name = "get_procedure_code", description = "Get code for a single procedure in a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, procedure_name = new { type = "string" } }, required = new string[] { "module_name", "procedure_name" } } },
                new { name = "get_module_declarations", description = "Get module-level declarations from a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" } }, required = new string[] { "module_name" } } },
                new { name = "insert_lines", description = "Insert lines into a VBA module at a specific line number.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, line_number = new { type = "integer" }, code = new { type = "string" } }, required = new string[] { "module_name", "line_number", "code" } } },
                new { name = "delete_lines", description = "Delete one or more lines from a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, start_line = new { type = "integer" }, line_count = new { type = "integer" } }, required = new string[] { "module_name", "start_line" } } },
                new { name = "replace_line", description = "Replace a line in a VBA module.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, line_number = new { type = "integer" }, code = new { type = "string" } }, required = new string[] { "module_name", "line_number", "code" } } },
                new { name = "find_text_in_module", description = "Find text in a VBA module using CodeModule.Find.", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, find_text = new { type = "string" }, start_line = new { type = "integer" }, start_column = new { type = "integer" }, end_line = new { type = "integer" }, end_column = new { type = "integer" }, whole_word = new { type = "boolean" }, match_case = new { type = "boolean" }, pattern_search = new { type = "boolean" } }, required = new string[] { "module_name", "find_text" } } },
                new { name = "list_import_export_specs", description = "List saved import/export specifications.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_import_export_spec", description = "Get details for a saved import/export specification.", inputSchema = new { type = "object", properties = new { specification_name = new { type = "string" } }, required = new string[] { "specification_name" } } },
                new { name = "create_import_export_spec", description = "Create an import/export specification from XML.", inputSchema = new { type = "object", properties = new { specification_name = new { type = "string" }, specification_xml = new { type = "string" } }, required = new string[] { "specification_name", "specification_xml" } } },
                new { name = "delete_import_export_spec", description = "Delete a saved import/export specification.", inputSchema = new { type = "object", properties = new { specification_name = new { type = "string" } }, required = new string[] { "specification_name" } } },
                new { name = "run_import_export_spec", description = "Run a saved import/export specification.", inputSchema = new { type = "object", properties = new { specification_name = new { type = "string" } }, required = new string[] { "specification_name" } } },
                new { name = "get_system_tables", description = "Get list of system tables", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_containers", description = "List DAO Containers in the current database.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_container_documents", description = "List DAO Documents in a Container.", inputSchema = new { type = "object", properties = new { container_name = new { type = "string" } }, required = new string[] { "container_name" } } },
                new { name = "get_document_properties", description = "Get DAO Document properties.", inputSchema = new { type = "object", properties = new { container_name = new { type = "string" }, document_name = new { type = "string" } }, required = new string[] { "container_name", "document_name" } } },
                new { name = "set_document_property", description = "Set DAO Document property.", inputSchema = new { type = "object", properties = new { container_name = new { type = "string" }, document_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" }, property_type = new { type = "string" }, create_if_missing = new { type = "boolean" } }, required = new string[] { "container_name", "document_name", "property_name", "value" } } },
                new { name = "get_object_metadata", description = "Get metadata for database objects", inputSchema = new { type = "object", properties = new { } } },
                new { name = "form_exists", description = "Check if a form exists", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_form_controls", description = "Get list of controls in a form", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_control_properties", description = "Get properties of a control", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "form_name", "control_name" } } },
                new { name = "set_control_property", description = "Set a property of a control", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "form_name", "control_name", "property_name", "value" } } },
                new { name = "get_report_controls", description = "Get list of controls in a report", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "get_report_control_properties", description = "Get properties of a report control", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "report_name", "control_name" } } },
                new { name = "set_report_control_property", description = "Set a property of a report control", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, control_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "report_name", "control_name", "property_name", "value" } } },
                new { name = "get_form_sections", description = "Get design-time sections from a form.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_report_sections", description = "Get design-time sections from a report.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "set_section_property", description = "Set a form/report section property in design view.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "form or report" }, object_name = new { type = "string" }, section = new { type = "string", description = "Section name or section index" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "object_type", "object_name", "section", "property_name", "value" } } },
                new { name = "create_control", description = "Create a control on a form in design view.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_type = new { type = "string" }, control_name = new { type = "string" }, section = new { type = "integer" }, parent_control_name = new { type = "string" }, column_name = new { type = "string" }, left = new { type = "integer" }, top = new { type = "integer" }, width = new { type = "integer" }, height = new { type = "integer" } }, required = new string[] { "form_name", "control_type" } } },
                new { name = "create_report_control", description = "Create a control on a report in design view.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, control_type = new { type = "string" }, control_name = new { type = "string" }, section = new { type = "integer" }, parent_control_name = new { type = "string" }, column_name = new { type = "string" }, left = new { type = "integer" }, top = new { type = "integer" }, width = new { type = "integer" }, height = new { type = "integer" } }, required = new string[] { "report_name", "control_type" } } },
                new { name = "delete_control", description = "Delete a control from a form in design view.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "form_name", "control_name" } } },
                new { name = "delete_report_control", description = "Delete a control from a report in design view.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "report_name", "control_name" } } },
                new { name = "get_form_properties", description = "Get form design properties.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "set_form_property", description = "Set a form design property.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "form_name", "property_name", "value" } } },
                new { name = "set_form_record_source", description = "Set form RecordSource at design time.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, record_source = new { type = "string" } }, required = new string[] { "form_name", "record_source" } } },
                new { name = "get_report_properties", description = "Get report design properties.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "set_report_property", description = "Set a report design property.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "report_name", "property_name", "value" } } },
                new { name = "set_report_record_source", description = "Set report RecordSource at design time.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, record_source = new { type = "string" } }, required = new string[] { "report_name", "record_source" } } },
                new { name = "get_report_grouping", description = "Get GroupLevel configuration for a report.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "set_report_grouping", description = "Add or modify report GroupLevel configuration.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, expression = new { type = "string" }, index = new { type = "integer" }, sort_order = new { type = "integer" }, group_on = new { type = "integer" }, group_interval = new { type = "integer" }, group_header = new { type = "boolean" }, group_footer = new { type = "boolean" }, keep_together = new { type = "integer" } }, required = new string[] { "report_name" } } },
                new { name = "delete_report_grouping", description = "Delete a GroupLevel from a report.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, index = new { type = "integer" } }, required = new string[] { "report_name", "index" } } },
                new { name = "get_report_sorting", description = "Get report sorting configuration.", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "get_tab_order", description = "Get tab order for form controls.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "set_tab_order", description = "Set tab order for form controls.", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_names = new { type = "array", items = new { type = "string" } } }, required = new string[] { "form_name", "control_names" } } },
                new { name = "get_page_setup", description = "Get page setup properties for a form or report.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "form or report" }, object_name = new { type = "string" } }, required = new string[] { "object_type", "object_name" } } },
                new { name = "set_page_setup", description = "Set page setup properties for a form or report.", inputSchema = new { type = "object", properties = new { object_type = new { type = "string", description = "form or report" }, object_name = new { type = "string" }, top_margin = new { type = "integer" }, bottom_margin = new { type = "integer" }, left_margin = new { type = "integer" }, right_margin = new { type = "integer" }, orientation = new { type = "integer" }, paper_size = new { type = "integer" }, data_only = new { type = "boolean" } }, required = new string[] { "object_type", "object_name" } } },
                new { name = "get_printer_info", description = "Get current printer and installed printer details.", inputSchema = new { type = "object", properties = new { } } },
                new { name = "export_form_to_text", description = "Export a form to text format", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "form_name" } } },
                new { name = "import_form_from_text", description = "Import a form from text format", inputSchema = new { type = "object", properties = new { form_data = new { type = "string" }, form_name = new { type = "string", description = "Optional form name override. Required for some access_text payloads." }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "form_data" } } },
                new { name = "delete_form", description = "Delete a form from the database", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "export_report_to_text", description = "Export a report to text format", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "report_name" } } },
                new { name = "import_report_from_text", description = "Import a report from text format", inputSchema = new { type = "object", properties = new { report_data = new { type = "string" }, report_name = new { type = "string", description = "Optional report name override. Required for some access_text payloads." }, mode = new { type = "string", @enum = new[] { "json", "access_text" }, description = "Optional mode. Defaults to json." } }, required = new string[] { "report_data" } } },
                new { name = "delete_report", description = "Delete a report from the database", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                // Priority 17: DoCmd Remaining + Domain Aggregates
                new { name = "find_next", description = "Continue search after FindRecord (DoCmd.FindNext). Must be preceded by a FindRecord call.", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "search_for_record", description = "Search for a record matching criteria in the active or specified object (DoCmd.SearchForRecord).", inputSchema = new { type = "object", properties = new { object_type = new { type = "integer", description = "Optional AcDataObjectType enum: -1=ActiveDataObject (default), 0=Table, 1=Query, 2=Form, 3=ServerView, 4=StoredProcedure" }, object_name = new { type = "string", description = "Optional name of the table, query, or form to search" }, record = new { type = "string", @enum = new[] { "", "First", "Last", "Next", "Previous" }, description = "Optional search direction. Default is empty (searches all)." }, where_condition = new { type = "string", description = "WHERE clause criteria string (without the WHERE keyword), e.g. \"[ID] = 5\"" } }, required = new string[] { "where_condition" } } },
                new { name = "set_filter_docmd", description = "Apply a filter to the active datasheet, form, report, or table via DoCmd.SetFilter.", inputSchema = new { type = "object", properties = new { filter_name = new { type = "string", description = "Optional name of a saved filter (query) to apply" }, where_condition = new { type = "string", description = "Optional WHERE clause (without WHERE keyword) to filter records" } }, required = new string[] { } } },
                new { name = "set_order_by", description = "Set the sort order for the active form, report, datasheet, or server view (DoCmd.SetOrderBy).", inputSchema = new { type = "object", properties = new { order_by = new { type = "string", description = "ORDER BY clause (without ORDER BY keyword), e.g. \"[LastName] ASC, [FirstName] DESC\"" } }, required = new string[] { "order_by" } } },
                new { name = "set_parameter", description = "Set a parameter value before opening a parameterized query, form, or report (DoCmd.SetParameter). Must be called before the OpenQuery/OpenForm/OpenReport call.", inputSchema = new { type = "object", properties = new { name = new { type = "string", description = "Parameter name as defined in the query/form/report" }, expression = new { type = "string", description = "Expression or value for the parameter" } }, required = new string[] { "name", "expression" } } },
                new { name = "set_runtime_property", description = "Set a control property at runtime by name (DoCmd.SetProperty). Works on the active form/report.", inputSchema = new { type = "object", properties = new { control_name = new { type = "string", description = "Name of the control whose property to set" }, property_id = new { type = "integer", description = "AcProperty enum value: 0=Value, 1=Enabled, 2=Visible, 3=Locked, 4=Left, 5=Top, 6=Width, 7=Height, 8=ForeColor, 9=BackColor, 10=Caption" }, value = new { type = "string", description = "New value for the property" } }, required = new string[] { "control_name", "property_id", "value" } } },
                new { name = "refresh_record", description = "Refresh the data in the active form by requerying the underlying record source (DoCmd.RefreshRecord).", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "close_database", description = "Close the current database without quitting the Access application (DoCmd.CloseDatabase). Resets the connection state.", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "domain_aggregate", description = "Execute a domain aggregate function (DLookup, DCount, DSum, DAvg, DMin, DMax, DFirst, DLast) against a table or query domain.", inputSchema = new { type = "object", properties = new { function = new { type = "string", @enum = new[] { "DLookup", "DCount", "DSum", "DAvg", "DMin", "DMax", "DFirst", "DLast" }, description = "Domain aggregate function to call" }, expression = new { type = "string", description = "Expression to evaluate, e.g. \"[Price]\" or \"Count(*)\"" }, domain = new { type = "string", description = "Table or query name to use as the domain" }, criteria = new { type = "string", description = "Optional WHERE clause criteria (without WHERE keyword), e.g. \"[Category] = 'Books'\"" } }, required = new string[] { "function", "expression", "domain" } } },
                new { name = "access_error", description = "Get the error description string for a given Access/VBA error number (Application.AccessError).", inputSchema = new { type = "object", properties = new { error_number = new { type = "integer", description = "The Access or VBA error number to look up" } }, required = new string[] { "error_number" } } },
                new { name = "build_criteria", description = "Build a properly formatted criteria string for use in domain aggregates, filters, or queries (Application.BuildCriteria).", inputSchema = new { type = "object", properties = new { field = new { type = "string", description = "Field name, e.g. \"[LastName]\"" }, field_type = new { type = "integer", description = "DAO DataTypeEnum value: 1=dbBoolean, 2=dbByte, 3=dbInteger, 4=dbLong, 5=dbCurrency, 6=dbSingle, 7=dbDouble, 8=dbDate, 10=dbText, 12=dbMemo" }, expression = new { type = "string", description = "Value or expression to match, e.g. \"Smith\"" } }, required = new string[] { "field", "field_type", "expression" } } },
                // Priority 18: Screen Object + Visibility + App Info
                new { name = "get_active_form", description = "Get info about the form that currently has focus (Screen.ActiveForm). Returns form name, record source, caption, current record, and dirty state.", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "get_active_report", description = "Get info about the active report (Screen.ActiveReport). Returns report name, record source, and caption.", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "get_active_control", description = "Get info about the control that currently has focus (Screen.ActiveControl). Returns control name, type, and value.", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "get_active_datasheet", description = "Get info about the active datasheet (Screen.ActiveDatasheet). Returns datasheet name and record source.", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "set_hidden_attribute", description = "Hide or unhide a database object in the Navigation Pane (Application.SetHiddenAttribute). object_type: 0=Table, 1=Query, 2=Form, 3=Report, 4=Macro, 5=Module.", inputSchema = new { type = "object", properties = new { object_type = new { type = "integer", description = "AcObjectType enum value: 0=Table, 1=Query, 2=Form, 3=Report, 4=Macro, 5=Module" }, object_name = new { type = "string", description = "Name of the database object" }, hidden = new { type = "boolean", description = "True to hide, false to unhide" } }, required = new string[] { "object_type", "object_name", "hidden" } } },
                new { name = "get_hidden_attribute", description = "Check if a database object is hidden in the Navigation Pane (Application.GetHiddenAttribute). object_type: 0=Table, 1=Query, 2=Form, 3=Report, 4=Macro, 5=Module.", inputSchema = new { type = "object", properties = new { object_type = new { type = "integer", description = "AcObjectType enum value: 0=Table, 1=Query, 2=Form, 3=Report, 4=Macro, 5=Module" }, object_name = new { type = "string", description = "Name of the database object" } }, required = new string[] { "object_type", "object_name" } } },
                new { name = "get_current_object", description = "Get the name and type of the currently selected/active database object (Application.CurrentObjectName and CurrentObjectType).", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "get_current_user", description = "Get the current user name (Application.CurrentUser).", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } },
                new { name = "set_access_visible", description = "Show or hide the Access application window (Application.Visible).", inputSchema = new { type = "object", properties = new { visible = new { type = "boolean", description = "True to show, false to hide the Access window" } }, required = new string[] { "visible" } } },
                new { name = "get_access_hwnd", description = "Get the window handle (hWnd) of the Access application window (Application.hWndAccessApp).", inputSchema = new { type = "object", properties = new { }, required = new string[] { } } }
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
            "open_table" => HandleOpenTable(accessService, toolArguments),
            "open_module" => HandleOpenModule(accessService, toolArguments),
            "copy_object" => HandleCopyObject(accessService, toolArguments),
            "delete_object" => HandleDeleteObject(accessService, toolArguments),
            "rename_object" => HandleRenameObject(accessService, toolArguments),
            "select_object" => HandleSelectObject(accessService, toolArguments),
            "save_object" => HandleSaveObject(accessService, toolArguments),
            "close_object" => HandleCloseObject(accessService, toolArguments),
            "transfer_database" => HandleTransferDatabase(accessService, toolArguments),
            "run_command" => HandleRunCommand(accessService, toolArguments),
            "sys_cmd" => HandleSysCmd(accessService, toolArguments),
            "goto_page" => HandleGoToPage(accessService, toolArguments),
            "goto_control" => HandleGoToControl(accessService, toolArguments),
            "move_size" => HandleMoveSize(accessService, toolArguments),
            "requery" => HandleRequery(accessService, toolArguments),
            "repaint_object" => HandleRepaintObject(accessService, toolArguments),
            "send_object" => HandleSendObject(accessService, toolArguments),
            "browse_to" => HandleBrowseTo(accessService, toolArguments),
            "lock_navigation_pane" => HandleLockNavigationPane(accessService, toolArguments),
            "navigate_to" => HandleNavigateTo(accessService, toolArguments),
            "beep" => HandleBeep(accessService, toolArguments),
            "get_database_summary_properties" => HandleGetDatabaseSummaryProperties(accessService, toolArguments),
            "set_database_summary_properties" => HandleSetDatabaseSummaryProperties(accessService, toolArguments),
            "get_database_properties" => HandleGetDatabaseProperties(accessService, toolArguments),
            "get_database_property" => HandleGetDatabaseProperty(accessService, toolArguments),
            "set_database_property" => HandleSetDatabaseProperty(accessService, toolArguments),
            "get_table_properties" => HandleGetTableProperties(accessService, toolArguments),
            "set_table_properties" => HandleSetTableProperties(accessService, toolArguments),
            "get_table_validation" => HandleGetTableValidation(accessService, toolArguments),
            "get_table_description" => HandleGetTableDescription(accessService, toolArguments),
            "set_table_description" => HandleSetTableDescription(accessService, toolArguments),
            "get_all_field_descriptions" => HandleGetAllFieldDescriptions(accessService, toolArguments),
            "get_query_properties" => HandleGetQueryProperties(accessService, toolArguments),
            "get_query_parameters" => HandleGetQueryParameters(accessService, toolArguments),
            "set_query_properties" => HandleSetQueryProperties(accessService, toolArguments),
            "set_field_validation" => HandleSetFieldValidation(accessService, toolArguments),
            "set_field_default" => HandleSetFieldDefault(accessService, toolArguments),
            "set_field_input_mask" => HandleSetFieldInputMask(accessService, toolArguments),
            "set_field_caption" => HandleSetFieldCaption(accessService, toolArguments),
            "get_field_properties" => HandleGetFieldProperties(accessService, toolArguments),
            "get_field_attributes" => HandleGetFieldAttributes(accessService, toolArguments),
            "detect_multi_value_fields" => HandleDetectMultiValueFields(accessService, toolArguments),
            "get_multi_value_field_values" => HandleGetMultiValueFieldValues(accessService, toolArguments),
            "set_multi_value_field_values" => HandleSetMultiValueFieldValues(accessService, toolArguments),
            "set_lookup_properties" => HandleSetLookupProperties(accessService, toolArguments),
            "get_vba_references" => HandleGetVbaReferences(accessService, toolArguments),
            "add_vba_reference" => HandleAddVbaReference(accessService, toolArguments),
            "remove_vba_reference" => HandleRemoveVbaReference(accessService, toolArguments),
            "get_startup_properties" => HandleGetStartupProperties(accessService, toolArguments),
            "set_startup_properties" => HandleSetStartupProperties(accessService, toolArguments),
            "get_ribbon_xml" => HandleGetRibbonXml(accessService, toolArguments),
            "set_ribbon_xml" => HandleSetRibbonXml(accessService, toolArguments),
            "get_application_info" => HandleGetApplicationInfo(accessService, toolArguments),
            "get_current_project_data" => HandleGetCurrentProjectData(accessService, toolArguments),
            "get_application_option" => HandleGetApplicationOption(accessService, toolArguments),
            "set_application_option" => HandleSetApplicationOption(accessService, toolArguments),
            "get_temp_vars" => HandleGetTempVars(accessService, toolArguments),
            "set_temp_var" => HandleSetTempVar(accessService, toolArguments),
            "remove_temp_var" => HandleRemoveTempVar(accessService, toolArguments),
            "clear_temp_vars" => HandleClearTempVars(accessService, toolArguments),
            "export_data_macro_axl" => HandleExportDataMacroAxl(accessService, toolArguments),
            "import_data_macro_axl" => HandleImportDataMacroAxl(accessService, toolArguments),
            "run_data_macro" => HandleRunDataMacro(accessService, toolArguments),
            "get_table_data_macros" => HandleGetTableDataMacros(accessService, toolArguments),
            "delete_data_macro" => HandleDeleteDataMacro(accessService, toolArguments),
            "get_autoexec_info" => HandleGetAutoExecInfo(accessService, toolArguments),
            "run_autoexec" => HandleRunAutoExec(accessService, toolArguments),
            "get_database_security" => HandleGetDatabaseSecurity(accessService, toolArguments),
            "set_database_password" => HandleSetDatabasePassword(accessService, toolArguments),
            "remove_database_password" => HandleRemoveDatabasePassword(accessService, toolArguments),
            "encrypt_database" => HandleEncryptDatabase(accessService, toolArguments),
            "get_navigation_groups" => HandleGetNavigationGroups(accessService, toolArguments),
            "set_display_categories" => HandleSetDisplayCategories(accessService, toolArguments),
            "refresh_database_window" => HandleRefreshDatabaseWindow(accessService, toolArguments),
            "create_navigation_group" => HandleCreateNavigationGroup(accessService, toolArguments),
            "add_navigation_group_object" => HandleAddNavigationGroupObject(accessService, toolArguments),
            "delete_navigation_group" => HandleDeleteNavigationGroup(accessService, toolArguments),
            "remove_navigation_group_object" => HandleRemoveNavigationGroupObject(accessService, toolArguments),
            "set_navigation_pane_visibility" => HandleSetNavigationPaneVisibility(accessService, toolArguments),
            "get_navigation_group_objects" => HandleGetNavigationGroupObjects(accessService, toolArguments),
            "get_conditional_formatting" => HandleGetConditionalFormatting(accessService, toolArguments),
            "add_conditional_formatting" => HandleAddConditionalFormatting(accessService, toolArguments),
            "delete_conditional_formatting" => HandleDeleteConditionalFormatting(accessService, toolArguments),
            "update_conditional_formatting" => HandleUpdateConditionalFormatting(accessService, toolArguments),
            "clear_conditional_formatting" => HandleClearConditionalFormatting(accessService, toolArguments),
            "list_all_conditional_formats" => HandleListAllConditionalFormats(accessService, toolArguments),
            "get_attachment_files" => HandleGetAttachmentFiles(accessService, toolArguments),
            "add_attachment_file" => HandleAddAttachmentFile(accessService, toolArguments),
            "remove_attachment_file" => HandleRemoveAttachmentFile(accessService, toolArguments),
            "save_attachment_to_disk" => HandleSaveAttachmentToDisk(accessService, toolArguments),
            "get_attachment_metadata" => HandleGetAttachmentMetadata(accessService, toolArguments),
            "get_object_events" => HandleGetObjectEvents(accessService, toolArguments),
            "set_object_event" => HandleSetObjectEvent(accessService, toolArguments),
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
            "get_form_record_count" => HandleGetFormRecordCount(accessService, toolArguments),
            "get_form_current_record" => HandleGetFormCurrentRecord(accessService, toolArguments),
            "set_form_filter" => HandleSetFormFilter(accessService, toolArguments),
            "get_open_objects" => HandleGetOpenObjects(accessService, toolArguments),
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
            "execute_vba" => HandleExecuteVba(accessService, toolArguments),
            "run_vba_procedure" => HandleRunVbaProcedure(accessService, toolArguments),
            "create_module" => HandleCreateModule(accessService, toolArguments),
            "delete_module" => HandleDeleteModule(accessService, toolArguments),
            "rename_module" => HandleRenameModule(accessService, toolArguments),
            "get_compilation_errors" => HandleGetCompilationErrors(accessService, toolArguments),
            "list_all_procedures" => HandleListAllProcedures(accessService, toolArguments),
            "get_vba_project_properties" => HandleGetVbaProjectProperties(accessService, toolArguments),
            "set_vba_project_properties" => HandleSetVbaProjectProperties(accessService, toolArguments),
            "get_module_info" => HandleGetModuleInfo(accessService, toolArguments),
            "list_procedures" => HandleListProcedures(accessService, toolArguments),
            "get_procedure_code" => HandleGetProcedureCode(accessService, toolArguments),
            "get_module_declarations" => HandleGetModuleDeclarations(accessService, toolArguments),
            "insert_lines" => HandleInsertLines(accessService, toolArguments),
            "delete_lines" => HandleDeleteLines(accessService, toolArguments),
            "replace_line" => HandleReplaceLine(accessService, toolArguments),
            "find_text_in_module" => HandleFindTextInModule(accessService, toolArguments),
            "list_import_export_specs" => HandleListImportExportSpecs(accessService, toolArguments),
            "get_import_export_spec" => HandleGetImportExportSpec(accessService, toolArguments),
            "create_import_export_spec" => HandleCreateImportExportSpec(accessService, toolArguments),
            "delete_import_export_spec" => HandleDeleteImportExportSpec(accessService, toolArguments),
            "run_import_export_spec" => HandleRunImportExportSpec(accessService, toolArguments),
            "get_system_tables" => HandleGetSystemTables(accessService, toolArguments),
            "get_containers" => HandleGetContainers(accessService, toolArguments),
            "get_container_documents" => HandleGetContainerDocuments(accessService, toolArguments),
            "get_document_properties" => HandleGetDocumentProperties(accessService, toolArguments),
            "set_document_property" => HandleSetDocumentProperty(accessService, toolArguments),
            "get_object_metadata" => HandleGetObjectMetadata(accessService, toolArguments),
            "form_exists" => HandleFormExists(accessService, toolArguments),
            "get_form_controls" => HandleGetFormControls(accessService, toolArguments),
            "get_control_properties" => HandleGetControlProperties(accessService, toolArguments),
            "set_control_property" => HandleSetControlProperty(accessService, toolArguments),
            "get_report_controls" => HandleGetReportControls(accessService, toolArguments),
            "get_report_control_properties" => HandleGetReportControlProperties(accessService, toolArguments),
            "set_report_control_property" => HandleSetReportControlProperty(accessService, toolArguments),
            "get_form_sections" => HandleGetFormSections(accessService, toolArguments),
            "get_report_sections" => HandleGetReportSections(accessService, toolArguments),
            "set_section_property" => HandleSetSectionProperty(accessService, toolArguments),
            "create_control" => HandleCreateControl(accessService, toolArguments),
            "create_report_control" => HandleCreateReportControl(accessService, toolArguments),
            "delete_control" => HandleDeleteControl(accessService, toolArguments),
            "delete_report_control" => HandleDeleteReportControl(accessService, toolArguments),
            "get_form_properties" => HandleGetFormProperties(accessService, toolArguments),
            "set_form_property" => HandleSetFormProperty(accessService, toolArguments),
            "set_form_record_source" => HandleSetFormRecordSource(accessService, toolArguments),
            "get_report_properties" => HandleGetReportProperties(accessService, toolArguments),
            "set_report_property" => HandleSetReportProperty(accessService, toolArguments),
            "set_report_record_source" => HandleSetReportRecordSource(accessService, toolArguments),
            "get_report_grouping" => HandleGetReportGrouping(accessService, toolArguments),
            "set_report_grouping" => HandleSetReportGrouping(accessService, toolArguments),
            "delete_report_grouping" => HandleDeleteReportGrouping(accessService, toolArguments),
            "get_report_sorting" => HandleGetReportSorting(accessService, toolArguments),
            "get_tab_order" => HandleGetTabOrder(accessService, toolArguments),
            "set_tab_order" => HandleSetTabOrder(accessService, toolArguments),
            "get_page_setup" => HandleGetPageSetup(accessService, toolArguments),
            "set_page_setup" => HandleSetPageSetup(accessService, toolArguments),
            "get_printer_info" => HandleGetPrinterInfo(accessService, toolArguments),
            "export_form_to_text" => HandleExportFormToText(accessService, toolArguments),
            "import_form_from_text" => HandleImportFormFromText(accessService, toolArguments),
            "delete_form" => HandleDeleteForm(accessService, toolArguments),
            "export_report_to_text" => HandleExportReportToText(accessService, toolArguments),
            "import_report_from_text" => HandleImportReportFromText(accessService, toolArguments),
            "delete_report" => HandleDeleteReport(accessService, toolArguments),
            // Priority 17: DoCmd Remaining + Domain Aggregates
            "find_next" => HandleFindNext(accessService, toolArguments),
            "search_for_record" => HandleSearchForRecord(accessService, toolArguments),
            "set_filter_docmd" => HandleSetFilterDoCmd(accessService, toolArguments),
            "set_order_by" => HandleSetOrderBy(accessService, toolArguments),
            "set_parameter" => HandleSetParameter(accessService, toolArguments),
            "set_runtime_property" => HandleSetRuntimeProperty(accessService, toolArguments),
            "refresh_record" => HandleRefreshRecord(accessService, toolArguments),
            "close_database" => HandleCloseDatabase(accessService, toolArguments),
            "domain_aggregate" => HandleDomainAggregate(accessService, toolArguments),
            "access_error" => HandleAccessError(accessService, toolArguments),
            "build_criteria" => HandleBuildCriteria(accessService, toolArguments),
            // Priority 18: Screen Object + Visibility + App Info
            "get_active_form" => HandleGetActiveForm(accessService, toolArguments),
            "get_active_report" => HandleGetActiveReport(accessService, toolArguments),
            "get_active_control" => HandleGetActiveControl(accessService, toolArguments),
            "get_active_datasheet" => HandleGetActiveDatasheet(accessService, toolArguments),
            "set_hidden_attribute" => HandleSetHiddenAttribute(accessService, toolArguments),
            "get_hidden_attribute" => HandleGetHiddenAttribute(accessService, toolArguments),
            "get_current_object" => HandleGetCurrentObject(accessService, toolArguments),
            "get_current_user" => HandleGetCurrentUser(accessService, toolArguments),
            "set_access_visible" => HandleSetAccessVisible(accessService, toolArguments),
            "get_access_hwnd" => HandleGetAccessHwnd(accessService, toolArguments),
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

    static object HandleOpenTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            _ = TryGetOptionalString(arguments, "view", out var view);
            _ = TryGetOptionalString(arguments, "data_mode", out var dataMode);

            accessService.OpenTable(
                tableName,
                string.IsNullOrWhiteSpace(view) ? null : view,
                string.IsNullOrWhiteSpace(dataMode) ? null : dataMode);

            return new
            {
                success = true,
                table_name = tableName,
                view = string.IsNullOrWhiteSpace(view) ? null : view,
                data_mode = string.IsNullOrWhiteSpace(dataMode) ? null : dataMode
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("open_table", ex);
        }
    }

    static object HandleOpenModule(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;

            _ = TryGetOptionalString(arguments, "procedure_name", out var procedureName);
            accessService.OpenModule(moduleName, string.IsNullOrWhiteSpace(procedureName) ? null : procedureName);

            return new
            {
                success = true,
                module_name = moduleName,
                procedure_name = string.IsNullOrWhiteSpace(procedureName) ? null : procedureName
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("open_module", ex);
        }
    }

    static object HandleCopyObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "source_object_name", out var sourceObjectName, out var sourceObjectNameError))
                return sourceObjectNameError;

            _ = TryGetOptionalString(arguments, "source_object_type", out var sourceObjectType);
            _ = TryGetOptionalString(arguments, "destination_database_path", out var destinationDatabasePath);
            _ = TryGetOptionalString(arguments, "new_name", out var newName);

            accessService.CopyObject(
                string.IsNullOrWhiteSpace(destinationDatabasePath) ? null : destinationDatabasePath,
                string.IsNullOrWhiteSpace(newName) ? null : newName,
                string.IsNullOrWhiteSpace(sourceObjectType) ? null : sourceObjectType,
                sourceObjectName);

            return new
            {
                success = true,
                source_object_name = sourceObjectName,
                source_object_type = string.IsNullOrWhiteSpace(sourceObjectType) ? null : sourceObjectType,
                destination_database_path = string.IsNullOrWhiteSpace(destinationDatabasePath) ? null : destinationDatabasePath,
                new_name = string.IsNullOrWhiteSpace(newName) ? null : newName
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("copy_object", ex);
        }
    }

    static object HandleDeleteObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            accessService.DeleteObject(objectName, string.IsNullOrWhiteSpace(objectType) ? null : objectType);

            return new
            {
                success = true,
                object_name = objectName,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_object", ex);
        }
    }

    static object HandleRenameObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "new_name", out var newName, out var newNameError))
                return newNameError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            accessService.RenameObject(newName, objectName, string.IsNullOrWhiteSpace(objectType) ? null : objectType);

            return new
            {
                success = true,
                new_name = newName,
                object_name = objectName,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("rename_object", ex);
        }
    }

    static object HandleSelectObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            var inDatabaseWindow = GetOptionalBool(arguments, "in_database_window", true);

            accessService.SelectObject(
                objectName,
                string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                inDatabaseWindow);

            return new
            {
                success = true,
                object_name = objectName,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                in_database_window = inDatabaseWindow
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("select_object", ex);
        }
    }

    static object HandleSaveObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            _ = TryGetOptionalString(arguments, "object_name", out var objectName);

            accessService.SaveObject(
                string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                string.IsNullOrWhiteSpace(objectName) ? null : objectName);

            return new
            {
                success = true,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                object_name = string.IsNullOrWhiteSpace(objectName) ? null : objectName
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("save_object", ex);
        }
    }

    static object HandleCloseObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            _ = TryGetOptionalString(arguments, "object_name", out var objectName);
            _ = TryGetOptionalString(arguments, "save", out var save);

            accessService.CloseObject(
                string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                string.IsNullOrWhiteSpace(objectName) ? null : objectName,
                string.IsNullOrWhiteSpace(save) ? null : save);

            return new
            {
                success = true,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                object_name = string.IsNullOrWhiteSpace(objectName) ? null : objectName,
                save = string.IsNullOrWhiteSpace(save) ? null : save
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("close_object", ex);
        }
    }

    static object HandleTransferDatabase(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "transfer_type", out var transferType, out var transferTypeError))
                return transferTypeError;
            if (!TryGetRequiredString(arguments, "database_type", out var databaseType, out var databaseTypeError))
                return databaseTypeError;
            if (!TryGetRequiredString(arguments, "database_name", out var databaseName, out var databaseNameError))
                return databaseNameError;
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "source", out var source, out var sourceError))
                return sourceError;

            _ = TryGetOptionalString(arguments, "destination", out var destination);
            var structureOnly = GetOptionalBool(arguments, "structure_only", false);
            var storeLogin = GetOptionalBool(arguments, "store_login", false);

            var result = accessService.TransferDatabase(
                transferType,
                databaseType,
                databaseName,
                objectType,
                source,
                string.IsNullOrWhiteSpace(destination) ? null : destination,
                structureOnly,
                storeLogin);

            return new
            {
                success = true,
                transfer_type = result.TransferType,
                database_type = result.DatabaseType,
                database_name = result.DatabaseName,
                object_type = result.ObjectType,
                source = result.Source,
                destination = result.Destination,
                structure_only = result.StructureOnly,
                store_login = result.StoreLogin
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("transfer_database", ex);
        }
    }

    static object HandleRunCommand(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "command", out var command, out var commandError))
                return commandError;

            accessService.RunCommand(command);
            return new { success = true, command };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("run_command", ex);
        }
    }

    static object HandleSysCmd(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "command", out var command, out var commandError))
                return commandError;
            if (!TryGetOptionalPrimitiveValue(arguments, "arg1", out var arg1, out var arg1Error))
                return arg1Error;
            if (!TryGetOptionalPrimitiveValue(arguments, "arg2", out var arg2, out var arg2Error))
                return arg2Error;
            if (!TryGetOptionalPrimitiveValue(arguments, "arg3", out var arg3, out var arg3Error))
                return arg3Error;

            var result = accessService.SysCmd(command, arg1, arg2, arg3);
            return new { success = true, result = result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("sys_cmd", ex);
        }
    }

    static object HandleGoToPage(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "page_number", out var pageNumber, out var pageNumberError))
                return pageNumberError;

            _ = TryGetOptionalString(arguments, "right", out var right);
            _ = TryGetOptionalString(arguments, "down", out var down);

            accessService.GoToPage(
                pageNumber,
                string.IsNullOrWhiteSpace(right) ? null : right,
                string.IsNullOrWhiteSpace(down) ? null : down);

            return new
            {
                success = true,
                page_number = pageNumber,
                right = string.IsNullOrWhiteSpace(right) ? null : right,
                down = string.IsNullOrWhiteSpace(down) ? null : down
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("goto_page", ex);
        }
    }

    static object HandleGoToControl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;

            accessService.GoToControl(controlName);
            return new { success = true, control_name = controlName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("goto_control", ex);
        }
    }

    static object HandleMoveSize(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetOptionalInt(arguments, "right", out var right, out var rightError))
                return rightError;
            if (!TryGetOptionalInt(arguments, "down", out var down, out var downError))
                return downError;
            if (!TryGetOptionalInt(arguments, "width", out var width, out var widthError))
                return widthError;
            if (!TryGetOptionalInt(arguments, "height", out var height, out var heightError))
                return heightError;

            accessService.MoveSize(right, down, width, height);
            return new { success = true, right, down, width, height };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("move_size", ex);
        }
    }

    static object HandleRequery(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "control_name", out var controlName);
            accessService.Requery(string.IsNullOrWhiteSpace(controlName) ? null : controlName);
            return new { success = true, control_name = string.IsNullOrWhiteSpace(controlName) ? null : controlName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("requery", ex);
        }
    }

    static object HandleRepaintObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            _ = TryGetOptionalString(arguments, "object_name", out var objectName);

            accessService.RepaintObject(
                string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                string.IsNullOrWhiteSpace(objectName) ? null : objectName);

            return new
            {
                success = true,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                object_name = string.IsNullOrWhiteSpace(objectName) ? null : objectName
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("repaint_object", ex);
        }
    }

    static object HandleSendObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            _ = TryGetOptionalString(arguments, "object_name", out var objectName);
            _ = TryGetOptionalString(arguments, "output_format", out var outputFormat);
            _ = TryGetOptionalString(arguments, "to", out var to);
            _ = TryGetOptionalString(arguments, "cc", out var cc);
            _ = TryGetOptionalString(arguments, "bcc", out var bcc);
            _ = TryGetOptionalString(arguments, "subject", out var subject);
            _ = TryGetOptionalString(arguments, "message_text", out var messageText);
            _ = TryGetOptionalString(arguments, "template_file", out var templateFile);

            if (!TryGetOptionalBoolNullable(arguments, "edit_message", out var editMessage, out var editMessageError))
                return editMessageError;

            accessService.SendObject(
                string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                string.IsNullOrWhiteSpace(objectName) ? null : objectName,
                string.IsNullOrWhiteSpace(outputFormat) ? null : outputFormat,
                string.IsNullOrWhiteSpace(to) ? null : to,
                string.IsNullOrWhiteSpace(cc) ? null : cc,
                string.IsNullOrWhiteSpace(bcc) ? null : bcc,
                string.IsNullOrWhiteSpace(subject) ? null : subject,
                string.IsNullOrWhiteSpace(messageText) ? null : messageText,
                editMessage,
                string.IsNullOrWhiteSpace(templateFile) ? null : templateFile);

            return new
            {
                success = true,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                object_name = string.IsNullOrWhiteSpace(objectName) ? null : objectName,
                output_format = string.IsNullOrWhiteSpace(outputFormat) ? null : outputFormat,
                to = string.IsNullOrWhiteSpace(to) ? null : to,
                cc = string.IsNullOrWhiteSpace(cc) ? null : cc,
                bcc = string.IsNullOrWhiteSpace(bcc) ? null : bcc,
                subject = string.IsNullOrWhiteSpace(subject) ? null : subject,
                message_text = string.IsNullOrWhiteSpace(messageText) ? null : messageText,
                edit_message = editMessage,
                template_file = string.IsNullOrWhiteSpace(templateFile) ? null : templateFile
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("send_object", ex);
        }
    }

    static object HandleBrowseTo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            _ = TryGetOptionalString(arguments, "path_to_subform_control", out var pathToSubformControl);
            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            _ = TryGetOptionalString(arguments, "page", out var page);

            accessService.BrowseTo(
                objectName,
                string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                string.IsNullOrWhiteSpace(pathToSubformControl) ? null : pathToSubformControl,
                string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition,
                string.IsNullOrWhiteSpace(page) ? null : page);

            return new
            {
                success = true,
                object_name = objectName,
                object_type = string.IsNullOrWhiteSpace(objectType) ? null : objectType,
                path_to_subform_control = string.IsNullOrWhiteSpace(pathToSubformControl) ? null : pathToSubformControl,
                where_condition = string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition,
                page = string.IsNullOrWhiteSpace(page) ? null : page
            };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("browse_to", ex);
        }
    }

    static object HandleLockNavigationPane(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var lockNavigationPane = GetOptionalBool(arguments, "lock_navigation_pane", true);
            accessService.LockNavigationPane(lockNavigationPane);
            return new { success = true, lock_navigation_pane = lockNavigationPane };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("lock_navigation_pane", ex);
        }
    }

    static object HandleNavigateTo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "navigation_category", out var navigationCategory, out var navigationCategoryError))
                return navigationCategoryError;

            accessService.NavigateTo(navigationCategory);
            return new { success = true, navigation_category = navigationCategory };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("navigate_to", ex);
        }
    }

    static object HandleBeep(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.Beep();
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("beep", ex);
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

    static object HandleGetTableValidation(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var validation = accessService.GetTableValidation(tableName);
            return new { success = true, validation = validation };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_table_validation", ex);
        }
    }

    static object HandleGetTableDescription(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var description = accessService.GetTableDescription(tableName);
            return new { success = true, table_name = tableName, description = description };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_table_description", ex);
        }
    }

    static object HandleSetTableDescription(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "description", out var description, out var descriptionError))
                return descriptionError;

            accessService.SetTableDescription(tableName, description);
            return new { success = true, table_name = tableName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_table_description", ex);
        }
    }

    static object HandleGetAllFieldDescriptions(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var fields = accessService.GetAllFieldDescriptions(tableName);
            return new { success = true, fields = fields.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_all_field_descriptions", ex);
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

    static object HandleGetQueryParameters(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "query_name", out var queryName, out var queryNameError))
                return queryNameError;

            var parameters = accessService.GetQueryParameters(queryName);
            return new { success = true, query_name = queryName, parameters = parameters.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_query_parameters", ex);
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

    static object HandleGetFieldAttributes(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;

            var attributes = accessService.GetFieldAttributes(tableName, fieldName);
            return new { success = true, attributes = attributes };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_field_attributes", ex);
        }
    }

    static object HandleDetectMultiValueFields(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var fields = accessService.DetectMultiValueFields(tableName);
            return new { success = true, fields = fields.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("detect_multi_value_fields", ex);
        }
    }

    static object HandleGetMultiValueFieldValues(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetOptionalInt(arguments, "max_rows", out var maxRows, out var maxRowsError))
                return maxRowsError;

            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            var values = accessService.GetMultiValueFieldValues(
                tableName,
                fieldName,
                string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition,
                maxRows ?? 100);

            return new { success = true, values = values.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_multi_value_field_values", ex);
        }
    }

    static object HandleSetMultiValueFieldValues(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredPrimitiveArray(arguments, "values", out var values, out var valuesError))
                return valuesError;

            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            var result = accessService.SetMultiValueFieldValues(
                tableName,
                fieldName,
                values,
                string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition);

            return new { success = true, result = result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_multi_value_field_values", ex);
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

    static object HandleGetVbaReferences(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            var references = accessService.GetVbaReferences(string.IsNullOrWhiteSpace(projectName) ? null : projectName);
            return new { success = true, references = references.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_vba_references", ex);
        }
    }

    static object HandleAddVbaReference(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            _ = TryGetOptionalString(arguments, "reference_path", out var referencePath);
            _ = TryGetOptionalString(arguments, "reference_guid", out var referenceGuid);

            if (string.IsNullOrWhiteSpace(referencePath) && string.IsNullOrWhiteSpace(referenceGuid))
                return new { success = false, error = "reference_path or reference_guid is required" };

            if (!TryGetOptionalInt(arguments, "major", out var major, out var majorError))
                return majorError;
            if (!TryGetOptionalInt(arguments, "minor", out var minor, out var minorError))
                return minorError;

            accessService.AddVbaReference(
                string.IsNullOrWhiteSpace(projectName) ? null : projectName,
                string.IsNullOrWhiteSpace(referencePath) ? null : referencePath,
                string.IsNullOrWhiteSpace(referenceGuid) ? null : referenceGuid,
                major ?? 1,
                minor ?? 0);

            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("add_vba_reference", ex);
        }
    }

    static object HandleRemoveVbaReference(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            if (!TryGetRequiredString(arguments, "reference_identifier", out var referenceIdentifier, out var referenceIdentifierError))
                return referenceIdentifierError;

            accessService.RemoveVbaReference(
                string.IsNullOrWhiteSpace(projectName) ? null : projectName,
                referenceIdentifier);

            return new { success = true, reference_identifier = referenceIdentifier };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("remove_vba_reference", ex);
        }
    }

    static object HandleGetStartupProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var properties = accessService.GetStartupProperties();
            return new { success = true, properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_startup_properties", ex);
        }
    }

    static object HandleSetStartupProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var hasStartupForm = TryGetOptionalString(arguments, "startup_form", out var startupForm);
            var hasAppTitle = TryGetOptionalString(arguments, "app_title", out var appTitle);
            var hasAppIcon = TryGetOptionalString(arguments, "app_icon", out var appIcon);

            if (!hasStartupForm && !hasAppTitle && !hasAppIcon)
                return new { success = false, error = "At least one startup property is required" };

            accessService.SetStartupProperties(
                hasStartupForm ? startupForm : null,
                hasAppTitle ? appTitle : null,
                hasAppIcon ? appIcon : null);

            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_startup_properties", ex);
        }
    }

    static object HandleGetRibbonXml(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "ribbon_name", out var ribbonName);
            var ribbon = accessService.GetRibbonXml(string.IsNullOrWhiteSpace(ribbonName) ? null : ribbonName);
            return new { success = true, ribbon };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_ribbon_xml", ex);
        }
    }

    static object HandleSetRibbonXml(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "ribbon_name", out var ribbonName, out var ribbonNameError))
                return ribbonNameError;
            if (!TryGetRequiredString(arguments, "ribbon_xml", out var ribbonXml, out var ribbonXmlError))
                return ribbonXmlError;

            var applyAsDefault = GetOptionalBool(arguments, "apply_as_default", false);
            accessService.SetRibbonXml(ribbonName, ribbonXml, applyAsDefault);
            return new { success = true, ribbon_name = ribbonName, apply_as_default = applyAsDefault };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_ribbon_xml", ex);
        }
    }

    static object HandleGetApplicationInfo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var info = accessService.GetApplicationInfo();
            return new { success = true, application = info };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_application_info", ex);
        }
    }

    static object HandleGetCurrentProjectData(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var info = accessService.GetCurrentProjectData();
            return new { success = true, data = info };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_current_project_data", ex);
        }
    }

    static object HandleGetApplicationOption(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "option_name", out var optionName, out var optionNameError))
                return optionNameError;

            var value = accessService.GetApplicationOption(optionName);
            return new { success = true, option_name = optionName, value };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_application_option", ex);
        }
    }

    static object HandleSetApplicationOption(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "option_name", out var optionName, out var optionNameError))
                return optionNameError;
            if (!TryGetRequiredString(arguments, "value", out var value, out var valueError))
                return valueError;

            accessService.SetApplicationOption(optionName, value);
            return new { success = true, option_name = optionName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_application_option", ex);
        }
    }

    static object HandleGetTempVars(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var tempVars = accessService.GetTempVars();
            return new { success = true, temp_vars = tempVars.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_temp_vars", ex);
        }
    }

    static object HandleSetTempVar(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "name", out var name, out var nameError))
                return nameError;
            if (!TryGetOptionalPrimitiveValue(arguments, "value", out var value, out var valueError))
                return valueError;

            accessService.SetTempVar(name, value);
            return new { success = true, name = name };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_temp_var", ex);
        }
    }

    static object HandleRemoveTempVar(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "name", out var name, out var nameError))
                return nameError;

            accessService.RemoveTempVar(name);
            return new { success = true, name = name };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("remove_temp_var", ex);
        }
    }

    static object HandleClearTempVars(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.ClearTempVars();
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("clear_temp_vars", ex);
        }
    }

    static object HandleExportDataMacroAxl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var axlXml = accessService.ExportDataMacroAxl(tableName);
            return new { success = true, table_name = tableName, axl_xml = axlXml };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("export_data_macro_axl", ex);
        }
    }

    static object HandleImportDataMacroAxl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "axl_xml", out var axlXml, out var axlXmlError))
                return axlXmlError;

            accessService.ImportDataMacroAxl(tableName, axlXml);
            return new { success = true, table_name = tableName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("import_data_macro_axl", ex);
        }
    }

    static object HandleRunDataMacro(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;

            accessService.RunDataMacro(macroName);
            return new { success = true, macro_name = macroName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("run_data_macro", ex);
        }
    }

    static object HandleGetTableDataMacros(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;

            var macros = accessService.GetTableDataMacros(tableName);
            return new { success = true, table_name = tableName, macros = macros.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_table_data_macros", ex);
        }
    }

    static object HandleDeleteDataMacro(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "macro_name", out var macroName, out var macroNameError))
                return macroNameError;

            accessService.DeleteDataMacro(tableName, macroName);
            return new { success = true, table_name = tableName, macro_name = macroName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_data_macro", ex);
        }
    }

    static object HandleGetAutoExecInfo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var info = accessService.GetAutoExecInfo();
            return new { success = true, autoexec = info };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_autoexec_info", ex);
        }
    }

    static object HandleRunAutoExec(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.RunAutoExec();
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("run_autoexec", ex);
        }
    }

    static object HandleGetDatabaseSecurity(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var security = accessService.GetDatabaseSecurityInfo();
            return new { success = true, security };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_database_security", ex);
        }
    }

    static object HandleSetDatabasePassword(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "new_password", out var newPassword, out var newPasswordError))
                return newPasswordError;

            accessService.SetDatabasePassword(newPassword);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_database_password", ex);
        }
    }

    static object HandleRemoveDatabasePassword(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.RemoveDatabasePassword();
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("remove_database_password", ex);
        }
    }

    static object HandleEncryptDatabase(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "password", out var password);
            accessService.EncryptDatabase(string.IsNullOrWhiteSpace(password) ? null : password);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("encrypt_database", ex);
        }
    }

    static object HandleGetNavigationGroups(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var groups = accessService.GetNavigationGroups();
            return new { success = true, groups = groups.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_navigation_groups", ex);
        }
    }

    static object HandleSetDisplayCategories(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var showCategories = GetOptionalBool(arguments, "show_categories", true);
            accessService.SetDisplayCategories(showCategories);
            return new { success = true, show_categories = showCategories };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_display_categories", ex);
        }
    }

    static object HandleRefreshDatabaseWindow(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.RefreshDatabaseWindow();
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("refresh_database_window", ex);
        }
    }

    static object HandleCreateNavigationGroup(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "group_name", out var groupName, out var groupNameError))
                return groupNameError;

            accessService.CreateNavigationGroup(groupName);
            return new { success = true, group_name = groupName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("create_navigation_group", ex);
        }
    }

    static object HandleAddNavigationGroupObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "group_name", out var groupName, out var groupNameError))
                return groupNameError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            _ = TryGetOptionalString(arguments, "object_type", out var objectType);
            accessService.AddNavigationGroupObject(groupName, objectName, string.IsNullOrWhiteSpace(objectType) ? null : objectType);
            return new { success = true, group_name = groupName, object_name = objectName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("add_navigation_group_object", ex);
        }
    }

    static object HandleDeleteNavigationGroup(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "group_name", out var groupName, out var groupNameError))
                return groupNameError;

            accessService.DeleteNavigationGroup(groupName);
            return new { success = true, group_name = groupName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_navigation_group", ex);
        }
    }

    static object HandleRemoveNavigationGroupObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "group_name", out var groupName, out var groupNameError))
                return groupNameError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            accessService.RemoveNavigationGroupObject(groupName, objectName);
            return new { success = true, group_name = groupName, object_name = objectName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("remove_navigation_group_object", ex);
        }
    }

    static object HandleSetNavigationPaneVisibility(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var visible = GetOptionalBool(arguments, "visible", true);
            accessService.SetNavigationPaneVisibility(visible);
            return new { success = true, visible = visible };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_navigation_pane_visibility", ex);
        }
    }

    static object HandleGetNavigationGroupObjects(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "group_name", out var groupName, out var groupNameError))
                return groupNameError;

            var objects = accessService.GetNavigationGroupObjects(groupName);
            return new { success = true, objects = objects.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_navigation_group_objects", ex);
        }
    }

    static object HandleGetConditionalFormatting(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;

            var rules = accessService.GetConditionalFormatting(objectType, objectName, controlName);
            return new { success = true, rules = rules.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_conditional_formatting", ex);
        }
    }

    static object HandleAddConditionalFormatting(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;
            if (!TryGetRequiredString(arguments, "expression", out var expression, out var expressionError))
                return expressionError;

            if (!TryGetOptionalInt(arguments, "fore_color", out var foreColor, out var foreColorError))
                return foreColorError;
            if (!TryGetOptionalInt(arguments, "back_color", out var backColor, out var backColorError))
                return backColorError;

            accessService.AddConditionalFormatting(objectType, objectName, controlName, expression, foreColor, backColor);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("add_conditional_formatting", ex);
        }
    }

    static object HandleDeleteConditionalFormatting(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;
            if (!TryGetOptionalInt(arguments, "rule_index", out var ruleIndex, out var ruleIndexError))
                return ruleIndexError;
            if (!ruleIndex.HasValue || ruleIndex.Value <= 0)
                return new { success = false, error = "rule_index must be an integer greater than 0" };

            accessService.DeleteConditionalFormatting(objectType, objectName, controlName, ruleIndex.Value);
            return new { success = true, rule_index = ruleIndex.Value };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_conditional_formatting", ex);
        }
    }

    static object HandleUpdateConditionalFormatting(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;
            if (!TryGetOptionalInt(arguments, "rule_index", out var ruleIndex, out var ruleIndexError))
                return ruleIndexError;
            if (!ruleIndex.HasValue || ruleIndex.Value <= 0)
                return new { success = false, error = "rule_index must be an integer greater than 0" };

            _ = TryGetOptionalString(arguments, "expression", out var expression);
            if (!TryGetOptionalInt(arguments, "fore_color", out var foreColor, out var foreColorError))
                return foreColorError;
            if (!TryGetOptionalInt(arguments, "back_color", out var backColor, out var backColorError))
                return backColorError;
            if (!TryGetOptionalBoolNullable(arguments, "enabled", out var enabled, out var enabledError))
                return enabledError;

            var updatedRule = accessService.UpdateConditionalFormatting(
                objectType,
                objectName,
                controlName,
                ruleIndex.Value,
                expression: string.IsNullOrWhiteSpace(expression) ? null : expression,
                foreColor: foreColor,
                backColor: backColor,
                enabled: enabled);

            return new { success = true, rule = updatedRule };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("update_conditional_formatting", ex);
        }
    }

    static object HandleClearConditionalFormatting(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;

            accessService.ClearConditionalFormatting(objectType, objectName, controlName);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("clear_conditional_formatting", ex);
        }
    }

    static object HandleListAllConditionalFormats(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            var controls = accessService.ListAllConditionalFormats(objectType, objectName);
            return new { success = true, controls = controls.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("list_all_conditional_formats", ex);
        }
    }

    static object HandleGetAttachmentFiles(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;

            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            var files = accessService.GetAttachmentFieldFiles(tableName, fieldName, string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition);
            return new { success = true, files = files.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_attachment_files", ex);
        }
    }

    static object HandleAddAttachmentFile(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "file_path", out var filePath, out var filePathError))
                return filePathError;

            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            accessService.AddAttachmentFile(tableName, fieldName, filePath, string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("add_attachment_file", ex);
        }
    }

    static object HandleRemoveAttachmentFile(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "file_name", out var fileName, out var fileNameError))
                return fileNameError;

            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            accessService.RemoveAttachmentFile(tableName, fieldName, fileName, string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("remove_attachment_file", ex);
        }
    }

    static object HandleSaveAttachmentToDisk(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;
            if (!TryGetRequiredString(arguments, "file_path", out var filePath, out var filePathError))
                return filePathError;

            _ = TryGetOptionalString(arguments, "file_name", out var fileName);
            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            var saveResult = accessService.SaveAttachmentToDisk(
                tableName,
                fieldName,
                filePath,
                string.IsNullOrWhiteSpace(fileName) ? null : fileName,
                string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition);

            return new { success = true, result = saveResult };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("save_attachment_to_disk", ex);
        }
    }

    static object HandleGetAttachmentMetadata(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "table_name", out var tableName, out var tableNameError))
                return tableNameError;
            if (!TryGetRequiredString(arguments, "field_name", out var fieldName, out var fieldNameError))
                return fieldNameError;

            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);
            var files = accessService.GetAttachmentMetadata(
                tableName,
                fieldName,
                string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition);

            return new { success = true, files = files.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_attachment_metadata", ex);
        }
    }

    static object HandleGetObjectEvents(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            var events = accessService.GetObjectEvents(objectType, objectName);
            return new { success = true, events = events.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_object_events", ex);
        }
    }

    static object HandleSetObjectEvent(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;
            if (!TryGetRequiredString(arguments, "event_name", out var eventName, out var eventNameError))
                return eventNameError;
            if (!TryGetRequiredString(arguments, "event_value", out var eventValue, out var eventValueError))
                return eventValueError;

            accessService.SetObjectEvent(objectType, objectName, eventName, eventValue);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_object_event", ex);
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

    static object HandleGetFormRecordCount(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;

            var recordCount = accessService.GetFormRecordCount(formName);
            return new { success = true, form_name = formName, record_count = recordCount };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_form_record_count", ex);
        }
    }

    static object HandleGetFormCurrentRecord(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;

            var record = accessService.GetFormCurrentRecord(formName);
            return new { success = true, record = record };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_form_current_record", ex);
        }
    }

    static object HandleSetFormFilter(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;

            _ = TryGetOptionalString(arguments, "filter", out var filter);
            if (!TryGetOptionalBoolNullable(arguments, "filter_on", out var filterOn, out var filterOnError))
                return filterOnError;

            accessService.SetFormFilter(
                formName,
                string.IsNullOrWhiteSpace(filter) ? null : filter,
                filterOn);

            return new { success = true, form_name = formName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_form_filter", ex);
        }
    }

    static object HandleGetOpenObjects(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var openObjects = accessService.GetOpenObjects();
            return new { success = true, open_objects = openObjects.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_open_objects", ex);
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

    static object HandleExecuteVba(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "expression", out var expression, out var expressionError))
                return expressionError;

            var result = accessService.ExecuteVba(expression);
            return new { success = true, result = result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("execute_vba", ex);
        }
    }

    static object HandleRunVbaProcedure(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "procedure_name", out var procedureName, out var procedureNameError))
                return procedureNameError;

            if (!TryGetOptionalPrimitiveArray(arguments, "args", out var args, out var argsError))
                return argsError;

            var result = accessService.RunVbaProcedure(procedureName, args);
            return new { success = true, result = result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("run_vba_procedure", ex);
        }
    }

    static object HandleCreateModule(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            accessService.CreateModule(moduleName, string.IsNullOrWhiteSpace(projectName) ? null : projectName);
            return new { success = true, module_name = moduleName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("create_module", ex);
        }
    }

    static object HandleDeleteModule(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            accessService.DeleteModule(moduleName, string.IsNullOrWhiteSpace(projectName) ? null : projectName);
            return new { success = true, module_name = moduleName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_module", ex);
        }
    }

    static object HandleRenameModule(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;
            if (!TryGetRequiredString(arguments, "new_module_name", out var newModuleName, out var newModuleNameError))
                return newModuleNameError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            accessService.RenameModule(moduleName, newModuleName, string.IsNullOrWhiteSpace(projectName) ? null : projectName);
            return new { success = true, module_name = moduleName, new_module_name = newModuleName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("rename_module", ex);
        }
    }

    static object HandleGetCompilationErrors(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var compilation = accessService.GetCompilationErrors();
            return new { success = true, compilation = compilation };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_compilation_errors", ex);
        }
    }

    static object HandleListAllProcedures(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            var procedures = accessService.ListAllProcedures(string.IsNullOrWhiteSpace(projectName) ? null : projectName);
            return new { success = true, procedures = procedures.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("list_all_procedures", ex);
        }
    }

    static object HandleGetVbaProjectProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            var properties = accessService.GetVbaProjectProperties(string.IsNullOrWhiteSpace(projectName) ? null : projectName);
            return new { success = true, properties = properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_vba_project_properties", ex);
        }
    }

    static object HandleSetVbaProjectProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            _ = TryGetOptionalString(arguments, "name", out var name);
            _ = TryGetOptionalString(arguments, "description", out var description);
            _ = TryGetOptionalString(arguments, "help_file", out var helpFile);
            if (!TryGetOptionalInt(arguments, "help_context_id", out var helpContextId, out var helpContextIdError))
                return helpContextIdError;

            var hasName = !string.IsNullOrWhiteSpace(name);
            var hasDescription = !string.IsNullOrWhiteSpace(description);
            var hasHelpFile = !string.IsNullOrWhiteSpace(helpFile);
            if (!hasName && !hasDescription && !hasHelpFile && !helpContextId.HasValue)
                return new { success = false, error = "At least one of name, description, help_file, or help_context_id is required" };

            var properties = accessService.SetVbaProjectProperties(
                string.IsNullOrWhiteSpace(projectName) ? null : projectName,
                hasName ? name : null,
                hasDescription ? description : null,
                hasHelpFile ? helpFile : null,
                helpContextId);

            return new { success = true, properties = properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_vba_project_properties", ex);
        }
    }

    static object HandleGetModuleInfo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            var moduleInfo = accessService.GetModuleInfo(string.IsNullOrWhiteSpace(projectName) ? null : projectName, moduleName);
            return new { success = true, module_info = moduleInfo };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_module_info", ex);
        }
    }

    static object HandleListProcedures(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            var procedures = accessService.ListProcedures(string.IsNullOrWhiteSpace(projectName) ? null : projectName, moduleName);
            return new { success = true, procedures = procedures.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("list_procedures", ex);
        }
    }

    static object HandleGetProcedureCode(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;
            if (!TryGetRequiredString(arguments, "procedure_name", out var procedureName, out var procedureNameError))
                return procedureNameError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            var code = accessService.GetProcedureCode(string.IsNullOrWhiteSpace(projectName) ? null : projectName, moduleName, procedureName);
            return new { success = true, procedure_code = code };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_procedure_code", ex);
        }
    }

    static object HandleGetModuleDeclarations(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            var declarations = accessService.GetModuleDeclarations(string.IsNullOrWhiteSpace(projectName) ? null : projectName, moduleName);
            return new { success = true, declarations };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_module_declarations", ex);
        }
    }

    static object HandleInsertLines(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;
            if (!TryGetRequiredString(arguments, "code", out var code, out var codeError))
                return codeError;
            if (!TryGetOptionalInt(arguments, "line_number", out var lineNumber, out var lineNumberError))
                return lineNumberError;

            if (!lineNumber.HasValue)
                return new { success = false, error = "line_number is required" };

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            accessService.InsertLines(string.IsNullOrWhiteSpace(projectName) ? null : projectName, moduleName, lineNumber.Value, code);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("insert_lines", ex);
        }
    }

    static object HandleDeleteLines(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;
            if (!TryGetOptionalInt(arguments, "start_line", out var startLine, out var startLineError))
                return startLineError;
            if (!TryGetOptionalInt(arguments, "line_count", out var lineCount, out var lineCountError))
                return lineCountError;

            if (!startLine.HasValue)
                return new { success = false, error = "start_line is required" };

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            accessService.DeleteLines(string.IsNullOrWhiteSpace(projectName) ? null : projectName, moduleName, startLine.Value, lineCount ?? 1);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_lines", ex);
        }
    }

    static object HandleReplaceLine(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;
            if (!TryGetRequiredString(arguments, "code", out var code, out var codeError))
                return codeError;
            if (!TryGetOptionalInt(arguments, "line_number", out var lineNumber, out var lineNumberError))
                return lineNumberError;

            if (!lineNumber.HasValue)
                return new { success = false, error = "line_number is required" };

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);
            accessService.ReplaceLine(string.IsNullOrWhiteSpace(projectName) ? null : projectName, moduleName, lineNumber.Value, code);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("replace_line", ex);
        }
    }

    static object HandleFindTextInModule(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "module_name", out var moduleName, out var moduleNameError))
                return moduleNameError;
            if (!TryGetRequiredString(arguments, "find_text", out var findText, out var findTextError))
                return findTextError;

            _ = TryGetOptionalString(arguments, "project_name", out var projectName);

            if (!TryGetOptionalInt(arguments, "start_line", out var startLine, out var startLineError))
                return startLineError;
            if (!TryGetOptionalInt(arguments, "start_column", out var startColumn, out var startColumnError))
                return startColumnError;
            if (!TryGetOptionalInt(arguments, "end_line", out var endLine, out var endLineError))
                return endLineError;
            if (!TryGetOptionalInt(arguments, "end_column", out var endColumn, out var endColumnError))
                return endColumnError;

            var wholeWord = GetOptionalBool(arguments, "whole_word", false);
            var matchCase = GetOptionalBool(arguments, "match_case", false);
            var patternSearch = GetOptionalBool(arguments, "pattern_search", false);

            var result = accessService.FindTextInModule(
                string.IsNullOrWhiteSpace(projectName) ? null : projectName,
                moduleName,
                findText,
                startLine ?? 1,
                startColumn ?? 1,
                endLine,
                endColumn,
                wholeWord,
                matchCase,
                patternSearch);

            return new { success = true, result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("find_text_in_module", ex);
        }
    }

    static object HandleListImportExportSpecs(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var specs = accessService.ListImportExportSpecs();
            return new { success = true, specifications = specs.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("list_import_export_specs", ex);
        }
    }

    static object HandleGetImportExportSpec(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "specification_name", out var specificationName, out var specificationNameError))
                return specificationNameError;

            var specification = accessService.GetImportExportSpec(specificationName);
            return new { success = true, specification };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_import_export_spec", ex);
        }
    }

    static object HandleCreateImportExportSpec(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "specification_name", out var specificationName, out var specificationNameError))
                return specificationNameError;
            if (!TryGetRequiredString(arguments, "specification_xml", out var specificationXml, out var specificationXmlError))
                return specificationXmlError;

            accessService.CreateImportExportSpec(specificationName, specificationXml);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("create_import_export_spec", ex);
        }
    }

    static object HandleDeleteImportExportSpec(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "specification_name", out var specificationName, out var specificationNameError))
                return specificationNameError;

            accessService.DeleteImportExportSpec(specificationName);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_import_export_spec", ex);
        }
    }

    static object HandleRunImportExportSpec(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "specification_name", out var specificationName, out var specificationNameError))
                return specificationNameError;

            accessService.RunImportExportSpec(specificationName);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("run_import_export_spec", ex);
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

    static object HandleGetContainers(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var containers = accessService.GetContainers();
            return new { success = true, containers = containers.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_containers", ex);
        }
    }

    static object HandleGetContainerDocuments(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "container_name", out var containerName, out var containerNameError))
                return containerNameError;

            var documents = accessService.GetContainerDocuments(containerName);
            return new { success = true, container_name = containerName, documents = documents.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_container_documents", ex);
        }
    }

    static object HandleGetDocumentProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "container_name", out var containerName, out var containerNameError))
                return containerNameError;
            if (!TryGetRequiredString(arguments, "document_name", out var documentName, out var documentNameError))
                return documentNameError;

            var properties = accessService.GetDocumentProperties(containerName, documentName);
            return new { success = true, container_name = containerName, document_name = documentName, properties = properties.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_document_properties", ex);
        }
    }

    static object HandleSetDocumentProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "container_name", out var containerName, out var containerNameError))
                return containerNameError;
            if (!TryGetRequiredString(arguments, "document_name", out var documentName, out var documentNameError))
                return documentNameError;
            if (!TryGetRequiredString(arguments, "property_name", out var propertyName, out var propertyNameError))
                return propertyNameError;
            if (!TryGetRequiredString(arguments, "value", out var value, out var valueError))
                return valueError;

            _ = TryGetOptionalString(arguments, "property_type", out var propertyType);
            var createIfMissing = GetOptionalBool(arguments, "create_if_missing", false);
            var updated = accessService.SetDocumentProperty(
                containerName,
                documentName,
                propertyName,
                value,
                string.IsNullOrWhiteSpace(propertyType) ? null : propertyType,
                createIfMissing);

            return new { success = true, property = updated };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_document_property", ex);
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

    static object HandleGetFormSections(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;

            var sections = accessService.GetFormSections(formName);
            return new { success = true, sections = sections.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_form_sections", ex);
        }
    }

    static object HandleGetReportSections(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;

            var sections = accessService.GetReportSections(reportName);
            return new { success = true, sections = sections.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_report_sections", ex);
        }
    }

    static object HandleSetSectionProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;
            if (!TryGetRequiredString(arguments, "section", out var section, out var sectionError))
                return sectionError;
            if (!TryGetRequiredString(arguments, "property_name", out var propertyName, out var propertyNameError))
                return propertyNameError;
            if (!TryGetRequiredString(arguments, "value", out var value, out var valueError))
                return valueError;

            accessService.SetSectionProperty(objectType, objectName, section, propertyName, value);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_section_property", ex);
        }
    }

    static object HandleCreateControl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;
            if (!TryGetRequiredString(arguments, "control_type", out var controlType, out var controlTypeError))
                return controlTypeError;

            _ = TryGetOptionalString(arguments, "control_name", out var controlName);
            _ = TryGetOptionalString(arguments, "parent_control_name", out var parentControlName);
            _ = TryGetOptionalString(arguments, "column_name", out var columnName);

            if (!TryGetOptionalInt(arguments, "section", out var section, out var sectionError))
                return sectionError;
            if (!TryGetOptionalInt(arguments, "left", out var left, out var leftError))
                return leftError;
            if (!TryGetOptionalInt(arguments, "top", out var top, out var topError))
                return topError;
            if (!TryGetOptionalInt(arguments, "width", out var width, out var widthError))
                return widthError;
            if (!TryGetOptionalInt(arguments, "height", out var height, out var heightError))
                return heightError;

            var control = accessService.CreateControl(
                formName,
                controlType,
                string.IsNullOrWhiteSpace(controlName) ? null : controlName,
                section ?? 0,
                string.IsNullOrWhiteSpace(parentControlName) ? null : parentControlName,
                string.IsNullOrWhiteSpace(columnName) ? null : columnName,
                left,
                top,
                width,
                height);

            return new { success = true, control };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("create_control", ex);
        }
    }

    static object HandleCreateReportControl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;
            if (!TryGetRequiredString(arguments, "control_type", out var controlType, out var controlTypeError))
                return controlTypeError;

            _ = TryGetOptionalString(arguments, "control_name", out var controlName);
            _ = TryGetOptionalString(arguments, "parent_control_name", out var parentControlName);
            _ = TryGetOptionalString(arguments, "column_name", out var columnName);

            if (!TryGetOptionalInt(arguments, "section", out var section, out var sectionError))
                return sectionError;
            if (!TryGetOptionalInt(arguments, "left", out var left, out var leftError))
                return leftError;
            if (!TryGetOptionalInt(arguments, "top", out var top, out var topError))
                return topError;
            if (!TryGetOptionalInt(arguments, "width", out var width, out var widthError))
                return widthError;
            if (!TryGetOptionalInt(arguments, "height", out var height, out var heightError))
                return heightError;

            var control = accessService.CreateReportControl(
                reportName,
                controlType,
                string.IsNullOrWhiteSpace(controlName) ? null : controlName,
                section ?? 0,
                string.IsNullOrWhiteSpace(parentControlName) ? null : parentControlName,
                string.IsNullOrWhiteSpace(columnName) ? null : columnName,
                left,
                top,
                width,
                height);

            return new { success = true, control };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("create_report_control", ex);
        }
    }

    static object HandleDeleteControl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;

            accessService.DeleteControl(formName, controlName);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_control", ex);
        }
    }

    static object HandleDeleteReportControl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlNameError))
                return controlNameError;

            accessService.DeleteReportControl(reportName, controlName);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_report_control", ex);
        }
    }

    static object HandleGetFormProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;

            var properties = accessService.GetFormProperties(formName);
            return new { success = true, properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_form_properties", ex);
        }
    }

    static object HandleSetFormProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;
            if (!TryGetRequiredString(arguments, "property_name", out var propertyName, out var propertyNameError))
                return propertyNameError;
            if (!TryGetRequiredString(arguments, "value", out var value, out var valueError))
                return valueError;

            accessService.SetFormProperty(formName, propertyName, value);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_form_property", ex);
        }
    }

    static object HandleSetFormRecordSource(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;
            if (!TryGetRequiredString(arguments, "record_source", out var recordSource, out var recordSourceError))
                return recordSourceError;

            accessService.SetFormRecordSource(formName, recordSource);
            return new { success = true, form_name = formName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_form_record_source", ex);
        }
    }

    static object HandleGetReportProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;

            var properties = accessService.GetReportProperties(reportName);
            return new { success = true, properties };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_report_properties", ex);
        }
    }

    static object HandleSetReportProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;
            if (!TryGetRequiredString(arguments, "property_name", out var propertyName, out var propertyNameError))
                return propertyNameError;
            if (!TryGetRequiredString(arguments, "value", out var value, out var valueError))
                return valueError;

            accessService.SetReportProperty(reportName, propertyName, value);
            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_report_property", ex);
        }
    }

    static object HandleSetReportRecordSource(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;
            if (!TryGetRequiredString(arguments, "record_source", out var recordSource, out var recordSourceError))
                return recordSourceError;

            accessService.SetReportRecordSource(reportName, recordSource);
            return new { success = true, report_name = reportName };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_report_record_source", ex);
        }
    }

    static object HandleGetReportGrouping(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;

            var grouping = accessService.GetReportGrouping(reportName);
            return new { success = true, grouping = grouping.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_report_grouping", ex);
        }
    }

    static object HandleSetReportGrouping(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;

            _ = TryGetOptionalString(arguments, "expression", out var expression);
            if (!TryGetOptionalInt(arguments, "index", out var index, out var indexError))
                return indexError;
            if (!TryGetOptionalInt(arguments, "sort_order", out var sortOrder, out var sortOrderError))
                return sortOrderError;
            if (!TryGetOptionalInt(arguments, "group_on", out var groupOn, out var groupOnError))
                return groupOnError;
            if (!TryGetOptionalInt(arguments, "group_interval", out var groupInterval, out var groupIntervalError))
                return groupIntervalError;
            if (!TryGetOptionalBoolNullable(arguments, "group_header", out var groupHeader, out var groupHeaderError))
                return groupHeaderError;
            if (!TryGetOptionalBoolNullable(arguments, "group_footer", out var groupFooter, out var groupFooterError))
                return groupFooterError;
            if (!TryGetOptionalInt(arguments, "keep_together", out var keepTogether, out var keepTogetherError))
                return keepTogetherError;

            var grouping = accessService.SetReportGrouping(
                reportName,
                expression: string.IsNullOrWhiteSpace(expression) ? null : expression,
                index: index,
                sortOrder: sortOrder,
                groupOn: groupOn,
                groupInterval: groupInterval,
                groupHeader: groupHeader,
                groupFooter: groupFooter,
                keepTogether: keepTogether);

            return new { success = true, grouping = grouping };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_report_grouping", ex);
        }
    }

    static object HandleDeleteReportGrouping(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;
            if (!TryGetOptionalInt(arguments, "index", out var index, out var indexError))
                return indexError;
            if (!index.HasValue || index.Value < 0)
                return new { success = false, error = "index must be an integer greater than or equal to 0" };

            accessService.DeleteReportGrouping(reportName, index.Value);
            return new { success = true, report_name = reportName, index = index.Value };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("delete_report_grouping", ex);
        }
    }

    static object HandleGetReportSorting(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "report_name", out var reportName, out var reportNameError))
                return reportNameError;

            var sorting = accessService.GetReportSorting(reportName);
            return new { success = true, sorting = sorting };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_report_sorting", ex);
        }
    }

    static object HandleGetTabOrder(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;

            var tabOrder = accessService.GetTabOrder(formName);
            return new { success = true, tab_order = tabOrder.ToArray() };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_tab_order", ex);
        }
    }

    static object HandleSetTabOrder(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "form_name", out var formName, out var formNameError))
                return formNameError;
            if (!TryGetRequiredStringArray(arguments, "control_names", out var controlNames, out var controlNamesError))
                return controlNamesError;

            accessService.SetTabOrder(formName, controlNames);
            return new { success = true, form_name = formName, control_count = controlNames.Count };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_tab_order", ex);
        }
    }

    static object HandleGetPageSetup(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            var pageSetup = accessService.GetPageSetup(objectType, objectName);
            return new { success = true, page_setup = pageSetup };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_page_setup", ex);
        }
    }

    static object HandleSetPageSetup(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "object_type", out var objectType, out var objectTypeError))
                return objectTypeError;
            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            if (!TryGetOptionalInt(arguments, "top_margin", out var topMargin, out var topMarginError))
                return topMarginError;
            if (!TryGetOptionalInt(arguments, "bottom_margin", out var bottomMargin, out var bottomMarginError))
                return bottomMarginError;
            if (!TryGetOptionalInt(arguments, "left_margin", out var leftMargin, out var leftMarginError))
                return leftMarginError;
            if (!TryGetOptionalInt(arguments, "right_margin", out var rightMargin, out var rightMarginError))
                return rightMarginError;
            if (!TryGetOptionalInt(arguments, "orientation", out var orientation, out var orientationError))
                return orientationError;
            if (!TryGetOptionalInt(arguments, "paper_size", out var paperSize, out var paperSizeError))
                return paperSizeError;
            if (!TryGetOptionalBoolNullable(arguments, "data_only", out var dataOnly, out var dataOnlyError))
                return dataOnlyError;

            accessService.SetPageSetup(
                objectType,
                objectName,
                topMargin,
                bottomMargin,
                leftMargin,
                rightMargin,
                orientation,
                paperSize,
                dataOnly);

            return new { success = true };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_page_setup", ex);
        }
    }

    static object HandleGetPrinterInfo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var printerInfo = accessService.GetPrinterInfo();
            return new { success = true, printer_info = printerInfo };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_printer_info", ex);
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

    // ===== Priority 17: DoCmd Remaining + Domain Aggregates =====

    static object HandleFindNext(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.FindNext();
            return new { success = true, message = "FindNext executed" };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("find_next", ex);
        }
    }

    static object HandleSearchForRecord(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "where_condition", out var whereCondition, out var whereError))
                return whereError;

            var objectType = GetOptionalIntFromAliases(arguments, new[] { "object_type" }, -1);
            _ = TryGetOptionalString(arguments, "object_name", out var objectName);
            _ = TryGetOptionalString(arguments, "record", out var record);

            accessService.SearchForRecord(objectType, objectName, record, whereCondition);
            return new { success = true, message = "SearchForRecord executed", where_condition = whereCondition };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("search_for_record", ex);
        }
    }

    static object HandleSetFilterDoCmd(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            _ = TryGetOptionalString(arguments, "filter_name", out var filterName);
            _ = TryGetOptionalString(arguments, "where_condition", out var whereCondition);

            accessService.SetFilterDoCmd(
                string.IsNullOrWhiteSpace(filterName) ? null : filterName,
                string.IsNullOrWhiteSpace(whereCondition) ? null : whereCondition);
            return new { success = true, message = "SetFilter applied" };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_filter_docmd", ex);
        }
    }

    static object HandleSetOrderBy(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "order_by", out var orderBy, out var orderByError))
                return orderByError;

            accessService.SetOrderBy(orderBy);
            return new { success = true, message = "SetOrderBy applied", order_by = orderBy };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_order_by", ex);
        }
    }

    static object HandleSetParameter(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "name", out var name, out var nameError))
                return nameError;
            if (!TryGetRequiredString(arguments, "expression", out var expression, out var exprError))
                return exprError;

            accessService.SetParameter(name, expression);
            return new { success = true, message = $"Parameter '{name}' set", name, expression };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_parameter", ex);
        }
    }

    static object HandleSetRuntimeProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "control_name", out var controlName, out var controlError))
                return controlError;
            if (!TryGetRequiredString(arguments, "value", out var value, out var valueError))
                return valueError;

            var propertyId = GetOptionalIntFromAliases(arguments, new[] { "property_id" }, -1);
            if (propertyId < 0)
                return new { success = false, error = "property_id is required and must be a non-negative integer" };

            accessService.SetRuntimeProperty(controlName, propertyId, value);
            return new { success = true, message = $"Property {propertyId} set on '{controlName}'", control_name = controlName, property_id = propertyId };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_runtime_property", ex);
        }
    }

    static object HandleRefreshRecord(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.RefreshRecord();
            return new { success = true, message = "RefreshRecord executed" };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("refresh_record", ex);
        }
    }

    static object HandleCloseDatabase(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.CloseDatabase();
            return new { success = true, message = "Database closed" };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("close_database", ex);
        }
    }

    static object HandleDomainAggregate(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "function", out var function, out var funcError))
                return funcError;
            if (!TryGetRequiredString(arguments, "expression", out var expression, out var exprError))
                return exprError;
            if (!TryGetRequiredString(arguments, "domain", out var domain, out var domainError))
                return domainError;
            _ = TryGetOptionalString(arguments, "criteria", out var criteria);

            var allowedFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "DLookup", "DCount", "DSum", "DAvg", "DMin", "DMax", "DFirst", "DLast" };
            if (!allowedFunctions.Contains(function))
                return new { success = false, error = $"Invalid function '{function}'. Allowed: DLookup, DCount, DSum, DAvg, DMin, DMax, DFirst, DLast" };

            var normalizedFunction = allowedFunctions.First(f => string.Equals(f, function, StringComparison.OrdinalIgnoreCase));

            var result = accessService.DomainAggregate(
                normalizedFunction,
                expression,
                domain,
                string.IsNullOrWhiteSpace(criteria) ? null : criteria);

            return new { success = true, message = $"{normalizedFunction} executed", value = result, function = normalizedFunction, expression, domain };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("domain_aggregate", ex);
        }
    }

    static object HandleAccessError(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var errorNumber = GetOptionalIntFromAliases(arguments, new[] { "error_number" }, int.MinValue);
            if (errorNumber == int.MinValue)
                return new { success = false, error = "error_number is required" };

            var description = accessService.AccessError(errorNumber);
            return new { success = true, error_number = errorNumber, description };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("access_error", ex);
        }
    }

    static object HandleBuildCriteria(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!TryGetRequiredString(arguments, "field", out var field, out var fieldError))
                return fieldError;
            if (!TryGetRequiredString(arguments, "expression", out var expression, out var exprError))
                return exprError;

            var fieldType = GetOptionalIntFromAliases(arguments, new[] { "field_type" }, int.MinValue);
            if (fieldType == int.MinValue)
                return new { success = false, error = "field_type is required" };

            var result = accessService.BuildCriteria(field, fieldType, expression);
            return new { success = true, criteria = result, field, field_type = fieldType, expression };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("build_criteria", ex);
        }
    }

    // ===== Priority 18: Screen Object + Visibility + App Info =====

    static object HandleGetActiveForm(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var result = accessService.GetActiveForm();
            return new { success = true, result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_active_form", ex);
        }
    }

    static object HandleGetActiveReport(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var result = accessService.GetActiveReport();
            return new { success = true, result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_active_report", ex);
        }
    }

    static object HandleGetActiveControl(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var result = accessService.GetActiveControl();
            return new { success = true, result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_active_control", ex);
        }
    }

    static object HandleGetActiveDatasheet(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var result = accessService.GetActiveDatasheet();
            return new { success = true, result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_active_datasheet", ex);
        }
    }

    static object HandleSetHiddenAttribute(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!arguments.TryGetProperty("object_type", out var objectTypeElement) ||
                objectTypeElement.ValueKind != JsonValueKind.Number ||
                !objectTypeElement.TryGetInt32(out var objectType))
                return new { success = false, error = "object_type is required (integer: 0=Table, 1=Query, 2=Form, 3=Report, 4=Macro, 5=Module)" };

            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            var hidden = GetOptionalBool(arguments, "hidden", true);

            accessService.SetHiddenAttribute(objectType, objectName, hidden);
            return new { success = true, message = $"Set hidden attribute for {objectName} to {hidden}", object_type = objectType, object_name = objectName, hidden };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_hidden_attribute", ex);
        }
    }

    static object HandleGetHiddenAttribute(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!arguments.TryGetProperty("object_type", out var objectTypeElement) ||
                objectTypeElement.ValueKind != JsonValueKind.Number ||
                !objectTypeElement.TryGetInt32(out var objectType))
                return new { success = false, error = "object_type is required (integer: 0=Table, 1=Query, 2=Form, 3=Report, 4=Macro, 5=Module)" };

            if (!TryGetRequiredString(arguments, "object_name", out var objectName, out var objectNameError))
                return objectNameError;

            var hidden = accessService.GetHiddenAttribute(objectType, objectName);
            return new { success = true, object_type = objectType, object_name = objectName, hidden };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_hidden_attribute", ex);
        }
    }

    static object HandleGetCurrentObject(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var result = accessService.GetCurrentObject();
            return new { success = true, result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_current_object", ex);
        }
    }

    static object HandleGetCurrentUser(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var result = accessService.GetCurrentUser();
            return new { success = true, user = result };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_current_user", ex);
        }
    }

    static object HandleSetAccessVisible(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var visible = GetOptionalBool(arguments, "visible", true);
            accessService.SetAccessVisible(visible);
            return new { success = true, message = $"Access application visibility set to {visible}", visible };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("set_access_visible", ex);
        }
    }

    static object HandleGetAccessHwnd(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var hwnd = accessService.GetAccessHwnd();
            return new { success = true, hwnd };
        }
        catch (Exception ex)
        {
            return BuildOperationErrorResponse("get_access_hwnd", ex);
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

    static bool TryGetRequiredPrimitiveArray(JsonElement arguments, string propertyName, out List<object?> values, out object error)
    {
        values = new List<object?>();

        if (!arguments.TryGetProperty(propertyName, out var element) || element.ValueKind != JsonValueKind.Array)
        {
            error = new { success = false, error = $"{propertyName} is required" };
            return false;
        }

        foreach (var item in element.EnumerateArray())
        {
            switch (item.ValueKind)
            {
                case JsonValueKind.Null:
                    values.Add(null);
                    break;
                case JsonValueKind.String:
                    values.Add(item.GetString());
                    break;
                case JsonValueKind.True:
                case JsonValueKind.False:
                    values.Add(item.GetBoolean());
                    break;
                case JsonValueKind.Number:
                    if (item.TryGetInt64(out var intValue))
                    {
                        values.Add(intValue);
                    }
                    else if (item.TryGetDouble(out var doubleValue))
                    {
                        values.Add(doubleValue);
                    }
                    else
                    {
                        values.Add(item.GetRawText());
                    }
                    break;
                default:
                    error = new { success = false, error = $"{propertyName} must be an array of primitive JSON values" };
                    return false;
            }
        }

        error = new { success = true };
        return true;
    }

    static bool TryGetOptionalPrimitiveArray(JsonElement arguments, string propertyName, out List<object?> values, out object error)
    {
        values = new List<object?>();
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

        if (element.ValueKind != JsonValueKind.Array)
        {
            error = new { success = false, error = $"{propertyName} must be an array when provided" };
            return false;
        }

        foreach (var item in element.EnumerateArray())
        {
            switch (item.ValueKind)
            {
                case JsonValueKind.Null:
                    values.Add(null);
                    break;
                case JsonValueKind.String:
                    values.Add(item.GetString());
                    break;
                case JsonValueKind.True:
                case JsonValueKind.False:
                    values.Add(item.GetBoolean());
                    break;
                case JsonValueKind.Number:
                    if (item.TryGetInt64(out var intValue))
                        values.Add(intValue);
                    else if (item.TryGetDouble(out var doubleValue))
                        values.Add(doubleValue);
                    else
                        values.Add(item.GetRawText());
                    break;
                default:
                    error = new { success = false, error = $"{propertyName} must contain only primitive JSON values" };
                    return false;
            }
        }

        error = new { success = true };
        return true;
    }

    static bool TryGetOptionalPrimitiveValue(JsonElement arguments, string propertyName, out object? value, out object error)
    {
        value = null;
        if (!arguments.TryGetProperty(propertyName, out var element))
        {
            error = new { success = true };
            return true;
        }

        switch (element.ValueKind)
        {
            case JsonValueKind.Undefined:
            case JsonValueKind.Null:
                value = null;
                error = new { success = true };
                return true;
            case JsonValueKind.String:
                value = element.GetString();
                error = new { success = true };
                return true;
            case JsonValueKind.True:
            case JsonValueKind.False:
                value = element.GetBoolean();
                error = new { success = true };
                return true;
            case JsonValueKind.Number:
                if (element.TryGetInt64(out var intValue))
                    value = intValue;
                else if (element.TryGetDouble(out var doubleValue))
                    value = doubleValue;
                else
                    value = element.GetRawText();
                error = new { success = true };
                return true;
            default:
                error = new { success = false, error = $"{propertyName} must be a primitive JSON value when provided" };
                return false;
        }
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
