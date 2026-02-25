using System.Text.Json;
using System.Text.Json.Serialization;
using System.Linq;
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
                    var document = JsonDocument.Parse(line);
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
                new { name = "execute_sql", description = "Execute a SQL statement against the connected Access database. For SELECT queries, returns columns and rows. For action queries, returns rows_affected.", inputSchema = new { type = "object", properties = new { sql = new { type = "string" }, max_rows = new { type = "integer" } }, required = new string[] { "sql" } } },
                new { name = "execute_query_md", description = "Execute a SQL statement and return result as a markdown table (or action-query summary).", inputSchema = new { type = "object", properties = new { sql = new { type = "string" }, max_rows = new { type = "integer" } }, required = new string[] { "sql" } } },
                new { name = "describe_table", description = "Describe a table schema including columns, nullability, defaults, and primary key columns.", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "create_table", description = "Create a new table in the database", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, fields = new { type = "array", items = new { type = "object", properties = new { name = new { type = "string" }, type = new { type = "string" }, size = new { type = "integer" }, required = new { type = "boolean" }, allow_zero_length = new { type = "boolean" } } } } }, required = new string[] { "table_name", "fields" } } },
                new { name = "delete_table", description = "Delete a table from the database", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "launch_access", description = "Launch Microsoft Access application", inputSchema = new { type = "object", properties = new { } } },
                new { name = "close_access", description = "Close Microsoft Access application", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_forms", description = "Get list of all forms in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_reports", description = "Get list of all reports in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_macros", description = "Get list of all macros in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_modules", description = "Get list of all modules in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "open_form", description = "Open a form in Access", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "close_form", description = "Close a form in Access", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
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
                new { name = "export_form_to_text", description = "Export a form to text format", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "import_form_from_text", description = "Import a form from text format", inputSchema = new { type = "object", properties = new { form_data = new { type = "string" } }, required = new string[] { "form_data" } } },
                new { name = "delete_form", description = "Delete a form from the database", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "export_report_to_text", description = "Export a report to text format", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "import_report_from_text", description = "Import a report from text format", inputSchema = new { type = "object", properties = new { report_data = new { type = "string" } }, required = new string[] { "report_data" } } },
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
            "execute_sql" => HandleExecuteSql(accessService, toolArguments),
            "execute_query_md" => HandleExecuteQueryMd(accessService, toolArguments),
            "describe_table" => HandleDescribeTable(accessService, toolArguments),
            "create_table" => HandleCreateTable(accessService, toolArguments),
            "delete_table" => HandleDeleteTable(accessService, toolArguments),
            "launch_access" => HandleLaunchAccess(accessService, toolArguments),
            "close_access" => HandleCloseAccess(accessService, toolArguments),
            "get_forms" => HandleGetForms(accessService, toolArguments),
            "get_reports" => HandleGetReports(accessService, toolArguments),
            "get_macros" => HandleGetMacros(accessService, toolArguments),
            "get_modules" => HandleGetModules(accessService, toolArguments),
            "open_form" => HandleOpenForm(accessService, toolArguments),
            "close_form" => HandleCloseForm(accessService, toolArguments),
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
            return new { success = false, error = ex.Message };
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
            return new { success = false, error = ex.Message };
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
            return new { success = false, error = ex.Message };
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
            return new { success = false, error = ex.Message };
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
            return new { success = false, error = ex.Message };
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
            return new { success = false, error = ex.Message };
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

    static object HandleExportFormToText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            var formData = accessService.ExportFormToText(formName);
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
                
            accessService.ImportFormFromText(formData);
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
                
            var reportData = accessService.ExportReportToText(reportName);
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
                
            accessService.ImportReportFromText(reportData);
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
