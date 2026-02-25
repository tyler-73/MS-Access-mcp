using System.Data;
using System.Data.OleDb;
using Microsoft.VisualBasic.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace MS.Access.MCP.Interop
{
    public class AccessInteropService : IDisposable
    {
        private OleDbConnection? _oleDbConnection;
        private dynamic? _accessApplication;
        private string? _currentDatabasePath;
        private int _oleDbReleaseDepth = 0;
        private bool _restoreOleDbAfterRelease = false;
        private string? _accessDatabasePath;
        private bool _accessDatabaseOpenedExclusive = false;
        private bool _disposed = false;

        #region 1. Connection Management

        public void Connect(string databasePath)
        {
            if (!File.Exists(databasePath))
                throw new FileNotFoundException($"Database file not found: {databasePath}");

            _currentDatabasePath = databasePath;
            OpenOleDbConnection(databasePath);
        }

        public void Disconnect()
        {
            _oleDbConnection?.Close();
            _oleDbConnection?.Dispose();
            _oleDbConnection = null;
            _currentDatabasePath = null;
            _accessDatabasePath = null;
            _accessDatabaseOpenedExclusive = false;
        }

        public bool IsConnected => !string.IsNullOrWhiteSpace(_currentDatabasePath);

        #endregion

        #region 2. Data Access Object Models

        public List<TableInfo> GetTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var tables = new List<TableInfo>();
            
            // Use OleDb to get table information
            var schema = _oleDbConnection!.GetSchema("Tables");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(tableName) && !tableName.StartsWith("~"))
                {
                    var fields = GetTableFields(tableName);
                    tables.Add(new TableInfo
                    {
                        Name = tableName,
                        Fields = fields,
                        RecordCount = GetTableRecordCount(tableName)
                    });
                }
            }

            return tables;
        }

        public List<QueryInfo> GetQueries()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var queries = new List<QueryInfo>();
            
            // Use OleDb to get query information
            var schema = _oleDbConnection!.GetSchema("Views");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var queryName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(queryName))
                {
                    queries.Add(new QueryInfo
                    {
                        Name = queryName,
                        SQL = "", // SQL not available through schema
                        Type = "Query"
                    });
                }
            }

            return queries;
        }

        public List<RelationshipInfo> GetRelationships()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var relationships = new List<RelationshipInfo>();

            try
            {
                // Not all ACE providers expose this collection; return empty on unsupported providers.
                var schema = _oleDbConnection!.GetSchema("ForeignKeys");

                foreach (DataRow row in schema.Rows)
                {
                    relationships.Add(new RelationshipInfo
                    {
                        Name = row["FK_NAME"]?.ToString() ?? "",
                        Table = row["TABLE_NAME"]?.ToString() ?? "",
                        ForeignTable = row["REFERENCED_TABLE_NAME"]?.ToString() ?? "",
                        Attributes = ""
                    });
                }
            }
            catch
            {
                // Keep compatibility with providers that do not publish ForeignKeys metadata.
            }

            return relationships;
        }

        public void CreateTable(string tableName, List<FieldInfo> fields)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var fieldDefinitions = new List<string>();
            foreach (var field in fields)
            {
                var fieldDef = $"[{field.Name}] {field.Type}";
                if (field.Size > 0 && field.Type.ToLower() == "text")
                    fieldDef += $"({field.Size})";
                if (field.Required)
                    fieldDef += " NOT NULL";
                fieldDefinitions.Add(fieldDef);
            }

            var createSql = $"CREATE TABLE [{tableName}] ({string.Join(", ", fieldDefinitions)})";
            var command = new OleDbCommand(createSql, _oleDbConnection);
            command.ExecuteNonQuery();
        }

        public void DeleteTable(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();
            var command = new OleDbCommand($"DROP TABLE [{tableName}]", _oleDbConnection);
            command.ExecuteNonQuery();
        }

        public SqlExecutionResult ExecuteSql(string sql, int maxRows = 200)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required", nameof(sql));
            if (maxRows <= 0) throw new ArgumentOutOfRangeException(nameof(maxRows), "maxRows must be greater than 0");
            EnsureOleDbConnection();

            using var command = new OleDbCommand(sql, _oleDbConnection);
            using var reader = command.ExecuteReader();

            if (reader == null)
            {
                return new SqlExecutionResult
                {
                    IsQuery = false,
                    RowsAffected = 0
                };
            }

            if (reader.FieldCount == 0)
            {
                while (reader.Read())
                {
                    // Consume any provider-side status rows to finalize RecordsAffected.
                }

                return new SqlExecutionResult
                {
                    IsQuery = false,
                    RowsAffected = reader.RecordsAffected
                };
            }

            var rawColumnNames = new string[reader.FieldCount];
            for (var i = 0; i < reader.FieldCount; i++)
            {
                rawColumnNames[i] = reader.GetName(i);
            }

            var columnNames = MakeUniqueColumnNames(rawColumnNames);
            var rows = new List<Dictionary<string, object?>>();
            var truncated = false;

            while (reader.Read())
            {
                if (rows.Count >= maxRows)
                {
                    truncated = true;
                    break;
                }

                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < columnNames.Count; i++)
                {
                    var value = reader.IsDBNull(i) ? null : NormalizeValue(reader.GetValue(i));
                    row[columnNames[i]] = value;
                }

                rows.Add(row);
            }

            return new SqlExecutionResult
            {
                IsQuery = true,
                Columns = columnNames,
                Rows = rows,
                RowCount = rows.Count,
                Truncated = truncated
            };
        }

        public string ExecuteQueryMarkdown(string sql, int maxRows = 100)
        {
            var result = ExecuteSql(sql, maxRows);

            if (!result.IsQuery)
            {
                return $"Statement executed successfully. Rows affected: {result.RowsAffected}.";
            }

            if (result.Columns.Count == 0)
            {
                return "No columns returned.";
            }

            var builder = new StringBuilder();
            builder.Append("| ");
            builder.Append(string.Join(" | ", result.Columns.Select(EscapeMarkdownCell)));
            builder.AppendLine(" |");

            builder.Append("| ");
            builder.Append(string.Join(" | ", result.Columns.Select(_ => "---")));
            builder.AppendLine(" |");

            foreach (var row in result.Rows)
            {
                builder.Append("| ");
                builder.Append(string.Join(" | ", result.Columns.Select(column =>
                {
                    row.TryGetValue(column, out var value);
                    return EscapeMarkdownCell(value?.ToString());
                })));
                builder.AppendLine(" |");
            }

            if (result.Truncated)
            {
                builder.AppendLine();
                builder.AppendLine($"_Results truncated to {maxRows} rows._");
            }

            return builder.ToString().TrimEnd();
        }

        public TableDefinition DescribeTable(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            EnsureOleDbConnection();

            var columnsSchema = _oleDbConnection!.GetSchema("Columns", new string[] { null!, null!, tableName, null! });
            if (columnsSchema.Rows.Count == 0)
            {
                throw new InvalidOperationException($"Table not found or has no visible columns: {tableName}");
            }

            var primaryKeyColumns = GetPrimaryKeyColumns(tableName);
            var columns = new List<TableColumnDefinition>();

            foreach (DataRow row in columnsSchema.Rows)
            {
                var columnName = GetRowString(row, "COLUMN_NAME") ?? string.Empty;
                var dataTypeCode = GetRowInt(row, "DATA_TYPE");
                var dataTypeName = dataTypeCode.HasValue
                    ? ((OleDbType)dataTypeCode.Value).ToString()
                    : "Unknown";

                columns.Add(new TableColumnDefinition
                {
                    Name = columnName,
                    DataType = dataTypeName,
                    DataTypeCode = dataTypeCode,
                    OrdinalPosition = GetRowInt(row, "ORDINAL_POSITION"),
                    MaxLength = GetRowInt(row, "CHARACTER_MAXIMUM_LENGTH"),
                    NumericPrecision = GetRowInt(row, "NUMERIC_PRECISION"),
                    NumericScale = GetRowInt(row, "NUMERIC_SCALE"),
                    IsNullable = string.Equals(GetRowString(row, "IS_NULLABLE"), "YES", StringComparison.OrdinalIgnoreCase),
                    IsPrimaryKey = primaryKeyColumns.Contains(columnName),
                    HasDefault = GetRowBool(row, "COLUMN_HASDEFAULT"),
                    DefaultValue = GetRowString(row, "COLUMN_DEFAULT")
                });
            }

            columns = columns
                .OrderBy(c => c.OrdinalPosition ?? int.MaxValue)
                .ThenBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new TableDefinition
            {
                TableName = tableName,
                Columns = columns,
                PrimaryKeyColumns = primaryKeyColumns.OrderBy(c => c, StringComparer.OrdinalIgnoreCase).ToList()
            };
        }

        #endregion

        #region 3. COM Automation (Simplified)

        public void LaunchAccess()
        {
            var accessApp = EnsureAccessApplication(openCurrentDatabase: true, requireExclusive: false);
            accessApp.Visible = true;
        }

        public void CloseAccess()
        {
            ResetAccessApplication();
        }

        public List<FormInfo> GetForms()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var forms = new List<FormInfo>();

            try
            {
                var accessApp = EnsureAccessApplication(openCurrentDatabase: true);
                foreach (var form in accessApp.CurrentProject.AllForms)
                {
                    forms.Add(new FormInfo
                    {
                        Name = form.Name ?? "",
                        FullName = form.FullName ?? form.Name ?? "",
                        Type = "Form"
                    });
                }

                if (forms.Count > 0)
                    return forms;
            }
            catch
            {
                // Fall back to OleDb system table scan.
            }

            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32768", _oleDbConnection);
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    forms.Add(new FormInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Form"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible.
            }

            return forms;
        }

        public List<ReportInfo> GetReports()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reports = new List<ReportInfo>();

            try
            {
                var accessApp = EnsureAccessApplication(openCurrentDatabase: true);
                foreach (var report in accessApp.CurrentProject.AllReports)
                {
                    reports.Add(new ReportInfo
                    {
                        Name = report.Name ?? "",
                        FullName = report.FullName ?? report.Name ?? "",
                        Type = "Report"
                    });
                }

                if (reports.Count > 0)
                    return reports;
            }
            catch
            {
                // Fall back to OleDb system table scan.
            }

            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32764", _oleDbConnection);
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    reports.Add(new ReportInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Report"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible.
            }

            return reports;
        }

        public List<MacroInfo> GetMacros()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var macros = new List<MacroInfo>();

            try
            {
                var accessApp = EnsureAccessApplication(openCurrentDatabase: true);
                foreach (var macro in accessApp.CurrentProject.AllMacros)
                {
                    macros.Add(new MacroInfo
                    {
                        Name = macro.Name ?? "",
                        FullName = macro.FullName ?? macro.Name ?? "",
                        Type = "Macro"
                    });
                }

                if (macros.Count > 0)
                    return macros;
            }
            catch
            {
                // Fall back to OleDb system table scan.
            }

            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32766", _oleDbConnection);
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    macros.Add(new MacroInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Macro"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible.
            }

            return macros;
        }

        public List<ModuleInfo> GetModules()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var modules = new List<ModuleInfo>();

            try
            {
                var accessApp = EnsureAccessApplication(openCurrentDatabase: true);
                foreach (var module in accessApp.CurrentProject.AllModules)
                {
                    modules.Add(new ModuleInfo
                    {
                        Name = module.Name ?? "",
                        FullName = module.FullName ?? module.Name ?? "",
                        Type = "Module"
                    });
                }

                if (modules.Count > 0)
                    return modules;
            }
            catch
            {
                // Fall back to OleDb system table scan.
            }

            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32761", _oleDbConnection);
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    modules.Add(new ModuleInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Module"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible.
            }

            return modules;
        }

        public void OpenForm(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required", nameof(formName));

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.OpenForm(formName),
                requireExclusive: false,
                releaseOleDb: false);
        }

        public void CloseForm(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required", nameof(formName));

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.Close(2, formName),
                requireExclusive: false,
                releaseOleDb: false);
        }

        #endregion

        #region 4. VBA Extensibility

        public List<VBAProjectInfo> GetVBAProjects()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var projects = new List<VBAProjectInfo>();

            try
            {
                var accessApp = EnsureAccessApplication(openCurrentDatabase: true);
                foreach (var project in accessApp.VBE.VBProjects)
                {
                    var modules = new List<VBAModuleInfo>();
                    foreach (var component in project.VBComponents)
                    {
                        modules.Add(new VBAModuleInfo
                        {
                            Name = SafeToString(TryGetDynamicProperty(component, "Name")) ?? "",
                            Type = MapVbComponentType(ToInt32(TryGetDynamicProperty(component, "Type"))),
                            HasCode = ToInt32(TryGetDynamicProperty(TryGetDynamicProperty(component, "CodeModule"), "CountOfLines")) > 0
                        });
                    }

                    projects.Add(new VBAProjectInfo
                    {
                        Name = SafeToString(TryGetDynamicProperty(project, "Name")) ?? "VBAProject",
                        Description = SafeToString(TryGetDynamicProperty(project, "Description")) ?? "",
                        Modules = modules.OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase).ToList()
                    });
                }

                if (projects.Count > 0)
                    return projects;
            }
            catch
            {
                // Fall back to lightweight module listing through system tables.
            }

            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32761", _oleDbConnection);
                using var reader = command.ExecuteReader();
                var modules = new List<VBAModuleInfo>();
                while (reader.Read())
                {
                    modules.Add(new VBAModuleInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        Type = "Module",
                        HasCode = true
                    });
                }

                projects.Add(new VBAProjectInfo
                {
                    Name = "CurrentProject",
                    Description = "Current Access Project",
                    Modules = modules.OrderBy(m => m.Name, StringComparer.OrdinalIgnoreCase).ToList()
                });
            }
            catch
            {
                // MSysObjects might not be accessible.
            }

            return projects;
        }

        public string GetVBACode(string projectName, string moduleName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required", nameof(moduleName));

            var accessApp = EnsureAccessApplication(openCurrentDatabase: true);
            var component = FindOrCreateVbComponent(accessApp, projectName, moduleName, false)
                ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
            var codeModule = TryGetDynamicProperty(component, "CodeModule")
                ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

            var lineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfLines"));
            if (lineCount <= 0)
                return string.Empty;

            return SafeToString(TryGetDynamicProperty(codeModule, "Lines", 1, lineCount)) ?? string.Empty;
        }

        public void SetVBACode(string projectName, string moduleName, string code)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required", nameof(moduleName));

            ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName, moduleName, true)
                    ?? throw new InvalidOperationException($"Unable to create or locate VBA module '{moduleName}'.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                var existingLines = ToInt32(TryGetDynamicProperty(codeModule, "CountOfLines"));
                if (existingLines > 0)
                {
                    InvokeDynamicMethod(codeModule, "DeleteLines", 1, existingLines);
                }

                if (!string.IsNullOrWhiteSpace(code))
                {
                    InvokeDynamicMethod(codeModule, "AddFromString", NormalizeLineEndings(code));
                }

                TrySaveModule(accessApp, moduleName);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void AddVBAProcedure(string projectName, string moduleName, string procedureName, string code)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required", nameof(moduleName));
            if (string.IsNullOrWhiteSpace(procedureName)) throw new ArgumentException("Procedure name is required", nameof(procedureName));
            if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Procedure code is required", nameof(code));

            ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName, moduleName, true)
                    ?? throw new InvalidOperationException($"Unable to create or locate VBA module '{moduleName}'.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                var lineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfLines"));
                var normalized = NormalizeLineEndings(code);
                if (lineCount > 0)
                {
                    normalized = "\r\n" + normalized;
                }

                InvokeDynamicMethod(codeModule, "AddFromString", normalized);
                TrySaveModule(accessApp, moduleName);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void CompileVBA()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                // 125 = acCmdCompileAndSaveAllModules
                try
                {
                    accessApp.DoCmd.RunCommand(125);
                }
                catch
                {
                    accessApp.RunCommand(125);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        #endregion

        #region 5. System Table Metadata Access

        public List<SystemTableInfo> GetSystemTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var systemTables = new List<SystemTableInfo>();
            var schema = _oleDbConnection!.GetSchema("Tables");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(tableName) && (tableName.StartsWith("~") || tableName.StartsWith("MSys")))
                {
                    systemTables.Add(new SystemTableInfo
                    {
                        Name = tableName,
                        DateCreated = DateTime.Now, // Not available through OleDb
                        LastUpdated = DateTime.Now, // Not available through OleDb
                        RecordCount = GetTableRecordCount(tableName)
                    });
                }
            }

            return systemTables;
        }

        public List<MetadataInfo> GetObjectMetadata()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var metadata = new List<MetadataInfo>();
            
            try
            {
                // Query MSysObjects table for object metadata
                var command = new OleDbCommand("SELECT * FROM MSysObjects", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    metadata.Add(new MetadataInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        Type = reader["Type"]?.ToString() ?? "",
                        Flags = reader["Flags"]?.ToString() ?? "",
                        DateCreated = reader["DateCreate"]?.ToString() ?? "",
                        DateModified = reader["DateUpdate"]?.ToString() ?? ""
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible, return empty list
            }

            return metadata;
        }

        #endregion

        #region 6. Form & Control Discovery & Editing APIs

        public bool FormExists(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            try
            {
                var accessApp = EnsureAccessApplication(openCurrentDatabase: true);
                foreach (var form in accessApp.CurrentProject.AllForms)
                {
                    if (string.Equals(form.Name, formName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch
            {
                // Fall back to OleDb query.
            }

            try
            {
                EnsureOleDbConnection();
                var command = new OleDbCommand("SELECT COUNT(*) FROM MSysObjects WHERE Name = ? AND Type = -32768", _oleDbConnection);
                command.Parameters.AddWithValue("@Name", formName);
                var count = Convert.ToInt32(command.ExecuteScalar());
                return count > 0;
            }
            catch
            {
                return false;
            }
        }

        public List<ControlInfo> GetFormControls(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required", nameof(formName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    var controlObjects = GetControlObjects((object)form);
                    return controlObjects
                        .Select(BuildControlInfo)
                        .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
                finally
                {
                    if (openedHere)
                    {
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        private List<ControlInfo> GetReportControls(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required", nameof(reportName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    var controlObjects = GetControlObjects((object)report);
                    return controlObjects
                        .Select(BuildControlInfo)
                        .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
                finally
                {
                    if (openedHere)
                    {
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public ControlProperties GetControlProperties(string formName, string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required", nameof(formName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("Control name is required", nameof(controlName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var form = EnsureFormOpen(accessApp, formName, false, out openedHere);
                try
                {
                    var control = GetControlByName(form, controlName)
                        ?? throw new InvalidOperationException($"Control '{controlName}' was not found on form '{formName}'.");

                    return new ControlProperties
                    {
                        Name = SafeToString(TryGetDynamicProperty(control, "Name")) ?? controlName,
                        Type = MapControlType(ToInt32(TryGetDynamicProperty(control, "ControlType"))),
                        Left = ToInt32(TryGetDynamicProperty(control, "Left")),
                        Top = ToInt32(TryGetDynamicProperty(control, "Top")),
                        Width = ToInt32(TryGetDynamicProperty(control, "Width")),
                        Height = ToInt32(TryGetDynamicProperty(control, "Height")),
                        Visible = ToBool(TryGetDynamicProperty(control, "Visible"), true),
                        Enabled = ToBool(TryGetDynamicProperty(control, "Enabled"), true),
                        BackColor = ToInt32(TryGetDynamicProperty(control, "BackColor")),
                        ForeColor = ToInt32(TryGetDynamicProperty(control, "ForeColor")),
                        FontName = SafeToString(TryGetDynamicProperty(control, "FontName")) ?? "",
                        FontSize = ToInt32(TryGetDynamicProperty(control, "FontSize")),
                        FontBold = ToBool(TryGetDynamicProperty(control, "FontBold"), false),
                        FontItalic = ToBool(TryGetDynamicProperty(control, "FontItalic"), false)
                    };
                }
                finally
                {
                    if (openedHere)
                    {
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetControlProperty(string formName, string controlName, string propertyName, object value)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required", nameof(formName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("Control name is required", nameof(controlName));
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property name is required", nameof(propertyName));

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                try
                {
                    var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
                    var control = GetControlByName(form, controlName)
                        ?? throw new InvalidOperationException($"Control '{controlName}' was not found on form '{formName}'.");

                    var existingValue = TryGetDynamicProperty(control, propertyName);
                    var convertedValue = ConvertValueForProperty(value, existingValue);
                    SetDynamicProperty(control, propertyName, convertedValue);
                    accessApp.DoCmd.Save(2, formName);
                }
                finally
                {
                    if (openedHere)
                    {
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        #endregion

        #region 7. Persistence & Versioning

        public string ExportFormToText(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var formData = new
            {
                Name = formName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetFormControls(formName),
                VBA = TryGetFormVbaCode(formName)
            };

            return JsonSerializer.Serialize(formData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportFormFromText(string formData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var formInfo = JsonSerializer.Deserialize<FormExportData>(formData);
            if (formInfo == null) throw new ArgumentException("Invalid form data");
            if (string.IsNullOrWhiteSpace(formInfo.Name)) throw new ArgumentException("Form name is required in form data");

            ExecuteComOperation(accessApp =>
            {
                TryDeleteObject(accessApp, 2, formInfo.Name);

                var form = accessApp.CreateForm();
                var temporaryName = SafeToString(TryGetDynamicProperty(form, "Name")) ?? throw new InvalidOperationException("Failed to create temporary form.");

                foreach (var control in formInfo.Controls ?? new List<ControlInfo>())
                {
                    TryCreateFormControl(accessApp, temporaryName, control);
                }

                accessApp.DoCmd.Close(2, temporaryName, 1);
                accessApp.DoCmd.Rename(formInfo.Name, 2, temporaryName);

                if (!string.IsNullOrWhiteSpace(formInfo.VBA))
                {
                    SetVBACode("CurrentProject", formInfo.Name, formInfo.VBA);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteForm(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.DeleteObject(2, formName),
                requireExclusive: true,
                releaseOleDb: true);
        }

        public string ExportReportToText(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reportData = new
            {
                Name = reportName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetReportControls(reportName)
            };

            return JsonSerializer.Serialize(reportData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportReportFromText(string reportData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reportInfo = JsonSerializer.Deserialize<ReportExportData>(reportData);
            if (reportInfo == null) throw new ArgumentException("Invalid report data");
            if (string.IsNullOrWhiteSpace(reportInfo.Name)) throw new ArgumentException("Report name is required in report data");

            ExecuteComOperation(accessApp =>
            {
                TryDeleteObject(accessApp, 3, reportInfo.Name);

                var report = accessApp.CreateReport();
                var temporaryName = SafeToString(TryGetDynamicProperty(report, "Name")) ?? throw new InvalidOperationException("Failed to create temporary report.");

                foreach (var control in reportInfo.Controls ?? new List<ControlInfo>())
                {
                    TryCreateReportControl(accessApp, temporaryName, control);
                }

                accessApp.DoCmd.Close(3, temporaryName, 1);
                accessApp.DoCmd.Rename(reportInfo.Name, 3, temporaryName);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteReport(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.DeleteObject(3, reportName),
                requireExclusive: true,
                releaseOleDb: true);
        }

        #endregion

        #region Helper Methods

        private List<FieldInfo> GetTableFields(string tableName)
        {
            var fields = new List<FieldInfo>();
            
            try
            {
                EnsureOleDbConnection();
                var schema = _oleDbConnection!.GetSchema("Columns", new string[] { null!, null!, tableName });
                
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    fields.Add(new FieldInfo
                    {
                        Name = row["COLUMN_NAME"]?.ToString() ?? "",
                        Type = row["DATA_TYPE"]?.ToString() ?? "",
                        Size = Convert.ToInt32(row["CHARACTER_MAXIMUM_LENGTH"] ?? 0),
                        Required = row["IS_NULLABLE"]?.ToString() == "NO",
                        AllowZeroLength = true // Default value
                    });
                }
            }
            catch
            {
                // Return empty list if table doesn't exist or can't be accessed
            }

            return fields;
        }

        private long GetTableRecordCount(string tableName)
        {
            try
            {
                EnsureOleDbConnection();
                var command = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", _oleDbConnection);
                return Convert.ToInt64(command.ExecuteScalar());
            }
            catch
            {
                return 0;
            }
        }

        private void OpenOleDbConnection(string databasePath)
        {
            _oleDbConnection?.Close();
            _oleDbConnection?.Dispose();

            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};";
            _oleDbConnection = new OleDbConnection(connectionString);
            _oleDbConnection.Open();
        }

        private void EnsureOleDbConnection()
        {
            if (_oleDbConnection?.State == ConnectionState.Open)
                return;

            if (string.IsNullOrWhiteSpace(_currentDatabasePath))
                throw new InvalidOperationException("Not connected to database");

            try
            {
                OpenOleDbConnection(_currentDatabasePath);
            }
            catch (Exception ex) when (IsRecoverableOleDbLockError(ex) && TryReleaseExclusiveAccessLock())
            {
                OpenOleDbConnection(_currentDatabasePath);
            }
        }

        private void ExecuteWithOleDbReleased(Action action)
        {
            var isOuterScope = _oleDbReleaseDepth == 0;
            if (isOuterScope)
            {
                _restoreOleDbAfterRelease = IsConnected && !string.IsNullOrWhiteSpace(_currentDatabasePath);
                if (_restoreOleDbAfterRelease)
                {
                    _oleDbConnection?.Close();
                    _oleDbConnection?.Dispose();
                    _oleDbConnection = null;
                }
            }

            _oleDbReleaseDepth++;
            try
            {
                action();
            }
            finally
            {
                _oleDbReleaseDepth--;
                if (isOuterScope)
                {
                    if (_restoreOleDbAfterRelease && _oleDbConnection == null && !string.IsNullOrWhiteSpace(_currentDatabasePath))
                    {
                        try
                        {
                            OpenOleDbConnection(_currentDatabasePath);
                        }
                        catch
                        {
                            // Defer reconnection until next OleDb operation.
                        }
                    }

                    _restoreOleDbAfterRelease = false;
                }
            }
        }

        private void ExecuteComOperation(Action<dynamic> operation, bool requireExclusive, bool releaseOleDb)
        {
            _ = ExecuteComOperation<object?>(
                accessApp =>
                {
                    operation(accessApp);
                    return null;
                },
                requireExclusive,
                releaseOleDb);
        }

        private T ExecuteComOperation<T>(Func<dynamic, T> operation, bool requireExclusive, bool releaseOleDb)
        {
            if (!releaseOleDb)
            {
                return ExecuteComOperationCore(operation, requireExclusive);
            }

            T result = default!;
            ExecuteWithOleDbReleased(() =>
            {
                result = ExecuteComOperationCore(operation, requireExclusive);
            });

            return result;
        }

        private T ExecuteComOperationCore<T>(Func<dynamic, T> operation, bool requireExclusive)
        {
            Exception? lastRecoverableError = null;

            for (var attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    var accessApp = EnsureAccessApplication(openCurrentDatabase: true, requireExclusive: requireExclusive);
                    return operation(accessApp);
                }
                catch (Exception ex) when (attempt == 0 && IsRecoverableAccessStateError(ex))
                {
                    lastRecoverableError = ex;
                    ResetAccessApplication();
                }
            }

            throw lastRecoverableError ?? new InvalidOperationException("COM operation failed.");
        }

        private static bool IsRecoverableAccessStateError(Exception ex)
        {
            if (ex is COMException comException)
            {
                var errorCode = unchecked((uint)comException.ErrorCode);
                if (errorCode == 0x800ADEB9 || errorCode == 0x800A0BB9)
                {
                    return true;
                }
            }

            var message = ex.Message ?? string.Empty;
            if (message.IndexOf("exclusive access", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("opened or locked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("opened or locked by another user", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("cannot be opened or locked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("prevents it from being opened or locked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("has been placed in a state", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return ex.InnerException != null && IsRecoverableAccessStateError(ex.InnerException);
        }

        private void ResetAccessApplication()
        {
            if (_accessApplication != null)
            {
                try
                {
                    // 2 = acQuitSaveNone (avoid save prompts during recovery/cleanup)
                    _accessApplication.Quit(2);
                }
                catch
                {
                    // Ignore shutdown failures.
                }

                try
                {
                    if (Marshal.IsComObject(_accessApplication))
                    {
                        Marshal.FinalReleaseComObject(_accessApplication);
                    }
                }
                catch
                {
                    // Ignore RCW cleanup failures.
                }

                _accessApplication = null;
            }

            _accessDatabasePath = null;
            _accessDatabaseOpenedExclusive = false;
        }

        private dynamic EnsureAccessApplication(bool openCurrentDatabase, bool requireExclusive = false)
        {
            if (_accessApplication == null)
            {
                var accessType = Type.GetTypeFromProgID("Access.Application", throwOnError: false);
                if (accessType == null)
                    throw new InvalidOperationException("Microsoft Access COM automation is not available on this machine.");

                _accessApplication = Activator.CreateInstance(accessType);
                if (_accessApplication == null)
                    throw new InvalidOperationException("Failed to create Access.Application COM instance.");
            }

            if (openCurrentDatabase && !string.IsNullOrWhiteSpace(_currentDatabasePath))
            {
                EnsureCurrentDatabaseOpen(_accessApplication, _currentDatabasePath, requireExclusive);
            }

            return _accessApplication;
        }

        private void EnsureCurrentDatabaseOpen(dynamic accessApplication, string databasePath, bool requireExclusive)
        {
            bool shouldOpen = true;
            bool shouldCloseCurrent = false;

            try
            {
                var currentPath = accessApplication.CurrentProject?.FullName;
                if (!string.IsNullOrWhiteSpace(currentPath))
                {
                    if (PathsMatch(currentPath, databasePath))
                    {
                        var alreadyKnownExclusive =
                            _accessDatabaseOpenedExclusive &&
                            !string.IsNullOrWhiteSpace(_accessDatabasePath) &&
                            PathsMatch(_accessDatabasePath, databasePath);

                        if (requireExclusive && !alreadyKnownExclusive)
                        {
                            shouldCloseCurrent = true;
                        }
                        else
                        {
                            shouldOpen = false;
                        }
                    }
                    else
                    {
                        shouldCloseCurrent = true;
                    }
                }
            }
            catch
            {
                // CurrentProject may throw if no database is currently open.
            }

            if (shouldCloseCurrent)
            {
                try
                {
                    accessApplication.CloseCurrentDatabase();
                }
                catch
                {
                    // Continue and attempt to reopen regardless.
                }
            }

            if (shouldOpen)
            {
                accessApplication.OpenCurrentDatabase(databasePath, requireExclusive);
                _accessDatabasePath = databasePath;
                _accessDatabaseOpenedExclusive = requireExclusive;
            }
            else
            {
                _accessDatabasePath = databasePath;
                if (requireExclusive)
                {
                    _accessDatabaseOpenedExclusive = true;
                }
            }
        }

        private static bool PathsMatch(string left, string right)
        {
            try
            {
                var leftFull = Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                var rightFull = Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                return string.Equals(leftFull, rightFull, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static bool IsRecoverableOleDbLockError(Exception ex)
        {
            var message = ex.Message ?? string.Empty;
            if (message.IndexOf("file already in use", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("could not use", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("opened or locked", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return ex.InnerException != null && IsRecoverableOleDbLockError(ex.InnerException);
        }

        private bool TryReleaseExclusiveAccessLock()
        {
            if (_accessApplication == null || !_accessDatabaseOpenedExclusive)
                return false;

            try
            {
                _accessApplication!.CloseCurrentDatabase();
                _accessDatabasePath = null;
                _accessDatabaseOpenedExclusive = false;
                return true;
            }
            catch
            {
                ResetAccessApplication();
                return false;
            }
        }

        private dynamic? FindOrCreateVbComponent(dynamic accessApp, string projectName, string moduleName, bool createIfMissing)
        {
            var project = FindVbProject(accessApp, projectName);
            if (project == null)
                throw new InvalidOperationException("No VBA project is available in the current Access database.");

            var component = FindVbComponent(project, moduleName);
            if (component != null || !createIfMissing)
                return component;

            // 1 = Standard Module (vbext_ct_StdModule)
            var newComponent = InvokeDynamicMethod(project.VBComponents, "Add", 1);
            if (newComponent == null)
                throw new InvalidOperationException($"Failed to create VBA module '{moduleName}'.");

            SetDynamicProperty(newComponent, "Name", moduleName);
            var actualName = SafeToString(TryGetDynamicProperty(newComponent, "Name"));
            if (!string.Equals(actualName, moduleName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"Created VBA module but could not assign requested name '{moduleName}'.");

            return newComponent;
        }

        private static dynamic? FindVbProject(dynamic accessApp, string? projectName)
        {
            var normalizedName = string.IsNullOrWhiteSpace(projectName) || string.Equals(projectName, "CurrentProject", StringComparison.OrdinalIgnoreCase)
                ? null
                : projectName;

            if (normalizedName == null)
            {
                try
                {
                    var activeProject = TryGetDynamicProperty(accessApp.VBE, "ActiveVBProject");
                    if (activeProject != null)
                        return activeProject;
                }
                catch
                {
                    // Fall back to iterating VBProjects.
                }
            }

            foreach (var project in accessApp.VBE.VBProjects)
            {
                var name = SafeToString(TryGetDynamicProperty(project, "Name"));
                if (normalizedName == null || string.Equals(name, normalizedName, StringComparison.OrdinalIgnoreCase))
                    return project;
            }

            return null;
        }

        private static dynamic? FindVbComponent(dynamic vbProject, string moduleName)
        {
            foreach (var component in vbProject.VBComponents)
            {
                var name = SafeToString(TryGetDynamicProperty(component, "Name"));
                if (string.Equals(name, moduleName, StringComparison.OrdinalIgnoreCase))
                    return component;
            }

            return null;
        }

        private static string MapVbComponentType(int componentType)
        {
            return componentType switch
            {
                1 => "StandardModule",
                2 => "ClassModule",
                3 => "Form",
                11 => "ActiveXDesigner",
                100 => "Document",
                _ => $"Unknown({componentType})"
            };
        }

        private dynamic EnsureFormOpen(dynamic accessApp, string formName, bool openInDesignView, out bool openedHere)
        {
            openedHere = false;

            if (!IsFormLoaded(accessApp, formName))
            {
                var view = openInDesignView ? 1 : 0; // 1 = Design view, 0 = Normal view
                accessApp.DoCmd.OpenForm(formName, view);
                openedHere = true;
            }

            return FindObjectByName(accessApp.Forms, formName)
                ?? throw new InvalidOperationException($"Form '{formName}' is not loaded.");
        }

        private dynamic EnsureReportOpen(dynamic accessApp, string reportName, bool openInDesignView, out bool openedHere)
        {
            openedHere = false;

            if (!IsReportLoaded(accessApp, reportName))
            {
                var view = openInDesignView ? 1 : 0; // 1 = Design view, 0 = Normal view
                accessApp.DoCmd.OpenReport(reportName, view);
                openedHere = true;
            }

            return FindObjectByName(accessApp.Reports, reportName)
                ?? throw new InvalidOperationException($"Report '{reportName}' is not loaded.");
        }

        private static bool IsFormLoaded(dynamic accessApp, string formName)
        {
            foreach (var form in accessApp.CurrentProject.AllForms)
            {
                var name = SafeToString(TryGetDynamicProperty(form, "Name"));
                if (!string.Equals(name, formName, StringComparison.OrdinalIgnoreCase))
                    continue;

                return ToBool(TryGetDynamicProperty(form, "IsLoaded"), false);
            }

            return false;
        }

        private static bool IsReportLoaded(dynamic accessApp, string reportName)
        {
            foreach (var report in accessApp.CurrentProject.AllReports)
            {
                var name = SafeToString(TryGetDynamicProperty(report, "Name"));
                if (!string.Equals(name, reportName, StringComparison.OrdinalIgnoreCase))
                    continue;

                return ToBool(TryGetDynamicProperty(report, "IsLoaded"), false);
            }

            return false;
        }

        private static void CloseFormInternal(dynamic accessApp, string formName, bool saveChanges)
        {
            try
            {
                accessApp.DoCmd.Close(2, formName, saveChanges ? 1 : 2); // 2 = acForm
            }
            catch
            {
                // Ignore close failures during cleanup.
            }
        }

        private static void CloseReportInternal(dynamic accessApp, string reportName, bool saveChanges)
        {
            try
            {
                accessApp.DoCmd.Close(3, reportName, saveChanges ? 1 : 2); // 3 = acReport
            }
            catch
            {
                // Ignore close failures during cleanup.
            }
        }

        private static object? FindObjectByName(object collection, string objectName)
        {
            foreach (var item in (dynamic)collection)
            {
                var name = SafeToString(TryGetDynamicProperty(item, "Name"));
                if (string.Equals(name, objectName, StringComparison.OrdinalIgnoreCase))
                    return item;
            }

            return null;
        }

        private static object? GetControlByName(object formOrReport, string controlName)
        {
            var controlsCollection = GetControlsCollection(formOrReport);
            if (controlsCollection == null)
                return null;

            try
            {
                var byItem = InvokeDynamicMethod(controlsCollection, "Item", controlName);
                if (byItem != null)
                    return byItem;
            }
            catch
            {
                // Fall back to manual enumeration.
            }

            foreach (var control in GetControlObjects(formOrReport))
            {
                var name = SafeToString(TryGetDynamicProperty(control, "Name"));
                if (string.Equals(name, controlName, StringComparison.OrdinalIgnoreCase))
                    return control;
            }

            return null;
        }

        private static ControlInfo BuildControlInfo(object control)
        {
            return new ControlInfo
            {
                Name = SafeToString(TryGetDynamicProperty(control, "Name")) ?? "",
                Type = MapControlType(ToInt32(TryGetDynamicProperty(control, "ControlType"))),
                Left = ToInt32(TryGetDynamicProperty(control, "Left")),
                Top = ToInt32(TryGetDynamicProperty(control, "Top")),
                Width = ToInt32(TryGetDynamicProperty(control, "Width")),
                Height = ToInt32(TryGetDynamicProperty(control, "Height")),
                Visible = ToBool(TryGetDynamicProperty(control, "Visible"), true),
                Enabled = ToBool(TryGetDynamicProperty(control, "Enabled"), true)
            };
        }

        private static List<object> GetControlObjects(object formOrReport)
        {
            var controlsCollection = GetControlsCollection(formOrReport)
                ?? throw new InvalidOperationException("Controls collection is not available for this Access object.");

            var controls = new List<object>();
            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var count = ToInt32(TryGetDynamicProperty(controlsCollection, "Count"));
            for (var i = 0; i < count; i++)
            {
                var control = TryGetControlByIndex(controlsCollection, i);
                if (control == null)
                    continue;

                var key = SafeToString(TryGetDynamicProperty(control, "Name")) ?? $"index:{i}";
                if (seenNames.Add(key))
                {
                    controls.Add(control);
                }
            }

            if (controls.Count > 0)
                return controls;

            foreach (var control in (dynamic)controlsCollection)
            {
                var key = SafeToString(TryGetDynamicProperty(control, "Name")) ?? $"ref:{controls.Count}";
                if (seenNames.Add(key))
                {
                    controls.Add(control);
                }
            }

            return controls;
        }

        private static object? GetControlsCollection(object formOrReport)
        {
            var controls = TryGetDynamicProperty(formOrReport, "Controls");
            if (controls != null)
                return controls;

            try
            {
                return InvokeDynamicMethod(formOrReport, "Controls");
            }
            catch
            {
                return null;
            }
        }

        private static object? TryGetControlByIndex(object controlsCollection, int zeroBasedIndex)
        {
            try
            {
                var byZeroBased = InvokeDynamicMethod(controlsCollection, "Item", zeroBasedIndex);
                if (byZeroBased != null)
                    return byZeroBased;
            }
            catch
            {
                // Access collection may be 1-based for this object.
            }

            try
            {
                return InvokeDynamicMethod(controlsCollection, "Item", zeroBasedIndex + 1);
            }
            catch
            {
                return null;
            }
        }

        private string TryGetFormVbaCode(string formName)
        {
            var candidates = new[]
            {
                formName,
                $"Form_{formName}"
            };

            foreach (var candidate in candidates)
            {
                try
                {
                    var code = GetVBACode("CurrentProject", candidate);
                    if (!string.IsNullOrWhiteSpace(code))
                        return code;
                }
                catch
                {
                    // Continue trying alternate component naming conventions.
                }
            }

            return string.Empty;
        }

        private static string MapControlType(int controlType)
        {
            return controlType switch
            {
                100 => "Label",
                101 => "Line",
                102 => "Rectangle",
                103 => "Image",
                104 => "CommandButton",
                105 => "OptionButton",
                106 => "CheckBox",
                107 => "OptionGroup",
                108 => "BoundObjectFrame",
                109 => "TextBox",
                110 => "ListBox",
                111 => "ComboBox",
                112 => "SubForm",
                122 => "ToggleButton",
                _ => $"ControlType({controlType})"
            };
        }

        private static int MapControlTypeToConstant(string? controlType)
        {
            return (controlType ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                "label" => 100,
                "line" => 101,
                "rectangle" => 102,
                "image" => 103,
                "commandbutton" => 104,
                "optionbutton" => 105,
                "checkbox" => 106,
                "optiongroup" => 107,
                "boundobjectframe" => 108,
                "textbox" => 109,
                "listbox" => 110,
                "combobox" => 111,
                "subform" => 112,
                "togglebutton" => 122,
                _ => 109 // default to TextBox
            };
        }

        private static void TryCreateFormControl(dynamic accessApp, string formName, ControlInfo control)
        {
            try
            {
                var created = accessApp.CreateControl(
                    formName,
                    MapControlTypeToConstant(control.Type),
                    0,
                    Type.Missing,
                    Type.Missing,
                    control.Left,
                    control.Top,
                    control.Width,
                    control.Height);

                if (!string.IsNullOrWhiteSpace(control.Name))
                {
                    SetDynamicProperty(created, "Name", control.Name);
                }

                SetDynamicProperty(created, "Visible", control.Visible);
                SetDynamicProperty(created, "Enabled", control.Enabled);
            }
            catch
            {
                // Keep import best-effort when a specific control cannot be created.
            }
        }

        private static void TryCreateReportControl(dynamic accessApp, string reportName, ControlInfo control)
        {
            try
            {
                var created = accessApp.CreateReportControl(
                    reportName,
                    MapControlTypeToConstant(control.Type),
                    0,
                    Type.Missing,
                    Type.Missing,
                    control.Left,
                    control.Top,
                    control.Width,
                    control.Height);

                if (!string.IsNullOrWhiteSpace(control.Name))
                {
                    SetDynamicProperty(created, "Name", control.Name);
                }

                SetDynamicProperty(created, "Visible", control.Visible);
                SetDynamicProperty(created, "Enabled", control.Enabled);
            }
            catch
            {
                // Keep import best-effort when a specific control cannot be created.
            }
        }

        private static void TryDeleteObject(dynamic accessApp, int objectType, string objectName)
        {
            try
            {
                accessApp.DoCmd.DeleteObject(objectType, objectName);
            }
            catch
            {
                // Object may not exist; ignore.
            }
        }

        private static void TrySaveModule(dynamic accessApp, string moduleName)
        {
            try
            {
                // 5 = acModule
                accessApp.DoCmd.Save(5, moduleName);
            }
            catch
            {
                // Saving modules can fail when the object isn't active; ignore best-effort failures.
            }
        }

        private static int ToInt32(object? value)
        {
            if (value == null || value == DBNull.Value)
                return 0;

            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        private static bool ToBool(object? value, bool defaultValue)
        {
            if (value == null || value == DBNull.Value)
                return defaultValue;

            try
            {
                return value switch
                {
                    bool b => b,
                    string s when bool.TryParse(s, out var parsed) => parsed,
                    string s when int.TryParse(s, out var intValue) => intValue != 0,
                    _ => Convert.ToInt32(value) != 0
                };
            }
            catch
            {
                return defaultValue;
            }
        }

        private static string? SafeToString(object? value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            return Convert.ToString(value);
        }

        private static object? TryGetDynamicProperty(object target, string propertyName, params object?[]? args)
        {
            if (target == null)
                return null;

            try
            {
                var lateArgs = args ?? Array.Empty<object?>();
                return NewLateBinding.LateGet(
                    target,
                    null,
                    propertyName,
                    lateArgs,
                    null,
                    null,
                    null);
            }
            catch
            {
                return null;
            }
        }

        private static void SetDynamicProperty(object target, string propertyName, object? value)
        {
            if (target == null)
                throw new ArgumentNullException(nameof(target));

            NewLateBinding.LateSet(
                target,
                null,
                propertyName,
                new object?[] { value },
                null,
                null);
        }

        private static object? InvokeDynamicMethod(object target, string methodName, params object?[] args)
        {
            if (target == null)
                throw new ArgumentNullException(nameof(target));

            return NewLateBinding.LateCall(
                target,
                null,
                methodName,
                args,
                null,
                null,
                null,
                false);
        }

        private static object? ConvertValueForProperty(object value, object? existingValue)
        {
            if (value is not string raw)
                return value;

            if (existingValue == null)
            {
                if (bool.TryParse(raw, out var boolValue))
                    return boolValue;
                if (int.TryParse(raw, out var intValue))
                    return intValue;
                if (double.TryParse(raw, out var doubleValue))
                    return doubleValue;
                return raw;
            }

            var targetType = existingValue.GetType();

            if (targetType == typeof(string))
                return raw;
            if (targetType == typeof(bool))
                return ToBool(raw, false);
            if (targetType == typeof(int) || targetType == typeof(short) || targetType == typeof(long))
                return int.TryParse(raw, out var intValue) ? intValue : existingValue;
            if (targetType == typeof(double) || targetType == typeof(float) || targetType == typeof(decimal))
                return double.TryParse(raw, out var doubleValue) ? doubleValue : existingValue;

            return raw;
        }

        private static string NormalizeLineEndings(string code)
        {
            return code
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Replace("\r", "\n", StringComparison.Ordinal)
                .Replace("\n", "\r\n", StringComparison.Ordinal);
        }

        private static object? NormalizeValue(object value)
        {
            return value switch
            {
                DBNull => null,
                byte[] bytes => Convert.ToBase64String(bytes),
                _ => value
            };
        }

        private static string EscapeMarkdownCell(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            return value
                .Replace("\\", "\\\\", StringComparison.Ordinal)
                .Replace("|", "\\|", StringComparison.Ordinal)
                .Replace("\r", " ", StringComparison.Ordinal)
                .Replace("\n", "<br/>", StringComparison.Ordinal);
        }

        private static List<string> MakeUniqueColumnNames(IReadOnlyList<string> rawNames)
        {
            var result = new List<string>(rawNames.Count);
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var raw in rawNames)
            {
                var baseName = string.IsNullOrWhiteSpace(raw) ? "column" : raw;
                if (!seen.TryGetValue(baseName, out var count))
                {
                    seen[baseName] = 1;
                    result.Add(baseName);
                    continue;
                }

                count++;
                seen[baseName] = count;
                result.Add($"{baseName}_{count}");
            }

            return result;
        }

        private HashSet<string> GetPrimaryKeyColumns(string tableName)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var indexes = _oleDbConnection!.GetSchema("Indexes");
                foreach (DataRow row in indexes.Rows)
                {
                    var indexedTable = GetRowString(row, "TABLE_NAME");
                    if (!string.Equals(indexedTable, tableName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var isPrimaryKey = GetRowBool(row, "PRIMARY_KEY");
                    if (!isPrimaryKey)
                        continue;

                    var columnName = GetRowString(row, "COLUMN_NAME");
                    if (!string.IsNullOrWhiteSpace(columnName))
                    {
                        keys.Add(columnName);
                    }
                }
            }
            catch
            {
                // Primary key metadata may not be available for all providers/tables.
            }

            return keys;
        }

        private static string? GetRowString(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName))
                return null;

            var value = row[columnName];
            if (value == DBNull.Value)
                return null;

            return value.ToString();
        }

        private static int? GetRowInt(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName))
                return null;

            var value = row[columnName];
            if (value == DBNull.Value)
                return null;

            return Convert.ToInt32(value);
        }

        private static bool GetRowBool(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName))
                return false;

            var value = row[columnName];
            if (value == DBNull.Value)
                return false;

            return value switch
            {
                bool b => b,
                string s when bool.TryParse(s, out var parsed) => parsed,
                _ => Convert.ToInt32(value) != 0
            };
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                try
                {
                    CloseAccess();
                }
                catch
                {
                    // No-op during disposal cleanup.
                }

                Disconnect();
                _disposed = true;
            }
        }
    }

    #region Data Models

    public class TableInfo
    {
        public string Name { get; set; } = "";
        public List<FieldInfo> Fields { get; set; } = new();
        public long RecordCount { get; set; }
    }

    public class FieldInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Size { get; set; }
        public bool Required { get; set; }
        public bool AllowZeroLength { get; set; }
    }

    public class QueryInfo
    {
        public string Name { get; set; } = "";
        public string SQL { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class RelationshipInfo
    {
        public string Name { get; set; } = "";
        public string Table { get; set; } = "";
        public string ForeignTable { get; set; } = "";
        public string Attributes { get; set; } = "";
    }

    public class FormInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ReportInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class MacroInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ModuleInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class VBAProjectInfo
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public List<VBAModuleInfo> Modules { get; set; } = new();
    }

    public class VBAModuleInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public bool HasCode { get; set; }
    }

    public class SystemTableInfo
    {
        public string Name { get; set; } = "";
        public DateTime DateCreated { get; set; }
        public DateTime LastUpdated { get; set; }
        public long RecordCount { get; set; }
    }

    public class MetadataInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Flags { get; set; } = "";
        public string DateCreated { get; set; } = "";
        public string DateModified { get; set; } = "";
    }

    public class ControlInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
    }

    public class ControlProperties
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
        public int BackColor { get; set; }
        public int ForeColor { get; set; }
        public string FontName { get; set; } = "";
        public int FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
    }

    public class FormExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
        public string VBA { get; set; } = "";
    }

    public class ReportExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
    }

    public class SqlExecutionResult
    {
        public bool IsQuery { get; set; }
        public List<string> Columns { get; set; } = new();
        public List<Dictionary<string, object?>> Rows { get; set; } = new();
        public int RowCount { get; set; }
        public bool Truncated { get; set; }
        public int RowsAffected { get; set; } = -1;
    }

    public class TableDefinition
    {
        public string TableName { get; set; } = "";
        public List<TableColumnDefinition> Columns { get; set; } = new();
        public List<string> PrimaryKeyColumns { get; set; } = new();
    }

    public class TableColumnDefinition
    {
        public string Name { get; set; } = "";
        public string DataType { get; set; } = "";
        public int? DataTypeCode { get; set; }
        public int? OrdinalPosition { get; set; }
        public int? MaxLength { get; set; }
        public int? NumericPrecision { get; set; }
        public int? NumericScale { get; set; }
        public bool IsNullable { get; set; }
        public bool IsPrimaryKey { get; set; }
        public bool HasDefault { get; set; }
        public string? DefaultValue { get; set; }
    }

    #endregion
} 
