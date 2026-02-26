using System.Data;
using System.Data.OleDb;
using Microsoft.VisualBasic.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

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
        private const int DaoRelationAttributeDontEnforce = 2;
        private const int DaoRelationAttributeUpdateCascade = 256;
        private const int DaoRelationAttributeDeleteCascade = 4096;
        private const string TextModeJson = "json";
        private const string TextModeAccessText = "access_text";

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

            try
            {
                var comQueries = ExecuteComOperation(accessApp =>
                {
                    var list = new List<QueryInfo>();
                    var currentDb = TryGetCurrentDb(accessApp);
                    if (currentDb == null)
                        return list;

                    var queryDefs = TryGetDynamicProperty(currentDb, "QueryDefs");
                    if (queryDefs == null)
                        return list;

                    foreach (var queryDef in queryDefs)
                    {
                        var queryName = SafeToString(TryGetDynamicProperty(queryDef, "Name"));
                        if (string.IsNullOrWhiteSpace(queryName) || queryName.StartsWith("~", StringComparison.Ordinal))
                            continue;

                        var sql = SafeToString(TryGetDynamicProperty(queryDef, "SQL")) ?? string.Empty;
                        var typeCode = ToInt32(TryGetDynamicProperty(queryDef, "Type"));

                        list.Add(new QueryInfo
                        {
                            Name = queryName,
                            SQL = sql.Trim(),
                            Type = MapQueryDefType(typeCode)
                        });
                    }

                    return list;
                },
                requireExclusive: false,
                releaseOleDb: false);

                if (comQueries.Count > 0)
                    return comQueries.OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase).ToList();
            }
            catch
            {
                // Fall back to OleDb schema query when DAO QueryDefs are unavailable.
            }

            // Use OleDb to get query information
            var schema = _oleDbConnection!.GetSchema("Views");
            foreach (DataRow row in schema.Rows)
            {
                var queryName = row["TABLE_NAME"]?.ToString();
                if (string.IsNullOrWhiteSpace(queryName) || queryName.StartsWith("~", StringComparison.Ordinal))
                    continue;

                queries.Add(new QueryInfo
                {
                    Name = queryName,
                    SQL = string.Empty,
                    Type = "Query"
                });
            }

            return queries.OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        public void CreateQuery(string queryName, string sql)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("Query name is required", nameof(queryName));
            if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required", nameof(sql));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");

                if (FindQueryDef(currentDb, queryName) != null)
                    throw new InvalidOperationException($"Query already exists: {queryName}");

                _ = InvokeDynamicMethod(currentDb, "CreateQueryDef", queryName, sql);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void UpdateQuery(string queryName, string sql)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("Query name is required", nameof(queryName));
            if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required", nameof(sql));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");

                var queryDef = FindQueryDef(currentDb, queryName)
                    ?? throw new InvalidOperationException($"Query not found: {queryName}");

                SetDynamicProperty(queryDef, "SQL", sql);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void DeleteQuery(string queryName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("Query name is required", nameof(queryName));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");

                DeleteQueryInternal(currentDb, queryName);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<RelationshipInfo> GetRelationships()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var relationships = new List<RelationshipInfo>();

            try
            {
                var comRelationships = ExecuteComOperation(accessApp =>
                {
                    var list = new List<RelationshipInfo>();
                    var currentDb = TryGetCurrentDb(accessApp);
                    if (currentDb == null)
                        return list;

                    var relationCollection = TryGetDynamicProperty(currentDb, "Relations");
                    if (relationCollection == null)
                        return list;

                    foreach (var relation in relationCollection)
                    {
                        var relationName = SafeToString(TryGetDynamicProperty(relation, "Name"));
                        if (string.IsNullOrWhiteSpace(relationName) || relationName.StartsWith("~", StringComparison.Ordinal))
                            continue;

                        var fieldName = string.Empty;
                        var foreignFieldName = string.Empty;
                        var relationFields = TryGetDynamicProperty(relation, "Fields");
                        if (relationFields != null)
                        {
                            foreach (var relationField in relationFields)
                            {
                                // DAO stores primary-side column as Name and foreign-side column as ForeignName.
                                // MCP APIs expose table/field as primary side and foreign_* as dependent side.
                                fieldName = SafeToString(TryGetDynamicProperty(relationField, "Name")) ?? string.Empty;
                                foreignFieldName = SafeToString(TryGetDynamicProperty(relationField, "ForeignName")) ?? string.Empty;
                                break;
                            }
                        }

                        var attributesValue = ToInt32(TryGetDynamicProperty(relation, "Attributes"));

                        list.Add(new RelationshipInfo
                        {
                            Name = relationName,
                            Table = SafeToString(TryGetDynamicProperty(relation, "Table")) ?? string.Empty,
                            ForeignTable = SafeToString(TryGetDynamicProperty(relation, "ForeignTable")) ?? string.Empty,
                            Field = fieldName,
                            ForeignField = foreignFieldName,
                            EnforceIntegrity = !HasRelationshipAttribute(attributesValue, DaoRelationAttributeDontEnforce),
                            CascadeUpdate = HasRelationshipAttribute(attributesValue, DaoRelationAttributeUpdateCascade),
                            CascadeDelete = HasRelationshipAttribute(attributesValue, DaoRelationAttributeDeleteCascade),
                            Attributes = attributesValue.ToString()
                        });
                    }

                    return list;
                },
                requireExclusive: false,
                releaseOleDb: false);

                if (comRelationships.Count > 0)
                    return comRelationships.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase).ToList();
            }
            catch
            {
                // Fall back to OleDb metadata where DAO relationships are unavailable.
            }

            try
            {
                // Not all ACE providers expose this collection; return empty on unsupported providers.
                var schema = _oleDbConnection!.GetSchema("ForeignKeys");

                foreach (DataRow row in schema.Rows)
                {
                    relationships.Add(new RelationshipInfo
                    {
                        Name = row["FK_NAME"]?.ToString() ?? string.Empty,
                        Table = row["REFERENCED_TABLE_NAME"]?.ToString() ?? string.Empty,
                        ForeignTable = row["TABLE_NAME"]?.ToString() ?? string.Empty,
                        Field = row["PK_COLUMN_NAME"]?.ToString() ?? string.Empty,
                        ForeignField = row["FK_COLUMN_NAME"]?.ToString() ?? string.Empty,
                        EnforceIntegrity = true,
                        CascadeUpdate = false,
                        CascadeDelete = false,
                        Attributes = string.Empty
                    });
                }
            }
            catch
            {
                // Keep compatibility with providers that do not publish ForeignKeys metadata.
            }

            return relationships.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase).ToList();
        }

        public string CreateRelationship(
            string tableName,
            string fieldName,
            string foreignTableName,
            string foreignFieldName,
            string? relationshipName = null,
            bool enforceIntegrity = true,
            bool cascadeUpdate = false,
            bool cascadeDelete = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required", nameof(fieldName));
            if (string.IsNullOrWhiteSpace(foreignTableName)) throw new ArgumentException("Foreign table name is required", nameof(foreignTableName));
            if (string.IsNullOrWhiteSpace(foreignFieldName)) throw new ArgumentException("Foreign field name is required", nameof(foreignFieldName));

            var effectiveRelationshipName = string.IsNullOrWhiteSpace(relationshipName)
                ? BuildRelationshipName(tableName, fieldName, foreignTableName, foreignFieldName)
                : relationshipName;

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");

                CreateRelationshipInternal(
                    currentDb,
                    effectiveRelationshipName!,
                    tableName,
                    fieldName,
                    foreignTableName,
                    foreignFieldName,
                    enforceIntegrity,
                    cascadeUpdate,
                    cascadeDelete);
            },
            requireExclusive: false,
            releaseOleDb: false);

            return effectiveRelationshipName!;
        }

        public string UpdateRelationship(
            string relationshipName,
            string tableName,
            string fieldName,
            string foreignTableName,
            string foreignFieldName,
            bool enforceIntegrity = true,
            bool cascadeUpdate = false,
            bool cascadeDelete = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(relationshipName)) throw new ArgumentException("Relationship name is required", nameof(relationshipName));
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required", nameof(fieldName));
            if (string.IsNullOrWhiteSpace(foreignTableName)) throw new ArgumentException("Foreign table name is required", nameof(foreignTableName));
            if (string.IsNullOrWhiteSpace(foreignFieldName)) throw new ArgumentException("Foreign field name is required", nameof(foreignFieldName));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");

                if (!DeleteRelationshipInternal(currentDb, relationshipName))
                    throw new InvalidOperationException($"Relationship not found: {relationshipName}");

                CreateRelationshipInternal(
                    currentDb,
                    relationshipName,
                    tableName,
                    fieldName,
                    foreignTableName,
                    foreignFieldName,
                    enforceIntegrity,
                    cascadeUpdate,
                    cascadeDelete);
            },
            requireExclusive: false,
            releaseOleDb: false);

            return relationshipName;
        }

        public void DeleteRelationship(string relationshipName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(relationshipName)) throw new ArgumentException("Relationship name is required", nameof(relationshipName));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");

                if (!DeleteRelationshipInternal(currentDb, relationshipName))
                    throw new InvalidOperationException($"Relationship not found: {relationshipName}");
            },
            requireExclusive: false,
            releaseOleDb: false);
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

        public void AddField(string tableName, FieldInfo field)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (field == null) throw new ArgumentNullException(nameof(field));

            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            var normalizedFieldName = NormalizeSchemaIdentifier(field.Name, nameof(field), "Field name is required");
            var typeDeclaration = BuildAccessDataTypeDeclaration(field.Type, field.Size, nameof(field.Type), nameof(field.Size));

            EnsureTableExists(normalizedTableName);
            if (FieldExists(normalizedTableName, normalizedFieldName))
                throw new InvalidOperationException($"Field already exists: {normalizedTableName}.{normalizedFieldName}");

            if (typeDeclaration == "COUNTER" && field.Required)
                throw new ArgumentException("COUNTER fields cannot be explicitly marked as required.", nameof(field));

            var notNullClause = field.Required ? " NOT NULL" : string.Empty;
            var sql = $"ALTER TABLE [{EscapeSqlIdentifier(normalizedTableName)}] ADD COLUMN [{EscapeSqlIdentifier(normalizedFieldName)}] {typeDeclaration}{notNullClause}";
            ExecuteSchemaNonQuery(sql);
        }

        public void AlterField(string tableName, string fieldName, string newType, int size = 0)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            var normalizedFieldName = NormalizeSchemaIdentifier(fieldName, nameof(fieldName), "Field name is required");
            var typeDeclaration = BuildAccessDataTypeDeclaration(newType, size, nameof(newType), nameof(size));

            EnsureTableExists(normalizedTableName);
            if (!FieldExists(normalizedTableName, normalizedFieldName))
                throw new InvalidOperationException($"Field not found: {normalizedTableName}.{normalizedFieldName}");

            if (typeDeclaration == "COUNTER")
                throw new ArgumentException("Altering a field to COUNTER is not supported by Access DDL.", nameof(newType));

            var sql = $"ALTER TABLE [{EscapeSqlIdentifier(normalizedTableName)}] ALTER COLUMN [{EscapeSqlIdentifier(normalizedFieldName)}] {typeDeclaration}";
            ExecuteSchemaNonQuery(sql);
        }

        public void DropField(string tableName, string fieldName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            var normalizedFieldName = NormalizeSchemaIdentifier(fieldName, nameof(fieldName), "Field name is required");

            EnsureTableExists(normalizedTableName);
            if (!FieldExists(normalizedTableName, normalizedFieldName))
                throw new InvalidOperationException($"Field not found: {normalizedTableName}.{normalizedFieldName}");

            var sql = $"ALTER TABLE [{EscapeSqlIdentifier(normalizedTableName)}] DROP COLUMN [{EscapeSqlIdentifier(normalizedFieldName)}]";
            ExecuteSchemaNonQuery(sql);
        }

        public void RenameTable(string tableName, string newTableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            var normalizedNewTableName = NormalizeSchemaIdentifier(newTableName, nameof(newTableName), "New table name is required");
            if (string.Equals(normalizedTableName, normalizedNewTableName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("New table name must be different from the existing table name.", nameof(newTableName));

            EnsureTableExists(normalizedTableName);
            if (TableExists(normalizedNewTableName))
                throw new InvalidOperationException($"Table already exists: {normalizedNewTableName}");

            try
            {
                ExecuteComOperation(accessApp =>
                {
                    var tableDef = FindTableDefWithRetry(accessApp, normalizedTableName);
                    if (tableDef == null)
                        throw new InvalidOperationException($"Table not found: {normalizedTableName}");

                    SetDynamicProperty(tableDef, "Name", normalizedNewTableName);
                },
                requireExclusive: true,
                releaseOleDb: true);
            }
            catch (Exception ex) when (ShouldUseOleDbRenameFallback(ex))
            {
                RenameTableViaOleDbCopy(normalizedTableName, normalizedNewTableName);
            }
        }

        public void RenameField(string tableName, string fieldName, string newFieldName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            var normalizedFieldName = NormalizeSchemaIdentifier(fieldName, nameof(fieldName), "Field name is required");
            var normalizedNewFieldName = NormalizeSchemaIdentifier(newFieldName, nameof(newFieldName), "New field name is required");
            if (string.Equals(normalizedFieldName, normalizedNewFieldName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("New field name must be different from the existing field name.", nameof(newFieldName));

            EnsureTableExists(normalizedTableName);
            if (!FieldExists(normalizedTableName, normalizedFieldName))
                throw new InvalidOperationException($"Field not found: {normalizedTableName}.{normalizedFieldName}");
            if (FieldExists(normalizedTableName, normalizedNewFieldName))
                throw new InvalidOperationException($"Field already exists: {normalizedTableName}.{normalizedNewFieldName}");

            try
            {
                ExecuteComOperation(accessApp =>
                {
                    var tableDef = FindTableDefWithRetry(accessApp, normalizedTableName);
                    if (tableDef == null)
                        throw new InvalidOperationException($"Table not found: {normalizedTableName}");

                    var sourceField = FindTableField(tableDef, normalizedFieldName)
                        ?? throw new InvalidOperationException($"Field not found: {normalizedTableName}.{normalizedFieldName}");

                    SetDynamicProperty(sourceField, "Name", normalizedNewFieldName);
                },
                requireExclusive: true,
                releaseOleDb: true);
            }
            catch (Exception ex) when (ShouldUseOleDbRenameFallback(ex))
            {
                RenameFieldViaOleDbCopy(normalizedTableName, normalizedFieldName, normalizedNewFieldName);
            }
        }

        public List<IndexInfo> GetIndexes(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            EnsureOleDbConnection();

            var indexesByName = new Dictionary<string, IndexInfo>(StringComparer.OrdinalIgnoreCase);
            var indexColumns = new Dictionary<string, List<(int Ordinal, string Column)>>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var schema = _oleDbConnection!.GetSchema("Indexes");
                foreach (DataRow row in schema.Rows)
                {
                    var indexedTable = GetRowString(row, "TABLE_NAME");
                    if (!string.Equals(indexedTable, tableName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var indexName = GetRowString(row, "INDEX_NAME");
                    if (string.IsNullOrWhiteSpace(indexName) || indexName.StartsWith("~", StringComparison.Ordinal))
                        continue;

                    if (!indexesByName.TryGetValue(indexName, out var index))
                    {
                        index = new IndexInfo
                        {
                            Name = indexName,
                            Table = tableName,
                            IsUnique = GetRowBool(row, "UNIQUE"),
                            IsPrimaryKey = GetRowBool(row, "PRIMARY_KEY")
                        };
                        indexesByName[indexName] = index;
                        indexColumns[indexName] = new List<(int Ordinal, string Column)>();
                    }

                    var columnName = GetRowString(row, "COLUMN_NAME");
                    if (string.IsNullOrWhiteSpace(columnName))
                        continue;

                    var ordinal = GetRowInt(row, "ORDINAL_POSITION") ?? int.MaxValue;
                    indexColumns[indexName].Add((ordinal, columnName));
                }
            }
            catch
            {
                // Index metadata is provider-dependent; return what is available.
            }

            foreach (var kvp in indexesByName)
            {
                var columns = indexColumns[kvp.Key]
                    .OrderBy(c => c.Ordinal)
                    .ThenBy(c => c.Column, StringComparer.OrdinalIgnoreCase)
                    .Select(c => c.Column)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                kvp.Value.Columns = columns;
            }

            return indexesByName.Values
                .OrderBy(i => i.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public void CreateIndex(string tableName, string indexName, List<string> columns, bool unique = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            if (string.IsNullOrWhiteSpace(indexName)) throw new ArgumentException("Index name is required", nameof(indexName));
            if (columns == null || columns.Count == 0) throw new ArgumentException("At least one column is required", nameof(columns));
            EnsureOleDbConnection();

            var normalizedColumns = columns
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (normalizedColumns.Count == 0)
                throw new ArgumentException("At least one non-empty column is required", nameof(columns));

            var uniqueSql = unique ? "UNIQUE " : string.Empty;
            var columnSql = string.Join(", ", normalizedColumns.Select(c => $"[{EscapeSqlIdentifier(c)}]"));
            var sql = $"CREATE {uniqueSql}INDEX [{EscapeSqlIdentifier(indexName)}] ON [{EscapeSqlIdentifier(tableName)}] ({columnSql})";
            using var command = new OleDbCommand(sql, _oleDbConnection);
            command.ExecuteNonQuery();
        }

        public void DeleteIndex(string tableName, string indexName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            if (string.IsNullOrWhiteSpace(indexName)) throw new ArgumentException("Index name is required", nameof(indexName));
            EnsureOleDbConnection();

            var sql = $"DROP INDEX [{EscapeSqlIdentifier(indexName)}] ON [{EscapeSqlIdentifier(tableName)}]";
            using var command = new OleDbCommand(sql, _oleDbConnection);
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

        public string ExportMacroToText(string macroName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required", nameof(macroName));

            return ExecuteComOperation(accessApp =>
            {
                var tempPath = BuildTemporaryTextPath("macro_export");
                try
                {
                    accessApp.SaveAsText(4, macroName, tempPath); // 4 = acMacro
                    return File.ReadAllText(tempPath, Encoding.UTF8);
                }
                finally
                {
                    TryDeleteFile(tempPath);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void ImportMacroFromText(string macroName, string macroData, bool overwrite = true)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required", nameof(macroName));
            if (string.IsNullOrWhiteSpace(macroData)) throw new ArgumentException("Macro data is required", nameof(macroData));

            ExecuteComOperation(accessApp =>
            {
                if (!overwrite && MacroExists(accessApp, macroName))
                    throw new InvalidOperationException($"Macro already exists: {macroName}");

                if (overwrite)
                {
                    TryDeleteObject(accessApp, 4, macroName); // 4 = acMacro
                }

                var tempPath = BuildTemporaryTextPath("macro_import");
                try
                {
                    File.WriteAllText(tempPath, macroData, Encoding.UTF8);
                    accessApp.LoadFromText(4, macroName, tempPath); // 4 = acMacro
                }
                finally
                {
                    TryDeleteFile(tempPath);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void CreateMacro(string macroName, string macroData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required", nameof(macroName));
            if (string.IsNullOrWhiteSpace(macroData)) throw new ArgumentException("Macro data is required", nameof(macroData));

            ExecuteComOperation(accessApp =>
            {
                if (MacroExists(accessApp, macroName))
                    throw new InvalidOperationException($"Macro already exists: {macroName}");

                var tempPath = BuildTemporaryTextPath("macro_create");
                try
                {
                    File.WriteAllText(tempPath, macroData, Encoding.UTF8);
                    accessApp.LoadFromText(4, macroName, tempPath); // 4 = acMacro
                }
                finally
                {
                    TryDeleteFile(tempPath);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void UpdateMacro(string macroName, string macroData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required", nameof(macroName));
            if (string.IsNullOrWhiteSpace(macroData)) throw new ArgumentException("Macro data is required", nameof(macroData));

            ExecuteComOperation(accessApp =>
            {
                if (!MacroExists(accessApp, macroName))
                    throw new InvalidOperationException($"Macro not found: {macroName}");

                var tempPath = BuildTemporaryTextPath("macro_update");
                try
                {
                    File.WriteAllText(tempPath, macroData, Encoding.UTF8);
                    accessApp.LoadFromText(4, macroName, tempPath); // 4 = acMacro
                }
                finally
                {
                    TryDeleteFile(tempPath);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void RunMacro(string macroName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required", nameof(macroName));

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.RunMacro(macroName),
                requireExclusive: false,
                releaseOleDb: false);
        }

        public void DeleteMacro(string macroName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required", nameof(macroName));

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.DeleteObject(4, macroName), // 4 = acMacro
                requireExclusive: true,
                releaseOleDb: true);
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

        public void OpenReport(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required", nameof(reportName));

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.OpenReport(reportName, 1), // 1 = acViewDesign
                requireExclusive: false,
                releaseOleDb: false);
        }

        public void CloseReport(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required", nameof(reportName));

            ExecuteComOperation(
                accessApp => accessApp.DoCmd.Close(3, reportName),
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

        public List<ControlInfo> GetReportControls(string reportName)
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

        public ControlProperties GetReportControlProperties(string reportName, string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required", nameof(reportName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("Control name is required", nameof(controlName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    var control = GetControlByName(report, controlName)
                        ?? throw new InvalidOperationException($"Control '{controlName}' was not found on report '{reportName}'.");

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
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
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

        public void SetReportControlProperty(string reportName, string controlName, string propertyName, object value)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required", nameof(reportName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("Control name is required", nameof(controlName));
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property name is required", nameof(propertyName));

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                try
                {
                    var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                    var control = GetControlByName(report, controlName)
                        ?? throw new InvalidOperationException($"Control '{controlName}' was not found on report '{reportName}'.");

                    var existingValue = TryGetDynamicProperty(control, propertyName);
                    var convertedValue = ConvertValueForProperty(value, existingValue);
                    SetDynamicProperty(control, propertyName, convertedValue);
                    accessApp.DoCmd.Save(3, reportName);
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

        #endregion

        #region 7. Persistence & Versioning

        public string ExportFormToText(string formName, string? mode = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required", nameof(formName));

            var normalizedMode = NormalizeTextTransferMode(mode);
            if (normalizedMode == TextModeAccessText)
            {
                return ExecuteComOperation(accessApp =>
                {
                    var tempPath = BuildTemporaryTextPath("form_export");
                    try
                    {
                        accessApp.SaveAsText(2, formName, tempPath); // 2 = acForm
                        return File.ReadAllText(tempPath, Encoding.UTF8);
                    }
                    finally
                    {
                        TryDeleteFile(tempPath);
                    }
                },
                requireExclusive: true,
                releaseOleDb: true);
            }

            var formData = new
            {
                Name = formName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetFormControls(formName),
                VBA = TryGetFormVbaCode(formName)
            };

            return JsonSerializer.Serialize(formData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportFormFromText(string formData, string? mode = null, string? formName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formData)) throw new ArgumentException("Form data is required", nameof(formData));

            var normalizedMode = NormalizeTextTransferMode(mode);
            if (normalizedMode == TextModeAccessText)
            {
                var resolvedFormName = ResolveAccessTextImportObjectName(formName, formData, "form", "Form_");

                ExecuteComOperation(accessApp =>
                {
                    TryDeleteObject(accessApp, 2, resolvedFormName); // 2 = acForm

                    var tempPath = BuildTemporaryTextPath("form_import");
                    try
                    {
                        File.WriteAllText(tempPath, formData, Encoding.UTF8);
                        accessApp.LoadFromText(2, resolvedFormName, tempPath); // 2 = acForm
                    }
                    finally
                    {
                        TryDeleteFile(tempPath);
                    }
                },
                requireExclusive: true,
                releaseOleDb: true);

                return;
            }

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

        public string ExportReportToText(string reportName, string? mode = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required", nameof(reportName));

            var normalizedMode = NormalizeTextTransferMode(mode);
            if (normalizedMode == TextModeAccessText)
            {
                return ExecuteComOperation(accessApp =>
                {
                    var tempPath = BuildTemporaryTextPath("report_export");
                    try
                    {
                        accessApp.SaveAsText(3, reportName, tempPath); // 3 = acReport
                        return File.ReadAllText(tempPath, Encoding.UTF8);
                    }
                    finally
                    {
                        TryDeleteFile(tempPath);
                    }
                },
                requireExclusive: true,
                releaseOleDb: true);
            }

            var reportData = new
            {
                Name = reportName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetReportControls(reportName)
            };

            return JsonSerializer.Serialize(reportData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportReportFromText(string reportData, string? mode = null, string? reportName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportData)) throw new ArgumentException("Report data is required", nameof(reportData));

            var normalizedMode = NormalizeTextTransferMode(mode);
            if (normalizedMode == TextModeAccessText)
            {
                var resolvedReportName = ResolveAccessTextImportObjectName(reportName, reportData, "report", "Report_");

                ExecuteComOperation(accessApp =>
                {
                    TryDeleteObject(accessApp, 3, resolvedReportName); // 3 = acReport

                    var tempPath = BuildTemporaryTextPath("report_import");
                    try
                    {
                        File.WriteAllText(tempPath, reportData, Encoding.UTF8);
                        accessApp.LoadFromText(3, resolvedReportName, tempPath); // 3 = acReport
                    }
                    finally
                    {
                        TryDeleteFile(tempPath);
                    }
                },
                requireExclusive: true,
                releaseOleDb: true);

                return;
            }

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

        private void ExecuteSchemaNonQuery(string sql)
        {
            Exception? lastRecoverableError = null;

            for (var attempt = 0; attempt < 2; attempt++)
            {
                EnsureOleDbConnection();

                try
                {
                    using var command = new OleDbCommand(sql, _oleDbConnection);
                    command.ExecuteNonQuery();
                    RefreshOleDbConnectionAfterSchemaMutation();
                    return;
                }
                catch (Exception ex) when (attempt == 0 && IsRecoverableOleDbLockError(ex) && TryReleaseExclusiveAccessLock())
                {
                    lastRecoverableError = ex;
                    _oleDbConnection?.Close();
                    _oleDbConnection?.Dispose();
                    _oleDbConnection = null;
                }
            }

            throw lastRecoverableError ?? new InvalidOperationException("Failed to execute schema command.");
        }

        private void RefreshOleDbConnectionAfterSchemaMutation()
        {
            if (string.IsNullOrWhiteSpace(_currentDatabasePath))
                return;

            try
            {
                OpenOleDbConnection(_currentDatabasePath);
            }
            catch
            {
                // Defer refresh to the next operation when immediate reopen is unavailable.
                _oleDbConnection?.Close();
                _oleDbConnection?.Dispose();
                _oleDbConnection = null;
            }
        }

        private void EnsureTableExists(string tableName)
        {
            if (!TableExists(tableName))
                throw new InvalidOperationException($"Table not found: {tableName}");
        }

        private bool TableExists(string tableName)
        {
            EnsureOleDbConnection();

            var schema = _oleDbConnection!.GetSchema("Tables");
            foreach (DataRow row in schema.Rows)
            {
                var currentName = GetRowString(row, "TABLE_NAME");
                if (!string.Equals(currentName, tableName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var tableType = GetRowString(row, "TABLE_TYPE");
                if (string.IsNullOrWhiteSpace(tableType))
                    return true;

                if (tableType.IndexOf("TABLE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    tableType.IndexOf("LINK", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private bool FieldExists(string tableName, string fieldName)
        {
            EnsureOleDbConnection();

            var schema = _oleDbConnection!.GetSchema("Columns", new string[] { null!, null!, tableName, null! });
            foreach (DataRow row in schema.Rows)
            {
                var currentName = GetRowString(row, "COLUMN_NAME");
                if (string.Equals(currentName, fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private static dynamic? FindTableDef(dynamic currentDb, string tableName)
        {
            var tableDefs = TryGetDynamicProperty(currentDb, "TableDefs");
            if (tableDefs == null)
                return null;

            try
            {
                var item = InvokeDynamicMethod(tableDefs, "Item", tableName);
                if (item != null)
                    return item;
            }
            catch
            {
                // Fall back to enumeration when direct keyed lookup is unavailable.
            }

            foreach (var tableDef in tableDefs)
            {
                var currentName = SafeToString(TryGetDynamicProperty(tableDef, "Name"));
                if (string.Equals(currentName, tableName, StringComparison.OrdinalIgnoreCase))
                    return tableDef;
            }

            return null;
        }

        private dynamic? FindTableDefWithRetry(dynamic accessApp, string tableName)
        {
            for (var attempt = 0; attempt < 3; attempt++)
            {
                var currentDb = TryGetCurrentDb(accessApp);
                if (currentDb != null)
                {
                    TryRefreshTableDefs(currentDb);
                    var tableDef = FindTableDef(currentDb, tableName);
                    if (tableDef != null)
                        return tableDef;
                }

                if (attempt == 0)
                {
                    ReopenCurrentDatabase(accessApp);
                }
                else
                {
                    System.Threading.Thread.Sleep(100);
                }
            }

            return null;
        }

        private static void TryRefreshTableDefs(dynamic currentDb)
        {
            try
            {
                var tableDefs = TryGetDynamicProperty(currentDb, "TableDefs");
                if (tableDefs != null)
                {
                    _ = InvokeDynamicMethod(tableDefs, "Refresh");
                }
            }
            catch
            {
                // Best-effort metadata refresh.
            }
        }

        private void ReopenCurrentDatabase(dynamic accessApp)
        {
            if (string.IsNullOrWhiteSpace(_currentDatabasePath))
                return;

            try
            {
                accessApp.CloseCurrentDatabase();
            }
            catch
            {
                // Continue and attempt reopen.
            }

            accessApp.OpenCurrentDatabase(_currentDatabasePath, false);
            _accessDatabasePath = _currentDatabasePath;
            _accessDatabaseOpenedExclusive = false;
        }

        private static dynamic? FindTableField(dynamic tableDef, string fieldName)
        {
            var fields = TryGetDynamicProperty(tableDef, "Fields");
            if (fields == null)
                return null;

            try
            {
                var item = InvokeDynamicMethod(fields, "Item", fieldName);
                if (item != null)
                    return item;
            }
            catch
            {
                // Fall back to enumeration when direct keyed lookup is unavailable.
            }

            foreach (var field in fields)
            {
                var currentName = SafeToString(TryGetDynamicProperty(field, "Name"));
                if (string.Equals(currentName, fieldName, StringComparison.OrdinalIgnoreCase))
                    return field;
            }

            return null;
        }

        private static string NormalizeSchemaIdentifier(string identifier, string paramName, string requiredMessage)
        {
            if (string.IsNullOrWhiteSpace(identifier))
                throw new ArgumentException(requiredMessage, paramName);

            var normalized = identifier.Trim();
            if (normalized.Length > 64)
                throw new ArgumentException("Access object names must be 64 characters or fewer.", paramName);

            return normalized;
        }

        private static string BuildAccessDataTypeDeclaration(string typeName, int size, string typeParamName, string sizeParamName)
        {
            if (string.IsNullOrWhiteSpace(typeName))
                throw new ArgumentException("Field type is required.", typeParamName);

            var normalized = typeName.Trim().ToLowerInvariant();
            return normalized switch
            {
                "text" or "char" or "varchar" or "string" => $"TEXT({ValidateSizedType(size, 1, 255, 255, sizeParamName, "TEXT")})",
                "memo" or "longtext" or "note" => ValidateUnsizedType(size, sizeParamName, "LONGTEXT"),
                "byte" => ValidateUnsizedType(size, sizeParamName, "BYTE"),
                "short" or "smallint" => ValidateUnsizedType(size, sizeParamName, "SHORT"),
                "long" or "integer" or "int" => ValidateUnsizedType(size, sizeParamName, "INTEGER"),
                "single" or "float" => ValidateUnsizedType(size, sizeParamName, "SINGLE"),
                "double" or "real" => ValidateUnsizedType(size, sizeParamName, "DOUBLE"),
                "decimal" or "numeric" => ValidateUnsizedType(size, sizeParamName, "DECIMAL"),
                "currency" or "money" => ValidateUnsizedType(size, sizeParamName, "CURRENCY"),
                "datetime" or "date" or "time" => ValidateUnsizedType(size, sizeParamName, "DATETIME"),
                "yesno" or "boolean" or "bool" or "bit" => ValidateUnsizedType(size, sizeParamName, "YESNO"),
                "guid" or "uniqueidentifier" => ValidateUnsizedType(size, sizeParamName, "GUID"),
                "counter" or "autoincrement" or "identity" => ValidateUnsizedType(size, sizeParamName, "COUNTER"),
                "binary" => $"BINARY({ValidateSizedType(size, 1, 510, 255, sizeParamName, "BINARY")})",
                "varbinary" => $"VARBINARY({ValidateSizedType(size, 1, 510, 255, sizeParamName, "VARBINARY")})",
                _ => throw new ArgumentException($"Unsupported Access field type: {typeName}", typeParamName)
            };
        }

        private void RenameTableViaOleDbCopy(string sourceTableName, string targetTableName)
        {
            var sourceIndexes = CaptureIndexSnapshots(sourceTableName);
            var sourceRelationships = CaptureForeignKeySnapshots(snapshot =>
                string.Equals(snapshot.PrimaryTable, sourceTableName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(snapshot.ForeignTable, sourceTableName, StringComparison.OrdinalIgnoreCase));

            var escapedSource = EscapeSqlIdentifier(sourceTableName);
            var escapedTarget = EscapeSqlIdentifier(targetTableName);

            ExecuteSchemaNonQuery($"SELECT * INTO [{escapedTarget}] FROM [{escapedSource}]");

            try
            {
                DropForeignKeyConstraints(sourceRelationships);
                ExecuteSchemaNonQuery($"DROP TABLE [{escapedSource}]");
            }
            catch
            {
                try
                {
                    ExecuteSchemaNonQuery($"DROP TABLE [{escapedTarget}]");
                }
                catch
                {
                    // Ignore cleanup failures and surface the original drop error.
                }

                throw;
            }

            RecreateIndexes(targetTableName, sourceIndexes);
            RecreateForeignKeyConstraints(
                sourceRelationships,
                sourceTableName: sourceTableName,
                targetTableName: targetTableName);
        }

        private void RenameFieldViaOleDbCopy(string tableName, string sourceFieldName, string targetFieldName)
        {
            var tableDefinition = DescribeTable(tableName);
            var sourceColumn = tableDefinition.Columns
                .FirstOrDefault(column => string.Equals(column.Name, sourceFieldName, StringComparison.OrdinalIgnoreCase));

            if (sourceColumn == null)
                throw new InvalidOperationException($"Field not found: {tableName}.{sourceFieldName}");

            var affectedIndexes = CaptureIndexSnapshots(
                tableName,
                index => index.Columns.Any(column => string.Equals(column, sourceFieldName, StringComparison.OrdinalIgnoreCase)));
            var affectedRelationships = CaptureForeignKeySnapshots(snapshot =>
                (string.Equals(snapshot.PrimaryTable, tableName, StringComparison.OrdinalIgnoreCase) &&
                 snapshot.PrimaryColumns.Any(column => string.Equals(column, sourceFieldName, StringComparison.OrdinalIgnoreCase))) ||
                (string.Equals(snapshot.ForeignTable, tableName, StringComparison.OrdinalIgnoreCase) &&
                 snapshot.ForeignColumns.Any(column => string.Equals(column, sourceFieldName, StringComparison.OrdinalIgnoreCase))));

            var escapedTableName = EscapeSqlIdentifier(tableName);
            var escapedSourceFieldName = EscapeSqlIdentifier(sourceFieldName);
            var escapedTargetFieldName = EscapeSqlIdentifier(targetFieldName);
            var typeDeclaration = BuildAccessDataTypeDeclarationFromColumn(sourceColumn);

            DropForeignKeyConstraints(affectedRelationships);
            DropIndexes(tableName, affectedIndexes);

            ExecuteSchemaNonQuery($"ALTER TABLE [{escapedTableName}] ADD COLUMN [{escapedTargetFieldName}] {typeDeclaration}");
            ExecuteSchemaNonQuery($"UPDATE [{escapedTableName}] SET [{escapedTargetFieldName}] = [{escapedSourceFieldName}]");
            ExecuteSchemaNonQuery($"ALTER TABLE [{escapedTableName}] DROP COLUMN [{escapedSourceFieldName}]");

            RecreateIndexes(
                tableName,
                affectedIndexes,
                sourceFieldName: sourceFieldName,
                targetFieldName: targetFieldName);
            RecreateForeignKeyConstraints(
                affectedRelationships,
                fieldRenameTableName: tableName,
                sourceFieldName: sourceFieldName,
                targetFieldName: targetFieldName);
        }

        private List<IndexSnapshot> CaptureIndexSnapshots(string tableName, Func<IndexInfo, bool>? predicate = null)
        {
            var indexInfos = GetIndexes(tableName);
            if (predicate != null)
            {
                indexInfos = indexInfos.Where(predicate).ToList();
            }

            return indexInfos
                .Where(index => !string.IsNullOrWhiteSpace(index.Name))
                .Select(index => new IndexSnapshot
                {
                    Name = index.Name,
                    IsUnique = index.IsUnique,
                    IsPrimaryKey = index.IsPrimaryKey,
                    Columns = index.Columns
                        .Where(column => !string.IsNullOrWhiteSpace(column))
                        .Select(column => column.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList()
                })
                .Where(index => index.Columns.Count > 0)
                .OrderBy(index => index.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<ForeignKeySnapshot> CaptureForeignKeySnapshots(Func<ForeignKeySnapshot, bool>? predicate = null)
        {
            EnsureOleDbConnection();

            var foreignKeyBuilders = new Dictionary<string, ForeignKeySnapshotBuilder>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var schema = _oleDbConnection!.GetSchema("ForeignKeys");
                foreach (DataRow row in schema.Rows)
                {
                    var relationshipName = GetRowStringByCandidates(row, "FK_NAME", "CONSTRAINT_NAME", "RELATIONSHIP_NAME");
                    if (string.IsNullOrWhiteSpace(relationshipName) || relationshipName.StartsWith("~", StringComparison.Ordinal))
                        continue;

                    var primaryTable = GetRowStringByCandidates(row, "PK_TABLE_NAME", "REFERENCED_TABLE_NAME");
                    var foreignTable = GetRowStringByCandidates(row, "FK_TABLE_NAME", "TABLE_NAME");
                    var primaryColumn = GetRowStringByCandidates(row, "PK_COLUMN_NAME", "REFERENCED_COLUMN_NAME");
                    var foreignColumn = GetRowStringByCandidates(row, "FK_COLUMN_NAME", "COLUMN_NAME");

                    if (string.IsNullOrWhiteSpace(primaryTable) ||
                        string.IsNullOrWhiteSpace(foreignTable) ||
                        string.IsNullOrWhiteSpace(primaryColumn) ||
                        string.IsNullOrWhiteSpace(foreignColumn))
                    {
                        continue;
                    }

                    var ordinal = GetRowIntByCandidates(row, "ORDINAL", "KEY_SEQ", "ORDINAL_POSITION") ?? int.MaxValue;
                    var updateRule = GetRowIntByCandidates(row, "UPDATE_RULE");
                    var deleteRule = GetRowIntByCandidates(row, "DELETE_RULE");

                    var dictionaryKey = $"{relationshipName}\u001F{primaryTable}\u001F{foreignTable}";
                    if (!foreignKeyBuilders.TryGetValue(dictionaryKey, out var builder))
                    {
                        builder = new ForeignKeySnapshotBuilder
                        {
                            Name = relationshipName,
                            PrimaryTable = primaryTable,
                            ForeignTable = foreignTable,
                            UpdateRule = updateRule,
                            DeleteRule = deleteRule
                        };
                        foreignKeyBuilders[dictionaryKey] = builder;
                    }

                    builder.Columns.Add((ordinal, primaryColumn, foreignColumn));
                }
            }
            catch
            {
                // Foreign key metadata is provider-dependent.
                return new List<ForeignKeySnapshot>();
            }

            var snapshots = new List<ForeignKeySnapshot>();
            foreach (var builder in foreignKeyBuilders.Values)
            {
                var orderedColumns = builder.Columns
                    .OrderBy(column => column.Ordinal)
                    .ThenBy(column => column.PrimaryColumn, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(column => column.ForeignColumn, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var primaryColumns = orderedColumns
                    .Select(column => column.PrimaryColumn)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var foreignColumns = orderedColumns
                    .Select(column => column.ForeignColumn)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (primaryColumns.Count == 0 || primaryColumns.Count != foreignColumns.Count)
                    continue;

                var snapshot = new ForeignKeySnapshot
                {
                    Name = builder.Name,
                    PrimaryTable = builder.PrimaryTable,
                    ForeignTable = builder.ForeignTable,
                    PrimaryColumns = primaryColumns,
                    ForeignColumns = foreignColumns,
                    CascadeUpdate = builder.UpdateRule == 0,
                    CascadeDelete = builder.DeleteRule == 0
                };

                if (predicate == null || predicate(snapshot))
                {
                    snapshots.Add(snapshot);
                }
            }

            return snapshots
                .OrderBy(snapshot => snapshot.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(snapshot => snapshot.ForeignTable, StringComparer.OrdinalIgnoreCase)
                .ThenBy(snapshot => snapshot.PrimaryTable, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void DropIndexes(string tableName, IReadOnlyList<IndexSnapshot> indexes)
        {
            if (indexes.Count == 0)
                return;

            var escapedTableName = EscapeSqlIdentifier(tableName);
            foreach (var index in indexes
                .OrderBy(index => index.IsPrimaryKey)
                .ThenBy(index => index.Name, StringComparer.OrdinalIgnoreCase))
            {
                var escapedIndexName = EscapeSqlIdentifier(index.Name);
                if (index.IsPrimaryKey)
                {
                    ExecuteSchemaNonQuery($"ALTER TABLE [{escapedTableName}] DROP CONSTRAINT [{escapedIndexName}]");
                    continue;
                }

                ExecuteSchemaNonQuery($"DROP INDEX [{escapedIndexName}] ON [{escapedTableName}]");
            }
        }

        private void RecreateIndexes(
            string tableName,
            IReadOnlyList<IndexSnapshot> indexes,
            string? sourceFieldName = null,
            string? targetFieldName = null)
        {
            if (indexes.Count == 0)
                return;

            var existingIndexes = CaptureIndexSnapshots(tableName);
            var existingIndexNames = new HashSet<string>(existingIndexes.Select(index => index.Name), StringComparer.OrdinalIgnoreCase);
            var hasPrimaryKey = existingIndexes.Any(index => index.IsPrimaryKey);
            var tableColumns = GetTableColumnSet(tableName);
            var escapedTableName = EscapeSqlIdentifier(tableName);

            foreach (var index in indexes
                .OrderByDescending(snapshot => snapshot.IsPrimaryKey)
                .ThenBy(snapshot => snapshot.Name, StringComparer.OrdinalIgnoreCase))
            {
                if (existingIndexNames.Contains(index.Name))
                    continue;

                var mappedColumns = RemapColumns(index.Columns, sourceFieldName, targetFieldName);
                if (mappedColumns.Count == 0 || mappedColumns.Any(column => !tableColumns.Contains(column)))
                    continue;

                var escapedIndexName = EscapeSqlIdentifier(index.Name);
                var columnSql = BuildColumnListSql(mappedColumns);

                if (index.IsPrimaryKey)
                {
                    if (hasPrimaryKey)
                        continue;

                    ExecuteSchemaNonQuery($"ALTER TABLE [{escapedTableName}] ADD CONSTRAINT [{escapedIndexName}] PRIMARY KEY ({columnSql})");
                    hasPrimaryKey = true;
                    existingIndexNames.Add(index.Name);
                    continue;
                }

                var uniqueSql = index.IsUnique ? "UNIQUE " : string.Empty;
                ExecuteSchemaNonQuery($"CREATE {uniqueSql}INDEX [{escapedIndexName}] ON [{escapedTableName}] ({columnSql})");
                existingIndexNames.Add(index.Name);
            }
        }

        private void DropForeignKeyConstraints(IReadOnlyList<ForeignKeySnapshot> foreignKeys)
        {
            if (foreignKeys.Count == 0)
                return;

            foreach (var foreignKey in foreignKeys
                .OrderBy(snapshot => snapshot.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(snapshot => snapshot.ForeignTable, StringComparer.OrdinalIgnoreCase))
            {
                var escapedForeignTable = EscapeSqlIdentifier(foreignKey.ForeignTable);
                var escapedRelationshipName = EscapeSqlIdentifier(foreignKey.Name);
                ExecuteSchemaNonQuery($"ALTER TABLE [{escapedForeignTable}] DROP CONSTRAINT [{escapedRelationshipName}]");
            }
        }

        private void RecreateForeignKeyConstraints(
            IReadOnlyList<ForeignKeySnapshot> foreignKeys,
            string? sourceTableName = null,
            string? targetTableName = null,
            string? fieldRenameTableName = null,
            string? sourceFieldName = null,
            string? targetFieldName = null)
        {
            if (foreignKeys.Count == 0)
                return;

            foreach (var foreignKey in foreignKeys
                .OrderBy(snapshot => snapshot.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(snapshot => snapshot.ForeignTable, StringComparer.OrdinalIgnoreCase)
                .ThenBy(snapshot => snapshot.PrimaryTable, StringComparer.OrdinalIgnoreCase))
            {
                var mappedPrimaryTable = RemapIdentifier(foreignKey.PrimaryTable, sourceTableName, targetTableName);
                var mappedForeignTable = RemapIdentifier(foreignKey.ForeignTable, sourceTableName, targetTableName);

                var mappedPrimaryColumns = string.Equals(mappedPrimaryTable, fieldRenameTableName, StringComparison.OrdinalIgnoreCase)
                    ? RemapColumns(foreignKey.PrimaryColumns, sourceFieldName, targetFieldName)
                    : foreignKey.PrimaryColumns.ToList();
                var mappedForeignColumns = string.Equals(mappedForeignTable, fieldRenameTableName, StringComparison.OrdinalIgnoreCase)
                    ? RemapColumns(foreignKey.ForeignColumns, sourceFieldName, targetFieldName)
                    : foreignKey.ForeignColumns.ToList();

                if (mappedPrimaryColumns.Count == 0 || mappedForeignColumns.Count == 0 || mappedPrimaryColumns.Count != mappedForeignColumns.Count)
                    continue;

                if (!TableExists(mappedPrimaryTable) || !TableExists(mappedForeignTable))
                    continue;

                var primaryTableColumns = GetTableColumnSet(mappedPrimaryTable);
                var foreignTableColumns = GetTableColumnSet(mappedForeignTable);
                if (mappedPrimaryColumns.Any(column => !primaryTableColumns.Contains(column)) ||
                    mappedForeignColumns.Any(column => !foreignTableColumns.Contains(column)))
                {
                    continue;
                }

                if (ForeignKeyConstraintExists(mappedForeignTable, foreignKey.Name))
                    continue;

                var foreignColumnsSql = BuildColumnListSql(mappedForeignColumns);
                var primaryColumnsSql = BuildColumnListSql(mappedPrimaryColumns);
                var escapedConstraintName = EscapeSqlIdentifier(foreignKey.Name);
                var escapedForeignTable = EscapeSqlIdentifier(mappedForeignTable);
                var escapedPrimaryTable = EscapeSqlIdentifier(mappedPrimaryTable);

                var sql = $"ALTER TABLE [{escapedForeignTable}] ADD CONSTRAINT [{escapedConstraintName}] FOREIGN KEY ({foreignColumnsSql}) REFERENCES [{escapedPrimaryTable}] ({primaryColumnsSql})";
                if (foreignKey.CascadeUpdate)
                    sql += " ON UPDATE CASCADE";
                if (foreignKey.CascadeDelete)
                    sql += " ON DELETE CASCADE";

                ExecuteSchemaNonQuery(sql);
            }
        }

        private bool ForeignKeyConstraintExists(string foreignTableName, string relationshipName)
        {
            EnsureOleDbConnection();

            try
            {
                var schema = _oleDbConnection!.GetSchema("ForeignKeys");
                foreach (DataRow row in schema.Rows)
                {
                    var existingName = GetRowStringByCandidates(row, "FK_NAME", "CONSTRAINT_NAME", "RELATIONSHIP_NAME");
                    if (!string.Equals(existingName, relationshipName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var existingForeignTable = GetRowStringByCandidates(row, "FK_TABLE_NAME", "TABLE_NAME");
                    if (string.Equals(existingForeignTable, foreignTableName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch
            {
                // Keep behavior best-effort when ForeignKeys metadata is unavailable.
            }

            return false;
        }

        private HashSet<string> GetTableColumnSet(string tableName)
        {
            EnsureOleDbConnection();
            var columnNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var schema = _oleDbConnection!.GetSchema("Columns", new string[] { null!, null!, tableName, null! });
                foreach (DataRow row in schema.Rows)
                {
                    var columnName = GetRowString(row, "COLUMN_NAME");
                    if (!string.IsNullOrWhiteSpace(columnName))
                    {
                        columnNames.Add(columnName);
                    }
                }
            }
            catch
            {
                // Column metadata may be unavailable for provider-specific objects.
            }

            return columnNames;
        }

        private static string BuildColumnListSql(IEnumerable<string> columns)
        {
            return string.Join(", ", columns.Select(column => $"[{EscapeSqlIdentifier(column)}]"));
        }

        private static string RemapIdentifier(string value, string? sourceValue, string? targetValue)
        {
            if (string.IsNullOrWhiteSpace(sourceValue) || string.IsNullOrWhiteSpace(targetValue))
                return value;

            return string.Equals(value, sourceValue, StringComparison.OrdinalIgnoreCase)
                ? targetValue
                : value;
        }

        private static List<string> RemapColumns(IReadOnlyList<string> columns, string? sourceColumn, string? targetColumn)
        {
            var result = new List<string>(columns.Count);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var column in columns)
            {
                var mappedColumn = string.IsNullOrWhiteSpace(sourceColumn) || string.IsNullOrWhiteSpace(targetColumn)
                    ? column
                    : string.Equals(column, sourceColumn, StringComparison.OrdinalIgnoreCase)
                        ? targetColumn
                        : column;

                if (string.IsNullOrWhiteSpace(mappedColumn) || !seen.Add(mappedColumn))
                    continue;

                result.Add(mappedColumn);
            }

            return result;
        }

        private static string? GetRowStringByCandidates(DataRow row, params string[] candidateColumns)
        {
            foreach (var candidateColumn in candidateColumns)
            {
                var value = GetRowString(row, candidateColumn);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return null;
        }

        private static int? GetRowIntByCandidates(DataRow row, params string[] candidateColumns)
        {
            foreach (var candidateColumn in candidateColumns)
            {
                var value = GetRowInt(row, candidateColumn);
                if (value.HasValue)
                    return value;
            }

            return null;
        }

        private static string BuildAccessDataTypeDeclarationFromColumn(TableColumnDefinition column)
        {
            var dataTypeCode = column.DataTypeCode ?? (int)OleDbType.VarWChar;
            var oleDbType = (OleDbType)dataTypeCode;

            return oleDbType switch
            {
                OleDbType.Boolean => "YESNO",
                OleDbType.UnsignedTinyInt => "BYTE",
                OleDbType.SmallInt => "SHORT",
                OleDbType.Integer => "LONG",
                OleDbType.Single => "SINGLE",
                OleDbType.Double => "DOUBLE",
                OleDbType.Currency => "CURRENCY",
                OleDbType.Decimal or OleDbType.Numeric => BuildDecimalTypeDeclaration(column),
                OleDbType.Date or OleDbType.DBDate or OleDbType.DBTime or OleDbType.DBTimeStamp => "DATETIME",
                OleDbType.Guid => "GUID",
                OleDbType.Binary => "BINARY",
                OleDbType.LongVarBinary or OleDbType.VarBinary => "LONGBINARY",
                OleDbType.LongVarChar or OleDbType.LongVarWChar => "LONGTEXT",
                OleDbType.Char or OleDbType.VarChar or OleDbType.WChar or OleDbType.VarWChar or OleDbType.BSTR =>
                    $"TEXT({NormalizeTextLength(column.MaxLength)})",
                _ => "LONGTEXT"
            };
        }

        private static string BuildDecimalTypeDeclaration(TableColumnDefinition column)
        {
            var precision = Math.Clamp(column.NumericPrecision ?? 18, 1, 28);
            var scale = Math.Clamp(column.NumericScale ?? 0, 0, precision);
            return $"DECIMAL({precision},{scale})";
        }

        private static int NormalizeTextLength(int? maxLength)
        {
            if (!maxLength.HasValue || maxLength.Value <= 0)
                return 255;

            return Math.Clamp(maxLength.Value, 1, 255);
        }

        private static bool ShouldUseOleDbRenameFallback(Exception ex)
        {
            var message = ex.Message ?? string.Empty;
            if (message.IndexOf("table not found", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("field not found", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("active content", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return ex.InnerException != null && ShouldUseOleDbRenameFallback(ex.InnerException);
        }

        private static int ValidateSizedType(int size, int min, int max, int defaultValue, string sizeParamName, string dataTypeName)
        {
            var effectiveSize = size == 0 ? defaultValue : size;
            if (effectiveSize < min || effectiveSize > max)
                throw new ArgumentOutOfRangeException(sizeParamName, $"{dataTypeName} size must be between {min} and {max}.");

            return effectiveSize;
        }

        private static string ValidateUnsizedType(int size, string sizeParamName, string dataTypeName)
        {
            if (size != 0)
                throw new ArgumentOutOfRangeException(sizeParamName, $"{dataTypeName} does not support a size argument.");

            return dataTypeName;
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

        private static dynamic? TryGetCurrentDb(dynamic accessApp)
        {
            try
            {
                return accessApp.CurrentDb();
            }
            catch
            {
                return null;
            }
        }

        private static dynamic? FindQueryDef(dynamic currentDb, string queryName)
        {
            var queryDefs = TryGetDynamicProperty(currentDb, "QueryDefs");
            if (queryDefs == null)
                return null;

            foreach (var queryDef in queryDefs)
            {
                var name = SafeToString(TryGetDynamicProperty(queryDef, "Name"));
                if (string.Equals(name, queryName, StringComparison.OrdinalIgnoreCase))
                    return queryDef;
            }

            return null;
        }

        private static void DeleteQueryInternal(dynamic currentDb, string queryName)
        {
            if (FindQueryDef(currentDb, queryName) == null)
                throw new InvalidOperationException($"Query not found: {queryName}");

            var queryDefs = TryGetDynamicProperty(currentDb, "QueryDefs")
                ?? throw new InvalidOperationException("DAO QueryDefs collection is unavailable.");

            _ = InvokeDynamicMethod(queryDefs, "Delete", queryName);
        }

        private static dynamic? FindRelationship(dynamic currentDb, string relationshipName)
        {
            var relationships = TryGetDynamicProperty(currentDb, "Relations");
            if (relationships == null)
                return null;

            foreach (var relationship in relationships)
            {
                var name = SafeToString(TryGetDynamicProperty(relationship, "Name"));
                if (string.Equals(name, relationshipName, StringComparison.OrdinalIgnoreCase))
                    return relationship;
            }

            return null;
        }

        private void CreateRelationshipInternal(
            dynamic currentDb,
            string relationshipName,
            string tableName,
            string fieldName,
            string foreignTableName,
            string foreignFieldName,
            bool enforceIntegrity,
            bool cascadeUpdate,
            bool cascadeDelete)
        {
            if (RelationshipExists(currentDb, relationshipName))
                throw new InvalidOperationException($"Relationship already exists: {relationshipName}");

            var attributes = BuildRelationshipAttributes(enforceIntegrity, cascadeUpdate, cascadeDelete);
            // DAO and MCP APIs both use (primaryTable, foreignTable) here.
            var relationship = InvokeDynamicMethod(currentDb, "CreateRelation", relationshipName, tableName, foreignTableName, attributes)
                ?? throw new InvalidOperationException("Failed to create DAO Relationship object.");

            // DAO field mapping: Name = primary key column, ForeignName = foreign key column.
            var relationshipField = InvokeDynamicMethod(relationship, "CreateField", fieldName)
                ?? throw new InvalidOperationException("Failed to create DAO Relationship field object.");
            SetDynamicProperty(relationshipField, "ForeignName", foreignFieldName);

            var relationshipFields = TryGetDynamicProperty(relationship, "Fields")
                ?? throw new InvalidOperationException("Relationship fields collection is unavailable.");
            _ = InvokeDynamicMethod(relationshipFields, "Append", relationshipField);

            var relationships = TryGetDynamicProperty(currentDb, "Relations")
                ?? throw new InvalidOperationException("DAO Relations collection is unavailable.");
            _ = InvokeDynamicMethod(relationships, "Append", relationship);
        }

        private bool DeleteRelationshipInternal(dynamic currentDb, string relationshipName)
        {
            if (FindRelationship(currentDb, relationshipName) is { } relationship)
            {
                var daoRelationshipName = SafeToString(TryGetDynamicProperty(relationship, "Name")) ?? relationshipName;
                var relationships = TryGetDynamicProperty(currentDb, "Relations")
                    ?? throw new InvalidOperationException("DAO Relations collection is unavailable.");
                _ = InvokeDynamicMethod(relationships, "Delete", daoRelationshipName);
                return true;
            }

            return DeleteRelationshipViaOleDb(relationshipName);
        }

        private bool RelationshipExists(dynamic currentDb, string relationshipName)
        {
            if (FindRelationship(currentDb, relationshipName) != null)
                return true;

            try
            {
                EnsureOleDbConnection();
                var schema = _oleDbConnection!.GetSchema("ForeignKeys");
                foreach (DataRow row in schema.Rows)
                {
                    var schemaName = row["FK_NAME"]?.ToString();
                    if (string.Equals(schemaName, relationshipName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch
            {
                // Ignore fallback metadata errors and treat as "not found".
            }

            return false;
        }

        private bool DeleteRelationshipViaOleDb(string relationshipName)
        {
            try
            {
                EnsureOleDbConnection();
                var schema = _oleDbConnection!.GetSchema("ForeignKeys");
                var candidateTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (DataRow row in schema.Rows)
                {
                    var schemaName = row["FK_NAME"]?.ToString();
                    if (!string.Equals(schemaName, relationshipName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var tableName = row["TABLE_NAME"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(tableName))
                        candidateTables.Add(tableName);
                }

                foreach (var tableName in candidateTables)
                {
                    var sql = $"ALTER TABLE [{EscapeSqlIdentifier(tableName)}] DROP CONSTRAINT [{EscapeSqlIdentifier(relationshipName)}]";
                    try
                    {
                        using var command = new OleDbCommand(sql, _oleDbConnection);
                        command.ExecuteNonQuery();
                        return true;
                    }
                    catch
                    {
                        // Continue trying other candidate tables.
                    }
                }
            }
            catch
            {
                // Ignore metadata errors and report not found.
            }

            return false;
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

        private static bool MacroExists(dynamic accessApp, string macroName)
        {
            try
            {
                foreach (var macro in accessApp.CurrentProject.AllMacros)
                {
                    var currentName = SafeToString(TryGetDynamicProperty(macro, "Name"));
                    if (string.Equals(currentName, macroName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch
            {
                // Return false when macro enumeration is unavailable.
            }

            return false;
        }

        private static string BuildTemporaryTextPath(string prefix)
        {
            return Path.Combine(Path.GetTempPath(), $"{prefix}_{Guid.NewGuid():N}.txt");
        }

        private static string NormalizeTextTransferMode(string? mode)
        {
            if (string.IsNullOrWhiteSpace(mode))
                return TextModeJson;

            var normalized = mode.Trim().ToLowerInvariant();
            return normalized switch
            {
                TextModeJson => TextModeJson,
                TextModeAccessText => TextModeAccessText,
                _ => throw new ArgumentException("mode must be either 'json' or 'access_text'.", nameof(mode))
            };
        }

        private static string ResolveAccessTextImportObjectName(string? explicitName, string objectText, string objectKind, string vbNamePrefix)
        {
            if (!string.IsNullOrWhiteSpace(explicitName))
                return explicitName.Trim();

            return ExtractObjectNameFromAccessText(objectText, objectKind, vbNamePrefix);
        }

        private static string ExtractObjectNameFromAccessText(string objectText, string objectKind, string vbNamePrefix)
        {
            var vbNameMatch = Regex.Match(
                objectText,
                "^\\s*Attribute\\s+VB_Name\\s*=\\s*\"(?<name>[^\"]+)\"\\s*$",
                RegexOptions.Multiline | RegexOptions.IgnoreCase);

            if (vbNameMatch.Success)
            {
                var parsedName = vbNameMatch.Groups["name"].Value.Trim();
                if (parsedName.StartsWith(vbNamePrefix, StringComparison.OrdinalIgnoreCase))
                    parsedName = parsedName.Substring(vbNamePrefix.Length);

                if (!string.IsNullOrWhiteSpace(parsedName))
                    return parsedName;
            }

            throw new ArgumentException($"Unable to determine {objectKind} name from access_text payload. Provide {objectKind}_name.");
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // Best-effort temp cleanup.
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

        private static string EscapeSqlIdentifier(string identifier)
        {
            return identifier.Replace("]", "]]", StringComparison.Ordinal);
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

        private static bool HasRelationshipAttribute(int attributes, int flag)
        {
            return (attributes & flag) == flag;
        }

        private static int BuildRelationshipAttributes(bool enforceIntegrity, bool cascadeUpdate, bool cascadeDelete)
        {
            var attributes = 0;
            if (!enforceIntegrity)
                attributes |= DaoRelationAttributeDontEnforce;
            if (cascadeUpdate)
                attributes |= DaoRelationAttributeUpdateCascade;
            if (cascadeDelete)
                attributes |= DaoRelationAttributeDeleteCascade;

            return attributes;
        }

        private static string BuildRelationshipName(string tableName, string fieldName, string foreignTableName, string foreignFieldName)
        {
            var rawName = $"rel_{NormalizeNameFragment(tableName)}_{NormalizeNameFragment(fieldName)}_{NormalizeNameFragment(foreignTableName)}_{NormalizeNameFragment(foreignFieldName)}";
            return rawName.Length <= 64 ? rawName : rawName.Substring(0, 64);
        }

        private static string NormalizeNameFragment(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "x";

            var builder = new StringBuilder(value.Length);
            foreach (var character in value)
            {
                builder.Append(char.IsLetterOrDigit(character) ? char.ToLowerInvariant(character) : '_');
            }

            return builder.ToString().Trim('_');
        }

        private static string MapQueryDefType(int typeCode)
        {
            return typeCode switch
            {
                0 => "Select",
                16 => "Crosstab",
                32 => "Delete",
                48 => "Update",
                64 => "Append",
                80 => "MakeTable",
                96 => "DDL",
                112 => "PassThrough",
                128 => "Union",
                _ => $"QueryType({typeCode})"
            };
        }

        private sealed class IndexSnapshot
        {
            public string Name { get; set; } = string.Empty;
            public bool IsUnique { get; set; }
            public bool IsPrimaryKey { get; set; }
            public List<string> Columns { get; set; } = new();
        }

        private sealed class ForeignKeySnapshot
        {
            public string Name { get; set; } = string.Empty;
            public string PrimaryTable { get; set; } = string.Empty;
            public string ForeignTable { get; set; } = string.Empty;
            public List<string> PrimaryColumns { get; set; } = new();
            public List<string> ForeignColumns { get; set; } = new();
            public bool CascadeUpdate { get; set; }
            public bool CascadeDelete { get; set; }
        }

        private sealed class ForeignKeySnapshotBuilder
        {
            public string Name { get; set; } = string.Empty;
            public string PrimaryTable { get; set; } = string.Empty;
            public string ForeignTable { get; set; } = string.Empty;
            public int? UpdateRule { get; set; }
            public int? DeleteRule { get; set; }
            public List<(int Ordinal, string PrimaryColumn, string ForeignColumn)> Columns { get; } = new();
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

    public class IndexInfo
    {
        public string Name { get; set; } = "";
        public string Table { get; set; } = "";
        public bool IsUnique { get; set; }
        public bool IsPrimaryKey { get; set; }
        public List<string> Columns { get; set; } = new();
    }

    public class RelationshipInfo
    {
        public string Name { get; set; } = "";
        public string Table { get; set; } = "";
        public string Field { get; set; } = "";
        public string ForeignTable { get; set; } = "";
        public string ForeignField { get; set; } = "";
        public bool EnforceIntegrity { get; set; }
        public bool CascadeUpdate { get; set; }
        public bool CascadeDelete { get; set; }
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
