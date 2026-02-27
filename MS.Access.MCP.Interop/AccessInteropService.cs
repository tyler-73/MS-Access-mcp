using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.Odbc;
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
        private OdbcConnection? _odbcConnection;
        private dynamic? _accessApplication;
        private string? _currentDatabasePath;
        private string? _databasePassword;
        private string? _systemDatabasePath;
        private int _oleDbReleaseDepth = 0;
        private bool _restoreOleDbAfterRelease = false;
        private DataProviderKind _providerToRestoreAfterRelease = DataProviderKind.None;
        private string? _accessDatabasePath;
        private bool _accessDatabaseOpenedExclusive = false;
        private OleDbTransaction? _oleDbTransaction;
        private OdbcTransaction? _odbcTransaction;
        private DataProviderKind _activeDataProvider = DataProviderKind.None;
        private DataProviderKind _preferredDataProvider = DataProviderKind.OleDb;
        private DateTimeOffset? _transactionStartedAtUtc;
        private bool _disposed = false;
        private const int DaoRelationAttributeDontEnforce = 2;
        private const int DaoRelationAttributeUpdateCascade = 256;
        private const int DaoRelationAttributeDeleteCascade = 4096;
        private const string TextModeJson = "json";
        private const string TextModeAccessText = "access_text";
        private static readonly HashSet<string> SupportedDatabaseExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".accdb",
            ".mdb"
        };

        private enum DataProviderKind
        {
            None = 0,
            OleDb = 1,
            Odbc = 2
        }

        #region 1. Connection Management

        public void Connect(string databasePath)
        {
            Connect(databasePath, null, null);
        }

        public void Connect(string databasePath, string? databasePassword, string? systemDatabasePath)
        {
            var normalizedDatabasePath = NormalizeDatabasePath(databasePath, nameof(databasePath), requireExists: true);
            var normalizedSystemDatabasePath = NormalizeSystemDatabasePath(systemDatabasePath);

            _currentDatabasePath = normalizedDatabasePath;
            _databasePassword = string.IsNullOrWhiteSpace(databasePassword) ? null : databasePassword;
            _systemDatabasePath = normalizedSystemDatabasePath;
            try
            {
                OpenPreferredConnection(normalizedDatabasePath);
            }
            catch
            {
                _currentDatabasePath = null;
                _databasePassword = null;
                _systemDatabasePath = null;
                throw;
            }
        }

        public void Disconnect()
        {
            ResetTransactionState(attemptRollback: true);
            CloseSqlConnections();
            _currentDatabasePath = null;
            _databasePassword = null;
            _systemDatabasePath = null;
            _accessDatabasePath = null;
            _accessDatabaseOpenedExclusive = false;
        }

        public bool IsConnected => !string.IsNullOrWhiteSpace(_currentDatabasePath);
        public string? CurrentDatabasePath => _currentDatabasePath;

        public DatabaseCreateResult CreateDatabase(string databasePath, bool overwrite = false)
        {
            var normalizedDatabasePath = NormalizeDatabasePath(databasePath, nameof(databasePath), requireExists: false);
            var existedBefore = File.Exists(normalizedDatabasePath);
            if (existedBefore && !overwrite)
                throw new IOException($"Destination database already exists: {normalizedDatabasePath}. Set overwrite=true to replace it.");

            var directory = Path.GetDirectoryName(normalizedDatabasePath);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            if (existedBefore)
                File.Delete(normalizedDatabasePath);

            ExecuteWithTemporaryAccessApplication(accessApp =>
            {
                accessApp.NewCurrentDatabase(normalizedDatabasePath);
                accessApp.CloseCurrentDatabase();
            });

            var fileInfo = new FileInfo(normalizedDatabasePath);
            return new DatabaseCreateResult
            {
                DatabasePath = normalizedDatabasePath,
                ExistedBefore = existedBefore,
                SizeBytes = fileInfo.Exists ? fileInfo.Length : 0,
                LastWriteTimeUtc = fileInfo.Exists ? fileInfo.LastWriteTimeUtc : DateTime.MinValue
            };
        }

        public DatabaseBackupResult BackupDatabase(string sourceDatabasePath, string destinationDatabasePath, bool overwrite = false)
        {
            var normalizedSourcePath = NormalizeDatabasePath(sourceDatabasePath, nameof(sourceDatabasePath), requireExists: true);
            var normalizedDestinationPath = NormalizeDatabasePath(destinationDatabasePath, nameof(destinationDatabasePath), requireExists: false);
            EnsureDistinctDatabasePaths(normalizedSourcePath, normalizedDestinationPath, nameof(sourceDatabasePath), nameof(destinationDatabasePath));

            var destinationDirectory = Path.GetDirectoryName(normalizedDestinationPath);
            if (!string.IsNullOrWhiteSpace(destinationDirectory))
                Directory.CreateDirectory(destinationDirectory);

            var operatedOnConnectedDatabase = IsConnected &&
                !string.IsNullOrWhiteSpace(_currentDatabasePath) &&
                PathsMatch(_currentDatabasePath, normalizedSourcePath);

            return ExecuteWithConnectedDatabaseReleased(
                normalizedSourcePath,
                nameof(BackupDatabase),
                () =>
                {
                    if (File.Exists(normalizedDestinationPath) && !overwrite)
                        throw new IOException($"Destination database already exists: {normalizedDestinationPath}. Set overwrite=true to replace it.");

                    File.Copy(normalizedSourcePath, normalizedDestinationPath, overwrite);

                    var sourceInfo = new FileInfo(normalizedSourcePath);
                    var destinationInfo = new FileInfo(normalizedDestinationPath);

                    return new DatabaseBackupResult
                    {
                        SourceDatabasePath = normalizedSourcePath,
                        DestinationDatabasePath = normalizedDestinationPath,
                        BytesCopied = destinationInfo.Exists ? destinationInfo.Length : 0,
                        SourceLastWriteTimeUtc = sourceInfo.Exists ? sourceInfo.LastWriteTimeUtc : DateTime.MinValue,
                        DestinationLastWriteTimeUtc = destinationInfo.Exists ? destinationInfo.LastWriteTimeUtc : DateTime.MinValue,
                        OperatedOnConnectedDatabase = operatedOnConnectedDatabase
                    };
                });
        }

        public DatabaseCompactRepairResult CompactRepairDatabase(string sourceDatabasePath, string? destinationDatabasePath = null, bool overwrite = false)
        {
            var normalizedSourcePath = NormalizeDatabasePath(sourceDatabasePath, nameof(sourceDatabasePath), requireExists: true);
            var inPlace = string.IsNullOrWhiteSpace(destinationDatabasePath);
            var normalizedDestinationPath = inPlace
                ? BuildCompactTemporaryPath(normalizedSourcePath)
                : NormalizeDatabasePath(destinationDatabasePath!, nameof(destinationDatabasePath), requireExists: false);

            EnsureDistinctDatabasePaths(normalizedSourcePath, normalizedDestinationPath, nameof(sourceDatabasePath), nameof(destinationDatabasePath));

            var finalDestinationPath = inPlace ? normalizedSourcePath : normalizedDestinationPath;
            var destinationDirectory = Path.GetDirectoryName(normalizedDestinationPath);
            if (!string.IsNullOrWhiteSpace(destinationDirectory))
                Directory.CreateDirectory(destinationDirectory);

            var operatedOnConnectedDatabase = IsConnected &&
                !string.IsNullOrWhiteSpace(_currentDatabasePath) &&
                PathsMatch(_currentDatabasePath, normalizedSourcePath);

            return ExecuteWithConnectedDatabaseReleased(
                normalizedSourcePath,
                nameof(CompactRepairDatabase),
                () =>
                {
                    if (!inPlace && File.Exists(normalizedDestinationPath) && !overwrite)
                        throw new IOException($"Destination database already exists: {normalizedDestinationPath}. Set overwrite=true to replace it.");

                    if (File.Exists(normalizedDestinationPath))
                        File.Delete(normalizedDestinationPath);

                    RunCompactRepair(normalizedSourcePath, normalizedDestinationPath);

                    if (inPlace)
                    {
                        ReplaceFileInPlace(normalizedDestinationPath, normalizedSourcePath);
                    }

                    var sourceInfo = new FileInfo(normalizedSourcePath);
                    var destinationInfo = new FileInfo(finalDestinationPath);
                    return new DatabaseCompactRepairResult
                    {
                        SourceDatabasePath = normalizedSourcePath,
                        DestinationDatabasePath = finalDestinationPath,
                        InPlace = inPlace,
                        SourceSizeBytes = sourceInfo.Exists ? sourceInfo.Length : 0,
                        DestinationSizeBytes = destinationInfo.Exists ? destinationInfo.Length : 0,
                        DestinationLastWriteTimeUtc = destinationInfo.Exists ? destinationInfo.LastWriteTimeUtc : DateTime.MinValue,
                        OperatedOnConnectedDatabase = operatedOnConnectedDatabase
                    };
                });
        }

        #endregion

        #region 2. Data Access Object Models

        public List<TableInfo> GetTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var tables = new List<TableInfo>();
            
            // Use OleDb to get table information
            var schema = GetSchema("Tables");
            
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
            var schema = GetSchema("Views");
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
                var schema = GetSchema("ForeignKeys");

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

        public List<LinkedTableInfo> GetLinkedTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var linkedTables = new List<LinkedTableInfo>();

            try
            {
                var daoLinkedTables = ExecuteComOperation(accessApp =>
                {
                    var list = new List<LinkedTableInfo>();
                    var currentDb = TryGetCurrentDb(accessApp);
                    if (currentDb == null)
                        return list;

                    var tableDefs = TryGetDynamicProperty(currentDb, "TableDefs");
                    if (tableDefs == null)
                        return list;

                    foreach (var tableDef in tableDefs)
                    {
                        var tableName = SafeToString(TryGetDynamicProperty(tableDef, "Name"));
                        if (string.IsNullOrWhiteSpace(tableName) || IsSystemOrTemporaryTableName(tableName))
                            continue;

                        var connectString = SafeToString(TryGetDynamicProperty(tableDef, "Connect"));
                        if (string.IsNullOrWhiteSpace(connectString))
                            continue;

                        var sourceTableName = SafeToString(TryGetDynamicProperty(tableDef, "SourceTableName")) ?? string.Empty;
                        list.Add(new LinkedTableInfo
                        {
                            Name = tableName,
                            SourceTableName = sourceTableName,
                            ConnectString = connectString,
                            SourceDatabasePath = ExtractDatabasePathFromConnectString(connectString) ?? string.Empty,
                            Attributes = ToInt32(TryGetDynamicProperty(tableDef, "Attributes"))
                        });
                    }

                    return list;
                },
                requireExclusive: false,
                releaseOleDb: false);

                if (daoLinkedTables.Count > 0)
                {
                    return daoLinkedTables
                        .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
            }
            catch
            {
                // Fall back to OleDb metadata when DAO TableDefs are unavailable.
            }

            EnsureOleDbConnection();
            var schema = GetSchema("Tables");
            foreach (DataRow row in schema.Rows)
            {
                var tableName = GetRowString(row, "TABLE_NAME");
                if (string.IsNullOrWhiteSpace(tableName) || IsSystemOrTemporaryTableName(tableName))
                    continue;

                var tableType = GetRowString(row, "TABLE_TYPE");
                if (string.IsNullOrWhiteSpace(tableType) || tableType.IndexOf("LINK", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                linkedTables.Add(new LinkedTableInfo
                {
                    Name = tableName,
                    SourceTableName = string.Empty,
                    ConnectString = string.Empty,
                    SourceDatabasePath = string.Empty,
                    Attributes = 0
                });
            }

            return linkedTables
                .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public LinkedTableInfo LinkTable(
            string tableName,
            string sourceDatabasePath,
            string sourceTableName,
            string? connectString = null,
            bool overwrite = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            var normalizedSourceTableName = NormalizeSchemaIdentifier(sourceTableName, nameof(sourceTableName), "Source table name is required");
            var normalizedSourcePath = NormalizeLinkSourceDatabasePath(sourceDatabasePath, nameof(sourceDatabasePath));
            var normalizedConnectString = NormalizeLinkConnectString(connectString, normalizedSourcePath);

            if (!string.IsNullOrWhiteSpace(_currentDatabasePath) &&
                PathsMatch(_currentDatabasePath, normalizedSourcePath) &&
                string.Equals(normalizedTableName, normalizedSourceTableName, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Cannot create a linked table that points to itself.");
            }

            EnsureNoActiveTransaction(nameof(LinkTable));

            var linkedInfo = ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");
                var tableDefs = TryGetDynamicProperty(currentDb, "TableDefs")
                    ?? throw new InvalidOperationException("DAO TableDefs collection is unavailable.");

                var existing = FindTableDefWithRetry(accessApp, normalizedTableName);
                if (existing != null)
                {
                    if (!overwrite)
                        throw new InvalidOperationException($"Table already exists: {normalizedTableName}");

                    if (!IsLinkedTableDef(existing))
                        throw new InvalidOperationException($"Table '{normalizedTableName}' exists and is not a linked table. Refusing to overwrite.");

                    var existingName = SafeToString(TryGetDynamicProperty(existing, "Name")) ?? normalizedTableName;
                    _ = InvokeDynamicMethod(tableDefs, "Delete", existingName);
                }

                var tableDef = InvokeDynamicMethod(currentDb, "CreateTableDef", normalizedTableName)
                    ?? throw new InvalidOperationException("Failed to create DAO TableDef.");
                SetDynamicProperty(tableDef, "Connect", normalizedConnectString);
                SetDynamicProperty(tableDef, "SourceTableName", normalizedSourceTableName);
                _ = InvokeDynamicMethod(tableDefs, "Append", tableDef);
                _ = InvokeDynamicMethod(tableDefs, "Refresh");

                return new LinkedTableInfo
                {
                    Name = normalizedTableName,
                    SourceTableName = normalizedSourceTableName,
                    ConnectString = normalizedConnectString,
                    SourceDatabasePath = normalizedSourcePath,
                    Attributes = ToInt32(TryGetDynamicProperty(tableDef, "Attributes"))
                };
            },
            requireExclusive: false,
            releaseOleDb: true);

            RefreshOleDbConnectionAfterSchemaMutation();
            return linkedInfo;
        }

        public LinkedTableInfo RefreshLink(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            EnsureNoActiveTransaction(nameof(RefreshLink));

            var linkedInfo = ExecuteComOperation(accessApp =>
            {
                var tableDef = FindTableDefWithRetry(accessApp, normalizedTableName)
                    ?? throw new InvalidOperationException($"Table not found: {normalizedTableName}");
                if (!IsLinkedTableDef(tableDef))
                    throw new InvalidOperationException($"Table '{normalizedTableName}' is not a linked table.");

                _ = InvokeDynamicMethod(tableDef, "RefreshLink");

                var connectString = SafeToString(TryGetDynamicProperty(tableDef, "Connect")) ?? string.Empty;
                return new LinkedTableInfo
                {
                    Name = normalizedTableName,
                    SourceTableName = SafeToString(TryGetDynamicProperty(tableDef, "SourceTableName")) ?? string.Empty,
                    ConnectString = connectString,
                    SourceDatabasePath = ExtractDatabasePathFromConnectString(connectString) ?? string.Empty,
                    Attributes = ToInt32(TryGetDynamicProperty(tableDef, "Attributes"))
                };
            },
            requireExclusive: false,
            releaseOleDb: false);

            RefreshOleDbConnectionAfterSchemaMutation();
            return linkedInfo;
        }

        public LinkedTableInfo RelinkTable(
            string tableName,
            string sourceDatabasePath,
            string? sourceTableName = null,
            string? connectString = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");
            var normalizedSourcePath = NormalizeLinkSourceDatabasePath(sourceDatabasePath, nameof(sourceDatabasePath));
            var normalizedConnectString = NormalizeLinkConnectString(connectString, normalizedSourcePath);

            string? normalizedSourceTableName = null;
            if (!string.IsNullOrWhiteSpace(sourceTableName))
                normalizedSourceTableName = NormalizeSchemaIdentifier(sourceTableName, nameof(sourceTableName), "Source table name is required");

            EnsureNoActiveTransaction(nameof(RelinkTable));

            var linkedInfo = ExecuteComOperation(accessApp =>
            {
                var tableDef = FindTableDefWithRetry(accessApp, normalizedTableName)
                    ?? throw new InvalidOperationException($"Table not found: {normalizedTableName}");
                if (!IsLinkedTableDef(tableDef))
                    throw new InvalidOperationException($"Table '{normalizedTableName}' is not a linked table.");

                SetDynamicProperty(tableDef, "Connect", normalizedConnectString);
                if (!string.IsNullOrWhiteSpace(normalizedSourceTableName))
                {
                    var currentSourceTableName = SafeToString(TryGetDynamicProperty(tableDef, "SourceTableName")) ?? string.Empty;
                    if (!string.Equals(currentSourceTableName, normalizedSourceTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            SetDynamicProperty(tableDef, "SourceTableName", normalizedSourceTableName);
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException(
                                "Updating source_table_name on an existing linked table is not supported by this Access provider. Recreate the linked table instead.",
                                ex);
                        }
                    }
                }

                _ = InvokeDynamicMethod(tableDef, "RefreshLink");

                var effectiveSourceTableName = SafeToString(TryGetDynamicProperty(tableDef, "SourceTableName")) ?? string.Empty;
                var effectiveConnectString = SafeToString(TryGetDynamicProperty(tableDef, "Connect")) ?? normalizedConnectString;
                return new LinkedTableInfo
                {
                    Name = normalizedTableName,
                    SourceTableName = effectiveSourceTableName,
                    ConnectString = effectiveConnectString,
                    SourceDatabasePath = ExtractDatabasePathFromConnectString(effectiveConnectString) ?? normalizedSourcePath,
                    Attributes = ToInt32(TryGetDynamicProperty(tableDef, "Attributes"))
                };
            },
            requireExclusive: false,
            releaseOleDb: true);

            RefreshOleDbConnectionAfterSchemaMutation();
            return linkedInfo;
        }

        public void UnlinkTable(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            var normalizedTableName = NormalizeSchemaIdentifier(tableName, nameof(tableName), "Table name is required");

            EnsureNoActiveTransaction(nameof(UnlinkTable));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("Failed to get current DAO database.");
                var tableDefs = TryGetDynamicProperty(currentDb, "TableDefs")
                    ?? throw new InvalidOperationException("DAO TableDefs collection is unavailable.");
                var tableDef = FindTableDefWithRetry(accessApp, normalizedTableName)
                    ?? throw new InvalidOperationException($"Table not found: {normalizedTableName}");

                if (!IsLinkedTableDef(tableDef))
                    throw new InvalidOperationException($"Table '{normalizedTableName}' is not a linked table.");

                var daoName = SafeToString(TryGetDynamicProperty(tableDef, "Name")) ?? normalizedTableName;
                _ = InvokeDynamicMethod(tableDefs, "Delete", daoName);
                _ = InvokeDynamicMethod(tableDefs, "Refresh");
            },
            requireExclusive: false,
            releaseOleDb: true);

            RefreshOleDbConnectionAfterSchemaMutation();
        }

        public TransactionStatusInfo BeginTransaction(string? isolationLevel = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            PruneInvalidTransactionState();
            if (HasActiveTransaction())
                throw new InvalidOperationException("A transaction is already active. Commit or rollback it before starting a new one.");

            EnsureOleDbConnection();
            var parsedIsolationLevel = ParseIsolationLevel(isolationLevel);

            if (_activeDataProvider == DataProviderKind.Odbc)
            {
                try
                {
                    _odbcTransaction = _odbcConnection!.BeginTransaction(parsedIsolationLevel);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Transactions are not supported by the active ODBC Access provider.", ex);
                }
            }
            else
            {
                _oleDbTransaction = _oleDbConnection!.BeginTransaction(parsedIsolationLevel);
            }

            _transactionStartedAtUtc = DateTimeOffset.UtcNow;

            return GetTransactionStatus();
        }

        public TransactionStatusInfo CommitTransaction()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var transaction = GetActiveTransaction()
                ?? throw new InvalidOperationException("No active transaction to commit.");

            try
            {
                if (transaction is OleDbTransaction oleDbTransaction)
                {
                    oleDbTransaction.Commit();
                }
                else if (transaction is OdbcTransaction odbcTransaction)
                {
                    odbcTransaction.Commit();
                }
            }
            finally
            {
                ResetTransactionState(attemptRollback: false);
            }

            return GetTransactionStatus();
        }

        public TransactionStatusInfo RollbackTransaction()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var transaction = GetActiveTransaction()
                ?? throw new InvalidOperationException("No active transaction to rollback.");

            try
            {
                if (transaction is OleDbTransaction oleDbTransaction)
                {
                    oleDbTransaction.Rollback();
                }
                else if (transaction is OdbcTransaction odbcTransaction)
                {
                    odbcTransaction.Rollback();
                }
            }
            finally
            {
                ResetTransactionState(attemptRollback: false);
            }

            return GetTransactionStatus();
        }

        public TransactionStatusInfo GetTransactionStatus()
        {
            PruneInvalidTransactionState();
            var transaction = GetActiveTransaction();

            return new TransactionStatusInfo
            {
                Active = transaction != null,
                IsolationLevel = transaction?.IsolationLevel.ToString(),
                StartedAtUtc = _transactionStartedAtUtc
            };
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
            using var command = CreateCommand(createSql);
            command.ExecuteNonQuery();
        }

        public void DeleteTable(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();
            using var command = CreateCommand($"DROP TABLE [{tableName}]");
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
                var schema = GetSchema("Indexes");
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
            using var command = CreateCommand(sql);
            command.ExecuteNonQuery();
        }

        public void DeleteIndex(string tableName, string indexName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            if (string.IsNullOrWhiteSpace(indexName)) throw new ArgumentException("Index name is required", nameof(indexName));
            EnsureOleDbConnection();

            var sql = $"DROP INDEX [{EscapeSqlIdentifier(indexName)}] ON [{EscapeSqlIdentifier(tableName)}]";
            using var command = CreateCommand(sql);
            command.ExecuteNonQuery();
        }

        public SqlExecutionResult ExecuteSql(string sql, int maxRows = 200)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required", nameof(sql));
            if (maxRows <= 0) throw new ArgumentOutOfRangeException(nameof(maxRows), "maxRows must be greater than 0");
            EnsureOleDbConnection();

            using var command = CreateCommand(sql);
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

            var columnsSchema = GetSchema("Columns", new string[] { null!, null!, tableName, null! });
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
                var dataTypeName = GetProviderDataTypeName(row, dataTypeCode);

                columns.Add(new TableColumnDefinition
                {
                    Name = columnName,
                    DataType = dataTypeName,
                    DataTypeCode = dataTypeCode,
                    OrdinalPosition = GetRowInt(row, "ORDINAL_POSITION"),
                    MaxLength = GetRowInt(row, "CHARACTER_MAXIMUM_LENGTH") ?? GetRowInt(row, "COLUMN_SIZE"),
                    NumericPrecision = GetRowInt(row, "NUMERIC_PRECISION"),
                    NumericScale = GetRowInt(row, "NUMERIC_SCALE"),
                    IsNullable = IsColumnNullable(row),
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

        public Dictionary<string, object?> GetDatabaseSummaryProperties()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");

                return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Title"] = GetDaoPropertyValue(currentDb, "Title"),
                    ["Author"] = GetDaoPropertyValue(currentDb, "Author"),
                    ["Subject"] = GetDaoPropertyValue(currentDb, "Subject"),
                    ["Keywords"] = GetDaoPropertyValue(currentDb, "Keywords"),
                    ["Comments"] = GetDaoPropertyValue(currentDb, "Comments")
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SetDatabaseSummaryProperties(string? title = null, string? author = null, string? subject = null, string? keywords = null, string? comments = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");

                if (title != null)
                    SetDaoPropertyValue(currentDb, "Title", title, daoType: 10, createIfMissing: true);
                if (author != null)
                    SetDaoPropertyValue(currentDb, "Author", author, daoType: 10, createIfMissing: true);
                if (subject != null)
                    SetDaoPropertyValue(currentDb, "Subject", subject, daoType: 10, createIfMissing: true);
                if (keywords != null)
                    SetDaoPropertyValue(currentDb, "Keywords", keywords, daoType: 12, createIfMissing: true);
                if (comments != null)
                    SetDaoPropertyValue(currentDb, "Comments", comments, daoType: 12, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public List<DatabasePropertyInfo> GetDatabaseProperties(bool includeSystem = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var propertiesCollection = TryGetDynamicProperty(currentDb, "Properties")
                    ?? throw new InvalidOperationException("DAO properties collection is unavailable.");

                var properties = new List<DatabasePropertyInfo>();
                foreach (var property in propertiesCollection)
                {
                    var name = SafeToString(TryGetDynamicProperty(property, "Name"));
                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    var isSystem = IsLikelySystemDatabaseProperty(name);
                    if (!includeSystem && isSystem)
                        continue;

                    properties.Add(new DatabasePropertyInfo
                    {
                        Name = name,
                        Value = NormalizeValue(TryGetDynamicProperty(property, "Value")),
                        TypeCode = ToInt32(TryGetDynamicProperty(property, "Type")),
                        IsSystem = isSystem
                    });
                }

                return properties
                    .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public DatabasePropertyInfo GetDatabaseProperty(string propertyName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property name is required.", nameof(propertyName));

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var property = FindDaoProperty(currentDb, propertyName)
                    ?? throw new InvalidOperationException($"Database property not found: {propertyName}");

                var name = SafeToString(TryGetDynamicProperty(property, "Name")) ?? propertyName;
                return new DatabasePropertyInfo
                {
                    Name = name,
                    Value = NormalizeValue(TryGetDynamicProperty(property, "Value")),
                    TypeCode = ToInt32(TryGetDynamicProperty(property, "Type")),
                    IsSystem = IsLikelySystemDatabaseProperty(name)
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SetDatabaseProperty(string propertyName, string value, string? propertyType = null, bool createIfMissing = true)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property name is required.", nameof(propertyName));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");

                var existingProperty = FindDaoProperty(currentDb, propertyName);
                var existingValue = existingProperty != null
                    ? TryGetDynamicProperty(existingProperty, "Value")
                    : null;

                var daoType = ParseDaoDataType(propertyType);
                var convertedValue = ConvertPropertyValue(value, propertyType, existingValue);
                SetDaoPropertyValue(currentDb, propertyName, convertedValue, daoType, createIfMissing);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public TablePropertiesInfo GetTableProperties(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));

            return ExecuteComOperation(accessApp =>
            {
                var tableDef = FindTableDefWithRetry(accessApp, tableName)
                    ?? throw new InvalidOperationException($"Table not found: {tableName}");

                var actualName = SafeToString(TryGetDynamicProperty(tableDef, "Name")) ?? tableName;
                return new TablePropertiesInfo
                {
                    TableName = actualName,
                    Description = SafeToString(GetDaoPropertyValue(tableDef, "Description")),
                    ValidationRule = SafeToString(GetDaoPropertyValue(tableDef, "ValidationRule")),
                    ValidationText = SafeToString(GetDaoPropertyValue(tableDef, "ValidationText"))
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SetTableProperties(string tableName, string? description = null, string? validationRule = null, string? validationText = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));

            ExecuteComOperation(accessApp =>
            {
                var tableDef = FindTableDefWithRetry(accessApp, tableName)
                    ?? throw new InvalidOperationException($"Table not found: {tableName}");

                if (description != null)
                    SetDaoPropertyValue(tableDef, "Description", description, daoType: 12, createIfMissing: true);
                if (validationRule != null)
                    SetDaoPropertyValue(tableDef, "ValidationRule", validationRule, daoType: 12, createIfMissing: true);
                if (validationText != null)
                    SetDaoPropertyValue(tableDef, "ValidationText", validationText, daoType: 12, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public TableValidationInfo GetTableValidation(string tableName)
        {
            var tableProperties = GetTableProperties(tableName);
            return new TableValidationInfo
            {
                TableName = tableProperties.TableName,
                ValidationRule = tableProperties.ValidationRule,
                ValidationText = tableProperties.ValidationText
            };
        }

        public string? GetTableDescription(string tableName)
        {
            return GetTableProperties(tableName).Description;
        }

        public void SetTableDescription(string tableName, string description)
        {
            if (string.IsNullOrWhiteSpace(description)) throw new ArgumentException("description is required.", nameof(description));
            SetTableProperties(tableName, description: description);
        }

        public List<TableFieldDescriptionInfo> GetAllFieldDescriptions(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));

            return ExecuteComOperation(accessApp =>
            {
                var tableDef = FindTableDefWithRetry(accessApp, tableName)
                    ?? throw new InvalidOperationException($"Table not found: {tableName}");
                var fields = TryGetDynamicProperty(tableDef, "Fields")
                    ?? throw new InvalidOperationException("DAO Fields collection is unavailable.");

                var results = new List<TableFieldDescriptionInfo>();
                foreach (var field in fields)
                {
                    var fieldName = SafeToString(TryGetDynamicProperty(field, "Name"));
                    if (string.IsNullOrWhiteSpace(fieldName))
                        continue;

                    results.Add(new TableFieldDescriptionInfo
                    {
                        TableName = tableName,
                        FieldName = fieldName,
                        Description = SafeToString(GetDaoPropertyValue(field, "Description"))
                    });
                }

                return results
                    .OrderBy(item => item.FieldName, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public QueryPropertiesInfo GetQueryProperties(string queryName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("Query name is required.", nameof(queryName));

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var queryDef = FindQueryDef(currentDb, queryName)
                    ?? throw new InvalidOperationException($"Query not found: {queryName}");

                var actualName = SafeToString(TryGetDynamicProperty(queryDef, "Name")) ?? queryName;
                var parameters = new List<QueryParameterInfo>();
                var parameterCollection = TryGetDynamicProperty(queryDef, "Parameters");
                if (parameterCollection != null)
                {
                    foreach (var parameter in parameterCollection)
                    {
                        parameters.Add(new QueryParameterInfo
                        {
                            Name = SafeToString(TryGetDynamicProperty(parameter, "Name")) ?? "",
                            TypeCode = ToInt32(TryGetDynamicProperty(parameter, "Type")),
                            Value = NormalizeValue(TryGetDynamicProperty(parameter, "Value"))
                        });
                    }
                }

                return new QueryPropertiesInfo
                {
                    QueryName = actualName,
                    Description = SafeToString(GetDaoPropertyValue(queryDef, "Description")),
                    Sql = SafeToString(TryGetDynamicProperty(queryDef, "SQL")) ?? "",
                    Parameters = parameters
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SetQueryProperties(string queryName, string? description = null, string? sql = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("Query name is required.", nameof(queryName));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var queryDef = FindQueryDef(currentDb, queryName)
                    ?? throw new InvalidOperationException($"Query not found: {queryName}");

                if (description != null)
                    SetDaoPropertyValue(queryDef, "Description", description, daoType: 12, createIfMissing: true);
                if (sql != null)
                    SetDynamicProperty(queryDef, "SQL", sql);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public FieldPropertiesInfo GetFieldProperties(string tableName, string fieldName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required.", nameof(fieldName));

            return ExecuteComOperation(accessApp =>
            {
                var field = ResolveField(accessApp, tableName, fieldName);
                return new FieldPropertiesInfo
                {
                    TableName = tableName,
                    FieldName = SafeToString(TryGetDynamicProperty(field, "Name")) ?? fieldName,
                    TypeCode = ToInt32(TryGetDynamicProperty(field, "Type")),
                    Size = ToInt32(TryGetDynamicProperty(field, "Size")),
                    Required = ToBool(TryGetDynamicProperty(field, "Required"), false),
                    AllowZeroLength = ToBool(TryGetDynamicProperty(field, "AllowZeroLength"), false),
                    DefaultValue = SafeToString(GetDaoPropertyValue(field, "DefaultValue")),
                    ValidationRule = SafeToString(GetDaoPropertyValue(field, "ValidationRule")),
                    ValidationText = SafeToString(GetDaoPropertyValue(field, "ValidationText")),
                    InputMask = SafeToString(GetDaoPropertyValue(field, "InputMask")),
                    Caption = SafeToString(GetDaoPropertyValue(field, "Caption")),
                    RowSource = SafeToString(GetDaoPropertyValue(field, "RowSource")),
                    BoundColumn = ToNullableInt(GetDaoPropertyValue(field, "BoundColumn")),
                    ColumnCount = ToNullableInt(GetDaoPropertyValue(field, "ColumnCount")),
                    ColumnWidths = SafeToString(GetDaoPropertyValue(field, "ColumnWidths")),
                    LimitToList = ToNullableBool(GetDaoPropertyValue(field, "LimitToList")),
                    AllowMultipleValues = ToNullableBool(GetDaoPropertyValue(field, "AllowMultipleValues")),
                    DisplayControl = ToNullableInt(GetDaoPropertyValue(field, "DisplayControl"))
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public FieldAttributesInfo GetFieldAttributes(string tableName, string fieldName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required.", nameof(fieldName));

            return ExecuteComOperation(accessApp =>
            {
                var field = ResolveField(accessApp, tableName, fieldName);
                var allowMultipleValues = ToNullableBool(GetDaoPropertyValue(field, "AllowMultipleValues"));
                var isComplex = ToNullableBool(TryGetDynamicProperty(field, "IsComplex")) ?? allowMultipleValues;
                return new FieldAttributesInfo
                {
                    TableName = tableName,
                    FieldName = SafeToString(TryGetDynamicProperty(field, "Name")) ?? fieldName,
                    TypeCode = ToInt32(TryGetDynamicProperty(field, "Type")),
                    Size = ToInt32(TryGetDynamicProperty(field, "Size")),
                    Required = ToBool(TryGetDynamicProperty(field, "Required"), false),
                    AllowZeroLength = ToBool(TryGetDynamicProperty(field, "AllowZeroLength"), false),
                    Attributes = ToInt32(TryGetDynamicProperty(field, "Attributes")),
                    AllowMultipleValues = allowMultipleValues,
                    IsComplex = isComplex
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<MultiValueFieldInfo> DetectMultiValueFields(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));

            return ExecuteComOperation(accessApp =>
            {
                var tableDef = FindTableDefWithRetry(accessApp, tableName)
                    ?? throw new InvalidOperationException($"Table not found: {tableName}");
                var fields = TryGetDynamicProperty(tableDef, "Fields")
                    ?? throw new InvalidOperationException("DAO Fields collection is unavailable.");

                var results = new List<MultiValueFieldInfo>();
                foreach (var field in fields)
                {
                    var fieldName = SafeToString(TryGetDynamicProperty(field, "Name"));
                    if (string.IsNullOrWhiteSpace(fieldName))
                        continue;

                    var allowMultipleValues = ToNullableBool(GetDaoPropertyValue(field, "AllowMultipleValues"));
                    var isComplex = ToNullableBool(TryGetDynamicProperty(field, "IsComplex")) ?? allowMultipleValues;
                    if (allowMultipleValues != true && isComplex != true)
                        continue;

                    results.Add(new MultiValueFieldInfo
                    {
                        TableName = tableName,
                        FieldName = fieldName,
                        Attributes = ToInt32(TryGetDynamicProperty(field, "Attributes")),
                        AllowMultipleValues = allowMultipleValues,
                        IsComplex = isComplex
                    });
                }

                return results
                    .OrderBy(item => item.FieldName, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<MultiValueFieldRecordInfo> GetMultiValueFieldValues(string tableName, string fieldName, string? whereCondition = null, int maxRows = 100)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("tableName is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("fieldName is required.", nameof(fieldName));
            if (maxRows <= 0) throw new ArgumentOutOfRangeException(nameof(maxRows), "maxRows must be greater than 0.");

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var sql = BuildAttachmentQuery(tableName, whereCondition);
                var recordset = InvokeDynamicMethod(currentDb, "OpenRecordset", sql)
                    ?? throw new InvalidOperationException("Failed to open source recordset.");

                var results = new List<MultiValueFieldRecordInfo>();
                var rowIndex = 0;
                while (!ToBool(TryGetDynamicProperty(recordset, "EOF"), true) && results.Count < maxRows)
                {
                    rowIndex++;
                    var field = GetRecordsetField(recordset, fieldName)
                        ?? throw new InvalidOperationException($"Field not found: {fieldName}");
                    var complexValueRecordset = TryGetDynamicProperty(field, "Value");
                    var entries = complexValueRecordset == null
                        ? new List<Dictionary<string, object?>>()
                        : ReadComplexFieldEntries((object)complexValueRecordset);

                    results.Add(new MultiValueFieldRecordInfo
                    {
                        TableName = tableName,
                        FieldName = fieldName,
                        RowIndex = rowIndex,
                        Values = entries.Select(GetPreferredComplexFieldValue).ToList(),
                        Entries = entries
                    });

                    _ = InvokeDynamicMethod(recordset, "MoveNext");
                }

                return results;
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public MultiValueFieldUpdateInfo SetMultiValueFieldValues(string tableName, string fieldName, List<object?> values, string? whereCondition = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("tableName is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("fieldName is required.", nameof(fieldName));
            if (values == null) throw new ArgumentNullException(nameof(values));

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var sql = BuildAttachmentQuery(tableName, whereCondition);
                var recordset = InvokeDynamicMethod(currentDb, "OpenRecordset", sql, 2)
                    ?? throw new InvalidOperationException("Failed to open source recordset.");

                if (ToBool(TryGetDynamicProperty(recordset, "EOF"), true))
                    throw new InvalidOperationException("No matching row was found for multi-value update.");

                var field = GetRecordsetField(recordset, fieldName)
                    ?? throw new InvalidOperationException($"Field not found: {fieldName}");
                var complexValueRecordset = TryGetDynamicProperty(field, "Value")
                    ?? throw new InvalidOperationException($"Field '{fieldName}' is not a complex/multi-value field.");

                while (!ToBool(TryGetDynamicProperty(complexValueRecordset, "EOF"), true))
                {
                    _ = InvokeDynamicMethod(complexValueRecordset, "Delete");
                    _ = InvokeDynamicMethod(complexValueRecordset, "MoveNext");
                }

                foreach (var value in values)
                {
                    _ = InvokeDynamicMethod(complexValueRecordset, "AddNew");
                    SetComplexFieldEntryValue(complexValueRecordset, value);
                    _ = InvokeDynamicMethod(complexValueRecordset, "Update");
                }

                return new MultiValueFieldUpdateInfo
                {
                    TableName = tableName,
                    FieldName = fieldName,
                    ValuesWritten = values.Count
                };
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetFieldValidation(string tableName, string fieldName, string validationRule, string? validationText = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required.", nameof(fieldName));

            ExecuteComOperation(accessApp =>
            {
                var field = ResolveField(accessApp, tableName, fieldName);
                SetDaoPropertyValue(field, "ValidationRule", validationRule, daoType: 12, createIfMissing: true);
                if (validationText != null)
                    SetDaoPropertyValue(field, "ValidationText", validationText, daoType: 12, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetFieldDefault(string tableName, string fieldName, string defaultValue)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required.", nameof(fieldName));

            ExecuteComOperation(accessApp =>
            {
                var field = ResolveField(accessApp, tableName, fieldName);
                SetDaoPropertyValue(field, "DefaultValue", defaultValue, daoType: 12, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetFieldInputMask(string tableName, string fieldName, string inputMask)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required.", nameof(fieldName));

            ExecuteComOperation(accessApp =>
            {
                var field = ResolveField(accessApp, tableName, fieldName);
                SetDaoPropertyValue(field, "InputMask", inputMask, daoType: 12, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetFieldCaption(string tableName, string fieldName, string caption)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required.", nameof(fieldName));

            ExecuteComOperation(accessApp =>
            {
                var field = ResolveField(accessApp, tableName, fieldName);
                SetDaoPropertyValue(field, "Caption", caption, daoType: 12, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetLookupProperties(
            string tableName,
            string fieldName,
            string? rowSource = null,
            int? boundColumn = null,
            int? columnCount = null,
            string? columnWidths = null,
            bool? limitToList = null,
            bool? allowMultipleValues = null,
            int? displayControl = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Field name is required.", nameof(fieldName));

            ExecuteComOperation(accessApp =>
            {
                var field = ResolveField(accessApp, tableName, fieldName);
                if (rowSource != null)
                    SetDaoPropertyValue(field, "RowSource", rowSource, daoType: 12, createIfMissing: true);
                if (boundColumn.HasValue)
                    SetDaoPropertyValue(field, "BoundColumn", boundColumn.Value, daoType: 4, createIfMissing: true);
                if (columnCount.HasValue)
                    SetDaoPropertyValue(field, "ColumnCount", columnCount.Value, daoType: 4, createIfMissing: true);
                if (columnWidths != null)
                    SetDaoPropertyValue(field, "ColumnWidths", columnWidths, daoType: 12, createIfMissing: true);
                if (limitToList.HasValue)
                    SetDaoPropertyValue(field, "LimitToList", limitToList.Value, daoType: 1, createIfMissing: true);
                if (allowMultipleValues.HasValue)
                    SetDaoPropertyValue(field, "AllowMultipleValues", allowMultipleValues.Value, daoType: 1, createIfMissing: true);
                if (displayControl.HasValue)
                    SetDaoPropertyValue(field, "DisplayControl", displayControl.Value, daoType: 4, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
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

        public TransferSpreadsheetResult TransferSpreadsheet(
            string transferType,
            string tableName,
            string fileName,
            string? spreadsheetType = null,
            bool hasFieldNames = true,
            string? range = null,
            bool useOA = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name is required", nameof(fileName));

            var transferTypeValue = ParseTransferType(transferType, nameof(transferType));
            var spreadsheetTypeValue = ParseSpreadsheetType(spreadsheetType);
            var normalizedRange = string.IsNullOrWhiteSpace(range) ? null : range.Trim();

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");

                _ = InvokeDynamicMethod(
                    doCmd,
                    "TransferSpreadsheet",
                    transferTypeValue,
                    spreadsheetTypeValue,
                    tableName,
                    fileName,
                    hasFieldNames,
                    normalizedRange ?? Type.Missing,
                    useOA);
            },
            requireExclusive: false,
            releaseOleDb: false);

            return new TransferSpreadsheetResult
            {
                TransferType = transferTypeValue,
                SpreadsheetType = spreadsheetTypeValue,
                TableName = tableName,
                FileName = fileName,
                HasFieldNames = hasFieldNames,
                Range = normalizedRange,
                UseOA = useOA
            };
        }

        public TransferTextResult TransferText(
            string transferType,
            string tableName,
            string fileName,
            string? specificationName = null,
            bool hasFieldNames = true,
            string? htmlTableName = null,
            int? codePage = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name is required", nameof(fileName));

            var transferTypeValue = ParseTransferType(transferType, nameof(transferType));
            var normalizedSpecificationName = string.IsNullOrWhiteSpace(specificationName) ? null : specificationName.Trim();
            var normalizedHtmlTableName = string.IsNullOrWhiteSpace(htmlTableName) ? null : htmlTableName.Trim();

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");

                _ = InvokeDynamicMethod(
                    doCmd,
                    "TransferText",
                    transferTypeValue,
                    normalizedSpecificationName ?? Type.Missing,
                    tableName,
                    fileName,
                    hasFieldNames,
                    normalizedHtmlTableName ?? Type.Missing,
                    codePage.HasValue ? codePage.Value : Type.Missing);
            },
            requireExclusive: false,
            releaseOleDb: false);

            return new TransferTextResult
            {
                TransferType = transferTypeValue,
                SpecificationName = normalizedSpecificationName,
                TableName = tableName,
                FileName = fileName,
                HasFieldNames = hasFieldNames,
                HtmlTableName = normalizedHtmlTableName,
                CodePage = codePage
            };
        }

        public OutputToResult OutputTo(
            string objectType,
            string? objectName,
            string outputFormat,
            string? outputFile = null,
            bool autoStart = false,
            string? templateFile = null,
            string? encoding = null,
            string? outputQuality = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("Object type is required", nameof(objectType));
            if (string.IsNullOrWhiteSpace(outputFormat)) throw new ArgumentException("Output format is required", nameof(outputFormat));

            var objectTypeValue = ParseOutputObjectType(objectType, nameof(objectType));
            var normalizedObjectName = string.IsNullOrWhiteSpace(objectName) ? null : objectName.Trim();
            var normalizedOutputFile = string.IsNullOrWhiteSpace(outputFile) ? null : outputFile.Trim();
            var normalizedTemplateFile = string.IsNullOrWhiteSpace(templateFile) ? null : templateFile.Trim();
            var normalizedEncoding = string.IsNullOrWhiteSpace(encoding) ? null : encoding.Trim();
            var normalizedOutputFormat = ParseOutputFormat(outputFormat);
            var outputQualityValue = ParseOutputQuality(outputQuality);

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");

                _ = InvokeDynamicMethod(
                    doCmd,
                    "OutputTo",
                    objectTypeValue,
                    normalizedObjectName ?? Type.Missing,
                    normalizedOutputFormat,
                    normalizedOutputFile ?? Type.Missing,
                    autoStart,
                    normalizedTemplateFile ?? Type.Missing,
                    normalizedEncoding ?? Type.Missing,
                    outputQualityValue.HasValue ? outputQualityValue.Value : Type.Missing);
            },
            requireExclusive: false,
            releaseOleDb: false);

            return new OutputToResult
            {
                ObjectType = objectTypeValue,
                ObjectName = normalizedObjectName,
                OutputFormat = normalizedOutputFormat?.ToString() ?? outputFormat,
                OutputFile = normalizedOutputFile,
                AutoStart = autoStart,
                TemplateFile = normalizedTemplateFile,
                Encoding = normalizedEncoding,
                OutputQuality = outputQualityValue
            };
        }

        public void SetWarnings(bool warningsOn)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "SetWarnings", warningsOn);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void Echo(bool echoOn, string? statusBarText = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            var normalizedStatusBarText = string.IsNullOrWhiteSpace(statusBarText) ? null : statusBarText.Trim();

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Echo", echoOn, normalizedStatusBarText ?? Type.Missing);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void Hourglass(bool hourglassOn)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Hourglass", hourglassOn);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void GoToRecord(string? objectType = null, string? objectName = null, string? record = null, string? offset = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "GoToRecord",
                    NormalizeDoCmdVariant(objectType),
                    NormalizeDoCmdVariant(objectName),
                    NormalizeDoCmdVariant(record),
                    NormalizeDoCmdVariant(offset));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void FindRecord(
            string findWhat,
            string? match = null,
            string? matchCase = null,
            string? search = null,
            string? searchAsFormatted = null,
            string? onlyCurrentField = null,
            string? findFirst = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(findWhat)) throw new ArgumentException("Find text is required.", nameof(findWhat));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "FindRecord",
                    findWhat,
                    NormalizeDoCmdVariant(match),
                    NormalizeDoCmdVariant(matchCase),
                    NormalizeDoCmdVariant(search),
                    NormalizeDoCmdVariant(searchAsFormatted),
                    NormalizeDoCmdVariant(onlyCurrentField),
                    NormalizeDoCmdVariant(findFirst));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void ApplyFilter(string? filterName = null, string? whereCondition = null, string? controlName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "ApplyFilter",
                    NormalizeDoCmdVariant(filterName),
                    NormalizeDoCmdVariant(whereCondition),
                    NormalizeDoCmdVariant(controlName));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void ShowAllRecords()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "ShowAllRecords");
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void MaximizeWindow()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Maximize");
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void MinimizeWindow()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Minimize");
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void RestoreWindow()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Restore");
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void PrintOut(string? printRange = null, int? pageFrom = null, int? pageTo = null, string? printQuality = null, int? copies = null, bool collateCopies = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");

                _ = InvokeDynamicMethod(
                    doCmd,
                    "PrintOut",
                    NormalizeDoCmdVariant(printRange),
                    pageFrom.HasValue ? pageFrom.Value : Type.Missing,
                    pageTo.HasValue ? pageTo.Value : Type.Missing,
                    NormalizeDoCmdVariant(printQuality),
                    copies.HasValue ? copies.Value : Type.Missing,
                    collateCopies);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void OpenQuery(string queryName, string? view = null, string? dataMode = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("Query name is required.", nameof(queryName));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");

                _ = InvokeDynamicMethod(
                    doCmd,
                    "OpenQuery",
                    queryName,
                    NormalizeDoCmdVariant(view),
                    NormalizeDoCmdVariant(dataMode));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void RunSqlDoCmd(string sql, bool useTransaction = true)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL is required.", nameof(sql));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "RunSQL", sql, useTransaction);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void OpenTable(string tableName, string? view = null, string? dataMode = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));

            var viewValue = ParseOpenTableView(view);
            var dataModeValue = ParseOpenDataMode(dataMode);

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "OpenTable", tableName, viewValue, dataModeValue);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void OpenModule(string moduleName, string? procedureName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "OpenModule",
                    moduleName,
                    string.IsNullOrWhiteSpace(procedureName) ? Type.Missing : procedureName.Trim());
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void CopyObject(string? destinationDatabasePath = null, string? newName = null, string? sourceObjectType = null, string? sourceObjectName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(sourceObjectName)) throw new ArgumentException("Source object name is required.", nameof(sourceObjectName));

            var sourceObjectTypeValue = string.IsNullOrWhiteSpace(sourceObjectType)
                ? Type.Missing
                : ParseDoCmdObjectType(sourceObjectType, nameof(sourceObjectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "CopyObject",
                    string.IsNullOrWhiteSpace(destinationDatabasePath) ? Type.Missing : destinationDatabasePath.Trim(),
                    string.IsNullOrWhiteSpace(newName) ? Type.Missing : newName.Trim(),
                    sourceObjectTypeValue,
                    sourceObjectName.Trim());
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteObject(string objectName, string? objectType = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("Object name is required.", nameof(objectName));

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "DeleteObject", objectTypeValue, objectName.Trim());
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void RenameObject(string newName, string oldName, string? objectType = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(newName)) throw new ArgumentException("New name is required.", nameof(newName));
            if (string.IsNullOrWhiteSpace(oldName)) throw new ArgumentException("Old name is required.", nameof(oldName));

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Rename", newName.Trim(), objectTypeValue, oldName.Trim());
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SelectObject(string objectName, string? objectType = null, bool inDatabaseWindow = true)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("Object name is required.", nameof(objectName));

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "SelectObject", objectTypeValue, objectName.Trim(), inDatabaseWindow);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SaveObject(string? objectType = null, string? objectName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "Save",
                    objectTypeValue,
                    string.IsNullOrWhiteSpace(objectName) ? Type.Missing : objectName.Trim());
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void CloseObject(string? objectType = null, string? objectName = null, string? save = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));
            var saveValue = ParseCloseSaveOption(save);

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "Close",
                    objectTypeValue,
                    string.IsNullOrWhiteSpace(objectName) ? Type.Missing : objectName.Trim(),
                    saveValue);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public TransferDatabaseResult TransferDatabase(
            string transferType,
            string databaseType,
            string databaseName,
            string objectType,
            string source,
            string? destination = null,
            bool structureOnly = false,
            bool storeLogin = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(transferType)) throw new ArgumentException("Transfer type is required.", nameof(transferType));
            if (string.IsNullOrWhiteSpace(databaseType)) throw new ArgumentException("Database type is required.", nameof(databaseType));
            if (string.IsNullOrWhiteSpace(databaseName)) throw new ArgumentException("Database name is required.", nameof(databaseName));
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("Object type is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("Source is required.", nameof(source));

            var transferTypeValue = ParseTransferType(transferType, nameof(transferType));
            var objectTypeValue = ParseDoCmdObjectType(objectType, nameof(objectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "TransferDatabase",
                    transferTypeValue,
                    databaseType.Trim(),
                    databaseName.Trim(),
                    objectTypeValue,
                    source.Trim(),
                    string.IsNullOrWhiteSpace(destination) ? Type.Missing : destination.Trim(),
                    structureOnly,
                    storeLogin);
            },
            requireExclusive: true,
            releaseOleDb: true);

            return new TransferDatabaseResult
            {
                TransferType = transferTypeValue,
                DatabaseType = databaseType.Trim(),
                DatabaseName = databaseName.Trim(),
                ObjectType = objectTypeValue,
                Source = source.Trim(),
                Destination = string.IsNullOrWhiteSpace(destination) ? null : destination.Trim(),
                StructureOnly = structureOnly,
                StoreLogin = storeLogin
            };
        }

        public void RunCommand(string command)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(command)) throw new ArgumentException("Command is required.", nameof(command));

            var commandValue = ParseRunCommandValue(command);
            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "RunCommand", commandValue);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void GoToPage(string pageNumber, string? right = null, string? down = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(pageNumber)) throw new ArgumentException("Page number is required.", nameof(pageNumber));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "GoToPage",
                    NormalizeDoCmdVariant(pageNumber),
                    NormalizeDoCmdVariant(right),
                    NormalizeDoCmdVariant(down));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void GoToControl(string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("Control name is required.", nameof(controlName));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "GoToControl", controlName.Trim());
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void MoveSize(int? right = null, int? down = null, int? width = null, int? height = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "MoveSize",
                    right.HasValue ? right.Value : Type.Missing,
                    down.HasValue ? down.Value : Type.Missing,
                    width.HasValue ? width.Value : Type.Missing,
                    height.HasValue ? height.Value : Type.Missing);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void Requery(string? controlName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Requery", NormalizeDoCmdVariant(controlName));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void RepaintObject(string? objectType = null, string? objectName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "RepaintObject",
                    objectTypeValue,
                    string.IsNullOrWhiteSpace(objectName) ? Type.Missing : objectName.Trim());
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SendObject(
            string? objectType = null,
            string? objectName = null,
            string? outputFormat = null,
            string? to = null,
            string? cc = null,
            string? bcc = null,
            string? subject = null,
            string? messageText = null,
            bool? editMessage = null,
            string? templateFile = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));
            var outputFormatValue = string.IsNullOrWhiteSpace(outputFormat)
                ? Type.Missing
                : ParseOutputFormat(outputFormat);

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "SendObject",
                    objectTypeValue,
                    string.IsNullOrWhiteSpace(objectName) ? Type.Missing : objectName.Trim(),
                    outputFormatValue,
                    string.IsNullOrWhiteSpace(to) ? Type.Missing : to.Trim(),
                    string.IsNullOrWhiteSpace(cc) ? Type.Missing : cc.Trim(),
                    string.IsNullOrWhiteSpace(bcc) ? Type.Missing : bcc.Trim(),
                    string.IsNullOrWhiteSpace(subject) ? Type.Missing : subject.Trim(),
                    string.IsNullOrWhiteSpace(messageText) ? Type.Missing : messageText.Trim(),
                    editMessage.HasValue ? editMessage.Value : Type.Missing,
                    string.IsNullOrWhiteSpace(templateFile) ? Type.Missing : templateFile.Trim());
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void BrowseTo(string objectName, string? objectType = null, string? pathToSubformControl = null, string? whereCondition = null, string? page = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("Object name is required.", nameof(objectName));

            var objectTypeValue = string.IsNullOrWhiteSpace(objectType)
                ? Type.Missing
                : ParseDoCmdObjectType(objectType, nameof(objectType));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(
                    doCmd,
                    "BrowseTo",
                    objectName.Trim(),
                    objectTypeValue,
                    string.IsNullOrWhiteSpace(pathToSubformControl) ? Type.Missing : pathToSubformControl.Trim(),
                    string.IsNullOrWhiteSpace(whereCondition) ? Type.Missing : whereCondition.Trim(),
                    NormalizeDoCmdVariant(page));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void LockNavigationPane(bool lockNavigationPane = true)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "LockNavigationPane", lockNavigationPane);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void NavigateTo(string navigationCategory)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(navigationCategory)) throw new ArgumentException("Navigation category is required.", nameof(navigationCategory));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "NavigateTo", navigationCategory.Trim());
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void Beep()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "Beep");
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public StartupPropertiesInfo GetStartupProperties()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");

                return new StartupPropertiesInfo
                {
                    StartupForm = SafeToString(GetDaoPropertyValue(currentDb, "StartupForm")),
                    AppTitle = SafeToString(GetDaoPropertyValue(currentDb, "AppTitle")),
                    AppIcon = SafeToString(GetDaoPropertyValue(currentDb, "AppIcon"))
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SetStartupProperties(string? startupForm = null, string? appTitle = null, string? appIcon = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");

                if (startupForm != null)
                    SetDaoPropertyValue(currentDb, "StartupForm", startupForm, daoType: 10, createIfMissing: true);
                if (appTitle != null)
                    SetDaoPropertyValue(currentDb, "AppTitle", appTitle, daoType: 10, createIfMissing: true);
                if (appIcon != null)
                    SetDaoPropertyValue(currentDb, "AppIcon", appIcon, daoType: 10, createIfMissing: true);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public RibbonInfo GetRibbonXml(string? ribbonName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var defaultRibbonName = ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                return SafeToString(GetDaoPropertyValue(currentDb, "RibbonName"));
            },
            requireExclusive: false,
            releaseOleDb: false);

            var effectiveRibbonName = string.IsNullOrWhiteSpace(ribbonName) ? defaultRibbonName : ribbonName.Trim();
            if (string.IsNullOrWhiteSpace(effectiveRibbonName))
            {
                return new RibbonInfo
                {
                    RibbonName = null,
                    RibbonXml = null,
                    DefaultRibbonName = defaultRibbonName,
                    Exists = false
                };
            }

            EnsureOleDbConnection();
            if (!TableExists("USysRibbons"))
            {
                return new RibbonInfo
                {
                    RibbonName = effectiveRibbonName,
                    RibbonXml = null,
                    DefaultRibbonName = defaultRibbonName,
                    Exists = false
                };
            }

            string? ribbonXml = null;
            using (var command = CreateCommand("SELECT TOP 1 [RibbonXML] FROM [USysRibbons] WHERE [RibbonName] = ? ORDER BY [ID] DESC"))
            {
                AddCommandParameter(command, "@p1", effectiveRibbonName);
                using var reader = command.ExecuteReader();
                if (reader.Read())
                {
                    ribbonXml = reader[0] == DBNull.Value ? null : reader[0]?.ToString();
                }
            }

            return new RibbonInfo
            {
                RibbonName = effectiveRibbonName,
                RibbonXml = ribbonXml,
                DefaultRibbonName = defaultRibbonName,
                Exists = !string.IsNullOrWhiteSpace(ribbonXml)
            };
        }

        public void SetRibbonXml(string ribbonName, string ribbonXml, bool applyAsDefault = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(ribbonName)) throw new ArgumentException("Ribbon name is required.", nameof(ribbonName));
            if (string.IsNullOrWhiteSpace(ribbonXml)) throw new ArgumentException("Ribbon XML is required.", nameof(ribbonXml));

            EnsureNoActiveTransaction("Ribbon XML update");
            EnsureUsysRibbonsTable();

            using (var deleteCommand = CreateCommand("DELETE FROM [USysRibbons] WHERE [RibbonName] = ?"))
            {
                AddCommandParameter(deleteCommand, "@p1", ribbonName);
                _ = deleteCommand.ExecuteNonQuery();
            }

            using (var insertCommand = CreateCommand("INSERT INTO [USysRibbons] ([RibbonName], [RibbonXML]) VALUES (?, ?)"))
            {
                AddCommandParameter(insertCommand, "@p1", ribbonName);
                AddCommandParameter(insertCommand, "@p2", ribbonXml);
                _ = insertCommand.ExecuteNonQuery();
            }

            if (applyAsDefault)
            {
                ExecuteComOperation(accessApp =>
                {
                    var currentDb = TryGetCurrentDb(accessApp)
                        ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                    SetDaoPropertyValue(currentDb, "RibbonName", ribbonName, daoType: 10, createIfMissing: true);
                },
                requireExclusive: true,
                releaseOleDb: true);
            }
        }

        public ApplicationInfo GetApplicationInfo()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var currentProject = TryGetDynamicProperty(accessApp, "CurrentProject");
                var currentData = TryGetDynamicProperty(accessApp, "CurrentData");

                return new ApplicationInfo
                {
                    Name = SafeToString(TryGetDynamicProperty(accessApp, "Name")) ?? "Microsoft Access",
                    Version = SafeToString(TryGetDynamicProperty(accessApp, "Version")),
                    Build = SafeToString(TryGetDynamicProperty(accessApp, "Build")),
                    IsTrusted = ToBool(TryGetDynamicProperty(accessApp, "IsTrusted"), false),
                    CurrentDatabasePath = _currentDatabasePath,
                    CurrentProjectName = SafeToString(TryGetDynamicProperty(currentProject, "Name")),
                    CurrentProjectPath = SafeToString(TryGetDynamicProperty(currentProject, "Path")),
                    CurrentProjectFullName = SafeToString(TryGetDynamicProperty(currentProject, "FullName")),
                    CurrentDataName = SafeToString(TryGetDynamicProperty(currentData, "Name")),
                    CurrentDataPath = SafeToString(TryGetDynamicProperty(currentData, "Path"))
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public CurrentProjectDataInfo GetCurrentProjectData()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var currentProject = TryGetDynamicProperty(accessApp, "CurrentProject");
                var currentData = TryGetDynamicProperty(accessApp, "CurrentData");

                return new CurrentProjectDataInfo
                {
                    CurrentProjectName = SafeToString(TryGetDynamicProperty(currentProject, "Name")),
                    CurrentProjectPath = SafeToString(TryGetDynamicProperty(currentProject, "Path")),
                    CurrentProjectFullName = SafeToString(TryGetDynamicProperty(currentProject, "FullName")),
                    CurrentDataName = SafeToString(TryGetDynamicProperty(currentData, "Name")),
                    CurrentDataPath = SafeToString(TryGetDynamicProperty(currentData, "Path")),
                    CurrentDataAllTablesCount = ToNullableInt(TryGetDynamicProperty(TryGetDynamicProperty(currentData, "AllTables"), "Count")),
                    CurrentDataAllQueriesCount = ToNullableInt(TryGetDynamicProperty(TryGetDynamicProperty(currentData, "AllQueries"), "Count"))
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public object? GetApplicationOption(string optionName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(optionName)) throw new ArgumentException("Option name is required.", nameof(optionName));

            return ExecuteComOperation(accessApp =>
            {
                var optionValue = InvokeDynamicMethod(accessApp, "GetOption", optionName.Trim());
                return NormalizeValue(optionValue ?? DBNull.Value);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void SetApplicationOption(string optionName, object value)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(optionName)) throw new ArgumentException("Option name is required.", nameof(optionName));

            ExecuteComOperation(accessApp =>
            {
                _ = InvokeDynamicMethod(accessApp, "SetOption", optionName.Trim(), value);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void RunDataMacro(string macroName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required.", nameof(macroName));

            ExecuteComOperation(accessApp =>
            {
                try
                {
                    _ = InvokeDynamicMethod(accessApp, "RunDataMacro", macroName.Trim());
                }
                catch
                {
                    var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                        ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                    _ = InvokeDynamicMethod(doCmd, "RunMacro", macroName.Trim());
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public List<DataMacroInfo> GetTableDataMacros(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));

            return ExecuteComOperation(accessApp =>
            {
                var currentData = TryGetDynamicProperty(accessApp, "CurrentData");
                var allTables = TryGetDynamicProperty(currentData, "AllTables");
                var table = allTables == null ? null : FindObjectByName(allTables, tableName);
                var dataMacros = table == null ? null : TryGetDynamicProperty(table, "DataMacros");
                var results = new List<DataMacroInfo>();

                if (dataMacros != null)
                {
                    foreach (var dataMacro in dataMacros)
                    {
                        var name = SafeToString(TryGetDynamicProperty(dataMacro, "Name"));
                        if (string.IsNullOrWhiteSpace(name))
                            continue;

                        results.Add(new DataMacroInfo
                        {
                            Name = name,
                            MacroType = SafeToString(TryGetDynamicProperty(dataMacro, "Type"))
                        });
                    }
                }

                if (results.Count > 0)
                {
                    return results
                        .OrderBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }

                var axlObject = InvokeDynamicMethod(accessApp, "SaveAsAXL", 0, tableName);
                var axl = SafeToString(axlObject) ?? string.Empty;
                foreach (Match match in Regex.Matches(axl, "<[^>]*DataMacro[^>]*\\bName\\s*=\\s*\"(?<name>[^\"]+)\"", RegexOptions.IgnoreCase))
                {
                    var name = match.Groups["name"].Value.Trim();
                    if (string.IsNullOrWhiteSpace(name) || results.Any(entry => string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    results.Add(new DataMacroInfo { Name = name, MacroType = null });
                }

                return results
                    .OrderBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void DeleteDataMacro(string tableName, string macroName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(macroName)) throw new ArgumentException("Macro name is required.", nameof(macroName));

            ExecuteComOperation(accessApp =>
            {
                var currentData = TryGetDynamicProperty(accessApp, "CurrentData");
                var allTables = TryGetDynamicProperty(currentData, "AllTables")
                    ?? throw new InvalidOperationException("CurrentData.AllTables is unavailable.");
                var table = FindObjectByName(allTables, tableName)
                    ?? throw new InvalidOperationException($"Table '{tableName}' was not found in CurrentData.AllTables.");
                var dataMacros = TryGetDynamicProperty(table, "DataMacros")
                    ?? throw new InvalidOperationException("DataMacros collection is unavailable for this table.");

                dynamic? targetMacro = null;
                foreach (var dataMacro in dataMacros)
                {
                    var name = SafeToString(TryGetDynamicProperty(dataMacro, "Name"));
                    if (string.Equals(name, macroName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetMacro = dataMacro;
                        break;
                    }
                }

                if (targetMacro == null)
                    throw new InvalidOperationException($"Data macro '{macroName}' was not found on table '{tableName}'.");

                _ = InvokeDynamicMethod(targetMacro, "Delete");
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public AutoExecInfo GetAutoExecInfo()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var exists = MacroExists(accessApp, "AutoExec");
                return new AutoExecInfo
                {
                    Exists = exists,
                    MacroName = "AutoExec"
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void RunAutoExec()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                if (!MacroExists(accessApp, "AutoExec"))
                    throw new InvalidOperationException("AutoExec macro does not exist.");

                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "RunMacro", "AutoExec");
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<QueryParameterInfo> GetQueryParameters(string queryName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("Query name is required.", nameof(queryName));

            var properties = GetQueryProperties(queryName);
            return properties.Parameters
                .OrderBy(parameter => parameter.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public List<DaoContainerInfo> GetContainers()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var containers = TryGetDynamicProperty(currentDb, "Containers")
                    ?? throw new InvalidOperationException("DAO Containers collection is unavailable.");

                var results = new List<DaoContainerInfo>();
                foreach (var container in containers)
                {
                    var name = SafeToString(TryGetDynamicProperty(container, "Name"));
                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    results.Add(new DaoContainerInfo
                    {
                        Name = name,
                        DocumentCount = ToInt32(TryGetDynamicProperty(TryGetDynamicProperty(container, "Documents"), "Count"))
                    });
                }

                return results
                    .OrderBy(container => container.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<DaoContainerDocumentInfo> GetContainerDocuments(string containerName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(containerName)) throw new ArgumentException("Container name is required.", nameof(containerName));

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var containers = TryGetDynamicProperty(currentDb, "Containers")
                    ?? throw new InvalidOperationException("DAO Containers collection is unavailable.");

                dynamic? targetContainer = null;
                foreach (var container in containers)
                {
                    var name = SafeToString(TryGetDynamicProperty(container, "Name"));
                    if (string.Equals(name, containerName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetContainer = container;
                        break;
                    }
                }

                if (targetContainer == null)
                    throw new InvalidOperationException($"Container not found: {containerName}");

                var documents = TryGetDynamicProperty(targetContainer, "Documents")
                    ?? throw new InvalidOperationException($"Documents collection is unavailable for container '{containerName}'.");

                var results = new List<DaoContainerDocumentInfo>();
                foreach (var document in documents)
                {
                    var name = SafeToString(TryGetDynamicProperty(document, "Name"));
                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    results.Add(new DaoContainerDocumentInfo
                    {
                        ContainerName = containerName,
                        Name = name,
                        Owner = SafeToString(TryGetDynamicProperty(document, "Owner")),
                        DateCreated = SafeToString(TryGetDynamicProperty(document, "DateCreated")),
                        LastUpdated = SafeToString(TryGetDynamicProperty(document, "LastUpdated"))
                    });
                }

                return results
                    .OrderBy(document => document.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<DaoDocumentPropertyInfo> GetDocumentProperties(string containerName, string documentName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(containerName)) throw new ArgumentException("containerName is required.", nameof(containerName));
            if (string.IsNullOrWhiteSpace(documentName)) throw new ArgumentException("documentName is required.", nameof(documentName));

            return ExecuteComOperation(accessApp =>
            {
                var document = ResolveDaoDocument(accessApp, containerName, documentName);
                var properties = TryGetDynamicProperty(document, "Properties")
                    ?? throw new InvalidOperationException("DAO document properties collection is unavailable.");

                var results = new List<DaoDocumentPropertyInfo>();
                foreach (var property in properties)
                {
                    var name = SafeToString(TryGetDynamicProperty(property, "Name"));
                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    results.Add(new DaoDocumentPropertyInfo
                    {
                        ContainerName = containerName,
                        DocumentName = documentName,
                        Name = name,
                        TypeCode = ToInt32(TryGetDynamicProperty(property, "Type")),
                        Value = NormalizeValue(TryGetDynamicProperty(property, "Value"))
                    });
                }

                return results
                    .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public DaoDocumentPropertyInfo SetDocumentProperty(string containerName, string documentName, string propertyName, string value, string? propertyType = null, bool createIfMissing = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(containerName)) throw new ArgumentException("containerName is required.", nameof(containerName));
            if (string.IsNullOrWhiteSpace(documentName)) throw new ArgumentException("documentName is required.", nameof(documentName));
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("propertyName is required.", nameof(propertyName));

            return ExecuteComOperation(accessApp =>
            {
                var document = ResolveDaoDocument(accessApp, containerName, documentName);
                var existingProperty = FindDaoProperty(document, propertyName);
                var existingValue = existingProperty != null
                    ? TryGetDynamicProperty(existingProperty, "Value")
                    : null;

                var daoType = ParseDaoDataType(propertyType);
                var convertedValue = ConvertPropertyValue(value, propertyType, existingValue);
                SetDaoPropertyValue(document, propertyName, convertedValue, daoType, createIfMissing);

                var updatedProperty = FindDaoProperty(document, propertyName)
                    ?? throw new InvalidOperationException($"Property not found after update: {propertyName}");
                return new DaoDocumentPropertyInfo
                {
                    ContainerName = containerName,
                    DocumentName = documentName,
                    Name = SafeToString(TryGetDynamicProperty(updatedProperty, "Name")) ?? propertyName,
                    TypeCode = ToInt32(TryGetDynamicProperty(updatedProperty, "Type")),
                    Value = NormalizeValue(TryGetDynamicProperty(updatedProperty, "Value"))
                };
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetDisplayCategories(bool showCategories)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                _ = InvokeDynamicMethod(accessApp, "SetOption", "Show Navigation Pane Grouping Options", showCategories);
                _ = InvokeDynamicMethod(accessApp, "SetOption", "Show Navigation Pane Object Tabs", showCategories);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public string ExportDataMacroAxl(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));

            return ExecuteComOperation(accessApp =>
            {
                var axl = InvokeDynamicMethod(accessApp, "SaveAsAXL", 0, tableName);
                var text = SafeToString(axl);
                if (string.IsNullOrWhiteSpace(text))
                    throw new InvalidOperationException($"SaveAsAXL returned empty output for table '{tableName}'.");

                return text;
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void ImportDataMacroAxl(string tableName, string axlXml)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(axlXml)) throw new ArgumentException("AXL XML is required.", nameof(axlXml));

            ExecuteComOperation(accessApp =>
            {
                _ = InvokeDynamicMethod(accessApp, "LoadFromAXL", 0, tableName, axlXml);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public DatabaseSecurityInfo GetDatabaseSecurityInfo()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return new DatabaseSecurityInfo
            {
                DatabasePath = _currentDatabasePath ?? "",
                PasswordConfigured = !string.IsNullOrWhiteSpace(_databasePassword),
                Encrypted = !string.IsNullOrWhiteSpace(_databasePassword)
            };
        }

        public void SetDatabasePassword(string newPassword)
        {
            if (string.IsNullOrWhiteSpace(newPassword)) throw new ArgumentException("newPassword is required.", nameof(newPassword));
            ApplyDatabasePasswordInternal(newPassword);
        }

        public void RemoveDatabasePassword()
        {
            ApplyDatabasePasswordInternal(null);
        }

        public void EncryptDatabase(string? password = null)
        {
            var effectivePassword = string.IsNullOrWhiteSpace(password) ? _databasePassword : password;
            if (string.IsNullOrWhiteSpace(effectivePassword))
                throw new InvalidOperationException("A password is required to compact/encrypt the database.");

            ApplyDatabasePasswordInternal(effectivePassword);
        }

        public List<NavigationGroupInfo> GetNavigationGroups()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var groups = GetNavigationGroupsCollection(accessApp);
                var results = new List<NavigationGroupInfo>();
                if (groups == null)
                    return results;

                foreach (var group in groups)
                {
                    var name = SafeToString(TryGetDynamicProperty(group, "Name"));
                    var itemCount = ToNullableInt(TryGetDynamicProperty(TryGetDynamicProperty(group, "NavigationButtons"), "Count")) ?? 0;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        results.Add(new NavigationGroupInfo
                        {
                            Name = name,
                            ItemCount = itemCount
                        });
                    }
                }

                return results
                    .OrderBy(g => g.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void CreateNavigationGroup(string groupName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(groupName)) throw new ArgumentException("groupName is required.", nameof(groupName));

            ExecuteComOperation(accessApp =>
            {
                var groups = GetNavigationGroupsCollection(accessApp)
                    ?? throw new InvalidOperationException("NavigationGroups collection is unavailable.");
                _ = InvokeDynamicMethod(groups, "Add", groupName);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void AddNavigationGroupObject(string groupName, string objectName, string? objectType = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(groupName)) throw new ArgumentException("groupName is required.", nameof(groupName));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));

            ExecuteComOperation(accessApp =>
            {
                var groups = GetNavigationGroupsCollection(accessApp)
                    ?? throw new InvalidOperationException("NavigationGroups collection is unavailable.");
                var targetGroup = FindNavigationGroup(groups, groupName)
                    ?? throw new InvalidOperationException($"Navigation group not found: {groupName}");

                var buttons = TryGetDynamicProperty(targetGroup, "NavigationButtons")
                    ?? throw new InvalidOperationException("NavigationButtons collection is unavailable.");
                _ = InvokeDynamicMethod(buttons, "Add", objectName, NormalizeDoCmdVariant(objectType));
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void DeleteNavigationGroup(string groupName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(groupName)) throw new ArgumentException("groupName is required.", nameof(groupName));

            ExecuteComOperation(accessApp =>
            {
                var groups = GetNavigationGroupsCollection(accessApp)
                    ?? throw new InvalidOperationException("NavigationGroups collection is unavailable.");
                var targetGroup = FindNavigationGroup(groups, groupName)
                    ?? throw new InvalidOperationException($"Navigation group not found: {groupName}");

                var actualName = SafeToString(TryGetDynamicProperty(targetGroup, "Name")) ?? groupName;
                var removed = false;
                try
                {
                    _ = InvokeDynamicMethod(groups, "Delete", actualName);
                    removed = true;
                }
                catch
                {
                    // Fall back to deleting from the group object directly.
                }

                if (!removed)
                {
                    try
                    {
                        _ = InvokeDynamicMethod(targetGroup, "Delete");
                        removed = true;
                    }
                    catch
                    {
                        // Fall through to final error below.
                    }
                }

                if (!removed)
                    throw new InvalidOperationException($"Failed to delete navigation group: {groupName}");
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void RemoveNavigationGroupObject(string groupName, string objectName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(groupName)) throw new ArgumentException("groupName is required.", nameof(groupName));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));

            ExecuteComOperation(accessApp =>
            {
                var groups = GetNavigationGroupsCollection(accessApp)
                    ?? throw new InvalidOperationException("NavigationGroups collection is unavailable.");
                var targetGroup = FindNavigationGroup(groups, groupName)
                    ?? throw new InvalidOperationException($"Navigation group not found: {groupName}");

                var buttons = TryGetDynamicProperty(targetGroup, "NavigationButtons")
                    ?? throw new InvalidOperationException("NavigationButtons collection is unavailable.");
                var targetButton = FindNavigationGroupObject(buttons, objectName)
                    ?? throw new InvalidOperationException($"Navigation object not found in group '{groupName}': {objectName}");

                var removed = false;
                var buttonName = SafeToString(TryGetDynamicProperty(targetButton, "Name")) ?? objectName;
                if (!string.IsNullOrWhiteSpace(buttonName))
                {
                    try
                    {
                        _ = InvokeDynamicMethod(buttons, "Delete", buttonName);
                        removed = true;
                    }
                    catch
                    {
                        // Fall through to additional deletion attempts.
                    }
                }

                if (!removed)
                {
                    var index = ToNullableInt(TryGetDynamicProperty(targetButton, "Index"));
                    if (index.HasValue)
                    {
                        try
                        {
                            _ = InvokeDynamicMethod(buttons, "Delete", index.Value);
                            removed = true;
                        }
                        catch
                        {
                            // Fall through to final attempt.
                        }
                    }
                }

                if (!removed)
                {
                    try
                    {
                        _ = InvokeDynamicMethod(targetButton, "Delete");
                        removed = true;
                    }
                    catch
                    {
                        // Fall through to final error below.
                    }
                }

                if (!removed)
                    throw new InvalidOperationException($"Failed to remove navigation object '{objectName}' from group '{groupName}'.");
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetNavigationPaneVisibility(bool visible)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            ExecuteComOperation(accessApp =>
            {
                var updated = false;
                try
                {
                    _ = InvokeDynamicMethod(accessApp, "SetOption", "Show Navigation Pane", visible);
                    updated = true;
                }
                catch
                {
                    // Fall through to alternate strategy.
                }

                if (!updated)
                {
                    var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                        ?? throw new InvalidOperationException("DoCmd object is unavailable.");
                    _ = InvokeDynamicMethod(doCmd, "SelectObject", Type.Missing, Type.Missing, visible);
                }
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<NavigationGroupObjectInfo> GetNavigationGroupObjects(string groupName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(groupName)) throw new ArgumentException("groupName is required.", nameof(groupName));

            return ExecuteComOperation(accessApp =>
            {
                var groups = GetNavigationGroupsCollection(accessApp)
                    ?? throw new InvalidOperationException("NavigationGroups collection is unavailable.");
                var targetGroup = FindNavigationGroup(groups, groupName)
                    ?? throw new InvalidOperationException($"Navigation group not found: {groupName}");
                var buttons = TryGetDynamicProperty(targetGroup, "NavigationButtons")
                    ?? throw new InvalidOperationException("NavigationButtons collection is unavailable.");

                var results = new List<NavigationGroupObjectInfo>();
                var fallbackIndex = 0;
                foreach (var button in buttons)
                {
                    fallbackIndex++;
                    var name = SafeToString(TryGetDynamicProperty(button, "Name"));
                    var objectName = SafeToString(TryGetDynamicProperty(button, "ObjectName"))
                        ?? SafeToString(TryGetDynamicProperty(button, "Object"));
                    var objectType = SafeToString(TryGetDynamicProperty(button, "ObjectType"))
                        ?? SafeToString(TryGetDynamicProperty(button, "Type"));
                    var index = ToNullableInt(TryGetDynamicProperty(button, "Index")) ?? fallbackIndex;

                    results.Add(new NavigationGroupObjectInfo
                    {
                        GroupName = groupName,
                        Name = name ?? objectName ?? "",
                        ObjectName = objectName ?? name ?? "",
                        ObjectType = objectType,
                        Index = index
                    });
                }

                return results
                    .OrderBy(item => item.Index)
                    .ThenBy(item => item.ObjectName, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<ConditionalFormattingRuleInfo> GetConditionalFormatting(string objectType, string objectName, string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("controlName is required.", nameof(controlName));

            return ExecuteComOperation(accessApp =>
            {
                var isReport = string.Equals(objectType, "report", StringComparison.OrdinalIgnoreCase);
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }

                try
                {
                    var control = GetControlByName(targetObject, controlName)
                        ?? throw new InvalidOperationException($"Control not found: {controlName}");
                    var formatConditions = TryGetDynamicProperty(control, "FormatConditions");
                    var results = new List<ConditionalFormattingRuleInfo>();
                    if (formatConditions == null)
                        return results;

                    var index = 0;
                    foreach (var condition in (dynamic)formatConditions)
                    {
                        index++;
                        results.Add(new ConditionalFormattingRuleInfo
                        {
                            Index = index,
                            TypeCode = ToInt32(TryGetDynamicProperty(condition, "Type")),
                            Expression1 = SafeToString(TryGetDynamicProperty(condition, "Expression1")),
                            Expression2 = SafeToString(TryGetDynamicProperty(condition, "Expression2")),
                            ForeColor = ToNullableInt(TryGetDynamicProperty(condition, "ForeColor")),
                            BackColor = ToNullableInt(TryGetDynamicProperty(condition, "BackColor")),
                            Enabled = ToBool(TryGetDynamicProperty(condition, "Enabled"), true)
                        });
                    }

                    return results;
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: true);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: true);
                    }
                }
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void AddConditionalFormatting(string objectType, string objectName, string controlName, string expression, int? foreColor = null, int? backColor = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("controlName is required.", nameof(controlName));
            if (string.IsNullOrWhiteSpace(expression)) throw new ArgumentException("expression is required.", nameof(expression));

            ExecuteComOperation(accessApp =>
            {
                var isReport = string.Equals(objectType, "report", StringComparison.OrdinalIgnoreCase);
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }

                try
                {
                    var control = GetControlByName(targetObject, controlName)
                        ?? throw new InvalidOperationException($"Control not found: {controlName}");
                    var formatConditions = TryGetDynamicProperty(control, "FormatConditions")
                        ?? throw new InvalidOperationException("FormatConditions collection is unavailable.");

                    var condition = InvokeDynamicMethod(formatConditions, "Add", 1, Type.Missing, expression);
                    if (condition == null)
                        throw new InvalidOperationException("Failed to create conditional formatting rule.");

                    if (foreColor.HasValue)
                        SetDynamicProperty(condition, "ForeColor", foreColor.Value);
                    if (backColor.HasValue)
                        SetDynamicProperty(condition, "BackColor", backColor.Value);
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: true);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: true);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteConditionalFormatting(string objectType, string objectName, string controlName, int ruleIndex)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("controlName is required.", nameof(controlName));
            if (ruleIndex <= 0) throw new ArgumentOutOfRangeException(nameof(ruleIndex), "ruleIndex must be 1 or greater.");

            ExecuteComOperation(accessApp =>
            {
                var isReport = string.Equals(objectType, "report", StringComparison.OrdinalIgnoreCase);
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }

                try
                {
                    var control = GetControlByName(targetObject, controlName)
                        ?? throw new InvalidOperationException($"Control not found: {controlName}");
                    var formatConditions = TryGetDynamicProperty(control, "FormatConditions")
                        ?? throw new InvalidOperationException("FormatConditions collection is unavailable.");

                    if (!TryDeleteFormatConditionByIndex(formatConditions, ruleIndex))
                        throw new InvalidOperationException($"Conditional formatting rule index {ruleIndex} was not found.");
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: true);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: true);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public ConditionalFormattingRuleInfo UpdateConditionalFormatting(string objectType, string objectName, string controlName, int ruleIndex, string? expression = null, int? foreColor = null, int? backColor = null, bool? enabled = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("controlName is required.", nameof(controlName));
            if (ruleIndex <= 0) throw new ArgumentOutOfRangeException(nameof(ruleIndex), "ruleIndex must be 1 or greater.");

            return ExecuteComOperation(accessApp =>
            {
                var isReport = string.Equals(objectType, "report", StringComparison.OrdinalIgnoreCase);
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }

                try
                {
                    var control = GetControlByName(targetObject, controlName)
                        ?? throw new InvalidOperationException($"Control not found: {controlName}");
                    var formatConditions = TryGetDynamicProperty(control, "FormatConditions")
                        ?? throw new InvalidOperationException("FormatConditions collection is unavailable.");
                    var condition = GetFormatConditionByIndex(formatConditions, ruleIndex)
                        ?? throw new InvalidOperationException($"Conditional formatting rule index {ruleIndex} was not found.");

                    if (expression != null)
                        SetDynamicProperty(condition, "Expression1", expression);
                    if (foreColor.HasValue)
                        SetDynamicProperty(condition, "ForeColor", foreColor.Value);
                    if (backColor.HasValue)
                        SetDynamicProperty(condition, "BackColor", backColor.Value);
                    if (enabled.HasValue)
                        SetDynamicProperty(condition, "Enabled", enabled.Value);

                    return BuildConditionalFormattingRuleInfo(condition, ruleIndex);
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: true);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: true);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void ClearConditionalFormatting(string objectType, string objectName, string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("controlName is required.", nameof(controlName));

            ExecuteComOperation(accessApp =>
            {
                var isReport = string.Equals(objectType, "report", StringComparison.OrdinalIgnoreCase);
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }

                try
                {
                    var control = GetControlByName(targetObject, controlName)
                        ?? throw new InvalidOperationException($"Control not found: {controlName}");
                    var formatConditions = TryGetDynamicProperty(control, "FormatConditions")
                        ?? throw new InvalidOperationException("FormatConditions collection is unavailable.");

                    while (ToInt32(TryGetDynamicProperty(formatConditions, "Count")) > 0)
                    {
                        if (!TryDeleteFormatConditionByIndex(formatConditions, 1))
                        {
                            _ = InvokeDynamicMethod(formatConditions, "Delete");
                            break;
                        }
                    }
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: true);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: true);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public List<ConditionalFormattingControlInfo> ListAllConditionalFormats(string objectType, string objectName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));

            return ExecuteComOperation(accessApp =>
            {
                var isReport = string.Equals(objectType, "report", StringComparison.OrdinalIgnoreCase);
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }

                try
                {
                    var results = new List<ConditionalFormattingControlInfo>();
                    foreach (var control in GetControlObjects(targetObject))
                    {
                        var controlName = SafeToString(TryGetDynamicProperty(control, "Name"));
                        if (string.IsNullOrWhiteSpace(controlName))
                            continue;

                        var formatConditions = TryGetDynamicProperty(control, "FormatConditions");
                        if (formatConditions == null || ToInt32(TryGetDynamicProperty(formatConditions, "Count")) <= 0)
                            continue;

                        var rules = new List<ConditionalFormattingRuleInfo>();
                        var index = 0;
                        foreach (var condition in (dynamic)formatConditions)
                        {
                            index++;
                            rules.Add(BuildConditionalFormattingRuleInfo(condition, index));
                        }

                        if (rules.Count == 0)
                            continue;

                        results.Add(new ConditionalFormattingControlInfo
                        {
                            ObjectType = isReport ? "report" : "form",
                            ObjectName = objectName,
                            ControlName = controlName,
                            RuleCount = rules.Count,
                            Rules = rules
                        });
                    }

                    return results
                        .OrderBy(result => result.ControlName, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: false);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: false);
                    }
                }
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<AttachmentFileInfo> GetAttachmentFieldFiles(string tableName, string fieldName, string? whereCondition = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("tableName is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("fieldName is required.", nameof(fieldName));

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var sql = BuildAttachmentQuery(tableName, whereCondition);
                var recordset = InvokeDynamicMethod(currentDb, "OpenRecordset", sql)
                    ?? throw new InvalidOperationException("Failed to open attachment source recordset.");

                var files = new List<AttachmentFileInfo>();
                while (!ToBool(TryGetDynamicProperty(recordset, "EOF"), true))
                {
                    var attachmentRecordset = TryGetAttachmentRecordset(recordset, fieldName);
                    if (attachmentRecordset != null)
                    {
                        while (!ToBool(TryGetDynamicProperty(attachmentRecordset, "EOF"), true))
                        {
                            files.Add(new AttachmentFileInfo
                            {
                                FileName = SafeToString(GetRecordsetFieldValue(attachmentRecordset, "FileName")) ?? "",
                                FileType = SafeToString(GetRecordsetFieldValue(attachmentRecordset, "FileType")),
                                FileSize = ToNullableInt(GetRecordsetFieldValue(attachmentRecordset, "FileData")) is int sizeFromData && sizeFromData > 0
                                    ? sizeFromData
                                    : ToNullableInt(GetRecordsetFieldValue(attachmentRecordset, "FileSize"))
                            });
                            _ = InvokeDynamicMethod(attachmentRecordset, "MoveNext");
                        }
                    }

                    _ = InvokeDynamicMethod(recordset, "MoveNext");
                }

                return files;
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void AddAttachmentFile(string tableName, string fieldName, string filePath, string? whereCondition = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("tableName is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("fieldName is required.", nameof(fieldName));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException($"Attachment file not found: {filePath}");

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var sql = BuildAttachmentQuery(tableName, whereCondition);
                var recordset = InvokeDynamicMethod(currentDb, "OpenRecordset", sql, 2)
                    ?? throw new InvalidOperationException("Failed to open attachment source recordset.");

                if (ToBool(TryGetDynamicProperty(recordset, "EOF"), true))
                    throw new InvalidOperationException("No matching row was found for attachment update.");

                var attachmentRecordset = TryGetAttachmentRecordset(recordset, fieldName)
                    ?? throw new InvalidOperationException($"Attachment field not found or not accessible: {fieldName}");

                _ = InvokeDynamicMethod(attachmentRecordset, "AddNew");
                SetRecordsetFieldValue(attachmentRecordset, "FileData", File.ReadAllBytes(filePath));
                SetRecordsetFieldValue(attachmentRecordset, "FileName", Path.GetFileName(filePath));
                _ = InvokeDynamicMethod(attachmentRecordset, "Update");
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void RemoveAttachmentFile(string tableName, string fieldName, string fileName, string? whereCondition = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("tableName is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("fieldName is required.", nameof(fieldName));
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("fileName is required.", nameof(fileName));

            ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var sql = BuildAttachmentQuery(tableName, whereCondition);
                var recordset = InvokeDynamicMethod(currentDb, "OpenRecordset", sql, 2)
                    ?? throw new InvalidOperationException("Failed to open attachment source recordset.");

                if (ToBool(TryGetDynamicProperty(recordset, "EOF"), true))
                    throw new InvalidOperationException("No matching row was found for attachment update.");

                var attachmentRecordset = TryGetAttachmentRecordset(recordset, fieldName)
                    ?? throw new InvalidOperationException($"Attachment field not found or not accessible: {fieldName}");

                var removed = false;
                while (!ToBool(TryGetDynamicProperty(attachmentRecordset, "EOF"), true))
                {
                    var currentFileName = SafeToString(GetRecordsetFieldValue(attachmentRecordset, "FileName"));
                    if (string.Equals(currentFileName, fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        _ = InvokeDynamicMethod(attachmentRecordset, "Delete");
                        removed = true;
                        break;
                    }

                    _ = InvokeDynamicMethod(attachmentRecordset, "MoveNext");
                }

                if (!removed)
                    throw new InvalidOperationException($"Attachment file not found: {fileName}");
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public AttachmentSaveResult SaveAttachmentToDisk(string tableName, string fieldName, string filePath, string? fileName = null, string? whereCondition = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("tableName is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("fieldName is required.", nameof(fieldName));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var sql = BuildAttachmentQuery(tableName, whereCondition);
                var recordset = InvokeDynamicMethod(currentDb, "OpenRecordset", sql)
                    ?? throw new InvalidOperationException("Failed to open attachment source recordset.");

                while (!ToBool(TryGetDynamicProperty(recordset, "EOF"), true))
                {
                    var attachmentRecordset = TryGetAttachmentRecordset(recordset, fieldName);
                    if (attachmentRecordset != null)
                    {
                        while (!ToBool(TryGetDynamicProperty(attachmentRecordset, "EOF"), true))
                        {
                            var currentFileName = SafeToString(GetRecordsetFieldValue(attachmentRecordset, "FileName")) ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(fileName) || string.Equals(currentFileName, fileName, StringComparison.OrdinalIgnoreCase))
                            {
                                var fileData = GetRecordsetFieldValue(attachmentRecordset, "FileData");
                                var bytes = ConvertAttachmentFieldToBytes(fileData);
                                if (bytes == null || bytes.Length == 0)
                                    throw new InvalidOperationException($"Attachment '{currentFileName}' has no file data.");

                                var fullPath = Path.GetFullPath(filePath);
                                var directory = Path.GetDirectoryName(fullPath);
                                if (!string.IsNullOrWhiteSpace(directory))
                                    Directory.CreateDirectory(directory);

                                File.WriteAllBytes(fullPath, bytes);

                                return new AttachmentSaveResult
                                {
                                    TableName = tableName,
                                    FieldName = fieldName,
                                    FileName = string.IsNullOrWhiteSpace(currentFileName) ? Path.GetFileName(fullPath) : currentFileName,
                                    FilePath = fullPath,
                                    FileSize = bytes.Length
                                };
                            }

                            _ = InvokeDynamicMethod(attachmentRecordset, "MoveNext");
                        }
                    }

                    _ = InvokeDynamicMethod(recordset, "MoveNext");
                }

                if (!string.IsNullOrWhiteSpace(fileName))
                    throw new InvalidOperationException($"Attachment file not found: {fileName}");

                throw new InvalidOperationException("No attachment file was available to save.");
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<AttachmentMetadataInfo> GetAttachmentMetadata(string tableName, string fieldName, string? whereCondition = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("tableName is required.", nameof(tableName));
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("fieldName is required.", nameof(fieldName));

            return ExecuteComOperation(accessApp =>
            {
                var currentDb = TryGetCurrentDb(accessApp)
                    ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
                var sql = BuildAttachmentQuery(tableName, whereCondition);
                var recordset = InvokeDynamicMethod(currentDb, "OpenRecordset", sql)
                    ?? throw new InvalidOperationException("Failed to open attachment source recordset.");

                var results = new List<AttachmentMetadataInfo>();
                while (!ToBool(TryGetDynamicProperty(recordset, "EOF"), true))
                {
                    var attachmentRecordset = TryGetAttachmentRecordset(recordset, fieldName);
                    if (attachmentRecordset != null)
                    {
                        while (!ToBool(TryGetDynamicProperty(attachmentRecordset, "EOF"), true))
                        {
                            var item = new AttachmentMetadataInfo
                            {
                                FileName = SafeToString(GetRecordsetFieldValue(attachmentRecordset, "FileName")) ?? "",
                                FileType = SafeToString(GetRecordsetFieldValue(attachmentRecordset, "FileType")),
                                FileUrl = SafeToString(GetRecordsetFieldValue(attachmentRecordset, "FileURL")),
                                FileSize = ToNullableInt(GetRecordsetFieldValue(attachmentRecordset, "FileSize")),
                                Fields = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                            };

                            var fields = TryGetDynamicProperty(attachmentRecordset, "Fields");
                            if (fields != null)
                            {
                                foreach (var field in fields)
                                {
                                    var currentName = SafeToString(TryGetDynamicProperty(field, "Name"));
                                    if (string.IsNullOrWhiteSpace(currentName))
                                        continue;

                                    item.Fields[currentName] = NormalizeValue(TryGetDynamicProperty(field, "Value"));
                                }
                            }

                            if (!item.FileSize.HasValue && item.Fields.TryGetValue("FileData", out var dataValue))
                            {
                                var bytes = ConvertAttachmentFieldToBytes(dataValue);
                                if (bytes != null)
                                    item.FileSize = bytes.Length;
                            }

                            results.Add(item);
                            _ = InvokeDynamicMethod(attachmentRecordset, "MoveNext");
                        }
                    }

                    _ = InvokeDynamicMethod(recordset, "MoveNext");
                }

                return results;
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<ObjectEventInfo> GetObjectEvents(string objectType, string objectName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));

            var normalizedType = NormalizeEnumToken(objectType);
            string sourceText = normalizedType switch
            {
                "form" => ExportFormToText(objectName, TextModeAccessText),
                "report" => ExportReportToText(objectName, TextModeAccessText),
                _ => throw new ArgumentException("objectType must be form or report.", nameof(objectType))
            };

            var events = new List<ObjectEventInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in Regex.Matches(sourceText, "^\\s*(?<event>On[A-Za-z0-9_]+)\\s*=\\s*\"(?<handler>[^\"]*)\"\\s*$", RegexOptions.Multiline))
            {
                var eventName = match.Groups["event"].Value.Trim();
                var handler = match.Groups["handler"].Value.Trim();
                if (string.IsNullOrWhiteSpace(eventName) || !seen.Add(eventName))
                    continue;

                events.Add(new ObjectEventInfo
                {
                    EventName = eventName,
                    Handler = handler
                });
            }

            return events
                .OrderBy(e => e.EventName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public void SetObjectEvent(string objectType, string objectName, string eventName, string eventValue)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("objectType is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("objectName is required.", nameof(objectName));
            if (string.IsNullOrWhiteSpace(eventName)) throw new ArgumentException("eventName is required.", nameof(eventName));

            ExecuteComOperation(accessApp =>
            {
                var normalizedType = NormalizeEnumToken(objectType);
                if (normalizedType == "form")
                {
                    bool openedHere;
                    var form = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                    SetDynamicProperty(form, eventName, eventValue);
                    accessApp.DoCmd.Save(2, objectName); // 2 = acForm
                    if (openedHere)
                        CloseFormInternal(accessApp, objectName, saveChanges: true);
                    return;
                }

                if (normalizedType == "report")
                {
                    bool openedHere;
                    var report = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                    SetDynamicProperty(report, eventName, eventValue);
                    accessApp.DoCmd.Save(3, objectName); // 3 = acReport
                    if (openedHere)
                        CloseReportInternal(accessApp, objectName, saveChanges: true);
                    return;
                }

                throw new ArgumentException("objectType must be form or report.", nameof(objectType));
            },
            requireExclusive: true,
            releaseOleDb: true);
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
                using var command = CreateCommand("SELECT Name FROM MSysObjects WHERE Type = -32768");
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
                using var command = CreateCommand("SELECT Name FROM MSysObjects WHERE Type = -32764");
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
                using var command = CreateCommand("SELECT Name FROM MSysObjects WHERE Type = -32766");
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
                using var command = CreateCommand("SELECT Name FROM MSysObjects WHERE Type = -32761");
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
                accessApp => accessApp.DoCmd.Close(2, formName, 2), // 2 = acSaveNo
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
                accessApp => accessApp.DoCmd.Close(3, reportName, 2), // 2 = acSaveNo
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
                using var command = CreateCommand("SELECT Name FROM MSysObjects WHERE Type = -32761");
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

        public ModuleAnalysisInfo GetModuleInfo(string? projectName, string moduleName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));

            return ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                var lineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfLines"));
                var declarationLineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfDeclarationLines"));
                var code = lineCount > 0
                    ? SafeToString(TryGetDynamicProperty(codeModule, "Lines", 1, lineCount)) ?? string.Empty
                    : string.Empty;
                var procedures = ParseModuleProcedures(code);

                return new ModuleAnalysisInfo
                {
                    ProjectName = SafeToString(TryGetDynamicProperty(TryGetDynamicProperty(component, "Collection"), "Parent")) ?? (projectName ?? "CurrentProject"),
                    ModuleName = moduleName,
                    ModuleType = MapVbComponentType(ToInt32(TryGetDynamicProperty(component, "Type"))),
                    LineCount = lineCount,
                    DeclarationLineCount = declarationLineCount,
                    ProcedureCount = procedures.Count
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<ModuleProcedureInfo> ListProcedures(string? projectName, string moduleName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));

            return ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                var lineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfLines"));
                if (lineCount <= 0)
                    return new List<ModuleProcedureInfo>();

                var code = SafeToString(TryGetDynamicProperty(codeModule, "Lines", 1, lineCount)) ?? string.Empty;
                return ParseModuleProcedures(code);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public string GetProcedureCode(string? projectName, string moduleName, string procedureName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));
            if (string.IsNullOrWhiteSpace(procedureName)) throw new ArgumentException("Procedure name is required.", nameof(procedureName));

            return ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                var lineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfLines"));
                if (lineCount <= 0)
                    return string.Empty;

                var code = SafeToString(TryGetDynamicProperty(codeModule, "Lines", 1, lineCount)) ?? string.Empty;
                var procedures = ParseModuleProcedures(code);
                ModuleProcedureInfo? procedure = null;
                foreach (var candidate in procedures)
                {
                    if (string.Equals(candidate.Name, procedureName, StringComparison.OrdinalIgnoreCase))
                    {
                        procedure = candidate;
                        break;
                    }
                }

                if (procedure == null)
                    throw new InvalidOperationException($"Procedure '{procedureName}' was not found in module '{moduleName}'.");

                return SafeToString(TryGetDynamicProperty(codeModule, "Lines", procedure.StartLine, procedure.LineCount)) ?? string.Empty;
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public string GetModuleDeclarations(string? projectName, string moduleName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));

            return ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                var declarationLineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfDeclarationLines"));
                if (declarationLineCount <= 0)
                    return string.Empty;

                return SafeToString(TryGetDynamicProperty(codeModule, "Lines", 1, declarationLineCount)) ?? string.Empty;
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void InsertLines(string? projectName, string moduleName, int lineNumber, string code)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));
            if (lineNumber <= 0) throw new ArgumentException("lineNumber must be greater than zero.", nameof(lineNumber));
            if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Code is required.", nameof(code));

            ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                _ = InvokeDynamicMethod(codeModule, "InsertLines", lineNumber, NormalizeLineEndings(code));
                TrySaveModule(accessApp, moduleName);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteLines(string? projectName, string moduleName, int startLine, int lineCount = 1)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));
            if (startLine <= 0) throw new ArgumentException("startLine must be greater than zero.", nameof(startLine));
            if (lineCount <= 0) throw new ArgumentException("lineCount must be greater than zero.", nameof(lineCount));

            ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                _ = InvokeDynamicMethod(codeModule, "DeleteLines", startLine, lineCount);
                TrySaveModule(accessApp, moduleName);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void ReplaceLine(string? projectName, string moduleName, int lineNumber, string code)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));
            if (lineNumber <= 0) throw new ArgumentException("lineNumber must be greater than zero.", nameof(lineNumber));
            if (code == null) throw new ArgumentNullException(nameof(code));

            ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                _ = InvokeDynamicMethod(codeModule, "ReplaceLine", lineNumber, NormalizeLineEndings(code));
                TrySaveModule(accessApp, moduleName);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public ModuleFindResult FindTextInModule(
            string? projectName,
            string moduleName,
            string findText,
            int startLine = 1,
            int startColumn = 1,
            int? endLine = null,
            int? endColumn = null,
            bool wholeWord = false,
            bool matchCase = false,
            bool patternSearch = false)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(moduleName)) throw new ArgumentException("Module name is required.", nameof(moduleName));
            if (string.IsNullOrWhiteSpace(findText)) throw new ArgumentException("Find text is required.", nameof(findText));
            if (startLine <= 0) throw new ArgumentException("startLine must be greater than zero.", nameof(startLine));
            if (startColumn <= 0) throw new ArgumentException("startColumn must be greater than zero.", nameof(startColumn));

            return ExecuteComOperation(accessApp =>
            {
                var component = FindOrCreateVbComponent(accessApp, projectName ?? "CurrentProject", moduleName, false)
                    ?? throw new InvalidOperationException($"VBA module '{moduleName}' was not found.");
                var codeModule = TryGetDynamicProperty(component, "CodeModule")
                    ?? throw new InvalidOperationException($"Code module for '{moduleName}' is not accessible.");

                var totalLineCount = ToInt32(TryGetDynamicProperty(codeModule, "CountOfLines"));
                var effectiveEndLine = endLine ?? Math.Max(totalLineCount, startLine);
                var effectiveEndColumn = endColumn ?? 255;

                var target = (object)codeModule;
                var methodArgs = new object?[]
                {
                    findText,
                    startLine,
                    startColumn,
                    effectiveEndLine,
                    effectiveEndColumn,
                    wholeWord,
                    matchCase,
                    patternSearch
                };

                var parameterModifier = new System.Reflection.ParameterModifier(methodArgs.Length);
                parameterModifier[1] = true;
                parameterModifier[2] = true;
                parameterModifier[3] = true;
                parameterModifier[4] = true;

                var foundObject = target.GetType().InvokeMember(
                    "Find",
                    System.Reflection.BindingFlags.InvokeMethod,
                    binder: null,
                    target: target,
                    args: methodArgs,
                    modifiers: new[] { parameterModifier },
                    culture: null,
                    namedParameters: null);

                var found = foundObject is bool boolResult && boolResult;
                return new ModuleFindResult
                {
                    Found = found,
                    StartLine = found ? ToInt32(methodArgs[1]) : null,
                    StartColumn = found ? ToInt32(methodArgs[2]) : null,
                    EndLine = found ? ToInt32(methodArgs[3]) : null,
                    EndColumn = found ? ToInt32(methodArgs[4]) : null
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<ImportExportSpecificationInfo> ListImportExportSpecs()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var specs = GetImportExportSpecificationsCollection(accessApp)
                    ?? throw new InvalidOperationException("ImportExportSpecifications collection is unavailable.");

                var results = new List<ImportExportSpecificationInfo>();
                foreach (var spec in specs)
                {
                    results.Add(new ImportExportSpecificationInfo
                    {
                        Name = SafeToString(TryGetDynamicProperty(spec, "Name")) ?? string.Empty,
                        Description = SafeToString(TryGetDynamicProperty(spec, "Description")),
                        Xml = SafeToString(TryGetDynamicProperty(spec, "XML"))
                    });
                }

                return results
                    .Where(spec => !string.IsNullOrWhiteSpace(spec.Name))
                    .OrderBy(spec => spec.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public ImportExportSpecificationInfo GetImportExportSpec(string specificationName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(specificationName)) throw new ArgumentException("Specification name is required.", nameof(specificationName));

            return ExecuteComOperation(accessApp =>
            {
                var specs = GetImportExportSpecificationsCollection(accessApp)
                    ?? throw new InvalidOperationException("ImportExportSpecifications collection is unavailable.");
                var spec = FindImportExportSpecification(specs, specificationName)
                    ?? throw new InvalidOperationException($"Import/export specification not found: {specificationName}");

                return new ImportExportSpecificationInfo
                {
                    Name = SafeToString(TryGetDynamicProperty(spec, "Name")) ?? specificationName,
                    Description = SafeToString(TryGetDynamicProperty(spec, "Description")),
                    Xml = SafeToString(TryGetDynamicProperty(spec, "XML"))
                };
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void CreateImportExportSpec(string specificationName, string specificationXml)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(specificationName)) throw new ArgumentException("Specification name is required.", nameof(specificationName));
            if (string.IsNullOrWhiteSpace(specificationXml)) throw new ArgumentException("Specification XML is required.", nameof(specificationXml));

            ExecuteComOperation(accessApp =>
            {
                var specs = GetImportExportSpecificationsCollection(accessApp)
                    ?? throw new InvalidOperationException("ImportExportSpecifications collection is unavailable.");

                var existing = FindImportExportSpecification(specs, specificationName);
                if (existing != null)
                    _ = InvokeDynamicMethod(existing, "Delete");

                _ = InvokeDynamicMethod(specs, "Add", specificationName.Trim(), specificationXml);
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteImportExportSpec(string specificationName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(specificationName)) throw new ArgumentException("Specification name is required.", nameof(specificationName));

            ExecuteComOperation(accessApp =>
            {
                var specs = GetImportExportSpecificationsCollection(accessApp)
                    ?? throw new InvalidOperationException("ImportExportSpecifications collection is unavailable.");
                var spec = FindImportExportSpecification(specs, specificationName)
                    ?? throw new InvalidOperationException($"Import/export specification not found: {specificationName}");

                _ = InvokeDynamicMethod(spec, "Delete");
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void RunImportExportSpec(string specificationName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(specificationName)) throw new ArgumentException("Specification name is required.", nameof(specificationName));

            ExecuteComOperation(accessApp =>
            {
                var doCmd = TryGetDynamicProperty(accessApp, "DoCmd")
                    ?? throw new InvalidOperationException("DoCmd is unavailable on the Access application instance.");
                _ = InvokeDynamicMethod(doCmd, "RunSavedImportExport", specificationName.Trim());
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        #endregion

        #region 5. System Table Metadata Access

        public List<SystemTableInfo> GetSystemTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            EnsureOleDbConnection();

            var systemTables = new List<SystemTableInfo>();
            var schema = GetSchema("Tables");
            
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
                using var command = CreateCommand("SELECT * FROM MSysObjects");
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
                using var command = CreateCommand("SELECT COUNT(*) FROM MSysObjects WHERE Name = ? AND Type = -32768");
                AddCommandParameter(command, "@Name", formName);
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
                var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
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

        public List<SectionInfo> GetFormSections(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required", nameof(formName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    return GetSectionObjects((object)form)
                        .Select((section, index) => BuildSectionInfo(section, index, isReport: false))
                        .OrderBy(s => s.Index)
                        .ToList();
                }
                finally
                {
                    if (openedHere)
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public List<SectionInfo> GetReportSections(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required", nameof(reportName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    return GetSectionObjects((object)report)
                        .Select((section, index) => BuildSectionInfo(section, index, isReport: true))
                        .OrderBy(s => s.Index)
                        .ToList();
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetSectionProperty(string objectType, string objectName, string section, string propertyName, object value)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("Object type is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("Object name is required.", nameof(objectName));
            if (string.IsNullOrWhiteSpace(section)) throw new ArgumentException("Section is required.", nameof(section));
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property name is required.", nameof(propertyName));

            ExecuteComOperation(accessApp =>
            {
                var normalizedType = NormalizeEnumToken(objectType);
                var isReport = normalizedType == "report" || normalizedType == "acreport";
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else if (normalizedType == "form" || normalizedType == "acform")
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    throw new ArgumentException("objectType must be form or report.", nameof(objectType));
                }

                try
                {
                    var targetSection = FindSection(targetObject, section)
                        ?? throw new InvalidOperationException($"Section '{section}' was not found on {objectType} '{objectName}'.");

                    var existingValue = TryGetDynamicProperty(targetSection, propertyName);
                    var convertedValue = ConvertValueForProperty(value, existingValue);
                    SetDynamicProperty(targetSection, propertyName, convertedValue);
                    accessApp.DoCmd.Save(isReport ? 3 : 2, objectName);
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: false);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: false);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public ControlInfo CreateControl(
            string formName,
            string controlType,
            string? controlName = null,
            int section = 0,
            string? parentControlName = null,
            string? columnName = null,
            int? left = null,
            int? top = null,
            int? width = null,
            int? height = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required.", nameof(formName));
            if (string.IsNullOrWhiteSpace(controlType)) throw new ArgumentException("Control type is required.", nameof(controlType));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    var created = accessApp.CreateControl(
                        formName,
                        MapControlTypeToConstant(controlType),
                        section,
                        string.IsNullOrWhiteSpace(parentControlName) ? Type.Missing : parentControlName.Trim(),
                        string.IsNullOrWhiteSpace(columnName) ? Type.Missing : columnName.Trim(),
                        left.HasValue ? left.Value : 1000,
                        top.HasValue ? top.Value : 1000,
                        width.HasValue ? width.Value : 1500,
                        height.HasValue ? height.Value : 300);

                    if (!string.IsNullOrWhiteSpace(controlName))
                        SetDynamicProperty(created, "Name", controlName.Trim());

                    accessApp.DoCmd.Save(2, formName);
                    return BuildControlInfo((object)created);
                }
                finally
                {
                    if (openedHere)
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public ControlInfo CreateReportControl(
            string reportName,
            string controlType,
            string? controlName = null,
            int section = 0,
            string? parentControlName = null,
            string? columnName = null,
            int? left = null,
            int? top = null,
            int? width = null,
            int? height = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required.", nameof(reportName));
            if (string.IsNullOrWhiteSpace(controlType)) throw new ArgumentException("Control type is required.", nameof(controlType));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    var created = accessApp.CreateReportControl(
                        reportName,
                        MapControlTypeToConstant(controlType),
                        section,
                        string.IsNullOrWhiteSpace(parentControlName) ? Type.Missing : parentControlName.Trim(),
                        string.IsNullOrWhiteSpace(columnName) ? Type.Missing : columnName.Trim(),
                        left.HasValue ? left.Value : 1000,
                        top.HasValue ? top.Value : 1000,
                        width.HasValue ? width.Value : 1500,
                        height.HasValue ? height.Value : 300);

                    if (!string.IsNullOrWhiteSpace(controlName))
                        SetDynamicProperty(created, "Name", controlName.Trim());

                    accessApp.DoCmd.Save(3, reportName);
                    return BuildControlInfo((object)created);
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteControl(string formName, string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required.", nameof(formName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("Control name is required.", nameof(controlName));

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    _ = InvokeDynamicMethod(accessApp, "DeleteControl", formName, controlName.Trim());
                    accessApp.DoCmd.Save(2, formName);
                }
                finally
                {
                    if (openedHere)
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteReportControl(string reportName, string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required.", nameof(reportName));
            if (string.IsNullOrWhiteSpace(controlName)) throw new ArgumentException("Control name is required.", nameof(controlName));

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    _ = InvokeDynamicMethod(accessApp, "DeleteReportControl", reportName, controlName.Trim());
                    accessApp.DoCmd.Save(3, reportName);
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public FormDesignPropertiesInfo GetFormProperties(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required.", nameof(formName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    return new FormDesignPropertiesInfo
                    {
                        FormName = formName,
                        RecordSource = SafeToString(TryGetDynamicProperty(form, "RecordSource")),
                        DefaultView = ToNullableInt(TryGetDynamicProperty(form, "DefaultView")),
                        AllowEdits = ToNullableBool(TryGetDynamicProperty(form, "AllowEdits")),
                        AllowAdditions = ToNullableBool(TryGetDynamicProperty(form, "AllowAdditions")),
                        AllowDeletions = ToNullableBool(TryGetDynamicProperty(form, "AllowDeletions")),
                        DataEntry = ToNullableBool(TryGetDynamicProperty(form, "DataEntry")),
                        NavigationButtons = ToNullableBool(TryGetDynamicProperty(form, "NavigationButtons")),
                        DividingLines = ToNullableBool(TryGetDynamicProperty(form, "DividingLines")),
                        ScrollBars = ToNullableInt(TryGetDynamicProperty(form, "ScrollBars")),
                        Caption = SafeToString(TryGetDynamicProperty(form, "Caption")),
                        Modal = ToNullableBool(TryGetDynamicProperty(form, "Modal")),
                        PopUp = ToNullableBool(TryGetDynamicProperty(form, "PopUp"))
                    };
                }
                finally
                {
                    if (openedHere)
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetFormProperty(string formName, string propertyName, object value)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required.", nameof(formName));
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property name is required.", nameof(propertyName));

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    var existingValue = TryGetDynamicProperty(form, propertyName);
                    var convertedValue = ConvertValueForProperty(value, existingValue);
                    SetDynamicProperty(form, propertyName, convertedValue);
                    accessApp.DoCmd.Save(2, formName);
                }
                finally
                {
                    if (openedHere)
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetFormRecordSource(string formName, string recordSource)
        {
            if (string.IsNullOrWhiteSpace(recordSource)) throw new ArgumentException("recordSource is required.", nameof(recordSource));
            SetFormProperty(formName, "RecordSource", recordSource);
        }

        public ReportDesignPropertiesInfo GetReportProperties(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required.", nameof(reportName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    return new ReportDesignPropertiesInfo
                    {
                        ReportName = reportName,
                        RecordSource = SafeToString(TryGetDynamicProperty(report, "RecordSource")),
                        DefaultView = ToNullableInt(TryGetDynamicProperty(report, "DefaultView")),
                        Caption = SafeToString(TryGetDynamicProperty(report, "Caption")),
                        ForceNewPage = ToNullableInt(TryGetDynamicProperty(report, "ForceNewPage")),
                        KeepTogether = ToNullableInt(TryGetDynamicProperty(report, "KeepTogether")),
                        PopUp = ToNullableBool(TryGetDynamicProperty(report, "PopUp")),
                        Modal = ToNullableBool(TryGetDynamicProperty(report, "Modal")),
                        ColumnCount = ToNullableInt(TryGetDynamicProperty(report, "ColumnCount")),
                        ColumnSpacing = ToNullableInt(TryGetDynamicProperty(report, "ColumnSpacing"))
                    };
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetReportRecordSource(string reportName, string recordSource)
        {
            if (string.IsNullOrWhiteSpace(recordSource)) throw new ArgumentException("recordSource is required.", nameof(recordSource));
            SetReportProperty(reportName, "RecordSource", recordSource);
        }

        public void SetReportProperty(string reportName, string propertyName, object value)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("Report name is required.", nameof(reportName));
            if (string.IsNullOrWhiteSpace(propertyName)) throw new ArgumentException("Property name is required.", nameof(propertyName));

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    var existingValue = TryGetDynamicProperty(report, propertyName);
                    var convertedValue = ConvertValueForProperty(value, existingValue);
                    SetDynamicProperty(report, propertyName, convertedValue);
                    accessApp.DoCmd.Save(3, reportName);
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public List<TabOrderEntryInfo> GetTabOrder(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required.", nameof(formName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    return GetControlObjects((object)form)
                        .Select(control => new TabOrderEntryInfo
                        {
                            ControlName = SafeToString(TryGetDynamicProperty(control, "Name")) ?? string.Empty,
                            TabIndex = ToInt32(TryGetDynamicProperty(control, "TabIndex")),
                            TabStop = ToBool(TryGetDynamicProperty(control, "TabStop"), true)
                        })
                        .Where(entry => !string.IsNullOrWhiteSpace(entry.ControlName))
                        .OrderBy(entry => entry.TabIndex)
                        .ThenBy(entry => entry.ControlName, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                }
                finally
                {
                    if (openedHere)
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetTabOrder(string formName, IEnumerable<string> controlNames)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(formName)) throw new ArgumentException("Form name is required.", nameof(formName));

            var normalizedControlNames = controlNames
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (normalizedControlNames.Count == 0)
                throw new ArgumentException("At least one control name is required.", nameof(controlNames));

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var form = EnsureFormOpen(accessApp, formName, true, out openedHere);
                try
                {
                    var controlsByName = GetControlObjects((object)form)
                        .Select(control => new
                        {
                            Name = SafeToString(TryGetDynamicProperty(control, "Name")),
                            Control = control
                        })
                        .Where(entry => !string.IsNullOrWhiteSpace(entry.Name))
                        .ToDictionary(entry => entry.Name!, entry => entry.Control, StringComparer.OrdinalIgnoreCase);

                    for (var i = 0; i < normalizedControlNames.Count; i++)
                    {
                        if (!controlsByName.TryGetValue(normalizedControlNames[i], out var control))
                            throw new InvalidOperationException($"Control '{normalizedControlNames[i]}' was not found on form '{formName}'.");

                        SetDynamicProperty(control, "TabIndex", i);
                    }

                    accessApp.DoCmd.Save(2, formName);
                }
                finally
                {
                    if (openedHere)
                        CloseFormInternal(accessApp, formName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public PageSetupInfo GetPageSetup(string objectType, string objectName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("Object type is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("Object name is required.", nameof(objectName));

            return ExecuteComOperation(accessApp =>
            {
                var normalizedType = NormalizeEnumToken(objectType);
                var isReport = normalizedType == "report" || normalizedType == "acreport";
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else if (normalizedType == "form" || normalizedType == "acform")
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    throw new ArgumentException("objectType must be form or report.", nameof(objectType));
                }

                try
                {
                    var printer = TryGetDynamicProperty(targetObject, "Printer");
                    return new PageSetupInfo
                    {
                        ObjectType = isReport ? "report" : "form",
                        ObjectName = objectName,
                        TopMargin = ToNullableInt(TryGetDynamicProperty(printer ?? targetObject, "TopMargin")),
                        BottomMargin = ToNullableInt(TryGetDynamicProperty(printer ?? targetObject, "BottomMargin")),
                        LeftMargin = ToNullableInt(TryGetDynamicProperty(printer ?? targetObject, "LeftMargin")),
                        RightMargin = ToNullableInt(TryGetDynamicProperty(printer ?? targetObject, "RightMargin")),
                        Orientation = ToNullableInt(TryGetDynamicProperty(printer ?? targetObject, "Orientation")),
                        PaperSize = ToNullableInt(TryGetDynamicProperty(printer ?? targetObject, "PaperSize")),
                        DataOnly = ToNullableBool(TryGetDynamicProperty(printer ?? targetObject, "DataOnly"))
                    };
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: false);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: false);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void SetPageSetup(string objectType, string objectName, int? topMargin = null, int? bottomMargin = null, int? leftMargin = null, int? rightMargin = null, int? orientation = null, int? paperSize = null, bool? dataOnly = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectType)) throw new ArgumentException("Object type is required.", nameof(objectType));
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("Object name is required.", nameof(objectName));

            ExecuteComOperation(accessApp =>
            {
                var normalizedType = NormalizeEnumToken(objectType);
                var isReport = normalizedType == "report" || normalizedType == "acreport";
                bool openedHere;
                object targetObject;
                if (isReport)
                {
                    targetObject = EnsureReportOpen(accessApp, objectName, true, out openedHere);
                }
                else if (normalizedType == "form" || normalizedType == "acform")
                {
                    targetObject = EnsureFormOpen(accessApp, objectName, true, out openedHere);
                }
                else
                {
                    throw new ArgumentException("objectType must be form or report.", nameof(objectType));
                }

                try
                {
                    var printer = TryGetDynamicProperty(targetObject, "Printer");
                    var pageSetupTarget = printer ?? targetObject;

                    if (topMargin.HasValue)
                        SetDynamicProperty(pageSetupTarget, "TopMargin", topMargin.Value);
                    if (bottomMargin.HasValue)
                        SetDynamicProperty(pageSetupTarget, "BottomMargin", bottomMargin.Value);
                    if (leftMargin.HasValue)
                        SetDynamicProperty(pageSetupTarget, "LeftMargin", leftMargin.Value);
                    if (rightMargin.HasValue)
                        SetDynamicProperty(pageSetupTarget, "RightMargin", rightMargin.Value);
                    if (orientation.HasValue)
                        SetDynamicProperty(pageSetupTarget, "Orientation", orientation.Value);
                    if (paperSize.HasValue)
                        SetDynamicProperty(pageSetupTarget, "PaperSize", paperSize.Value);
                    if (dataOnly.HasValue)
                        SetDynamicProperty(pageSetupTarget, "DataOnly", dataOnly.Value);

                    accessApp.DoCmd.Save(isReport ? 3 : 2, objectName);
                }
                finally
                {
                    if (openedHere)
                    {
                        if (isReport)
                            CloseReportInternal(accessApp, objectName, saveChanges: false);
                        else
                            CloseFormInternal(accessApp, objectName, saveChanges: false);
                    }
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public List<ReportGroupingInfo> GetReportGrouping(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("reportName is required.", nameof(reportName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    var results = new List<ReportGroupingInfo>();
                    for (var index = 0; index < 128; index++)
                    {
                        var groupLevel = TryGetReportGroupLevel(report, index);
                        if (groupLevel == null)
                            break;

                        results.Add(BuildReportGroupingInfo(groupLevel, index));
                    }

                    return results;
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public ReportGroupingInfo SetReportGrouping(string reportName, string? expression = null, int? index = null, int? sortOrder = null, int? groupOn = null, int? groupInterval = null, bool? groupHeader = null, bool? groupFooter = null, int? keepTogether = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("reportName is required.", nameof(reportName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    dynamic groupLevel;
                    int effectiveIndex;
                    if (index.HasValue)
                    {
                        if (index.Value < 0)
                            throw new ArgumentOutOfRangeException(nameof(index), "index must be greater than or equal to 0.");
                        groupLevel = TryGetReportGroupLevel(report, index.Value)
                            ?? throw new InvalidOperationException($"Group level index {index.Value} was not found.");
                        effectiveIndex = index.Value;
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(expression))
                            throw new ArgumentException("expression is required when creating a new group level.", nameof(expression));

                        var headerValue = groupHeader ?? true;
                        var footerValue = groupFooter ?? true;
                        groupLevel = InvokeDynamicMethod(accessApp, "CreateGroupLevel", reportName, expression, headerValue, footerValue)
                            ?? throw new InvalidOperationException("Failed to create report group level.");
                        effectiveIndex = ToNullableInt(TryGetDynamicProperty(groupLevel, "GroupLevel")) ??
                            ToNullableInt(TryGetDynamicProperty(groupLevel, "Index")) ??
                            Math.Max(0, ToInt32(TryGetDynamicProperty(TryGetDynamicProperty(report, "GroupLevels"), "Count")) - 1);
                    }

                    if (expression != null)
                        SetDynamicProperty(groupLevel, "ControlSource", expression);
                    if (sortOrder.HasValue)
                        SetDynamicProperty(groupLevel, "SortOrder", sortOrder.Value);
                    if (groupOn.HasValue)
                        SetDynamicProperty(groupLevel, "GroupOn", groupOn.Value);
                    if (groupInterval.HasValue)
                        SetDynamicProperty(groupLevel, "GroupInterval", groupInterval.Value);
                    if (groupHeader.HasValue)
                        SetDynamicProperty(groupLevel, "GroupHeader", groupHeader.Value);
                    if (groupFooter.HasValue)
                        SetDynamicProperty(groupLevel, "GroupFooter", groupFooter.Value);
                    if (keepTogether.HasValue)
                        SetDynamicProperty(groupLevel, "KeepTogether", keepTogether.Value);

                    accessApp.DoCmd.Save(3, reportName);
                    return BuildReportGroupingInfo(groupLevel, effectiveIndex);
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public void DeleteReportGrouping(string reportName, int index)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("reportName is required.", nameof(reportName));
            if (index < 0) throw new ArgumentOutOfRangeException(nameof(index), "index must be greater than or equal to 0.");

            ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    var groupLevel = TryGetReportGroupLevel(report, index)
                        ?? throw new InvalidOperationException($"Group level index {index} was not found.");
                    var deleted = false;
                    try
                    {
                        _ = InvokeDynamicMethod(groupLevel, "Delete");
                        deleted = true;
                    }
                    catch
                    {
                        // Fall through to legacy fallback.
                    }

                    if (!deleted)
                    {
                        var groupLevels = TryGetDynamicProperty(report, "GroupLevels");
                        if (groupLevels != null)
                        {
                            _ = InvokeDynamicMethod(groupLevels, "Delete", index);
                            deleted = true;
                        }
                    }

                    if (!deleted)
                        throw new InvalidOperationException($"Failed to delete group level index {index}.");

                    accessApp.DoCmd.Save(3, reportName);
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: true,
            releaseOleDb: true);
        }

        public ReportSortingInfo GetReportSorting(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(reportName)) throw new ArgumentException("reportName is required.", nameof(reportName));

            return ExecuteComOperation(accessApp =>
            {
                var openedHere = false;
                var report = EnsureReportOpen(accessApp, reportName, true, out openedHere);
                try
                {
                    var groupingLevels = new List<ReportGroupingInfo>();
                    for (var index = 0; index < 128; index++)
                    {
                        var groupLevel = TryGetReportGroupLevel(report, index);
                        if (groupLevel == null)
                            break;
                        groupingLevels.Add(BuildReportGroupingInfo(groupLevel, index));
                    }

                    return new ReportSortingInfo
                    {
                        ReportName = reportName,
                        OrderBy = SafeToString(TryGetDynamicProperty(report, "OrderBy")),
                        OrderByOn = ToNullableBool(TryGetDynamicProperty(report, "OrderByOn")),
                        GroupLevels = groupingLevels
                    };
                }
                finally
                {
                    if (openedHere)
                        CloseReportInternal(accessApp, reportName, saveChanges: false);
                }
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public PrinterInfoResult GetPrinterInfo()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var result = new PrinterInfoResult();
                var activePrinter = TryGetDynamicProperty(accessApp, "Printer");
                if (activePrinter != null)
                    result.ActivePrinter = BuildPrinterInfo(activePrinter);

                var printers = TryGetDynamicProperty(accessApp, "Printers");
                if (printers != null)
                {
                    foreach (var printer in printers)
                    {
                        result.Printers.Add(BuildPrinterInfo(printer));
                    }
                }

                result.Printers = result.Printers
                    .OrderBy(p => p.DeviceName, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                return result;
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public List<VBAReferenceInfo> GetVbaReferences(string? projectName = null)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            return ExecuteComOperation(accessApp =>
            {
                var project = FindVbProject(accessApp, projectName)
                    ?? throw new InvalidOperationException("No VBA project is available in the current Access database.");
                var references = TryGetDynamicProperty(project, "References")
                    ?? throw new InvalidOperationException("VBA references collection is unavailable.");

                var results = new List<VBAReferenceInfo>();
                foreach (var reference in references)
                {
                    results.Add(new VBAReferenceInfo
                    {
                        Name = SafeToString(TryGetDynamicProperty(reference, "Name")) ?? "",
                        Guid = SafeToString(TryGetDynamicProperty(reference, "Guid")) ?? "",
                        Major = ToInt32(TryGetDynamicProperty(reference, "Major")),
                        Minor = ToInt32(TryGetDynamicProperty(reference, "Minor")),
                        FullPath = SafeToString(TryGetDynamicProperty(reference, "FullPath")) ?? "",
                        Description = SafeToString(TryGetDynamicProperty(reference, "Description")) ?? "",
                        BuiltIn = ToBool(TryGetDynamicProperty(reference, "BuiltIn"), false),
                        IsBroken = ToBool(TryGetDynamicProperty(reference, "IsBroken"), false)
                    });
                }

                return results
                    .OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void AddVbaReference(string? projectName, string? referencePath, string? referenceGuid, int major = 1, int minor = 0)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(referencePath) && string.IsNullOrWhiteSpace(referenceGuid))
                throw new ArgumentException("Either referencePath or referenceGuid must be provided.");

            ExecuteComOperation(accessApp =>
            {
                var project = FindVbProject(accessApp, projectName)
                    ?? throw new InvalidOperationException("No VBA project is available in the current Access database.");
                var references = TryGetDynamicProperty(project, "References")
                    ?? throw new InvalidOperationException("VBA references collection is unavailable.");

                if (!string.IsNullOrWhiteSpace(referencePath))
                {
                    _ = InvokeDynamicMethod(references, "AddFromFile", referencePath.Trim());
                    return;
                }

                _ = InvokeDynamicMethod(references, "AddFromGuid", referenceGuid!.Trim(), major, minor);
            },
            requireExclusive: false,
            releaseOleDb: false);
        }

        public void RemoveVbaReference(string? projectName, string referenceIdentifier)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(referenceIdentifier)) throw new ArgumentException("referenceIdentifier is required.", nameof(referenceIdentifier));

            ExecuteComOperation(accessApp =>
            {
                var project = FindVbProject(accessApp, projectName)
                    ?? throw new InvalidOperationException("No VBA project is available in the current Access database.");
                var references = TryGetDynamicProperty(project, "References")
                    ?? throw new InvalidOperationException("VBA references collection is unavailable.");

                dynamic? targetReference = null;
                foreach (var reference in references)
                {
                    var name = SafeToString(TryGetDynamicProperty(reference, "Name"));
                    var guid = SafeToString(TryGetDynamicProperty(reference, "Guid"));
                    var fullPath = SafeToString(TryGetDynamicProperty(reference, "FullPath"));
                    if (string.Equals(name, referenceIdentifier, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(guid, referenceIdentifier, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(fullPath, referenceIdentifier, StringComparison.OrdinalIgnoreCase))
                    {
                        targetReference = reference;
                        break;
                    }
                }

                if (targetReference == null)
                    throw new InvalidOperationException($"VBA reference not found: {referenceIdentifier}");

                _ = InvokeDynamicMethod(references, "Remove", targetReference);
            },
            requireExclusive: false,
            releaseOleDb: false);
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
                var schema = GetSchema("Columns", new string[] { null!, null!, tableName, null! });
                
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
                using var command = CreateCommand($"SELECT COUNT(*) FROM [{tableName}]");
                return Convert.ToInt64(command.ExecuteScalar());
            }
            catch
            {
                return 0;
            }
        }

        private void ExecuteSchemaNonQuery(string sql)
        {
            EnsureNoActiveTransaction("Schema mutation");

            Exception? lastRecoverableError = null;

            for (var attempt = 0; attempt < 2; attempt++)
            {
                EnsureOleDbConnection();

                try
                {
                    using var command = CreateCommand(sql);
                    command.ExecuteNonQuery();
                    RefreshOleDbConnectionAfterSchemaMutation();
                    return;
                }
                catch (Exception ex) when (attempt == 0 && IsRecoverableOleDbLockError(ex) && TryReleaseExclusiveAccessLock())
                {
                    lastRecoverableError = ex;
                    CloseSqlConnections();
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
                OpenPreferredConnection(_currentDatabasePath);
            }
            catch
            {
                // Defer refresh to the next operation when immediate reopen is unavailable.
                CloseSqlConnections();
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

            var schema = GetSchema("Tables");
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

        private void EnsureUsysRibbonsTable()
        {
            if (TableExists("USysRibbons"))
                return;

            ExecuteSchemaNonQuery(
                "CREATE TABLE [USysRibbons] (" +
                "[ID] COUNTER CONSTRAINT [PrimaryKey] PRIMARY KEY, " +
                "[RibbonName] TEXT(255), " +
                "[RibbonXML] LONGTEXT)");
        }

        private void ApplyDatabasePasswordInternal(string? newPassword)
        {
            if (!IsConnected || string.IsNullOrWhiteSpace(_currentDatabasePath))
                throw new InvalidOperationException("Not connected to database");

            EnsureNoActiveTransaction("Database password update");

            var sourcePath = _currentDatabasePath!;
            var previousPassword = _databasePassword;
            var systemDatabasePath = _systemDatabasePath;
            var normalizedPassword = string.IsNullOrWhiteSpace(newPassword) ? null : newPassword.Trim();
            var tempPath = BuildCompactTemporaryPath(sourcePath);

            Disconnect();
            try
            {
                ExecuteWithTemporaryAccessApplication(accessApp =>
                {
                    var options = string.IsNullOrWhiteSpace(normalizedPassword) ? null : $";PWD={normalizedPassword}";
                    Exception? lastError = null;
                    var compacted = false;

                    try
                    {
                        var primaryResult = string.IsNullOrWhiteSpace(options)
                            ? InvokeDynamicMethod(accessApp, "CompactRepair", sourcePath, tempPath, true)
                            : InvokeDynamicMethod(accessApp, "CompactRepair", sourcePath, tempPath, true, Type.Missing, options);

                        compacted = primaryResult is bool compactedBool ? compactedBool : File.Exists(tempPath);
                    }
                    catch (Exception ex)
                    {
                        lastError = ex;
                    }

                    if (!compacted)
                    {
                        try
                        {
                            var fallbackResult = string.IsNullOrWhiteSpace(options)
                                ? InvokeDynamicMethod(accessApp, "CompactRepair", sourcePath, tempPath)
                                : InvokeDynamicMethod(accessApp, "CompactRepair", sourcePath, tempPath, true, options);

                            compacted = fallbackResult is bool fallbackBool ? fallbackBool : File.Exists(tempPath);
                            lastError = null;
                        }
                        catch (Exception fallbackEx)
                        {
                            lastError = fallbackEx;
                        }
                    }

                    if (!compacted && !File.Exists(tempPath))
                    {
                        throw new InvalidOperationException("Failed to compact database while applying password/encryption.", lastError);
                    }
                });

                ReplaceFileInPlace(tempPath, sourcePath);
                Connect(sourcePath, normalizedPassword, systemDatabasePath);
            }
            catch
            {
                if (!IsConnected)
                {
                    try
                    {
                        Connect(sourcePath, previousPassword, systemDatabasePath);
                    }
                    catch
                    {
                        // Surface original operation failure when reconnect also fails.
                    }
                }

                throw;
            }
            finally
            {
                TryDeleteFile(tempPath);
            }
        }

        private static dynamic? GetNavigationGroupsCollection(dynamic accessApp)
        {
            var currentProject = TryGetDynamicProperty(accessApp, "CurrentProject");
            return TryGetDynamicProperty(currentProject, "NavigationGroups")
                ?? TryGetDynamicProperty(accessApp, "NavigationGroups");
        }

        private static dynamic? FindNavigationGroup(dynamic groupsCollection, string groupName)
        {
            foreach (var group in groupsCollection)
            {
                var currentName = SafeToString(TryGetDynamicProperty(group, "Name"));
                if (string.Equals(currentName, groupName, StringComparison.OrdinalIgnoreCase))
                    return group;
            }

            return null;
        }

        private static dynamic? FindNavigationGroupObject(dynamic navigationButtonsCollection, string objectName)
        {
            foreach (var button in navigationButtonsCollection)
            {
                var currentName = SafeToString(TryGetDynamicProperty(button, "Name"));
                var currentObjectName = SafeToString(TryGetDynamicProperty(button, "ObjectName"))
                    ?? SafeToString(TryGetDynamicProperty(button, "Object"));
                if (string.Equals(currentName, objectName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(currentObjectName, objectName, StringComparison.OrdinalIgnoreCase))
                {
                    return button;
                }
            }

            return null;
        }

        private static bool TryDeleteFormatConditionByIndex(dynamic formatConditions, int ruleIndex)
        {
            if (ruleIndex <= 0)
                return false;

            try
            {
                _ = InvokeDynamicMethod(formatConditions, "Delete", ruleIndex);
                return true;
            }
            catch
            {
                // Try zero-based index fallback.
            }

            try
            {
                _ = InvokeDynamicMethod(formatConditions, "Delete", ruleIndex - 1);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static dynamic? GetFormatConditionByIndex(dynamic formatConditions, int ruleIndex)
        {
            if (ruleIndex <= 0)
                return null;

            try
            {
                var item = InvokeDynamicMethod(formatConditions, "Item", ruleIndex);
                if (item != null)
                    return item;
            }
            catch
            {
                // Try zero-based index fallback.
            }

            try
            {
                var item = InvokeDynamicMethod(formatConditions, "Item", ruleIndex - 1);
                if (item != null)
                    return item;
            }
            catch
            {
                // Fall back to enumeration.
            }

            var index = 0;
            foreach (var condition in formatConditions)
            {
                index++;
                if (index == ruleIndex)
                    return condition;
            }

            return null;
        }

        private static ConditionalFormattingRuleInfo BuildConditionalFormattingRuleInfo(dynamic condition, int index)
        {
            return new ConditionalFormattingRuleInfo
            {
                Index = index,
                TypeCode = ToInt32(TryGetDynamicProperty(condition, "Type")),
                Expression1 = SafeToString(TryGetDynamicProperty(condition, "Expression1")),
                Expression2 = SafeToString(TryGetDynamicProperty(condition, "Expression2")),
                ForeColor = ToNullableInt(TryGetDynamicProperty(condition, "ForeColor")),
                BackColor = ToNullableInt(TryGetDynamicProperty(condition, "BackColor")),
                Enabled = ToBool(TryGetDynamicProperty(condition, "Enabled"), true)
            };
        }

        private static string BuildAttachmentQuery(string tableName, string? whereCondition)
        {
            var escapedTable = EscapeSqlIdentifier(tableName);
            if (string.IsNullOrWhiteSpace(whereCondition))
                return $"SELECT * FROM [{escapedTable}]";

            return $"SELECT * FROM [{escapedTable}] WHERE {whereCondition}";
        }

        private static dynamic? TryGetAttachmentRecordset(dynamic recordset, string fieldName)
        {
            var field = GetRecordsetField(recordset, fieldName);
            if (field == null)
                return null;

            return TryGetDynamicProperty(field, "Value");
        }

        private static dynamic? GetRecordsetField(dynamic recordset, string fieldName)
        {
            var fields = TryGetDynamicProperty(recordset, "Fields");
            if (fields == null)
                return null;

            try
            {
                var byName = InvokeDynamicMethod(fields, "Item", fieldName);
                if (byName != null)
                    return byName;
            }
            catch
            {
                // Fall back to enumeration.
            }

            foreach (var field in fields)
            {
                var currentName = SafeToString(TryGetDynamicProperty(field, "Name"));
                if (string.Equals(currentName, fieldName, StringComparison.OrdinalIgnoreCase))
                    return field;
            }

            return null;
        }

        private static object? GetRecordsetFieldValue(dynamic recordset, string fieldName)
        {
            var field = GetRecordsetField(recordset, fieldName);
            return field == null ? null : TryGetDynamicProperty(field, "Value");
        }

        private static byte[]? ConvertAttachmentFieldToBytes(object? value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            if (value is byte[] bytes)
                return bytes;

            if (value is Array array)
            {
                try
                {
                    var result = new byte[array.Length];
                    for (var i = 0; i < array.Length; i++)
                    {
                        var element = array.GetValue(i);
                        result[i] = Convert.ToByte(element ?? 0);
                    }

                    return result;
                }
                catch
                {
                    return null;
                }
            }

            if (value is string text && !string.IsNullOrWhiteSpace(text))
            {
                try
                {
                    return Convert.FromBase64String(text);
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }

        private static void SetRecordsetFieldValue(dynamic recordset, string fieldName, object? value)
        {
            var field = GetRecordsetField(recordset, fieldName)
                ?? throw new InvalidOperationException($"Recordset field not found: {fieldName}");
            SetDynamicProperty(field, "Value", value);
        }

        private static List<Dictionary<string, object?>> ReadComplexFieldEntries(dynamic complexRecordset)
        {
            var results = new List<Dictionary<string, object?>>();
            while (!ToBool(TryGetDynamicProperty(complexRecordset, "EOF"), true))
            {
                var entry = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                var fields = TryGetDynamicProperty(complexRecordset, "Fields");
                if (fields != null)
                {
                    foreach (var field in fields)
                    {
                        var name = SafeToString(TryGetDynamicProperty(field, "Name"));
                        if (string.IsNullOrWhiteSpace(name))
                            continue;
                        entry[name] = NormalizeValue(TryGetDynamicProperty(field, "Value"));
                    }
                }

                results.Add(entry);
                _ = InvokeDynamicMethod(complexRecordset, "MoveNext");
            }

            return results;
        }

        private static object? GetPreferredComplexFieldValue(Dictionary<string, object?> entry)
        {
            if (entry.TryGetValue("Value", out var preferred))
                return preferred;

            foreach (var pair in entry)
            {
                if (!string.Equals(pair.Key, "ID", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(pair.Key, "GUID", StringComparison.OrdinalIgnoreCase))
                {
                    return pair.Value;
                }
            }

            return entry.Values.FirstOrDefault();
        }

        private static void SetComplexFieldEntryValue(dynamic complexRecordset, object? value)
        {
            var field = GetRecordsetField(complexRecordset, "Value");
            if (field == null)
            {
                var fields = TryGetDynamicProperty(complexRecordset, "Fields")
                    ?? throw new InvalidOperationException("Complex field entry fields collection is unavailable.");
                foreach (var currentField in fields)
                {
                    var name = SafeToString(TryGetDynamicProperty(currentField, "Name"));
                    if (string.Equals(name, "ID", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(name, "GUID", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    field = currentField;
                    break;
                }
            }

            if (field == null)
                throw new InvalidOperationException("Unable to identify writable value field for complex data entry.");

            SetDynamicProperty(field, "Value", value ?? DBNull.Value);
        }

        private bool FieldExists(string tableName, string fieldName)
        {
            EnsureOleDbConnection();

            var schema = GetSchema("Columns", new string[] { null!, null!, tableName, null! });
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

            if (string.IsNullOrWhiteSpace(_databasePassword))
            {
                accessApp.OpenCurrentDatabase(_currentDatabasePath, false);
            }
            else
            {
                accessApp.OpenCurrentDatabase(_currentDatabasePath, false, _databasePassword);
            }
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

        private dynamic ResolveField(dynamic accessApp, string tableName, string fieldName)
        {
            var tableDef = FindTableDefWithRetry(accessApp, tableName)
                ?? throw new InvalidOperationException($"Table not found: {tableName}");
            var field = FindTableField(tableDef, fieldName)
                ?? throw new InvalidOperationException($"Field not found: {tableName}.{fieldName}");
            return field;
        }

        private static dynamic ResolveDaoDocument(dynamic accessApp, string containerName, string documentName)
        {
            var currentDb = TryGetCurrentDb(accessApp)
                ?? throw new InvalidOperationException("DAO CurrentDb is unavailable.");
            var containers = TryGetDynamicProperty(currentDb, "Containers")
                ?? throw new InvalidOperationException("DAO Containers collection is unavailable.");

            dynamic? targetContainer = null;
            foreach (var container in containers)
            {
                var currentName = SafeToString(TryGetDynamicProperty(container, "Name"));
                if (string.Equals(currentName, containerName, StringComparison.OrdinalIgnoreCase))
                {
                    targetContainer = container;
                    break;
                }
            }

            if (targetContainer == null)
                throw new InvalidOperationException($"Container not found: {containerName}");

            var documents = TryGetDynamicProperty(targetContainer, "Documents")
                ?? throw new InvalidOperationException($"Documents collection is unavailable for container '{containerName}'.");
            foreach (var document in documents)
            {
                var currentName = SafeToString(TryGetDynamicProperty(document, "Name"));
                if (string.Equals(currentName, documentName, StringComparison.OrdinalIgnoreCase))
                    return document;
            }

            throw new InvalidOperationException($"Document not found: {containerName}.{documentName}");
        }

        private static dynamic? FindDaoProperty(dynamic owner, string propertyName)
        {
            var properties = TryGetDynamicProperty(owner, "Properties");
            if (properties == null)
                return null;

            try
            {
                var byItem = InvokeDynamicMethod(properties, "Item", propertyName);
                if (byItem != null)
                    return byItem;
            }
            catch
            {
                // Fall back to manual enumeration.
            }

            foreach (var property in properties)
            {
                var currentName = SafeToString(TryGetDynamicProperty(property, "Name"));
                if (string.Equals(currentName, propertyName, StringComparison.OrdinalIgnoreCase))
                    return property;
            }

            return null;
        }

        private static object? GetDaoPropertyValue(dynamic owner, string propertyName)
        {
            var property = FindDaoProperty(owner, propertyName);
            return property == null ? null : TryGetDynamicProperty(property, "Value");
        }

        private static void SetDaoPropertyValue(dynamic owner, string propertyName, object? value, int? daoType, bool createIfMissing)
        {
            var property = FindDaoProperty(owner, propertyName);
            if (property != null)
            {
                SetDynamicProperty(property, "Value", value);
                return;
            }

            if (!createIfMissing)
                throw new InvalidOperationException($"Property not found: {propertyName}");

            var properties = TryGetDynamicProperty(owner, "Properties")
                ?? throw new InvalidOperationException("DAO properties collection is unavailable.");

            var effectiveType = daoType ?? (value is string s && s.Length > 255 ? 12 : 10);
            var createdProperty = InvokeDynamicMethod(owner, "CreateProperty", propertyName, effectiveType, value)
                ?? throw new InvalidOperationException($"Failed to create property: {propertyName}");
            _ = InvokeDynamicMethod(properties, "Append", createdProperty);
        }

        private static int? ParseDaoDataType(string? propertyType)
        {
            if (string.IsNullOrWhiteSpace(propertyType))
                return null;

            var normalized = NormalizeEnumToken(propertyType);
            return normalized switch
            {
                "boolean" or "bool" => 1,
                "byte" => 2,
                "short" or "smallint" or "integer" => 3,
                "long" or "int" => 4,
                "currency" => 5,
                "single" => 6,
                "double" or "float" => 7,
                "date" or "datetime" or "time" => 8,
                "binary" => 9,
                "text" or "string" => 10,
                "longbinary" or "blob" => 11,
                "memo" or "longtext" or "note" => 12,
                "guid" => 15,
                _ => throw new ArgumentException($"Unsupported property_type: {propertyType}", nameof(propertyType))
            };
        }

        private static object ConvertPropertyValue(string value, string? propertyType, object? existingValue)
        {
            if (string.IsNullOrWhiteSpace(propertyType))
                return ConvertValueForProperty(value, existingValue) ?? value;

            var normalized = NormalizeEnumToken(propertyType);
            return normalized switch
            {
                "boolean" or "bool" => ToBool(value, false),
                "byte" => byte.TryParse(value, out var byteValue) ? byteValue : throw new ArgumentException("Value must be a byte."),
                "short" or "smallint" or "integer" => short.TryParse(value, out var shortValue) ? shortValue : throw new ArgumentException("Value must be a short integer."),
                "long" or "int" => int.TryParse(value, out var intValue) ? intValue : throw new ArgumentException("Value must be an integer."),
                "currency" or "single" or "double" or "float" => double.TryParse(value, out var doubleValue) ? doubleValue : throw new ArgumentException("Value must be numeric."),
                "date" or "datetime" or "time" => DateTime.TryParse(value, out var dateValue) ? dateValue : throw new ArgumentException("Value must be a date/time."),
                _ => value
            };
        }

        private static bool IsLikelySystemDatabaseProperty(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
                return true;

            if (string.Equals(propertyName, "Title", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(propertyName, "Author", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(propertyName, "Subject", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(propertyName, "Keywords", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(propertyName, "Comments", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (propertyName.StartsWith("{", StringComparison.Ordinal) ||
                propertyName.StartsWith("Jet ", StringComparison.OrdinalIgnoreCase) ||
                propertyName.StartsWith("Access", StringComparison.OrdinalIgnoreCase) ||
                propertyName.StartsWith("NameMap", StringComparison.OrdinalIgnoreCase) ||
                propertyName.StartsWith("db", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static int? ToNullableInt(object? value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return null;
            }
        }

        private static bool? ToNullableBool(object? value)
        {
            if (value == null || value == DBNull.Value)
                return null;

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
                return null;
            }
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
                var schema = GetSchema("ForeignKeys");
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
                var schema = GetSchema("ForeignKeys");
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
                var schema = GetSchema("Columns", new string[] { null!, null!, tableName, null! });
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

        private IDbCommand CreateCommand(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
                throw new ArgumentException("SQL is required.", nameof(sql));

            EnsureOleDbConnection();
            PruneInvalidTransactionState();

            IDbCommand command;
            if (_activeDataProvider == DataProviderKind.Odbc)
            {
                command = new OdbcCommand(sql, _odbcConnection);
            }
            else
            {
                command = new OleDbCommand(sql, _oleDbConnection);
            }

            if (_oleDbTransaction != null)
            {
                var transactionConnection = _oleDbTransaction.Connection;
                if (transactionConnection == null || !ReferenceEquals(transactionConnection, _oleDbConnection))
                {
                    ResetTransactionState(attemptRollback: false);
                    throw new InvalidOperationException("Active transaction is no longer valid because the database connection changed.");
                }

                command.Transaction = _oleDbTransaction;
            }
            else if (_odbcTransaction != null)
            {
                var transactionConnection = _odbcTransaction.Connection;
                if (transactionConnection == null || !ReferenceEquals(transactionConnection, _odbcConnection))
                {
                    ResetTransactionState(attemptRollback: false);
                    throw new InvalidOperationException("Active transaction is no longer valid because the database connection changed.");
                }

                command.Transaction = _odbcTransaction;
            }

            return command;
        }

        private static void AddCommandParameter(IDbCommand command, string parameterName, object? value)
        {
            var parameter = command.CreateParameter();
            parameter.ParameterName = parameterName;
            parameter.Value = value ?? DBNull.Value;
            command.Parameters.Add(parameter);
        }

        private void EnsureNoActiveTransaction(string operationName)
        {
            PruneInvalidTransactionState();
            if (HasActiveTransaction())
                throw new InvalidOperationException($"{operationName} is not allowed while a transaction is active. Commit or rollback first.");
        }

        private T ExecuteWithConnectedDatabaseReleased<T>(string sourceDatabasePath, string operationName, Func<T> operation)
        {
            var shouldDisconnectCurrent = IsConnected &&
                !string.IsNullOrWhiteSpace(_currentDatabasePath) &&
                PathsMatch(_currentDatabasePath, sourceDatabasePath);

            if (!shouldDisconnectCurrent)
                return operation();

            EnsureNoActiveTransaction(operationName);

            var reconnectPath = _currentDatabasePath!;
            var reconnectPassword = _databasePassword;
            var reconnectSystemDatabasePath = _systemDatabasePath;
            Disconnect();

            Exception? operationError = null;
            try
            {
                return operation();
            }
            catch (Exception ex)
            {
                operationError = ex;
                throw;
            }
            finally
            {
                try
                {
                    Connect(reconnectPath, reconnectPassword, reconnectSystemDatabasePath);
                }
                catch (Exception reconnectEx)
                {
                    if (operationError != null)
                    {
                        throw new AggregateException(
                            $"{operationName} failed and reconnecting to {reconnectPath} also failed.",
                            operationError,
                            reconnectEx);
                    }

                    throw;
                }
            }
        }

        private static void ExecuteWithTemporaryAccessApplication(Action<dynamic> action)
        {
            var accessType = Type.GetTypeFromProgID("Access.Application", throwOnError: false);
            if (accessType == null)
                throw new InvalidOperationException("Microsoft Access COM automation is not available on this machine.");

            dynamic? accessApp = null;
            try
            {
                accessApp = Activator.CreateInstance(accessType);
                if (accessApp == null)
                    throw new InvalidOperationException("Failed to create Access.Application COM instance.");

                try
                {
                    accessApp.Visible = false;
                }
                catch
                {
                    // Best-effort: keep temporary automation instances headless.
                }

                try
                {
                    accessApp.UserControl = false;
                }
                catch
                {
                    // Best-effort: keep temporary automation instances non-interactive.
                }

                action(accessApp);
            }
            finally
            {
                if (accessApp != null)
                {
                    try
                    {
                        accessApp.Quit(2);
                    }
                    catch
                    {
                        // Ignore shutdown failures while releasing COM resources.
                    }

                    try
                    {
                        if (Marshal.IsComObject(accessApp))
                            Marshal.FinalReleaseComObject(accessApp);
                    }
                    catch
                    {
                        // Ignore RCW cleanup failures.
                    }
                }
            }
        }

        private static void RunCompactRepair(string sourceDatabasePath, string destinationDatabasePath)
        {
            ExecuteWithTemporaryAccessApplication(accessApp =>
            {
                var result = accessApp.CompactRepair(sourceDatabasePath, destinationDatabasePath, true);
                if (result is bool compacted && !compacted)
                    throw new InvalidOperationException($"Compact/repair operation returned false for destination: {destinationDatabasePath}");
            });

            if (!File.Exists(destinationDatabasePath))
                throw new InvalidOperationException($"Compact/repair did not produce destination database: {destinationDatabasePath}");
        }

        private static void ReplaceFileInPlace(string compactedDatabasePath, string sourceDatabasePath)
        {
            var sourceDirectory = Path.GetDirectoryName(sourceDatabasePath);
            var backupFileName = $"{Path.GetFileName(sourceDatabasePath)}.precompact.{Guid.NewGuid():N}.bak";
            var backupPath = string.IsNullOrWhiteSpace(sourceDirectory)
                ? Path.Combine(Path.GetTempPath(), backupFileName)
                : Path.Combine(sourceDirectory, backupFileName);

            try
            {
                File.Replace(compactedDatabasePath, sourceDatabasePath, backupPath, ignoreMetadataErrors: true);
            }
            finally
            {
                if (File.Exists(backupPath))
                {
                    try
                    {
                        File.Delete(backupPath);
                    }
                    catch
                    {
                        // Ignore cleanup failures for temporary backup files.
                    }
                }

                if (File.Exists(compactedDatabasePath))
                {
                    try
                    {
                        File.Delete(compactedDatabasePath);
                    }
                    catch
                    {
                        // Ignore cleanup failures for temporary compacted files.
                    }
                }
            }
        }

        private static string BuildCompactTemporaryPath(string sourceDatabasePath)
        {
            var sourceDirectory = Path.GetDirectoryName(sourceDatabasePath);
            var temporaryFileName = $"{Path.GetFileNameWithoutExtension(sourceDatabasePath)}.compact.{Guid.NewGuid():N}{Path.GetExtension(sourceDatabasePath)}";
            if (string.IsNullOrWhiteSpace(sourceDirectory))
                return Path.Combine(Path.GetTempPath(), temporaryFileName);

            return Path.Combine(sourceDirectory, temporaryFileName);
        }

        private static string NormalizeDatabasePath(string databasePath, string paramName, bool requireExists)
        {
            if (string.IsNullOrWhiteSpace(databasePath))
                throw new ArgumentException("Database path is required.", paramName);

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(databasePath.Trim());
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Database path is invalid.", paramName, ex);
            }

            var extension = Path.GetExtension(fullPath);
            if (!SupportedDatabaseExtensions.Contains(extension))
                throw new ArgumentException($"Database path must use a .accdb or .mdb extension: {fullPath}", paramName);

            if (requireExists && !File.Exists(fullPath))
                throw new FileNotFoundException($"Database file not found: {fullPath}");

            return fullPath;
        }

        private static string? NormalizeSystemDatabasePath(string? systemDatabasePath)
        {
            if (string.IsNullOrWhiteSpace(systemDatabasePath))
                return null;

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(systemDatabasePath.Trim());
            }
            catch (Exception ex)
            {
                throw new ArgumentException("System database path is invalid.", nameof(systemDatabasePath), ex);
            }

            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"System database file not found: {fullPath}");

            return fullPath;
        }

        private static void EnsureDistinctDatabasePaths(string sourcePath, string destinationPath, string sourceParamName, string destinationParamName)
        {
            if (PathsMatch(sourcePath, destinationPath))
                throw new ArgumentException($"{sourceParamName} and {destinationParamName} must refer to different files.", destinationParamName);
        }

        private static string BuildOdbcSecuritySegment(string? databasePassword, string? systemDatabasePath)
        {
            var segment = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(databasePassword))
                segment.Append("PWD=").Append(EscapeOdbcValue(databasePassword)).Append(';');

            if (!string.IsNullOrWhiteSpace(systemDatabasePath))
                segment.Append("SystemDB=").Append(EscapeOdbcValue(systemDatabasePath)).Append(';');

            return segment.ToString();
        }

        private static string EscapeOdbcValue(string value)
        {
            if (value.IndexOfAny(new[] { ';', '{', '}', ' ' }) >= 0)
                return "{" + value.Replace("}", "}}") + "}";

            return value;
        }

        private static string NormalizeLinkSourceDatabasePath(string sourceDatabasePath, string paramName)
        {
            if (string.IsNullOrWhiteSpace(sourceDatabasePath))
                throw new ArgumentException("Source database path is required.", paramName);

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(sourceDatabasePath.Trim());
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Source database path is invalid.", paramName, ex);
            }

            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"Source database file not found: {fullPath}");

            return fullPath;
        }

        private static string NormalizeLinkConnectString(string? connectString, string normalizedSourceDatabasePath)
        {
            if (!string.IsNullOrWhiteSpace(connectString))
            {
                var normalized = connectString.Trim();

                // Accept common caller formats like:
                //   DATABASE=C:\db.accdb
                //   ;DATABASE=C:\db.accdb
                //   MS Access;DATABASE=C:\db.accdb
                if (normalized.StartsWith("MS Access;", StringComparison.OrdinalIgnoreCase))
                {
                    normalized = normalized.Substring("MS Access".Length);
                }

                if (!normalized.StartsWith(";", StringComparison.Ordinal))
                {
                    normalized = ";" + normalized;
                }

                if (!normalized.EndsWith(";", StringComparison.Ordinal))
                {
                    normalized += ";";
                }

                return normalized;
            }

            return BuildAccessLinkConnectString(normalizedSourceDatabasePath);
        }

        private static string BuildAccessLinkConnectString(string normalizedSourceDatabasePath)
        {
            // DAO link strings expect a trailing semicolon terminator.
            return $";DATABASE={normalizedSourceDatabasePath};";
        }

        private static string? ExtractDatabasePathFromConnectString(string? connectString)
        {
            if (string.IsNullOrWhiteSpace(connectString))
                return null;

            var match = Regex.Match(connectString, @"(?:^|;)\s*(DATABASE|DBQ|Data Source)\s*=\s*(?<value>[^;]+)", RegexOptions.IgnoreCase);
            if (!match.Success)
                return null;

            var value = match.Groups["value"].Value.Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(value))
                return null;

            try
            {
                return Path.GetFullPath(value);
            }
            catch
            {
                return value;
            }
        }

        private static bool IsLinkedTableDef(dynamic tableDef)
        {
            var connectString = SafeToString(TryGetDynamicProperty(tableDef, "Connect"));
            return !string.IsNullOrWhiteSpace(connectString);
        }

        private static bool IsSystemOrTemporaryTableName(string tableName)
        {
            return tableName.StartsWith("~", StringComparison.Ordinal) ||
                   tableName.StartsWith("MSys", StringComparison.OrdinalIgnoreCase);
        }

        private static IsolationLevel ParseIsolationLevel(string? isolationLevel)
        {
            if (string.IsNullOrWhiteSpace(isolationLevel))
                return IsolationLevel.ReadCommitted;

            var normalized = isolationLevel
                .Trim()
                .Replace("_", string.Empty, StringComparison.Ordinal)
                .Replace("-", string.Empty, StringComparison.Ordinal)
                .ToLowerInvariant();

            return normalized switch
            {
                "chaos" => IsolationLevel.Chaos,
                "readuncommitted" => IsolationLevel.ReadUncommitted,
                "readcommitted" => IsolationLevel.ReadCommitted,
                "repeatableread" => IsolationLevel.RepeatableRead,
                "serializable" => IsolationLevel.Serializable,
                "unspecified" => IsolationLevel.Unspecified,
                _ => throw new ArgumentException("Unsupported isolation level. Valid values: read_committed, read_uncommitted, repeatable_read, serializable, chaos, unspecified.", nameof(isolationLevel))
            };
        }

        private void PruneInvalidTransactionState()
        {
            if (_oleDbTransaction != null &&
                (_oleDbConnection?.State != ConnectionState.Open || _oleDbTransaction.Connection == null))
            {
                ResetTransactionState(attemptRollback: false);
            }

            if (_odbcTransaction != null &&
                (_odbcConnection?.State != ConnectionState.Open || _odbcTransaction.Connection == null))
            {
                ResetTransactionState(attemptRollback: false);
            }
        }

        private bool HasActiveTransaction()
        {
            return _oleDbTransaction != null || _odbcTransaction != null;
        }

        private DbTransaction? GetActiveTransaction()
        {
            if (_oleDbTransaction != null)
                return _oleDbTransaction;

            return _odbcTransaction;
        }

        private void ResetTransactionState(bool attemptRollback)
        {
            if (_oleDbTransaction != null)
            {
                try
                {
                    if (attemptRollback)
                        _oleDbTransaction.Rollback();
                }
                catch
                {
                    // Ignore rollback failures during cleanup.
                }

                try
                {
                    _oleDbTransaction.Dispose();
                }
                catch
                {
                    // Ignore disposal failures during cleanup.
                }

                _oleDbTransaction = null;
            }

            if (_odbcTransaction != null)
            {
                try
                {
                    if (attemptRollback)
                        _odbcTransaction.Rollback();
                }
                catch
                {
                    // Ignore rollback failures during cleanup.
                }

                try
                {
                    _odbcTransaction.Dispose();
                }
                catch
                {
                    // Ignore disposal failures during cleanup.
                }

                _odbcTransaction = null;
            }

            _transactionStartedAtUtc = null;
        }

        private void OpenOleDbConnection(string databasePath)
        {
            ResetTransactionState(attemptRollback: true);
            CloseSqlConnections();

            var connectionStringBuilder = new OleDbConnectionStringBuilder
            {
                Provider = "Microsoft.ACE.OLEDB.12.0",
                DataSource = databasePath
            };

            if (!string.IsNullOrWhiteSpace(_databasePassword))
                connectionStringBuilder["Jet OLEDB:Database Password"] = _databasePassword;

            if (!string.IsNullOrWhiteSpace(_systemDatabasePath))
                connectionStringBuilder["Jet OLEDB:System Database"] = _systemDatabasePath;

            _oleDbConnection = new OleDbConnection(connectionStringBuilder.ConnectionString);
            _oleDbConnection.Open();
            _activeDataProvider = DataProviderKind.OleDb;
            _preferredDataProvider = DataProviderKind.OleDb;
        }

        private void OpenOdbcConnection(string databasePath)
        {
            ResetTransactionState(attemptRollback: true);
            CloseSqlConnections();

            Exception? lastError = null;
            foreach (var connectionString in BuildOdbcConnectionStrings(databasePath, _databasePassword, _systemDatabasePath))
            {
                try
                {
                    var connection = new OdbcConnection(connectionString);
                    connection.Open();
                    _odbcConnection = connection;
                    _activeDataProvider = DataProviderKind.Odbc;
                    _preferredDataProvider = DataProviderKind.Odbc;
                    return;
                }
                catch (Exception ex)
                {
                    lastError = ex;
                }
            }

            throw new InvalidOperationException(
                $"Failed to open Access database via ODBC. Install a Microsoft Access ODBC driver (for example: Microsoft Access Driver (*.mdb, *.accdb)). Last ODBC error: {lastError?.Message}",
                lastError);
        }

        private void OpenPreferredConnection(string databasePath)
        {
            var preferred = _preferredDataProvider == DataProviderKind.None ? DataProviderKind.OleDb : _preferredDataProvider;

            if (preferred == DataProviderKind.Odbc)
            {
                Exception? odbcError = null;
                try
                {
                    OpenOdbcConnection(databasePath);
                    return;
                }
                catch (Exception ex)
                {
                    odbcError = ex;
                }

                try
                {
                    OpenOleDbConnection(databasePath);
                    return;
                }
                catch (Exception oleDbError)
                {
                    throw new AggregateException("Unable to open Access database using ODBC or OleDb providers.", odbcError!, oleDbError);
                }
            }

            Exception? primaryError = null;
            try
            {
                OpenOleDbConnection(databasePath);
                return;
            }
            catch (Exception ex) when (IsRecoverableOleDbLockError(ex) && TryReleaseExclusiveAccessLock())
            {
                OpenOleDbConnection(databasePath);
                return;
            }
            catch (Exception ex) when (ShouldUseOdbcFallbackForOleDbOpen(ex))
            {
                primaryError = ex;
            }

            if (primaryError == null)
                throw new InvalidOperationException("Failed to open Access database.");

            try
            {
                OpenOdbcConnection(databasePath);
            }
            catch (Exception odbcError)
            {
                throw new AggregateException("Unable to open Access database using OleDb or ODBC providers.", primaryError, odbcError);
            }
        }

        private static IEnumerable<string> BuildOdbcConnectionStrings(string databasePath, string? databasePassword, string? systemDatabasePath)
        {
            var securitySegment = BuildOdbcSecuritySegment(databasePassword, systemDatabasePath);
            yield return $"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={databasePath};{securitySegment}";
            yield return $"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={databasePath};ExtendedAnsiSQL=1;{securitySegment}";
            yield return $"Driver={{Microsoft Access Driver (*.mdb)}};Dbq={databasePath};{securitySegment}";
        }

        private void EnsureOleDbConnection()
        {
            PruneInvalidTransactionState();
            if (_oleDbConnection?.State == ConnectionState.Open)
            {
                _activeDataProvider = DataProviderKind.OleDb;
                return;
            }

            if (_odbcConnection?.State == ConnectionState.Open)
            {
                _activeDataProvider = DataProviderKind.Odbc;
                return;
            }

            if (string.IsNullOrWhiteSpace(_currentDatabasePath))
                throw new InvalidOperationException("Not connected to database");

            OpenPreferredConnection(_currentDatabasePath);
        }

        private void CloseSqlConnections()
        {
            _oleDbConnection?.Close();
            _oleDbConnection?.Dispose();
            _oleDbConnection = null;

            _odbcConnection?.Close();
            _odbcConnection?.Dispose();
            _odbcConnection = null;

            _activeDataProvider = DataProviderKind.None;
        }

        private void ExecuteWithOleDbReleased(Action action)
        {
            EnsureNoActiveTransaction("Temporarily releasing OleDb connection");

            var isOuterScope = _oleDbReleaseDepth == 0;
            if (isOuterScope)
            {
                _restoreOleDbAfterRelease = IsConnected && !string.IsNullOrWhiteSpace(_currentDatabasePath);
                if (_restoreOleDbAfterRelease)
                {
                    _providerToRestoreAfterRelease = _activeDataProvider == DataProviderKind.None
                        ? _preferredDataProvider
                        : _activeDataProvider;
                    CloseSqlConnections();
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
                    if (_restoreOleDbAfterRelease && _oleDbConnection == null && _odbcConnection == null && !string.IsNullOrWhiteSpace(_currentDatabasePath))
                    {
                        try
                        {
                            _preferredDataProvider = _providerToRestoreAfterRelease == DataProviderKind.None
                                ? _preferredDataProvider
                                : _providerToRestoreAfterRelease;
                            OpenPreferredConnection(_currentDatabasePath);
                        }
                        catch
                        {
                            // Defer reconnection until next SQL operation.
                        }
                    }

                    _restoreOleDbAfterRelease = false;
                    _providerToRestoreAfterRelease = DataProviderKind.None;
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

                try
                {
                    _accessApplication.Visible = false;
                }
                catch
                {
                    // Best effort only.
                }

                try
                {
                    _accessApplication.UserControl = false;
                }
                catch
                {
                    // Best effort only.
                }
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
                if (string.IsNullOrWhiteSpace(_databasePassword))
                {
                    accessApplication.OpenCurrentDatabase(databasePath, requireExclusive);
                }
                else
                {
                    accessApplication.OpenCurrentDatabase(databasePath, requireExclusive, _databasePassword);
                }
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

        private DataTable GetSchema(string collectionName)
        {
            EnsureOleDbConnection();

            return _activeDataProvider switch
            {
                DataProviderKind.Odbc => _odbcConnection!.GetSchema(collectionName),
                _ => _oleDbConnection!.GetSchema(collectionName)
            };
        }

        private DataTable GetSchema(string collectionName, string[] restrictions)
        {
            EnsureOleDbConnection();

            return _activeDataProvider switch
            {
                DataProviderKind.Odbc => _odbcConnection!.GetSchema(collectionName, restrictions),
                _ => _oleDbConnection!.GetSchema(collectionName, restrictions)
            };
        }

        private string GetProviderDataTypeName(DataRow schemaRow, int? dataTypeCode)
        {
            var providerTypeName = GetRowString(schemaRow, "TYPE_NAME");
            if (!string.IsNullOrWhiteSpace(providerTypeName))
                return providerTypeName;

            if (!dataTypeCode.HasValue)
                return "Unknown";

            if (_activeDataProvider == DataProviderKind.Odbc)
                return dataTypeCode.Value.ToString();

            return ((OleDbType)dataTypeCode.Value).ToString();
        }

        private static bool ShouldUseOdbcFallbackForOleDbOpen(Exception ex)
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

            return ex.InnerException != null && ShouldUseOdbcFallbackForOleDbOpen(ex.InnerException);
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
                var schema = GetSchema("ForeignKeys");
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
                var schema = GetSchema("ForeignKeys");
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
                        using var command = CreateCommand(sql);
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

        private static List<ModuleProcedureInfo> ParseModuleProcedures(string code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return new List<ModuleProcedureInfo>();

            var normalized = NormalizeLineEndings(code).Replace("\r\n", "\n", StringComparison.Ordinal);
            var lines = normalized.Split('\n');
            var procedureStartRegex = new Regex(
                "^\\s*(?:Public|Private|Friend|Static)?\\s*(Sub|Function|Property\\s+Get|Property\\s+Let|Property\\s+Set)\\s+([A-Za-z_][A-Za-z0-9_]*)\\b",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

            var procedures = new List<ModuleProcedureInfo>();
            var index = 0;
            while (index < lines.Length)
            {
                var line = lines[index];
                var match = procedureStartRegex.Match(line);
                if (!match.Success)
                {
                    index++;
                    continue;
                }

                var declarationType = match.Groups[1].Value.Trim();
                var procedureName = match.Groups[2].Value.Trim();
                var startLine = index + 1;
                var endPattern = declarationType.StartsWith("Property", StringComparison.OrdinalIgnoreCase)
                    ? new Regex("^\\s*End\\s+Property\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled)
                    : declarationType.Equals("Sub", StringComparison.OrdinalIgnoreCase)
                        ? new Regex("^\\s*End\\s+Sub\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled)
                        : new Regex("^\\s*End\\s+Function\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

                var endLine = startLine;
                for (var scan = index + 1; scan < lines.Length; scan++)
                {
                    if (endPattern.IsMatch(lines[scan]))
                    {
                        endLine = scan + 1;
                        break;
                    }
                }

                var lineCount = Math.Max(1, endLine - startLine + 1);
                procedures.Add(new ModuleProcedureInfo
                {
                    Name = procedureName,
                    ProcedureType = declarationType,
                    StartLine = startLine,
                    LineCount = lineCount
                });

                index = Math.Max(index + 1, endLine);
            }

            return procedures;
        }

        private static dynamic? GetImportExportSpecificationsCollection(dynamic accessApp)
        {
            var currentProject = TryGetDynamicProperty(accessApp, "CurrentProject");
            var collection = TryGetDynamicProperty(currentProject, "ImportExportSpecifications");
            if (collection != null)
                return collection;

            return TryGetDynamicProperty(accessApp, "ImportExportSpecifications");
        }

        private static dynamic? FindImportExportSpecification(dynamic specifications, string specificationName)
        {
            foreach (var specification in specifications)
            {
                var name = SafeToString(TryGetDynamicProperty(specification, "Name"));
                if (string.Equals(name, specificationName, StringComparison.OrdinalIgnoreCase))
                    return specification;
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

        private static int ParseTransferType(string transferType, string paramName)
        {
            if (string.IsNullOrWhiteSpace(transferType))
                throw new ArgumentException("Transfer type is required.", paramName);

            var trimmed = transferType.Trim();
            if (int.TryParse(trimmed, out var numericTransferType))
                return numericTransferType;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "import" or "acimport" => 0,
                "export" or "acexport" => 1,
                "link" or "aclink" => 2,
                _ => throw new ArgumentException("transfer_type must be import, export, link, or an Access enum integer value.", paramName)
            };
        }

        private static int ParseSpreadsheetType(string? spreadsheetType)
        {
            if (string.IsNullOrWhiteSpace(spreadsheetType))
                return 10; // acSpreadsheetTypeExcel12Xml

            var trimmed = spreadsheetType.Trim();
            if (int.TryParse(trimmed, out var numericSpreadsheetType))
                return numericSpreadsheetType;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "excel3" or "acspreadsheettypeexcel3" => 0,
                "excel4" or "acspreadsheettypeexcel4" => 6,
                "excel5" or "excel7" or "acspreadsheettypeexcel5" or "acspreadsheettypeexcel7" => 5,
                "excel8" or "excel9" or "acspreadsheettypeexcel8" or "acspreadsheettypeexcel9" => 8,
                "excel12" or "acspreadsheettypeexcel12" => 9,
                "excel12xml" or "acspreadsheettypeexcel12xml" => 10,
                _ => throw new ArgumentException("spreadsheet_type must be a supported name (excel12xml, excel12, excel8, excel5, etc.) or an Access enum integer value.", nameof(spreadsheetType))
            };
        }

        private static int ParseOutputObjectType(string objectType, string paramName)
        {
            var trimmed = objectType.Trim();
            if (int.TryParse(trimmed, out var numericObjectType))
                return numericObjectType;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "table" or "acoutputtable" => 0,
                "query" or "acoutputquery" => 1,
                "form" or "acoutputform" => 2,
                "report" or "acoutputreport" => 3,
                "module" or "acoutputmodule" => 5,
                "dataaccesspage" or "acoutputdataaccesspage" => 6,
                "serverview" or "acoutputserverview" => 7,
                "storedprocedure" or "acoutputstoredprocedure" => 9,
                "function" or "acoutputfunction" => 10,
                _ => throw new ArgumentException("object_type must be table, query, form, report, module, data_access_page, server_view, stored_procedure, function, or an Access enum integer value.", paramName)
            };
        }

        private static object ParseOutputFormat(string outputFormat)
        {
            var trimmed = outputFormat.Trim();
            if (int.TryParse(trimmed, out var numericFormat))
                return numericFormat;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "pdf" or "acformatpdf" => "PDF Format (*.pdf)",
                "xps" or "acformatxps" => "XPS Format (*.xps)",
                "rtf" or "acformatrtf" => "Rich Text Format (*.rtf)",
                "txt" or "text" or "acformattxt" => "Text Files (*.txt)",
                "html" or "htm" or "acformathtml" => "HTML Files (*.html)",
                "xls" or "acformatxls" => "Microsoft Excel 97-2003 Workbook(*.xls)",
                "xlsx" or "acformatxlsx" => "Microsoft Excel Workbook(*.xlsx)",
                _ => trimmed
            };
        }

        private static int? ParseOutputQuality(string? outputQuality)
        {
            if (string.IsNullOrWhiteSpace(outputQuality))
                return null;

            var trimmed = outputQuality.Trim();
            if (int.TryParse(trimmed, out var numericQuality))
                return numericQuality;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "print" or "acexportqualityprint" => 0,
                "screen" or "acexportqualityscreen" => 1,
                _ => throw new ArgumentException("output_quality must be print, screen, or an Access enum integer value.", nameof(outputQuality))
            };
        }

        private static object ParseOpenTableView(string? view)
        {
            if (string.IsNullOrWhiteSpace(view))
                return Type.Missing;

            var trimmed = view.Trim();
            if (int.TryParse(trimmed, out var numericView))
                return numericView;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "datasheet" or "normal" or "acviewnormal" => 0,
                "design" or "acviewdesign" => 1,
                "printpreview" or "preview" or "acviewpreview" => 2,
                "pivottable" or "acviewpivottable" => 3,
                "pivotchart" or "acviewpivotchart" => 4,
                _ => throw new ArgumentException("view must be datasheet, design, print_preview, pivot_table, pivot_chart, or an Access enum integer value.", nameof(view))
            };
        }

        private static object ParseOpenDataMode(string? dataMode)
        {
            if (string.IsNullOrWhiteSpace(dataMode))
                return Type.Missing;

            var trimmed = dataMode.Trim();
            if (int.TryParse(trimmed, out var numericDataMode))
                return numericDataMode;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "add" or "acformadd" => 0,
                "edit" or "acformedit" => 1,
                "readonly" or "acformreadonly" => 2,
                _ => throw new ArgumentException("data_mode must be add, edit, read_only, or an Access enum integer value.", nameof(dataMode))
            };
        }

        private static int ParseDoCmdObjectType(string objectType, string paramName)
        {
            var trimmed = objectType.Trim();
            if (int.TryParse(trimmed, out var numericObjectType))
                return numericObjectType;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "table" or "actable" => 0,
                "query" or "acquery" => 1,
                "form" or "acform" => 2,
                "report" or "acreport" => 3,
                "macro" or "acmacro" => 4,
                "module" or "acmodule" => 5,
                "serverview" or "acserverview" => 7,
                "storedprocedure" or "acstoredprocedure" => 9,
                "function" or "acfunction" => 10,
                _ => throw new ArgumentException("object_type must be table, query, form, report, macro, module, server_view, stored_procedure, function, or an Access enum integer value.", paramName)
            };
        }

        private static object ParseCloseSaveOption(string? save)
        {
            if (string.IsNullOrWhiteSpace(save))
                return Type.Missing;

            var trimmed = save.Trim();
            if (int.TryParse(trimmed, out var numericSaveOption))
                return numericSaveOption;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "prompt" or "acsaveprompt" => 0,
                "yes" or "true" or "acsaveyes" => 1,
                "no" or "false" or "acsaveno" => 2,
                _ => throw new ArgumentException("save must be prompt, yes, no, or an Access enum integer value.", nameof(save))
            };
        }

        private static int ParseRunCommandValue(string command)
        {
            var trimmed = command.Trim();
            if (int.TryParse(trimmed, out var numericCommand))
                return numericCommand;

            var normalized = NormalizeEnumToken(trimmed);
            return normalized switch
            {
                "compileandsaveallmodules" or "accmdcompileandsaveallmodules" => 125,
                _ => throw new ArgumentException("command must be an acCommand integer value (or a supported acCmd constant name).", nameof(command))
            };
        }

        private static string NormalizeEnumToken(string value)
        {
            return value
                .Trim()
                .Replace("_", string.Empty, StringComparison.Ordinal)
                .Replace("-", string.Empty, StringComparison.Ordinal)
                .Replace(" ", string.Empty, StringComparison.Ordinal)
                .ToLowerInvariant();
        }

        private static object NormalizeDoCmdVariant(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return Type.Missing;

            var trimmed = value.Trim();
            if (int.TryParse(trimmed, out var numericValue))
                return numericValue;
            if (bool.TryParse(trimmed, out var boolValue))
                return boolValue;

            return trimmed;
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

        private static SectionInfo BuildSectionInfo(object section, int index, bool isReport)
        {
            var name = SafeToString(TryGetDynamicProperty(section, "Name"));
            var typeCode = ToNullableInt(TryGetDynamicProperty(section, "Section"));

            return new SectionInfo
            {
                Name = string.IsNullOrWhiteSpace(name) ? BuildSectionName(index, typeCode, isReport) : name!,
                Index = typeCode ?? index,
                Height = ToInt32(TryGetDynamicProperty(section, "Height")),
                Visible = ToBool(TryGetDynamicProperty(section, "Visible"), true),
                BackColor = ToNullableInt(TryGetDynamicProperty(section, "BackColor")),
                KeepTogether = ToNullableInt(TryGetDynamicProperty(section, "KeepTogether"))
            };
        }

        private static dynamic? TryGetReportGroupLevel(dynamic report, int index)
        {
            if (index < 0)
                return null;

            try
            {
                var groupLevels = TryGetDynamicProperty(report, "GroupLevels");
                if (groupLevels != null)
                {
                    var byItem = InvokeDynamicMethod(groupLevels, "Item", index);
                    if (byItem != null)
                        return byItem;
                }
            }
            catch
            {
                // Fall through to alternate access patterns.
            }

            try
            {
                var byMethod = InvokeDynamicMethod(report, "GroupLevel", index);
                if (byMethod != null)
                    return byMethod;
            }
            catch
            {
                // Fall through to final null.
            }

            return null;
        }

        private static ReportGroupingInfo BuildReportGroupingInfo(dynamic groupLevel, int index)
        {
            return new ReportGroupingInfo
            {
                Index = index,
                Expression = SafeToString(TryGetDynamicProperty(groupLevel, "ControlSource")) ?? "",
                SortOrder = ToNullableInt(TryGetDynamicProperty(groupLevel, "SortOrder")),
                GroupOn = ToNullableInt(TryGetDynamicProperty(groupLevel, "GroupOn")),
                GroupInterval = ToNullableInt(TryGetDynamicProperty(groupLevel, "GroupInterval")),
                GroupHeader = ToNullableBool(TryGetDynamicProperty(groupLevel, "GroupHeader")),
                GroupFooter = ToNullableBool(TryGetDynamicProperty(groupLevel, "GroupFooter")),
                KeepTogether = ToNullableInt(TryGetDynamicProperty(groupLevel, "KeepTogether"))
            };
        }

        private static PrinterInfoEntry BuildPrinterInfo(dynamic printer)
        {
            return new PrinterInfoEntry
            {
                DeviceName = SafeToString(TryGetDynamicProperty(printer, "DeviceName")) ?? "",
                DriverName = SafeToString(TryGetDynamicProperty(printer, "DriverName")),
                Port = SafeToString(TryGetDynamicProperty(printer, "Port")),
                Orientation = ToNullableInt(TryGetDynamicProperty(printer, "Orientation")),
                PaperSize = ToNullableInt(TryGetDynamicProperty(printer, "PaperSize")),
                TopMargin = ToNullableInt(TryGetDynamicProperty(printer, "TopMargin")),
                BottomMargin = ToNullableInt(TryGetDynamicProperty(printer, "BottomMargin")),
                LeftMargin = ToNullableInt(TryGetDynamicProperty(printer, "LeftMargin")),
                RightMargin = ToNullableInt(TryGetDynamicProperty(printer, "RightMargin"))
            };
        }

        private static string BuildSectionName(int index, int? typeCode, bool isReport)
        {
            var effective = typeCode ?? index;
            return (effective, isReport) switch
            {
                (0, _) => "Detail",
                (1, false) => "FormHeader",
                (2, false) => "FormFooter",
                (1, true) => "ReportHeader",
                (2, true) => "ReportFooter",
                (3, _) => "PageHeader",
                (4, _) => "PageFooter",
                _ => $"Section{effective}"
            };
        }

        private static List<object> GetSectionObjects(object formOrReport)
        {
            var sections = TryGetDynamicProperty(formOrReport, "Sections") ?? InvokeDynamicMethod(formOrReport, "Sections");
            if (sections == null)
                throw new InvalidOperationException("Sections collection is not available for this Access object.");

            var sectionObjects = new List<object>();
            foreach (var section in (dynamic)sections)
            {
                sectionObjects.Add(section);
            }

            return sectionObjects;
        }

        private static object? FindSection(object formOrReport, string sectionIdentifier)
        {
            var sections = GetSectionObjects(formOrReport);
            if (sections.Count == 0)
                return null;

            if (int.TryParse(sectionIdentifier.Trim(), out var parsedIndex))
            {
                return sections.FirstOrDefault(section =>
                {
                    var typeCode = ToNullableInt(TryGetDynamicProperty(section, "Section"));
                    if (typeCode.HasValue)
                        return typeCode.Value == parsedIndex;

                    var index = ToNullableInt(TryGetDynamicProperty(section, "Index"));
                    return index.HasValue && index.Value == parsedIndex;
                });
            }

            var normalizedIdentifier = NormalizeEnumToken(sectionIdentifier);
            var fallbackIndex = ParseSectionAlias(normalizedIdentifier);

            return sections.FirstOrDefault(section =>
            {
                var name = SafeToString(TryGetDynamicProperty(section, "Name"));
                if (!string.IsNullOrWhiteSpace(name) && NormalizeEnumToken(name) == normalizedIdentifier)
                    return true;

                if (!fallbackIndex.HasValue)
                    return false;

                var typeCode = ToNullableInt(TryGetDynamicProperty(section, "Section"));
                if (typeCode.HasValue)
                    return typeCode.Value == fallbackIndex.Value;

                var index = ToNullableInt(TryGetDynamicProperty(section, "Index"));
                return index.HasValue && index.Value == fallbackIndex.Value;
            });
        }

        private static int? ParseSectionAlias(string normalizedIdentifier)
        {
            return normalizedIdentifier switch
            {
                "detail" or "acdetail" => 0,
                "header" or "formheader" or "reportheader" or "acheader" => 1,
                "footer" or "formfooter" or "reportfooter" or "acfooter" => 2,
                "pageheader" or "acpageheader" => 3,
                "pagefooter" or "acpagefooter" => 4,
                _ => null
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
                var indexes = GetSchema("Indexes");
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

        private static bool IsColumnNullable(DataRow row)
        {
            var isNullable = GetRowString(row, "IS_NULLABLE");
            if (!string.IsNullOrWhiteSpace(isNullable))
                return string.Equals(isNullable, "YES", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(isNullable, "TRUE", StringComparison.OrdinalIgnoreCase);

            var nullableCode = GetRowInt(row, "NULLABLE");
            return nullableCode.HasValue && nullableCode.Value != 0;
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

    public class LinkedTableInfo
    {
        public string Name { get; set; } = "";
        public string SourceTableName { get; set; } = "";
        public string ConnectString { get; set; } = "";
        public string SourceDatabasePath { get; set; } = "";
        public int Attributes { get; set; }
    }

    public class TransactionStatusInfo
    {
        public bool Active { get; set; }
        public string? IsolationLevel { get; set; }
        public DateTimeOffset? StartedAtUtc { get; set; }
    }

    public class DatabaseCreateResult
    {
        public string DatabasePath { get; set; } = "";
        public bool ExistedBefore { get; set; }
        public long SizeBytes { get; set; }
        public DateTime LastWriteTimeUtc { get; set; }
    }

    public class DatabaseBackupResult
    {
        public string SourceDatabasePath { get; set; } = "";
        public string DestinationDatabasePath { get; set; } = "";
        public long BytesCopied { get; set; }
        public DateTime SourceLastWriteTimeUtc { get; set; }
        public DateTime DestinationLastWriteTimeUtc { get; set; }
        public bool OperatedOnConnectedDatabase { get; set; }
    }

    public class DatabaseCompactRepairResult
    {
        public string SourceDatabasePath { get; set; } = "";
        public string DestinationDatabasePath { get; set; } = "";
        public bool InPlace { get; set; }
        public long SourceSizeBytes { get; set; }
        public long DestinationSizeBytes { get; set; }
        public DateTime DestinationLastWriteTimeUtc { get; set; }
        public bool OperatedOnConnectedDatabase { get; set; }
    }

    public class TransferSpreadsheetResult
    {
        public int TransferType { get; set; }
        public int SpreadsheetType { get; set; }
        public string TableName { get; set; } = "";
        public string FileName { get; set; } = "";
        public bool HasFieldNames { get; set; }
        public string? Range { get; set; }
        public bool UseOA { get; set; }
    }

    public class TransferTextResult
    {
        public int TransferType { get; set; }
        public string? SpecificationName { get; set; }
        public string TableName { get; set; } = "";
        public string FileName { get; set; } = "";
        public bool HasFieldNames { get; set; }
        public string? HtmlTableName { get; set; }
        public int? CodePage { get; set; }
    }

    public class OutputToResult
    {
        public int ObjectType { get; set; }
        public string? ObjectName { get; set; }
        public string OutputFormat { get; set; } = "";
        public string? OutputFile { get; set; }
        public bool AutoStart { get; set; }
        public string? TemplateFile { get; set; }
        public string? Encoding { get; set; }
        public int? OutputQuality { get; set; }
    }

    public class TransferDatabaseResult
    {
        public int TransferType { get; set; }
        public string DatabaseType { get; set; } = "";
        public string DatabaseName { get; set; } = "";
        public int ObjectType { get; set; }
        public string Source { get; set; } = "";
        public string? Destination { get; set; }
        public bool StructureOnly { get; set; }
        public bool StoreLogin { get; set; }
    }

    public class DatabasePropertyInfo
    {
        public string Name { get; set; } = "";
        public object? Value { get; set; }
        public int TypeCode { get; set; }
        public bool IsSystem { get; set; }
    }

    public class TablePropertiesInfo
    {
        public string TableName { get; set; } = "";
        public string? Description { get; set; }
        public string? ValidationRule { get; set; }
        public string? ValidationText { get; set; }
    }

    public class TableValidationInfo
    {
        public string TableName { get; set; } = "";
        public string? ValidationRule { get; set; }
        public string? ValidationText { get; set; }
    }

    public class TableFieldDescriptionInfo
    {
        public string TableName { get; set; } = "";
        public string FieldName { get; set; } = "";
        public string? Description { get; set; }
    }

    public class QueryPropertiesInfo
    {
        public string QueryName { get; set; } = "";
        public string? Description { get; set; }
        public string Sql { get; set; } = "";
        public List<QueryParameterInfo> Parameters { get; set; } = new();
    }

    public class QueryParameterInfo
    {
        public string Name { get; set; } = "";
        public int TypeCode { get; set; }
        public object? Value { get; set; }
    }

    public class FieldPropertiesInfo
    {
        public string TableName { get; set; } = "";
        public string FieldName { get; set; } = "";
        public int TypeCode { get; set; }
        public int Size { get; set; }
        public bool Required { get; set; }
        public bool AllowZeroLength { get; set; }
        public string? DefaultValue { get; set; }
        public string? ValidationRule { get; set; }
        public string? ValidationText { get; set; }
        public string? InputMask { get; set; }
        public string? Caption { get; set; }
        public string? RowSource { get; set; }
        public int? BoundColumn { get; set; }
        public int? ColumnCount { get; set; }
        public string? ColumnWidths { get; set; }
        public bool? LimitToList { get; set; }
        public bool? AllowMultipleValues { get; set; }
        public int? DisplayControl { get; set; }
    }

    public class FieldAttributesInfo
    {
        public string TableName { get; set; } = "";
        public string FieldName { get; set; } = "";
        public int TypeCode { get; set; }
        public int Size { get; set; }
        public bool Required { get; set; }
        public bool AllowZeroLength { get; set; }
        public int Attributes { get; set; }
        public bool? AllowMultipleValues { get; set; }
        public bool? IsComplex { get; set; }
    }

    public class MultiValueFieldInfo
    {
        public string TableName { get; set; } = "";
        public string FieldName { get; set; } = "";
        public int Attributes { get; set; }
        public bool? AllowMultipleValues { get; set; }
        public bool? IsComplex { get; set; }
    }

    public class MultiValueFieldRecordInfo
    {
        public string TableName { get; set; } = "";
        public string FieldName { get; set; } = "";
        public int RowIndex { get; set; }
        public List<object?> Values { get; set; } = new();
        public List<Dictionary<string, object?>> Entries { get; set; } = new();
    }

    public class MultiValueFieldUpdateInfo
    {
        public string TableName { get; set; } = "";
        public string FieldName { get; set; } = "";
        public int ValuesWritten { get; set; }
    }

    public class VBAReferenceInfo
    {
        public string Name { get; set; } = "";
        public string Guid { get; set; } = "";
        public int Major { get; set; }
        public int Minor { get; set; }
        public string FullPath { get; set; } = "";
        public string Description { get; set; } = "";
        public bool BuiltIn { get; set; }
        public bool IsBroken { get; set; }
    }

    public class StartupPropertiesInfo
    {
        public string? StartupForm { get; set; }
        public string? AppTitle { get; set; }
        public string? AppIcon { get; set; }
    }

    public class RibbonInfo
    {
        public string? RibbonName { get; set; }
        public string? RibbonXml { get; set; }
        public string? DefaultRibbonName { get; set; }
        public bool Exists { get; set; }
    }

    public class ApplicationInfo
    {
        public string Name { get; set; } = "";
        public string? Version { get; set; }
        public string? Build { get; set; }
        public bool IsTrusted { get; set; }
        public string? CurrentDatabasePath { get; set; }
        public string? CurrentProjectName { get; set; }
        public string? CurrentProjectPath { get; set; }
        public string? CurrentProjectFullName { get; set; }
        public string? CurrentDataName { get; set; }
        public string? CurrentDataPath { get; set; }
    }

    public class CurrentProjectDataInfo
    {
        public string? CurrentProjectName { get; set; }
        public string? CurrentProjectPath { get; set; }
        public string? CurrentProjectFullName { get; set; }
        public string? CurrentDataName { get; set; }
        public string? CurrentDataPath { get; set; }
        public int? CurrentDataAllTablesCount { get; set; }
        public int? CurrentDataAllQueriesCount { get; set; }
    }

    public class DatabaseSecurityInfo
    {
        public string DatabasePath { get; set; } = "";
        public bool PasswordConfigured { get; set; }
        public bool Encrypted { get; set; }
    }

    public class NavigationGroupInfo
    {
        public string Name { get; set; } = "";
        public int ItemCount { get; set; }
    }

    public class NavigationGroupObjectInfo
    {
        public string GroupName { get; set; } = "";
        public string Name { get; set; } = "";
        public string ObjectName { get; set; } = "";
        public string? ObjectType { get; set; }
        public int Index { get; set; }
    }

    public class ConditionalFormattingRuleInfo
    {
        public int Index { get; set; }
        public int TypeCode { get; set; }
        public string? Expression1 { get; set; }
        public string? Expression2 { get; set; }
        public int? ForeColor { get; set; }
        public int? BackColor { get; set; }
        public bool Enabled { get; set; }
    }

    public class ConditionalFormattingControlInfo
    {
        public string ObjectType { get; set; } = "";
        public string ObjectName { get; set; } = "";
        public string ControlName { get; set; } = "";
        public int RuleCount { get; set; }
        public List<ConditionalFormattingRuleInfo> Rules { get; set; } = new();
    }

    public class AttachmentFileInfo
    {
        public string FileName { get; set; } = "";
        public string? FileType { get; set; }
        public int? FileSize { get; set; }
    }

    public class AttachmentSaveResult
    {
        public string TableName { get; set; } = "";
        public string FieldName { get; set; } = "";
        public string FileName { get; set; } = "";
        public string FilePath { get; set; } = "";
        public int FileSize { get; set; }
    }

    public class AttachmentMetadataInfo
    {
        public string FileName { get; set; } = "";
        public string? FileType { get; set; }
        public string? FileUrl { get; set; }
        public int? FileSize { get; set; }
        public Dictionary<string, object?> Fields { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    public class ObjectEventInfo
    {
        public string EventName { get; set; } = "";
        public string Handler { get; set; } = "";
    }

    public class DataMacroInfo
    {
        public string Name { get; set; } = "";
        public string? MacroType { get; set; }
    }

    public class AutoExecInfo
    {
        public bool Exists { get; set; }
        public string MacroName { get; set; } = "";
    }

    public class DaoContainerInfo
    {
        public string Name { get; set; } = "";
        public int DocumentCount { get; set; }
    }

    public class DaoContainerDocumentInfo
    {
        public string ContainerName { get; set; } = "";
        public string Name { get; set; } = "";
        public string? Owner { get; set; }
        public string? DateCreated { get; set; }
        public string? LastUpdated { get; set; }
    }

    public class DaoDocumentPropertyInfo
    {
        public string ContainerName { get; set; } = "";
        public string DocumentName { get; set; } = "";
        public string Name { get; set; } = "";
        public int TypeCode { get; set; }
        public object? Value { get; set; }
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

    public class ModuleAnalysisInfo
    {
        public string ProjectName { get; set; } = "";
        public string ModuleName { get; set; } = "";
        public string ModuleType { get; set; } = "";
        public int LineCount { get; set; }
        public int DeclarationLineCount { get; set; }
        public int ProcedureCount { get; set; }
    }

    public class ModuleProcedureInfo
    {
        public string Name { get; set; } = "";
        public string ProcedureType { get; set; } = "";
        public int StartLine { get; set; }
        public int LineCount { get; set; }
    }

    public class ModuleFindResult
    {
        public bool Found { get; set; }
        public int? StartLine { get; set; }
        public int? StartColumn { get; set; }
        public int? EndLine { get; set; }
        public int? EndColumn { get; set; }
    }

    public class ImportExportSpecificationInfo
    {
        public string Name { get; set; } = "";
        public string? Description { get; set; }
        public string? Xml { get; set; }
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

    public class SectionInfo
    {
        public string Name { get; set; } = "";
        public int Index { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public int? BackColor { get; set; }
        public int? KeepTogether { get; set; }
    }

    public class FormDesignPropertiesInfo
    {
        public string FormName { get; set; } = "";
        public string? RecordSource { get; set; }
        public int? DefaultView { get; set; }
        public bool? AllowEdits { get; set; }
        public bool? AllowAdditions { get; set; }
        public bool? AllowDeletions { get; set; }
        public bool? DataEntry { get; set; }
        public bool? NavigationButtons { get; set; }
        public bool? DividingLines { get; set; }
        public int? ScrollBars { get; set; }
        public string? Caption { get; set; }
        public bool? Modal { get; set; }
        public bool? PopUp { get; set; }
    }

    public class ReportDesignPropertiesInfo
    {
        public string ReportName { get; set; } = "";
        public string? RecordSource { get; set; }
        public int? DefaultView { get; set; }
        public string? Caption { get; set; }
        public int? ForceNewPage { get; set; }
        public int? KeepTogether { get; set; }
        public bool? PopUp { get; set; }
        public bool? Modal { get; set; }
        public int? ColumnCount { get; set; }
        public int? ColumnSpacing { get; set; }
    }

    public class TabOrderEntryInfo
    {
        public string ControlName { get; set; } = "";
        public int TabIndex { get; set; }
        public bool TabStop { get; set; }
    }

    public class PageSetupInfo
    {
        public string ObjectType { get; set; } = "";
        public string ObjectName { get; set; } = "";
        public int? TopMargin { get; set; }
        public int? BottomMargin { get; set; }
        public int? LeftMargin { get; set; }
        public int? RightMargin { get; set; }
        public int? Orientation { get; set; }
        public int? PaperSize { get; set; }
        public bool? DataOnly { get; set; }
    }

    public class ReportGroupingInfo
    {
        public int Index { get; set; }
        public string Expression { get; set; } = "";
        public int? SortOrder { get; set; }
        public int? GroupOn { get; set; }
        public int? GroupInterval { get; set; }
        public bool? GroupHeader { get; set; }
        public bool? GroupFooter { get; set; }
        public int? KeepTogether { get; set; }
    }

    public class ReportSortingInfo
    {
        public string ReportName { get; set; } = "";
        public string? OrderBy { get; set; }
        public bool? OrderByOn { get; set; }
        public List<ReportGroupingInfo> GroupLevels { get; set; } = new();
    }

    public class PrinterInfoResult
    {
        public PrinterInfoEntry? ActivePrinter { get; set; }
        public List<PrinterInfoEntry> Printers { get; set; } = new();
    }

    public class PrinterInfoEntry
    {
        public string DeviceName { get; set; } = "";
        public string? DriverName { get; set; }
        public string? Port { get; set; }
        public int? Orientation { get; set; }
        public int? PaperSize { get; set; }
        public int? TopMargin { get; set; }
        public int? BottomMargin { get; set; }
        public int? LeftMargin { get; set; }
        public int? RightMargin { get; set; }
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
