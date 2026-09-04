using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Data.Sqlite;
using ML.ImportExcel.Core.Models;

namespace ML.ImportExcel.Core.Storage
{
    /// <summary>
    /// 基于 SQLite 的导入配置存储。默认库文件为应用目录下的 excelimport.db，可通过选项配置路径。
    /// </summary>
    public sealed class SqliteImportStore : IDisposable
    {
        private const string DateTimeFormat = "yyyy-MM-dd HH:mm:ss.fff";
        private readonly string _dbPath;
        private readonly string _connString;
        private static int _sqliteInit;

        /// <summary>当前数据库文件完整路径。</summary>
        public string DatabasePath => _dbPath;

        public SqliteImportStore(string databasePath)
        {
            if (string.IsNullOrWhiteSpace(databasePath))
                throw new ArgumentNullException(nameof(databasePath));

            _dbPath = Path.GetFullPath(databasePath);
            var dir = Path.GetDirectoryName(_dbPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            EnsureSqliteInitialized();
            _connString = new SqliteConnectionStringBuilder
            {
                DataSource = _dbPath,
                Mode = SqliteOpenMode.ReadWriteCreate,
                Cache = SqliteCacheMode.Shared
            }.ToString();

            CreateSchema();
        }

        private static void EnsureSqliteInitialized()
        {
            if (System.Threading.Interlocked.Exchange(ref _sqliteInit, 1) == 0)
            {
                try
                {
                    // bundle_e_sqlite3 自带的初始化，加载 e_sqlite3 原生库
                    SQLitePCL.Batteries_V2.Init();
                }
                catch
                {
                    // 若宿主已自行初始化过可忽略
                }
            }
        }

        private SqliteConnection OpenConnection()
        {
            EnsureSqliteInitialized();
            var c = new SqliteConnection(_connString);
            c.Open();
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "PRAGMA journal_mode=WAL; PRAGMA foreign_keys=ON;";
                cmd.ExecuteNonQuery();
            }
            return c;
        }

        private static void AddParameter(SqliteCommand cmd, string name, object value)
        {
            cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        }

        private void CreateSchema()
        {
            using (var c = OpenConnection())
            {
                var sql = @"
CREATE TABLE IF NOT EXISTS import_configs (
    Id              INTEGER PRIMARY KEY AUTOINCREMENT,
    Name            TEXT    NOT NULL UNIQUE,
    EntityName      TEXT    NULL,
    Description     TEXT    NULL,
    SheetIndex      INTEGER NOT NULL DEFAULT 0,
    SheetName       TEXT    NULL,
    HeaderRowIndex  INTEGER NOT NULL DEFAULT 0,
    FolderId        INTEGER NULL,
    CreatedAt       TEXT    NOT NULL,
    UpdatedAt       TEXT    NOT NULL
);

CREATE TABLE IF NOT EXISTS import_config_fields (
    Id           INTEGER PRIMARY KEY AUTOINCREMENT,
    ConfigId     INTEGER NOT NULL REFERENCES import_configs(Id) ON DELETE CASCADE,
    ColumnIndex  INTEGER NOT NULL,
    Header       TEXT    NOT NULL,
    Field        TEXT    NOT NULL,
    Required     INTEGER NOT NULL DEFAULT 0,
    DataType     TEXT    NULL
);

CREATE TABLE IF NOT EXISTS import_folders (
    Id        INTEGER PRIMARY KEY AUTOINCREMENT,
    ParentId  INTEGER NULL REFERENCES import_folders(Id) ON DELETE CASCADE,
    Name      TEXT    NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_import_config_fields_configId ON import_config_fields(ConfigId);
CREATE INDEX IF NOT EXISTS idx_import_folders_parentId ON import_folders(ParentId);
CREATE UNIQUE INDEX IF NOT EXISTS uq_import_folders_sibling
    ON import_folders(COALESCE(ParentId, -1), Name);
";
                using (var cmd = c.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }

                // 老库升级：补充 FolderId 列（幂等），之后才能建其索引
                EnsureColumn(c, "import_configs", "FolderId", "ALTER TABLE import_configs ADD COLUMN FolderId INTEGER NULL;");

                using (var idx = c.CreateCommand())
                {
                    idx.CommandText = "CREATE INDEX IF NOT EXISTS idx_import_configs_folderId ON import_configs(FolderId);";
                    idx.ExecuteNonQuery();
                }
            }
        }

        /// <summary>老库升级：若列不存在则补充（幂等）。</summary>
        private static void EnsureColumn(SqliteConnection c, string table, string column, string alterSql)
        {
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(1) FROM pragma_table_info($table) WHERE name = $col;";
                AddParameter(cmd, "$table", table);
                AddParameter(cmd, "$col", column);
                var exists = Convert.ToInt64(cmd.ExecuteScalar()) > 0;
                if (!exists)
                {
                    using (var alt = c.CreateCommand())
                    {
                        alt.CommandText = alterSql;
                        alt.ExecuteNonQuery();
                    }
                }
            }
        }

        // ---------------- 查询 ----------------

        /// <summary>列出全部导入配置（含字段映射）。</summary>
        public List<ImportConfig> ListConfigs()
        {
            var result = new List<ImportConfig>();
            using (var c = OpenConnection())
            {
                using (var cmd = c.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM import_configs ORDER BY Id ASC;";
                    using (var rd = cmd.ExecuteReader())
                    {
                        while (rd.Read())
                            result.Add(ReadConfig(rd));
                    }
                }
                if (result.Count == 0) return result;
                LoadFields(c, result);
            }
            return result;
        }

        /// <summary>按 Id 获取配置。</summary>
        public ImportConfig GetConfig(long id)
        {
            using (var c = OpenConnection())
            {
                var cfg = QuerySingle(c, "Id = $id", new[] { ("$id", (object)id) });
                if (cfg != null) LoadFields(c, new List<ImportConfig> { cfg });
                return cfg;
            }
        }

        /// <summary>按配置名称或实体名称获取（先按 Name，再按 EntityName）。</summary>
        public ImportConfig GetByNameOrEntity(string nameOrEntity)
        {
            using (var c = OpenConnection())
            {
                var cfg = QuerySingle(c, "Name = $name", new[] { ("$name", (object)nameOrEntity) });
                if (cfg == null)
                    cfg = QuerySingle(c, "EntityName = $name", new[] { ("$name", (object)nameOrEntity) });
                if (cfg != null) LoadFields(c, new List<ImportConfig> { cfg });
                return cfg;
            }
        }

        private ImportConfig QuerySingle(SqliteConnection c, string where, (string, object)[] prms)
        {
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM import_configs WHERE " + where + " LIMIT 1;";
                foreach (var (k, v) in prms) AddParameter(cmd, k, v);
                using (var rd = cmd.ExecuteReader())
                {
                    return rd.Read() ? ReadConfig(rd) : null;
                }
            }
        }

        private static void LoadFields(SqliteConnection c, IList<ImportConfig> configs)
        {
            if (configs.Count == 0) return;
            var byId = configs.ToDictionary(x => x.Id);
            using (var cmd = c.CreateCommand())
            {
                var ids = string.Join(",", byId.Keys);
                cmd.CommandText = "SELECT * FROM import_config_fields WHERE ConfigId IN (" + ids + ") ORDER BY ColumnIndex ASC;";
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        var f = ReadField(rd);
                        if (byId.TryGetValue(f.ConfigId, out var cfg))
                            cfg.Fields.Add(f);
                    }
                }
            }
        }

        private static ImportConfig ReadConfig(SqliteDataReader rd)
        {
            var cfg = new ImportConfig
            {
                Id = rd.GetInt64(rd.GetOrdinal("Id")),
                Name = rd.IsDBNull(rd.GetOrdinal("Name")) ? "" : rd.GetString(rd.GetOrdinal("Name")),
                EntityName = rd.IsDBNull(rd.GetOrdinal("EntityName")) ? "" : rd.GetString(rd.GetOrdinal("EntityName")),
                Description = rd.IsDBNull(rd.GetOrdinal("Description")) ? "" : rd.GetString(rd.GetOrdinal("Description")),
                SheetIndex = rd.GetInt32(rd.GetOrdinal("SheetIndex")),
                SheetName = rd.IsDBNull(rd.GetOrdinal("SheetName")) ? "" : rd.GetString(rd.GetOrdinal("SheetName")),
                HeaderRowIndex = rd.GetInt32(rd.GetOrdinal("HeaderRowIndex"))
            };
            var fid = rd.GetOrdinal("FolderId");
            cfg.FolderId = rd.IsDBNull(fid) ? (long?)null : rd.GetInt64(fid);
            cfg.CreatedAt = rd.IsDBNull(rd.GetOrdinal("CreatedAt")) ? DateTime.Now : ParseDate(rd.GetString(rd.GetOrdinal("CreatedAt")));
            cfg.UpdatedAt = rd.IsDBNull(rd.GetOrdinal("UpdatedAt")) ? DateTime.Now : ParseDate(rd.GetString(rd.GetOrdinal("UpdatedAt")));
            return cfg;
        }

        private static ImportFieldMapping ReadField(SqliteDataReader rd)
        {
            return new ImportFieldMapping
            {
                Id = rd.GetInt64(rd.GetOrdinal("Id")),
                ConfigId = rd.GetInt64(rd.GetOrdinal("ConfigId")),
                ColumnIndex = rd.GetInt32(rd.GetOrdinal("ColumnIndex")),
                Header = rd.IsDBNull(rd.GetOrdinal("Header")) ? "" : rd.GetString(rd.GetOrdinal("Header")),
                Field = rd.IsDBNull(rd.GetOrdinal("Field")) ? "" : rd.GetString(rd.GetOrdinal("Field")),
                Required = rd.GetInt32(rd.GetOrdinal("Required")) != 0,
                DataType = rd.IsDBNull(rd.GetOrdinal("DataType")) ? "" : rd.GetString(rd.GetOrdinal("DataType"))
            };
        }

        // ---------------- 写操作 ----------------

        /// <summary>
        /// 新增或更新一份配置（含字段映射）。Id &lt;= 0 视为新增。
        /// 保存后返回持久化后的完整配置。
        /// </summary>
        public ImportConfig SaveConfig(ImportConfig cfg)
        {
            if (cfg == null) throw new ArgumentNullException(nameof(cfg));
            if (string.IsNullOrWhiteSpace(cfg.Name))
                throw new ArgumentException("配置名称不能为空。", nameof(cfg));

            using (var c = OpenConnection())
            using (var tx = c.BeginTransaction())
            {
                long id;
                var now = DateTime.Now;
                using (var cmd = c.CreateCommand())
                {
                    cmd.Transaction = tx;
                    if (cfg.Id > 0 && Exists(c, tx, cfg.Id))
                    {
                        cmd.CommandText = @"UPDATE import_configs SET Name=$name, EntityName=$entity, Description=$desc,
                                              SheetIndex=$sheetIndex, SheetName=$sheetName, HeaderRowIndex=$headerRow,
                                              FolderId=$folder, UpdatedAt=$updated WHERE Id=$id;";
                        AddParameter(cmd, "$id", cfg.Id);
                        id = cfg.Id;
                    }
                    else
                    {
                        cmd.CommandText = @"INSERT INTO import_configs (Name, EntityName, Description, SheetIndex, SheetName, HeaderRowIndex, FolderId, CreatedAt, UpdatedAt)
                                            VALUES ($name,$entity,$desc,$sheetIndex,$sheetName,$headerRow,$folder,$created,$updated);";
                        AddParameter(cmd, "$created", now.ToString(DateTimeFormat));
                        id = 0;
                    }
                    AddParameter(cmd, "$name", cfg.Name?.Trim());
                    AddParameter(cmd, "$entity", string.IsNullOrWhiteSpace(cfg.EntityName) ? null : cfg.EntityName.Trim());
                    AddParameter(cmd, "$desc", string.IsNullOrWhiteSpace(cfg.Description) ? null : cfg.Description);
                    AddParameter(cmd, "$sheetIndex", cfg.SheetIndex);
                    AddParameter(cmd, "$sheetName", string.IsNullOrWhiteSpace(cfg.SheetName) ? null : cfg.SheetName);
                    AddParameter(cmd, "$headerRow", cfg.HeaderRowIndex);
                    AddParameter(cmd, "$folder", cfg.FolderId.HasValue ? (object)cfg.FolderId.Value : null);
                    AddParameter(cmd, "$updated", now.ToString(DateTimeFormat));
                    cmd.ExecuteNonQuery();

                    if (id == 0)
                    {
                        id = (long)LastInsertRowId(c, tx);
                    }
                }

                // 删除旧的字段映射并重建
                using (var del = c.CreateCommand())
                {
                    del.Transaction = tx;
                    del.CommandText = "DELETE FROM import_config_fields WHERE ConfigId=$id;";
                    AddParameter(del, "$id", id);
                    del.ExecuteNonQuery();
                }

                if (cfg.Fields != null)
                {
                    foreach (var f in cfg.Fields)
                    {
                        var header = string.IsNullOrWhiteSpace(f.Header) ? "" : f.Header.Trim();
                        var field = string.IsNullOrWhiteSpace(f.Field) ? "" : f.Field.Trim();
                        if (header.Length == 0 && field.Length == 0) continue;

                        using (var ins = c.CreateCommand())
                        {
                            ins.Transaction = tx;
                            ins.CommandText = @"INSERT INTO import_config_fields (ConfigId, ColumnIndex, Header, Field, Required, DataType)
                                                VALUES ($configId,$col,$header,$field,$required,$dataType);";
                            AddParameter(ins, "$configId", id);
                            AddParameter(ins, "$col", f.ColumnIndex);
                            AddParameter(ins, "$header", header);
                            AddParameter(ins, "$field", field);
                            AddParameter(ins, "$required", f.Required ? 1 : 0);
                            AddParameter(ins, "$dataType", string.IsNullOrWhiteSpace(f.DataType) ? null : f.DataType.Trim());
                            ins.ExecuteNonQuery();
                        }
                    }
                }

                tx.Commit();

                var saved = GetConfig(id);
                if (saved == null) throw new InvalidOperationException("保存配置失败。");
                return saved;
            }
        }

        private static bool Exists(SqliteConnection c, SqliteTransaction tx, long id)
        {
            using (var cmd = c.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "SELECT COUNT(1) FROM import_configs WHERE Id=$id;";
                AddParameter(cmd, "$id", id);
                return Convert.ToInt64(cmd.ExecuteScalar()) > 0;
            }
        }

        private static long LastInsertRowId(SqliteConnection c, SqliteTransaction tx)
        {
            using (var cmd = c.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "SELECT last_insert_rowid();";
                return Convert.ToInt64(cmd.ExecuteScalar());
            }
        }

        /// <summary>删除一份配置。</summary>
        public void DeleteConfig(long id)
        {
            using (var c = OpenConnection())
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM import_configs WHERE Id=$id;";
                AddParameter(cmd, "$id", id);
                cmd.ExecuteNonQuery();
            }
        }

        // ================= 文件夹（树）管理 =================

        /// <summary>列出全部文件夹（扁平结构，由前端按 ParentId 组装树）。</summary>
        public List<ImportFolder> ListFolders()
        {
            var list = new List<ImportFolder>();
            using (var c = OpenConnection())
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "SELECT Id, ParentId, Name FROM import_folders ORDER BY Name ASC;";
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                        list.Add(ReadFolder(rd));
                }
            }
            return list;
        }

        /// <summary>按 Id 获取文件夹。</summary>
        public ImportFolder GetFolder(long id)
        {
            using (var c = OpenConnection())
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "SELECT Id, ParentId, Name FROM import_folders WHERE Id=$id LIMIT 1;";
                AddParameter(cmd, "$id", id);
                using (var rd = cmd.ExecuteReader())
                    return rd.Read() ? ReadFolder(rd) : null;
            }
        }

        /// <summary>新增或更新文件夹；返回持久化后的对象。</summary>
        public ImportFolder SaveFolder(ImportFolder folder)
        {
            if (folder == null) throw new ArgumentNullException(nameof(folder));
            if (string.IsNullOrWhiteSpace(folder.Name))
                throw new ArgumentException("文件夹名称不能为空。", nameof(folder));
            folder.Name = folder.Name.Trim();

            using (var c = OpenConnection())
            {
                // 校验父文件夹存在
                if (folder.ParentId.HasValue)
                {
                    var parent = GetFolder(folder.ParentId.Value);
                    if (parent == null)
                        throw new ArgumentException("父文件夹不存在。");
                    // 防止把文件夹移动到自身或其子文件夹下造成循环
                    var cur = parent;
                    while (cur != null)
                    {
                        if (cur.Id == folder.Id)
                            throw new ArgumentException("不能将文件夹移动到自身或其子文件夹内。");
                        cur = cur.ParentId.HasValue ? GetFolder(cur.ParentId.Value) : null;
                    }
                }

                // 同级名称唯一
                using (var chk = c.CreateCommand())
                {
                    chk.CommandText = @"SELECT COUNT(1) FROM import_folders
                                        WHERE COALESCE(ParentId,-1) = COALESCE($parent,-1) AND Name = $name AND Id <> $id;";
                    AddParameter(chk, "$parent", folder.ParentId.HasValue ? (object)folder.ParentId.Value : null);
                    AddParameter(chk, "$name", folder.Name);
                    AddParameter(chk, "$id", folder.Id);
                    if (Convert.ToInt64(chk.ExecuteScalar()) > 0)
                        throw new ArgumentException($"同级已存在文件夹“{folder.Name}”。");
                }

                using (var cmd = c.CreateCommand())
                {
                    if (folder.Id > 0 && GetFolder(folder.Id) != null)
                    {
                        cmd.CommandText = "UPDATE import_folders SET Name=$name, ParentId=$parent WHERE Id=$id;";
                        AddParameter(cmd, "$id", folder.Id);
                    }
                    else
                    {
                        cmd.CommandText = "INSERT INTO import_folders (Name, ParentId) VALUES ($name, $parent);";
                        folder.Id = 0;
                    }
                    AddParameter(cmd, "$name", folder.Name);
                    AddParameter(cmd, "$parent", folder.ParentId.HasValue ? (object)folder.ParentId.Value : null);
                    cmd.ExecuteNonQuery();
                    if (folder.Id == 0)
                        folder.Id = (long)LastInsertRowId(c, null);
                }

                var saved = GetFolder(folder.Id);
                if (saved == null) throw new InvalidOperationException("保存文件夹失败。");
                return saved;
            }
        }

        /// <summary>
        /// 删除文件夹及其全部子孙文件夹；其下配置移动到根目录（FolderId 置空）。
        /// </summary>
        public void DeleteFolderTree(long id)
        {
            using (var c = OpenConnection())
            {
                // 收集子树所有 Id
                var ids = new List<long> { id };
                CollectSubFolderIds(c, id, ids);

                var inExpr = string.Join(",", ids);
                using (var up = c.CreateCommand())
                {
                    up.CommandText = "UPDATE import_configs SET FolderId = NULL WHERE FolderId IN (" + inExpr + ");";
                    up.ExecuteNonQuery();
                }
                using (var del = c.CreateCommand())
                {
                    del.CommandText = "DELETE FROM import_folders WHERE Id IN (" + inExpr + ");";
                    del.ExecuteNonQuery();
                }
            }
        }

        private static void CollectSubFolderIds(SqliteConnection c, long parentId, List<long> acc)
        {
            using (var cmd = c.CreateCommand())
            {
                cmd.CommandText = "SELECT Id FROM import_folders WHERE ParentId = $p;";
                AddParameter(cmd, "$p", parentId);
                using (var rd = cmd.ExecuteReader())
                {
                    var kids = new List<long>();
                    while (rd.Read()) kids.Add(rd.GetInt64(0));
                    foreach (var k in kids)
                    {
                        acc.Add(k);
                        CollectSubFolderIds(c, k, acc);
                    }
                }
            }
        }

        private static ImportFolder ReadFolder(SqliteDataReader rd)
        {
            return new ImportFolder
            {
                Id = rd.GetInt64(rd.GetOrdinal("Id")),
                ParentId = rd.IsDBNull(rd.GetOrdinal("ParentId")) ? (long?)null : rd.GetInt64(rd.GetOrdinal("ParentId")),
                Name = rd.IsDBNull(rd.GetOrdinal("Name")) ? "" : rd.GetString(rd.GetOrdinal("Name"))
            };
        }

        private static DateTime ParseDate(string s)
        {
            return DateTime.TryParse(s, out var dt) ? dt : DateTime.MinValue;
        }

        public void Dispose()
        {
            // 无长连接资源需要释放
        }
    }
}
