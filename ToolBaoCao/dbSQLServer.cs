using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace ToolBaoCao
{
    public class dbSQLServer : IDisposable
    {
        public SqlConnectionStringBuilder ConnectionString = new SqlConnectionStringBuilder();
        private SqlConnection connection = new SqlConnection();
        public int CommandTimeOut = 0;

        public static bool IsUpdateData(string tsql)
        {
            /* Xóa các chuỗi trong nháy đơn (để tránh bắt từ khóa bên trong văn bản) */
            string sanitizedSql = Regex.Replace(tsql, @"'[^']*'", string.Empty, RegexOptions.IgnoreCase);
            return Regex.IsMatch(sanitizedSql, @"\b(UPDATE|DELETE|INSERT|MERGE|DROP)\b(?!.*')", RegexOptions.IgnoreCase);
        }

        public dbSQLServer(string connectionString = "")
        {
            if (string.IsNullOrEmpty(connectionString) == false)
            {
                try { ConnectionString = new SqlConnectionStringBuilder(connectionString); }
                catch { connectionString = ""; }
                Close();
                connection = new SqlConnection(connectionString);
            }
        }

        public string like(string field, string value)
        {
            if (Regex.IsMatch(value, "[%*_?]") == false) { return $"{field} = '{value.Replace("'", "''")}'"; }
            value = value.Replace("*", "%").Replace("?", "_");
            value = Regex.Replace(value, "[%]+", "%");
            return $"{field} LIKE '{value.Replace("'", "''")}'";
        }

        private SqlParameter[] ConvertObjectToParameter(object parameters)
        {
            if (parameters == null) { return null; }
            if (parameters is KeyValuePair<string, string> obj1) { return new SqlParameter[] { new SqlParameter(obj1.Key, obj1.Value) }; }
            if (parameters is KeyValuePair<string, object> obj2) { return new SqlParameter[] { new SqlParameter(obj2.Key, obj2.Value) }; }
            if (parameters is SqlParameter obj5) { return new SqlParameter[] { obj5 }; }
            if (parameters is SqlParameter[] obj6) { return obj6; }
            if (parameters is Dictionary<string, string> obj3)
            {
                var par = new List<SqlParameter>();
                foreach (var v in obj3) { par.Add(new SqlParameter(v.Key, v.Value)); }
                return par.ToArray();
            }
            if (parameters is Dictionary<string, object> obj4)
            {
                var par = new List<SqlParameter>();
                foreach (var v in obj4) { par.Add(new SqlParameter(v.Key, v.Value)); }
                return par.ToArray();
            }
            throw new Exception($"Not support SQLiteParameter ${parameters}");
        }

        public DataTable getDataTable(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            DataTable data = new DataTable();
            if (connection.State == ConnectionState.Closed) { connection.Open(); }
            using (var command = new SqlCommand(query, connection))
            {
                if (CommandTimeOut > 0) { command.CommandTimeout = CommandTimeOut; }
                if (par != null) { command.Parameters.AddRange(par); }
                using (var adapter = new SqlDataAdapter(command))
                {
                    var dataset = new System.Data.DataSet();
                    adapter.Fill(dataset);
                    data = dataset.Tables[0];
                }
            }
            return data;
        }

        public int Execute(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            if (connection.State == ConnectionState.Closed) { connection.Open(); }
            using (var command = new SqlCommand(query, connection))
            {
                if (par != null) { command.Parameters.AddRange(par); }
                return command.ExecuteNonQuery();
            }
        }

        public object getValue(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            if (connection.State == ConnectionState.Closed) { connection.Open(); }
            using (var command = new SqlCommand(query, connection))
            {
                if (CommandTimeOut > 0) { command.CommandTimeout = CommandTimeOut; }
                if (par != null) { command.Parameters.AddRange(par); }
                return command.ExecuteScalar();
            }
        }

        public string getValueField(string valueField)
        { if (string.IsNullOrEmpty(valueField)) { return ""; } return valueField.Replace("'", "''"); }

        public List<string> getColumnNames(string tableName)
        {
            var l = new List<string>();
            var tsql = $@"SELECT [COLUMN_NAME] FROM [INFORMATION_SCHEMA].[COLUMNS] WHERE [TABLE_CATALOG]='{getValueField(ConnectionString.InitialCatalog)}' AND [TABLE_NAME] = '{getValueField(tableName)}'";
            var dt = getDataTable(tsql);
            foreach (DataRow r in dt.Columns) l.Add(r[0].ToString());
            return l;
        }

        public List<DataColumn> getColumns(string tableName)
        {
            var l = new List<DataColumn>();
            var tsql = $"SELECT TOP 1 * FROM {getValueField(tableName)}";
            var dt = getDataTable(tsql);
            foreach (DataColumn c in dt.Columns) l.Add(c);
            return l;
        }

        public List<string> getAllTables(bool views = false)
        {
            var l = new List<string>();
            string type = "BASE TABLE";
            if (views) { type = "'table', 'view'"; }
            var tsql = $@"SELECT [TABLE_NAME] FROM [INFORMATION_SCHEMA].[TABLES] WHERE [TABLE_CATALOG]='{getValueField(ConnectionString.InitialCatalog)}' AND [TABLE_TYPE] = '{type}'";
            var dt = getDataTable(tsql);
            foreach (DataRow r in dt.Rows) l.Add($"{r[0]}");
            return l;
        }

        public int Update(string tableName, Dictionary<string, string> data, string where = "")
        {
            if (string.IsNullOrEmpty(tableName)) { return 0; }
            if (data.Count == 0) { return 0; }
            var tsql = "";
            var fields = new List<string>();
            var par = new List<SqlParameter>();
            if (string.IsNullOrEmpty(where))
            {
                /* Addnew */
                var parV = new List<string>();
                foreach (var v in data) { fields.Add($"{v.Key}"); parV.Add($"@{v.Key}"); par.Add(new SqlParameter($"@{v.Key}", v.Value)); }
                tsql = $"INSERT INTO {tableName} ({string.Join(",", fields)}) VALUES ({string.Join(", ", parV)});";
            }
            else
            {
                /* Update */
                foreach (var v in data) { fields.Add($"{v.Key} = @{v.Key}"); par.Add(new SqlParameter($"@{v.Key}", v.Value)); }
                where = where.Trim(); if (Regex.IsMatch(where, "^where", RegexOptions.IgnoreCase)) { where = where.Substring(5).Trim(); }
                tsql = $"UPDATE {tableName} SET {string.Join(",", fields)} WHERE {where}";
            }
            return Execute(tsql, par.ToArray());
        }

        public void Close()
        {
            if (connection.State != ConnectionState.Closed)
            {
                try
                {
                    if (connection.State == ConnectionState.Executing || connection.State == ConnectionState.Fetching)
                    {
                        foreach (SqlCommand command in connection.GetSchema("Commands").Rows) { command.Cancel(); }
                    }
                    connection.Close();
                }
                catch { }
            }
        }

        public void Dispose()
        {
            Close();
            connection.Dispose();
        }

        public Dictionary<string, object> getItem(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new Dictionary<string, object>(); }
            var result = new Dictionary<string, object>();
            foreach (DataColumn c in dt.Columns) { result.Add(c.ColumnName, dt.Rows[0][c.ColumnName]); }
            return result;
        }

        public List<string> getListValue(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new List<string>(); }
            var result = new List<string>();
            foreach (DataRow r in dt.Rows) { result.Add($"{r[0]}"); }
            return result;
        }

        public List<string> getListValueItem(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new List<string>(); }
            var result = new List<string>();
            foreach (DataColumn c in dt.Columns) { result.Add($"{dt.Rows[0][c.ColumnName]}"); }
            return result;
        }

        public Dictionary<string, object> getKeyValue(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new Dictionary<string, object>(); }
            if (dt.Columns.Count < 2) { return new Dictionary<string, object>(); }
            var result = new Dictionary<string, object>();
            foreach (DataRow r in dt.Rows) { result.Add($"{r[0]}", r[1]); }
            return result;
        }

        public List<KeyValuePair<string, string>> getListKeyValue(string query, object parameters = null)
        {
            var par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            var result = new List<KeyValuePair<string, string>>();
            if (dt.Rows.Count == 0) { return result; }
            if (dt.Columns.Count < 2) { return result; }
            foreach (DataRow r in dt.Rows) { result.Add(new KeyValuePair<string, string>($"{r[0]}", $"{r[1]}")); }
            return result;
        }

        public void backup(string pathFile)
        {
            Execute($"BACKUP DATABASE [{ConnectionString.InitialCatalog}] TO DISK = N'{pathFile}' WITH NOFORMAT, NOINIT, NAME = N'{ConnectionString.InitialCatalog} - Full Backup', SKIP, NOREWIND, NOUNLOAD, STATS = 10");
        }

        public void ExportDataToSqlFile(string outputFilePath, string version = "")
        {
            if (string.IsNullOrEmpty(version)) { version = DateTime.Now.ToString("yyyyMMdd"); }
            var CurrentCultureInfoName = Thread.CurrentThread.CurrentUICulture.Name;
            if (CurrentCultureInfoName != "en-US")
            {
                CultureInfo culture = CultureInfo.CreateSpecificCulture("en-US");
                Thread.CurrentThread.CurrentCulture = culture;
                Thread.CurrentThread.CurrentUICulture = culture;
            }
            try
            {
                /* Lấy danh sách các bảng trong cơ sở dữ liệu */
                var tables = getAllTables();
                int pageSize = 500;
                if (connection.State == ConnectionState.Closed) { connection.Open(); }
                using (StreamWriter writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                {
                    writer.WriteLine("--" + typeof(dbSQLServer).Namespace + " v" + version);
                    foreach (var tableName in tables)
                    {
                        var fields = new List<string>();
                        var tsqls = new List<string>();
                        var joinFields = "";
                        var tsql = "";
                        /* Với mỗi bảng, tạo một truy vấn SQL để tạo bảng và điền dữ liệu vào tập tin .sql */
                        SqlCommand dataCommand = new SqlCommand($"SELECT * FROM {tableName}", connection);
                        SqlDataReader dataReader = dataCommand.ExecuteReader();
                        writer.WriteLine($"TRUNCATE TABLE [{tableName}];");
                        writer.WriteLine("GO");
                        while (dataReader.Read())
                        {
                            if (joinFields == "")
                            {
                                /* Tạo danh sách trường import */
                                for (int i = 0; i < dataReader.FieldCount; i++) { fields.Add(dataReader.GetName(i)); }
                                joinFields = string.Join(",", fields);
                            }
                            var vals = new List<string>();
                            if (fields.Count == 0)
                            {
                                for (int i = 0; i < dataReader.FieldCount; i++)
                                {
                                    fields.Add(dataReader.GetName(i));
                                }
                                tsql = $"INSERT INTO [{tableName}] ({joinFields}) ({string.Join(",", fields)}) VALUES";
                            }

                            if (tsqls.Count >= pageSize)
                            {
                                writer.WriteLine($"{tsql} {string.Join(",", tsqls)};");
                                tsqls = new List<string>();
                            }
                            for (int i = 0; i < dataReader.FieldCount; i++)
                            {
                                if (dataReader.IsDBNull(i)) { vals.Add("NULL"); continue; }
                                if (dataReader.GetDataTypeName(i) == "datetime") { vals.Add($"'{dataReader.GetDateTime(i):yyyy-MM-dd HH:mm:ss}'"); continue; }
                                if (dataReader.GetDataTypeName(i) == "nvarchar" || dataReader.GetDataTypeName(i) == "nchar") { vals.Add($"N'{dataReader.GetString(i).Replace("'", "''")}'"); continue; }
                                vals.Add($"'{dataReader.GetValue(i)}'");
                            }
                            tsqls.Add($"({string.Join(",", vals)})");
                        }
                        if (tsqls.Count > 0)
                        {
                            writer.WriteLine($"{tsql} {string.Join(",", tsqls)};");
                            writer.WriteLine("GO");
                        }
                        dataReader.Close();
                        writer.Flush();
                    }
                }
                if (CurrentCultureInfoName != "en-US")
                {
                    CultureInfo culture = CultureInfo.CreateSpecificCulture(CurrentCultureInfoName);
                    Thread.CurrentThread.CurrentCulture = culture;
                    Thread.CurrentThread.CurrentUICulture = culture;
                }
            }
            catch (Exception ex)
            {
                if (CurrentCultureInfoName != "en-US")
                {
                    CultureInfo culture = CultureInfo.CreateSpecificCulture(CurrentCultureInfoName);
                    Thread.CurrentThread.CurrentCulture = culture;
                    Thread.CurrentThread.CurrentUICulture = culture;
                }
                throw new Exception(ex.Message);
            }
        }

        private string GetSqlType(string dataType)
        {
            switch (dataType)
            {
                case "int": return "INT";
                case "bigint": return "BIGINT";
                case "bit": return "BIT";
                case "decimal": return "DECIMAL";
                case "float": return "FLOAT";
                case "money": return "MONEY";
                case "numeric": return "NUMERIC";
                case "real": return "REAL";
                case "smallint": return "SMALLINT";
                case "tinyint": return "TINYINT";
                case "date": return "DATE";
                case "datetime": return "DATETIME";
                case "datetime2": return "DATETIME2";
                case "datetimeoffset": return "DATETIMEOFFSET";
                case "smalldatetime": return "SMALLDATETIME";
                case "time": return "TIME";
                case "char": return "CHAR";
                case "nchar": return "NCHAR";
                case "ntext": return "NTEXT";
                case "nvarchar": return "NVARCHAR(MAX)";
                case "text": return "TEXT";
                case "varchar": return "VARCHAR(MAX)";
                case "binary": return "BINARY";
                case "image": return "IMAGE";
                case "varbinary": return "VARBINARY(MAX)";
                case "uniqueidentifier": return "UNIQUEIDENTIFIER";
                default: throw new ArgumentException($"Data type {dataType} is not supported");
            }
        }

        public void ShrinkDatabase() => Execute($"DBCC SHRINKDATABASE ({ConnectionString.InitialCatalog}, 0);");

        public void Restore(string fileName, int CommandTimeout = 300)
        {
            string ext = Path.GetExtension(fileName).ToLower();
            if (ext == ".sql")
            {
                if (connection.State == ConnectionState.Closed) { connection.Open(); }
                int chunkSize = 1024 * 1024; // 1 MB
                int count = 0;
                using (StreamReader reader = new StreamReader(fileName, Encoding.UTF8))
                {
                    string line = $"{reader.ReadLine()}".Trim();
                    if (line.StartsWith($"--SoThongHanh ") == false) { reader.Close(); throw new Exception("Không phải tập tin sao lưu của phần mềm."); }
                    string sql = "";
                    // Đọc từng dòng trong tập tin .sql
                    while (!reader.EndOfStream)
                    {
                        line = reader.ReadLine().Trim();
                        if (line == "" || line == "GO")
                        {
                            if (!string.IsNullOrEmpty(sql))
                            {
                                SqlCommand command = new SqlCommand(sql, connection);
                                command.CommandTimeout = CommandTimeout;
                                command.ExecuteNonQuery();
                                count += sql.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
                                sql = "";
                            }
                            continue;
                        }
                        if (line.StartsWith("--"))
                        {
                            if (!string.IsNullOrEmpty(sql))
                            {
                                SqlCommand command = new SqlCommand(sql, connection);
                                command.CommandTimeout = CommandTimeout;
                                command.ExecuteNonQuery();
                                count += sql.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
                                sql = "";
                            }
                            continue;
                        }
                        sql += line + Environment.NewLine;
                        if (sql.Length > chunkSize)
                        {
                            SqlCommand command = new SqlCommand(sql, connection);
                            command.CommandTimeout = CommandTimeout;
                            command.ExecuteNonQuery();
                            count += sql.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
                            sql = "";
                        }
                    }
                    if (!string.IsNullOrEmpty(sql))
                    {
                        SqlCommand command = new SqlCommand(sql, connection);
                        command.CommandTimeout = CommandTimeout;
                        command.ExecuteNonQuery();
                        count += sql.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
                    }
                }
                return;
            }
            if (ext == ".bak")
            {
                var dataName = ConnectionString.InitialCatalog;
                Execute($"ALTER DATABASE [{dataName}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE " + Environment.NewLine +
                        $"RESTORE DATABASE [{dataName}] FROM DISK = N'{fileName} ' WITH FILE = 1, NORECOVERY, NOUNLOAD, REPLACE, STATS = 10" + Environment.NewLine +
                        $"ALTER DATABASE [{dataName}] SET MULTI_USER WITH ROLLBACK IMMEDIATE ");
            }
            throw new Exception($"Hiện phần mềm chưa hỗ trợ phụ hồi từ tập tin có kiểu '{ext}'");
        }

        public SqlDataReader getDataReader(string tsql)
        {
            if (connection.State == ConnectionState.Closed) { connection.Open(); }
            var command = new SqlCommand(tsql, connection);
            if (CommandTimeOut > 0) { command.CommandTimeout = CommandTimeOut; }
            return command.ExecuteReader();
        }
    }
}