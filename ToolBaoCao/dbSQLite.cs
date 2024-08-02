using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ToolBaoCao
{
    public class dbSQLite
    {
        public SQLiteConnectionStringBuilder connectString = new SQLiteConnectionStringBuilder();
        private SQLiteConnection connection = new SQLiteConnection();
        private string fileDataName = "";

        public long getTimestamp(DateTime time) => ((DateTimeOffset)time).ToUnixTimeSeconds();

        public dbSQLite(string pathOrConnectionString = "main.data", string password = "")
        {
            var cs = new SQLiteConnectionStringBuilder();
            try { cs = new SQLiteConnectionStringBuilder(pathOrConnectionString); } catch { }
            if (cs.DataSource == "")
            {
                if (string.IsNullOrEmpty(pathOrConnectionString)) { pathOrConnectionString = "main.db"; }
                cs.DataSource = pathOrConnectionString;
                if (string.IsNullOrEmpty(password) == false) { cs.Password = password; }
            }
            connectString = cs;
            connection.ConnectionString = cs.ConnectionString;
            fileDataName = Path.GetFileName(cs.DataSource);
        }

        public string getConnectionString(string databasePath = "main.data", string password = "")
        {
            var cs = new SQLiteConnectionStringBuilder();
            if (string.IsNullOrEmpty(databasePath)) { databasePath = "main.db"; }
            cs.DataSource = databasePath;
            if (string.IsNullOrEmpty(password) == false) { cs.Password = password; }
            return cs.ConnectionString;
        }

        public string getValueField(string valueField)
        { if (string.IsNullOrEmpty(valueField)) { return ""; } return valueField.Replace("'", "''"); }

        public string getPathDataFile() => connectString.DataSource;

        public void checkTableViewExists()
        { }

        public void Close()
        { if (connection.State != ConnectionState.Closed) { connection.Close(); } }

        private SQLiteParameter[] ConvertObjectToParameter(object parameters)
        {
            if (parameters == null) { return null; }
            if (parameters is KeyValuePair<string, string> obj1) { return new SQLiteParameter[] { new SQLiteParameter(obj1.Key, obj1.Value) }; }
            if (parameters is KeyValuePair<string, object> obj2) { return new SQLiteParameter[] { new SQLiteParameter(obj2.Key, obj2.Value) }; }
            if (parameters is SQLiteParameter obj5) { return new SQLiteParameter[] { obj5 }; }
            if (parameters is SQLiteParameter[] obj6) { return obj6; }
            if (parameters is Dictionary<string, string> obj3)
            {
                List<SQLiteParameter> par = new List<SQLiteParameter>();
                foreach (var v in obj3) { par.Add(new SQLiteParameter(v.Key, v.Value)); }
                return par.ToArray();
            }
            if (parameters is Dictionary<string, object> obj4)
            {
                List<SQLiteParameter> par = new List<SQLiteParameter>();
                foreach (var v in obj4) { par.Add(new SQLiteParameter(v.Key, v.Value)); }
                return par.ToArray();
            }
            throw new Exception($"Not support SQLiteParameter ${parameters}");
        }

        public DataTable getDataTable(string query, object parameters = null)
        {
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            DataTable data = new DataTable("DataTable");
            if (string.IsNullOrEmpty(query)) { return data; }
            var parstring = new List<string>();
            if (par != null) { foreach (var p in par) { parstring.Add($"{p.ParameterName}:{p.Value}"); } }
            var fileCache = AppHelper.GetPathFileCacheQuery($"{query} {string.Join(",", parstring)}");
            if (fileCache != "")
            {
                try
                {
                    if (File.Exists(fileCache))
                    {
                        data.ReadXml(fileCache);
                        return data;
                    }
                }
                catch { try { File.Delete(fileCache); } catch { } }
            }
            if (connection.State == ConnectionState.Closed) { connection.Open(); }
            using (var command = new SQLiteCommand(query, connection))
            {
                if (par != null) { command.Parameters.AddRange(par); }
                using (var adapter = new SQLiteDataAdapter(command))
                {
                    var dataset = new System.Data.DataSet();
                    adapter.Fill(dataset);
                    data = dataset.Tables[0];
                }
            }
            if (fileCache != "") { data.WriteXml(fileCache); }
            return data;
        }

        public int Execute(string query, object parameters = null)
        {
            var rs = 0;
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            if (connection.State == ConnectionState.Closed) { connection.Open(); }
            using (var command = new SQLiteCommand(query, connection))
            {
                if (par != null) { command.Parameters.AddRange(par); }
                rs = command.ExecuteNonQuery();
            }
            AppHelper.DeleteFileCacheQuery(query, fileDataName);
            return rs;
        }

        public object getValue(string query, object parameters = null)
        {
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            if (connection.State == ConnectionState.Closed) { connection.Open(); }
            using (var command = new SQLiteCommand(query, connection))
            {
                if (par != null) { command.Parameters.AddRange(par); }
                return command.ExecuteScalar();
            }
        }

        public List<DataColumn> getColumns(string tableName)
        {
            var l = new List<DataColumn>();
            var dt = getDataTable($"SELECT * FROM {tableName} limit 1");
            foreach (DataColumn c in dt.Columns) l.Add(c);
            return l;
        }

        public List<string> getAllTables(bool views = false)
        {
            var l = new List<string>();
            string type = "'table'";
            if (views) { type = "'table', 'view'"; }
            var dt = getDataTable($"SELECT [name] FROM [sqlite_master] WHERE type IN ({type}) AND name not like 'sqlite_%'");
            foreach (DataRow r in dt.Rows) l.Add($"{r[0]}");
            return l;
        }

        public void backup(string pathsave)
        {
            using (var destination = new SQLiteConnection($"Data Source={pathsave};Version=3;"))
            {
                if (connection.State == ConnectionState.Closed) { connection.Open(); }
                destination.Open();
                connection.BackupDatabase(destination, "main", "main", -1, null, 0);
            }
        }

        public void Restore(string fileName, int CommandTimeout = 300)
        {
            if (File.Exists(fileName) == false) { throw new Exception($"Tập tin {fileName} không tồn tại"); }
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
                                sql = sql.Replace("N'", "'");
                                SQLiteCommand command = new SQLiteCommand(sql, connection);
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
                                sql = sql.Replace("N'", "'");
                                SQLiteCommand command = new SQLiteCommand(sql, connection);
                                command.CommandTimeout = CommandTimeout;
                                command.ExecuteNonQuery();
                                count += sql.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
                                sql = "";
                            }
                            continue;
                        }
                        if (line.StartsWith("TRUNCATE TABLE")) { line = "DELETE FROM" + line.Substring("TRUNCATE TABLE".Length + 1); }
                        sql += line + Environment.NewLine;
                        if (sql.Length > chunkSize)
                        {
                            sql = sql.Replace("N'", "'");
                            SQLiteCommand command = new SQLiteCommand(sql, connection);
                            command.CommandTimeout = CommandTimeout;
                            command.ExecuteNonQuery();
                            count += sql.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
                            sql = "";
                        }
                    }
                    if (!string.IsNullOrEmpty(sql))
                    {
                        sql = sql.Replace("N'", "'");
                        SQLiteCommand command = new SQLiteCommand(sql, connection);
                        command.CommandTimeout = CommandTimeout;
                        command.ExecuteNonQuery();
                        count += sql.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
                    }
                }
                return;
            }
            if (ext == ".bak")
            {
                connection.Close();
                File.Copy(fileName, connectString.DataSource, overwrite: true);
                return;
            }
            throw new Exception($"Hiện phần mềm chưa hỗ trợ phụ hồi từ tập tin có kiểu '{ext}'");
        }

        public int Update(string tableName, Dictionary<string, string> data, string whereOrReplaceIgnore = "")
        {
            if (string.IsNullOrEmpty(tableName)) { return 0; }
            if (data.Count == 0) { return 0; }
            var tsql = ""; string tmp = ""; string where = whereOrReplaceIgnore;
            var fields = new List<string>();
            List<SQLiteParameter> par = new List<SQLiteParameter>();
            if (string.IsNullOrEmpty(where) || where.ToLower() == "replace" || where.ToLower() == "ignore")
            {
                /* Addnew */
                var parV = new List<string>();
                foreach (var v in data)
                {
                    tmp = Regex.Replace(v.Key, "[{}]", "");
                    fields.Add($"{tmp}");
                    parV.Add($"@{tmp}");
                    par.Add(new SQLiteParameter($"@{tmp}", v.Value));
                }
                switch (where.ToLower())
                {
                    case "replace":
                        tsql = $"INSERT OR REPLACE INTO {tableName} ({string.Join(", ", fields)}) VALUES ({string.Join(", ", parV)});";
                        break;

                    case "ignore":
                        tsql = $"INSERT OR IGNORE INTO {tableName} ({string.Join(", ", fields)}) VALUES ({string.Join(", ", parV)});";
                        break;

                    default:
                        tsql = $"INSERT INTO {tableName} ({string.Join(", ", fields)}) VALUES ({string.Join(", ", parV)});";
                        break;
                }
            }
            else
            {
                /* Update */
                foreach (var v in data)
                {
                    tmp = Regex.Replace(v.Key, "[{}]", "");
                    fields.Add($"{tmp} = @{tmp}");
                    par.Add(new SQLiteParameter($"@{tmp}", v.Value));
                }
                where = where.Trim(); if (Regex.IsMatch(where, "^where", RegexOptions.IgnoreCase)) { where = where.Substring(5).Trim(); }
                tsql = $"UPDATE {tableName} SET {string.Join(",", fields)} WHERE {where}";
            }
            return Execute(tsql, par.ToArray());
        }

        public int Insert(string tableName, DataTable data, string orRepalceIgnore = "", int packetSize = 1000)
        {
            if (string.IsNullOrEmpty(tableName)) { return 0; }
            if (data.Rows.Count == 0) { return 0; }
            var tsql = "";
            int rs = 0;
            var fields = new List<string>();
            foreach (DataColumn c in data.Columns) { fields.Add($"[{c.ColumnName}]"); }
            var tsqlInert = "";
            switch (orRepalceIgnore.ToLower())
            {
                case "replace":
                    tsqlInert = $"INSERT OR REPLACE INTO {tableName} ({string.Join(",", fields)}) VALUES ";
                    break;

                case "ignore":
                    tsqlInert = $"INSERT OR IGNORE INTO {tableName} ({string.Join(",", fields)}) VALUES ";
                    break;

                default:
                    tsqlInert = $"INSERT INTO {tableName} ({string.Join(",", fields)}) VALUES ";
                    break;
            }
            var values = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                if (values.Count >= packetSize)
                {
                    tsql = tsqlInert + string.Join(", ", values);
                    values = new List<string>();
                    rs = Execute(tsql);
                }
                var v = new List<string>();
                foreach (var f in fields)
                {
                    if (row[f] is DBNull) { v.Add("NULL"); continue; }
                    if (row.Table.Columns[f].DataType == typeof(DateTime)) { v.Add($"'{row[f]:yyyy-MM-dd H:m:s}'"); continue; }
                    v.Add("'" + $"{row[f]}".sqliteGetValueField() + "'");
                }
                values.Add($"({string.Join(",", v)})");
            }
            if (values.Count > 0)
            {
                tsql = tsqlInert + string.Join(", ", values);
                rs = Execute(tsql);
            }
            return rs;
        }

        public Dictionary<string, object> getItem(string query, object parameters = null)
        {
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new Dictionary<string, object>(); }
            var result = new Dictionary<string, object>();
            foreach (DataColumn c in dt.Columns) { result.Add(c.ColumnName, dt.Rows[0][c.ColumnName]); }
            return result;
        }

        public List<string> getListValue(string query, object parameters = null)
        {
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new List<string>(); }
            var result = new List<string>();
            foreach (DataRow r in dt.Rows) { result.Add($"{r[0]}"); }
            return result;
        }

        public List<string> getListValueItem(string query, object parameters = null)
        {
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new List<string>(); }
            var result = new List<string>();
            foreach (DataColumn c in dt.Columns) { result.Add($"{dt.Rows[0][c.ColumnName]}"); }
            return result;
        }

        public Dictionary<string, object> getKeyValue(string query, object parameters = null)
        {
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            if (dt.Rows.Count == 0) { return new Dictionary<string, object>(); }
            if (dt.Columns.Count < 2) { return new Dictionary<string, object>(); }
            var result = new Dictionary<string, object>();
            foreach (DataRow r in dt.Rows) { result.Add($"{r[0]}", r[1]); }
            return result;
        }

        public List<KeyValuePair<string, string>> getListKeyValue(string query, object parameters = null)
        {
            SQLiteParameter[] par = ConvertObjectToParameter(parameters);
            var dt = getDataTable(query, par);
            var result = new List<KeyValuePair<string, string>>();
            if (dt.Rows.Count == 0) { return result; }
            if (dt.Columns.Count < 2) { return result; }
            foreach (DataRow r in dt.Rows) { result.Add(new KeyValuePair<string, string>($"{r[0]}", $"{r[1]}")); }
            return result;
        }

        public void ExportDataToSqlFile(string outputFilePath, string version = "")
        {
            if (string.IsNullOrEmpty(version)) { version = DateTime.Now.ToString("yyyyMMdd"); }
            /* Lấy danh sách các bảng trong cơ sở dữ liệu */
            var tables = getAllTables();
            using (StreamWriter writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
            {
                writer.WriteLine("--" + typeof(dbSQLite).Namespace + " v" + version);
                if (connection.State == ConnectionState.Closed) { connection.Open(); }
                int pageSizes = 500; int index = 0; List<string> tsql = new List<string>();
                foreach (var tableName in tables)
                {
                    /* Với mỗi bảng, tạo một truy vấn SQL để tạo bảng và điền dữ liệu vào tập tin .sql */
                    SQLiteCommand dataCommand = new SQLiteCommand($"SELECT * FROM {tableName}", connection);
                    SQLiteDataReader dataReader = dataCommand.ExecuteReader();

                    writer.WriteLine($"TRUNCATE TABLE [{tableName}];");
                    writer.WriteLine("GO");

                    while (dataReader.Read())
                    {
                        var v = new List<string>() { "(" };
                        for (int i = 0; i < dataReader.FieldCount; i++)
                        {
                            if (dataReader.IsDBNull(i)) { v.Add("NULL"); }
                            else { v.Add($"'{dataReader.GetValue(i).ToString().Replace("'", "''")}'"); }
                            if (i < dataReader.FieldCount - 1) { v.Add(","); }
                        }
                        v.Add(")");
                        tsql.Add(string.Join("", v));
                        index++;
                        if (index >= pageSizes) { writer.WriteLine($"INSERT INTO [{tableName}] VALUES {string.Join(",", tsql)};"); index = 0; tsql = new List<string>(); }
                    }
                    if (index > 0) { writer.WriteLine($"INSERT INTO [{tableName}] VALUES {string.Join(",", tsql)};"); }
                    writer.WriteLine("GO");
                    dataReader.Close();
                    writer.Flush();
                }
            }
        }
    }
}