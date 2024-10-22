using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace ToolBaoCao
{
    public class ItemTask
    {
        public ItemTask(string id, string name, string actionName = "", string args = "", long timeStart = 0)
        {
            ID = id;
            NameTask = name;
            ActionName = actionName;
            Args = args;
            if (timeStart == 0) { TimeStart = DateTime.Now; }
            else { TimeStart = timeStart.toDateTime(); }
        }

        public string ID { get; set; } = "";
        public string NameTask { get; set; } = "";
        public string ActionName { get; set; } = "";
        public string Args { get; set; } = "";
        public long Running { get; set; } = 0;
        public DateTime TimeStart { get; set; } = DateTime.Now;
    }

    public class TaskManage
    {
        private ConcurrentDictionary<string, ItemTask> _threads = new ConcurrentDictionary<string, ItemTask>();
        private Timer _timer;
        private dbSQLite dbTask = new dbSQLite(Path.Combine(AppHelper.pathAppData, "task.db"));
        public string IDRunning = "";

        public TaskManage()
        {
            Load();
            _timer = new Timer(_ => Call(), null, TimeSpan.Zero, TimeSpan.FromMinutes(30));
        }

        public void Load()
        {
            dbTask.Execute("CREATE TABLE IF NOT EXISTS task(id text not null primary key, nametask text not null default '', actionname text not null default '', args text not null default '', running integer not null default 0, timestart integer not null);");
            var data = dbTask.getDataTable("SELECT * FROM task ORDER BY timestart");
            foreach (DataRow row in data.Rows)
            {
                var item = new ItemTask(row["id"].ToString(), row["nametask"].ToString(), $"{row["actionname"]}", $"{row["args"]}", long.Parse($"{row["timestart"]}"));
                Add(item, false);
            }
        }

        public void Add(ItemTask item, bool callRun = true)
        {
            item.Running = 0;
            if (_threads.TryAdd(item.ID, item))
            {
                var tsql = $"INSERT OR IGNORE INTO task(id, nametask, actionname, args, timestart) VALUES ('{item.ID}', '{item.NameTask.sqliteGetValueField()}', '{item.ActionName.sqliteGetValueField()}', '{item.Args.sqliteGetValueField()}', '{item.TimeStart.toTimestamp()}')";
                try { dbTask.Execute(tsql); }
                catch (Exception ex)
                {
                    AppHelper.saveError($"Task({item.ID} - {item.ActionName} - {item.Args}): {tsql}{Environment.NewLine}{ex.Message}");
                    throw new Exception(ex.getLineHTML());
                }
                if (callRun) { Call(); }
            }
        }

        public void Delete(string ID)
        {
            if (_threads.TryGetValue(ID, out var item))
            {
                _threads.TryRemove(ID, out _);
                dbTask.Execute($"DELETE FROM task WHERE id='{item.ID}';");
            }
            if (ID == IDRunning) { IDRunning = ""; }
            Call();
        }

        public void Call()
        {
            if (IDRunning != "")
            {
                var obj = _threads.Values.FirstOrDefault(p => p.ID == IDRunning);
                if (obj != null) { return; }
                AppHelper.saveError($"Không tìm thấy ID Task '{IDRunning}'");
                IDRunning = "";
            }
            var item = _threads.Values.FirstOrDefault();
            if (item != null)
            {
                IDRunning = item.ID;
                try
                {
                    switch (item.ActionName.ToLower())
                    {
                        case "controller.xml":
                            Thread t = new Thread(new ThreadStart(() =>
                            {
                                try
                                {
                                    /* Kiểm tra xem có trong danh sách XMLThread không? */
                                    var tmp = item.Args.Split('|');
                                    var dbXML = BuildDatabase.getDataXML(tmp[0]);
                                    tmp[0] = $"{dbXML.getValue($"SELECT time2 FROM xmlthread WHERE id='{tmp[1]}';")}";
                                    dbXML.Close();
                                    if (tmp[0] == "") { Delete(tmp[1]); return; }
                                    if (tmp[0] != "0") { Delete(tmp[1]); return; }
                                    XMLThread(item.Args);
                                    Delete(tmp[1]);
                                }
                                catch (Exception exT) { AppHelper.saveError($"Lỗi XMLThread({item.ID} - {item.ActionName} - {item.Args}): {exT.Message}"); }
                            }));
                            t.Start();
                            break;

                        default: AppHelper.saveError($"Không tìm thấy Task({item.ID} - {item.ActionName} - {item.Args})"); break;
                    }
                }
                catch (Exception ex) { AppHelper.saveError($"Task({item.ID} - {item.ActionName} - {item.Args}) Lỗi: {ex.Message}"); }
            }
        }

        public void setFinshThreadInAppStart()
        {
            Thread t = new Thread(new ThreadStart(() =>
            {
                var d = new DirectoryInfo(Path.Combine(AppHelper.pathAppData, "xml"));
                if ((d.Exists == false)) { d.Create(); return; }
                foreach (var f in d.GetFiles("*.db"))
                {
                    var db = new dbSQLite(f.FullName);
                    try
                    {
                        var tables = db.getAllTables();
                        if (tables.Contains("xmlthread"))
                        {
                            /* AND title <> 'Thread was being aborted.' */
                            db.Execute($"UPDATE xmlthread SET title='Lỗi do do hệ thống bị ngắt đột ngột', time2={DateTime.Now.toTimestamp()} WHERE time2=0;");
                        }
                    }
                    catch { }
                    db.Close();
                }
            }));
            t.Start();
        }

        /// <summary>
        /// idThread = {MaTinh}|{ID table XML}
        /// </summary>
        /// <param name="idThread">{MaTinh}|{ID table XML}</param>
        public void XMLThread(string idThread)
        {
            string folderTemp = "", folderSave = "", id = "", matinh = "";
            var dbXML = BuildDatabase.getDataXML(matinh);
            try
            {
                var objs = idThread.Split('|');
                if (objs.Length != 2) { throw new Exception($"Tham số không đúng idThread XML '{idThread}'"); }
                id = objs[1];
                matinh = objs[0];
                folderTemp = Path.Combine(AppHelper.pathTemp, "xml", $"t{matinh}");
                folderSave = Path.Combine(AppHelper.pathAppData, "xml", $"t{matinh}");
                if (Directory.Exists(folderSave) == false) { Directory.CreateDirectory(folderSave); }
                dbXML = BuildDatabase.getDataXML(matinh);
                var data = dbXML.getDataTable($"SELECT * FROM xmlthread WHERE id='{id}'");
                if (data.Rows.Count == 0) { throw new Exception($"Thread XML có id '{id}' không tồn tại hoặc đã bị xoá khỏi hệ thống"); }
                var item = new Dictionary<string, string>();
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    item[data.Columns[i].ColumnName] = data.Rows[0][i].ToString();
                }
                var lfile = item["args"].Split('|').ToList();
                int ij = 0;
                var XMLdb = new dbSQLite(Path.Combine(AppHelper.pathAppData, "xml", $"t{matinh}", $"xml{id}.db"));
                XMLdb.Execute("DROP TABLE IF EXISTS xml123; DROP TABLE IF EXISTS xml7980a; DROP TABLE IF EXISTS xmlthread;");
                XMLdb.createTableXMLThread();
                foreach (string f in lfile)
                {
                    dbXML.Execute($"UPDATE xmlthread SET title = 'Đang thao tác tại {f} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                    ij++;
                    var fileName = AppHelper.pathApp + f;
                    if (System.IO.File.Exists(fileName) == false) { throw new Exception($"XMLThread '{id}' có tập tin '{f}' không tồn tại trong hệ thống"); }
                    var ext = Path.GetExtension(fileName);
                    if (ext == ".zip")
                    {
                        using (ZipArchive archive = ZipFile.OpenRead(fileName))
                        {
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                ij++;
                                if (entry.FullName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
                                {
                                    dbXML.Execute($"UPDATE xmlthread SET title = 'Đang giải nén {entry.FullName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                                    var fdbForm = Path.Combine(folderTemp, $"xml{id}_zip{ij}.db");
                                    entry.ExtractToFile(fdbForm, overwrite: true);
                                    var dbFrom = new dbSQLite(fdbForm);
                                    /* Kiểm tra có đúng cấu trúc dữ liệu không? */
                                    dbXML.Execute($"UPDATE xmlthread SET title = 'Kiểm tra cấu trúc {entry.FullName} ({DateTime.Now:dd/MM/yyyy HH:mm})', time2 = 0 WHERE id='{id}'");
                                    var tables = dbFrom.getAllTables();
                                    var tsql = "SELECT MIN(KY_QT) AS X1, MAX(KY_QT) AS X2 FROM ";
                                    var tableName = "xml123";
                                    if (tables.Contains("xml7980a")) { tableName = "xml7980a"; }
                                    else { if (tables.Contains("bhyt7980a")) { tableName = "bhyt7980a"; } }
                                    data = dbFrom.getDataTable($"{tsql}{tableName} LIMIT 1");
                                    if (data.Rows.Count == 0)
                                    {
                                        dbFrom.Close();
                                        throw new Exception($"XMLThread '{id}' có tập tin '{f}' không có dữ liệu");
                                    }
                                    /* Chuyển dữ liệu */
                                    XMLCopyTable(XMLdb, dbFrom, dbXML, id);
                                    dbFrom.Close();
                                    /* Xoá đi sau khi sao chép song */
                                    try { System.IO.File.Delete(fdbForm); } catch { }
                                }
                            }
                        }
                        continue;
                    }
                    if (ext == ".db")
                    {
                        var dbFrom = new dbSQLite(fileName);
                        /* Kiểm tra có đúng cấu trúc dữ liệu không? */
                        dbXML.Execute($"UPDATE xmlthread SET title = 'Kiểm tra cấu trúc {f} ({DateTime.Now:dd/MM/yyyy HH:mm})', time2 = 0 WHERE id='{id}'");
                        var tables = dbFrom.getAllTables();
                        var tsql = "SELECT MIN(KY_QT) AS X1, MAX(KY_QT) AS X2 FROM ";
                        if (tables.Contains("xml7980a")) { tsql += "xml7980a"; }
                        else
                        {
                            if (tables.Contains("bhyt7980a")) { tsql += "bhyt7980a"; }
                            else { tsql += "xml123"; }
                        }
                        data = dbFrom.getDataTable(tsql + " LIMIT 1");
                        if (data.Rows.Count == 0)
                        {
                            dbFrom.Close();
                            throw new Exception($"XMLThread '{id}' có tập tin '{f}' không có dữ liệu");
                        }
                        XMLCopyTable(XMLdb, dbFrom, dbXML, id);
                        dbFrom.Close();
                        continue;
                    }
                }
                dbXML.Execute($"UPDATE xmlthread SET title = 'Hoàn thành', pageindex = 1, time2='{DateTime.Now.toTimestamp()}' WHERE id='{id}'");
                data = dbXML.getDataTable($"SELECT * FROM xmlthread WHERE id='{id}'");
                XMLdb.Insert("xmlthread", data, "replace");
                XMLdb.Close();
            }
            catch (Exception ex)
            {
                ex.saveError();
                dbXML.Execute($"UPDATE xmlthread SET title = '{ex.Message.sqliteGetValueField()}', time2='{DateTime.Now.toTimestamp()}' WHERE id='{id}'");
            }
            dbXML.Close();
        }

        private string RemoveColumns(string tsql, HashSet<string> columnsToRemove)
        {
            tsql = tsql.Replace(Environment.NewLine, "");
            foreach (var column in columnsToRemove)
            {
                var p = $"[\"']?{column}[\"']?";
                tsql = Regex.Replace(tsql, $@"{p}\s+\w+\s*(,)?", "", RegexOptions.IgnoreCase);
            }
            return Regex.Replace(tsql, @"\s+", " ");
        }

        private void XMLCopyTable(dbSQLite dbTo, dbSQLite dbFrom, dbSQLite dbXML, string id)
        {
            var tablesTo = dbTo.getAllTables();
            var tablesFrom = dbFrom.getAllTables();
            var fileName = Path.GetFileName(dbFrom.getPathDataFile());
            var tmp = "";
            int batchSize = 1000; double rowCopyed = 0;
            var colsRemove = new HashSet<string> { "TEN_TINH", "TEN_CSKCB", "COSOKCB_ID", "MA_TINH_THE", "T_VUOTTRAN" };
            var colsMD5 = new List<string>() { "MA_THE", "NGAY_SINH", "HO_TEN", "DIA_CHI" };
            /* Tạo bảo nếu chưa có */
            if (tablesTo.Contains("xml123") == false)
            {
                if (tablesFrom.Contains("xml123"))
                {
                    tmp = $"{dbFrom.getValue("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'xml123'")}";
                    if (Regex.IsMatch(tmp, "primary key", RegexOptions.IgnoreCase) == false)
                    {
                        tmp = tmp.Replace(")", ", PRIMARY KEY(ID))");
                    }
                    if (Regex.IsMatch(tmp, "IF NOT EXISTS", RegexOptions.IgnoreCase) == false)
                    {
                        tmp = Regex.Replace(tmp, "CREATE TABLE ", "CREATE TABLE IF NOT EXISTS ", RegexOptions.IgnoreCase);
                    }
                    tmp = RemoveColumns(tmp, colsRemove);
                    AppHelper.saveError(tmp);
                    dbTo.Execute(tmp);
                    dbTo.Execute("CREATE INDEX xml123_index1 ON xml123(MA_TINH,KY_QT,MA_CHA,MA_CSKCB);");
                }
            }
            tmp = "";
            if (tablesTo.Contains("xml7980a") == false)
            {
                if (tablesFrom.Contains("xml7980a"))
                {
                    tmp = $"{dbFrom.getValue("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'xml7980a'")}";
                }
                else
                {
                    if (tablesFrom.Contains("bhyt7980a"))
                    {
                        tmp = $"{dbFrom.getValue("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'bhyt7980a'")}";
                        tmp = tmp.Replace("bhyt7980a", "xml7980a");
                    }
                }
                if (tmp != "")
                {
                    if (Regex.IsMatch(tmp, "primary key", RegexOptions.IgnoreCase) == false)
                    {
                        tmp = tmp.Replace(")", ", PRIMARY KEY(ID))");
                    }
                    if (Regex.IsMatch(tmp, "IF NOT EXISTS", RegexOptions.IgnoreCase) == false)
                    {
                        tmp = Regex.Replace(tmp, "CREATE TABLE ", "CREATE TABLE IF NOT EXISTS ", RegexOptions.IgnoreCase);
                    }
                    tmp = RemoveColumns(tmp, colsRemove);
                    AppHelper.saveError(tmp);
                    dbTo.Execute(tmp);
                    dbTo.Execute("CREATE INDEX xml7980a_index1 ON xml7980a(MA_TINH,KY_QT,MA_CSKCB);");
                }
            }
            string totalRow = "", tableName = "xml123";
            if (tablesFrom.Contains(tableName))
            {
                totalRow = $"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName}")}".FormatCultureVN();
                if (totalRow == "0") { throw new Exception($"{fileName}: Không có dữ liệu xml123"); }
                dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}(0/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                rowCopyed = 0;
                var lfield = dbTo.getColumns(tableName).Select(p => p.ColumnName).ToList();
                var data = dbFrom.getDataTable($"SELECT {string.Join(", ", lfield)} FROM {tableName} LIMIT 1");
                data.Rows.RemoveAt(0);
                var md5EncyptCols = colsMD5;
                foreach(var v in md5EncyptCols) { if (lfield.Contains(v) == false) { md5EncyptCols.Remove(v); } }
                var reader = dbFrom.getDataReader($"SELECT {string.Join(", ", lfield)} FROM {tableName}");
                while (reader.Read())
                {
                    if (data.Rows.Count >= batchSize)
                    {
                        /* Copy AND ignore */
                        dbTo.Insert(tableName, data, "IGNORE", batchSize);
                        rowCopyed += data.Rows.Count;
                        dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                        data.Rows.Clear();
                    }
                    DataRow dr = data.NewRow();
                    foreach (DataColumn c in data.Columns) { dr[c.ColumnName] = reader[c.ColumnName]; }
                    foreach(var c in md5EncyptCols) { dr[c] = $"{dr[c]}".MD5Encrypt(); }
                    data.Rows.Add(dr);
                }
                if (data.Rows.Count > 0)
                {
                    /* Copy AND ignore */
                    dbTo.Insert(tableName, data, "IGNORE", batchSize);
                    rowCopyed += data.Rows.Count;
                    dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                }
                reader.Close();
            }
            tableName = "xml7980a";
            if (tablesFrom.Contains(tableName))
            {
                var lfield = dbTo.getColumns(tableName).Select(p => p.ColumnName).ToList();
                totalRow = $"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName}")}".FormatCultureVN();
                if (totalRow == "0") { throw new Exception($"{fileName}: Không có dữ liệu xml123"); }
                dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}(0/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                rowCopyed = 0;
                var data = dbFrom.getDataTable($"SELECT {string.Join(", ", lfield)} FROM {tableName} LIMIT 1");
                data.Rows.RemoveAt(0);
                var md5EncyptCols = colsMD5;
                foreach (var v in md5EncyptCols) { if (lfield.Contains(v) == false) { md5EncyptCols.Remove(v); } }
                var reader = dbFrom.getDataReader($"SELECT {string.Join(", ", lfield)} FROM {tableName}");
                while (reader.Read())
                {
                    if (data.Rows.Count >= batchSize)
                    {
                        /* Copy AND ignore */
                        dbTo.Insert(tableName, data, "IGNORE", batchSize);
                        rowCopyed += data.Rows.Count;
                        dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                        data.Rows.Clear();
                    }
                    DataRow dr = data.NewRow();
                    foreach (DataColumn c in data.Columns) { dr[c.ColumnName] = reader[c.ColumnName]; }
                    foreach(var c in md5EncyptCols) { dr[c] = $"{dr[c]}".MD5Encrypt(); }
                    data.Rows.Add(dr);
                }
                if (data.Rows.Count > 0)
                {
                    /* Copy AND ignore */
                    dbTo.Insert(tableName, data, "IGNORE", batchSize);
                    rowCopyed += data.Rows.Count;
                    dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                }
                reader.Close();
            }
            tableName = "bhyt7980a";
            if (tablesFrom.Contains(tableName))
            {
                var lfield = dbTo.getColumns("xml7980a").Select(p => p.ColumnName).ToList();
                totalRow = $"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName}")}".FormatCultureVN();
                if (totalRow == "0") { throw new Exception($"{fileName}: Không có dữ liệu xml123"); }
                dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}(0/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                rowCopyed = 0;
                var data = dbFrom.getDataTable($"SELECT {string.Join(", ", lfield)} FROM {tableName} LIMIT 1");
                data.Rows.RemoveAt(0);
                var md5EncyptCols = colsMD5;
                foreach (var v in md5EncyptCols) { if (lfield.Contains(v) == false) { md5EncyptCols.Remove(v); } }
                var reader = dbFrom.getDataReader($"SELECT {string.Join(", ", lfield)} FROM {tableName}");
                while (reader.Read())
                {
                    if (data.Rows.Count >= batchSize)
                    {
                        /* Copy AND ignore */
                        dbTo.Insert("xml7980a", data, "IGNORE", batchSize);
                        rowCopyed += data.Rows.Count;
                        dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                        data.Rows.Clear();
                    }
                    DataRow dr = data.NewRow();
                    foreach (DataColumn c in data.Columns) { dr[c.ColumnName] = reader[c.ColumnName]; }
                    foreach(var c in md5EncyptCols) { dr[c] = $"{dr[c]}".MD5Encrypt(); }
                    data.Rows.Add(dr);
                }
                if (data.Rows.Count > 0)
                {
                    /* Copy AND ignore */
                    dbTo.Insert("xml7980a", data, "IGNORE", batchSize);
                    rowCopyed += data.Rows.Count;
                    dbXML.Execute($"UPDATE xmlthread SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                }
                reader.Close();
            }
            dbFrom.Close(); dbTo.Close(); dbXML.Close();
        }
    }
}