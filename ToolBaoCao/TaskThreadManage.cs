﻿using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Common;
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
        private dbSQLite dbTask = new dbSQLite(Path.Combine(AppHelper.pathAppData, "task.db"));
        public string IDRunning = "";

        public TaskManage()
        {
            Load();
            Thread t = new Thread(new ThreadStart(() =>
            {
                while (true)
                {
                    /* Mặc định 30 phút = 30 * 60 * 1000 */
                    int i = 1800000;
                    try
                    {
                        var tmp = AppHelper.getConfig("threadload.sleep", "666");
                        if (Regex.IsMatch(tmp, @"^\d+$") == false) { tmp = i.ToString(); }
                        i = int.Parse(tmp);
                        if (i < 600) { i = 600; }
                        i = i * 1000;
                    }
                    catch { }
                    Call();
                    Thread.Sleep(i);
                }
            }));
            t.Start();
        }

        public Dictionary<string, ItemTask> GetData() => _threads.ToDictionary(v => v.Key, v => v.Value);

        private void findThread()
        {
            /* XML */
            var d = new DirectoryInfo(Path.Combine(AppHelper.pathAppData, "xml"));
            if ((d.Exists == false)) { d.Create(); }
            else
            {
                foreach (var f in d.GetFiles("*.db"))
                {
                    var db = new dbSQLite(f.FullName);
                    try
                    {
                        if (db.tableExist("xmlthread") == false) { continue; }
                        var dt = db.getDataTable("SELECT * FROM xmlthread WHERE title='Thread was being aborted.';");
                        foreach (DataRow row in dt.Rows)
                        {
                            var itemTask = new ItemTask(row["id"].ToString(), $"Controller.XML.{row["id"]}", "Controller.XML", $"{row["matinh"]}|{row["id"]}", long.Parse(row["time1"].ToString()));
                            Add(itemTask, false);
                        }
                        db.Execute($"UPDATE xmlthread SET time2 = 0, args2 = args2 || '; Recall Thread {DateTime.Now:HH:mm:ss}' WHERE title='Thread was being aborted.';");
                    }
                    catch { }
                    db.Close();
                }
            }
        }

        private void Load()
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
            if (item == null) { findThread(); return; }
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

        /// <summary>
        /// idThread = {MaTinh}|{ID table XML}
        /// </summary>
        /// <param name="idThread">{MaTinh}|{ID table XML}</param>
        private void XMLThread(string idThread)
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
                var lfileTarget = item["name"].Split(',').ToList();
                int ij = 0; int indexFileTarget = int.Parse(item["pageindex"]);
                var dbTo = new dbSQLite(Path.Combine(AppHelper.pathAppData, "xml", $"t{matinh}", $"xml{id}.db"));
                dbTo.createTableXMLThread();
                data = dbXML.getDataTable($"SELECT * FROM xmlthread WHERE id='{id}'");
                dbTo.Insert("xmlthread", data, "ignore");
                for (; indexFileTarget < lfile.Count; indexFileTarget++)
                {
                    string f = lfile[indexFileTarget];
                    dbXML.Execute($"UPDATE xmlthread SET title = 'Đang thao tác tại {f} ({DateTime.Now:HH:mm:ss})', pageindex={indexFileTarget} WHERE id='{id}'");
                    ij++;
                    var fileName = AppHelper.pathApp + f;
                    if (System.IO.File.Exists(fileName) == false) { throw new Exception($"XMLThread '{id}' có tập tin '{f}' không tồn tại trong hệ thống"); }
                    var ext = Path.GetExtension(fileName);
                    if (ext == ".zip")
                    {
                        using (ZipArchive archive = ZipFile.OpenRead(fileName))
                        {
                            int indexDBZip = 0;
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                if (entry.FullName.EndsWith(".db", StringComparison.OrdinalIgnoreCase) == false) { continue; }
                                ij++; indexDBZip++;
                                dbXML.Execute($"UPDATE xmlthread SET title = 'Đang giải nén {entry.FullName} ({DateTime.Now:HH:mm:ss})' WHERE id='{id}'");
                                var fdbForm = Path.Combine(folderTemp, $"xml{id}_zip_{indexFileTarget}_{indexDBZip}.db");
                                bool extract = true;
                                var fi = new FileInfo(fdbForm);
                                if (fi.Exists) { if (fi.Length == entry.Length) { extract = false; } }
                                if (extract) { entry.ExtractToFile(fdbForm, overwrite: true); }
                                var dbFrom = new dbSQLite(fdbForm);
                                /* Chuyển dữ liệu */
                                XMLCopyTable(dbTo, dbFrom, dbXML, id, lfileTarget[indexFileTarget] + $"[{indexDBZip}]");
                                dbFrom.Close();
                                /* Xoá đi sau khi sao chép song */
                                try { System.IO.File.Delete(fdbForm); } catch { }
                            }
                        }
                    }
                    else if (ext == ".rar")
                    {
                        using (var archive = RarArchive.Open(fileName))
                        {
                            int indexDBZip = 0;
                            foreach (var entry in archive.Entries)
                            {
                                if (entry.Key.EndsWith(".db", StringComparison.OrdinalIgnoreCase) == false) { continue; }
                                ij++; indexDBZip++;
                                dbXML.Execute($"UPDATE xmlthread SET title = 'Đang giải nén {entry.Key} ({DateTime.Now:HH:mm:ss})' WHERE id='{id}'");
                                var fdbForm = Path.Combine(folderTemp, $"xml{id}_zip_{indexFileTarget}_{indexDBZip}.db");
                                bool extract = true;
                                var fi = new FileInfo(fdbForm);
                                if (fi.Exists) { if (fi.Length == entry.Size) { extract = false; } }
                                if (extract) { entry.WriteToFile(fdbForm, new ExtractionOptions { ExtractFullPath = true, Overwrite = true }); }
                                var dbFrom = new dbSQLite(fdbForm);
                                /* Chuyển dữ liệu */
                                XMLCopyTable(dbTo, dbFrom, dbXML, id, lfileTarget[indexFileTarget] + $"[{indexDBZip}]");
                                dbFrom.Close();
                                /* Xoá đi sau khi sao chép song */
                                try { System.IO.File.Delete(fdbForm); } catch { }
                            }
                        }
                    }
                    else if (ext == ".db")
                    {
                        var dbFrom = new dbSQLite(fileName);
                        XMLCopyTable(dbTo, dbFrom, dbXML, id, lfileTarget[indexFileTarget]);
                        dbFrom.Close();
                        continue;
                    }
                }
                dbXML.Execute($"UPDATE xmlthread SET title = 'Hoàn thành', time2='{DateTime.Now.toTimestamp()}' WHERE id='{id}'");
                data = dbXML.getDataTable($"SELECT * FROM xmlthread WHERE id='{id}'");
                dbTo.Insert("xmlthread", data, "replace");
                dbTo.Close();
                /* Xoá hết các tập tin tạm để giải phóng dung lượng */
                /*
                var d = new DirectoryInfo(folderTemp);
                foreach (var f in d.GetFiles($"xml{id}*.*")) { try { f.Delete(); } catch { } }
                */
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

        private void XMLCopyTable2(dbSQLite dbTo, dbSQLite dbFrom, dbSQLite dbXML, string idThread, string tableName, string fileName, int batchSize, List<string> colsMD5)
        {
            string totalRow = $"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName}")}".FormatCultureVN();
            string tmp = "";
            if (totalRow == "0")
            {
                tmp = $"UPDATE xmlthread SET args2 = args2 || '; {fileName.sqliteGetValueField()} không có dữ liệu {tableName}' WHERE id='{idThread}'";
                dbXML.Execute(tmp);
                return;
            }
            tmp = $"{dbTo.getValue($"SELECT pageindex FROM xmlthread WHERE id = '{idThread}' AND args2 = '{tableName}'")}";
            if (tmp == "") { tmp = "0"; }
            long idCopy = long.Parse(tmp);
            string tableTo = tableName == "xml123" ? "xml123" : "xml7980a";
            /* Chuyển dữ liệu */
            int soTrung = 0; int soExecute = 0;
            long rowCopyed = long.Parse($"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName} WHERE ID <= {idCopy}")}");
            dbXML.Execute($"UPDATE xmlthread SET title = '{fileName}: đã chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) {DateTime.Now:HH:mm:ss}' WHERE id='{idThread}'");
            var lfield = dbTo.getColumns(tableName).Select(p => p.ColumnName).ToList();
            var data = dbFrom.getDataTable($"SELECT {string.Join(", ", lfield)} FROM {tableName} LIMIT 1");
            data.Rows.RemoveAt(0);
            var md5EncyptCols = new List<string>();
            foreach (var v in colsMD5) { if (lfield.Contains(v)) { md5EncyptCols.Add(v); } }
            var reader = dbFrom.getDataReader($"SELECT {string.Join(", ", lfield)} FROM {tableName} WHERE ID > {idCopy} ORDER BY ID");
            try
            {
                while (reader.Read())
                {
                    if (data.Rows.Count >= batchSize)
                    {
                        /* Copy AND ignore */
                        soExecute = dbTo.Insert(tableTo, data, "IGNORE", batchSize);
                        rowCopyed += data.Rows.Count; soTrung = data.Rows.Count - soExecute;
                        dbXML.Execute($"UPDATE xmlthread SET title = '{fileName} đã chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow} Trùng {soTrung.FormatCultureVN()}) {DateTime.Now:HH:mm:ss}' WHERE id='{idThread}'");
                        dbTo.Execute($"UPDATE xmlthread SET args2='{tableName}', pageindex={data.Rows[data.Rows.Count - 1]["ID"]} WHERE id= '{idThread}'");
                        data.Rows.Clear();
                    }
                    DataRow dr = data.NewRow();
                    foreach (DataColumn c in data.Columns) { dr[c.ColumnName] = reader[c.ColumnName]; }
                    foreach (var c in md5EncyptCols) { dr[c] = $"{dr[c]}".MD5Encrypt(); }
                    data.Rows.Add(dr);
                }
                if (data.Rows.Count > 0)
                {
                    /* Copy AND ignore */
                    soExecute = dbTo.Insert(tableTo, data, "IGNORE", batchSize);
                    rowCopyed += data.Rows.Count; soTrung = data.Rows.Count - soExecute;
                    dbXML.Execute($"UPDATE xmlthread SET title = '{fileName} đã chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow} Trùng {soTrung.FormatCultureVN()}) {DateTime.Now:HH:mm:ss}' WHERE id='{idThread}'");
                }
                dbTo.Execute($"UPDATE xmlthread SET args2='', pageindex=0 WHERE id= '{idThread}'");
            }
            catch (Exception ex2)
            {
                dbXML.Execute($"UPDATE xmlthread SET args2 = args2 || '; {fileName.sqliteGetValueField()} {tableName}: {ex2.Message.sqliteGetValueField()}' WHERE id='{idThread}'");
                ex2.saveError();
            }
            reader.Close();
            dbXML.Execute($"UPDATE xmlthread SET args2 = args2 || '; {fileName.sqliteGetValueField()} {tableName} đã chép {rowCopyed.FormatCultureVN()} Trùng {soTrung.FormatCultureVN()}' WHERE id='{idThread}'");
        }

        private void XMLCopyTable(dbSQLite dbTo, dbSQLite dbFrom, dbSQLite dbXML, string id, string nameFile)
        {
            var tablesTo = dbTo.getAllTables();
            var tablesFrom = dbFrom.getAllTables();
            var tablesCopy = new List<string>();
            if (tablesFrom.Contains("xml123")) { tablesCopy.Add("xml123"); }
            if (tablesFrom.Contains("xml7980a")) { tablesCopy.Add("xml7980a"); }
            if (tablesFrom.Contains("bhyt7980a")) { tablesCopy.Add("bhyt7980a"); }
            var fileName = $"{nameFile}: ";
            if (tablesCopy.Count == 0)
            {
                dbXML.Execute($"UPDATE xmlthread SET args2 = args2 || '; {fileName.sqliteGetValueField()}: không có dữ liệu.' WHERE id='{id}'");
            }
            var tmp = "";
            int batchSize = 1000;
            var colsRemove = new HashSet<string> { "TEN_TINH", "TEN_CSKCB", "COSOKCB_ID", "MA_TINH_THE", "T_VUOTTRAN" };
            var colsMD5 = new List<string>() { "MA_THE", "NGAY_SINH", "HO_TEN", "DIA_CHI" };
            /* Tạo bảng nếu chưa có */
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
                    dbTo.Execute(tmp);
                    dbTo.Execute("CREATE INDEX xml7980a_index1 ON xml7980a(MA_TINH,KY_QT,MA_CSKCB);");
                }
            }
            /* Ghi lại danh sách bảng cần sao chép */
            dbTo.Execute($"UPDATE xmlthread SET args = '{string.Join(",", tablesCopy)}' WHERE id='{id}'");
            /* Loại bỏ các bảng đã sao chép */
            if (tablesCopy.Count > 1)
            {
                int i = -1;
                tmp = $"{dbTo.getValue($"SELECT args2 FROM WHERE id='{id}'")}";
                if (tmp != "")
                {
                    foreach (var table in tablesCopy) { i++; if (table == tmp) { break; } }
                    if (i > 0) { var obj = new List<string>(); for (; i < tablesCopy.Count; i++) { obj.Add(tablesCopy[i]); } tablesCopy = obj; }
                }
            }
            foreach (var table in tablesCopy) { XMLCopyTable2(dbTo, dbFrom, dbXML, id, table, fileName, batchSize, colsMD5); }
            dbFrom.Close();
        }
    }
}