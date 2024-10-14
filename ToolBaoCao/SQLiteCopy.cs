using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;

namespace ToolBaoCao
{
    public static class SQLiteCopy
    {
        public static void CopyTableXML(dbSQLite dbTo, dbSQLite dbFrom, dbSQLite dbXML, string id)
        {
            var tablesTo = dbTo.getAllTables();
            var tablesFrom = dbFrom.getAllTables();
            var fileName = Path.GetFileName(dbFrom.getPathDataFile());
            var tmp = "";
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
                    dbTo.Execute(tmp);
                }
            }
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
                    dbTo.Execute(tmp);
                }
            }
            var batchSize = 1000;
            var tsql = "";
            if (tablesFrom.Contains("xml123"))
            {
                dbXML.Execute($"UPDATE xml SET title = 'Đang thao sao chép dữ liệu bảng xml123 tại tập tin {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                var idStart = "0";
                while (true)
                {
                    tsql = $"SELECT * FROM xml123 WHERE ID > {idStart} ORDER BY ID LIMIT {batchSize}";
                    var data = dbFrom.getDataTable(tsql);
                    if (data.Rows.Count > 0)
                    {
                        foreach (DataRow r in data.Rows)
                        {
                            r["HO_TEN"] = $"{r["HO_TEN"]}".MD5Encrypt();
                            r["NGAY_SINH"] = $"{r["NGAY_SINH"]}".MD5Encrypt();
                        }
                        /* Copy AND ignore */
                        dbTo.Insert("xml123", data, "IGNORE", batchSize);
                    }
                    idStart = data.Rows[data.Rows.Count - 1]["ID"].ToString();
                    if (data.Rows.Count < batchSize) { break; }
                }
            }
            if (tablesFrom.Contains("xml7980a"))
            {
                dbXML.Execute($"UPDATE xml SET title = 'Đang thao sao chép dữ liệu bảng xml7980a tại tập tin {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                var idStart = "0";
                while (true)
                {
                    tsql = $"SELECT * FROM xml7980a WHERE ID > {idStart} ORDER BY ID LIMIT {batchSize}";
                    var data = dbFrom.getDataTable(tsql);
                    if (data.Rows.Count > 0)
                    {
                        foreach (DataRow r in data.Rows)
                        {
                            r["HO_TEN"] = $"{r["HO_TEN"]}".MD5Encrypt();
                            r["NGAY_SINH"] = $"{r["NGAY_SINH"]}".MD5Encrypt();
                        }
                        /* Copy AND ignore */
                        dbTo.Insert("xml7980a", data, "IGNORE", batchSize);
                    }
                    idStart = data.Rows[data.Rows.Count - 1]["ID"].ToString();
                    if (data.Rows.Count < batchSize) { break; }
                }
            }
            if (tablesFrom.Contains("bhyt7980a"))
            {
                dbXML.Execute($"UPDATE xml SET title = 'Đang thao sao chép dữ liệu bảng bhyt7980a tại tập tin {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                var idStart = "0";
                while (true)
                {
                    tsql = $"SELECT * FROM bhyt7980a WHERE ID > {idStart} ORDER BY ID LIMIT {batchSize}";
                    var data = dbFrom.getDataTable(tsql);
                    if (data.Rows.Count > 0)
                    {
                        foreach (DataRow r in data.Rows)
                        {
                            r["HO_TEN"] = $"{r["HO_TEN"]}".MD5Encrypt();
                            r["NGAY_SINH"] = $"{r["NGAY_SINH"]}".MD5Encrypt();
                        }
                        /* Copy AND ignore */
                        dbTo.Insert("xml7980a", data, "IGNORE", batchSize);
                    }
                    idStart = data.Rows[data.Rows.Count - 1]["ID"].ToString();
                    if (data.Rows.Count < batchSize) { break; }
                }
            }
        }

        /// <summary>
        /// idThread = {MaTinh}|{ID table XML}
        /// </summary>
        /// <param name="idThread">{MaTinh}|{ID table XML}</param>
        public static void threadCopyXML(string idThread)
        {
            string tmp = "", folderTemp = "", folderSave = "", id = "", matinh = "";
            var dbXML = BuildDatabase.getDataXML(matinh);
            try
            {
                var objs = idThread.Split('|');
                if (idThread.Length != 2) { throw new Exception($"Tham số không đúng idThread XML '{idThread}'"); }
                id = objs[1];
                matinh = objs[0];
                folderTemp = Path.Combine(AppHelper.pathTemp, "xml", $"t{matinh}");
                folderSave = Path.Combine(AppHelper.pathAppData, "xml", $"t{matinh}");
                var data = dbXML.getDataTable($"SELECT * FROM xml WHERE id='{id}'");
                if (data.Rows.Count == 0) { throw new Exception($"Thread XML có id '{id}' không tồn tại hoặc đã bị xoá khỏi hệ thống"); }
                var item = new Dictionary<string, string>();
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    item[data.Columns[i].ColumnName] = data.Rows[0][i].ToString();
                }
                var lfile = item["arg"].Split('|').ToList();
                int ij = 0;
                var xmldb = new dbSQLite(Path.Combine(AppHelper.pathAppData, "xml", $"xml_{id}.db"));
                foreach (string f in lfile)
                {
                    dbXML.Execute($"UPDATE xml SET title = 'Đang thao tác tại tập tin {f} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                    ij++;
                    var fileName = AppHelper.pathApp + f;
                    if (System.IO.File.Exists(fileName) == false) { throw new Exception($"Thread '{id}' có tập tin '{f}' không tồn tại trong hệ thống"); }
                    var ext = Path.GetExtension(fileName);
                    if (ext == ".zip")
                    {
                        using (ZipArchive archive = ZipFile.OpenRead(fileName))
                        {
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                if (entry.FullName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
                                {
                                    var fdb = Path.Combine(folderTemp, $"xml_{id}_{ij}.db");
                                    entry.ExtractToFile(fdb, overwrite: true);
                                    var db = new dbSQLite(fdb);
                                    try
                                    {
                                        /* Kiểm tra có đúng cấu trúc dữ liệu không? */
                                        var tables = db.getAllTables();
                                        var tsql = "SELECT MIN(KY_QT) AS X1, MAX(KY_QT) AS X2 FROM ";
                                        var tableName = "xml123";
                                        if (tables.Contains("xml7980a")) { tableName = "xml7980a"; }
                                        else { if (tables.Contains("bhyt7980a")) { tableName = "bhyt7980a"; } }
                                        data = db.getDataTable($"{tsql}{tableName} LIMIT 1");
                                        if (data.Rows.Count == 0) { db.Close(); throw new Exception($"Thread '{id}' có tập tin '{f}' không có dữ liệu"); }
                                        /* Chuyển dữ liệu */
                                        CopyTableXML(dbXML, db, xmldb, id);
                                        db.Close();
                                        /* Xoá đi sau khi sao chép song */
                                        if (System.IO.File.Exists(fdb)) { System.IO.File.Delete(fdb); }
                                    }
                                    catch (Exception exDB) { tmp = exDB.Message; continue; }
                                    db.Close();
                                }
                            }
                        }
                        continue;
                    }
                    if (ext == ".db")
                    {
                        var db = new dbSQLite(fileName);
                        try
                        {
                            /* Kiểm tra có đúng cấu trúc dữ liệu không? */
                            var tables = db.getAllTables();
                            var tsql = "SELECT MIN(KY_QT) AS X1, MAX(KY_QT) AS X2 FROM ";
                            if (tables.Contains("xml7980a")) { tsql += "xml7980a"; }
                            else
                            {
                                if (tables.Contains("bhyt7980a")) { tsql += "bhyt7980a"; }
                                else { tsql += "xml123"; }
                            }
                            data = db.getDataTable(tsql + " LIMIT 1");
                            if (data.Rows.Count == 0) { db.Close(); throw new Exception($"Thread '{id}' có tập tin '{f}' không có dữ liệu"); }
                            CopyTableXML(dbXML, db, xmldb, id);
                            db.Close();
                            System.IO.File.Delete(fileName);
                        }
                        catch (Exception exDB) { tmp = exDB.Message; }
                        db.Close();
                        continue;
                    }
                }
                dbXML.Execute($"UPDATE xml SET title = 'Hoàn thành', time2='{DateTime.Now.toTimestamp()}' WHERE id='{id}'");
            }
            catch (Exception ex)
            {
                dbXML.Execute($"UPDATE xml SET title = '{ex.Message.sqliteGetValueField()}', time2='{DateTime.Now.toTimestamp()}' WHERE id='{id}'");
                AppHelper.saveError(ex.getLineHTML());
            }
        }
    }
}