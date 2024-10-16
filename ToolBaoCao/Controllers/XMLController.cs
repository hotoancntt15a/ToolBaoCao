using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace ToolBaoCao.Controllers
{
    public class XMLController : ControllerCheckLogin
    {
        // GET: XML
        public ActionResult Index()
        {
            try
            {
                var d = new DirectoryInfo(Path.Combine(AppHelper.pathAppData, "xml"));
                if (d.Exists == false) { d.Create(); }
                d = new DirectoryInfo(Path.Combine(AppHelper.pathTemp, "xml"));
                if (d.Exists == false) { d.Create(); }

                if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }

        public ActionResult Buoc1()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            /* Tạo thư mục tạm */
            var folderTemp = Path.Combine(AppHelper.pathTemp, "xml", $"t{Session["idtinh"]}");
            var d = new System.IO.DirectoryInfo(folderTemp);
            if (d.Exists == false) { d.Create(); }
            return View();
        }

        public ActionResult Buoc2()
        {
            var timeStart = DateTime.Now;
            var idUser = $"{Session["iduser"]}";
            var matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            if (Request.Files.Count == 0) { ViewBag.Error = "Không có tập tin dữ liệu nào đẩy lên"; return View(); }
            var id = $"{timeStart:yyMMddHHmmss}_{matinh}_{timeStart.Millisecond:000}";
            var timeUp = timeStart.toTimestamp().ToString();
            var folderTemp = Path.Combine(AppHelper.pathTemp, "xml", $"t{matinh}");
            var tmp = "";
            ViewBag.id = id;
            try
            {
                var lWaitProcess = new List<string>();
                var lFilesProcess = new List<string>();
                /* Xoá hết các File có trong thư mục */
                var d = new System.IO.DirectoryInfo(folderTemp);
                if (d.Exists == false) { d.Create(); }
                /* Đọc và kiểm tra các tập tin */
                var list = new List<string>() { ".db", ".zip" };
                for (int i = 0; i < Request.Files.Count; i++)
                {
                    tmp = Path.GetExtension(Request.Files[i].FileName).ToLower();
                    if (list.Contains(tmp) == false) { continue; }
                    lFilesProcess.Add($"{Request.Files[i].FileName} ({Request.Files[i].ContentLength.getFileSize()})");
                    var fstmp = Path.Combine(folderTemp, $"xml{id}_{i}{tmp}");
                    Request.Files[i].SaveAs(fstmp);
                    lWaitProcess.Add(fstmp.Replace(AppHelper.pathApp, ""));
                }
                if (lWaitProcess.Count == 0) { throw new Exception("Không có dữ liệu đẩy lên phù hợp"); }
                var db = BuildDatabase.getDataXML(matinh);
                var item = new Dictionary<string, string>() {
                    { "id", id},
                    { "name", string.Join(", ", lFilesProcess)},
                    { "args", string.Join("|", lWaitProcess)},
                    { "title", "Đã vào hàng đợi xử lý"},
                    { "matinh", matinh},
                    { "time1", $"{DateTime.Now.toTimestamp()}"},
                    { "iduser", idUser}
                };
                db.Update("xml", item);
                db.Close();
                Thread t = new Thread(new ThreadStart(() =>
                {
                    try { XMLThread($"{matinh}|{id}"); }
                    catch (Exception exT) { AppHelper.saveError($"Error XMLThread({id}): {exT.Message}"); }
                }));
                t.Start();
                /*
                var itemTask = new ItemTask(id, $"Controller.XML.{id}", "Controller.XML", $"{matinh}|{id}", long.Parse(item["time1"]));
                AppHelper.threadManage.Add(itemTask);
                */
                ViewBag.files = lFilesProcess;
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }

        public ActionResult TruyVan()
        {
            var matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            try
            {
                var mode = Request.getValue("mode");
                if (mode == "truyvan")
                {
                    var dbXML = BuildDatabase.getDataXML(matinh);
                    /* Call Thread IF Exists */
                    var tsql = "SELECT *, datetime(time1, 'auto', '+7 hour') AS thoigian1 FROM xml ORDER BY time1 DESC LIMIT 50";
                    var data = dbXML.getDataTable(tsql);
                    ViewBag.data = data;
                    dbXML.Close();
                }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }

        public ActionResult Update()
        {
            var idtinh = $"{Session["idtinh"]}";
            if (idtinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            var id = Request.getValue("objectid");
            ViewBag.id = id;
            try
            {
                var item = new Dictionary<string, string>();
                ViewBag.data = item;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getErrorSave()}"; }
            return View();
        }

        public ActionResult Delete()
        {
            var timeStart = DateTime.Now;
            string ids = Request.getValue("id");
            var lid = new List<string>();
            string mode = Request.getValue("mode");
            try
            {
                if (string.IsNullOrEmpty(ids)) { return Content("Không có tham số".BootstrapAlter("warning")); }
                /* Kiểm tra danh sách nếu có */
                lid = ids.Split(new[] { '|', ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                ViewBag.data = string.Join(",", lid);
                if (mode == "force")
                {
                    foreach (string id in lid) { DeleteXML(id, true); }
                    return Content($"Xoá thành công dữ liệu có ID '{string.Join(", ", lid)}' ({timeStart.getTimeRun()})".BootstrapAlter());
                }
            }
            catch (Exception ex) { return Content(ex.getErrorSave().BootstrapAlter("warning")); }
            return View();
        }

        private void DeleteXML(string id, bool throwEx = false)
        {
            /* ID: {yyyyMMddHHmmss}_{idtinh}_{Milisecon}*/
            var tmpl = id.Split('_');
            if (tmpl.Length != 3)
            {
                if (throwEx == false) { return; }
                throw new Exception("ID Báo cáo không đúng định dạng {yyyyMMddHHmmss}_{idtinh}_{Milisecon}: " + id);
            }
            string idtinh = tmpl[1];
            /* Xoá hết các file trong mục lưu trữ App_Data/bcThang */
            var folder = new DirectoryInfo(Path.Combine(AppHelper.pathAppData, "xml", $"t{idtinh}"));
            if (folder.Exists)
            {
                foreach (var f in folder.GetFiles($"xml{id}*.*")) { try { f.Delete(); } catch { } }
            }
            folder = new DirectoryInfo(Path.Combine(AppHelper.pathTemp, "xml", $"t{idtinh}"));
            if (folder.Exists)
            {
                foreach (var f in folder.GetFiles($"xml{id}*.*")) { try { f.Delete(); } catch { } }
            }
            /* Xoá trong cơ sở dữ liệu */
            var db = BuildDatabase.getDataXML(idtinh);
            try
            {
                var idBaoCao = id.sqliteGetValueField();
                db.Execute($@"DELETE FROM xml WHERE id='{idBaoCao}' AND time2 <> 0;");
                db.Close();
            }
            catch (Exception ex)
            {
                var msg = ex.getErrorSave();
                if (throwEx) { throw new Exception(msg); }
            }
            finally { db.Close(); }
        }

        private void XMLCopyTable(dbSQLite dbTo, dbSQLite dbFrom, dbSQLite dbXML, string id)
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
                    /* CREATE INDEX MA_TINH,KY_QT,MA_CHA,MA_CSKCB*/
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
                    /* CREATE INDEX MA_TINH,KY_QT,MA_CHA,MA_CSKCB*/
                }
            }
            int batchSize = 1500; double rowCopyed = 0;
            string totalRow = "", tableName = "xml123";
            if (tablesFrom.Contains(tableName))
            {
                totalRow = $"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName}")}".FormatCultureVN();
                if (totalRow == "0") { throw new Exception($"{fileName}: Không có dữ liệu xml123"); }
                dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}(0/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                rowCopyed = 0;
                var data = dbFrom.getDataTable($"SELECT * FROM {tableName} LIMIT 1");
                data.Rows.RemoveAt(0);
                var reader = dbFrom.getDataReader($"SELECT * FROM {tableName}");
                while (reader.Read())
                {
                    if (data.Rows.Count >= batchSize)
                    {
                        /* Copy AND ignore */
                        dbTo.Insert(tableName, data, "IGNORE", batchSize);
                        rowCopyed += data.Rows.Count;
                        dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                        data.Rows.Clear();
                    }
                    DataRow dr = data.NewRow();
                    foreach (DataColumn c in data.Columns) { dr[c.ColumnName] = reader[c.ColumnName]; }
                    dr["HO_TEN"] = $"{dr["HO_TEN"]}".MD5Encrypt();
                    dr["NGAY_SINH"] = $"{dr["NGAY_SINH"]}".MD5Encrypt();
                    data.Rows.Add(dr);
                }
                if (data.Rows.Count > 0)
                {
                    /* Copy AND ignore */
                    dbTo.Insert(tableName, data, "IGNORE", batchSize);
                    rowCopyed += data.Rows.Count;
                    dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                }
                reader.Close();
            }
            tableName = "xml7980a";
            if (tablesFrom.Contains(tableName))
            {
                totalRow = $"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName}")}".FormatCultureVN();
                if (totalRow == "0") { throw new Exception($"{fileName}: Không có dữ liệu xml123"); }
                dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}(0/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                rowCopyed = 0;
                var data = dbFrom.getDataTable($"SELECT * FROM {tableName} LIMIT 1");
                data.Rows.RemoveAt(0);
                var reader = dbFrom.getDataReader($"SELECT * FROM {tableName}");
                while (reader.Read())
                {
                    if (data.Rows.Count >= batchSize)
                    {
                        /* Copy AND ignore */
                        dbTo.Insert(tableName, data, "IGNORE", batchSize);
                        rowCopyed += data.Rows.Count;
                        dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                        data.Rows.Clear();
                    }
                    DataRow dr = data.NewRow();
                    foreach (DataColumn c in data.Columns) { dr[c.ColumnName] = reader[c.ColumnName]; }
                    dr["HO_TEN"] = $"{dr["HO_TEN"]}".MD5Encrypt();
                    dr["NGAY_SINH"] = $"{dr["NGAY_SINH"]}".MD5Encrypt();
                    data.Rows.Add(dr);
                }
                if (data.Rows.Count > 0)
                {
                    /* Copy AND ignore */
                    dbTo.Insert(tableName, data, "IGNORE", batchSize);
                    rowCopyed += data.Rows.Count;
                    dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                }
                reader.Close();
            }
            tableName = "bhyt7980a";
            if (tablesFrom.Contains(tableName))
            {
                totalRow = $"{dbFrom.getValue($"SELECT COUNT(ID) AS X FROM {tableName}")}".FormatCultureVN();
                if (totalRow == "0") { throw new Exception($"{fileName}: Không có dữ liệu xml123"); }
                dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}(0/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                /* Chuyển dữ liệu */
                rowCopyed = 0;
                var data = dbFrom.getDataTable($"SELECT * FROM {tableName} LIMIT 1");
                data.Rows.RemoveAt(0);
                var reader = dbFrom.getDataReader($"SELECT * FROM {tableName}");
                while (reader.Read())
                {
                    if (data.Rows.Count >= batchSize)
                    {
                        /* Copy AND ignore */
                        dbTo.Insert("xml7980a", data, "IGNORE", batchSize);
                        rowCopyed += data.Rows.Count;
                        dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                        data.Rows.Clear();
                    }
                    DataRow dr = data.NewRow();
                    foreach (DataColumn c in data.Columns) { dr[c.ColumnName] = reader[c.ColumnName]; }
                    dr["HO_TEN"] = $"{dr["HO_TEN"]}".MD5Encrypt();
                    dr["NGAY_SINH"] = $"{dr["NGAY_SINH"]}".MD5Encrypt();
                    data.Rows.Add(dr);
                }
                if (data.Rows.Count > 0)
                {
                    /* Copy AND ignore */
                    dbTo.Insert("xml7980a", data, "IGNORE", batchSize);
                    rowCopyed += data.Rows.Count;
                    dbXML.Execute($"UPDATE xml SET title = 'Sao chép {tableName}({rowCopyed.FormatCultureVN()}/{totalRow}) từ {fileName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                }
                reader.Close();
            }
            dbFrom.Close(); dbTo.Close(); dbXML.Close();
        }

        /// <summary>
        /// idThread = {MaTinh}|{ID table XML}
        /// </summary>
        /// <param name="idThread">{MaTinh}|{ID table XML}</param>
        public void XMLThread(string idThread)
        {
            string tmp = "", folderTemp = "", folderSave = "", id = "", matinh = "";
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
                var data = dbXML.getDataTable($"SELECT * FROM xml WHERE id='{id}'");
                if (data.Rows.Count == 0) { throw new Exception($"Thread XML có id '{id}' không tồn tại hoặc đã bị xoá khỏi hệ thống"); }
                var item = new Dictionary<string, string>();
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    item[data.Columns[i].ColumnName] = data.Rows[0][i].ToString();
                }
                var lfile = item["args"].Split('|').ToList();
                int ij = 0;
                var xmldb = new dbSQLite(Path.Combine(AppHelper.pathAppData, "xml", $"t{matinh}", $"xml{id}.db"));
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
                                ij++;
                                if (entry.FullName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
                                {
                                    dbXML.Execute($"UPDATE xml SET title = 'Đang giải nén {entry.FullName} ({DateTime.Now:dd/MM/yyyy HH:mm})' WHERE id='{id}'");
                                    var fdb = Path.Combine(folderTemp, $"xml{id}_zip{ij}.db");
                                    entry.ExtractToFile(fdb, overwrite: true);
                                    var dbFrom = new dbSQLite(fdb);
                                    try
                                    {
                                        /* Kiểm tra có đúng cấu trúc dữ liệu không? */
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
                                        XMLCopyTable(xmldb, dbFrom, dbXML, id);
                                        dbFrom.Close();
                                    }
                                    catch (Exception exDB)
                                    {
                                        dbFrom.Close();
                                        AppHelper.saveError($"XMLThread({id}): {entry.FullName} IN {f} - {exDB.Message}");
                                        continue;
                                    }
                                    /* Xoá đi sau khi sao chép song */
                                    try { System.IO.File.Delete(fdb); } catch { }
                                }
                            }
                        }
                        continue;
                    }
                    if (ext == ".db")
                    {
                        var dbFrom = new dbSQLite(fileName);
                        try
                        {
                            /* Kiểm tra có đúng cấu trúc dữ liệu không? */
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
                            XMLCopyTable(xmldb, dbFrom, dbXML, id);
                        }
                        catch (Exception exDB) { tmp = exDB.Message; }
                        dbFrom.Close();
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
            dbXML.Close();
        }
    }
}