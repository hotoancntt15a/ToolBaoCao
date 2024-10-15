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
            var lWaitProcess = new List<string>();
            var lFilesProcess = new List<string>();
            var tmp = "";
            ViewBag.id = id;
            try
            {
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
                    try
                    {
                        AppHelper.saveError($"Running XMLThead({id})");
                        XMLThread($"{matinh}|{id}");
                    }
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
                foreach (var f in folder.GetFiles($"xml_{id}.*")) { try { f.Delete(); } catch { } }
            }
            folder = new DirectoryInfo(Path.Combine(AppHelper.pathTemp, "xml", $"t{idtinh}"));
            if (folder.Exists)
            {
                foreach (var f in folder.GetFiles($"xml_{id}.*")) { try { f.Delete(); } catch { } }
            }
            /* Xoá trong cơ sở dữ liệu */
            var db = BuildDatabase.getDataXML(idtinh);
            try
            {
                var idBaoCao = id.sqliteGetValueField();
                db.Execute($@"DELETE FROM xml WHERE id='{idBaoCao}';");
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
        public void XMLThread(string idThread)
        {
            AppHelper.saveError($"RUNNING XMLThead({idThread})");
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
                var xmldb = new dbSQLite(Path.Combine(AppHelper.pathAppData, "xml", $"t{matinh}", $"xml_{id}.db"));
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
                                        XMLCopyTable(dbXML, db, xmldb, id);
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
                            XMLCopyTable(dbXML, db, xmldb, id);
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