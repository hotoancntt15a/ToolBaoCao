using ICSharpCode.SharpZipLib.GZip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Web.Mvc;

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
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "xml", $"{Session["idtinh"]}_{Session["iduser"]}".GetMd5Hash());
            var d = new System.IO.DirectoryInfo(folderTemp);
            if (d.Exists == false) { d.Create(); }
            return View();
        }
        private void CopyData(dbSQLite dbTo, dbSQLite dbFrom, string tableName) { 
        
        }
        /// <summary>
        /// idThread = {MaTinh}|{ID table XML}
        /// </summary>
        /// <param name="idThread">{MaTinh}|{ID table XML}</param>
        private void threadCopyXML(string idThread)
        {
            string tmp = "", folderTemp = "", folderSave = "", id = "", matinh = "";
            try
            {
                var objs = idThread.Split('|');
                if (idThread.Length != 2) { throw new Exception($"Tham số không đúng idThread XML '{idThread}'"); }
                id = objs[1];
                matinh = objs[0];
                folderTemp = Path.Combine(AppHelper.pathTemp, "xml", $"t{matinh}");
                folderSave = Path.Combine(AppHelper.pathAppData, "xml", $"t{matinh}");
                var dbXML = BuildDatabase.getDataXML(matinh);
                var data = dbXML.getDataTable($"SELECT * FROM xml WHERE id='{id}'");
                if (data.Rows.Count == 0) { throw new Exception($"Thread XML có id '{id}' không tồn tại hoặc đã bị xoá khỏi hệ thống"); }
                var item = new Dictionary<string, string>();
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    if (data.Columns[i].ColumnName.StartsWith("time")) { continue; }
                    item[data.Columns[i].ColumnName] = data.Rows[0][i].ToString();
                }
                var lfile = item["arg"].Split('|').ToList();
                int ij = 0;
                var xmldb = new dbSQLite(Path.Combine(AppHelper.pathAppData, "xml", $"xml_{id}.db"));
                foreach (string f in lfile)
                {
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
                                        db.Close();
                                        tmp = Path.Combine(folderSave, $"xml_{id}.db");
                                        /* Xoá đi nếu tồn tại rồi */
                                        if (System.IO.File.Exists(tmp)) { System.IO.File.Delete(tmp); }
                                        /* Chuyển về thư mục chính */
                                        System.IO.File.Move(fdb, tmp);
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
                            db.Close();
                            if (data.Rows.Count == 0) { throw new Exception($"Thread '{id}' có tập tin '{f}' không có dữ liệu"); }
                        }
                        catch (Exception exDB) { tmp = exDB.Message; }
                        db.Close();
                        continue;
                    }
                }
            }
            catch (Exception ex) { AppHelper.saveError(ex.getLineHTML()); }
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
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "xml", $"t{matinh}");
            var folderSave = Path.Combine(AppHelper.pathApp, "temp", "xml", $"tinh{matinh}");
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

                ViewBag.files = lFilesProcess;
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }

        public ActionResult TruyVan()
        {
            var matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            /* Tài khoản system có thể xem được tất cả
             * Tài khoản admin tỉnh xem được toàn bộ của tỉnh được phân
             * Tải khoản người dùng chỉnh xem các báo cáo mình tạo ra
             */
            try
            {
                var mode = Request.getValue("mode");
                if (mode == "truyvan")
                {
                    var folderSave = Path.Combine(AppHelper.pathApp, "temp", "xml", $"tinh{matinh}");
                    var d = new DirectoryInfo(folderSave);
                    ViewBag.data = d.GetFiles().ToList();
                    return View();
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
    }
}