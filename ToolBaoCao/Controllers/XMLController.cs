using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using zModules.NPOIExcel;

namespace ToolBaoCao.Controllers
{
    public class XMLController : ControllerCheckLogin
    {
        /* GET: XML */

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
                var d = new DirectoryInfo(folderTemp);
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
                db.Update("xmlthread", item);
                db.Close();
                var itemTask = new ItemTask(id, $"Controller.XML.{id}", "Controller.XML", $"{matinh}|{id}", long.Parse(item["time1"]));
                AppHelper.threadManage.Add(itemTask);
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
                    var ngay = Request.getValue("ngay1");
                    var t = DateTime.Now;
                    var w = new List<string>();
                    if (ngay.isDateVN(out t) == false) { throw new Exception($"Từ ngày không đúng định dạng (Ngày/Tháng/Năm): {ngay}"); }
                    w.Add($"time1 >= {t.toTimestamp()}");
                    ngay = Request.getValue("ngay2");
                    if (ngay.isDateVN(out t) == false) { throw new Exception($"Đến ngày không đúng định dạng (Ngày/Tháng/Năm): {ngay}"); }
                    w.Add($"time1 < {t.AddDays(1).toTimestamp()}");
                    var tsql = "SELECT *, datetime(time1, 'auto', '+7 hour') AS thoigian1 FROM xmlthread";
                    if (w.Count > 0) { tsql += " WHERE " + string.Join(" AND ", w); }
                    var dbXML = BuildDatabase.getDataXML(matinh);
                    var data = dbXML.getDataTable(tsql + " ORDER BY time1 DESC LIMIT 50");
                    var view = data.AsEnumerable().Where(x => x.Field<string>("title") == "Thread was being aborted.").ToList();
                    ViewBag.threadabort = view.Count;
                    if (view.Count > 0)
                    {
                        foreach (DataRow row in view)
                        {
                            var itemTask = new ItemTask(row["id"].ToString(), $"Controller.XML.{row["id"]}", "Controller.XML", $"{row["matinh"]}|{row["id"]}", long.Parse(row["time1"].ToString()));
                            AppHelper.threadManage.Add(itemTask, false);
                            dbXML.Execute($"UPDATE xmlthread SET args2 = args2 || '; Recreate Thread {DateTime.Now:HH:mm:ss}', time2=0 WHERE id='{row["id"]}'");
                        }
                        AppHelper.threadManage.Call();
                        data = dbXML.getDataTable(tsql + " ORDER BY time1 DESC LIMIT 50");
                    }
                    dbXML.Close();
                    ViewBag.data = data;
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
            if (id == "") { ViewBag.Error = "Tham số bỏ trống"; return View(); }
            if (Regex.IsMatch(id, "^[a-z0-9_]+$") == false) { ViewBag.Error = $"Tham số không đúng {id}"; return View(); }
            var lsFieldRemove = new List<string>() { "ma_the", "ho_ten", "ngay_sinh" };
            var timeStart = DateTime.Now;
            var mode = Request.getValue("mode"); var pathDB = ""; var limit = 1000;
            var db = new dbSQLite();
            try
            {
                pathDB = Path.Combine(AppHelper.pathAppData, "xml", $"t{idtinh}", $"xml{id}.db");
                if (System.IO.File.Exists(pathDB) == false) { throw new Exception($"Không tìm thấy XMLThread có ID '{id}'"); }
                if (mode == "tsql")
                {
                    string dataName = Request.getValue("data");
                    string tsql = Request.getValue("tsql").Trim();
                    if (tsql == "") { throw new Exception($"<div class=\"alert alert-warning\">TSQL bỏ trống</div>"); }
                    if (AppHelper.IsUpdateData(tsql)) { throw new Exception($"<div class=\"alert alert-warning\">Hệ thống chặn cập nhật dữ liệu: {tsql}</div>"); }
                    db = new dbSQLite(pathDB);
                    if (Regex.IsMatch(tsql, "^pragma ", RegexOptions.IgnoreCase) == false)
                    {
                        if (!Regex.IsMatch(tsql, @"limit\s+[0-9]+;?$", RegexOptions.IgnoreCase)) { tsql += $" LIMIT {limit}"; }
                    }
                    var data = db.getDataTable(tsql);
                    data = data.RemoveColumns(lsFieldRemove, false);
                    ViewBag.content = $"Data {dataName}; Thao tác thành công ({timeStart.getTimeRun()}); TSQL: {tsql}";
                    ViewBag.data = data;
                    db.Close();
                    return View();
                }
                if (mode == "xlsx")
                {
                    string dataName = Request.getValue("data");
                    string tsql = Request.getValue("tsql").Trim();
                    if (tsql == "") { throw new Exception($"<div class=\"alert alert-warning\">TSQL bỏ trống</div>"); }
                    if (AppHelper.IsUpdateData(tsql)) { throw new Exception($"<div class=\"alert alert-warning\">Hệ thống chặn cập nhật dữ liệu: {tsql}</div>"); }
                    db = new dbSQLite(pathDB);
                    if (Regex.IsMatch(tsql, "^pragma ", RegexOptions.IgnoreCase) == false)
                    {
                        if (!Regex.IsMatch(tsql, @"limit\s+[0-9]+;?$", RegexOptions.IgnoreCase)) { tsql += $" LIMIT 63999"; }
                    }
                    var data = db.getDataTable(tsql);
                    db.Close();
                    data = data.RemoveColumns(lsFieldRemove, false);
                    var wb = XLSX.exportExcel(data);
                    using (MemoryStream stream = new MemoryStream())
                    {
                        wb.Write(stream);
                        return File(stream.ToArray(), "application/octet-stream", $"id{DateTime.Now.toTimestamp()}.xlsx");
                    }
                }
                else
                {
                    db = new dbSQLite(pathDB);
                    var tables = db.getAllTables();
                    var tablesInfo = new List<string>();
                    foreach (var table in tables)
                    {
                        var data = db.getDataTable($"SELECT name, type FROM pragma_table_info('{table}');");
                        var cols = new List<string>();
                        foreach (DataRow row in data.Rows) { cols.Add($"{row["name"]}({row["type"]})"); }
                        tablesInfo.Add($"{table} ({cols.Count} cột): {string.Join("; ", cols)}");
                    }
                    ViewBag.tables = tablesInfo;
                }
            }
            catch (Exception ex)
            {
                if (Request.getValue("layout") == "null") { return Content($"<div class=\"alert alert-warning\">Lỗi {pathDB}: {ex.getLineHTML()}</div>"); }
                ViewBag.Error = $"<div class=\"alert alert-warning\">Lỗi: {ex.getLineHTML()}</div>";
            }
            db.Close();
            return View();
        }

        public ActionResult StoreTSQL()
        {
            string mode = Request.getValue("mode");
            try
            {
                if (mode == "view")
                {
                    var db = BuildDatabase.getDataStoreTSQL();
                    var dt = db.getDataTable("SELECT * FROM storetsql ORDER BY timeup LIMIT 1000");
                    ViewBag.data = dt;
                    return View();
                }
            } catch(Exception ex) { ViewBag.Error = ex.getLineHTML(); }
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
            if (AppHelper.threadManage.IDRunning == id) { /* Thread đang chạy không thể xoá */ return; }
            var idBaoCao = id.sqliteGetValueField();
            /* Xoá trong cơ sở dữ liệu */
            var db = BuildDatabase.getDataXML(idtinh);
            try
            {
                db.Execute($@"DELETE FROM xmlthread WHERE id='{idBaoCao}' AND time2 > 0;");
                db.Close();
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
            }
            catch (Exception ex)
            {
                var msg = ex.getErrorSave();
                if (throwEx) { throw new Exception(msg); }
            }
            finally { db.Close(); }
        }
    }
}