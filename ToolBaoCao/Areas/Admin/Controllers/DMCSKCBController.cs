using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class DMCSKCBController : Controller
    {
        /* GET: Admin/DMCSKCB */

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Buoc1()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            return View();
        }

        public ActionResult Buoc2()
        {
            var timeStart = DateTime.Now;
            var idUser = $"{Session["iduser"]}";
            if (Request.Files.Count == 0) { ViewBag.Error = "Không có tập tin dữ liệu nào đẩy lên"; return View(); }
            var timeUp = timeStart.toTimestamp().ToString();
            var fileName = $"dmcskcb_{timeUp}.xlsx";
            var tmp = ""; IWorkbook workbook = null;
            try
            {
                /* Xoá hết các File có trong thư mục */
                if (Request.Files.Count == 0) { throw new Exception("Không có tập tin nào được đẩy lên"); }
                var lsFile = new List<string>();
                var ext = Path.GetExtension(Request.Files[0].FileName).ToLower();
                if (ext == ".xlsx")
                {
                    lsFile.Add($"{Request.Files[0].FileName} ({Request.Files[0].ContentLength.getFileSize()})");
                    ViewBag.files = lsFile;
                    /* Cập nhật dự toán được giao trong năm của csyt */
                    ViewBag.mode = "update";
                    workbook = new XSSFWorkbook(Request.Files[0].InputStream);
                    ISheet sheet = workbook.GetSheetAt(0);
                    if (sheet.LastRowNum < 5) { throw new Exception("Excel Không có dữ liệu để cập nhật."); }
                    var data = new DataTable();
                    for (int i = 0; i < 34; i++) { data.Columns.Add($"c{i}"); }
                    data.Columns[0].ColumnName = "ma_tinh";
                    data.Columns[1].ColumnName = "id";
                    data.Columns[2].ColumnName = "ten";
                    data.Columns[3].ColumnName = "tuyencmkt";
                    data.Columns[4].ColumnName = "hangbv";
                    data.Columns[5].ColumnName = "loaibv";
                    data.Columns[6].ColumnName = "tenhuyen";
                    data.Columns[7].ColumnName = "donvi";
                    data.Columns[8].ColumnName = "madinhdanh";
                    data.Columns[9].ColumnName = "macaptren";
                    data.Columns[10].ColumnName = "diachi";
                    data.Columns[11].ColumnName = "ttduyet";
                    data.Columns[12].ColumnName = "hieuluc";
                    data.Columns[13].ColumnName = "tuchu";
                    data.Columns[14].ColumnName = "trangthai";
                    data.Columns[15].ColumnName = "hangdv";
                    data.Columns[16].ColumnName = "hangthuoc";
                    data.Columns[17].ColumnName = "dangkykcb";
                    data.Columns[18].ColumnName = "hinhthuctochuc";
                    data.Columns[19].ColumnName = "hinhthucthanhtoan";
                    data.Columns[20].ColumnName = "ngaycapma";
                    data.Columns[21].ColumnName = "kcb";
                    data.Columns[22].ColumnName = "ngayngunghd";
                    data.Columns[23].ColumnName = "kt7";
                    data.Columns[24].ColumnName = "kcn";
                    data.Columns[25].ColumnName = "knl";
                    data.Columns[26].ColumnName = "cpdtt43";
                    data.Columns[27].ColumnName = "slthedacap";
                    data.Columns[28].ColumnName = "donvichuquan";
                    data.Columns[29].ColumnName = "mota";
                    /* ma_huyen	userid */
                    data.Columns[30].ColumnName = "loaichuyenkhoa";
                    data.Columns[31].ColumnName = "ngaykyhopdong";
                    data.Columns[32].ColumnName = "ngayhethieuluc";
                    data.Columns[33].ColumnName = "userid";
                    /* Kiểm tra dữ liệu */
                    IRow row = null;
                    var pattern = @"[0-9A-Z]+";
                    for (int i = 0; i < sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i); if (row == null) { continue; }
                        /* Cột thứ tự */
                        tmp = row.GetCell(0).GetValueAsString();
                        if (tmp.isNumberUSInt(true) == false) { continue; }
                        var dr = data.NewRow();
                        for (int j = 1; j < 33; j++) { dr[j] = row.GetCell(j).GetValueAsString(); }
                        if (Regex.IsMatch($"{dr["id"]}".Trim(), pattern) == false) { continue; }
                        if (Regex.IsMatch($"{dr["madinhdanh"]}".Trim(), pattern) == false) { continue; }
                        dr[0] = $"{dr["id"]}".Substring(0, 2);
                        data.Rows.Add(dr);
                    }
                    if (data.Rows.Count == 0) { throw new Exception("Không có dữ liệu để cập nhật."); }
                    AppHelper.dbSqliteMain.Insert("dmcskcb", data, "replace");
                    ViewBag.Message = $"Đã cập nhật DMCSKCB ({timeStart.getTimeRun()})";
                }
            }
            catch (Exception ex) { ViewBag.Error = "TryCatch: " + ex.getLineHTML(); }
            if (workbook != null) { workbook.Close(); workbook.Dispose(); }
            return View();
        }

        public ActionResult Update()
        {
            var timeStart = DateTime.Now;
            try
            {
                var mode = Request.getValue("mode");
                var id = Request.getValue("objectid");
                ViewBag.id = id;
                if (mode == "update")
                {
                    var item = new Dictionary<string, string>();
                    var lsrq = new List<string>() { "id", "ten", "tuyencmkt", "hangbv", "loaibv", "tenhuyen", "donvi", "madinhdanh", "macaptren", "diachi", "ttduyet", "hieuluc", "tuchu", "trangthai", "hangdv", "hangthuoc", "dangkykcb", "hinhthuctochuc", "hinhthucthanhtoan", "ngaycapma", "kcb", "ngayngunghd", "kt7", "kcn", "knl", "cpdtt43", "slthedacap", "donvichuquan", "mota", "loaichuyenkhoa", "ngaykyhopdong", "ngayhethieuluc", "ma_tinh", "ma_huyen" };
                    foreach (var v in lsrq) { item.Add(v, Request.getValue(v).Trim()); }
                    item["userid"] = $"{Session["iduser"]}";

                    if (item["id"] == "") { return Content("Mã bỏ trống".BootstrapAlter("warning")); }
                    if (item["ten"] == "") { return Content("Tên bỏ trống".BootstrapAlter("warning")); }
                    if (Regex.IsMatch(item["id"], @"^[0-9a-z]+$", RegexOptions.IgnoreCase) == false) { return Content($"Mã không đúng '{id}'".BootstrapAlter("warning")); }
                    if (Regex.IsMatch(item["slthedacap"], @"^\d+$") == false) { item["slthedacap"] = "0"; }
                    if (item["ngayngunghd"] != "") { if (item["ngayngunghd"].isDateVN() == false) { throw new Exception($"Ngày ngưng hoạt động không đúng định dạng Ngày/Tháng/Năm {item["ngayngunghd"]}"); } }
                    if (item["ngaycapma"] != "") { if (item["ngaycapma"].isDateVN() == false) { throw new Exception($"Ngày cấp mã không đúng định dạng Ngày/Tháng/Năm {item["ngaycapma"]}"); } }
                    if (item["madinhdanh"] != "") { if (Regex.IsMatch(item["madinhdanh"], @"^[0-9a-z]+$", RegexOptions.IgnoreCase) == false) { throw new Exception($"Mã định danh không đúng {item["madinhdanh"]}"); } }
                    if (item["macaptren"] != "") { if (Regex.IsMatch(item["macaptren"], @"^[0-9a-z]+$", RegexOptions.IgnoreCase) == false) { throw new Exception($"Mã cấp trên không đúng định dạng {item["macaptren"]}"); } }
                    if (item["ma_tinh"] != "") { if (Regex.IsMatch(item["ma_tinh"], @"^[0-9a-z]+$", RegexOptions.IgnoreCase) == false) { throw new Exception($"Mã tỉnh không đúng định dạng {item["ma_tinh"]}"); } }
                    if (item["ma_huyen"] != "") { if (Regex.IsMatch(item["ma_huyen"], @"^[0-9a-z]+$", RegexOptions.IgnoreCase) == false) { throw new Exception($"Mã huyện không đúng định dạng {item["ma_huyen"]}"); } }

                    if (id != "") { if (id != item["id"]) { return Content($"Mã '{id}' sửa chữa không khớp '{item["id"]}'"); } }
                    AppHelper.dbSqliteMain.Update("dmcskcb", item, "replace");
                    return Content($"Lưu thành công ({timeStart.getTimeRun()})".BootstrapAlter());
                }
                if (id != "")
                {
                    if (Regex.IsMatch(id, @"^[0-9a-z]+$", RegexOptions.IgnoreCase) == false) { return Content($"Mã không đúng '{id}'".BootstrapAlter("warning")); }
                    var data = AppHelper.dbSqliteMain.getDataTable($"SELECT * FROM dmcskcb WHERE id='{id}'");
                    if (data.Rows.Count == 0) { return Content($"Cơ sở KCB có mã '{id}' không tồn tại hoặc bị xoá khỏi hệ thống".BootstrapAlter("danger")); }
                    var item = new Dictionary<string, object>();
                    foreach (DataColumn c in data.Columns) { item.Add(c.ColumnName, data.Rows[0][c.ColumnName]); }
                    ViewBag.data = item;
                }
            }
            catch (Exception ex) { return Content(ex.getErrorSave().BootstrapAlter("warning")); }
            return View();
        }

        public ActionResult Delete()
        {
            var timeStart = DateTime.Now;
            var mode = Request.getValue("mode");
            var id = Request.getValue("id");
            if (Regex.IsMatch(id, @"^[0-9a-z,]+$", RegexOptions.IgnoreCase) == false) { return Content($"Mã không đúng '{id}'".BootstrapAlter("warning")); }
            try
            {
                if (mode == "force")
                {
                    var listID = id.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < listID.Length; i++) { listID[i] = $"'{listID[i].sqliteGetValueField()}'"; }
                    AppHelper.dbSqliteMain.Execute($"DELETE FROM dmcskcb WHERE id IN ({string.Join(",", listID)});");
                    return Content($"Xoá thành công ({timeStart.getTimeRun()})".BootstrapAlter());
                }
            }
            catch (Exception ex) { return Content(ex.getErrorSave().BootstrapAlter("warning")); }
            ViewBag.id = id;
            return View();
        }

        public ActionResult TruyVan()
        {
            var tsql = "";
            try
            {
                var msg = new List<string>();
                var w = new List<string>();
                var tmp = Request.getValue("id");
                if (!string.IsNullOrEmpty(tmp))
                {
                    tmp = Regex.Replace(tmp, "[, /|]+", ",");
                    if (Regex.IsMatch(tmp, "^[0-9a-z,]+$", RegexOptions.IgnoreCase))
                    {
                        w.Add($"id IN ('{tmp.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)}')");
                    }
                    else { msg.Add($"Mã '{tmp}' không đúng định dạng"); }
                }
                if (w.Count == 0)
                {
                    tmp = Request.getValue("ten");
                    if (string.IsNullOrEmpty(tmp) == false) { w.Add(AppHelper.dbSqliteMain.like("ten", tmp)); }
                    tmp = Request.getValue("ma_tinh");
                    if (string.IsNullOrEmpty(tmp) == false) { w.Add(AppHelper.dbSqliteMain.like("ma_tinh", tmp)); }
                    tmp = Request.getValue("macaptren");
                    if (string.IsNullOrEmpty(tmp) == false) { w.Add(AppHelper.dbSqliteMain.like("macaptren", tmp)); }
                    tmp = Request.getValue("tenhuyen");
                    if (string.IsNullOrEmpty(tmp) == false) { w.Add(AppHelper.dbSqliteMain.like("tenhuyen", tmp)); }
                }
                tsql = $"SELECT * FROM dmcskcb {(w.Count == 0 ? "" : "WHERE " + string.Join(" AND ", w))} ORDER BY ma_tinh, ten";
                ViewBag.Data = AppHelper.dbSqliteMain.getDataTable(tsql);
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(tsql); }
            return View();
        }
    }
}