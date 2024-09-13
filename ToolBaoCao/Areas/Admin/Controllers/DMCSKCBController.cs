using System;
using System.Collections.Generic;
using System.Data;
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