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
            try
            {
                var msg = new List<string>();
                var w = new List<string>();
                var idObject = Request.getValue("id");
                if (!string.IsNullOrEmpty(idObject))
                {
                    idObject = Regex.Replace(idObject, "[, /|]+", ",");
                    if (Regex.IsMatch(idObject, "^[0-9a-z,]+$", RegexOptions.IgnoreCase) == false) { msg.Add($"Mã '{idObject}' không đúng định dạng"); }
                }

                var ten = Request.getValue("ten");
                var ma_tinh = Request.getValue("ma_tinh");
                var macaptren = Request.getValue("macaptren");
                var tenhuyen = Request.getValue("tenhuyen");
                ViewBag.Data = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM dmcskcb ORDER BY ma_tinh, ten");
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }
    }
}