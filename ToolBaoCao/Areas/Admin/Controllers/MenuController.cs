using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class MenuController : Controller
    {
        // GET: Admin/Menu
        public ActionResult Index()
        {
            try
            {
                var data = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM wmenu ORDER BY link");
                ViewBag.Data = data;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getErrorSave()}"; }
            return View();
        }

        private string showMenuTree(DataTable dataMenu, int idMenuFather = -1, string jsfunction = "selectMenu", bool viewUrl = true)
        {
            if (dataMenu.Rows.Count == 0) { return ""; }
            var dcopy = dataMenu.Copy(); dcopy.Rows.Clear();
            string html = "<div class=\"viewmenutree\">";
            if (idMenuFather == -1)
            {
                foreach (DataRow r in dataMenu.Rows)
                {
                    var childen = dataMenu.Select($"idfather={r["id"]}");
                    var note = $"{r["note"]}".Trim(); if (note != "") { note = $" (<i>{note}</i>)"; }
                    if (childen.Length == 0) { html += $"<li> <a href=\"javascript:{jsfunction}(this,'{r["id"]}');\">{r["title"]}</a>{(note == "" ? "" : "")}"; continue; }
                    dcopy.Rows.Add(childen);
                    html += "<ul>";
                    showMenuTree(dcopy, (int)r["id"], jsfunction, viewUrl);
                    html += "</ul>";
                }
            }
            html += "</div>";
            return html;
        }

        public ActionResult Select()
        {
            try {
                var dataMenu = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM wmenu");
                return Content(showMenuTree(dataMenu));
            }
            catch(Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getLineHTML()}</div>"); }
        }

        public ActionResult Update(string id = "")
        {
            var timeStart = DateTime.Now;
            ViewBag.id = id;
            var objectid = Request.getValue("objectid");
            try
            {
                var mode = Request.getValue("mode");
                if (mode == "delete")
                {
                    if(Regex.IsMatch(objectid, @"^\d+$") == false) { throw new Exception($"ID menu không đúng {objectid}"); }
                    AppHelper.dbSqliteMain.Execute($"DELETE FROM wmenu WHERE id={objectid}");
                    /* Xóa tài khoản */
                    return Content($"<div class=\"alert alert-info\">Xóa menu có ID '{objectid}' thành công ({timeStart.getTimeRun()})</div>");
                }
                if (mode != "update")
                {
                    if (id != "")
                    {
                        DataTable items = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM wmenu WHERE id = @iduser LIMIT 1", new KeyValuePair<string, string>("@iduser", id));
                        if (items.Rows.Count == 0) { throw new Exception($"Tài khoản có tên đăng nhập '{id}' đã bị xoá hoặc không tồn tại trên hệ thống"); }
                        var data = new Dictionary<string, string>();
                        foreach (DataColumn c in items.Columns) { data.Add(c.ColumnName, items.Rows[0][c.ColumnName].ToString()); }
                        ViewBag.Data = data;
                    }
                    return View();
                }
                string where = "";
                var item = new Dictionary<string, string>
                {
                    { "mat_khau", Request.getValue("mat_khau").Trim() },
                    { "ten_hien_thi", Request.getValue("ten_hien_thi") },
                    { "gioi_tinh", Request.getValue("gioi_tinh") },
                    { "ngay_sinh", Request.getValue("ngay_sinh") },
                    { "email", Request.getValue("email") },
                    { "dien_thoai", Request.getValue("dien_thoai") },
                    { "dia_chi", Request.getValue("dia_chi") },
                    { "ghi_chu", Request.getValue("ghi_chu") },
                    { "hinh_dai_dien", "" }
                };
                if (idObject == "")
                {
                    item.Add("iduser", Request.getValue("iduser"));
                    if (item["iduser"] == "") { throw new Exception("Tên đăng nhập bỏ trống"); }
                    if (Regex.IsMatch(item["iduser"], "^[a-z0-9@_.]+$", RegexOptions.IgnoreCase) == false) { throw new Exception("Tên đăng nhập có các ký tự không thuộc [a-z0-9@_.] các từ cho phép"); }
                    if (item["mat_khau"] == "") { throw new Exception("Mật khẩu để trống"); }
                    item.Add("time_create", DateTime.Now.toTimestamp().ToString());
                    idObject = item["iduser"];
                }
                else { where = $"iduser = '{idObject.sqliteGetValueField()}'"; }
                /* Kiểm tra dữ liệu đầu vào */
                if (item["mat_khau"] != "") { item["mat_khau"] = item["mat_khau"].GetMd5Hash(); }
                else { item.Remove("mat_khau"); }
                if (item["ten_hien_thi"] == "") { throw new Exception("Tên hiển thị để trống"); }
                if (item["ngay_sinh"] == "") { throw new Exception("Ngày sinh để trống"); }

                AppHelper.dbSqliteMain.Update("taikhoan", item, where);
                return Content($"<div class=\"alert alert-info\">Thao tác thành công với tài khoản '{idObject}'</div>");
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getErrorSave()}</div>"); }
        }
    }
}