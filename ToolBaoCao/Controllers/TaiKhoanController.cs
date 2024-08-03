using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class TaiKhoanController : Controller
    {
        /* GET: QuanTri */

        public ActionResult Index()
        {
            try
            {
                string idtinh = $"{Session["idtinh"]}";
                var data = AppHelper.dbSqliteMain.getDataTable($"SELECT * FROM taikhoan WHERE idtinh='{idtinh.sqliteGetValueField()}' ORDER BY iduser");
                ViewBag.Data = data;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getErrorSave()}"; }
            return View();
        }

        public ActionResult Update(string id = "")
        {
            var tmp = $"{Session["nhom"]}";
            if (tmp != "0" && tmp != "1") { return Content($"<div class=\"alert alert-warning\">Tài khoản bạn không có quyền khóa tài khoản</div>"); }
            ViewBag.id = id;
            var idObject = Request.getValue("idobject");
            try
            {
                string idtinh = $"{Session["idtinh"]}";
                var mode = Request.getValue("mode");
                if (mode == "delete")
                {
                    /* Kiểm tra tài khoản đã sử dụng chưa, Nếu đã sử dụng thì không thể xóa */
                    if (idObject == "") { throw new Exception("Tham số tài khoản không đúng"); }
                    var listAccoutNotAccess = new List<string>() { "admin", "administrator", "system" };
                    if (listAccoutNotAccess.Contains(idObject.ToLower())) { throw new Exception("Tài khoản có tên đăng nhập đặc biệt không thể khóa"); }
                    AppHelper.dbSqliteMain.Execute("UPDATE taikhoan SET locked=1 WHERE iduser=@iduser", new KeyValuePair<string, string>("@iduser", idObject));
                    /* Xóa tài khoản */
                    return Content($"<div class=\"alert alert-info\">Khóa tài khoản {idObject} thành công</div>");
                }
                if (mode != "update")
                {
                    ViewBag.dmTinh = AppHelper.dbSqliteMain.getDataTable($"SELECT id, ten FROM dmTinh WHERE id='{idtinh.sqliteGetValueField()}' ORDER BY tt, ten");
                    if (id != "")
                    {
                        var items = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM taikhoan WHERE iduser = @iduser LIMIT 1", new KeyValuePair<string, string>("@iduser", id));
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
                    { "idtinh", idtinh },
                    { "ghi_chu", Request.getValue("ghi_chu") },
                    { "hinh_dai_dien", "" }
                };
                if (idObject == "")
                {
                    item.Add("iduser", Request.getValue("iduser"));
                    if (item["iduser"] == "") { throw new Exception("Tên đăng nhập bỏ trống"); }
                    if (Regex.IsMatch(item["iduser"], "^[a-z0-9@_.]+$", RegexOptions.IgnoreCase) == false) { throw new Exception("Tên đăng nhập có các ký tự không thuộc [a-zA-Z0-9@_.] các từ cho phép"); }
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
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getLineHTML()}</div>"); }
        }
    }
}