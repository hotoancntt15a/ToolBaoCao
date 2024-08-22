using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class TaiKhoanController : ControllerCheckLogin
    {
        /* GET: QuanTri */

        public ActionResult Index()
        {
            var tmp = $"{Session["nhom"]}";
            try
            {
                if(tmp != "0" && tmp != "1") { ViewBag.Error = "Bạn không có quyền sử dụng tính năng này"; }
                var idtinh = $"{Session["idtinh"]}".sqliteGetValueField();
                var iduser = $"{Session["iduser"]}".sqliteGetValueField();
                var tsql = $"SELECT tk.*, datetime(tk.time_create, 'auto', '+7 hour') as timecreate, IFNULL(p2.timelogin, 0) as timelogin FROM taikhoan tk LEFT JOIN logintime p2 ON tk.iduser=p2.iduser WHERE tk.idtinh='{idtinh}' AND tk.iduser <> '{iduser}' ORDER BY tk.iduser";
                ViewBag.tsql = tsql;
                var data = AppHelper.dbSqliteMain.getDataTable(tsql);
                ViewBag.Data = data;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getLineHTML()}"; }
            return View();
        }

        public ActionResult Update(string id = "")
        {
            var timeStart = DateTime.Now;
            var tmp = $"{Session["nhom"]}";
            if (tmp != "0" && tmp != "1") { return Content($"Tài khoản bạn không có quyền khóa tài khoản".BootstrapAlter("warning")); }
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
                    { "email", Request.getValue("email") },
                    { "dien_thoai", Request.getValue("dien_thoai") },
                    { "idtinh", idtinh },
                    { "vitrilamviec", Request.getValue("vitrilamviec") }
                };
                if (idObject == "")
                {
                    item.Add("iduser", Request.getValue("iduser"));
                    if (item["iduser"] == "") { throw new Exception("Tên đăng nhập bỏ trống"); }
                    if (Regex.IsMatch(item["iduser"], "^[a-z0-9@_.]+$", RegexOptions.IgnoreCase) == false) { throw new Exception("Tên đăng nhập có các ký tự không thuộc [a-zA-Z0-9@_.] các từ cho phép"); }
                    if (item["mat_khau"] == "") { throw new Exception("Mật khẩu để trống"); }
                    item.Add("time_create", DateTime.Now.toTimestamp().ToString());
                    item.Add("nhom", "3"); /* Mặc định 3 - Nhóm người sử dụng */
                    idObject = item["iduser"];
                }
                else { where = $"iduser = '{idObject.sqliteGetValueField()}'"; }
                /* Kiểm tra dữ liệu đầu vào */
                if (item["mat_khau"] != "") { item["mat_khau"] = item["mat_khau"].GetMd5Hash(); }
                else { item.Remove("mat_khau"); }
                if (item["ten_hien_thi"] == "") { throw new Exception("Tên hiển thị để trống"); }

                AppHelper.dbSqliteMain.Update("taikhoan", item, where);
                return Content($"<div class=\"alert alert-info\">Thao tác thành công với tài khoản '{idObject}' ({timeStart.getTimeRun()})</div>");
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getLineHTML()}</div>"); }
        }
    }
}