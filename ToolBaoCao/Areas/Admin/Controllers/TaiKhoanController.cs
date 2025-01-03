﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class TaiKhoanController : ControllerCheckLogin
    {
        // GET: Admin/TaiKhoan
        public ActionResult Index()
        {
            try
            {
                /* Danh sách nhóm */
                var data = AppHelper.dbSqliteMain.getDataTable("SELECT id, ten FROM dmnhom");
                var dsNhom = new Dictionary<string, string>();
                foreach (DataRow dr in data.Rows) { dsNhom.Add(dr[0].ToString(), dr[1].ToString()); }
                ViewBag.dsnhom = dsNhom;
                data = AppHelper.dbSqliteMain.getDataTable("SELECT tk.*, datetime(tk.time_create, 'auto', '+7 hour') as timecreate, IFNULL(p2.timelogin, 0) as timelogin FROM taikhoan tk LEFT JOIN logintime p2 ON tk.iduser=p2.iduser ORDER BY tk.iduser");
                ViewBag.Data = data;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getErrorSave()}"; }
            return View();
        }

        // GET: Admin/TaiKhoan/Update
        public ActionResult Update(string id = "")
        {
            var timeStart = DateTime.Now;
            var tmp = $"{Session["nhom"]}";
            if (tmp != "0" && tmp != "1") { return Content($"<div class=\"alert alert-warning\">Tài khoản bạn không có quyền khóa tài khoản</div>"); }
            ViewBag.id = id;
            var idObject = Request.getValue("idobject");
            try
            {
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
                    ViewBag.dmTinh = AppHelper.dbSqliteMain.getDataTable("SELECT id, ten FROM dmTinh ORDER BY tt, ten"); ;
                    ViewBag.dmNhom = AppHelper.dbSqliteMain.getDataTable("SELECT id, ten FROM dmNhom ORDER BY id");
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
                    { "nhom", Request.getValue("nhom") },
                    { "email", Request.getValue("email") },
                    { "dien_thoai", Request.getValue("dien_thoai") },
                    { "idtinh", Request.getValue("idtinh") },
                    { "vitrilamviec", Request.getValue("vitrilamviec") }
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
                if (Regex.IsMatch(item["nhom"], @"^\d+$") == false) { throw new Exception($"Mã nhóm làm việc không hợp lệ {item["nhom"]}"); }

                AppHelper.dbSqliteMain.Update("taikhoan", item, where);
                return Content($"<div class=\"alert alert-info\">Thao tác thành công với tài khoản '{idObject}' ({timeStart.getTimeRun()}) </div>");
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getLineHTML()}</div>"); }
        }
    }
}