using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class HomeController : ControllerCheckLogin
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult GetMessages() {
            return View();
        }
        public ActionResult TaiKhoan()
        {
            var id = $"{Session["iduser"]}";
            if (id == "") { return Content($"<div class=\"alert alert-warning\">Bạn chưa đăng nhập hoặc phiên làm việc của bạn đã hết hạn</div>"); }
            string mode = Request.getValue("mode");
            try
            {
                if (mode == "update")
                {
                    var item = new Dictionary<string, string>
                    {
                        { "mat_khau", Request.getValue("mat_khau").Trim() },
                        { "ten_hien_thi", Request.getValue("ten_hien_thi") },
                        { "email", Request.getValue("email") },
                        { "dien_thoai", Request.getValue("dien_thoai") },
                        { "vitrilamviec", Request.getValue("vitrilamviec") }
                    };
                    if (item["mat_khau"] == "") { item.Remove("mat_khau"); } 
                    else { item["mat_khau"] = item["mat_khau"].GetMd5Hash(); }
                    if (item["ten_hien_thi"] == "") { return Content($"<div class=\"alert alert-warning\">Tên hiển thị để trống</div>"); }

                    AppHelper.dbSqliteMain.Update("taikhoan", item, $"iduser='{id.sqliteGetValueField()}'");
                    return Content($"<div class=\"alert alert-info\">Thay đổi thông tin thành công</div>");
                }
                var items = AppHelper.dbSqliteMain.getDataTable($"SELECT * FROM taikhoan WHERE iduser = '{id.sqliteGetValueField()}' LIMIT 1");
                if (items.Rows.Count == 0)
                {
                    return Content($"<div class=\"alert alert-warning\">Tài khoản có tên đăng nhập '{id}' đã bị xoá hoặc không tồn tại trên hệ thống</div>");
                }
                var data = new Dictionary<string, string>();
                foreach (DataColumn c in items.Columns) { data.Add(c.ColumnName, items.Rows[0][c.ColumnName].ToString()); }
                ViewBag.dmTinh = AppHelper.dbSqliteMain.getDataTable($"SELECT id, ten FROM dmTinh WHERE id='{data.getValue("idtinh").sqliteGetValueField()}' ORDER BY tt, ten");
                ViewBag.Data = data;
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getLineHTML()}</div>"); }
            return View();
        }
    }
}