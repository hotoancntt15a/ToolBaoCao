using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class NhomController : Controller
    {
        // GET: Admin/Nhom
        public ActionResult Index()
        {
            try
            {
                var data = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM dmnhom");
                ViewBag.Data = data;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getErrorSave()}"; }
            return View();
        }

        public ActionResult Update(string id = "")
        {
            var timeStart = DateTime.Now;
            ViewBag.id = id;
            try
            {
                if (id != "") { if (Regex.IsMatch(id, @"^\d+$") == false) { throw new Exception($"ID nhóm không đúng {id}"); } }
                var mode = Request.getValue("mode");
                if (mode == "delete")
                {
                    return Content($"<div class=\"alert alert-info\">Bạn có thực sự có muốn xoá Nhóm có ID '{id}' không? <br /><a href=\"javascript:postform('', '/Admin/Nhom/Update?id={id}&layout=null&mode=forcedel');\" class=\"btn btn-primary btn-sm\"> Có </a></div>");
                }
                if (mode == "forcedel")
                {
                    AppHelper.dbSqliteMain.Execute($"DELETE FROM dmnhom WHERE id={id}"); /* Xóa tài khoản */
                    return Content($"<div class=\"alert alert-info\">Xóa Nhóm có ID '{id}' thành công ({timeStart.getTimeRun()})</div>");
                }
                if (mode != "update")
                {
                    if (id != "")
                    {
                        /* Lấy thông tin nhóm cần sửa */
                        var items = AppHelper.dbSqliteMain.getDataTable($"SELECT * FROM dmnhom WHERE id = {id}");
                        if (items.Rows.Count == 0) { throw new Exception($"Nhóm có ID '{id}' không tồn tại hoặc bị xoá trong hệ thống"); }
                        var data = new Dictionary<string, string>();
                        foreach (DataColumn c in items.Columns) { data.Add(c.ColumnName, items.Rows[0][c.ColumnName].ToString()); }
                        ViewBag.Data = data;
                    }
                    return View();
                }
                string objectid = Request.getValue("objectid");
                if (objectid != "") { if (Regex.IsMatch(objectid, @"^\d+$") == false) { throw new Exception($"ID nhóm không đúng định dạng {objectid}"); } }
                var item = new Dictionary<string, string>
                {
                    { "id", objectid },
                    { "ten", Request.getValue("ten").Trim() },
                    { "idwmenu", Request.getValue("idwmenu") },
                    { "ghichu", Request.getValue("ghichu").Trim() }
                };
                AppHelper.dbSqliteMain.Update("dmnhom", item, "replace");
                return Content($"<div class=\"alert alert-info\">Cập nhật thành công Nhóm có ID '{objectid}' ({timeStart.getTimeRun()})</div>");
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getErrorSave()}</div>"); }
        }
    }
}