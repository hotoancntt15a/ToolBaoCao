using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
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
                var data = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM wmenu");
                ViewBag.Data = data;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getErrorSave()}"; }
            return View();
        }

        private string showMenuTree(DataTable dataMenu, long idMenuFather = 0, string jsfunction = "selectMenu", bool viewUrl = true)
        {
            var li = new List<string>();
            if (idMenuFather == 0)
            {
                string html = $"<div class=\"viewmenutree\"><li> <a href=\"javascript:{jsfunction}(this,'0');\"> -- NEW MENU GROUP -- </a></li>";
                if (dataMenu.Rows.Count > 0)
                {
                    li.Add("<ul>");
                    foreach (DataRow r in dataMenu.Rows)
                    {
                        var note = $"{r["note"]}".Trim();
                        li.Add($"<li> <a href=\"javascript:{jsfunction}(this,'{r["id"]}');\"> <i class=\"{r["css"]}\"></i> {r["title"]}</a>");
                        showMenuTree(dataMenu, (long)r["id"], jsfunction, viewUrl);
                        li.Add("</li>");
                    }
                    li.Add("</ul>");
                    html += string.Join("", li);
                }
                html += "</div>";
                return html;
            }
            var dcopy = dataMenu.AsEnumerable().Where(r => r.Field<long>("idfather") == 0).OrderBy(r => r.Field<long>("postion")).ToList();
            if (dcopy.Count > 0)
            {
                li.Add("<ul>");
                foreach (DataRow r in dcopy)
                {
                    var note = $"{r["note"]}".Trim();
                    li.Add($"<li> <a href=\"javascript:{jsfunction}(this,'{r["id"]}');\"> <i class=\"{r["css"]}\"></i> {r["title"]}</a>");
                    showMenuTree(dataMenu, (long)r["id"], jsfunction, viewUrl);
                    li.Add("</li>");
                }
                li.Add("</ul>");
                return string.Join("", li);
            }
            return "";
        }

        public ActionResult Select()
        {
            try
            {
                var dataMenu = AppHelper.dbSqliteMain.getDataTable("SELECT * FROM wmenu");
                return Content(showMenuTree(dataMenu));
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getLineHTML()}</div>"); }
        }

        public ActionResult Update(string id = "")
        {
            var timeStart = DateTime.Now;
            ViewBag.id = id;
            try
            {
                if (id != "") { if (Regex.IsMatch(id, @"^\d+$") == false) { throw new Exception($"ID menu không đúng {id}"); } }
                var mode = Request.getValue("mode");
                if (mode == "delete")
                {
                    return Content($"<div class=\"alert alert-info\">Bạn có thực sự có muốn xoá Menu có ID '{id}' không?</div>");
                }
                if (mode == "forcedel")
                {
                    AppHelper.dbSqliteMain.Execute($"DELETE FROM wmenu WHERE id={id}"); /* Xóa tài khoản */
                    return Content($"<div class=\"alert alert-info\">Xóa menu có ID '{id}' thành công ({timeStart.getTimeRun()})</div>");
                }
                if (mode != "update")
                {
                    if (id != "")
                    {
                        /* Lấy thông tin menu cần sửa */
                        var items = AppHelper.dbSqliteMain.getDataTable($"SELECT * FROM wmenu WHERE id = {id}");
                        if (items.Rows.Count == 0) { throw new Exception($"Menu có ID '{id}' không tồn tại hoặc bị xoá trong hệ thống"); }
                        var data = new Dictionary<string, string>();
                        foreach (DataColumn c in items.Columns) { data.Add(c.ColumnName, items.Rows[0][c.ColumnName].ToString()); }
                        ViewBag.Data = data;
                    }
                    return View();
                }
                string where = id == "" ? "" : $"id={id}";
                var item = new Dictionary<string, string>
                {
                    { "title", Request.getValue("title").Trim() },
                    { "link", Request.getValue("link").Trim() },
                    { "idfather", Request.getValue("idfather") },
                    { "paths", Request.getValue("paths").Trim() },
                    { "postion", Request.getValue("postion") },
                    { "note", Request.getValue("note").Trim() },
                    { "css", Request.getValue("css").Trim() }
                };
                AppHelper.dbSqliteMain.Update("wmenu", item, where);
                where = where == "" ? "Thêm mới thành công " : $"Thay đổi thành công menu có ID '{id}'";
                return Content($"<div class=\"alert alert-info\">{where} ({timeStart.getTimeRun()})</div>");
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getErrorSave()}</div>"); }
        }
    }
}