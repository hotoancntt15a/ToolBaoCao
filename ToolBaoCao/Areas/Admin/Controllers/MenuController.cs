using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class MenuController : ControllerCheckLogin
    {
        /* GET: Admin/Menu */

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

        private string showMenuTree(DataTable dataMenu, long idMenuFather = 0, string jsfunction = "selectMenu", bool showTree = true)
        {
            var li = new List<string>();
            var dcopy = dataMenu.AsEnumerable().Where(r => r.Field<long>("idfather") == idMenuFather).OrderBy(r => r.Field<long>("postion")).ToList();
            if (dcopy.Count > 0)
            {
                li.Add($"<ul>");
                if (idMenuFather == 0)
                {
                    li.Add($"<li> <a href=\"javascript:{jsfunction}(this,'0');\"> -- NEW MENU GROUP -- </a></li>");
                }
                var dt = dataMenu.Clone();
                foreach (DataRow r in dcopy) { dt.ImportRow(r); }
                foreach (DataRow r in dt.Rows)
                {
                    var link = $"{r["link"]}".Trim(); if (link != "") { link = $" ({link})"; }
                    li.Add($"<li> <a href=\"javascript:{jsfunction}(this,'{r["id"]}');\" title=\"{r["note"]}\"> <i class=\"{r["css"]}\"></i> {r["title"]}{link}</a>");
                    if (showTree) { li.Add(showMenuTree(dataMenu, (long)r["id"], jsfunction, showTree)); }
                    li.Add("</li>");
                }
                li.Add("</ul>");
            }
            return string.Join("", li);
        }

        public ActionResult Select()
        {
            try
            {
                string idfather = Request.getValue("father");
                if (idfather != "") { if (Regex.IsMatch(idfather, @"^\d+$") == false) { return Content($"<div class=\"alert alert-warning\">Tham số menu cha '{idfather}' không đứng</div>"); } }
                string showtree = Request.getValue("showtree");
                if (Regex.IsMatch(idfather, @"^\d+$") == false) { showtree = "1"; }
                string where = idfather == "" ? "" : $"WHERE idfather={idfather}";
                var dataMenu = AppHelper.dbSqliteMain.getDataTable($"SELECT * FROM wmenu {where}");
                return Content("<div class=\"viewmenutree\">" + showMenuTree(dataMenu, long.Parse(idfather == "" ? "0" : idfather), showTree: showtree == "1") + "</div>");
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">{ex.getLineHTML()}</div>"); }
        }

        public ActionResult Update(string id = "")
        {
            var timeStart = DateTime.Now;
            var mode = Request.getValue("mode");
            ViewBag.mode = mode;
            ViewBag.id = id;
            try
            {
                if (id != "") { if (Regex.IsMatch(id, @"^\d+$") == false) { throw new Exception($"ID menu không đúng {id}"); } }
                if (mode == "delete")
                {
                    return View();
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