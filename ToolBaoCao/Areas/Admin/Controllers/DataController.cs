using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class DataController : ControllerCheckLogin
    {
        // GET: Admin/Data
        public ActionResult Index()
        {
            if ($"{Session["nhom"]}" != "0") { return RedirectToAction("Index", "Error", new { area = "", Message = keyMSG.NotAccessControl }); }
            var timeStart = DateTime.Now;
            var mode = Request.getValue("mode");
            try
            {
                if (mode == "tsql")
                {
                    string dataName = Request.getValue("data");
                    string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", dataName);
                    string tsql = Request.getValue("tsql").Trim();
                    if (tsql == "") { return Content($"<div class=\"alert alert-warning\">TSQL bỏ trống</div>"); }
                    var db = new dbSQLite(pathDB);
                    if (Regex.IsMatch(tsql, "^select ", RegexOptions.IgnoreCase) == false)
                    {
                        var rs = db.Execute(tsql);
                        return Content($"<div class=\"alert alert-info\">Data {dataName}; TSQL: {tsql}<br />Thao tác thành công {rs} ({timeStart.getTimeRun()})</div>");
                    }
                    var data = db.getDataTable(tsql);
                    ViewBag.content = $"Data {dataName}; TSQL: {tsql}";
                    ViewBag.data = data;
                }
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">Lỗi: {ex.getLineHTML()}</div>"); }
            return View();
        }

        public ActionResult Cache()
        {
            string mode = Request.getValue("mode");
            DateTime timeStart = DateTime.Now;
            try
            {
                string pathCache = Path.Combine(AppHelper.pathApp, "cache");
                if (mode == "del")
                {
                    if (Request.getValue("listfile") == "") { throw new Exception("Bạn chưa chọn tập tin nào để thao tác"); }
                    var files = Request.Form.GetValues("listfile").ToList();
                    var list = new List<string>();
                    try { foreach (var file in files) { System.IO.File.Delete(Path.Combine(pathCache, file)); list.Add(file); } }
                    catch { }
                    return Content($"<div class=\"alert alert-info\">Xoá tất cả các tập tin đã chọn thành công ({timeStart.getTimeRun()}){(list.Count > 10 ? "" : "<br />" + string.Join("<br />", list))}</div>");
                }
                if (mode == "clear")
                {
                    var d = new System.IO.DirectoryInfo(pathCache);
                    try { foreach (var file in d.GetFiles()) { file.Delete(); } }
                    catch { }
                    return Content($"<div class=\"alert alert-info\">Xoá tất cả các tập tin thành công ({timeStart.getTimeRun()})</div>");
                }
                if (mode != "") { throw new Exception($"Tham số '{mode}' không hỗ trợ"); }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }
    }
}