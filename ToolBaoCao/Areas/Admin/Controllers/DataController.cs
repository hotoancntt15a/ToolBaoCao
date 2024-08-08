using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
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
                    if(Regex.IsMatch(tsql, "^select ", RegexOptions.IgnoreCase) == false)
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
    }
}