using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class TaiController : Controller
    {
        // GET: Tai
        public ActionResult Index()
        {
            var path = Request.getValue("object");
            try
            {
                if (path == "") { throw new Exception("Không có tham số"); }
                var p2 = path.MD5Decrypt(); if (p2.StartsWith("Lỗi:") == false) { path = p2; }
                path = path.Replace("/", @"\");
                if (path.StartsWith(@"\") == false) { path = $@"\{path}"; }
                path = Path.Combine(AppHelper.pathApp, Regex.Replace(path, @"^[\\]+", ""));
                if (System.IO.File.Exists(path))
                {
                    return File(path, "application/octet-stream", Path.GetFileName(path));
                }
                throw new Exception($"Không tìm thấy tập tin {AppHelper.pathApp} : {path}");
            }
            catch (Exception ex) { return Content(ex.getLineHTML().BootstrapAlter("warning")); }
        }
    }
}