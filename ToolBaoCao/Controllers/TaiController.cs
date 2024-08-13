using System;
using System.IO;
using System.Text.RegularExpressions;
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
                /* Không hỗ trợ \\{PC Share}\ Nếu có yêu cầu Map Network Drive */
                path = path.Replace("/", @"\");
                path = Path.Combine(AppHelper.pathApp, Regex.Replace(path, @"^\\+", ""));
                if (System.IO.File.Exists(path))
                {
                    return File(path, "application/octet-stream", Path.GetFileName(path));
                }
                return Content($"Không tìm thấy tập tin {path} với Object='{Request.getValue("object")}'; MD5Check: {p2}".BootstrapAlter("warning"));
            }
            catch (Exception ex) { return Content(ex.getErrorSave().BootstrapAlter("warning")); }
        }
    }
}