using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class TaiController : Controller
    {
        // GET: Tai
        public class DownloadController : Controller
        {
            public ActionResult Index()
            {
                var path = Request.getValue("object").Replace("\\", "/");
                try
                {
                    if (path == "") { throw new Exception("Không có tham số"); }
                    if (path.StartsWith("/") == false) { path = path.MD5Decrypt(); }
                    path = Server.MapPath("~" + path); if (System.IO.File.Exists(path)) { return File(path, "application/octet-stream", Path.GetFileName(path)); }
                    throw new Exception($"Không tìm thấy tập tin {Request.getValue("object")}");
                }
                catch (Exception ex) { return Content(ex.getLineHTML().BootstrapAlter("warning")); }
            }
        }
    }
}