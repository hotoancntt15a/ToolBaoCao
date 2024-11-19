using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class TapTinController : ControllerCheckLogin
    {
        // GET: Admin/TapTin
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ReadText()
        {
            var path = Request.getValue("path");
            try
            {
                if (path == "") { throw new Exception("Không có tham số"); }
                var p2 = path.MD5Decrypt(); if (p2.StartsWith("Lỗi:") == false) { path = p2; }
                /* Không hỗ trợ \\{PC Share}\ Nếu có yêu cầu Map Network Drive */
                path = path.Replace("/", @"\");
                path = Path.Combine(AppHelper.pathApp, Regex.Replace(path, @"^\\+", ""));
                if (System.IO.File.Exists(path))
                {
                    return Content(System.IO.File.ReadAllText(path).Replace(Environment.NewLine, "<br />"));
                }
                return Content($"Không tìm thấy tập tin {path} với Path='{Request.getValue("path")}'; MD5Check: {p2}".BootstrapAlter("warning"));
            }
            catch (Exception ex) { return Content(ex.getErrorSave().BootstrapAlter("warning")); }
        }

        public ActionResult views()
        {
            try
            {
                var listFolder = new List<DirectoryInfo>();
                /* Tham số đường dẫn */
                var pathFolder = Request.getValue("path");
                var folders = new List<string>();
                if (pathFolder != "")
                {
                    pathFolder = Regex.Replace(pathFolder, @"^\\+", "");
                    /* Bỏ qua thư mục ẩn */
                    folders = pathFolder.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
                ViewBag.folders = folders;
                ViewBag.path = pathFolder;
                /* Từ khoá */
                var key = Request.getValue("key");
                ViewBag.key = key;
                /* Thư mục cần truy vấn */
                var dir = new System.IO.DirectoryInfo(pathFolder == "" ? AppHelper.pathApp : Path.Combine(AppHelper.pathApp, pathFolder));
                if (key == "")
                {
                    ViewBag.listfolder = dir.GetDirectories().ToList();
                    ViewBag.listfile = dir.GetFiles().ToList();
                    return View();
                }
                ViewBag.listfolder = dir.GetDirectories(key, SearchOption.AllDirectories).ToList();
                ViewBag.listfile = dir.GetFiles(key, SearchOption.AllDirectories).ToList();
            }
            catch (Exception ex) { return Content(ex.getLineHTML()); }
            return View();
        }
    }
}