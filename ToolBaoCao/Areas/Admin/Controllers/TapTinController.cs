using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
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
        public ActionResult views()
        {
            try
            {
                var listFolder = new List<DirectoryInfo>();
                /* Tham số đường dẫn */
                var pathFolder = Request.getValue("path");
                var folders = new List<string>();
                if (pathFolder != "") { 
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
                if(key == "")
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