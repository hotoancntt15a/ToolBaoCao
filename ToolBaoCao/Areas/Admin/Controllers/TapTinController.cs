using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class TapTinController : Controller
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
                var listFolder = new List<string>();
                var listfile = new List<string>();
                /* Tham số đường dẫn */
                var pathFolder = Request.getValue("path");
                if(pathFolder != "") { pathFolder = Regex.Replace(pathFolder, @"^\\+", ""); }
                /* Từ khoá */
                var key = Request.getValue("key");
                /* Thư mục cần truy vấn */
                var d = new System.IO.DirectoryInfo(pathFolder == "" ? AppHelper.pathApp : Path.Combine(AppHelper.pathApp, pathFolder));
                if(key == "")
                {
                    return View();
                }

                int len = 0;
                var d = new System.IO.DirectoryInfo(AppHelper.pathApp);
                if (d.Exists == false) { throw new Exception($"Thư mục Ứng dụng không có quyền truy cập '{AppHelper.pathApp}'"); }
                len = AppHelper.pathApp.Length; ViewBag.len = len;
                var tmp = Request.getValue("p").Replace("/", @"\");
                ViewBag.path = tmp;
                d = new System.IO.DirectoryInfo(Path.Combine(AppHelper.pathApp, Regex.Replace(tmp, @"^\\+", "")));
                if (d.Exists == false) { return Content($"Thư mục: {tmp} không tồn tại trên hệ thống".BootstrapAlter("warning")); }
                tmp = Request.getValue("key");
                if (tmp == "")
                {
                    ViewBag.folders = d.GetDirectories().OrderBy(p => p.Name).ToList();
                    ViewBag.files = d.GetFiles().OrderBy(p => p.Name).ToList();
                    var ls = new List<string>();
                    var s = d.FullName.Substring(len - 1).Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                    tmp = "";
                    ls.Add($"<a href=\"javascript:viewfolders('{link}')\"> .. </a>");
                    foreach (var v in s)
                    {
                        tmp = $"{tmp}/{v}";
                        ls.Add($"<a href=\"javascript:viewfolders('{link}?p={Server.UrlPathEncode(tmp)}')\"> {v} </a>");
                    }
                    ViewBag.vitri = string.Join(" \\ ", ls);
                    return View();
                }
                ViewBag.folders = d.GetDirectories(tmp, System.IO.SearchOption.AllDirectories).OrderBy(p => p.Name).ToList();
                ViewBag.files = d.GetFiles(tmp, System.IO.SearchOption.AllDirectories).OrderBy(p => p.Name).ToList();
                ViewBag.vitri = $"Tìm kiếm thư mục/tập tin: <a href=\"javascript:viewfolders('{link}?p={Server.UrlPathEncode(d.FullName.Substring(len - 2)).Replace("\\", "/")}')\"> {d.FullName.Substring(len - 2)} </a>";
            }
            catch (Exception ex) { return Content(ex.getLineHTML()); }
            return View();
        }
    }
}