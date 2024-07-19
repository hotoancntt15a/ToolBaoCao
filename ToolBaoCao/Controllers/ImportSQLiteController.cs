using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class ImportSQLiteController : Controller
    {
        // GET: ImportSQLite
        public ActionResult Index()
        {
            ViewBag.Title = "Quản lý nhập dữ liệu từ SQLite(*.db, *.db3, *.bak, *.sqlite3, *.sqlite, *)";
            return View();
        }
        public ActionResult Update(string bieu, HttpPostedFileBase file)
        {
            ViewBag.data = "Đang thao tác";
            if (string.IsNullOrEmpty(bieu)) { ViewBag.Error = "Tham số biểu nhập không có chỉ định"; return View(); }
            if (file == null) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            if (file.ContentLength == 0) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            string fileName = Path.GetFileName(file.FileName);
            string fileExtension = Path.GetExtension(file.FileName);
            string fileNameSave = $"{bieu}{fileExtension}";
            file.SaveAs(Server.MapPath($"~/temp/excel/{fileNameSave}"));
            ViewBag.data = $"{bieu}: {fileName} size {file.ContentLength} b được lưu tại {fileNameSave}";
            return View();
        }
    }
}