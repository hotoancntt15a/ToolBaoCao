using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class ImportSQLiteController : ControllerCheckLogin
    {
        // GET: ImportSQLite
        public ActionResult Index()
        {
            ViewBag.Title = "Quản lý nhập dữ liệu từ SQLite(*.db, *.db3, *.bak, *.sqlite3, *.sqlite, *)";
            return View();
        }
        public ActionResult Update(string bieu, HttpPostedFileBase inputfile)
        {
            ViewBag.data = "Đang thao tác";
            if (string.IsNullOrEmpty(bieu)) { ViewBag.Error = "Tham số biểu nhập không có chỉ định"; return View(); }
            if (inputfile == null) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            if (inputfile.ContentLength == 0) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            string fileName = Path.GetFileName(inputfile.FileName);
            string fileExtension = Path.GetExtension(inputfile.FileName);
            string fileNameSave = $"{bieu}{fileExtension}";
            inputfile.SaveAs(Server.MapPath($"~/temp/excel/{fileNameSave}"));
            ViewBag.data = $"{bieu}: {fileName} size {inputfile.ContentLength} b được lưu tại {fileNameSave}";
            return View();
        }
    }
}