using System;
using System.Collections.Generic;
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
        public ActionResult Update()
        {
            var bieu = Request.getValue("bieu");
            if (string.IsNullOrEmpty(bieu)) { ViewBag.Error = "Tham số biểu nhập không có chỉ định"; return View(); }
            if (Request.Files.Count == 0) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            Request.Files[0].SaveAs(Server.MapPath($"~/temp/{bieu}.db"));
            ViewBag.Info = $"{bieu}: {Request.Files[0].FileName} size {Request.Files[0].ContentLength} b";
            return View();
        }
    }
}