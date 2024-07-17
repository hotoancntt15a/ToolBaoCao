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
    }
}