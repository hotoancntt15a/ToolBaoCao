using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class ImportExcelController : Controller
    {
        // GET: ImportExcel
        public ActionResult Index()
        {
            ViewBag.Title = "Quản lý nhập dữ liệu Excel";
            return View();
        }
        public ActionResult Update()
        {
            var bieu = Request.getValue("bieu");
            if (string.IsNullOrEmpty(bieu)) { ViewBag.Error = "Tham số biểu nhập không có chỉ định"; return View(); }
            if(Request.Files.Count == 0) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            ViewBag.Info = $"{bieu}: {Request.Files[0].FileName} size {Request.Files[0].ContentLength} b";
            return View();
        }
    }
}