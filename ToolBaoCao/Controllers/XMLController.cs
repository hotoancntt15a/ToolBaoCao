using Antlr.Runtime.Misc;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class XMLController : ControllerCheckLogin
    {
        // GET: XML
        public ActionResult Index()
        {
            try
            {
                var d = new DirectoryInfo(Path.Combine(AppHelper.pathAppData, "xml"));
                if (d.Exists == false) { d.Create(); }

                d = new DirectoryInfo(Path.Combine(AppHelper.pathTemp, "xml"));
                if (d.Exists == false) { d.Create(); }

                if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }
        public ActionResult Buoc1()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            /* Tạo thư mục tạm */
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "xml", $"{Session["idtinh"]}_{Session["iduser"]}".GetMd5Hash());
            var d = new System.IO.DirectoryInfo(folderTemp);
            if (d.Exists == false) { d.Create(); }
            return View();
        }
        public ActionResult Buoc2()
        {
            var timeStart = DateTime.Now;
            var idUser = $"{Session["iduser"]}";
            var matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            if (Request.Files.Count == 0) { ViewBag.Error = "Không có tập tin dữ liệu nào đẩy lên"; return View(); }
            var id = $"{timeStart:yyMMddHHmmss}_{matinh}_{timeStart.Millisecond:000}";
            var timeUp = timeStart.toTimestamp().ToString();
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "xml", $"{matinh}_{Session["iduser"]}".GetMd5Hash());
            var tmp = "";
            ViewBag.id = id;
            try
            {
                /* Xoá hết các File có trong thư mục */
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
                /* Khai báo dữ liệu tạm */
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                dbTemp.CreateImportBcThang();
                dbTemp.CreatePhucLucBcThang();
                dbTemp.CreateBcThang();
                /* Đọc và kiểm tra các tập tin */
                var list = new List<string>() { ".db", ".zip" };
                var lfile = new List<string>();
                for (int i = 0; i < Request.Files.Count; i++)
                {
                    tmp = Path.GetExtension(Request.Files[i].FileName).ToLower();
                    if(tmp ==".zip")
                    {
                        
                    }
                    if(list.Contains(tmp) == false) { continue; }
                    if (Path.GetExtension(Request.Files[i].FileName).ToLower() == ".db") { continue; }
                    lfile.Add($"{Request.Files[i].FileName} ({Request.Files[i].ContentLength.getFileSize()})");
                }
                ViewBag.files = lfile;
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.getLineHTML();
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
            }
            return View();
        }
    }
}