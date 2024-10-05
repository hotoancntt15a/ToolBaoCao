using Antlr.Runtime.Misc;
using ICSharpCode.SharpZipLib.GZip;
using NPOI.POIFS.FileSystem;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
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
            var folderSave = Path.Combine(AppHelper.pathApp, "temp", "xml", $"tinh{matinh}");
            var tmp = "";
            ViewBag.id = id;
            try
            {
                /* Xoá hết các File có trong thư mục */
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
                /* Đọc và kiểm tra các tập tin */
                var list = new List<string>() { ".db", ".zip" };
                var lfile = new List<string>();
                for (int i = 0; i < Request.Files.Count; i++)
                {
                    tmp = Path.GetExtension(Request.Files[i].FileName).ToLower();
                    if (tmp == ".zip")
                    {
                        lfile.Add($"{Request.Files[i].FileName} ({Request.Files[i].ContentLength.getFileSize()})");
                        var fzip = Path.Combine(folderTemp, "t" + Request.Files[i].FileName.GetMd5Hash() + ".zip");
                        Request.Files[i].SaveAs(fzip);
                        /* Giải nén tập tin */
                        using (ZipArchive archive = ZipFile.OpenRead(fzip))
                        {
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                if (entry.FullName.EndsWith(".db", StringComparison.OrdinalIgnoreCase))
                                {
                                    var fdb = Path.Combine(folderTemp, $"xml_{id}_{i}.db");
                                    entry.ExtractToFile(fdb, overwrite: true);
                                    var db = new dbSQLite(fdb);
                                    try
                                    {
                                        /* Kiểm tra có đúng cấu trúc dữ liệu không? */
                                        var data = db.getDataTable("SELECT MIN(KY_QT) AS X1, MAX(KY_QT) AS X2 FROM xml123");
                                        if (data.Rows.Count == 0) { continue; }
                                        tmp = $"{data.Rows[0][0]}";
                                        if (tmp == "" || tmp == "0") { continue; }
                                        db.Close();
                                        tmp = Path.Combine(folderSave, $"xml_{id}_{tmp}_{data.Rows[0][1]}");
                                        /* Xoá đi nếu tồn tại rồi */
                                        if (System.IO.File.Exists(tmp)) { System.IO.File.Delete(tmp); }
                                        /* Chuyển về thư mục chính */
                                        System.IO.File.Move(fdb, tmp);
                                    }
                                    catch (Exception exDB) { tmp = exDB.Message; continue; }
                                    db.Close();
                                }
                            }
                        }
                        continue;
                    }
                    if(tmp == ".db")
                    {
                        var fdb = Path.Combine(folderTemp, $"xml_{id}_{i}.db");
                        Request.Files[i].SaveAs(fdb);
                        var db = new dbSQLite(fdb);
                        try
                        {
                            /* Kiểm tra có đúng cấu trúc dữ liệu không? */
                            var data = db.getDataTable("SELECT MIN(KY_QT) AS X1, MAX(KY_QT) AS X2 FROM xml123");
                            if (data.Rows.Count == 0) { continue; }
                            tmp = $"{data.Rows[0][0]}";
                            if (tmp == "" || tmp == "0") { continue; }
                            db.Close();
                            tmp = Path.Combine(folderSave, $"xml_{id}_{tmp}_{data.Rows[0][1]}");
                            /* Xoá đi nếu tồn tại rồi */
                            if (System.IO.File.Exists(tmp) == false) { System.IO.File.Delete(tmp); }
                            /* Chuyển về thư mục chính */
                            System.IO.File.Move(fdb, tmp);
                        }
                        catch (Exception exDB) { tmp = exDB.Message; continue; }
                        db.Close();
                    }
                    if (list.Contains(tmp) == false) { continue; }
                    if (Path.GetExtension(Request.Files[i].FileName).ToLower() == ".db") { continue; }
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
        public ActionResult TruyVan()
        {
            var matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            /* Tài khoản system có thể xem được tất cả
             * Tài khoản admin tỉnh xem được toàn bộ của tỉnh được phân
             * Tải khoản người dùng chỉnh xem các báo cáo mình tạo ra
             */
            try
            {
                var mode = Request.getValue("mode");
                if (mode == "truyvan")
                {
                    var folderSave = Path.Combine(AppHelper.pathApp, "temp", "xml", $"tinh{matinh}");
                    var d = new DirectoryInfo(folderSave);
                    ViewBag.data = d.GetFiles().ToList();
                    return View();
                }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }
    }
}