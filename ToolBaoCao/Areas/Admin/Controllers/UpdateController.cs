using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class UpdateController : Controller
    {
        // GET: Admin/Update
        public ActionResult Index()
        {
            string mode = Request.getValue("mode");
            if (mode == "update")
            {
                try
                {
                    if (Request.Files.Count == 0) { throw new Exception("Không có tập tin nào đẩy lên"); }
                    if (Request.Files[0].FileName.ToLower().EndsWith(".zip") == false) { throw new Exception("Tập tin cập nhật không đúng định dạng nén (zip)"); }
                    string fileUpdate = Path.Combine(AppHelper.pathTemp, "hiatools.zip");
                    if (System.IO.File.Exists(fileUpdate))
                    {
                        try { System.IO.File.Delete(fileUpdate); } catch { }
                    }
                    if (System.IO.File.Exists(fileUpdate + ".txt"))
                    {
                        try { System.IO.File.Delete(fileUpdate + ".txt"); } catch { }
                    }
                    Request.Files[0].SaveAs(fileUpdate);
                    var lsfile = new List<string>();
                    using (ZipArchive archive = ZipFile.OpenRead(fileUpdate))
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            if (entry.Name == "") { continue; }
                            lsfile.Add(entry.FullName);
                            System.IO.File.AppendAllText(fileUpdate + ".txt", $"{entry.FullName}({entry.Length.getFileSize()}){Environment.NewLine}");
                        }
                    }
                    if (lsfile.Contains($"bin/{AppHelper.projectName}.dll") == false) { throw new Exception("Tập tin cập nhật không đúng"); }
                    /* Kiểm tra xem có đang nằm trong thư mực Developer không */
                    string tmp = Path.Combine(AppHelper.pathApp, "dbSQLite.cs");
                    if (System.IO.File.Exists(tmp)) { throw new Exception("Hệ thống đang chạy ở chế độ phát triển, không thể cập nhật"); }
                    AppHelper.Extract7z(fileUpdate, AppHelper.pathApp);
                }
                catch (Exception ex) { return Content(ex.getLineHTML()); }
                return Content("Đã đưa vào tiến trình nâng cấp. Vùi lòng chờ hoàn thành khoảng 5-15 phút tuỳ vào từng Server");
            }
            return View();
        }
    }
}