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
                    string fileExecute = Path.Combine(AppHelper.pathApp, "7z.exe");
                    if (System.IO.File.Exists(fileExecute) == false) { throw new Exception("Chương trình thực thi hệ thống không tồn tại (7z.exe)"); }
                    if (System.IO.File.Exists(Path.Combine(AppHelper.pathApp, "7z.dll")) == false) { throw new Exception("Thư viện chương trình thực thi hệ thống không tồn tại (7z.dll)"); }
                }
                catch (Exception ex)
                {
                    ViewBag.Error = ex.getLineHTML();
                    return View();
                }
                return Content("Thao tacs thanhf coong");
            }
            return View();
        }
    }
}