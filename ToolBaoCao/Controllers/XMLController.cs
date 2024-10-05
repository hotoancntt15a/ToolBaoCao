using System;
using System.IO;
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
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }
    }
}