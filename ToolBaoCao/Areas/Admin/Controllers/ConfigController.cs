using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class ConfigController : ControllerCheckLogin
    {
        // GET: Admin/Config
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult save()
        {
            string key = Request.getValue("key");
            string value = Request.getValue("value");
            if (key == "") { return Content(keyMSG.NotVariable); }
            if(value == "") { AppHelper.appConfig.Remove(key); }
            else { AppHelper.appConfig.Set(key, value); }            
            return Content("Lưu thành công".BootstrapAlter());
        }

        public ActionResult Variables()
        {
            return View();
        }
        public ActionResult views()
        {
            return View();
        }
    }
}