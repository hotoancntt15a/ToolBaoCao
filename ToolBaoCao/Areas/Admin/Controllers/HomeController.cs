using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Areas.Admin.Controllers
{
    public class HomeController : ControllerCheckLogin
    {
        // GET: Admin/Home
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult GetMessages()
        {
            return View();
        }
    }
}