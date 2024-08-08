using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class TaskController : ControllerCheckLogin
    {
        // GET: Task
        public ActionResult Index()
        {
            return View();
        }
    }
}