using System;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Index()
        {
            var mode = Request.getValue("mode");
            if (mode == "login")
            {
                var remember = Request.getValue("remember");
                var msg = AppHelper.setLogin(Request.getValue("username"), Request.getValue("password"), remember == "1");
                if (msg == "") { return RedirectToAction("Index", "Home"); }
                ViewBag.Error = msg;
            }
            /* Clear */
            Session.Clear();
            Request.Cookies.Clear();
            return View();
        }

        public ActionResult LogOut()
        {
            var mode = Request.getValue("mode");
            if (mode == "force")
            {
                Session.Clear();
                Request.Cookies.Clear();
                return RedirectToAction("Index");
            }
            return View();
        }
    }
}