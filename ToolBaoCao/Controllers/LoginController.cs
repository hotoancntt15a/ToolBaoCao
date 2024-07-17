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
            return View();
        }

        public ActionResult LogOut()
        {
            var mode = Request.getValue("mode");
            if (mode == "force")
            {
                Session.Clear();
                Session.Abandon();
                if (Request.Cookies != null)
                {
                    foreach (string cookieName in Request.Cookies.AllKeys)
                    {
                        HttpCookie cookie = Request.Cookies[cookieName];
                        cookie.Expires = DateTime.Now.AddDays(-1);
                        Response.Cookies.Add(cookie);
                    }
                }
                return RedirectToAction("Index");
            }
            return View();
        }
    }
}