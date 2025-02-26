using System;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace ToolBaoCao
{
    public class ControllerCheckLogin : Controller
    {
        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            AppHelper.CheckIsLogin();
            /* if (AppHelper.CheckIsLogin() != true) { filterContext.Result = new RedirectResult("/Login/"); } */
            base.OnActionExecuting(filterContext);
        }
    }

    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            AppHelper.LoadStart();
            try
            {
                var db = BuildDatabase.getDBUserOnline();
                db.Execute("DELETE FROM useronline;");
            }
            catch { }
        }

        protected void Application_Error()
        {
            var exS = Server.GetLastError();
            var httpException = exS as HttpException ?? exS.InnerException as HttpException;
            if (httpException == null) return;
            try
            {
                if (httpException.InnerException != null)
                {
                    var innerHttpException = httpException.InnerException as HttpException;
                    if (innerHttpException != null && innerHttpException.WebEventCode == System.Web.Management.WebEventCodes.RuntimeErrorPostTooLarge)
                    {
                        string message = $"Tập tin đẩy lên lớn hơn {WebConfigHelper.GetMaxAllowedContentLengthMB()}MB";
                        throw new Exception($"Message={HttpUtility.UrlEncode(message)}");
                    }
                }
                if (httpException.GetHttpCode() == 404) { throw new Exception($"UrlNotFound={HttpUtility.UrlEncode("Không tìm thấy trang " + HttpContext.Current.Request.Url.PathAndQuery)}"); }
                string errorMessage = $"Message={HttpUtility.UrlEncode(httpException.Message)}" +
                                      $"&StackTrace={HttpUtility.UrlEncode(httpException.StackTrace ?? "No stack trace")}" +
                                      $"&WebEventCode={httpException.WebEventCode}" +
                                      $"&ErrorCode={httpException.ErrorCode}";
                throw new Exception(errorMessage);
            }
            catch (Exception ex)
            {
                if (HttpContext.Current == null) { ex.saveError("Application_Error: "); }
                else
                {
                    if (HttpContext.Current.Session == null) { ex.saveError("Application_Error: "); }
                    else
                    {
                        HttpContext.Current.Session["ErrorMessage"] = ex.Message;
                        Response.Redirect($"~/Error");
                    }
                }
            }
            finally { Server.ClearError(); }
        }

        private void Session_Start(object sender, EventArgs e)
        {
            var http = HttpContext.Current;
            Session[keyMSG.SessionIPAddress] = http.GetUserIpAddress();
            Session[keyMSG.SessionBrowserInfo] = http.GetUserBrowser();
            var db = BuildDatabase.getDBUserOnline();
            int maxSeccondsOnline = 15 * 60;
            try { db.Execute($"DELETE useronline WHERE ({DateTime.Now.toTimestamp()} - time2) > {maxSeccondsOnline}"); } catch { }
        }

        private void Session_End(object sender, EventArgs e)
        {
            Session.Clear();
        }
    }
}