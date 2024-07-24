using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Security;
using System.Web.SessionState;

namespace ToolBaoCao
{
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
        }

        protected void Application_Error()
        {
            var ex = Server.GetLastError();
            var httpException = ex as HttpException ?? ex.InnerException as HttpException;
            if (httpException == null) return;
            if(httpException.InnerException != null)
            {
                if (((System.Web.HttpException)httpException.InnerException).WebEventCode == System.Web.Management.WebEventCodes.RuntimeErrorPostTooLarge)
                {
                    Response.Write("Too big a file, dude");
                }
            }            
            if (httpException.GetHttpCode() == 404) { Response.Redirect("~/Error"); }
        }

        private void Session_Start(object sender, EventArgs e)
        {
            Session[keyMSG.SessionIPAddress] = GetUserIpAddress();
            Session[keyMSG.SessionBrowserInfo] = GetUserBrowserInfo();
        }

        private void Session_End(object sender, EventArgs e)
        {
            Session.Clear();
        }

        private string GetUserIpAddress()
        {
            string ipAddress = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (string.IsNullOrEmpty(ipAddress)) { ipAddress = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"]; }
            // Trường hợp có nhiều địa chỉ IP trong X-Forwarded-For, lấy địa chỉ đầu tiên
            if (!string.IsNullOrEmpty(ipAddress) && ipAddress.Contains(",")) { ipAddress = ipAddress.Split(',')[0].Trim(); }
            return ipAddress;
        }
        private string GetUserBrowserInfo()
        {
            string userAgent = HttpContext.Current.Request.UserAgent;
            HttpBrowserCapabilities browser = HttpContext.Current.Request.Browser;
            string browserName = browser.Browser;
            string browserVersion = browser.Version;
            return $"{browserName} {browserVersion} ({userAgent})";
        }
    }
}