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
            
        }

        private void Session_End(object sender, EventArgs e)
        {
            Session.Clear();
        }
    }
}