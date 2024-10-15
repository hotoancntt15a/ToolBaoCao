using System.Text.RegularExpressions;
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
            if (value == "") { AppHelper.appConfig.Remove(key); }
            else
            {
                if (key == "maxRequestLengthMB")
                {
                    if (Regex.IsMatch(value, @"^\d+$") == false) { value = "0"; }
                    WebConfigHelper.UpdateMaxLength(int.Parse(value));
                }
                if (key == "maxAllowedContentLengthMB")
                {
                    if (Regex.IsMatch(value, @"^\d+$") == false) { value = "0"; }
                    WebConfigHelper.UpdateMaxLength(maxAllowedContentLengthMB: int.Parse(value));
                }
                if(key == "maxSizeFileUploadMB")
                {
                    if (Regex.IsMatch(value, @"^\d+$") == false) { value = "0"; }
                    WebConfigHelper.UpdateMaxLength(int.Parse(value), int.Parse(value));
                }
                AppHelper.appConfig.Set(key, value);
            }
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