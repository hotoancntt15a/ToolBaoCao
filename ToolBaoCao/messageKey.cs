using Org.BouncyCastle.Bcpg;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ToolBaoCao
{
    public static class keyMSG
    {
        public static string SessionIPAddress = "Connect.IpAddress";
        public static string SessionBrowserInfo = "Connect.BrowserInfo";
        public static string ErrorNotLogin = "Bạn chưa đăng nhập hoặc đã quá hạn đăng nhập";
        public static string ErrorNotLoginAccess = "Bạn vui lòng đăng nhập để sử dụng chức năng này";
        public static string ErrorHttpConnetNull = "Không xác định được HttpContext";
    }
}