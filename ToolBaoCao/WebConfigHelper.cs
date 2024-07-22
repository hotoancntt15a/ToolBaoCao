using Microsoft.Web.Administration;
using System.Configuration;
using System.Web;
using System.Web.Configuration;
using System.Xml;

namespace ToolBaoCao
{
    public class WebConfigHelper
    {
        public static int GetMaxRequestLength()
        {
            System.Configuration.Configuration config = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");
            var section = (HttpRuntimeSection)config.GetSection("system.web/httpRuntime");
            return section != null ? section.MaxRequestLength : 4096;
        }
        public static long GetMaxAllowedContentLength()
        {
            string webConfigPath = HttpContext.Current.Server.MapPath("~/web.config");
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(webConfigPath);

            XmlNode requestLimitsNode = xmlDoc.SelectSingleNode("//system.webServer/security/requestFiltering/requestLimits");
            if (requestLimitsNode != null)
            {
                XmlAttribute maxAllowedContentLengthAttr = requestLimitsNode.Attributes["maxAllowedContentLength"];
                if (maxAllowedContentLengthAttr != null)
                {
                    long maxAllowedContentLength;
                    if (long.TryParse(maxAllowedContentLengthAttr.Value, out maxAllowedContentLength))
                    {
                        return maxAllowedContentLength;
                    }
                }
            }
            return 10485760;
        }

        public static void UpdateMaxRequestLength(int maxRequestLengthKB)
        {
            System.Configuration.Configuration config = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");
            var section = (HttpRuntimeSection)config.GetSection("system.web/httpRuntime");
            section.MaxRequestLength = maxRequestLengthKB;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("system.web/httpRuntime");
        }

        public static void UpdateMaxAllowedContentLength(long maxAllowedContentLengthBytes)
        {
            using (ServerManager serverManager = new ServerManager())
            {
                var site = serverManager.Sites["Default Web Site"]; // Thay thế bằng tên site của bạn
                var config = site.GetWebConfiguration();

                var section = config.GetSection("system.webServer/security/requestFiltering");
                var requestLimitsElement = section.GetChildElement("requestLimits");
                requestLimitsElement["maxAllowedContentLength"] = maxAllowedContentLengthBytes;

                serverManager.CommitChanges();
            }
        }
    }
}