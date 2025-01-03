﻿using System.Web;
using System.Web.Configuration;
using System.Xml;

namespace ToolBaoCao
{
    public class WebConfigHelper
    {
        public static int GetMaxRequestLengthMB()
        {
            System.Configuration.Configuration config = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");
            var section = (HttpRuntimeSection)config.GetSection("system.web/httpRuntime");
            return section != null ? (int)(section.MaxRequestLength / 1024) : 4;
        }

        public static int GetMaxAllowedContentLengthMB()
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
                        return (int)(maxAllowedContentLength / (1024 * 1024));
                    }
                }
            }
            return 10;
        }

        public static void UpdateMaxLength(int maxRequestLengthMB = 0, int maxAllowedContentLengthMB = 0)
        {
            if (maxAllowedContentLengthMB == 0 && maxRequestLengthMB == 0) { return; }
            string webConfigPath = System.Web.HttpContext.Current.Server.MapPath("~/web.config");

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(webConfigPath);
            if (maxRequestLengthMB > 0)
            {
                /* Cập nhật maxRequestLength trong httpRuntime */
                XmlNode httpRuntimeNode = xmlDoc.SelectSingleNode("//system.web/httpRuntime");
                if (httpRuntimeNode != null)
                {
                    XmlAttribute maxRequestLengthAttr = httpRuntimeNode.Attributes["maxRequestLength"];
                    if (maxRequestLengthAttr != null) { maxRequestLengthAttr.Value = (maxRequestLengthMB * 1024).ToString(); }
                    else
                    {
                        maxRequestLengthAttr = xmlDoc.CreateAttribute("maxRequestLength");
                        maxRequestLengthAttr.Value = (maxRequestLengthMB * 1024).ToString();
                        httpRuntimeNode.Attributes.Append(maxRequestLengthAttr);
                    }
                }
            }
            if (maxAllowedContentLengthMB > 0)
            {
                /* Cập nhật maxAllowedContentLength trong requestLimits */
                XmlNode requestLimitsNode = xmlDoc.SelectSingleNode("//system.webServer/security/requestFiltering/requestLimits");
                if (requestLimitsNode != null)
                {
                    XmlAttribute maxAllowedContentLengthAttr = requestLimitsNode.Attributes["maxAllowedContentLength"];
                    if (maxAllowedContentLengthAttr != null)
                    {
                        maxAllowedContentLengthAttr.Value = (maxAllowedContentLengthMB * 1048576).ToString();
                    }
                    else
                    {
                        maxAllowedContentLengthAttr = xmlDoc.CreateAttribute("maxAllowedContentLength");
                        maxAllowedContentLengthAttr.Value = (maxAllowedContentLengthMB * 1048576).ToString();
                        requestLimitsNode.Attributes.Append(maxAllowedContentLengthAttr);
                    }
                }
            }
            xmlDoc.Save(webConfigPath);
        }
    }
}