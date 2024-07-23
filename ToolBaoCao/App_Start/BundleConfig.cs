using System.Web;
using System.Web.Optimization;

namespace ToolBaoCao
{
    public class BundleConfig
    {
        // For more information on bundling, visit https://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/jquery").Include(
                     "~/vendor/jquery/jquery.min.js",
                     "~/vendor/jquery-easing/jquery.easing.min.js",
                     "~/js/moment.js"));

            // Use the development version of Modernizr to develop with and learn from. Then, when you're
            // ready for production, use the build tool at https://modernizr.com to pick only the tests you need.
            bundles.Add(new ScriptBundle("~/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new ScriptBundle("~/bootstrap").Include(
                      "~/Scripts/bootstrap.bundle.min.js",
                        "~/vendor/datatables/jquery.dataTables.js",
                      "~/vendor/datatables/dataTables.bootstrap4.min.js",
                      "~/js/sb-admin-2.min.js"));

            bundles.Add(new StyleBundle("~/css").Include(
                      "~/vendor/fontawesome-free/css/all.min.css",
                      "~/Content/bootstrap.min.css",
                      "~/css/sb-admin-2.min.css"));
        }
    }
}
