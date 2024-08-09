using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    /* Dữ liệu báo cáo được lưu tại App_Data/import{matinh}.db; App_Data/bctuan{matinh}.db
     * File báo cáo được lưu tại App_Data/bctuan/{md5hash(iduser)}/{thoigantao}: _bctuan.docx, _plbctuan.xlsx, _b0200.xlsx, b02{matinh}.xlsx, ..
     * Dữ thống kê lưu tại App_Data/data.db
     * Nếu User bị thay đổi mã tỉnh làm việc sẽ huỷ toàn bộ tiến trình các bước nếu có
     * Dữ liệu bắt buộc B02 00, b02 cs, b04 00, b26 00, b26 cs
     */

    public class bcTuanController : ControllerCheckLogin
    {
        public ActionResult Index()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }

            return View();
        }

        public ActionResult Buoc1()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }

            return View();
        }

        public ActionResult Buoc2()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }

            return View();
        }

        public ActionResult Buoc3()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }

            return View();
        }
        public ActionResult Update()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }

            return View();
        }

        public ActionResult TruyVan()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            var matinh = $"{Session["idtinh"]}";
            try
            {
                var mode = Request.getValue("mode");
                if (mode == "truyvan")
                {
                    var ngay1 = Request.getValue("ngay1"); var ngay2 = Request.getValue("ngay2");
                    var time1 = DateTime.Now; var time2 = DateTime.Now;
                    if (ngay1.isDateVN(out time1) == false) { throw new Exception($"từ ngày không đúng định dạng ngày/tháng/năm '{ngay1}'"); }
                    if (ngay2.isDateVN(out time2) == false) { throw new Exception($"từ ngày không đúng định dạng ngày/tháng/năm '{ngay2}'"); }
                    ViewBag.ngay1 = ngay1;
                    ViewBag.ngay2 = ngay2;
                    return View();
                }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }
    }
}