using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class BaoCaoController : Controller
    {
        // GET: BaoCao
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Tuan()
        {
            if (Session["iduser"] == null)
            {
                ViewBag.Error = keyMSG.ErrorNotLoginAccess;
                return View();
            }
            var mode = Request.getValue("mode");
            string tmp = "";
            if (mode == "")
            {
                tmp = $"{Session["idtinh"]}".Trim();
                ViewBag.tinhSelect = tmp;
                tmp = $"{Session["nhom"]}".Trim() == "0" ? "" : $" WHERE idtinh = '{tmp}'";
                var dmTinh = AppHelper.dbSqliteMain.getDataTable($"SELECT id,ten FROM dmtinh{tmp} ORDER BY tt, ten");
                if (dmTinh.Rows.Count == 0) { ViewBag.Error = "Bạn chưa chọn hoặc được cấp tỉnh hoạt động"; return View(); }
                ViewBag.dmTinh = dmTinh;
                return View();
            }
            if (mode == "taive")
            {
                string pathFileTemplate = Server.MapPath("~/App_Data/baocaotuan.docx");
                if (System.IO.File.Exists(pathFileTemplate) == false)
                {
                    ViewBag.Error = "Không tìm thấy tập tin mẫu báo cáo 'baocaotuan.docx' trong thư mục App_Data";
                    return View();
                }
                string matinh = Request.getValue("matinh");
                string ngay = Request.getValue("thoigian");
                if (ngay.isDateVN() == false) { ViewBag.Error = $"Thời gian không đúng định dạng ngày/tháng/năm '{ngay}'"; return View(); }
                DateTime ngayTime = ngay.getFromDateVN();
                string thoigian = ngayTime.ToString("yyyyMMdd");
                string thang = ngayTime.ToString("MM");
                string nam = ngayTime.ToString("yyyy");
                var tailieu = new Dictionary<string, string>();
                /* Lấy tên tỉnh */
                tmp = $"{AppHelper.dbSqliteMain.getValue($"SELECT ten FROM dmTinh WHERE id='{matinh.sqliteGetValueField()}'")}";
                tailieu.Add("{X73}", tmp);
                /* Lấy ngày chọn báo cáo */
                tailieu.Add("{X74}", ngay);
                /*
                * X1={cột R (T-BHTT) bảng B02_TOANQUOC }
                * X71 = {cột S T_BHTT_NOI bảng B02_TOANQUOC }
                * X72={cột T T_BHTT_NGOAI bảng B02_TOANQUOC }
                */
                var data = AppHelper.dbSqliteWork.getDataTable($@"SELECT IFNULL(p1.t_bhtt, 0) AS x1, IFNULL(p1.t_bhtt_noi, 0) AS x71, IFNULL(p1.t_bhtt_ngoai, 0) AS x72
                    , IFNULL(p1.tyle_noitru, 0) AS x5
                    FROM b02chitiet p1 INNER JOIN b02 ON p1.id2=b02.id
                    WHERE b02.tu_thang={thang} AND b02.den_thang={thang} AND b02.nam={nam} AND b02.cs='1' AND p1.ma_tinh='{matinh.sqliteGetValueField()}' LIMIT 1");

                if (data.Rows.Count > 0)
                {
                    tailieu.Add("{X1}", data.Rows[0].ToString());
                    tailieu.Add("{X71}", data.Rows[1].ToString());
                    tailieu.Add("{X72}", data.Rows[2].ToString());
                    tailieu.Add("{X5}", data.Rows[3].ToString());
                }
                else
                {
                    tailieu.Add("{X1}", "0");
                    tailieu.Add("{X71}", "0");
                    tailieu.Add("{X72}", "0");
                    tailieu.Add("{X5}", "0");
                }
                /*
                 X2={“ Quyết định số: Nếu không tìm thấy dòng nào của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán thì “TW chưa giao dự toán, tạm lấy theo dự toán năm trước”, nếu thấy lấy số ký hiệu các dòng QĐ của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán}
                 X3={Như trên, ko thấy thì lấy tổng tiền các dòng dự toán năm trước, thấy thì lấy tổng số tiền các dòng quyết định năm nay}
                 */
                tailieu.Add("{X2}", "0"); /* Nhập từ bảng dữ liệu */
                tailieu.Add("{X3}", "0"); /* Nhập từ bảng dữ liệu */
                /* B02 */
                /*
                 * X5={Cột G, dòng MA_TINH=10};
                 * X6={cột G, dòng MA_TINH=00};
                 * X7={đoạn văn tùy thuộc X5> hay < X6. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số };
                 * X8={Sort cột G (TYLE_NOITRU) cao xuống thấp và lấy thứ tự};
                 * X9 ={tính toán: total cột F (TONG_LUOT_NOI) chia cho Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG=mã vùng của tỉnh báo cáo};
                 * X10 ={đoạn văn tùy thuộc X5> hay < X9. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số };
                 * X11= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort cột G (TYLE_NOITRU ) cao –thấp và lấy thứ tự}
                 * */
                data = AppHelper.dbSqliteWork.getDataTable($@"SELECT IFNULL(p1.tyle_noitru, 0) AS x6
                    FROM b02chitiet p1 INNER JOIN b02 ON p1.id2=b02.id
                    WHERE b02.tu_thang={thang} AND b02.den_thang={thang} AND b02.nam={nam} AND b02.cs='1' AND p1.ma_tinh='00' LIMIT 1");

                if (data.Rows.Count > 0)
                {
                    tailieu.Add("{X6}", data.Rows[0].ToString());
                }
                else
                {
                    tailieu.Add("{X6}", "0");
                }
                if (tailieu["{X5}"] == tailieu["{X6}"]) { tailieu["{X7}"] = "Bằng"; }
                else
                {
                    if (double.Parse(tailieu["{X5}"]) > double.Parse(tailieu["{X6}"])) { tailieu["{X7}"] = "Cao hơn"; }
                    else { tailieu["{X7}"] = "Thấp hơn"; }
                }
                if (tailieu["{X3}"] == "0") { tailieu.Add("{X4}", "0"); }
                else { tailieu.Add("{X4}", (double.Parse(tailieu["{X1}"])/double.Parse(tailieu["{X3}"])).ToString("0.###")); }

                using (var fileStream = new FileStream(pathFileTemplate, FileMode.Open, FileAccess.ReadWrite))
                {
                    var document = new XWPFDocument(fileStream);
                    foreach (var paragraph in document.Paragraphs)
                    {
                        foreach (var run in paragraph.Runs)
                        {
                            tmp = run.ToString();
                            foreach (var v in tailieu)
                            {
                                if (tmp.Contains(v.Key)) { tmp = tmp.Replace(v.Key, v.Value.Replace(".", ",")); }
                            }
                            /* Xóa hết các thông tin {X[0-9]+} nếu còn */
                            tmp = Regex.Replace(tmp, "{X[0-9]+}", "", RegexOptions.IgnoreCase);
                            run.SetText(tmp, 0);
                        }
                    }
                    MemoryStream memoryStream = new MemoryStream();
                    document.Write(memoryStream);
                    memoryStream.Position = 0;
                    return File(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{matinh}_{thoigian}.docx");
                }
            }
            return View();
        }
    }
}