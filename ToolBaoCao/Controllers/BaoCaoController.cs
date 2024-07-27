using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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
                var tailieu = new Dictionary<string, string>(); double so1 = 0; double so2 = 0;
                var tmpD = new Dictionary<string, string>();

                var b02TQ = AppHelper.dbSqliteWork.getDataTable($@"SELECT p1.* FROM b02chitiet p1 INNER JOIN b02 ON p1.id2=b02.id
                    WHERE b02.tu_thang={thang} AND b02.den_thang={thang} AND b02.nam={nam} AND b02.cs='0'").AsEnumerable();
                if (b02TQ.Count() == 0) { ViewBag.Error = "B02 Toàn Quốc không có dữ liệu phù hợp truy vấn"; return View(); }
                var b26TQ = AppHelper.dbSqliteWork.getDataTable($@"SELECT p1.* FROM b26chitiet p1 INNER JOIN b26 ON p1.id2=b26.id
                    WHERE b26.thoi_gian = '{thoigian}' AND b26.cs='0'").AsEnumerable();
                if (b26TQ.Count() == 0) { ViewBag.Error = "B26 Toàn quốc không có dữ liệu phù hợp truy vấn"; return View(); }

                var dataTinhB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == matinh).FirstOrDefault();
                if (dataTinhB02 == null) { ViewBag.Error = "B02 không có dữ liệu tỉnh phù hợp truy vấn"; return View(); }
                var dataTinhB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == matinh).FirstOrDefault();
                if (dataTinhB26 == null) { ViewBag.Error = "B26 không có dữ liệu tỉnh phù hợp truy vấn"; return View(); }

                var dataTQB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault();
                if (dataTQB02 == null) { ViewBag.Error = "B02 không có dữ liệu toàn quốc phù hợp truy vấn"; return View(); }
                var dataTQB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault();
                if (dataTQB26 == null) { ViewBag.Error = "B26 không có dữ liệu toàn quốc phù hợp truy vấn"; return View(); }

                string mavung = dataTinhB02["ma_vung"].ToString();

                /* X1 = {cột R (T-BHTT) bảng B02_TOANQUOC } */
                tailieu.Add("{X1}", dataTinhB02["t_bhtt"].ToString());
                /* X2 = {“ Quyết định số: Nếu không tìm thấy dòng nào của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán thì “TW chưa giao dự toán, tạm lấy theo dự toán năm trước”, nếu thấy lấy số ký hiệu các dòng QĐ của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán} */
                tailieu.Add("{X2}", "0");
                /* X3 = {Như trên, ko thấy thì lấy tổng tiền các dòng dự toán năm trước, thấy thì lấy tổng số tiền các dòng quyết định năm nay} */
                tailieu.Add("{X3}", "0");
                /* X4={X1/X2 %} So sánh với dự toán, tỉnh đã sử dụng */
                if (tailieu["{X2}"] == "0") { tailieu.Add("{X4}", "0"); }
                else { tailieu.Add("{X4}", (double.Parse(tailieu["{X1}"]) / double.Parse(tailieu["{X2}"])).ToString("0.###")); }
                /* X5 = {Cột tyle_noitru, dòng MA_TINH=10} bảng B02_TOANQUOC */
                tailieu.Add("{X5}", dataTinhB02["tyle_noitru"].ToString());
                /* X6 = {Cột tyle_noitru, dòng MA_TINH=00} bảng B02_TOANQUOC */
                tailieu.Add("{X6}", dataTQB02["tyle_noitru"].ToString());
                /* X7 = {đoạn văn tùy thuộc X5> hay < X6. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X7}", "bằng");
                so1 = double.Parse(tailieu["{X5}"]);
                so2 = double.Parse(tailieu["{X6}"]);
                if (so1 > so2) { tailieu["{X7}"] = $"cao hơn {(so1 - so2).ToString("0.###")}"; }
                else { if (so1 < so2) { tailieu["{X7}"] = $"thấp hơn {(so2 - so1).ToString("0.###")}"; } }
                /* X8={Sort cột G (TYLE_NOITRU) cao xuống thấp và lấy thứ tự}; */
                var sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
                int position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("X8", position.ToString());
                /* X9 ={tính toán: total cột F (TONG_LUOT_NOI) chia cho Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG=mã vùng của tỉnh báo cáo}; */
                tailieu.Add("{X9}", "0");
                so2 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung)
                            .Sum(row => row.Field<long>("tong_luot"));
                if (so2 != 0)
                {
                    so1 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung)
                                .Sum(row => row.Field<long>("tong_luot_noi"));
                    tailieu["{X9}"] = (so1 / so2).ToString("0.###");
                }
                /* X10 ={đoạn văn tùy thuộc X5> hay < X9. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X10}", "bằng");
                so1 = double.Parse(tailieu["{X5}"]);
                so2 = double.Parse(tailieu["{X9}"]);
                if (so1 > so2) { tailieu["{X10}"] = $"cao hơn {(so1 - so2).ToString("0.###")}"; }
                else { if (so1 < so2) { tailieu["{X10}"] = $"thấp hơn {(so2 - so1).ToString("0.###")}"; } }
                /* X11= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort cột G (TYLE_NOITRU ) cao –thấp và lấy thứ tự} */
                sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                    .OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X11}", position.ToString());
                /* X12 = Ngày điều trị bình quân X12={Cột H NGAY_DTRI_BQ , dòng MA_TINH=10}; */
                tailieu.Add("{X12}", dataTinhB02["ngay_dtri_bq"].ToString());
                /* X13 = Nbình quân toàn quốc X13={cột H NGAY_DTRI_BQ, dòng MA_TINH=00}; */
                tailieu.Add("{X13}", dataTQB02["ngay_dtri_bq"].ToString());
                /* X14 = Số chênh lệch X14={đoạn văn tùy thuộc X12> hay < X13. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X14}", "bằng");
                so1 = double.Parse(tailieu["{X12}"]);
                so2 = double.Parse(tailieu["{X13}"]);
                if (so1 > so2) { tailieu["{X14}"] = $"cao hơn {(so1 - so2).ToString("0.###")}"; }
                else { if (so1 < so2) { tailieu["{X14}"] = $"thấp hơn {(so2 - so1).ToString("0.###")}"; } }
                /* X15 = xếp thứ so toàn quốc X15={Sort cột H (NGAY_DTRI_BQ) cao xuống thấp và lấy thứ tự}; */
                sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X15}", position.ToString());
                /* X16 = Bình quân vùng X16 ={tính toán: A-Tổng ngày điều trị nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của cá tỉnh cùng mã vùng. A=Total(cột H (NGAY_DTRI_BQ) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                tailieu.Add("{X16}", "0");
                so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => r.Field<long>("tong_luot_noi"));
                if (so2 != 0)
                {
                    so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => (r.Field<double>("ngay_dtri_bq") * r.Field<long>("tong_luot_noi")));
                    tailieu["{X16}"] = (so1 / so2).ToString("0.###");
                }
                /* X17 = Số chênh lệch X17 ={đoạn văn tùy thuộc X12> hay < X16. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X17}", "bằng");
                so1 = double.Parse(tailieu["{X12}"]);
                so2 = double.Parse(tailieu["{X16}"]);
                if (so1 > so2) { tailieu["{X17}"] = $"cao hơn {(so1 - so2).ToString("0.###")}"; }
                else { if (so1 < so2) { tailieu["{X17}"] = $"thấp hơn {(so2 - so1).ToString("0.###")}"; } }
                /* X18 = đứng thứ so với vùng X18 = {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột H (NGAY_DTRI_BQ) cao –thấp và lấy thứ tự} */
                sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                    .OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X18}", position.ToString());
                /* X19 = Chi bình quân chung X19={Cột I (CHI_BQ_CHUNG), dòng MA_TINH=10}; */
                tailieu.Add("{X19}", dataTinhB02["chi_bq_chung"].ToString());
                /* X20 = bình quân toàn quốc X20={cột I (CHI_BQ_CHUNG), dòng MA_TINH=00}; */
                tailieu.Add("{X20}", dataTQB02["chi_bq_chung"].ToString());
                /* X21 = Số chênh lệch X21={đoạn văn tùy thuộc X19> hay < X20. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X21}", "bằng");
                so1 = double.Parse(tailieu["{X19}"]);
                so2 = double.Parse(tailieu["{X20}"]);
                if (so1 > so2) { tailieu["{X21}"] = $"cao hơn {(so1 - so2).ToString("0.###")}"; }
                else { if (so1 < so2) { tailieu["{X21}"] = $"thấp hơn {(so2 - so1).ToString("0.###")}"; } }
                /* X22 = xếp thứ so toàn quốc X22={Sort cột I (CHI_BQ_CHUNG) cao xuống thấp và lấy thứ tự}; */
                sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("chi_bq_chung")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X22}", position.ToString());
                /* X23 = Bình quân vùng X23={tính toán: A-Tổng chi các tỉnh cùng mã vùng / B- Tổng lượt kcb của các tỉnh cùng mã vùng. A=Total  (cột I (CHI_BQ_CHUNG) * cột D (TONG_LUOT)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                tailieu.Add("{X23}", "0");
                so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => r.Field<double>("chi_bq_chung"));
                if (so2 != 0)
                {
                    so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => (r.Field<double>("chi_bq_chung") * r.Field<long>("tong_luot")));
                    tailieu["{X23}"] = (so1 / so2).ToString("0.###");
                }
                /* X24 = Số chênh lệch X24 ={đoạn văn tùy thuộc X19> hay < X23. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X24}", "bằng");
                so1 = double.Parse(tailieu["{X19}"]);
                so2 = double.Parse(tailieu["{X23}"]);
                if (so1 > so2) { tailieu["{X24}"] = $"cao hơn {(so1 - so2).ToString("0.###")}"; }
                else { if (so1 < so2) { tailieu["{X24}"] = $"thấp hơn {(so2 - so1).ToString("0.###")}"; } }
                /* X25 đứng thứ so với vùng X25= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột I (CHI_BQ_CHUNG) cao –thấp và lấy thứ tự} */
                sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                   .OrderByDescending(row => row.Field<double>("chi_bq_chung")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X25}", position.ToString());
                /* X26 = Chi bình quân ngoại trú X26={Cột J (CHI_BQ_NGOAI), dòng MA_TINH=10}; */
                tailieu.Add("{X26}", dataTinhB02["chi_bq_ngoai"].ToString());
                /* X27 = bình quân toàn quốc X27={cột J (CHI_BQ_NGOAI), dòng MA_TINH=00}; */
                tailieu.Add("{X27}", dataTQB02["chi_bq_ngoai"].ToString());
                /* X28 = Số chênh lệch X28={đoạn văn tùy thuộc X26> hay < X27. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X28}", "bằng");
                so1 = double.Parse(tailieu["{X26}"]);
                so2 = double.Parse(tailieu["{X27}"]);
                if (so1 > so2) { tailieu["{X28}"] = $"cao hơn {(so1 - so2).ToString("0.###")}"; }
                else { if (so1 < so2) { tailieu["{X28}"] = $"thấp hơn {(so2 - so1).ToString("0.###")}"; } }
                /* X29 = xếp thứ so toàn quốc X29={Sort cột J(CHI_BQ_NGOAI) cao xuống thấp và lấy thứ tự}; */
                sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("chi_bq_ngoai")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X29}", position.ToString());
                /* X30 = Bình quân vùng X30={tính toán: A-Tổng chi ngoại trú các tỉnh cùng mã vùng / B- Tổng lượt kcb ngoại trú của các tỉnh cùng mã vùng. A=Total  (cột J (CHI_BQ_NGOAI) * cột E (TONG_LUOT_NGOAI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột E (TONG_LUOT_NGOAI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                tailieu.Add("{X30}", "0");
                so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => r.Field<double>("chi_bq_ngoai"));
                if (so2 != 0)
                {
                    so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => (r.Field<double>("chi_bq_ngoai") * r.Field<long>("tong_luot_ngoai")));
                    tailieu["{X30}"] = (so1 / so2).ToString("0.###");
                }
                /* X31 = Số chênh lệch X31 ={đoạn văn tùy thuộc X19> hay < X30. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X31}", "bằng");
                so1 = double.Parse(tailieu["{X19}"]);
                so2 = double.Parse(tailieu["{X30}"]);
                if (so1 > so2) { tailieu["{X31}"] = $"cao hơn {(so1 - so2)}"; }
                else { if (so1 < so2) { tailieu["{X31}"] = $"thấp hơn {(so2 - so1)}"; } }
                /* X32 = đứng thứ so với vùng X32= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột J (CHI_BQ_NGOAI) cao –thấp và lấy thứ tự} */
                sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                   .OrderByDescending(row => row.Field<double>("chi_bq_ngoai")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X32}", position.ToString());
                /* X33 = Chi bình quân nội trú X33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
                tailieu.Add("{X33}", dataTinhB02["chi_bq_noi"].ToString());
                /* X34 = bình quân toàn quốc X34={cột K (CHI_BQ_NOI), dòng MA_TINH=00}; */
                tailieu.Add("{X34}", dataTQB02["chi_bq_noi"].ToString());
                /* X35 = Số chênh lệch X35={đoạn văn tùy thuộc X33> hay < X34. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X35}", "bằng");
                so1 = double.Parse(tailieu["{X33}"]);
                so2 = double.Parse(tailieu["{X34}"]);
                if (so1 > so2) { tailieu["{X35}"] = $"cao hơn {(so1 - so2)}"; }
                else { if (so1 < so2) { tailieu["{X35}"] = $"thấp hơn {(so2 - so1)}"; } }
                /* X36= xếp thứ so toàn quốc X36={Sort cột K CHI_BQ_NOI cao xuống thấp và lấy thứ tự}; */
                sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("chi_bq_noi")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X36}", position.ToString());
                /* X37 = Bình quân vùng X37={tính toán: A-Tổng chi nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của các tỉnh cùng mã vùng. A=Total  (cột K (CHI_BQ_NOI) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                tailieu.Add("{X37}", "0");
                so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => r.Field<double>("chi_bq_noi"));
                if (so2 != 0)
                {
                    so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                            .Sum(r => (r.Field<double>("chi_bq_noi") * r.Field<long>("tong_luot_noi")));
                    tailieu["{X37}"] = (so1 / so2).ToString("0.###");
                }
                /* X38 = số chênh lệch X38 ={đoạn văn tùy thuộc X33> hay < X34. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X38}", "bằng");
                so1 = double.Parse(tailieu["{X33}"]);
                so2 = double.Parse(tailieu["{X34}"]);
                if (so1 > so2) { tailieu["{X38}"] = $"cao hơn {(so1 - so2)}"; }
                else { if (so1 < so2) { tailieu["{X38}"] = $"thấp hơn {(so2 - so1)}"; } }
                /* X39 đứng thứ so với vùng X39= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột K (CHI_BQ_NOI) cao –thấp và lấy thứ tự} */
                tailieu.Add("{X39}", getPosition(mavung, matinh, "chi_bq_noi", b02TQ));

                /* ----- Dữ liệu X40 trở lên lọc dữ liệu tù B26 ------- */
                /* X40 = Bình quân xét nghiệm X40= {cột P (bq_xn) dòng có mã tỉnh = 10}; B26 */
                tmpD = buildB26("{X40}", "{X41}", "{X42}", "bq_xn", "bq_xn_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X43 Bình quân CĐHA X43= {cột R(bq_cdha) dòng có mã tỉnh =10}; */
                tmpD = buildB26("{X43}", "{X44}", "{X45}", "bq_cdha", "bq_cdha_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X46 Bình quân thuốc X46= {cột T(bq_thuoc) dòng có mã tỉnh =10}; */
                tmpD = buildB26("{X46}", "{X47}", "{X48}", "bq_thuoc", "bq_thuoc_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X49 Bình quân chi phẫu thuật X49= {cột V(bq_pt) dòng có mã tỉnh =10}; */
                tmpD = buildB26("{X49}", "{X50}", "{X51}", "bq_pt", "bq_pt_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X52 Bình quân chi thủ thuật X52= {cột X(bq_tt) dòng có mã tỉnh =10}; */
                tmpD = buildB26("{X52}", "{X53}", "{X54}", "bq_tt", "bq_tt_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X55 Bình quân chi vật tư y tế X55= {cột Z(bq_vtyt) dòng có mã tỉnh =10}; */
                tmpD = buildB26("{X55}", "{X56}", "{X57}", "bq_vtyt", "bq_vtyt_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X58 Bình quân chi tiền giường X58= {cột AB(bq_giuong) dòng có mã tỉnh =10}; */
                tmpD = buildB26("{X58}", "{X59}", "{X60}", "bq_giuong", "bq_giuong_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X61 Chỉ định xét nghiệm X61={cột AD, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
                tailieu.Add("X61", ((double)dataTinhB26["chi_dinh_xn"] * 100).ToString("0.##"));

                /* X71 = {cột S T_BHTT_NOI bảng B02_TOANQUOC } */
                tailieu.Add("{X71}", dataTinhB02["t_bhtt_noi"].ToString());
                /* X72 = {cột T T_BHTT_NGOAI bảng B02_TOANQUOC } */
                tailieu.Add("{X72}", dataTinhB02["t_bhtt_ngoai"].ToString());
                /* X73 Lấy tên tỉnh */
                tmp = $"{AppHelper.dbSqliteMain.getValue($"SELECT ten FROM dmTinh WHERE id='{matinh.sqliteGetValueField()}'")}";
                tailieu.Add("{X73}", tmp);
                /* X74 Lấy ngày chọn báo cáo */
                tailieu.Add("{X74}", ngay);

                using (var fileStream = new FileStream(pathFileTemplate, FileMode.Open, FileAccess.ReadWrite))
                {
                    var document = new XWPFDocument(fileStream);
                    foreach (var paragraph in document.Paragraphs)
                    {
                        foreach (var run in paragraph.Runs)
                        {
                            tmp = run.ToString();
                            // Sử dụng Regex để tìm tất cả các match
                            MatchCollection matches = Regex.Matches(tmp, "{x[0-9]+}", RegexOptions.IgnoreCase);
                            foreach (Match match in matches) { tmp = tmp.Replace(match.Value, tailieu.getValue(match.Value, "")); }
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
        private string getPosition(string mavung, string matinh, string field, EnumerableRowCollection<DataRow> data)
        {
            if (mavung != "")
            {
                var sortedRows = data.Where(r => r.Field<string>("ma_vung") == mavung)
                   .OrderByDescending(row => row.Field<double>("chi_bq_noi")).ToList();
                return (sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1).ToString();
            }
            var s = data.OrderByDescending(row => row.Field<double>("chi_bq_noi")).ToList();
            return (s.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1).ToString();
        }
        private Dictionary<string, string> buildB26(string key1, string key2, string key3, string field1, string field2, DataRow row)
        {
            var d = new Dictionary<string, string>();
            /* X46 Bình quân cột [x] dòng có mã tỉnh = 10}; */
            var x = (double)row[field1];
            d.Add(key1, row[field1].ToString());
            /* X47 số tương đối X47={nếu cột [x+1] dòng có mã tỉnh=10 là số dương, “tăng “ & cột [x+1] & “%”, không thì “giảm “ & cột [x+1] %}; */
            d.Add(key2, "bằng");
            var x1 = (double)row[field2]; /* s */
            if (x1 > 0) { d[key2] = $"tăng {x1}%"; }
            else { if (x1 < 0) { d[key2] = $"giảm {Math.Abs(x1)}%"; } }
            /* X48 số tuyệt đối X48={nếu cột [x+1] là dương, “tăng “ & [cột [x] - (cột [x] / (cột [x+1] +100) *100 )] & “ đồng”, không thì “giảm “ & [cột [x]- (cột [x] / (cột [x+1]+100) *100 )] & “ đồng”} */
            d.Add(key3, "bằng");
            if (x1 > 0)
            {
                d[key3] = "tăng " + (x - (x / (x1 + 100) * 100)).ToString("0.##") + " đồng";
            }
            else
            {
                if (x1 < 0)
                {
                    d[key3] = "giảm " + (x - (x / (x1 + 100) * 100)).ToString("0.##") + " đồng";
                }
            }
            return d;
        }
    }
}