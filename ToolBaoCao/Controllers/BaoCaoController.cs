using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Mvc;
using System.Xml.Linq;
using zModules.NPOIExcel;

namespace ToolBaoCao.Controllers
{
    public class BaoCaoController : ControllerCheckLogin
    {
        // GET: BaoCao
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult BCTuanTruyVan(string matinh, string ngay1, string ngay2, string mode)
        {
            if (Session["iduser"] == null)
            {
                ViewBag.Error = keyMSG.NotLoginAccess;
                return View();
            }
            long time1 = 0, time2 = 0;
            try
            {
                if (string.IsNullOrEmpty(mode)) { throw new Exception("Tham số thực thi không đúng"); }
                if (mode != "truyvan") { throw new Exception("Tham số thực thi không đúng"); }
                if (string.IsNullOrEmpty(matinh)) { throw new Exception("Bạn chưa chọn tỉnh làm việc"); }
                if (string.IsNullOrEmpty(ngay1)) { throw new Exception("Bạn chưa chọn từ ngày"); }
                if (string.IsNullOrEmpty(ngay2)) { throw new Exception("Bạn chưa chọn đến ngày"); }
                if (ngay1.isDateVN() == false) { throw new Exception($"Từ ngày {ngay1} không đúng định dạng ngày/Tháng/Năm"); }
                if (ngay2.isDateVN() == false) { throw new Exception($"Đến ngày {ngay2} không đúng định dạng ngày/Tháng/Năm"); }
                time1 = ngay1.getFromDateVN().toTimestamp();
                time2 = ngay2.getFromDateVN().toTimestamp();
                if (time2 < time1) { throw new Exception($"Từ ngày {ngay1} phải nhỏ hơn đến ngày {ngay2}"); }
                if (time2 == 0) { throw new Exception($"Ngày không hợp lệ {ngay2}"); }
                if (Regex.IsMatch(matinh, @"^\d+$") == false) { throw new Exception($"Mã tỉnh '{matinh}' làm việc không hợp lệ"); }
            }
            catch (Exception ex) { ViewBag.Error = ex.Message; return View(); }
            try
            {
                ViewBag.ngay1 = ngay1;
                ViewBag.ngay2 = ngay2;
                ViewBag.matinh = matinh;
                var dbBaoCao = BuildDatabase.getDbSQLiteBaoCao();
                var data = dbBaoCao.getDataTable($"SELECT id, ma_tinh, userid, date(ngay, 'auto', '+7 hours') AS ngayGMT7, datetime(timecreate, 'auto', '+7 hours') AS taoLanCuoi FROM bctuandocx WHERE ma_tinh='{matinh}' AND (ngay >= {time1} AND ngay <= {time2})");
                ViewBag.data = data;
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); return View(); }
            return View();
        }

        public ActionResult BCTuanCreate(string objectid)
        {
            if (Session["iduser"] == null) { ViewBag.Error = keyMSG.NotLoginAccess; return View(); }
            DateTime timeStart = DateTime.Now;
            string tmp = "";
            string mode = Request.getValue("mode");
            try
            {
                if (mode == "thuchien")
                {
                    tmp = Request.getValue("x2");
                    if (Regex.IsMatch(tmp, @"^\d+$") == false) { ViewBag.Error = $"Số của quyết định giao dự toán không đúng {tmp}"; return View(); }
                    tmp = Request.getValue("x3");
                    if (Regex.IsMatch(tmp, @"^\d+$") == false) { ViewBag.Error = $"Tổng số tiền các dòng quyết định năm nay không đúng {tmp}"; return View(); }
                    string matinh = Request.getValue("matinh");
                    string ngay = Request.getValue("thoigian");
                    if (ngay.isDateVN() == false) { ViewBag.Error = $"Thời gian không đúng định dạng ngày/tháng/năm '{ngay}'"; return View(); }
                    DateTime ngayTime = ngay.getFromDateVN();
                    string thoigian = ngayTime.ToString("yyyyMMdd");

                    var tailieu = buildBaoCaoTuan(ngayTime, matinh, $"{Session["iduser"]}", Request.getValue("x2"), Request.getValue("x3"), Request.getValue("x67"), Request.getValue("x68"), Request.getValue("x69"), Request.getValue("x70"));
                    if (tailieu.ContainsKey("Error")) { ViewBag.Error = tailieu["Error"]; return View(); }
                    return Content($"<div class=\"alert alert-info\">Thao tác thành công ({timeStart.getTimeRun()})</div>");
                }
                if (mode == "save")
                {
                    if (string.IsNullOrEmpty(objectid) == false) { ViewBag.Error = "Không có tham số cập nhật"; }
                    var tailieu = new Dictionary<string, string>();
                    foreach (var key in Request.Form.AllKeys) { if (Regex.IsMatch(key, @"^x\d+$")) { tailieu[key] = Request.Form[key].Trim(); } }
                    /* kiểm tra dữ liệu trong cơ sở dữ liệu */
                    var dbBaoCao = BuildDatabase.getDbSQLiteBaoCao();
                    var item = dbBaoCao.getDataTable($"SELECT * FROM bctuandocx WHERE id='{objectid}'");
                    foreach (DataColumn c in item.Columns)
                    {
                        /*  Kiểm tra trường số */
                        if (c.ColumnName.StartsWith("x"))
                        {
                            if (tailieu.ContainsKey(c.ColumnName) && (c.DataType == typeof(long) || c.DataType == typeof(double)))
                            {
                                if (tailieu[c.ColumnName] == "") { tailieu[c.ColumnName] = "0"; }
                                else if (Regex.IsMatch(tailieu[c.ColumnName], @"^-?\d+([.]\d+)?$") == false)
                                {
                                    ViewBag.Error = $"Trường số {c.ColumnName} không đúng: '{tailieu[c.ColumnName]}'";
                                    return View();
                                }
                            }
                            continue;
                        }
                    }
                    /*  Kiểm tra trường ngày */
                    tailieu["timecreate"] = DateTime.Now.toTimestamp().ToString();
                    /* X4 = X1/X2 %*/
                    tailieu["x4"] = tailieu["x2"] == "0" ? "0" : (double.Parse(tailieu["x1"]) / double.Parse(tailieu["x2"])).ToString("0.###");
                    if (item.Rows.Count == 0)
                    {
                        tailieu["x74"] = Request.getValue("thoigian");
                        if (tailieu["x74"].isDateVN() == false)
                        {
                            ViewBag.Error = $"Ngày không đúng định dạng Ngày/Tháng/Năm x74: {tailieu["x74"]}";
                            return View();
                        }
                        tailieu["ngay"] = tailieu["x74"].getFromDateVN().toTimestamp().ToString();
                        tailieu["userid"] = $"{Session["iduser"]}";
                        tailieu["ma_tinh"] = Request.getValue("matinh").Trim();
                        tailieu["id"] = objectid;
                        dbBaoCao.Update("bctuandocx", tailieu, "replace");
                    }
                    else
                    {
                        var v = new List<string>();
                        foreach (var key in tailieu.Keys) { v.Add($"{key}='{tailieu[key].sqliteGetValueField()}'"); }
                        var tsql = $"UPDATE bctuandocx SET {string.Join(", ", v)} WHERE id='{objectid.sqliteGetValueField()}'";
                        dbBaoCao.Execute(tsql);
                    }
                    return Content($"<div class=\"alert alert-info\">Lưu thành công ({timeStart.getTimeRun()})</div>");
                }
                tmp = $"{Session["idtinh"]}".Trim();
                ViewBag.tinhSelect = tmp;
                tmp = $"{Session["nhom"]}".Trim() == "0" ? "WHERE id NOT IN ('', '00')" : $"WHERE id = '{tmp}'";
                var dmTinh = AppHelper.dbSqliteMain.getDataTable($"SELECT id,ten FROM dmtinh {tmp} ORDER BY tt, ten");
                if (dmTinh.Rows.Count == 0) { ViewBag.Error = "Bạn chưa chọn hoặc được cấp tỉnh hoạt động"; return View(); }
                ViewBag.dmTinh = dmTinh;
                if (string.IsNullOrEmpty(objectid) == false)
                {
                    var db = BuildDatabase.getDbSQLiteBaoCao();
                    var data = db.getDataTable($"SELECT * FROM bctuandocx WHERE id='{objectid.sqliteGetValueField()}';");
                    if (data.Rows.Count == 0) { ViewBag.Error = $"Báo cáo có mã '{objectid}' không tồn tại hoặc bị xoá trên hệ thống"; return View(); }
                    ViewBag.data = data.Rows[0];
                }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); return View(); }
            return View();
        }

        public ActionResult Tuan()
        {
            if ($"{Session["iduser"]}" == "") { ViewBag.Error = keyMSG.NotLoginAccess; return View(); }
            var mode = Request.getValue("mode");
            string tmp = "";
            if (mode == "")
            {
                tmp = $"{Session["idtinh"]}".Trim();
                ViewBag.tinhSelect = tmp;
                tmp = $"{Session["nhom"]}".Trim() == "0" ? "WHERE id NOT IN ('', '00')" : $"WHERE id = '{tmp}'";
                var dmTinh = AppHelper.dbSqliteMain.getDataTable($"SELECT id,ten FROM dmtinh {tmp} ORDER BY tt, ten");
                if (dmTinh.Rows.Count == 0) { ViewBag.Error = "Bạn chưa chọn hoặc được cấp tỉnh hoạt động"; return View(); }
                ViewBag.dmTinh = dmTinh;
                return View();
            }
            if (mode == "download")
            {
                try
                {
                    tmp = Request.getValue("idobject");
                    var dbBaoCao = BuildDatabase.getDbSQLiteBaoCao();
                    if (Request.getValue("type") == "xlsx")
                    {
                        var pl1 = dbBaoCao.getDataTable($"SELECT * FROM sheetpl01 WHERE id_bc='{tmp.sqliteGetValueField()}'");
                        pl1.TableName = "sheetpl01";
                        if (pl1.Rows.Count == 0) { ViewBag.Error = $"Báo cáo có ID '{tmp}' không tồn tại hoặc bị xoá trong hệ thống"; return View(); }
                        var pl2 = dbBaoCao.getDataTable($"SELECT * FROM sheetpl02 WHERE id_bc='{tmp.sqliteGetValueField()}'");
                        pl1.TableName = "sheetpl02";
                        var pl3 = dbBaoCao.getDataTable($"SELECT * FROM sheetpl03 WHERE id_bc='{tmp.sqliteGetValueField()}'");
                        pl1.TableName = "sheetpl03";
                        XSSFWorkbook xlsx = XLSX.exportExcel(new DataTable[] { pl1, pl2, pl3 });
                        var output = xlsx.WriteToStream();
                        return File(output.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"PL{tmp}.xlsx");
                    }
                    else
                    {
                        var data = dbBaoCao.getDataTable($"SELECT * FROM bctuandocx WHERE id='{tmp.sqliteGetValueField()}'");
                        if (data.Rows.Count == 0) { ViewBag.Error = $"Báo cáo có ID '{tmp}' không tồn tại hoặc bị xoá trong hệ thống"; return View(); }
                        var tailieu = new Dictionary<string, string>();
                        foreach (DataColumn c in data.Columns)
                        {
                            tailieu.Add("{" + c.ColumnName.ToUpper() + "}", $"{data.Rows[0][c]}");
                        }
                        string pathFileTemplate = Server.MapPath("~/App_Data/baocaotuan.docx");
                        if (System.IO.File.Exists(pathFileTemplate) == false)
                        {
                            ViewBag.Error = "Không tìm thấy tập tin mẫu báo cáo 'baocaotuan.docx' trong thư mục App_Data"; return View();
                        }
                        string thoigian = ((long)data.Rows[0]["ngay"]).toDateTime().ToString("yyyyMMdd");
                        using (var fileStream = new FileStream(pathFileTemplate, FileMode.Open, FileAccess.Read))
                        {
                            var document = new XWPFDocument(fileStream);
                            foreach (var paragraph in document.Paragraphs)
                            {
                                foreach (var run in paragraph.Runs)
                                {
                                    tmp = run.ToString();
                                    // Sử dụng Regex để tìm tất cả các match
                                    MatchCollection matches = Regex.Matches(tmp, "{x[0-9]+}", RegexOptions.IgnoreCase);
                                    foreach (Match match in matches) { tmp = tmp.Replace(match.Value, tailieu.getValue(match.Value, "", true)); }
                                    run.SetText(tmp, 0);
                                }
                            }
                            MemoryStream memoryStream = new MemoryStream();
                            document.Write(memoryStream);
                            memoryStream.Position = 0;
                            return File(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{data.Rows[0]["ma_tinh"]}_{thoigian}.docx");
                        }
                    }
                }
                catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); return View(); }
            }
            if (mode == "taive")
            {
                tmp = Request.getValue("x2");
                if (Regex.IsMatch(tmp, @"^\d+$") == false) { ViewBag.Error = $"Số của quyết định giao dự toán không đúng {tmp}"; return View(); }
                tmp = Request.getValue("x3");
                if (Regex.IsMatch(tmp, @"^\d+$") == false) { ViewBag.Error = $"Tổng số tiền các dòng quyết định năm nay không đúng {tmp}"; return View(); }
                string matinh = Request.getValue("matinh");
                string ngay = Request.getValue("thoigian");
                if (ngay.isDateVN() == false) { ViewBag.Error = $"Thời gian không đúng định dạng ngày/tháng/năm '{ngay}'"; return View(); }
                DateTime ngayTime = ngay.getFromDateVN();
                string thoigian = ngayTime.ToString("yyyyMMdd");

                var tailieu = buildBaoCaoTuan(ngayTime, matinh, $"{Session["iduser"]}", Request.getValue("x2"), Request.getValue("x3"), Request.getValue("x67"), Request.getValue("x68"), Request.getValue("x69"), Request.getValue("x70"));
                if (tailieu.ContainsKey("Error")) { ViewBag.Error = tailieu["Error"]; return View(); }
                string pathFileTemplate = Server.MapPath("~/App_Data/baocaotuan.docx");
                if (System.IO.File.Exists(pathFileTemplate) == false)
                {
                    ViewBag.Error = "Không tìm thấy tập tin mẫu báo cáo 'baocaotuan.docx' trong thư mục App_Data";
                    return View();
                }
                using (var fileStream = new FileStream(pathFileTemplate, FileMode.Open, FileAccess.Read))
                {
                    var document = new XWPFDocument(fileStream);
                    foreach (var paragraph in document.Paragraphs)
                    {
                        foreach (var run in paragraph.Runs)
                        {
                            tmp = run.ToString();
                            // Sử dụng Regex để tìm tất cả các match
                            MatchCollection matches = Regex.Matches(tmp, "{x[0-9]+}", RegexOptions.IgnoreCase);
                            foreach (Match match in matches) { tmp = tmp.Replace(match.Value, tailieu.getValue(match.Value, "", true)); }
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

        private string getPosition(string mavung, string matinh, string field, List<DataRow> data)
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

        private Dictionary<string, string> buildBCTuanB02(int iKey, string fieldChiBQ, string fieldTongLuot, string fieldChiBQChung, string mavung, string matinh, DataRow rowTinh, DataRow rowTQ, List<DataRow> data)
        {
            var d = new Dictionary<string, string>();
            var keys = new List<string>();
            for (int i = iKey; i <= (iKey + 6); i++) { keys.Add("{X" + i.ToString() + "}"); }
            /* X33 = Chi bình quân nội trú X33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
            d.Add(keys[0], rowTinh[fieldChiBQ].ToString()); /* "chi_bq_noi" */
            /* X34 = bình quân toàn quốc X34={cột K (CHI_BQ_NOI), dòng MA_TINH=00}; */
            d.Add(keys[1], rowTQ[fieldChiBQ].ToString());
            /* X35 = Số chênh lệch X35={đoạn văn tùy thuộc X33> hay < X34. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            d.Add(keys[2], "bằng");
            var so1 = double.Parse(d[keys[0]]);
            var so2 = double.Parse(d[keys[1]]);
            if (so1 > so2) { d[keys[2]] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { d[keys[2]] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X36= xếp thứ so toàn quốc X36={Sort cột K CHI_BQ_NOI cao xuống thấp và lấy thứ tự}; */
            d.Add(keys[3], getPosition("", matinh, fieldChiBQ, data));
            /* X37 = Bình quân vùng X37={tính toán: A-Tổng chi nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của các tỉnh cùng mã vùng. A=Total  (cột K (CHI_BQ_NOI) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
            d.Add(keys[4], "0");
            so2 = data.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<double>(fieldChiBQ));
            if (so2 != 0)
            {
                so1 = data.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => (r.Field<double>(fieldChiBQ) * r.Field<long>(fieldTongLuot)));
                d[keys[4]] = (so1 / so2).ToString();
            }
            /* X38 = số chênh lệch X38 ={đoạn văn tùy thuộc X33 > hay < X37. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            d.Add(keys[5], "bằng");
            so1 = double.Parse(d[keys[0]]);
            so2 = double.Parse(d[keys[4]]);
            if (so1 > so2) { d[keys[5]] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { d[keys[5]] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X39 đứng thứ so với vùng X39= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột K (CHI_BQ_NOI) cao –thấp và lấy thứ tự} */
            d.Add(keys[6], getPosition(mavung, matinh, fieldChiBQ, data));
            return d;
        }

        private Dictionary<string, string> buildBCTuanB26(int iKey, string field1, string field2, DataRow row)
        {
            var d = new Dictionary<string, string>();
            string key1 = "{X" + iKey.ToString() + "}", key2 = "{X" + (iKey + 1).ToString() + "}", key3 = "{X" + (iKey + 2).ToString() + "}";
            /* X46 Bình quân cột [x] dòng có mã tỉnh = 10}; */
            var x = (double)row[field1];
            d.Add(key1, row[field1].ToString());
            /* X47 số tương đối X47={nếu cột [x+1] dòng có mã tỉnh=10 là số dương, “tăng “ & cột [x+1] & “%”, không thì “giảm “ & cột [x+1] %}; */
            d.Add(key2, "bằng");
            var x1 = (double)row[field2]; /* s */
            if (x1 > 0) { d[key2] = $"tăng {x1.FormatCultureVN()}%"; }
            else { if (x1 < 0) { d[key2] = $"giảm {Math.Abs(x1).FormatCultureVN()}%"; } }
            /* X48 số tuyệt đối X48={nếu cột [x+1] là dương, “tăng “ & [cột [x] - (cột [x] / (cột [x+1] +100) *100 )] & “ đồng”, không thì “giảm “ & [cột [x]- (cột [x] / (cột [x+1]+100) *100 )] & “ đồng”} */
            d.Add(key3, "bằng");
            if (x1 > 0) { d[key3] = "tăng " + (x - (x / (x1 + 100) * 100)).FormatCultureVN() + " đồng"; }
            else { if (x1 < 0) { d[key3] = "giảm " + (x - (x / (x1 + 100) * 100)).FormatCultureVN() + " đồng"; } }
            return d;
        }

        private Dictionary<string, string> buildBCTuan02B26(int iKey, string field1, string field2, DataRow row)
        {
            var d = new Dictionary<string, string>();
            string key1 = "{X" + iKey.ToString() + "}", key2 = "{X" + (iKey + 1).ToString() + "}", key3 = "{X" + (iKey + 2).ToString() + "}";
            /* X61 Chỉ định xét nghiệm X61={cột AD, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
            var so1 = ((double)row[field1] * 100);
            d.Add(key1, so1.ToString());
            /* X62 số tương đối X62={cột AE dòng có mã tỉnh=10 & “%”}; */
            d.Add(key2, row[field2].ToString().FormatCultureVN() + "%");
            /* X63 = số tuyệt đối X63 {tính toán: [X61 trừ đi (X61 chia (cột AE+100)*100)] & “bệnh nhân”} */
            var so2 = (double)row[field2];
            d.Add(key3, (so1 - (so1 / (so2 + 100) * 100)).FormatCultureVN() + " bệnh nhân");
            return d;
        }

        private Dictionary<string, string> buildBaoCaoTuan(DateTime ngayTime, string matinh, string iduser, string x2, string x3, string x67, string x68, string x69, string x70)
        {
            var tailieu = new Dictionary<string, string>();
            try
            {
                if (Regex.IsMatch(x2, @"^\d+(\.\d+)?$") == false) { x2 = "0"; }
                if (Regex.IsMatch(x3, @"^\d+(\.\d+)?$") == false) { x3 = "0"; }
                string thoigian = ngayTime.ToString("yyyyMMdd");
                string thang = ngayTime.ToString("MM");
                string nam = ngayTime.ToString("yyyy");
                double so1 = 0; double so2 = 0;
                var tmpD = new Dictionary<string, string>();
                string tsql = string.Empty;
                string tmp = string.Empty;

                /* Bỏ qua các vùng */
                tsql = $"SELECT p1.* FROM b02chitiet p1 INNER JOIN b02 ON p1.id2=b02.id WHERE b02.tu_thang={thang} AND b02.den_thang={thang} AND b02.nam={nam} AND b02.cs='0' AND p1.ma_tinh NOT LIKE 'V%'";
                var b02TQ = AppHelper.dbSqliteWork.getDataTable(tsql).AsEnumerable().ToList();
                if (b02TQ.Count() == 0) { throw new Exception("B02 Toàn Quốc không có dữ liệu phù hợp truy vấn"); }
                /* Bỏ qua các vùng */
                tsql = $"SELECT p1.* FROM b26chitiet p1 INNER JOIN b26 ON p1.id2=b26.id WHERE b26.thoigian = '{thoigian}' AND b26.cs='0' AND p1.ma_tinh NOT LIKE 'V%'";
                var b26TQ = AppHelper.dbSqliteWork.getDataTable(tsql).AsEnumerable().ToList();
                if (b26TQ.Count() == 0) { throw new Exception("B26 Toàn quốc không có dữ liệu phù hợp truy vấn"); }

                var dataTinhB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == matinh).FirstOrDefault();
                if (dataTinhB02 == null) { throw new Exception("B02 không có dữ liệu tỉnh phù hợp truy vấn"); }
                var dataTinhB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == matinh).FirstOrDefault();
                if (dataTinhB26 == null) { throw new Exception("B26 không có dữ liệu tỉnh phù hợp truy vấn"); }

                var dataTQB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault();
                if (dataTQB02 == null) { throw new Exception("B02 không có dữ liệu toàn quốc phù hợp truy vấn"); }
                var dataTQB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault();
                if (dataTQB26 == null) { throw new Exception("B26 không có dữ liệu toàn quốc phù hợp truy vấn"); }

                /* Bỏ Toàn quốc ra khỏi danh sách */
                b02TQ = b02TQ.Where(p => p.Field<string>("ma_tinh") != "00").ToList();
                b26TQ = b26TQ.Where(p => p.Field<string>("ma_tinh") != "00").ToList();

                string mavung = dataTinhB02["ma_vung"].ToString();

                /* X1 = {cột R (T-BHTT) bảng B02_TOANQUOC } */
                tailieu.Add("{X1}", dataTinhB02["t_bhtt"].ToString());
                /* X2 = {“ Quyết định số: Nếu không tìm thấy dòng nào của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán thì “TW chưa giao dự toán, tạm lấy theo dự toán năm trước”, nếu thấy lấy số ký hiệu các dòng QĐ của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán} */
                tailieu.Add("{X2}", x2);
                /* X3 = {Như trên, ko thấy thì lấy tổng tiền các dòng dự toán năm trước, thấy thì lấy tổng số tiền các dòng quyết định năm nay} */
                tailieu.Add("{X3}", x3);
                /* X4={X1/X2 %} So sánh với dự toán, tỉnh đã sử dụng */
                so2 = double.Parse(x2);
                if (so2 == 0) { tailieu.Add("{X4}", "0"); }
                else { tailieu.Add("{X4}", (double.Parse(tailieu["{X1}"]) / so2).ToString()); }

                /* X5 = {Cột tyle_noitru, dòng MA_TINH=10} bảng B02_TOANQUOC */
                tailieu.Add("{X5}", dataTinhB02["tyle_noitru"].ToString());
                /* X6 = {Cột tyle_noitru, dòng MA_TINH=00} bảng B02_TOANQUOC */
                tailieu.Add("{X6}", dataTQB02["tyle_noitru"].ToString());
                /* X7 = {đoạn văn tùy thuộc X5> hay < X6. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X7}", "bằng");
                so1 = (double)dataTinhB02["tyle_noitru"];
                so2 = (double)dataTQB02["tyle_noitru"];
                if (so1 > so2) { tailieu["{X7}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
                else { if (so1 < so2) { tailieu["{X7}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
                /* X8={Sort cột G (TYLE_NOITRU) cao xuống thấp và lấy thứ tự}; */
                var sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
                int position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("X8", position.ToString());
                /* X9 ={tính toán: total cột F (TONG_LUOT_NOI) chia cho Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG=mã vùng của tỉnh báo cáo}; */
                tailieu.Add("{X9}", "0");
                so2 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung).Sum(row => row.Field<long>("tong_luot"));
                if (so2 != 0)
                {
                    so1 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung).Sum(row => row.Field<long>("tong_luot_noi"));
                    tailieu["{X9}"] = (so1 / so2).ToString();
                }
                /* X10 ={đoạn văn tùy thuộc X5> hay < X9. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X10}", "bằng");
                so1 = (double)dataTinhB02["tyle_noitru"];
                so2 = double.Parse(tailieu["{X9}"]); tailieu["{X9}"] = tailieu["{X9}"].ToString();
                if (so1 > so2) { tailieu["{X10}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
                else { if (so1 < so2) { tailieu["{X10}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
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
                so1 = (double)dataTinhB02["ngay_dtri_bq"];
                so2 = (double)dataTQB02["ngay_dtri_bq"];
                if (so1 > so2) { tailieu["{X14}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
                else { if (so1 < so2) { tailieu["{X14}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
                /* X15 = xếp thứ so toàn quốc X15={Sort cột H (NGAY_DTRI_BQ) cao xuống thấp và lấy thứ tự}; */
                sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X15}", position.ToString());
                /* X16 = Bình quân vùng X16 ={tính toán: A-Tổng ngày điều trị nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của cá tỉnh cùng mã vùng. A=Total(cột H (NGAY_DTRI_BQ) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                tailieu.Add("{X16}", "0");
                so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<long>("tong_luot_noi"));
                if (so2 != 0)
                {
                    so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => (r.Field<double>("ngay_dtri_bq") * r.Field<long>("tong_luot_noi")));
                    tailieu["{X16}"] = (so1 / so2).ToString();
                }
                /* X17 = Số chênh lệch X17 ={đoạn văn tùy thuộc X12> hay < X16. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                tailieu.Add("{X17}", "bằng");
                so1 = (double)dataTinhB02["ngay_dtri_bq"];
                so2 = double.Parse(tailieu["{X16}"]); tailieu["{X16}"] = tailieu["{X16}"].ToString();
                if (so1 > so2) { tailieu["{X17}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
                else { if (so1 < so2) { tailieu["{X17}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
                /* X18 = đứng thứ so với vùng X18 = {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột H (NGAY_DTRI_BQ) cao –thấp và lấy thứ tự} */
                sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                    .OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
                position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1;
                tailieu.Add("{X18}", position.ToString());

                /* X19 = Chi bình quân chung X19={Cột I (CHI_BQ_CHUNG), dòng MA_TINH=10}; */
                tmpD = buildBCTuanB02(19, "chi_bq_chung", "tong_luot", "chi_bq_chung", mavung, matinh, dataTinhB02, dataTQB02, b02TQ);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X26 = Chi bình quân ngoại trú X26={Cột J (CHI_BQ_NGOAI), dòng MA_TINH=10}; */
                tmpD = buildBCTuanB02(26, "chi_bq_ngoai", "tong_luot_ngoai", "chi_bq_chung", mavung, matinh, dataTinhB02, dataTQB02, b02TQ);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X33 = Chi bình quân nội trú X33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
                tmpD = buildBCTuanB02(33, "chi_bq_noi", "tong_luot_noi", "chi_bq_chung", mavung, matinh, dataTinhB02, dataTQB02, b02TQ);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }

                /* ----- Dữ liệu X40 trở lên lọc dữ liệu tù B26 ------- */
                /* X40 = Bình quân xét nghiệm X40= {cột P (bq_xn) dòng có mã tỉnh = 10}; B26 */
                tmpD = buildBCTuanB26(40, "bq_xn", "bq_xn_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X43 Bình quân CĐHA X43= {cột R(bq_cdha) dòng có mã tỉnh =10}; */
                tmpD = buildBCTuanB26(43, "bq_cdha", "bq_cdha_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X46 Bình quân thuốc X46= {cột T(bq_thuoc) dòng có mã tỉnh =10}; */
                tmpD = buildBCTuanB26(46, "bq_thuoc", "bq_thuoc_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X49 Bình quân chi phẫu thuật X49= {cột V(bq_pt) dòng có mã tỉnh =10}; */
                tmpD = buildBCTuanB26(49, "bq_pt", "bq_pt_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X52 Bình quân chi thủ thuật X52= {cột X(bq_tt) dòng có mã tỉnh =10}; */
                tmpD = buildBCTuanB26(52, "bq_tt", "bq_tt_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X55 Bình quân chi vật tư y tế X55= {cột Z(bq_vtyt) dòng có mã tỉnh =10}; */
                tmpD = buildBCTuanB26(55, "bq_vtyt", "bq_vtyt_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X58 Bình quân chi tiền giường X58= {cột AB(bq_giuong) dòng có mã tỉnh =10}; */
                tmpD = buildBCTuanB26(58, "bq_giuong", "bq_giuong_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }

                /* X61 Chỉ định xét nghiệm X61={cột AD, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
                tmpD = buildBCTuan02B26(61, "chi_dinh_xn", "chi_dinh_xn_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }
                /* X64 =  Chỉ định CĐHA X64={cột AF, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
                tmpD = buildBCTuan02B26(64, "chi_dinh_cdha", "chi_dinh_cdha_tang", dataTinhB26);
                foreach (var d in tmpD) { tailieu.Add(d.Key, d.Value); }

                /* X67 Công tác kiểm soát chi X67={lần đầu lập BC sẽ rỗng, người dùng tự trình bày văn bản, lưu lại ở bảng dữ liệu kết quả báo cáo, kỳ sau sẽ tự động lấy từ kỳ trước, để người dùng kế thừa, sửa và lưu dùng cho kỳ này và kỳ sau} */
                tailieu.Add("{X67}", x67);
                /* X68 Công tác thanh, quyết toán năm X68={tương tự X67} */
                tailieu.Add("{X68}", x68);
                /* X69 Phương hướng kỳ tiếp theo X69={tương tự X67} */
                tailieu.Add("{X69}", x69);
                /* X70 Khó khăn, vướng mắc, đề xuất (nếu có) X70={tương tự X67} */
                tailieu.Add("{X70}", x70);

                /* X71 = {cột S T_BHTT_NOI bảng B02_TOANQUOC } */
                tailieu.Add("{X71}", dataTinhB02["t_bhtt_noi"].ToString());
                /* X72 = {cột T T_BHTT_NGOAI bảng B02_TOANQUOC } */
                tailieu.Add("{X72}", dataTinhB02["t_bhtt_ngoai"].ToString());
                /* X73 Lấy tên tỉnh */
                tmp = $"{AppHelper.dbSqliteMain.getValue($"SELECT ten FROM dmTinh WHERE id='{matinh.sqliteGetValueField()}'")}";
                tailieu.Add("{X73}", tmp);
                /* X74 Lấy ngày chọn báo cáo */
                tailieu.Add("{X74}", ngayTime.ToString("dd/MM/yyyy"));

                string timeCreate = DateTime.Now.toTimestamp().ToString();
                tailieu.Add("id", $"{thoigian}|{iduser}");
                tailieu.Add("ma_tinh", matinh);
                tailieu.Add("userid", iduser);
                tailieu.Add("ngay", ngayTime.toTimestamp().ToString());
                tailieu.Add("timecreate", timeCreate);
                var dbBaoCao = BuildDatabase.getDbSQLiteBaoCao();
                dbBaoCao.Update("bctuandocx", tailieu, "replace");
                dbBaoCao.Execute($"DELETE FROM sheetpl01 WHERE id_bc='{tailieu["id"]}'; DELETE FROM sheetpl02 WHERE id_bc='{tailieu["id"]}'; DELETE FROM sheetpl03 WHERE id_bc='{tailieu["id"]}'; ");
                /* Tạo Phục lục sheetpl01 */
                tsql = $@"SELECT '{tailieu["id"]}' AS id_bc, '{matinh}' AS idtinh, p1.ma_tinh, p1.ten_tinh, p1.ma_vung, p1.tyle_noitru, p1.ngay_dtri_bq, p1.chi_bq_chung, p1.chi_bq_ngoai, p1.chi_bq_noi, '{iduser}' AS userid, '{timeCreate}' AS timecreate
                    FROM b02chitiet p1 INNER JOIN b02 ON p1.id2=b02.id WHERE b02.tu_thang={thang} AND b02.den_thang={thang} AND b02.nam={nam} AND b02.cs='0'";
                var pl = AppHelper.dbSqliteWork.getDataTable(tsql);
                dbBaoCao.Insert("sheetpl01", pl);

                /* Tạo Phục lục sheetpl02 */
                tsql = $@"SELECT '{tailieu["id"]}' AS id_bc, '{matinh}' AS idtinh, p1.ma_tinh, p1.ten_tinh, p1.ma_vung, p1.bq_xn AS chi_bq_xn, p1.bq_cdha AS chi_bq_cdha, p1.bq_thuoc AS chi_bq_thuoc, p1.bq_ptt AS chi_bq_pttt, p1.bq_vtyt AS chi_bq_vtyt, p1.bq_giuong AS chi_bq_giuong, p1.ngay_ttbq, '{iduser}' AS userid, '{timeCreate}' AS timecreate
                    FROM b04chitiet p1 INNER JOIN b04 ON p1.id2=b04.id WHERE b04.tu_thang={thang} AND b04.den_thang={thang} AND b04.nam={nam} AND b04.cs='0'";
                pl = AppHelper.dbSqliteWork.getDataTable(tsql);
                dbBaoCao.Insert("sheetpl02", pl);

                /* Tạo Phục lục sheetpl03 */
                tsql = $@"SELECT '{tailieu["id"]}' AS id_bc, '{matinh}' AS idtinh, p1.ma_cskcb, p1.ten_cskcb, p1.tyle_noitru, p1.ngay_dtri_bq, p1.chi_bq_chung, p1.chi_bq_ngoai, p1.chi_bq_noi, '{iduser}' AS userid, '{timeCreate}' AS timecreate
                        FROM b02chitiet p1 INNER JOIN b02 ON p1.id2=b02.id WHERE b02.tu_thang={thang} AND b02.den_thang={thang} AND b02.nam={nam} AND b02.cs='1'";
                pl = AppHelper.dbSqliteWork.getDataTable(tsql);
                dbBaoCao.Insert("sheetpl03", pl);

                dbBaoCao.Close();
                return tailieu;
            }
            catch (Exception ex) { tailieu.Add("Error", ex.getLineHTML()); return tailieu; }
        }
    }
}