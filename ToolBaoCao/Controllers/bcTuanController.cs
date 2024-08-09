using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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
            /* Tạo thư mục tạm */
            string folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bctuan", $"{Session["idtinh"]}_{Session["iduser"]}".GetMd5Hash());
            var d = new System.IO.DirectoryInfo(folderTemp);
            if (d.Exists == false) { d.Create(); }
            return View();
        }

        public ActionResult Buoc2()
        {
            DateTime timeStart = DateTime.Now;
            string matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            if (Request.Files.Count == 0) { ViewBag.Error = "Không có tập tin dữ liệu nào đẩy lên"; return View(); }
            string id = $"{timeStart:yyyyMMddHHmmss}_{matinh}_{timeStart.Millisecond:000}";
            string folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bctuan", $"{matinh}_{Session["iduser"]}".GetMd5Hash());
            ViewBag.id = id;
            try
            {
                /* Xoá hết các File có trong thư mục */
                var time = timeStart.AddMinutes(-30);
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { if (item.LastWriteTime < time) { try { item.Delete(); } catch { } } }
                /* Khai báo dữ liệu tạm */
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                dbTemp.CreateTableImport();
                dbTemp.CreateTablePhucLucBaoCao();
                dbTemp.CreateTableBaoCao();
                /* Đọc và kiểm tra các tập tin */
                var list = new List<string>();
                var bieus = new List<string>();
                for (int i = 0; i < Request.Files.Count; i++)
                {
                    if (Path.GetExtension(Request.Files[i].FileName).ToLower() != ".xlsx") { throw new Exception($"Hệ thống chỉ hỗ trợ dữ liệu Excel 2007 (Type of file: Microsoft Excel Worksheet) trở lên '{Request.Files[i].FileName}'"); }
                    list.Add($"{Request.Files[i].FileName} ({Request.Files[i].ContentLength.getFileSize()})");
                    bieus.Add(readExcelbcTuan(dbTemp, Request.Files[i], Session, id, folderTemp, timeStart));
                }
                ViewBag.files = list;
                list = new List<string>();
                if (bieus.Contains("b02_00") == false) { list.Add("Thiếu biểu B02 toàn quốc;"); }
                if (bieus.Contains($"b02_{matinh}") == false) { list.Add($"Thiếu biểu B02 của Tỉnh có mã {matinh};"); }
                if (bieus.Contains("b04_00") == false) { list.Add("Thiếu biểu B04 toàn quốc;"); }
                if (bieus.Contains("b26_00") == false) { list.Add("Thiếu biểu B26 toàn quốc;"); }
                if (bieus.Contains($"b26_{matinh}") == false) { list.Add($"Thiếu biểu B26 của Tỉnh có mã {matinh};"); }
                if(list.Count > 0) { throw new Exception(string.Join("<br />", list)); }
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.getLineHTML();
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
            }
            return View();
        }

        private string readExcelbcTuan(dbSQLite dbConnect, HttpPostedFileBase inputFile, HttpSessionStateBase Session, string idBaoCao, string folderTemp, DateTime timeStart)
        {
            string messageError = "";
            var timeUp = timeStart.toTimestamp().ToString();
            var userID = $"{Session["iduser"]}".Trim();
            string bieu = ""; string matinhImport = "";
            string fileExtension = Path.GetExtension(inputFile.FileName);
            int sheetIndex = 0; int packetSize = 1000;
            int indexRow = 0; int indexColumn = 0; int maxRow = 0; int jIndex = 0;
            int fieldCount = 50; var tsql = new List<string>();
            var tmp = "";
            IWorkbook workbook = null;
            try
            {
                try
                {
                    workbook = new XSSFWorkbook(inputFile.InputStream);
                    /* if (fileExtension.ToLower() == ".xls") { workbook = new HSSFWorkbook(fs); }
                    else { workbook = new XSSFWorkbook(fs); } */
                }
                catch (Exception ex) { throw new Exception($"Lỗi tập tin '{inputFile.FileName}' sai định dạng : {ex.Message}"); }
                var sheet = workbook.GetSheetAt(sheetIndex);
                var tsqlv = new List<string>(); maxRow = sheet.LastRowNum;
                var cs = true;
                IRow row = null;
                for (; indexRow <= maxRow; indexRow++)
                {
                    row = sheet.GetRow(indexRow); if (row == null) { continue; }
                    /* Xác định tên biểu */
                    /* Xác định vị trí hàng bắt đầu có dữ liệu */
                    foreach (var c in row.Cells)
                    {
                        tmp = c.GetValueAsString().Trim().ToLower();
                        if (tmp.StartsWith("b26")) { bieu = "b26"; }
                        if (tmp.StartsWith("b04")) { bieu = "b04"; }
                        if (tmp.StartsWith("b02")) { bieu = "b02"; }
                        if (tmp == "ma_tinh") { indexColumn = c.ColumnIndex; break; }
                    }
                    if (tmp == "ma_tinh") { break; }
                }
                if (indexRow >= maxRow) { throw new Exception("Không có dữ liệu"); }
                string pattern = "^20[0-9][0-9]$";
                int indexRegex = 3; int tmpInt = 0;
                /*
                 * Bắt đầu đọc dữ liệu
                 */
                /*
                 * - Đọc thông số biểu
                 * Biểu B04: ma_tinh ma_loai_kcb tu_thang den_thang nam loai_bv kieubv loaick hang_bv tuyen cs + userID
                 * Biểu B26: ma_tinh	loai_kcb	thoi_gian	loai_bv	kieubv	loaick	hang_bv	tuyen	loai_so_sanh	cs
                 */
                switch (bieu)
                {
                    /* Kiểm tra năm */
                    case "b02": fieldCount = 11; indexRegex = 4; pattern = "^20[0-9][0-9]$"; break;
                    /* Kiểm tra năm */
                    case "b04": fieldCount = 11; indexRegex = 3; pattern = "^20[0-9][0-9]$"; break;
                    /* Kiểm tra thoigian */
                    case "b26": fieldCount = 10; indexRegex = 2; pattern = "^20[0-9][0-9][0-1][0-9][0-3][0-9]$"; break;
                    default: fieldCount = 11; break;
                }
                indexRow++; /* Lấy dòng có dữ liệu */
                var listValue = new List<string>();
                row = sheet.GetRow(indexRow);
                for (jIndex = indexColumn; jIndex < indexColumn + fieldCount; jIndex++)
                {
                    ICell c = row.GetCell(jIndex);
                    listValue.Add(c.GetValueAsString().Trim());
                }
                /* Có phải là cơ sở không? */
                tmpInt = (fieldCount - 1);
                if (listValue[tmpInt] == "true" && listValue[tmpInt] != "1") { listValue[tmpInt] = "1"; }
                else { listValue[tmpInt] = "0"; cs = false; }
                /* Kiểm tra có đúng dữ liệu không */
                if (Regex.IsMatch(listValue[indexRegex], pattern) == false) { throw new Exception($"dữ liệu không đúng cấu trúc (năm, thời gian): {listValue[indexRegex]}"); }
                matinhImport = listValue[0];
                /* Lấy danh sách cột, bỏ cột ID */
                var allColumns = dbConnect.getColumns(bieu).Select(p => p.ColumnName).ToList();
                allColumns.RemoveAt(0);
                /* Thêm UserID */
                listValue.Add(userID);
                listValue.Add(timeUp);
                listValue.Add(idBaoCao);
                tsql.Add($"INSERT INTO {bieu} ({string.Join(",", allColumns)}) VALUES ('{string.Join("','", listValue)}');");
                /**
                 * Lấy dữ liệu chi tiết
                 */
                allColumns = dbConnect.getColumns(bieu + "chitiet").Select(p => p.ColumnName).ToList();
                allColumns.RemoveAt(0);
                /* id2 matinh tentinh macskcb tencskcb */
                if (cs) { allColumns.RemoveAt(1); allColumns.RemoveAt(1); } /* Loại bỏ ma_tinh, ten_tinh */
                else { allColumns.RemoveAt(3); allColumns.RemoveAt(3); } /* Loại bỏ ma_cskcb, ten_cskcb */
                var fieldNumbers = new List<int>();
                /* indexRegex + 1 do thêm cột {@id2} ID vào đằng trước */
                switch (bieu)
                {
                    /* Kiểm tra tổng số lượt KCB */
                    case "b02":
                        fieldCount = 20; indexRegex = 3 + 1; pattern = "^[0-9]+$";
                        fieldNumbers = new List<int>() { 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 };
                        break;
                    /* Kiểm tra ngày TTBQ */
                    case "b04":
                        fieldCount = 11; indexRegex = 9 + 1; pattern = "^[0-9]+[.,][0-9]+$|^[0-9]+$";
                        fieldNumbers = new List<int>() { 3, 4, 5, 6, 7, 8, 9, 10 };
                        break;
                    /* Kiểm tra BQ chung trong kỳ */
                    case "b26":
                        fieldCount = 34; indexRegex = 7 + 1; pattern = "^[0-9]+[.,][0-9]+$|^[0-9]+$";
                        fieldNumbers = new List<int>() { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33 };
                        break;

                    default: fieldCount = 11; break;
                }
                /* Bỏ qua dòng tiêu đề */
                indexRow++;
                var tsqlVaues = new List<string>();
                for (; indexRow <= maxRow; indexRow++)
                {
                    if (tsqlVaues.Count > packetSize)
                    {
                        tsql.Add($"INSERT INTO {bieu}chitiet ({string.Join(",", allColumns)}) VALUES {string.Join(",", tsqlVaues)};");
                        tsqlVaues = new List<string>();
                    }
                    /* Dòng không tồn tại */
                    row = sheet.GetRow(indexRow); if (row == null) { continue; }
                    /* Số cột ít hơn số trường cần lấy dữ liệu */
                    /* if ((int)row.LastCellNum < fieldCount) { continue; } */
                    /* Cột đầu tiên không phải là matinh dạng số */
                    string ma = row.GetCell(indexColumn).GetValueAsString().Trim();

                    if (Regex.IsMatch(ma, "^[0-9]+$|^V[0-9]+$") == false) { continue; }
                    /* Xây dựng tsql VALUES */
                    listValue = new List<string>() { "0", ma.sqliteGetValueField() };
                    for (jIndex = indexColumn + 1; jIndex < (indexColumn + fieldCount); jIndex++)
                    {
                        ICell c = row.GetCell(jIndex);
                        listValue.Add(c.GetValueAsString().Trim().sqliteGetValueField());
                    }
                    /* Cột lấy dữ liệu không đúng định dạng bỏ qua */
                    if (Regex.IsMatch(listValue[indexRegex], pattern) == false) { continue; }
                    /* Trường hợp trường số để trống thì cho bằng 0 */
                    foreach (int i in fieldNumbers) { if (Regex.IsMatch(listValue[i], "^[0-9]+$|^[0-9]+[.][0-9]+$") == false) { listValue[i] = "0"; } }
                    listValue.Add(idBaoCao);
                    tsqlVaues.Add($"('{string.Join("','", listValue)}')");
                }
                if (tsqlVaues.Count > 0) { tsql.Add($"INSERT INTO {bieu}chitiet ({string.Join(",", allColumns)}) VALUES {string.Join(",", tsqlVaues)};"); }
                tmp = string.Join(Environment.NewLine, tsql);
                System.IO.File.WriteAllText(Path.Combine(folderTemp, $"id{idBaoCao}_{bieu}_{matinhImport}.sql"), tmp);
                dbConnect.Execute(tmp);
                if (tsql.Count < 2) { throw new Exception("Không có dữ liệu chi tiết"); }
            }
            catch (Exception ex2) { messageError = $"Lỗi trong quá trình đọc, nhập dữ liệu từ Excel '{inputFile.FileName}': {ex2.getLineHTML()} <br />{tmp}"; }
            finally
            {
                if (workbook != null) { workbook.Close(); workbook = null; }
            }
            if (messageError != "") { throw new Exception(messageError); }
            inputFile.SaveAs(Path.Combine(folderTemp, $"id{idBaoCao}_{bieu}_{matinhImport}{fileExtension}"));
            return $"{bieu}_{matinhImport}";
        }

        public ActionResult Buoc3()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            ViewBag.id = Request.getValue("idobject");
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