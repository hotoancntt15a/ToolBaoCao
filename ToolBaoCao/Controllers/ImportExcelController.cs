using NPOI.HSSF.UserModel;
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
    public class ImportExcelController : ControllerCheckLogin
    {
        // GET: ImportExcel
        public ActionResult Index()
        {
            ViewBag.Title = "Quản lý nhập dữ liệu Excel";
            return View();
        }

        public ActionResult TruyVan()
        {
            string ngay1 = Request.getValue("ngay1");
            string ngay2 = Request.getValue("ngay2");
            var time1 = DateTime.Now;
            var time2 = time1;
            try
            {
                if (ngay1.isDateVN(out time1) == false) { throw new Exception($"Từ ngày không đúng định dạng ngày/tháng/năm '{ngay1}'"); }
                if (ngay2.isDateVN(out time2) == false) { throw new Exception($"Đến ngày không đúng định dạng ngày/tháng/năm '{ngay2}'"); }
                if (time2 < time1) { throw new Exception($"Đến ngày '{ngay2}' < từ ngày {ngay1}"); }
                var ts = time2 - time1;
                if (ts.Days > 365) { throw new Exception("Hệ thống không hỗ trợ truy vấn quá 365 ngày"); }
                var db = AppHelper.dbSqliteWork;
                if (time1.Year == time2.Year)
                {
                    ViewBag.b02 = db.getDataTable($"SELECT *, datetime(timeup, 'auto', '+7 hour') AS timeup2 FROM b02 WHERE nam = {time1.Year} AND (den_thang <= {time2.Month} AND tu_thang >= {time1.Month}) ORDER BY timeup DESC");
                    ViewBag.b04 = db.getDataTable($"SELECT *, datetime(timeup, 'auto', '+7 hour') AS timeup2 FROM b04 WHERE nam = {time1.Year} AND (den_thang <= {time2.Month} AND tu_thang >= {time1.Month}) ORDER BY timeup DESC");
                }
                else
                {
                    ViewBag.b02 = db.getDataTable($"SELECT *, datetime(timeup, 'auto', '+7 hour') AS timeup2 FROM b02 WHERE (nam > {time1.Year} AND tu_thang > {time1.Month}) AND (nam <= {time2.Year} AND den_thang >= {time2.Month}) ORDER BY timeup DESC");
                    ViewBag.b04 = db.getDataTable($"SELECT *, datetime(timeup, 'auto', '+7 hour') AS timeup2 FROM b04 WHERE (nam > {time1.Year} AND tu_thang > {time1.Month}) AND (nam <= {time2.Year} AND den_thang >= {time2.Month}) ORDER BY timeup DESC");
                }
                if (time1 == time2) { ViewBag.b26 = db.getDataTable($"SELECT *, datetime(timeup, 'auto', '+7 hour') AS timeup2 FROM b26 WHERE thoigian = {time1:yyyyMMdd} ORDER BY timeup DESC"); }
                else { ViewBag.b26 = db.getDataTable($"SELECT *, datetime(timeup, 'auto', '+7 hour') AS timeup2 FROM b26 WHERE thoigian >= {time1:yyyyMMdd} AND thoigian <= {time2:yyyyMMdd} ORDER BY timeup DESC"); }
            }
            catch (Exception ex) { return Content($"<div class=\"alert alert-warning\">Lỗi: {ex.getLineHTML()}</div>"); }
            return View();
        }

        public ActionResult Update(string bieu, HttpPostedFileBase inputfile)
        {
            if (Session["iduser"] == null) { ViewBag.Error = "Bạn chưa đăng nhập"; return View(); }
            DateTime timeStart = DateTime.Now;
            var timeUp = timeStart.toTimestamp().ToString();
            var userID = $"{Session["iduser"]}".Trim();
            ViewBag.data = "Đang thao tác";
            if (string.IsNullOrEmpty(bieu)) { ViewBag.Error = "Tham số biểu nhập không có chỉ định"; return View(); }
            if (inputfile == null) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            if (inputfile.ContentLength == 0) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            string fileName = Path.GetFileName(inputfile.FileName);
            string fileExtension = Path.GetExtension(inputfile.FileName);
            string fileNameSave = $"{userID}_{bieu}{fileExtension}";
            int sheetIndex = 0; int packetSize = 1000;
            int indexRow = 0; int indexColumn = 0; int maxRow = 0; int jIndex = 0;
            int fieldCount = 50; var tsql = new List<string>();
            string pathNameSave = Server.MapPath($"~/temp/excel/{fileNameSave}");
            inputfile.SaveAs(pathNameSave);
            var finfo = new FileInfo(pathNameSave);
            var tmp = "";
            using (FileStream fs = finfo.OpenRead())
            {
                IWorkbook workbook = null;
                try
                {
                    try
                    {
                        if (fileExtension.ToLower() == ".xls") { workbook = new HSSFWorkbook(fs); }
                        else { workbook = new XSSFWorkbook(fs); }
                    }
                    catch (Exception ex) { throw new Exception($"Lỗi sai định dạng tập tin {fileName}: {ex.Message}"); }
                    var sheet = workbook.GetSheetAt(sheetIndex);
                    var tsqlv = new List<string>(); maxRow = sheet.LastRowNum;
                    var cs = true;
                    IRow row = null;
                    for (; indexRow <= maxRow; indexRow++)
                    {
                        row = sheet.GetRow(indexRow); if (row == null) { continue; }
                        /* Xác định vị trí hàng bắt đầu có dữ liệu */
                        foreach (var c in row.Cells)
                        {
                            tmp = c.GetValueAsString().Trim().ToLower();
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
                    /* Lấy danh sách cột */
                    var allColumns = AppHelper.dbSqliteWork.getColumns(bieu).Select(p => p.ColumnName).ToList();
                    allColumns.RemoveAt(0);
                    /* Thêm UserID */
                    listValue.Add(userID);
                    listValue.Add(timeUp);
                    tsql.Add($"INSERT INTO {bieu} ({string.Join(",", allColumns)}) VALUES ('{string.Join("','", listValue)}')");
                    /**
                     * Lấy dữ liệu chi tiết
                     */
                    allColumns = AppHelper.dbSqliteWork.getColumns(bieu + "chitiet").Select(p => p.ColumnName).ToList();
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
                        listValue = new List<string>() { "{@id2}", ma.sqliteGetValueField() };
                        for (jIndex = indexColumn + 1; jIndex < (indexColumn + fieldCount); jIndex++)
                        {
                            ICell c = row.GetCell(jIndex);
                            listValue.Add(c.GetValueAsString().Trim().sqliteGetValueField());
                        }
                        /* Cột lấy dữ liệu không đúng định dạng bỏ qua */
                        if (Regex.IsMatch(listValue[indexRegex], pattern) == false) { continue; }
                        /* Trường hợp trường số để trống thì cho bằng 0 */
                        foreach (int i in fieldNumbers) { if (Regex.IsMatch(listValue[i], "^[0-9]+$|^[0-9]+[.][0-9]+$") == false) { listValue[i] = "0"; } }
                        tsqlVaues.Add($"('{string.Join("','", listValue)}')");
                    }
                    if (tsqlVaues.Count > 0) { tsql.Add($"INSERT INTO {bieu}chitiet ({string.Join(",", allColumns)}) VALUES {string.Join(",", tsqlVaues)};"); }

                    System.IO.File.WriteAllText(Server.MapPath($"~/temp/excel/{fileNameSave}.sql"), string.Join(Environment.NewLine, tsql));
                    if (tsql.Count < 2) { throw new Exception("Không có dữ liệu chi tiết"); }

                    AppHelper.dbSqliteWork.Execute(tsql[0]);
                    tmp = $"{AppHelper.dbSqliteWork.getValue($"SELECT id FROM {bieu} WHERE userid = '{userID}' AND timeup = {timeUp} LIMIT 1")}";
                    if (Regex.IsMatch(tmp, "^[0-9]+$") == false) { throw new Exception($"Không cấp được id cho lần nhập liệu này: {tmp}"); }
                    for (int i = 1; i < tsql.Count; i++) { AppHelper.dbSqliteWork.Execute(tsql[i].Replace("{@id2}", tmp)); }
                    tsql.Add("/* {@id2}: " + tmp + " */");
                }
                catch (Exception ex2)
                {
                    ViewBag.Error = $"{bieu}: {fileName} (size {inputfile.ContentLength.getFileSize()}; Thời gian xử lý là: {(DateTime.Now - timeStart).TotalSeconds:0.##} giây <br />Lỗi trong quá trình đọc, nhập dữ liệu từ Excel '{fileName}': {ex2.getLineHTML()}";
                    return View();
                }
                finally { if (workbook != null) { workbook.Close(); workbook = null; } }
            }
            ViewBag.data = $"{bieu}: {fileName} (size {inputfile.ContentLength} b) được lưu tại {fileNameSave}; Thời gian xử lý là: {(DateTime.Now - timeStart).TotalSeconds:0.##} giây";
            return View();
        }
    }
}