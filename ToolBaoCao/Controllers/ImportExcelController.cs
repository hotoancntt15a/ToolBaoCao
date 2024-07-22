using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace ToolBaoCao.Controllers
{
    public class ImportExcelController : Controller
    {
        // GET: ImportExcel
        public ActionResult Index()
        {
            ViewBag.Title = "Quản lý nhập dữ liệu Excel";
            return View();
        }

        public ActionResult Update(string bieu, HttpPostedFileBase file)
        {
            ViewBag.data = "Đang thao tác";
            if (string.IsNullOrEmpty(bieu)) { ViewBag.Error = "Tham số biểu nhập không có chỉ định"; return View(); }
            if (file == null) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            if (file.ContentLength == 0) { ViewBag.Error = "Không có tập tin nào được đẩy lên"; return View(); }
            string fileName = Path.GetFileName(file.FileName);
            string fileExtension = Path.GetExtension(file.FileName);
            string fileNameSave = $"{bieu}{fileExtension}";
            int sheetIndex = 0; int packetSize = 1000;
            int indexRow = 0; int indexColumn = 0; int maxRow = 0;
            int fieldCount = 50; var tsql = new List<string>();
            using (var fs = file.InputStream)
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
                    IRow row = null;
                    for (; indexRow <= maxRow; indexRow++)
                    {
                        row = sheet.GetRow(indexRow); if (row == null) { continue; }
                        /* Xác định vị trí hàng bắt đầu có dữ liệu */
                        foreach (var c in row.Cells)
                        {
                            if ($"{c}".ToLower() == "ma_tinh") { indexColumn = c.ColumnIndex; break; }
                        }
                    }
                    if (indexRow >= maxRow) { throw new Exception("Không có dữ liệu"); }
                    var listValue = new List<string>();
                    string pattern = "^20[0-9][0-9]$";
                    int indexRegex = 3; int tmpInt = 0;
                    /*
                     * Bắt đầu đọc dữ liệu
                     */
                    /* - Đọc thông số biểu */
                    /* Biểu B04: ma_tinh ma_loai_kcb tu_thang den_thang nam loai_bv kieubv loaick hang_bv tuyen cs + userID */
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
                    row = sheet.GetRow(indexRow);
                    for (int j = indexColumn; j < j + 11; j++)
                    {
                        ICell c = row.GetCell(j);
                        if (c == null) { listValue.Add(""); }
                        else { listValue.Add($"{c.GetValueAsString()}".Trim()); }
                    }
                    /* Có phải là cơ sở không? */
                    tmpInt = (fieldCount - 1);
                    if (listValue[tmpInt] == "true" && listValue[tmpInt] != "1") { listValue[tmpInt] = "1"; } else { listValue[tmpInt] = "0"; }
                    /* Kiểm tra có đúng dữ liệu không */
                    if (Regex.IsMatch(listValue[indexRegex], pattern) == false) { throw new Exception($"dữ liệu không đúng cấu trúc (năm, thời gian): {listValue[indexRegex]}"); }
                    /* Thêm UserID */
                    listValue.Add("0");
                    tsql.Add($"INSERT INTO {bieu} VALUES ('{string.Join("','", listValue)}')");
                    /**
                     * Lấy dữ liệu chi tiết
                     */
                    switch (bieu)
                    {
                        /* Kiểm tra tổng số lượt KCB */
                        case "b02": fieldCount = 11; indexRegex = 3; pattern = "^[0-9]+$"; break;
                        /* Kiểm tra ngày TTBQ */
                        case "b04": fieldCount = 11; indexRegex = 9; pattern = "^[0-9]+[.,][0-9]+$|^[0-9]+$"; break;
                        /* Kiểm tra BQ chung trong kỳ */
                        case "b26": fieldCount = 34; indexRegex = 7; pattern = "^[0-9]+[.,][0-9]+$|^[0-9]+$"; break;
                        default: fieldCount = 11; break;
                    }
                    /* Bỏ qua dòng tiêu đề */
                    indexRow++;
                    var tsqlVaues = new List<string>();
                    for (; indexRow <= maxRow; indexRow++)
                    {
                        if (tsqlVaues.Count > packetSize)
                        {
                            tsql.Add($"INSERT INTO {bieu}chitiet VALUES ('{string.Join("','", listValue)}')");
                            tsqlVaues = new List<string>();
                        }
                        /* Dòng không tồn tại */
                        row = sheet.GetRow(indexRow); if (row == null) { continue; }
                        /* Số cột ít hơn số trường cần lấy dữ liệu */
                        if ((int)row.LastCellNum < fieldCount) { continue; }
                        /* Cột đầu tiên không phải là matinh dạng số */
                        string ma = row.GetCell(indexColumn).GetValueAsString();
                        if (Regex.IsMatch(ma, "^[0-9]+$") == false) { continue; }
                        /* Xây dựng tsql VALUES */
                        listValue = new List<string>() { ma.sqliteGetValueField() };
                        for (int j = indexColumn + 1; j < (indexColumn + fieldCount); j++)
                        {
                            ICell c = row.GetCell(j);
                            if (c == null) { listValue.Add(""); }
                            else { listValue.Add($"{c.GetValueAsString()}".Trim().sqliteGetValueField()); }
                            /* Cột lấy dữ liệu không đúng định dạng bỏ qua */
                            if (Regex.IsMatch(listValue[indexRegex], pattern) == false) { continue; }
                        }
                        tsqlVaues.Add($"('{string.Join("','", listValue)}')");
                    }
                    if (tsqlVaues.Count > 0) { tsql.Add($"INSERT INTO {bieu}chitiet VALUES ('{string.Join("','", listValue)}')"); }
                    System.IO.File.WriteAllText(Server.MapPath($"~/temp/excel/{fileNameSave}.tsql"), string.Join(Environment.NewLine, tsql));
                }
                catch (Exception ex2) { throw new Exception($"Lỗi trong quá trình đọc, nhập dữ liệu từ Excel {fileName}: {ex2.Message}"); }
                if (workbook != null)
                {
                    workbook.Close();
                    workbook = null;
                }
            }
            file.SaveAs(Server.MapPath($"~/temp/excel/{fileNameSave}"));
            ViewBag.data = $"{bieu}: {fileName} size {file.ContentLength} b được lưu tại {fileNameSave}";
            return View();
        }
    }
}