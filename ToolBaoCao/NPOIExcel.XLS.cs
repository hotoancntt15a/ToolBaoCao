using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace zModules.NPOIExcel
{
    public static class XLS
    {
        private static Regex regField = new Regex("{@[a-zA-Z0-9_]+}");
        public static List<System.Type> typeNumber = new List<System.Type>() { Type.GetType("System.Int16"), Type.GetType("System.Int32"), Type.GetType("System.Int64"), Type.GetType("System.Decimal"), Type.GetType("System.Double"), Type.GetType("System.Single") };
        public static List<System.Type> typeDateTime = new List<System.Type>() { Type.GetType("System.DateTime") };

        public static ICellStyle CreateCellStyleThin(this HSSFWorkbook hw)
        {
            ICellStyle cell = hw.CreateCellStyle();
            cell.BorderLeft = BorderStyle.Thin;
            cell.BorderRight = BorderStyle.Thin;
            cell.BorderTop = BorderStyle.Thin;
            cell.BorderBottom = BorderStyle.Thin;
            IFont font = hw.CreateFont();
            font.FontName = "Times New Roman";
            cell.SetFont(font);
            return cell;
        }

        public static ICellStyle CreateCellStyleTitle(this HSSFWorkbook hw)
        {
            ICellStyle cell = hw.CreateCellStyleThin();
            cell.WrapText = true;
            cell.Alignment = HorizontalAlignment.Center;
            cell.VerticalAlignment = VerticalAlignment.Center;
            cell.SetFont(hw.CreateFontTahomaBold());
            return cell;
        }

        public static IFont CreateFontTahomaBold(this HSSFWorkbook hw)
        {
            IFont fb = hw.CreateFont();
            fb.IsBold = true;
            fb.FontName = "Tahoma";
            return fb;
        }

        public static HSSFWorkbook exportXLS(this DataTable dt, string FileTemplate = "", string PathSave = "", List<CellMerge> lsTieuDe = null, int RowIndex = 0, bool ShowHeader = true, bool addColumnAutoNumber = false, string formatDate = "dd/MM/yyyy HH:mm:ss")
        {
            /* Tạo mới */
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            if (FileTemplate != "")
            {
                var fs = new FileStream(FileTemplate, FileMode.Open, FileAccess.Read);
                hssfworkbook = new HSSFWorkbook(fs);
                fs.Close();
                fs.Dispose();
            }
            /* Mặc định từ dòng thứ 2 đi */
            int index = RowIndex <= 0 ? 1 : RowIndex;
            int pointIndex = index - 1;
            /* Tạo hoặc lấy Sheet */
            var sheet = FileTemplate == "" ? hssfworkbook.CreateSheet() : hssfworkbook.GetSheetAt(0);
            /* Tạo đường viền của ô */
            var cell = hssfworkbook.CreateCellStyleThin();
            /* tạo tiêu đề */
            var cellb = hssfworkbook.CreateCellStyleTitle();
            var fb = hssfworkbook.CreateFontTahomaBold();
            if (ShowHeader)
            {
                var cr = sheet.CreateRow(index - 1);
                int i = 0;
                if (addColumnAutoNumber)
                {
                    ICell stt = cr.CreateCell(0, CellType.String);
                    stt.SetCellValue("STT");
                    stt.CellStyle = cellb;
                    foreach (DataColumn c in dt.Columns) { var cc = cr.CreateCell(i + 1, CellType.String); cc.SetCellValue(c.ColumnName); cc.CellStyle = cellb; cc.CellStyle.SetFont(fb); i++; }
                }
                else { foreach (DataColumn c in dt.Columns) { var cc = cr.CreateCell(i, CellType.String); cc.SetCellValue(c.ColumnName); cc.CellStyle = cellb; i++; } }
            }
            /* Kiểm tra và đặt tiêu đề */
            if (lsTieuDe == null) { lsTieuDe = new List<CellMerge>(); }
            foreach (var item in lsTieuDe)
            {
                /* Lấy vị trí RowIndex dòng tiêu đề */
                IRow row = sheet.GetRow(item.RowIndex);
                if (row == null) row = sheet.CreateRow(item.RowIndex);
                /* Lấy vị trí ColumnIndex và đặt giá trị */
                var cTitle = row.GetCell(item.ColumnIndex);
                if (cTitle == null)
                {
                    var column = row.CreateCell(item.ColumnIndex);
                    column.SetCellValue(item.Value);
                    column.CellStyle = cellb;
                }
                else
                {
                    cTitle.SetCellValue(item.Value);
                    cTitle.CellStyle = cellb;
                }
                if (item.MergeColumnCount > 0 || item.MergeRowCount > 0)
                {
                    int LastRow = item.MergeRowCount > item.RowIndex ? item.MergeRowCount : item.RowIndex;
                    int LastColumn = item.MergeColumnCount > item.ColumnIndex ? item.MergeColumnCount : item.ColumnIndex;
                    var cellRange = new CellRangeAddress(item.RowIndex, LastRow, item.ColumnIndex, LastColumn);
                    sheet.AddMergedRegion(cellRange);
                }
            }
            /* Xuất dữ liệu */
            if (addColumnAutoNumber)
            {
                foreach (DataRow r in dt.Rows)
                {
                    var cr = sheet.CreateRow(index);
                    ICell stt = cr.CreateCell(0, CellType.String);
                    stt.SetCellValue((index - pointIndex).ToString());
                    stt.CellStyle = cell;
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (typeNumber.Contains(dt.Columns[i].DataType))
                        {
                            var cc = cr.CreateCell(i + 1, CellType.Numeric);
                            if (r[i] == DBNull.Value) cc.SetCellValue(0);
                            else cc.SetCellValue(double.Parse(r[i].ToString()));
                            cc.CellStyle = cell;
                        }
                        else
                        {
                            var cc = cr.CreateCell(i + 1, CellType.String);
                            if (typeDateTime.Contains(dt.Columns[i].DataType))
                            {
                                if (r[i] == DBNull.Value) { cc.SetCellValue(""); }
                                else { cc.SetCellValue(((DateTime)r[i]).ToString(formatDate)); }
                            }
                            else { cc.SetCellValue(r[i].ToString()); }
                            cc.CellStyle = cell;
                        }
                    }
                    index++;
                }
            }
            else { sheet.fillData(dt, index, cell, formatDate, ShowHeader); }
            return hssfworkbook;
        }
        public static void fillData(this HSSFWorkbook hw, DataTable dt, int sheetIndex, string formatDate = "dd/MM/yyyy HH:mm:ss", bool showHeader = false, bool autoNumber = false, Dictionary<string, string> headerText = null)
        {
            var sheet = hw.GetSheetAt(sheetIndex);
            sheet.fillData(dt, 0, hw.CreateCellStyleThin(), formatDate, showHeader, autoNumber, headerText);
        }

        private static void fillData(this ISheet sheet, DataTable dt, int index, ICellStyle cellStyle, string formatDate = "dd/MM/yyyy HH:mm:ss", bool showHeader = false, bool autoNumber = false, Dictionary<string, string> headerText = null)
        {
            if (index == 0)
            {
                /* Find Row FillData */
                for (index = 0; index < sheet.LastRowNum; index++)
                {
                    var row = sheet.GetRow(index);
                    if (row == null) { continue; }
                    for (int jc = 0; jc < row.LastCellNum; jc++)
                    {
                        try
                        {
                            var c = row.Cells[jc];
                            if (c == null) { continue; }
                            if (c.CellType != CellType.String) { continue; }
                            if (c.StringCellValue.Trim() == "<filldata>") { break; }
                        }
                        catch { continue; }
                    }
                }
            }
            if (showHeader)
            {
                var cr = sheet.CreateRow(index);
                if (headerText == null)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        var cc = cr.CreateCell(i, CellType.String);
                        cc.SetCellValue(dt.Columns[i].ColumnName);
                        cc.CellStyle = cellStyle;
                    }
                }
                else
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        var cc = cr.CreateCell(i, CellType.String);
                        if (headerText.ContainsKey(dt.Columns[i].ColumnName))
                        {
                            cc.SetCellValue(headerText[dt.Columns[i].ColumnName]);
                        }
                        else
                        {
                            cc.SetCellValue(dt.Columns[i].ColumnName);
                        }
                        cc.CellStyle = cellStyle;
                    }
                }
                index++;
            }
            if (autoNumber)
            {
                int tt = 1;
                foreach (DataRow r in dt.Rows)
                {
                    var cr = sheet.CreateRow(index);
                    var cc1 = cr.CreateCell(0, CellType.Numeric);
                    cc1.SetCellValue(tt);
                    cc1.CellStyle = cellStyle;
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (typeNumber.Contains(dt.Columns[i].DataType))
                        {
                            var cc = cr.CreateCell(i + 1, CellType.Numeric);
                            if (r[i] == DBNull.Value) cc.SetCellValue(0);
                            else cc.SetCellValue(double.Parse(r[i].ToString()));
                            cc.CellStyle = cellStyle;
                        }
                        else
                        {
                            var cc = cr.CreateCell(i + 1, CellType.String);
                            if (typeDateTime.Contains(dt.Columns[i].DataType))
                            {
                                if (r[i] == DBNull.Value) { cc.SetCellValue(""); }
                                else { cc.SetCellValue(((DateTime)r[i]).ToString(formatDate)); }
                            }
                            else { cc.SetCellValue(r[i].ToString()); }
                            cc.CellStyle = cellStyle;
                        }
                    }
                    index++;
                    tt++;
                }
            }
            else
            {
                foreach (DataRow r in dt.Rows)
                {
                    var cr = sheet.CreateRow(index);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (typeNumber.Contains(dt.Columns[i].DataType))
                        {
                            var cc = cr.CreateCell(i, CellType.Numeric);
                            if (r[i] == DBNull.Value) cc.SetCellValue(0);
                            else cc.SetCellValue(double.Parse(r[i].ToString()));
                            cc.CellStyle = cellStyle;
                        }
                        else
                        {
                            var cc = cr.CreateCell(i, CellType.String);
                            if (typeDateTime.Contains(dt.Columns[i].DataType))
                            {
                                if (r[i] == DBNull.Value) { cc.SetCellValue(""); }
                                else { cc.SetCellValue(((DateTime)r[i]).ToString(formatDate)); }
                            }
                            else { cc.SetCellValue(r[i].ToString()); }
                            cc.CellStyle = cellStyle;
                        }
                    }
                    index++;
                }
            }
        }

        public static void saveXLS(this DataTable dt, string FileTemplate = "", string PathSave = "", List<CellMerge> lsTieuDe = null, int RowIndex = 0, bool ShowHeader = true, bool addColumnAutoNumber = true)
        {
            var xls = dt.exportXLS(FileTemplate, PathSave, lsTieuDe, RowIndex, ShowHeader, addColumnAutoNumber);
            if (PathSave == "")
            {
                PathSave = Path.GetPathRoot(FileTemplate);
                PathSave = PathSave.EndsWith("\\") ? PathSave + "template.xls" : PathSave + "\\template.xls";
            }
            if (File.Exists(PathSave)) File.Delete(PathSave);
            using (var fs = new FileStream(PathSave, FileMode.Create, FileAccess.Write)) { xls.Write(fs); }
            xls.Clear();
        }

        public static void saveExcel(this DataTable dt, string PathSave, string FileTemplate = "", List<CellMerge> lsTieuDe = null, int RowIndex = 0, bool ShowHeader = true, bool addColumnAutoNumber = false)
        {
            if (Path.GetExtension(PathSave).ToLower() == ".xls")
            {
                dt.saveXLS(FileTemplate, PathSave, lsTieuDe, RowIndex, ShowHeader, addColumnAutoNumber);
                return;
            }
            dt.saveXLSX(FileTemplate, PathSave, lsTieuDe, RowIndex, ShowHeader, addColumnAutoNumber);
        }

        public static MemoryStream WriteToStream(this HSSFWorkbook hssfworkbook)
        {
            MemoryStream file = new MemoryStream();
            hssfworkbook.Write(file);
            return file;
        }
        public static List<string> getSheetNames(this HSSFWorkbook hw) { var rs = new List<string>(); for (int i = 0; i < hw.NumberOfSheets; i++) { rs.Add(hw.GetSheetName(i)); } return rs; }
        public static List<string> getSheetNames(string fileXls)
        {
            if(string.IsNullOrEmpty(fileXls)) { return new List<string>(); }
            fileXls = fileXls.Trim();
            if(Regex.IsMatch(fileXls, ".xlsx$", RegexOptions.IgnoreCase)) { return XLSX.getSheetNames(fileXls); }
            var fs = new FileStream(fileXls, FileMode.Open, FileAccess.Read);
            HSSFWorkbook hw = new HSSFWorkbook(fs);
            fs.Close(); fs.Dispose();
            var items = hw.getSheetNames();
            hw.Close();
            return items;
        }
        public static void copySheetFrom(this HSSFWorkbook hw, HSSFWorkbook hw2, int sheetIndex = 0, string newNameSheet = "")
        {
            if (sheetIndex < 0) { sheetIndex = 0; }
            if (string.IsNullOrEmpty(newNameSheet)) {
                /* Lấy danh sách tên sheet của HW */
                var names = hw.getSheetNames();
                newNameSheet = hw2.GetSheetName(sheetIndex);
                if (names.Contains(newNameSheet)) { newNameSheet += $"{DateTime.Now:yyMMddHHmmss}"; }
            }
            ((HSSFSheet)hw2.GetSheetAt(sheetIndex)).CopyTo(hw, newNameSheet, true, true);
        }

        /// <summary>
        /// Tên cột dtField chính là tên trường Excel có định dạng {@[a-zA-Z0-9]+}
        /// </summary>
        /// <param name="dt">Dữ liệu cần xuất</param>
        /// <param name="fileTemplate">Tập tin mẫu</param>
        /// <param name="pathSave">Tập tin cần lưu</param>
        /// <param name="dtField">Danh sách dữ liệu riêng biệt</param>
        /// <param name="formatDate">Định dạng ngày tháng</param>
        public static void saveFromXls(this DataTable dt, string fileTemplate, string pathSave, Dictionary<string, string> dtField = null, string formatDate = "dd/MM/yyyy HH:mm", bool showHeader = false, bool autoNumber = false, Dictionary<string, string> headerText = null, int sheetIndex = 0)
        {
            if (Path.GetExtension(fileTemplate).ToLower() == ".xlsx")
            {
                dt.saveFromXlsx(fileTemplate, pathSave, dtField, formatDate, showHeader, autoNumber, headerText, sheetIndex);
                return;
            }
            var xls = dt.fromXls(fileTemplate, dtField, formatDate, showHeader, autoNumber, headerText, sheetIndex);
            if (pathSave == "")
            {
                pathSave = Path.GetPathRoot(fileTemplate);
                pathSave = pathSave.EndsWith("\\") ? pathSave + "template.xls" : pathSave + "\\template.xls";
            }
            if (File.Exists(pathSave)) { File.Delete(pathSave); }
            using (var fs = new FileStream(pathSave, FileMode.Create, FileAccess.Write)) { xls.Write(fs); }
            xls.Clear();
        }

        private static void setField(this ISheet sheet, Dictionary<string, string> dtField, ref int rowIndex)
        {
            /* Tạo hoặc lấy Sheet */
            if (dtField == null) { dtField = new Dictionary<string, string>(); }
            for (int index = 0; index < sheet.LastRowNum; index++)
            {
                var row = sheet.GetRow(index);
                if (row == null) { continue; }
                for (int jc = 0; jc < row.LastCellNum; jc++)
                {
                    try
                    {
                        var c = row.Cells[jc];
                        if (c == null) { continue; }
                        if (c.CellType != CellType.String) { continue; }
                        var s = c.StringCellValue;
                        if (s.Trim() == "<filldata>") { rowIndex = int.Parse(index.ToString()); }
                        if (dtField.Count > 0)
                        {
                            var reg = regField.Matches(s);
                            if (reg.Count == 0) { continue; }
                            var fis = new List<string>();
                            for (int jf = 0; jf < reg.Count; jf++) { fis.Add(reg[jf].Value.Replace("{@", "").Replace("}", "")); }
                            fis = fis.Distinct().ToList();
                            foreach (var v in fis)
                            {
                                if (dtField.Keys.Contains(v) == false)
                                {
                                    s = s.Replace("{@" + v + "}", "");
                                    continue;
                                }
                                s = s.Replace("{@" + v + "}", dtField[v]);
                            }
                            c.SetCellValue(s);
                        }
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// fromTemplate Tên cột dtField chính là tên trường Excel có định dạng {@[a-zA-Z0-9]+}
        /// </summary>
        /// <param name="dt">Dữ liệu xuất</param>
        /// <param name="fileTemplate">Đường dẫn tập tin mẫu</param>
        /// <param name="dtField">Danh sách dữ liệu biến </param>
        /// <param name="formatDate">Định dạng ngày</param>
        /// <param name="showHeader">Hiển thị tiêu đề cột của dt</param>
        /// <param name="autoNumber">Tạo thêm cột STT</param>
        /// <param name="headerText">Danh sách tiêu đề thay thế</param>
        /// <param name="sheetIndex">Vị trí Sheet</param>
        /// <returns></returns>
        public static HSSFWorkbook fromXls(this DataTable dt, string fileTemplate, Dictionary<string, string> dtField = null, string formatDate = "dd/MM/yyyy HH:mm", bool showHeader = false, bool autoNumber = false, Dictionary<string, string> headerText = null, int sheetIndex = 0)
        {
            /* Tạo mới */
            HSSFWorkbook hw;
            ISheet sheet;
            if (string.IsNullOrEmpty(fileTemplate))
            {
                hw = new HSSFWorkbook();
                sheet = hw.CreateSheet();
                showHeader = true;
                var r = sheet.CreateRow(0);
                var c = r.CreateCell(0, CellType.String);
                c.SetCellValue("<filldata>");
            }
            else
            {
                var fs = new FileStream(fileTemplate, FileMode.Open, FileAccess.Read);
                hw = new HSSFWorkbook(fs);
                fs.Close();
                fs.Dispose();
                sheet = hw.GetSheetAt(sheetIndex);
            }
            /* Mặc định từ dòng thứ 2 đi */
            int index = 0;
            sheet.setField(dtField, ref index);
            /* Xuất dữ liệu */
            sheet.fillData(dt, index, hw.CreateCellStyleThin(), formatDate, showHeader, autoNumber, headerText: headerText);
            return hw;
        }

        /// <summary>
        /// fromXls
        /// </summary>
        /// <param name="dt">Dữ liệu xuất</param>
        /// <param name="dtField">Danh sách dữ liệu biến </param>
        /// <param name="formatDate">Định dạng ngày</param>
        /// <param name="showHeader">Hiển thị tiêu đề cột của dt</param>
        /// <param name="autoNumber">Tạo thêm cột STT</param>
        /// <param name="headerText">Danh sách tiêu đề thay thế</param>
        /// <param name="sheetIndex">Vị trí Sheet</param>
        /// <returns></returns>
        public static void fromXls(this HSSFWorkbook hw, DataTable dt, Dictionary<string, string> dtField = null, string formatDate = "dd/MM/yyyy HH:mm", bool showHeader = false, bool autoNumber = false, Dictionary<string, string> headerText = null, int sheetIndex = 0)
        {
            /* Mặc định từ dòng thứ 2 đi */
            int index = 0;
            /* Tạo hoặc lấy Sheet */
            var sheet = hw.GetSheetAt(sheetIndex);
            sheet.setField(dtField, ref index);
            /* Xuất dữ liệu */
            sheet.fillData(dt, index, hw.CreateCellStyleThin(), formatDate, showHeader, autoNumber, headerText: headerText);
        }

        public static DataTable getDataFromExcel97_2003(this FileInfo file, int fieldCount = 50, int sheetIndex = 0, int maxRow = 0)
        {
            if (file.Extension.ToLower().EndsWith("xlsx")) { return file.getDataFromExcel(fieldCount, sheetIndex); }
            var fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);
            HSSFWorkbook hw = new HSSFWorkbook(fs);
            fs.Close();
            fs.Dispose();
            int maxColumn = 0;
            var sheet = hw.GetSheetAt(sheetIndex);
            var dt = new DataTable();
            for (int i = 0; i < fieldCount; i++) { dt.Columns.Add($"f{i}"); }
            if(maxRow == 0) { maxRow = sheet.LastRowNum; }
            else if(maxRow < 1) { maxRow = maxRow > sheet.LastRowNum ? sheet.LastRowNum : maxRow; }
            for (int i = 0; i < maxRow; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null) { continue; }
                if(maxColumn < row.LastCellNum) { maxColumn = row.LastCellNum; }
                DataRow r = dt.NewRow();
                for (int j = 0; j < (row.LastCellNum >= fieldCount ? fieldCount : row.LastCellNum); j++)
                {
                    var cell = row.GetCell(j);
                    if (cell == null) { continue; }
                    r[j] = $"{cell}";
                }
                dt.Rows.Add(r);
            }
            if(dt.Columns.Count > maxColumn) { for(var i = dt.Columns.Count -1; i>= maxColumn; i--) { dt.Columns.RemoveAt(i); } }
            hw.Close();
            return dt;
        }

        public static ISheet getSheetFromExcel97_2003(this FileInfo file, int sheetIndex = 0)
        {
            if (file.Extension.ToLower().EndsWith("xlsx")) { return file.getSheetFromExcel(sheetIndex); }
            var fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);
            HSSFWorkbook hw = new HSSFWorkbook(fs);
            fs.Close();
            fs.Dispose();
            var sheet = hw.GetSheetAt(sheetIndex);
            hw.Close();
            return sheet;
        }
    }
}