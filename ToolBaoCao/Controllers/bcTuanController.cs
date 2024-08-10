using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using zModules.NPOIExcel;

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
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bctuan", $"{Session["idtinh"]}_{Session["iduser"]}".GetMd5Hash());
            var d = new System.IO.DirectoryInfo(folderTemp);
            if (d.Exists == false) { d.Create(); }
            return View();
        }

        public ActionResult Buoc2()
        {
            var timeStart = DateTime.Now;
            var idUser = $"{Session["iduser"]}";
            var matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            if (Request.Files.Count == 0) { ViewBag.Error = "Không có tập tin dữ liệu nào đẩy lên"; return View(); }
            var id = $"{timeStart:yyyyMMddHHmmss}_{matinh}_{timeStart.Millisecond:000}";
            var timeUp = timeStart.toTimestamp().ToString();
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bctuan", $"{matinh}_{Session["iduser"]}".GetMd5Hash());
            ViewBag.id = id;
            try
            {
                /* Xoá hết các File có trong thư mục */
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
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
                if (list.Count > 0) { throw new Exception(string.Join("<br />", list)); }
                /* Tạo Phục Lục 1 */
                dbTemp.Execute($@"INSERT INTO pl01 (id_bc, idtinh, ma_tinh, ten_tinh, ma_vung, tyle_noitru, ngay_dtri_bq, chi_bq_chung, chi_bq_ngoai, chi_bq_noi, userid) SELECT id_bc, '{matinh}' AS idtinh, ma_tinh, ten_tinh, ma_vung, tyle_noitru, ngay_dtri_bq, chi_bq_chung, chi_bq_ngoai, chi_bq_noi, '{idUser}' AS userid
                    FROM b02chitiet WHERE id_bc='{id}' AND ma_tinh <> '';");
                /* Tạo Phục Lục 2*/
                dbTemp.Execute($@"INSERT INTO pl02 (id_bc, idtinh, ma_tinh, ten_tinh, ma_vung, chi_bq_xn, chi_bq_cdha, chi_bq_thuoc, chi_bq_pttt, chi_bq_vtyt, chi_bq_giuong, ngay_ttbq, userid) SELECT id_bc, '{matinh}' AS idtinh, ma_tinh, ten_tinh, ma_vung, bq_xn AS chi_bq_xn, bq_cdha AS chi_bq_cdha, bq_thuoc AS chi_bq_thuoc, bq_ptt AS chi_bq_pttt, bq_vtyt AS chi_bq_vtyt, bq_giuong AS chi_bq_giuong, ngay_ttbq, '{idUser}' AS userid
                    FROM b04chitiet WHERE id_bc='{id}' AND ma_tinh <> '';");
                /* Tạo Phục Lục 3 */
                dbTemp.Execute($@"INSERT INTO pl03 (id_bc, idtinh, ma_cskcb, ten_cskcb, tyle_noitru, ngay_dtri_bq, chi_bq_chung, chi_bq_ngoai, chi_bq_noi, userid) SELECT id_bc, '{matinh}' AS idtinh, ma_cskcb, ten_cskcb, tyle_noitru, ngay_dtri_bq, chi_bq_chung, chi_bq_ngoai, chi_bq_noi, '{idUser}' AS userid
                        FROM b02chitiet WHERE id_bc='{id}' AND ma_cskcb <> '';");
                /* Đọc dữ liệu DuToanGiao dự theo thoigian của b26_00 */
                var namDuToan = $"{dbTemp.getValue($"SELECT thoigian FROM b26 WHERE id_bc = '{id}' AND ma_tinh <> '' LIMIT 1")}";
                namDuToan = namDuToan.Substring(0, 4);
                var data = AppHelper.dbSqliteWork.getDataTable($"SELECT so_kyhieu_qd, tong_dutoan, namqd FROM dutoangia WHERE namqd <= {namDuToan} AND idtinh='{matinh}' ORDER BY namqd DESC LIMIT 1;");
                if (data.Rows.Count > 0)
                {
                    var tmp = $"{data.Rows[0]["namqd"]}";
                    ViewBag.x2 = $"{data.Rows[0]["so_kyhieu_qd"]}";
                    ViewBag.x3 = $"{data.Rows[0]["tong_dutoan"]}";
                }
                dbTemp.Close();
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
                    NPOI.SS.UserModel.ICell c = row.GetCell(jIndex);
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
                        NPOI.SS.UserModel.ICell c = row.GetCell(jIndex);
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
                /* Lưu lại file */
                using (FileStream stream = new FileStream(Path.Combine(folderTemp, $"id{idBaoCao}_{bieu}_{matinhImport}{fileExtension}"), FileMode.Create, FileAccess.Write)) { workbook.Write(stream); }
            }
            catch (Exception ex2) { messageError = $"Lỗi trong quá trình đọc, nhập dữ liệu từ Excel '{inputFile.FileName}': {ex2.getLineHTML()} <br />{tmp}"; }
            finally
            {
                /* Xoá luôn dữ liệu tạm của IIS */
                if (workbook != null)
                {
                    workbook.Close(); workbook = null;
                }
            }
            if (messageError != "") { throw new Exception(messageError); }
            return $"{bieu}_{matinhImport}";
        }

        public ActionResult Buoc3()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            var idBaoCao = Request.getValue("idobject");
            ViewBag.id = idBaoCao;
            var iduser = $"{Session["iduser"]}"; var idtinh = $"{Session["idtinh"]}";
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bctuan", $"{idtinh}_{iduser}".GetMd5Hash());
            var dirTemp = new System.IO.DirectoryInfo(folderTemp); var list = new List<string>();
            foreach (var f in dirTemp.GetFiles()) { list.Add($"{f.Name} ({f.Length.getFileSize()})"); }
            ViewBag.files = list;
            try
            {
                var tmp = "";
                var pathDB = Path.Combine(folderTemp, "import.db");
                if (System.IO.File.Exists(pathDB) == false) { throw new Exception($"Dữ liệu tạo báo cáo có ID '{idBaoCao}' đã bị huỷ hoặc không tồn tại trên hệ thống"); }
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                /* Tạo bctuan */
                var bctuan = createBcTuan(dbTemp, idBaoCao, idtinh, iduser, Request.getValue("x2"), Request.getValue("x3"), Request.getValue("x67"), Request.getValue("x68"), Request.getValue("x69"), Request.getValue("x70"));
                /* Đường dẫn lưu */
                string folderSave = Path.Combine(AppHelper.pathApp, "App_Data", "bctuan");
                /* Tạo docx */
                string pathFileTemplate = Path.Combine(AppHelper.pathApp, "App_Data", "baocaotuan.docx");
                if (System.IO.File.Exists(pathFileTemplate) == false)
                {
                    ViewBag.Error = "Không tìm thấy tập tin mẫu báo cáo 'baocaotuan.docx' trong thư mục App_Data";
                    return View();
                }
                using (var fileStream = new FileStream(pathFileTemplate, FileMode.Open, FileAccess.Read))
                {
                    var document = new NPOI.XWPF.UserModel.XWPFDocument(fileStream);
                    foreach (var paragraph in document.Paragraphs)
                    {
                        foreach (var run in paragraph.Runs)
                        {
                            tmp = run.ToString();
                            // Sử dụng Regex để tìm tất cả các match
                            MatchCollection matches = Regex.Matches(tmp, "{x[0-9]+}", RegexOptions.IgnoreCase);
                            foreach (Match match in matches) { tmp = tmp.Replace(match.Value, bctuan.getValue(match.Value, "", true)); }
                            run.SetText(tmp, 0);
                        }
                    }
                    tmp = Path.Combine(folderSave, $"bctuan_{idBaoCao}.docx");
                    if (System.IO.File.Exists(tmp)) { System.IO.File.Delete(tmp); }
                    using (FileStream stream = new FileStream(tmp, FileMode.Create, FileAccess.Write)) { document.Write(stream); }
                    /*
                     * MemoryStream memoryStream = new MemoryStream();
                            document.Write(memoryStream);
                            memoryStream.Position = 0;
                            return File(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{data.Rows[0]["ma_tinh"]}_{thoigian}.docx");
                    */
                }
                string idBaoCaoVauleField = idBaoCao.sqliteGetValueField();
                /* Tạo phụ lục báo cáo */
                var pl1 = dbTemp.getDataTable($"SELECT * FROM pl01 WHERE id_bc='{idBaoCaoVauleField}'");
                pl1.TableName = "pl01";
                if (pl1.Rows.Count == 0) { ViewBag.Error = $"Báo cáo có ID '{idBaoCao}' không tồn tại hoặc bị xoá trong hệ thống"; return View(); }
                var pl2 = dbTemp.getDataTable($"SELECT * FROM pl02 WHERE id_bc='{idBaoCaoVauleField}'");
                pl2.TableName = "pl02";
                var pl3 = dbTemp.getDataTable($"SELECT * FROM pl03 WHERE id_bc='{idBaoCaoVauleField}'");
                pl3.TableName = "pl03";
                
                var xlsx = XLSX.exportExcel(pl1, pl2, pl3);
                tmp = Path.Combine(folderSave, $"bctuan_pl_{idBaoCao}.xlsx");
                if (System.IO.File.Exists(tmp)) { System.IO.File.Delete(tmp); }
                using (FileStream stream = new FileStream(tmp, FileMode.Create, FileAccess.Write)) { xlsx.Write(stream); }
                xlsx.Close(); xlsx.Clear();
                /*
                 * XSSFWorkbook xlsx = XLSX.exportExcel(pl1, pl2, pl3);
                        var output = xlsx.WriteToStream();
                        return File(output.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"PL{tmp}.xlsx"); */
                /* Di chuyển tập tin Excel */
                foreach (var f in dirTemp.GetFiles("*.xls*")) { f.MoveTo(Path.Combine(folderSave, f.Name)); }

                /** Chuyển sang dữ liệu chính */
                var dbBCTuan = BuildDatabase.getDataBaoCaoTuan(idtinh);
                /* Bỏ cột ID (Số tự động) */
                /* Phụ Lục chuyển */
                pl1.Columns.RemoveAt(0); dbBCTuan.Insert("pl01", pl1);
                pl2.Columns.RemoveAt(0); dbBCTuan.Insert("pl02", pl2);
                pl3.Columns.RemoveAt(0); dbBCTuan.Insert("pl03", pl3);

                /* Báo cáo tuần chuyển */
                dbBCTuan.Update("bctuandocx", bctuan);
                dbBCTuan.Close();

                var dbImport = BuildDatabase.getDataImportBaoCaoTuan(idtinh);
                /* Di chuyển dữ liệu import */
                var data = dbTemp.getDataTable($"SELECT * FROM b02 WHERE id_bc='{idBaoCaoVauleField}';");
                data.Columns.RemoveAt(0); dbImport.Insert("b02", data);
                data = dbTemp.getDataTable($"SELECT * FROM b04 WHERE id_bc='{idBaoCaoVauleField}';");
                data.Columns.RemoveAt(0); dbImport.Insert("b04", data);
                data = dbTemp.getDataTable($"SELECT * FROM b26 WHERE id_bc='{idBaoCaoVauleField}';");
                data.Columns.RemoveAt(0); dbImport.Insert("b26", data);
                /* Dữ liệu chi tiết */
                data = dbTemp.getDataTable($"SELECT * FROM b02chitiet WHERE id_bc='{idBaoCaoVauleField}';");
                data.Columns.RemoveAt(0); dbImport.Insert("b02chitiet", data);
                data = dbTemp.getDataTable($"SELECT * FROM b04chitiet WHERE id_bc='{idBaoCaoVauleField}';");
                data.Columns.RemoveAt(0); dbImport.Insert("b04chitiet", data);
                data = dbTemp.getDataTable($"SELECT * FROM b26chitiet WHERE id_bc='{idBaoCaoVauleField}';");
                data.Columns.RemoveAt(0); dbImport.Insert("b26chitiet", data);
                dbTemp.Close();
            }
            catch (Exception ex) { ViewBag.Error = ex.getErrorSave(); }
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

        private Dictionary<string, string> createBcTuan(dbSQLite dbConnect, string idBaoCao, string maTinh, string idUser, string x2 = "", string x3 = "", string x67 = "", string x68 = "", string x69 = "", string x70 = "")
        {
            var bctuan = new Dictionary<string, string>() { { "id", idBaoCao } };
            if (Regex.IsMatch(x3, @"^\d+(\.\d+)?$") == false) { x3 = "0"; }

            double so1 = 0; double so2 = 0;
            var tmpD = new Dictionary<string, string>();
            string tsql = string.Empty;
            string tmp = string.Empty;

            /* Bỏ qua các vùng */
            var idBaoCaoValueField = idBaoCao.sqliteGetValueField();
            var maTinhValueField = maTinh.sqliteGetValueField();
            /* Bỏ qua các vùng */
            tsql = $"SELECT * FROM b02chitiet WHERE id_bc='{idBaoCaoValueField}' AND (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%')";
            var b02TQ = dbConnect.getDataTable(tsql).AsEnumerable().ToList();
            if (b02TQ.Count() == 0) { throw new Exception("B02 Toàn Quốc không có dữ liệu phù hợp truy vấn"); }
            /* Bỏ qua các vùng */
            tsql = $"SELECT * FROM b04chitiet WHERE id_bc='{idBaoCaoValueField}' AND  (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%')";
            var b04TQ = dbConnect.getDataTable(tsql).AsEnumerable().ToList();
            if (b04TQ.Count() == 0) { throw new Exception("B04 Toàn quốc không có dữ liệu phù hợp truy vấn"); }
            /* Bỏ qua các vùng */
            tsql = $"SELECT * FROM b26chitiet WHERE id_bc='{idBaoCaoValueField}' AND  (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%')";
            var b26TQ = dbConnect.getDataTable(tsql).AsEnumerable().ToList();
            if (b26TQ.Count() == 0) { throw new Exception("B26 Toàn quốc không có dữ liệu phù hợp truy vấn"); }

            var dataTinhB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == maTinh).FirstOrDefault();
            if (dataTinhB02 == null) { throw new Exception("B02 không có dữ liệu tỉnh phù hợp truy vấn"); }
            var dataTinhB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == maTinh).FirstOrDefault();
            if (dataTinhB26 == null) { throw new Exception("B26 không có dữ liệu tỉnh phù hợp truy vấn"); }

            var dataTQB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault();
            if (dataTQB02 == null) { throw new Exception("B02 không có dữ liệu toàn quốc phù hợp truy vấn"); }
            var dataTQB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault();
            if (dataTQB26 == null) { throw new Exception("B26 không có dữ liệu toàn quốc phù hợp truy vấn"); }

            /* Bỏ Toàn quốc ra khỏi danh sách */
            b02TQ = b02TQ.Where(p => p.Field<string>("ma_tinh") != "00").ToList();
            b26TQ = b26TQ.Where(p => p.Field<string>("ma_tinh") != "00").ToList();

            string mavung = dataTinhB02["ma_vung"].ToString();
            var data = dbConnect.getDataTable($"SELECT thoigian, timeup FROM b26 WHERE id_bc='{idBaoCao}'");
            string timeCreate = $"{data.Rows[0]["timeup"]}";
            tmp = $"{data.Rows[0]["thoigian"]}";
            var ngayTime = new DateTime(int.Parse(tmp.Substring(0, 4)), int.Parse(tmp.Substring(4, 2)), int.Parse(tmp.Substring(6)));

            /* X1 = {cột R (T-BHTT) bảng B02_TOANQUOC } */
            bctuan.Add("{X1}", dataTinhB02["t_bhtt"].ToString());
            /* X2 = {“ Quyết định số: Nếu không tìm thấy dòng nào của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán thì “TW chưa giao dự toán, tạm lấy theo dự toán năm trước”, nếu thấy lấy số ký hiệu các dòng QĐ của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán} */
            bctuan.Add("{X2}", x2);
            /* X3 = {Như trên, ko thấy thì lấy tổng tiền các dòng dự toán năm trước, thấy thì lấy tổng số tiền các dòng quyết định năm nay} */
            bctuan.Add("{X3}", x3);
            /* X4={X1/X3 %} So sánh với dự toán, tỉnh đã sử dụng */
            so2 = double.Parse(x3);
            if (so2 == 0) { bctuan.Add("{X4}", "0"); }
            else { bctuan.Add("{X4}", (double.Parse(bctuan["{X1}"]) / so2).ToString()); }

            /* X5 = {Cột tyle_noitru, dòng MA_TINH=10} bảng B02_TOANQUOC */
            bctuan.Add("{X5}", dataTinhB02["tyle_noitru"].ToString());
            /* X6 = {Cột tyle_noitru, dòng MA_TINH=00} bảng B02_TOANQUOC */
            bctuan.Add("{X6}", dataTQB02["tyle_noitru"].ToString());
            /* X7 = {đoạn văn tùy thuộc X5> hay < X6. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bctuan.Add("{X7}", "bằng");
            so1 = (double)dataTinhB02["tyle_noitru"];
            so2 = (double)dataTQB02["tyle_noitru"];
            if (so1 > so2) { bctuan["{X7}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bctuan["{X7}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X8={Sort cột G (TYLE_NOITRU) cao xuống thấp và lấy thứ tự}; */
            var sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
            int position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bctuan.Add("X8", position.ToString());
            /* X9 ={tính toán: total cột F (TONG_LUOT_NOI) chia cho Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG=mã vùng của tỉnh báo cáo}; */
            bctuan.Add("{X9}", "0");
            so2 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung).Sum(row => row.Field<long>("tong_luot"));
            if (so2 != 0)
            {
                so1 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung).Sum(row => row.Field<long>("tong_luot_noi"));
                bctuan["{X9}"] = (so1 / so2).ToString();
            }
            /* X10 ={đoạn văn tùy thuộc X5> hay < X9. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bctuan.Add("{X10}", "bằng");
            so1 = (double)dataTinhB02["tyle_noitru"];
            so2 = double.Parse(bctuan["{X9}"]); bctuan["{X9}"] = bctuan["{X9}"].ToString();
            if (so1 > so2) { bctuan["{X10}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bctuan["{X10}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X11= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort cột G (TYLE_NOITRU ) cao –thấp và lấy thứ tự} */
            sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                .OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bctuan.Add("{X11}", position.ToString());

            /* X12 = Ngày điều trị bình quân X12={Cột H NGAY_DTRI_BQ , dòng MA_TINH=10}; */
            bctuan.Add("{X12}", dataTinhB02["ngay_dtri_bq"].ToString());
            /* X13 = Nbình quân toàn quốc X13={cột H NGAY_DTRI_BQ, dòng MA_TINH=00}; */
            bctuan.Add("{X13}", dataTQB02["ngay_dtri_bq"].ToString());
            /* X14 = Số chênh lệch X14={đoạn văn tùy thuộc X12> hay < X13. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bctuan.Add("{X14}", "bằng");
            so1 = (double)dataTinhB02["ngay_dtri_bq"];
            so2 = (double)dataTQB02["ngay_dtri_bq"];
            if (so1 > so2) { bctuan["{X14}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bctuan["{X14}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X15 = xếp thứ so toàn quốc X15={Sort cột H (NGAY_DTRI_BQ) cao xuống thấp và lấy thứ tự}; */
            sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bctuan.Add("{X15}", position.ToString());
            /* X16 = Bình quân vùng X16 ={tính toán: A-Tổng ngày điều trị nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của cá tỉnh cùng mã vùng. A=Total(cột H (NGAY_DTRI_BQ) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
            bctuan.Add("{X16}", "0");
            so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<long>("tong_luot_noi"));
            if (so2 != 0)
            {
                so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => (r.Field<double>("ngay_dtri_bq") * r.Field<long>("tong_luot_noi")));
                bctuan["{X16}"] = (so1 / so2).ToString();
            }
            /* X17 = Số chênh lệch X17 ={đoạn văn tùy thuộc X12> hay < X16. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bctuan.Add("{X17}", "bằng");
            so1 = (double)dataTinhB02["ngay_dtri_bq"];
            so2 = double.Parse(bctuan["{X16}"]); bctuan["{X16}"] = bctuan["{X16}"].ToString();
            if (so1 > so2) { bctuan["{X17}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bctuan["{X17}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X18 = đứng thứ so với vùng X18 = {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột H (NGAY_DTRI_BQ) cao –thấp và lấy thứ tự} */
            sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                .OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bctuan.Add("{X18}", position.ToString());

            /* X19 = Chi bình quân chung X19={Cột I (CHI_BQ_CHUNG), dòng MA_TINH=10}; */
            tmpD = buildBCTuanB02(19, "chi_bq_chung", "tong_luot", "chi_bq_chung", mavung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X26 = Chi bình quân ngoại trú X26={Cột J (CHI_BQ_NGOAI), dòng MA_TINH=10}; */
            tmpD = buildBCTuanB02(26, "chi_bq_ngoai", "tong_luot_ngoai", "chi_bq_chung", mavung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X33 = Chi bình quân nội trú X33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
            tmpD = buildBCTuanB02(33, "chi_bq_noi", "tong_luot_noi", "chi_bq_chung", mavung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }

            /* ----- Dữ liệu X40 trở lên lọc dữ liệu tù B26 ------- */
            /* X40 = Bình quân xét nghiệm X40= {cột P (bq_xn) dòng có mã tỉnh = 10}; B26 */
            tmpD = buildBCTuanB26(40, "bq_xn", "bq_xn_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X43 Bình quân CĐHA X43= {cột R(bq_cdha) dòng có mã tỉnh =10}; */
            tmpD = buildBCTuanB26(43, "bq_cdha", "bq_cdha_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X46 Bình quân thuốc X46= {cột T(bq_thuoc) dòng có mã tỉnh =10}; */
            tmpD = buildBCTuanB26(46, "bq_thuoc", "bq_thuoc_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X49 Bình quân chi phẫu thuật X49= {cột V(bq_pt) dòng có mã tỉnh =10}; */
            tmpD = buildBCTuanB26(49, "bq_pt", "bq_pt_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X52 Bình quân chi thủ thuật X52= {cột X(bq_tt) dòng có mã tỉnh =10}; */
            tmpD = buildBCTuanB26(52, "bq_tt", "bq_tt_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X55 Bình quân chi vật tư y tế X55= {cột Z(bq_vtyt) dòng có mã tỉnh =10}; */
            tmpD = buildBCTuanB26(55, "bq_vtyt", "bq_vtyt_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X58 Bình quân chi tiền giường X58= {cột AB(bq_giuong) dòng có mã tỉnh =10}; */
            tmpD = buildBCTuanB26(58, "bq_giuong", "bq_giuong_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }

            /* X61 Chỉ định xét nghiệm X61={cột AD, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
            tmpD = buildBCTuan02B26(61, "chi_dinh_xn", "chi_dinh_xn_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }
            /* X64 =  Chỉ định CĐHA X64={cột AF, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
            tmpD = buildBCTuan02B26(64, "chi_dinh_cdha", "chi_dinh_cdha_tang", dataTinhB26);
            foreach (var d in tmpD) { bctuan.Add(d.Key, d.Value); }

            /* X67 Công tác kiểm soát chi X67={lần đầu lập BC sẽ rỗng, người dùng tự trình bày văn bản, lưu lại ở bảng dữ liệu kết quả báo cáo, kỳ sau sẽ tự động lấy từ kỳ trước, để người dùng kế thừa, sửa và lưu dùng cho kỳ này và kỳ sau} */
            bctuan.Add("{X67}", x67);
            /* X68 Công tác thanh, quyết toán năm X68={tương tự X67} */
            bctuan.Add("{X68}", x68);
            /* X69 Phương hướng kỳ tiếp theo X69={tương tự X67} */
            bctuan.Add("{X69}", x69);
            /* X70 Khó khăn, vướng mắc, đề xuất (nếu có) X70={tương tự X67} */
            bctuan.Add("{X70}", x70);

            /* X71 = {cột S T_BHTT_NOI bảng B02_TOANQUOC } */
            bctuan.Add("{X71}", dataTinhB02["t_bhtt_noi"].ToString());
            /* X72 = {cột T T_BHTT_NGOAI bảng B02_TOANQUOC } */
            bctuan.Add("{X72}", dataTinhB02["t_bhtt_ngoai"].ToString());
            /* X73 Lấy tên tỉnh */
            tmp = $"{AppHelper.dbSqliteMain.getValue($"SELECT ten FROM dmTinh WHERE id='{maTinh.sqliteGetValueField()}'")}";
            bctuan.Add("{X73}", tmp);
            /* X74 Lấy ngày chọn báo cáo */
            bctuan.Add("{X74}", ngayTime.ToString("dd/MM/yyyy"));

            bctuan.Add("ma_tinh", maTinh);
            bctuan.Add("userid", idUser);
            bctuan.Add("ngay", ngayTime.toTimestamp().ToString());
            bctuan.Add("timecreate", timeCreate);
            /* Tự động cập nhật vào dữ tuyết giao */
            var item = new Dictionary<string, string>() {
                    { "namqd", $"{ngayTime.Year}" },
                    { "idtinh", maTinh },
                    { "idhuyen", "" },
                    { "so_kyhieu_qd", x2},
                    { "tong_dutoan", x3 },
                    { "iduser", idUser }
                };
            AppHelper.dbSqliteWork.Update("dutoangiao", item, "replace");
            return bctuan;
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