﻿using NPOI.SS.UserModel;
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
    public class bcThangController : Controller
    {
        public ActionResult Index()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            var folder = Path.Combine(AppHelper.pathAppData, "bcThang");
            if (Directory.Exists(folder) == false) { Directory.CreateDirectory(folder); }
            folder = Path.Combine(AppHelper.pathTemp, "bcThang");
            if (Directory.Exists(folder) == false) { Directory.CreateDirectory(folder); }
            return View();
        }

        public ActionResult Buoc1()
        {
            if ($"{Session["idtinh"]}" == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            /* Tạo thư mục tạm */
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bcThang", $"{Session["idtinh"]}_{Session["iduser"]}".GetMd5Hash());
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
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bcThang", $"{matinh}_{Session["iduser"]}".GetMd5Hash());
            ViewBag.id = id;
            try
            {
                /* Xoá hết các File có trong thư mục */
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
                /* Khai báo dữ liệu tạm */
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                dbTemp.CreateImportBcTuan();
                dbTemp.CreatePhucLucBcTuan();
                dbTemp.CreateBcTuan();
                /* Đọc và kiểm tra các tập tin */
                var list = new List<string>();
                var bieus = new List<string>();
                for (int i = 0; i < Request.Files.Count; i++)
                {
                    if (Path.GetExtension(Request.Files[i].FileName).ToLower() != ".xlsx") { continue; }
                    list.Add($"{Request.Files[i].FileName} ({Request.Files[i].ContentLength.getFileSize()})");
                    bieus.Add(readExcelbcThang(dbTemp, Request.Files[i], Session, id, folderTemp, timeStart));
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
                dbTemp.Execute($@"INSERT INTO pl01 (id_bc, idtinh, ma_tinh, ten_tinh, ma_vung, tyle_noitru, ngay_dtri_bq, chi_bq_chung, chi_bq_ngoai, chi_bq_noi, userid) SELECT id_bc, '{matinh}' AS idtinh, ma_tinh, ten_tinh, ma_vung
                    , ROUND(tyle_noitru, 2) AS tyle_noitru
                    , ROUND(ngay_dtri_bq, 2) AS ngay_dtri_bq
                    , ROUND(chi_bq_chung) AS chi_bq_chung
                    , ROUND(chi_bq_ngoai) AS chi_bq_ngoai
                    , ROUND(chi_bq_noi) AS chi_bq_noi
                    , '{idUser}' AS userid
                    FROM b02chitiet WHERE id_bc='{id}' AND ma_tinh <> '' AND ma_tinh NOT LIKE 'V%';");
                /* Tạo Phục Lục 2*/
                dbTemp.Execute($@"INSERT INTO pl02 (id_bc, idtinh, ma_tinh, ten_tinh, ma_vung, chi_bq_xn, chi_bq_cdha, chi_bq_thuoc, chi_bq_pttt, chi_bq_vtyt, chi_bq_giuong, ngay_ttbq, userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    , ROUND(bq_xn) AS chi_bq_xn
                    , ROUND(bq_cdha) AS chi_bq_cdha
                    , ROUND(bq_thuoc) AS chi_bq_thuoc
                    , ROUND(bq_ptt) AS chi_bq_pttt
                    , ROUND(bq_vtyt) AS chi_bq_vtyt
                    , ROUND(bq_giuong) AS chi_bq_giuong
                    , ROUND(ngay_ttbq, 2) AS ngay_ttbq
                    , '{idUser}' AS userid
                    FROM b04chitiet WHERE id_bc='{id}' AND (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%');");
                /* Thêm cột vùng */
                var mavung = $"{dbTemp.getValue($"SELECT ma_vung FROM pl02 WHERE ma_tinh='{matinh}'")}";
                dbTemp.Execute($@"INSERT INTO pl02 (id_bc, idtinh, ma_tinh, ten_tinh, ma_vung , chi_bq_xn , chi_bq_cdha , chi_bq_thuoc , chi_bq_pttt , chi_bq_vtyt , chi_bq_giuong , ngay_ttbq , userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, '' AS ma_vung
                    , ROUND(bq_xn) AS chi_bq_xn
                    , ROUND(bq_cdha) AS chi_bq_cdha
                    , ROUND(bq_thuoc) AS chi_bq_thuoc
                    , ROUND(bq_ptt) AS chi_bq_pttt
                    , ROUND(bq_vtyt) AS chi_bq_vtyt
                    , ROUND(bq_giuong) AS chi_bq_giuong
                    , ROUND(ngay_ttbq, 2) AS ngay_ttbq
                    , '{idUser}' AS userid
                    FROM b04chitiet WHERE id_bc='{id}' AND ma_tinh LIKE 'V%' AND ma_vung='{mavung}';");
                /* Tạo Phục Lục 3 */
                var tablePL03 = dbTemp.getDataTable($@"SELECT id_bc, '{matinh}' AS idtinh, ma_cskcb, ten_cskcb, ma_vung
                    , ROUND(tyle_noitru, 2) AS tyle_noitru
                    , ROUND(ngay_dtri_bq, 2) AS ngay_dtri_bq
                    , ROUND(chi_bq_chung) AS chi_bq_chung
                    , ROUND(chi_bq_ngoai) AS chi_bq_ngoai
                    , ROUND(chi_bq_noi) AS chi_bq_noi
                    , '{idUser}' AS userid
                        FROM b02chitiet WHERE id_bc='{id}' AND ma_cskcb <> ''");
                /* Lấy danh sách Ma_CSKCB */
                var listIDCSKCB = string.Join(",", tablePL03.AsEnumerable().Select(x => x.Field<string>("ma_cskcb")).ToList()).Replace("'", "");
                var data = AppHelper.dbSqliteMain.getDataTable($"SELECT id, tuyencmkt, hangdv FROM dmcskcb WHERE ma_tinh ='{matinh}' AND id IN ('{listIDCSKCB.Replace(",", "','")}')");
                tablePL03.Columns.Add("tuyen_bv");
                tablePL03.Columns.Add("hang_bv");
                var dsCSKCB = data.AsEnumerable().Select(x => new
                {
                    id = x.Field<string>("id"),
                    tuyen = string.IsNullOrEmpty(x.Field<string>("tuyencmkt")) ? "*" : x.Field<string>("tuyencmkt"),
                    hang = string.IsNullOrEmpty(x.Field<string>("hangdv")) ? "*" : x.Field<string>("hangdv")
                }).ToList();
                foreach (DataRow row in tablePL03.Rows)
                {
                    var idCSKCB = $"{row["ma_cskcb"]}";
                    var v = dsCSKCB.FirstOrDefault(x => x.id == idCSKCB);
                    if (v == null) { row["tuyen_bv"] = "*"; row["hang_bv"] = "*"; }
                    else
                    {
                        row["tuyen_bv"] = v.tuyen;
                        row["hang_bv"] = v.hang.ToLower().StartsWith("h") ? v.hang : "*";
                    }
                }
                dbTemp.Insert("pl03", tablePL03);
                /* Đọc dữ liệu DuToanGiao dự theo thoigian của b26_00 */
                var namDuToan = $"{dbTemp.getValue($"SELECT thoigian FROM b26 WHERE id_bc = '{id}' AND ma_tinh <> '' LIMIT 1")}";
                namDuToan = namDuToan.Substring(0, 4);
                data = AppHelper.dbSqliteWork.getDataTable($"SELECT so_kyhieu_qd, tong_dutoan, namqd FROM dutoangiao WHERE namqd <= {namDuToan} AND idtinh='{matinh}' ORDER BY namqd DESC LIMIT 1;");
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

        private void createFilePhuLucbcThang(string idBaoCao, string matinh, dbSQLite dbBaoCaoTuan = null, Dictionary<string, string> bcThang = null)
        {
            if (dbBaoCaoTuan == null) { dbBaoCaoTuan = BuildDatabase.getDataBaoCaoTuan(matinh); }
            var idBaoCaoVauleField = idBaoCao.sqliteGetValueField();
            if (bcThang == null)
            {
                bcThang = new Dictionary<string, string>();
                var data = dbBaoCaoTuan.getDataTable($"SELECT * FROM bcThangdocx WHERE id='{idBaoCaoVauleField}';");
                if (data.Rows.Count > 0)
                {
                    foreach (DataColumn c in data.Columns)
                    {
                        bcThang.Add("{" + c.ColumnName.ToUpper() + "}", $"{data.Rows[0][c.ColumnName]}");
                    }
                }
            }
            /* Tạo phụ lục báo cáo */
            var pl = dbBaoCaoTuan.getDataTable($"SELECT * FROM pl01 WHERE id_bc='{idBaoCaoVauleField}';");
            var phuluc01 = createPhuLuc01(pl, matinh, bcThang);

            pl = dbBaoCaoTuan.getDataTable($"SELECT * FROM pl02 WHERE id_bc='{idBaoCaoVauleField}';");
            var phuluc02 = createPhuLuc02(pl, matinh);

            pl = dbBaoCaoTuan.getDataTable($"SELECT * FROM pl03 WHERE id_bc='{idBaoCaoVauleField}';");
            var phuluc03 = createPhuLuc03(pl, matinh, phuluc01);

            var xlsx = exportPhuLucbcThang(phuluc01, phuluc02, phuluc03);

            var tmp = Path.Combine(AppHelper.pathApp, "App_Data", "bcThang", $"tinh{matinh}", $"bcThang_pl_{idBaoCao}.xlsx");
            if (System.IO.File.Exists(tmp)) { System.IO.File.Delete(tmp); }
            using (FileStream stream = new FileStream(tmp, FileMode.Create, FileAccess.Write)) { xlsx.Write(stream); }
            xlsx.Close(); xlsx.Clear();
        }

        private XSSFWorkbook exportPhuLucbcThang(params DataTable[] par)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            int i = 0; int rowIndex = 0;
            var names = new List<string>();
            string tmp = "";
            foreach (DataTable dt in par)
            {
                var sheet = names.Contains(dt.TableName) ? workbook.CreateSheet() : workbook.CreateSheet(dt.TableName);
                names.Add(dt.TableName);
                var listColRight = new List<int>();
                var listColWith = new List<int>();
                switch (dt.TableName.ToLower())
                {
                    case "phuluc01":
                        listColRight = new List<int>() { 0, 2, 4, 6, 8, 10 };
                        listColWith = new List<int>() { 9, 18, 10, 18, 13, 18, 10, 18, 10, 18, 10 };
                        break;

                    case "phuluc02":
                        listColRight = new List<int>() { 0, 2, 3, 4, 5, 6, 7, 8 };
                        listColWith = new List<int>() { 9, 18, 13, 13, 13, 13, 13, 13, 13 };
                        break;

                    case "phuluc03":
                        listColRight = new List<int>() { 0, 2, 4, 6, 8, 10 };
                        listColWith = new List<int>() { 9, 33, 10, 33, 13, 33, 10, 33, 10, 33, 10 };
                        break;

                    default: break;
                }
                for (int colIndex = 0; colIndex < listColWith.Count; colIndex++) { sheet.SetColumnWidth(colIndex, (listColWith[colIndex] * 256)); }
                /* Tạo tiêu đề */
                rowIndex = 0;
                var row = sheet.CreateRow(rowIndex);
                i = -1;
                foreach (DataColumn col in dt.Columns)
                {
                    i++;
                    var cell = row.CreateCell(i, CellType.String);
                    cell.CellStyle = workbook.CreateCellStyleThin(true, true, true);
                    cell.SetCellValue(Regex.Replace(col.ColumnName, @"[ ][(]\d+[)]", ""));
                }
                /* Đổ dữ liệu */
                foreach (DataRow r in dt.Rows)
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    i = -1;
                    if ($"{r[0]}{r[1]}" == "")
                    {
                        foreach (DataColumn col in dt.Columns)
                        {
                            i++;
                            var cell = row.CreateCell(i, CellType.String);
                            cell.CellStyle = workbook.CreateCellStyleThin();
                            cell.SetCellValue("");
                        }
                    }
                    else
                    {
                        foreach (DataColumn col in dt.Columns)
                        {
                            i++;
                            var cell = row.CreateCell(i, CellType.String);
                            tmp = $"{r[i]}";
                            if (tmp.StartsWith("<b>"))
                            {
                                cell.CellStyle = workbook.CreateCellStyleThin(true);
                                cell.SetCellValue(tmp.Substring(3));
                            }
                            else
                            {
                                cell.CellStyle = workbook.CreateCellStyleThin();
                                cell.SetCellValue(tmp);
                            }
                            if (listColRight.Contains(i)) { cell.CellStyle.Alignment = HorizontalAlignment.Right; }
                        }
                    }
                }
            }
            return workbook;
        }

        private string readExcelbcThang(dbSQLite dbConnect, HttpPostedFileBase inputFile, HttpSessionStateBase Session, string idBaoCao, string folderTemp, DateTime timeStart)
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
                /* Không xác định được biểu thì bỏ qua */
                if (bieu == "") { workbook.Close(); return ""; }
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
                /* Yêu cầu tháng từ là từ đầu năm dương lịch */
                if (bieu == "b02")
                {
                    if (listValue[2] != "1") { throw new Exception($"Biểu {bieu} yêu cầu từ tháng 1; Tháng từ của biểu là '{listValue[2]}'"); }
                }
                if (bieu == "b04")
                {
                    if (listValue[1] != "1") { throw new Exception($"Biểu {bieu} yêu cầu từ tháng 1; Tháng từ của biểu là '{listValue[1]}'"); }
                }
                /* Có phải là cơ sở không? */
                tmpInt = (fieldCount - 1);
                listValue[tmpInt] = "1";
                if (listValue[0] == "00") { listValue[tmpInt] = "0"; cs = false; }
                tmp = string.Join(",", listValue);
                if (tmp.Contains(",,")) { throw new Exception($"Biểu {bieu} không đúng định dạng."); }
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
                    foreach (int i in fieldNumbers) { if (Regex.IsMatch(listValue[i], @"^-?\d+([.]\d+)?$") == false) { listValue[i] = "0"; } }
                    listValue.Add(idBaoCao);
                    tsqlVaues.Add($"('{string.Join("','", listValue)}')");
                }
                if (tsqlVaues.Count > 0) { tsql.Add($"INSERT INTO {bieu}chitiet ({string.Join(",", allColumns)}) VALUES {string.Join(",", tsqlVaues)};"); }
                tmp = string.Join(Environment.NewLine, tsql);
                /* System.IO.File.WriteAllText(Path.Combine(folderTemp, $"id{idBaoCao}_{bieu}_{matinhImport}.sql"), tmp); */
                dbConnect.Execute(tmp);
                if (tsql.Count < 2) { throw new Exception("Không có dữ liệu chi tiết"); }
                /* Lưu lại file */
                using (FileStream stream = new FileStream(Path.Combine(folderTemp, $"id{idBaoCao}_{bieu}_{matinhImport}{fileExtension}"), FileMode.Create, FileAccess.Write)) { workbook.Write(stream); }
            }
            catch (Exception ex2) { messageError = $"Lỗi trong quá trình đọc, nhập dữ liệu từ Excel '{inputFile.FileName}': {ex2.getLineHTML()}"; }
            finally
            {
                /* Xoá luôn dữ liệu tạm của IIS */
                if (workbook != null) { workbook.Close(); workbook = null; }
            }
            if (messageError != "") { throw new Exception(messageError); }
            return $"{bieu}_{matinhImport}";
        }

        public ActionResult Tai()
        {
            var id = Request.getValue("idobject");
            if (id.Contains("_") == false) { ViewBag.Error = $"Tham số không đúng '{id}'"; return View(); }
            var tmp = id.Split('_')[1];
            try
            {
                var d = new DirectoryInfo(Path.Combine(AppHelper.pathAppData, "bcThang", $"tinh{tmp}"));
                if (d.Exists == false) { throw new Exception($"Thư mục '{d.FullName}' không tồn tại"); }
                ViewBag.path = d.FullName;
                /* Trường hợp không tìm thấy tập tin nào thì tạo lại nếu còn dữ liệu */
                var tsql = "";
                var matinh = tmp;
                if (System.IO.File.Exists(Path.Combine(d.FullName, $"bcThang_{id}.docx")) == false || System.IO.File.Exists(Path.Combine(d.FullName, $"bcThang_pl_{id}.docx")) == false)
                {
                    /* Tạo lại báo cáo */
                    var dbBaoCao = BuildDatabase.getDataBaoCaoTuan(matinh);
                    tsql = $"SELECT * FROM bcThangdocx WHERE id='{id.sqliteGetValueField()}'";
                    var data = dbBaoCao.getDataTable(tsql);
                    dbBaoCao.Close();
                    if (data.Rows.Count == 0)
                    {
                        ViewBag.Error = $"Báo cáo tuần có ID '{id}' thuộc tỉnh có mã '{matinh}' không tồn tại hoặc đã bị xoá khỏi hệ thống";
                        return View();
                    }
                    var bcThang = new Dictionary<string, string>();
                    foreach (DataColumn c in data.Columns) { bcThang.Add("{" + c.ColumnName.ToUpper() + "}", $"{data.Rows[0][c.ColumnName]}"); }
                    createFilebcThangDocx(id, matinh, bcThang);
                    createFilePhuLucbcThang(id, matinh, dbBaoCao, bcThang);
                    dbBaoCao.Close();
                }
                tmp = Path.Combine(d.FullName, $"id{id}_b26_00.xlsx");
                if (System.IO.File.Exists(tmp) == false)
                {
                    /* Tạo lại biểu 26 Toàn quốc */
                    var dbImport = BuildDatabase.getDataImportBaoCaoTuan(matinh);
                    var data = dbImport.getDataTable($"SELECT * FROM b26chitiet WHERE id_bc='{id.sqliteGetValueField()}' AND ma_tinh <> ''");
                    dbImport.Close();
                    data.saveXLSX(PathSave: Path.Combine(d.FullName, $"id{id}_b26_00.xlsx"), addColumnAutoNumber: false);
                }
            }
            catch (Exception ex) { ViewBag.Error = ex.Message; }
            return View();
        }

        public ActionResult Buoc3()
        {
            var idtinh = $"{Session["idtinh"]}";
            if (idtinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            ViewBag.idtinh = idtinh;
            var idBaoCao = Request.getValue("idobject");
            ViewBag.id = idBaoCao;
            var iduser = $"{Session["iduser"]}";
            /* Đường dẫn lưu */
            var folderSave = Path.Combine(AppHelper.pathApp, "App_Data", "bcThang", $"tinh{idtinh}");
            if (Directory.Exists(folderSave) == false) { Directory.CreateDirectory(folderSave); }
            ViewBag.forlderSave = folderSave;
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bcThang", $"{idtinh}_{iduser}".GetMd5Hash());
            var dirTemp = new System.IO.DirectoryInfo(folderTemp);
            var list = new List<string>();
            foreach (var f in dirTemp.GetFiles()) { list.Add($"{f.Name} ({f.Length.getFileSize()})"); }
            ViewBag.files = list;
            try
            {
                var tmp = "";
                var pathDB = Path.Combine(folderTemp, "import.db");
                if (System.IO.File.Exists(pathDB) == false) { throw new Exception($"Dữ liệu tạo báo cáo có ID '{idBaoCao}' đã bị huỷ hoặc không tồn tại trên hệ thống"); }
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                /* Tạo bcThang */
                var bcThang = createbcThang(dbTemp, idBaoCao, idtinh, iduser, Request.getValue("x2"), Request.getValue("x3"), Request.getValue("x67"), Request.getValue("x68"), Request.getValue("x69"), Request.getValue("x70"));
                /* Tạo docx */
                createFilebcThangDocx(idBaoCao, idtinh, bcThang);
                /* Tạo dữ liệu để xuất phụ lục */
                string idBaoCaoVauleField = idBaoCao.sqliteGetValueField();
                var dbbcThang = BuildDatabase.getDataBaoCaoTuan(idtinh);
                var dbImport = BuildDatabase.getDataImportBaoCaoTuan(idtinh);
                /* Tạo phụ lục báo cáo */
                var pl = dbTemp.getDataTable($"SELECT * FROM pl01 WHERE id_bc='{idBaoCaoVauleField}'");
                if (pl.Rows.Count == 0) { ViewBag.Error = $"Báo cáo có ID '{idBaoCao}' không tồn tại hoặc bị xoá trong hệ thống"; return View(); }
                var phuluc01 = createPhuLuc01(pl, idtinh, bcThang);
                pl.Columns.RemoveAt(0); dbbcThang.Insert("pl01", pl);

                pl = dbTemp.getDataTable($"SELECT * FROM pl02 WHERE id_bc='{idBaoCaoVauleField}'");
                var phuluc02 = createPhuLuc02(pl, idtinh);
                pl.Columns.RemoveAt(0); dbbcThang.Insert("pl02", pl);

                pl = dbTemp.getDataTable($"SELECT * FROM pl03 WHERE id_bc='{idBaoCaoVauleField}'");
                var phuluc03 = createPhuLuc03(pl, idtinh, phuluc01);
                pl.Columns.RemoveAt(0); dbbcThang.Insert("pl03", pl);

                var xlsx = exportPhuLucbcThang(phuluc01, phuluc02, phuluc03);
                phuluc01 = null; phuluc02 = null; phuluc03 = null;

                tmp = Path.Combine(folderSave, $"bcThang_pl_{idBaoCao}.xlsx");
                if (System.IO.File.Exists(tmp)) { System.IO.File.Delete(tmp); }
                using (FileStream stream = new FileStream(tmp, FileMode.Create, FileAccess.Write)) { xlsx.Write(stream); }
                xlsx.Close(); xlsx.Clear();
                /*
                 * XSSFWorkbook xlsx = XLSX.exportExcel(pl1, pl2, pl3);
                        var output = xlsx.WriteToStream();
                        return File(output.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"PL{tmp}.xlsx"); */
                /* Di chuyển tập tin Excel */
                foreach (var f in dirTemp.GetFiles("*.xls*")) { f.MoveTo(Path.Combine(folderSave, f.Name)); }

                /* Báo cáo tuần chuyển */
                dbbcThang.Update("bcThangdocx", bcThang);
                dbbcThang.Close();

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
            catch (Exception ex)
            {
                ViewBag.Error = ex.getErrorSave();
                DeletebcThang(idtinh);
            }
            return View();
        }

        private DataTable createPhuLuc01(DataTable pl1, string idtinh, Dictionary<string, string> bcThang)
        {
            var tmp = "";
            var phuluc01 = new DataTable("PhuLuc01");
            phuluc01.Columns.Add("Mã Tỉnh");
            phuluc01.Columns.Add("Tên tỉnh");
            phuluc01.Columns.Add("Tỷ lệ nội trú (%)");
            phuluc01.Columns.Add("Tên tỉnh (2)");
            phuluc01.Columns.Add("Ngày điều trị BQ (ngày)");
            phuluc01.Columns.Add("Tên tỉnh (3)");
            phuluc01.Columns.Add("Chi BQ chung");
            phuluc01.Columns.Add("Tên tỉnh (4)");
            phuluc01.Columns.Add("chi BQ nội trú");
            phuluc01.Columns.Add("Tên tỉnh (5)");
            phuluc01.Columns.Add("Chi BQ ngoại trú");
            /* Lấy mã vùng */
            var mavung = $"{pl1.AsEnumerable().Where(x => x.Field<string>("ma_tinh") == idtinh).Select(x => x.Field<string>("ma_vung")).FirstOrDefault()}";
            /* Lọc danh sách bỏ */
            var view = pl1.AsEnumerable().Where(x => x.Field<string>("ma_tinh") != "00" && x.Field<string>("ma_vung") == mavung).OrderByDescending(x => x.Field<double>("tyle_noitru")).ToList();
            /* Sắp xếp theo tỷ lệ nội trú */
            foreach (DataRow r in view)
            {
                var dr = phuluc01.NewRow();
                tmp = $"{r["ma_tinh"]}".Trim();
                if (idtinh == tmp)
                {
                    dr[0] = $"<b>{r["ma_tinh"]}";
                    dr[1] = $"<b>{r["ten_tinh"]}";
                    dr[2] = $"<b>{r["tyle_noitru"]}";
                }
                else
                {
                    dr[0] = $"{r["ma_tinh"]}";
                    dr[1] = $"{r["ten_tinh"]}";
                    dr[2] = $"{r["tyle_noitru"]}";
                }
                for (int i = 3; i < phuluc01.Columns.Count; i++) { dr[i] = ""; }
                phuluc01.Rows.Add(dr);
            }
            /* Sắp xếp theo Ngày điều trị BQ (ngày) */
            view = view.OrderByDescending(x => x.Field<double>("ngay_dtri_bq")).ToList();
            for (int i = 0; i < phuluc01.Rows.Count; i++)
            {
                tmp = $"{view[i]["ma_tinh"]}".Trim();
                if (tmp == idtinh)
                {
                    phuluc01.Rows[i][3] = $"<b>{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][4] = $"<b>{view[i]["ngay_dtri_bq"]}";
                }
                else
                {
                    phuluc01.Rows[i][3] = $"{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][4] = $"{view[i]["ngay_dtri_bq"]}";
                }
            }
            /* Sắp xếp theo Chi_bq_chung (ngày) */
            view = view.OrderByDescending(x => x.Field<double>("chi_bq_chung")).ToList();
            for (int i = 0; i < phuluc01.Rows.Count; i++)
            {
                tmp = $"{view[i]["ma_tinh"]}".Trim();
                if (tmp == idtinh)
                {
                    phuluc01.Rows[i][5] = $"<b>{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][6] = $"<b>{view[i]["chi_bq_chung"]}";
                }
                else
                {
                    phuluc01.Rows[i][5] = $"{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][6] = $"{view[i]["chi_bq_chung"]}";
                }
            }
            /* Sắp xếp theo chi BQ nội trú */
            view = view.OrderByDescending(x => x.Field<double>("chi_bq_noi")).ToList();
            for (int i = 0; i < phuluc01.Rows.Count; i++)
            {
                tmp = $"{view[i]["ma_tinh"]}".Trim();
                if (tmp == idtinh)
                {
                    phuluc01.Rows[i][7] = $"<b>{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][8] = $"<b>{view[i]["chi_bq_noi"]}";
                }
                else
                {
                    phuluc01.Rows[i][7] = $"{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][8] = $"{view[i]["chi_bq_noi"]}";
                }
            }
            /* Sắp xếp theo Chi BQ ngoại trú */
            view = view.OrderByDescending(x => x.Field<double>("chi_bq_ngoai")).ToList();
            for (int i = 0; i < phuluc01.Rows.Count; i++)
            {
                tmp = $"{view[i]["ma_tinh"]}".Trim();
                if (tmp == idtinh)
                {
                    phuluc01.Rows[i][9] = $"<b>{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][10] = $"<b>{view[i]["chi_bq_ngoai"]}";
                }
                else
                {
                    phuluc01.Rows[i][9] = $"{view[i]["ten_tinh"]}";
                    phuluc01.Rows[i][10] = $"{view[i]["chi_bq_ngoai"]}";
                }
            }
            /* Dòng trống */
            phuluc01.Rows.Add("", "", "0", "", "0", "", "0", "", "0", "", "0");
            /* Toàn Quốc */
            view = pl1.AsEnumerable().Where(x => x.Field<string>("ma_tinh") == "00").ToList().GetRange(0, 1);
            if (view.Count == 0) { phuluc01.Rows.Add("00", "00", "0", "00", "0", "00", "0", "00", "0", "00", "0"); }
            else
            {
                phuluc01.Rows.Add($"00"
                    , $"{view[0]["ten_tinh"]}", $"{view[0]["tyle_noitru"]}"
                    , $"{view[0]["ten_tinh"]}", $"{view[0]["ngay_dtri_bq"]}"
                    , $"{view[0]["ten_tinh"]}", $"{view[0]["chi_bq_chung"]}"
                    , $"{view[0]["ten_tinh"]}", $"{view[0]["chi_bq_noi"]}"
                    , $"{view[0]["ten_tinh"]}", $"{view[0]["chi_bq_ngoai"]}");
            }
            var row00 = phuluc01.Rows[phuluc01.Rows.Count - 1];
            /* Xây dựng Vùng */
            phuluc01.Rows.Add($"V{(mavung.Length == 1 ? $"0{mavung}" : mavung)}",
                    "Vùng", $"{Math.Round(double.Parse(bcThang["{X9}"]), 2)}",
                    "Vùng", $"{Math.Round(double.Parse(bcThang["{X16}"]), 2)}",
                    "Vùng", bcThang["{X23}"],
                    "Vùng", bcThang["{X37}"], /* Ngoại trú */
                    "Vùng", bcThang["{X30}"]); /* Nội trú */
            var rowV = phuluc01.Rows[phuluc01.Rows.Count - 1];
            /* Chỉ lấy dòng Tỉnh đã chọn */
            view = pl1.AsEnumerable().Where(x => x.Field<string>("ma_tinh") == idtinh).ToList().GetRange(0, 1);
            if (view.Count == 0) { phuluc01.Rows.Add($"<b>{idtinh}", $"<b>{idtinh}", "<b>0", $"<b>{idtinh}", "<b>0", $"<b>{idtinh}", "<b>0", $"<b>{idtinh}", "<b>0", $"<b>{idtinh}", "<b>0"); }
            else
            {
                phuluc01.Rows.Add($"<b>{idtinh}"
                    , $"<b>{view[0]["ten_tinh"]}", $"<b>{view[0]["tyle_noitru"]}"
                    , $"<b>{view[0]["ten_tinh"]}", $"<b>{view[0]["ngay_dtri_bq"]}"
                    , $"<b>{view[0]["ten_tinh"]}", $"<b>{view[0]["chi_bq_chung"]}"
                    , $"<b>{view[0]["ten_tinh"]}", $"<b>{view[0]["chi_bq_noi"]}"
                    , $"<b>{view[0]["ten_tinh"]}", $"<b>{view[0]["chi_bq_ngoai"]}");
            }
            var index = phuluc01.Rows.Count - 1;
            var rowTinh = phuluc01.NewRow();
            for (int i = 0; i < rowTinh.Table.Columns.Count; i++) { rowTinh[i] = $"{phuluc01.Rows[index][i]}".Substring(3); }
            /* Chênh với toàn quốc */
            phuluc01.Rows.Add("", "Chênh so toàn quốc"
                , $"{Math.Round(double.Parse($"{rowTinh[2]}") - double.Parse($"{row00[2]}"), 2)}",
                "", $"{Math.Round(double.Parse($"{rowTinh[4]}") - double.Parse($"{row00[4]}"), 2)}",
                "", $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{row00[6]}"))}",
                "", $"{(double.Parse($"{rowTinh[8]}") - double.Parse($"{row00[8]}"))}",
                "", $"{(double.Parse($"{rowTinh[10]}") - double.Parse($"{row00[10]}"))}");

            /* Chênh với Vùng */
            index++;
            phuluc01.Rows.Add("", "Chênh so vùng",
                $"{Math.Round(double.Parse($"{rowTinh[2]}") - double.Parse($"{rowV[2]}"), 2)}",
                "", $"{Math.Round(double.Parse($"{rowTinh[4]}") - double.Parse($"{rowV[4]}"), 2)}",
                "", $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{rowV[6]}"))}",
                "", $"{(double.Parse($"{rowTinh[8]}") - double.Parse($"{rowV[8]}"))}",
                "", $"{(double.Parse($"{rowTinh[10]}") - double.Parse($"{rowV[10]}"))}");
            return phuluc01;
        }

        private DataTable createPhuLuc02(DataTable pl2, string idtinh)
        {
            /* Bỏ [ma tỉnh] - ở cột tên tỉnh */
            for (int i = 0; i < pl2.Rows.Count; i++) { pl2.Rows[i]["ten_tinh"] = Regex.Replace($"{pl2.Rows[i]["ten_tinh"]}", @"^V?\d+[ -]+", ""); }
            var phuluc02 = new DataTable("PhuLuc02");
            phuluc02.Columns.Add("Mã Tỉnh");
            phuluc02.Columns.Add("Tên tỉnh");
            phuluc02.Columns.Add("BQ_XN (đồng)");
            phuluc02.Columns.Add("BQ_CĐHA (đồng)");
            phuluc02.Columns.Add("BQ_THUOC (đồng)");
            phuluc02.Columns.Add("BQ_PTTT (đồng)");
            phuluc02.Columns.Add("BQ_VTYT (đồng)");
            phuluc02.Columns.Add("BQ_GIUONG (đồng)");
            phuluc02.Columns.Add("Ngày thanh toán BQ");
            /* Lấy dòng tỉnh */
            var view = pl2.AsEnumerable().Where(x => x.Field<string>("ma_tinh") == idtinh).ToList().GetRange(0, 1);
            var mavung = "";
            if (view.Count == 0) { phuluc02.Rows.Add($"<b>{idtinh}", $"<b>{idtinh}", "<b>0", "<b>0", "<b>0", "<b>0", "<b>0", "<b>0", "<b>0"); }
            else
            {
                mavung = $"{view[0]["ma_vung"]}";
                phuluc02.Rows.Add($"<b>{idtinh}", $"<b>{view[0]["ten_tinh"]}"
                    , $"<b>{view[0]["chi_bq_xn"]}"
                    , $"<b>{view[0]["chi_bq_cdha"]}"
                    , $"<b>{view[0]["chi_bq_thuoc"]}"
                    , $"<b>{view[0]["chi_bq_pttt"]}"
                    , $"<b>{view[0]["chi_bq_vtyt"]}"
                    , $"<b>{view[0]["chi_bq_giuong"]}"
                    , $"<b>{view[0]["ngay_ttbq"]}");
            }
            var index = phuluc02.Rows.Count - 1;
            DataRow rowTinh = phuluc02.NewRow();
            for (int i = 0; i < rowTinh.Table.Columns.Count; i++) { rowTinh[i] = $"{phuluc02.Rows[index][i]}".Substring(3); }
            view = pl2.AsEnumerable().Where(x => (x.Field<string>("ma_tinh") != idtinh && x.Field<string>("ma_tinh") != "00") && x.Field<string>("ma_vung") == mavung).ToList();
            foreach (DataRow r in view)
            {
                phuluc02.Rows.Add($"{r["ma_tinh"]}", $"{r["ten_tinh"]}"
                    , $"{r["chi_bq_xn"]}"
                    , $"{r["chi_bq_cdha"]}"
                    , $"{r["chi_bq_thuoc"]}"
                    , $"{r["chi_bq_pttt"]}"
                    , $"{r["chi_bq_vtyt"]}"
                    , $"{r["chi_bq_giuong"]}"
                    , $"{r["ngay_ttbq"]}");
            }
            /* Dòng trống */
            phuluc02.Rows.Add("", "", "0", "0", "0", "0", "0", "0", "0");
            /* Toàn quốc */
            view = pl2.AsEnumerable().Where(x => x.Field<string>("ma_tinh") == "00").ToList().GetRange(0, 1);
            if (view.Count == 0) { phuluc02.Rows.Add("00", "00", "0", "0", "0", "0", "0", "0", "0"); }
            else
            {
                phuluc02.Rows.Add("00", view[0]["ten_tinh"]
                    , $"{view[0]["chi_bq_xn"]}"
                    , $"{view[0]["chi_bq_cdha"]}"
                    , $"{view[0]["chi_bq_thuoc"]}"
                    , $"{view[0]["chi_bq_pttt"]}"
                    , $"{view[0]["chi_bq_vtyt"]}"
                    , $"{view[0]["chi_bq_giuong"]}"
                    , $"{view[0]["ngay_ttbq"]}");
            }
            DataRow row00 = phuluc02.Rows[phuluc02.Rows.Count - 1];
            /* Vùng */
            var vung = pl2.AsEnumerable()
                .Where(x => x.Field<string>("ma_vung") == "" && x.Field<string>("ma_tinh") != "00")
                .Select(x => new
                {
                    matinh = x.Field<string>("ma_tinh"),
                    chi_bq_xn = x.Field<double>("chi_bq_xn"),
                    chi_bq_cdha = x.Field<double>("chi_bq_cdha"),
                    chi_bq_thuoc = x.Field<double>("chi_bq_thuoc"),
                    chi_bq_pttt = x.Field<double>("chi_bq_pttt"),
                    chi_bq_vtyt = x.Field<double>("chi_bq_vtyt"),
                    chi_bq_giuong = x.Field<double>("chi_bq_giuong"),
                    ngay_ttbq = x.Field<double>("ngay_ttbq"),
                })
                .FirstOrDefault();
            if (vung == null) { phuluc02.Rows.Add($"V{mavung}", "Vùng", "0", "0", "0", "0", "0", "0", "0"); }
            else
            {
                phuluc02.Rows.Add(vung.matinh, "Vùng",
                    $"{vung.chi_bq_xn}",
                    $"{vung.chi_bq_cdha}",
                    $"{vung.chi_bq_thuoc}",
                    $"{vung.chi_bq_pttt}",
                    $"{vung.chi_bq_vtyt}",
                    $"{vung.chi_bq_giuong}",
                    $"{vung.ngay_ttbq}");
            }
            DataRow rowVung = phuluc02.Rows[phuluc02.Rows.Count - 1];
            /* Tỉnh */
            phuluc02.Rows.Add($"{rowTinh[0]}", $"{rowTinh[1]}", $"{rowTinh[2]}", $"{rowTinh[3]}", $"{rowTinh[4]}", $"{rowTinh[5]}", $"{rowTinh[6]}", $"{rowTinh[7]}", $"{rowTinh[8]}");
            /* Chênh so toàn quốc */
            phuluc02.Rows.Add("", "Chênh so toàn quốc",
                $"{double.Parse($"{rowTinh[2]}") - double.Parse($"{row00[2]}")}",
                $"{(double.Parse($"{rowTinh[3]}") - double.Parse($"{row00[3]}"))}",
                $"{(double.Parse($"{rowTinh[4]}") - double.Parse($"{row00[4]}"))}",
                $"{(double.Parse($"{rowTinh[5]}") - double.Parse($"{row00[5]}"))}",
                $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{row00[6]}"))}",
                $"{(double.Parse($"{rowTinh[7]}") - double.Parse($"{row00[7]}"))}",
                $"{Math.Round(double.Parse($"{rowTinh[8]}") - double.Parse($"{row00[8]}"), 2)}");

            /* Chênh với Vùng */
            index++;
            phuluc02.Rows.Add("", "Chênh so vùng",
                $"{(double.Parse($"{rowTinh[2]}") - double.Parse($"{rowVung[2]}"))}",
                $"{(double.Parse($"{rowTinh[3]}") - double.Parse($"{rowVung[3]}"))}",
                $"{(double.Parse($"{rowTinh[4]}") - double.Parse($"{rowVung[4]}"))}",
                $"{(double.Parse($"{rowTinh[5]}") - double.Parse($"{rowVung[5]}"))}",
                $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{rowVung[6]}"))}",
                $"{(double.Parse($"{rowTinh[7]}") - double.Parse($"{rowVung[7]}"))}",
                $"{Math.Round(double.Parse($"{rowTinh[8]}") - double.Parse($"{rowVung[8]}"), 2)}");
            return phuluc02;
        }

        private DataTable createPhuLuc03(DataTable pl3, string idtinh, DataTable phuLuc01)
        {
            var phuluc03 = new DataTable("PhuLuc03");
            phuluc03.Columns.Add("Mã");
            phuluc03.Columns.Add("hạng BV /Tên CSKCB ");
            phuluc03.Columns.Add("Tỷ lệ nội trú (%)");
            phuluc03.Columns.Add("hạng BV /Tên CSKCB (1)");
            phuluc03.Columns.Add("Ngày điều trị BQ (ngày)");
            phuluc03.Columns.Add("hạng BV /Tên CSKCB (2)");
            phuluc03.Columns.Add("Chi BQ chung");
            phuluc03.Columns.Add("hạng BV /Tên CSKCB (3)");
            phuluc03.Columns.Add("chi BQ nội trú");
            phuluc03.Columns.Add("hạng BV /Tên CSKCB (4)");
            phuluc03.Columns.Add("Chi BQ ngoại trú");
            if (phuLuc01.Rows.Count > 5)
            {
                for (int i = phuLuc01.Rows.Count - 5; i < phuLuc01.Rows.Count - 2; i++)
                {
                    var dr = phuluc03.NewRow();
                    for (int j = 0; j < phuLuc01.Columns.Count; j++) { dr[j] = phuLuc01.Rows[i][j]; }
                    phuluc03.Rows.Add(dr);
                }
            }
            phuluc03.Rows.Add("", "", "0", "", "0", "", "0", "", "0", "", "0");
            var indexHeader = phuluc03.Rows.Count;

            List<string> listTuyen = pl3.AsEnumerable().Select(x => x.Field<string>("tuyen_bv")).Distinct().ToList();
            /* Trường hợp bênh viện quân y */
            if (listTuyen.Contains("*"))
            {
                listTuyen.Remove("*");
                var rTuyen = pl3.AsEnumerable().Where(x => x.Field<string>("tuyen_bv") == "*").ToList();
                var rd = getPhuLuc03(rTuyen, "*", phuluc03);
                foreach (DataRow r in rd) { phuluc03.Rows.Add(r); }
            }
            /* Trường hợp các tuyến tỉnh */
            List<string> tmpl = listTuyen.Where(x => x.ToLower().StartsWith("t") == true).OrderBy(x => x).ToList();
            foreach (var tuyen in tmpl)
            {
                var rTuyen = pl3.AsEnumerable().Where(x => x.Field<string>("tuyen_bv") == tuyen).ToList();
                var rd = getPhuLuc03(rTuyen, tuyen, phuluc03);
                foreach (DataRow r in rd) { phuluc03.Rows.Add(r); }
            }
            /* Trường hợp các tuyến huyện, xã */
            listTuyen = listTuyen.Where(x => x.ToLower().StartsWith("t") == false).OrderBy(x => x).ToList();
            foreach (var tuyen in listTuyen)
            {
                var rTuyen = pl3.AsEnumerable().Where(x => x.Field<string>("tuyen_bv") == tuyen).ToList();
                var rd = getPhuLuc03(rTuyen, tuyen, phuluc03);
                foreach (DataRow r in rd) { phuluc03.Rows.Add(r); }
            }
            return phuluc03;
        }

        private List<DataRow> getPhuLuc03(List<DataRow> rTuyen, string tuyen, DataTable phuLuc03)
        {
            List<DataRow> rs = new List<DataRow>();
            var rNew = phuLuc03.NewRow();
            if (tuyen == "*") { rNew.ItemArray = new object[] { "", "Tuyến (*)", "0", "", "0", "", "0", "", "0", "", "0" }; }
            else { rNew.ItemArray = new object[] { "", $"Tuyến {tuyen}", "0", "", "0", "", "0", "", "0", "", "0" }; }
            rs.Add(rNew);
            var listHang = rTuyen.Select(x => x.Field<string>("hang_bv")).Distinct().OrderBy(x => x).ToList();
            foreach (var hang in listHang)
            {
                int indexHeader = rs.Count();
                var view = rTuyen.Where(x => x.Field<string>("hang_bv") == hang).OrderByDescending(x => x.Field<double>("tyle_noitru")).ToList();
                /* Sắp xếp theo tỷ lệ nội trú */
                foreach (DataRow r in view)
                {
                    var dr = phuLuc03.NewRow();
                    dr[0] = $"{r["ma_cskcb"]}";
                    dr[1] = $"{hang}/{r["ten_cskcb"]}";
                    dr[2] = $"{r["tyle_noitru"]}";
                    for (int i = 3; i < phuLuc03.Columns.Count; i++) { dr[i] = ""; }
                    rs.Add(dr);
                }
                /* Sắp xếp theo Ngày điều trị BQ (ngày) */
                view = rTuyen.Where(x => x.Field<string>("hang_bv") == hang).OrderByDescending(x => x.Field<double>("ngay_dtri_bq")).ToList();
                for (int i = indexHeader; i < rs.Count; i++)
                {
                    rs[i][3] = $"{hang}/{view[(i - indexHeader)]["ten_cskcb"]}";
                    rs[i][4] = $"{view[(i - indexHeader)]["ngay_dtri_bq"]}";
                }
                /* Sắp xếp theo chi_bq_chung */
                view = rTuyen.Where(x => x.Field<string>("hang_bv") == hang).OrderByDescending(x => x.Field<double>("chi_bq_chung")).ToList();
                for (int i = indexHeader; i < rs.Count; i++)
                {
                    rs[i][5] = $"{hang}/{view[(i - indexHeader)]["ten_cskcb"]}";
                    rs[i][6] = $"{view[(i - indexHeader)]["chi_bq_chung"]}";
                }
                /* Sắp xếp theo chi BQ nội trú */
                view = rTuyen.Where(x => x.Field<string>("hang_bv") == hang).OrderByDescending(x => x.Field<double>("chi_bq_noi")).ToList();
                for (int i = indexHeader; i < rs.Count; i++)
                {
                    rs[i][7] = $"{hang}/{view[(i - indexHeader)]["ten_cskcb"]}";
                    rs[i][8] = $"{view[(i - indexHeader)]["chi_bq_noi"]}";
                }
                /* Sắp xếp theo Chi BQ ngoại trú */
                view = rTuyen.Where(x => x.Field<string>("hang_bv") == hang).OrderByDescending(x => x.Field<double>("chi_bq_ngoai")).ToList();
                for (int i = indexHeader; i < rs.Count; i++)
                {
                    rs[i][9] = $"{hang}/{view[(i - indexHeader)]["ten_cskcb"]}";
                    rs[i][10] = $"{view[(i - indexHeader)]["chi_bq_ngoai"]}";
                }
            }
            return rs;
        }

        private string getPosition(string mavung, string matinh, string fieldSortDesc, List<DataRow> data)
        {
            if (mavung != "")
            {
                var sortedRows = data.Where(r => r.Field<string>("ma_vung") == mavung)
                   .OrderByDescending(row => row.Field<double>(fieldSortDesc)).ToList();
                return (sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1).ToString();
            }
            var s = data.OrderByDescending(row => row.Field<double>(fieldSortDesc)).ToList();
            return (s.FindIndex(row => row.Field<string>("ma_tinh") == matinh) + 1).ToString();
        }

        private Dictionary<string, string> buildbcThangB02(int iKey, string fieldChiBQ, string fieldChiBQChung, string fieldTongLuotVung, string fieldTongChiVung, string mavung, string matinh, DataRow rowTinh, DataRow rowTQ, List<DataRow> data)
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
            if (so1 > so2) { d[keys[2]] = $"cao hơn {Math.Round(so1 - so2, 0).FormatCultureVN()}"; }
            else { if (so1 < so2) { d[keys[2]] = $"thấp hơn {Math.Round(so2 - so1, 0).FormatCultureVN()}"; } }
            /* X36= xếp thứ so toàn quốc X36={Sort cột K CHI_BQ_NOI cao xuống thấp và lấy thứ tự}; */
            d.Add(keys[3], getPosition("", matinh, fieldChiBQ, data));
            /*** Vùng
             = SUM(tong_chi)/SUM(tong_luot)
             */
            /* X37 = Bình quân vùng X37={tính toán: A-Tổng chi nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của các tỉnh cùng mã vùng. A=Total  (cột K (CHI_BQ_NOI) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
            d.Add(keys[4], "0");
            so2 = data.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<long>(fieldTongLuotVung));
            if (so2 != 0)
            {
                so1 = data.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<double>(fieldTongChiVung));
                d[keys[4]] = (so1 / so2).ToString();
            }
            /* X38 = số chênh lệch X38 ={đoạn văn tùy thuộc X33 > hay < X37. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            d.Add(keys[5], "bằng");
            so1 = double.Parse(d[keys[0]]);
            so2 = double.Parse(d[keys[4]]);
            if (so1 > so2) { d[keys[5]] = $"cao hơn {Math.Round(so1 - so2, 0).FormatCultureVN()}"; }
            else { if (so1 < so2) { d[keys[5]] = $"thấp hơn {Math.Round(so2 - so1, 0).FormatCultureVN()}"; } }
            /* X39 đứng thứ so với vùng X39= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột K (CHI_BQ_NOI) cao –thấp và lấy thứ tự} */
            d.Add(keys[6], getPosition(mavung, matinh, fieldChiBQ, data));
            return d;
        }

        private Dictionary<string, string> buildbcThangB26(int iKey, string field1, string field2, DataRow row)
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
            if (x1 > 0) { d[key3] = "tăng " + Math.Round(Math.Abs(x - (x / (x1 + 100) * 100)), 0).FormatCultureVN() + " đồng"; }
            else { if (x1 < 0) { d[key3] = "giảm " + Math.Round(Math.Abs(x - (x / (x1 + 100) * 100)), 0).FormatCultureVN() + " đồng"; } }
            return d;
        }

        private Dictionary<string, string> buildbcThang02B26(int iKey, string field1, string field2, DataRow row)
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
            d.Add(key3, Math.Abs(so1 - (so1 / (so2 + 100) * 100)).FormatCultureVN() + " bệnh nhân");
            return d;
        }

        private Dictionary<string, string> createbcThang(dbSQLite dbConnect, string idBaoCao, string maTinh, string idUser, string x2 = "", string x3 = "", string x67 = "", string x68 = "", string x69 = "", string x70 = "")
        {
            var bcThang = new Dictionary<string, string>() { { "id", idBaoCao } };
            if (x3.isNumberUSInt() == false) { x3 = "0"; }

            double so1 = 0; double so2 = 0;
            var tmpD = new Dictionary<string, string>();
            string tsql = string.Empty;
            string tmp = string.Empty;

            /* Bỏ qua các vùng */
            var idBaoCaoValueField = idBaoCao.sqliteGetValueField();
            var maTinhValueField = maTinh.sqliteGetValueField();
            /* Bỏ qua các vùng */
            tsql = $@"SELECT ma_tinh
                ,ten_tinh
                ,ma_vung
                ,tong_luot
                ,tong_luot_ngoai
                ,tong_luot_noi
                ,ROUND(tyle_noitru, 2) AS tyle_noitru
                ,ROUND(ngay_dtri_bq, 2) AS ngay_dtri_bq
                ,ROUND(chi_bq_chung) AS chi_bq_chung
                ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai
                ,ROUND(chi_bq_noi) AS chi_bq_noi
                ,ROUND(tong_chi) AS tong_chi
                ,ROUND(ty_trong, 2) AS ty_trong
                ,ROUND(tong_chi_ngoai) AS tong_chi_ngoai
                ,ROUND(ty_trong_kham, 2) AS ty_trong_kham
                ,ROUND(tong_chi_noi) AS tong_chi_noi
                ,ROUND(ty_trong_giuong, 2) AS ty_trong_giuong
                ,ROUND(t_bhtt) AS t_bhtt
                ,ROUND(t_bhtt_noi) AS t_bhtt_noi
                ,ROUND(t_bhtt_ngoai) AS t_bhtt_ngoai
                FROM b02chitiet WHERE id_bc='{idBaoCaoValueField}' AND (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%')";
            var b02TQ = dbConnect.getDataTable(tsql).AsEnumerable().ToList();
            if (b02TQ.Count() == 0) { throw new Exception("B02 Toàn Quốc không có dữ liệu phù hợp truy vấn"); }
            /* Bỏ qua các vùng
             *
            tsql = $"SELECT * FROM b04chitiet WHERE id_bc='{idBaoCaoValueField}' AND  (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%')";
            var b04TQ = dbConnect.getDataTable(tsql).AsEnumerable().ToList();
            if (b04TQ.Count() == 0) { throw new Exception("B04 Toàn quốc không có dữ liệu phù hợp truy vấn"); }
             */
            /* Bỏ qua các vùng */
            tsql = $@"SELECT ma_tinh
                ,ten_tinh
                ,vitri_chibq
                ,vitri_tyle_noitru
                ,vitri_tlxn
                ,vitri_tlcdha
                ,ROUND(tytrong, 2) AS tytrong
                ,ROUND(chi_bq_chung) AS chi_bq_chung
                ,ROUND(chi_bq_chung_tang, 2) AS chi_bq_chung_tang
                ,ROUND(tyle_noitru, 2) AS tyle_noitru
                ,ROUND(tyle_noitru_tang, 2) AS tyle_noitru_tang
                ,ROUND(lan_kham_bq, 2) AS lan_kham_bq
                ,ROUND(lan_kham_bq_tang, 2) AS lan_kham_bq_tang
                ,ROUND(ngay_dtri_bq, 2) AS ngay_dtri_bq
                ,ROUND(ngay_dtri_bq_tang, 2) AS ngay_dtri_bq_tang
                ,ROUND(bq_xn) AS bq_xn
                ,ROUND(bq_xn_tang, 2) AS bq_xn_tang
                ,ROUND(bq_cdha) AS bq_cdha
                ,ROUND(bq_cdha_tang, 2) AS bq_cdha_tang
                ,ROUND(bq_thuoc) AS bq_thuoc
                ,ROUND(bq_thuoc_tang, 2) AS bq_thuoc_tang
                ,ROUND(bq_pt) AS bq_pt
                ,ROUND(bq_pt_tang, 2) AS bq_pt_tang
                ,ROUND(bq_tt) AS bq_tt
                ,ROUND(bq_tt_tang, 2) AS bq_tt_tang
                ,ROUND(bq_vtyt) AS bq_vtyt
                ,ROUND(bq_vtyt_tang, 2) AS bq_vtyt_tang
                ,ROUND(bq_giuong) AS bq_giuong
                ,ROUND(bq_giuong_tang, 2) AS bq_giuong_tang
                ,ROUND(chi_dinh_xn, 2) AS chi_dinh_xn
                ,ROUND(chi_dinh_xn_tang, 2) AS chi_dinh_xn_tang
                ,ROUND(chi_dinh_cdha, 2) AS chi_dinh_cdha
                ,ROUND(chi_dinh_cdha_tang, 2) AS chi_dinh_cdha_tang
                ,ma_vung
                FROM b26chitiet WHERE id_bc='{idBaoCaoValueField}' AND (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%')";
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
            bcThang.Add("{X1}", dataTinhB02["t_bhtt"].ToString());
            /* X2 = {“ Quyết định số: Nếu không tìm thấy dòng nào của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán thì “TW chưa giao dự toán, tạm lấy theo dự toán năm trước”, nếu thấy lấy số ký hiệu các dòng QĐ của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán} */
            bcThang.Add("{X2}", x2);
            /* X3 = {Như trên, ko thấy thì lấy tổng tiền các dòng dự toán năm trước, thấy thì lấy tổng số tiền các dòng quyết định năm nay} */
            bcThang.Add("{X3}", x3.lamTronTrieuDong());
            /* X4={X1/X3 %} So sánh với dự toán, tỉnh đã sử dụng */
            so2 = double.Parse(bcThang["{X3}"]);
            if (so2 == 0) { bcThang.Add("{X4}", "0"); }
            else { bcThang.Add("{X4}", ((double.Parse(bcThang["{X1}"]) / so2) * 100).ToString("0.##")); }

            /* X5 = {Cột tyle_noitru, dòng MA_TINH=10} bảng B02_TOANQUOC */
            bcThang.Add("{X5}", dataTinhB02["tyle_noitru"].ToString());
            /* X6 = {Cột tyle_noitru, dòng MA_TINH=00} bảng B02_TOANQUOC */
            bcThang.Add("{X6}", dataTQB02["tyle_noitru"].ToString());
            /* X7 = {đoạn văn tùy thuộc X5> hay < X6. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThang.Add("{X7}", "bằng");
            so1 = (double)dataTinhB02["tyle_noitru"];
            so2 = (double)dataTQB02["tyle_noitru"];
            if (so1 > so2) { bcThang["{X7}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThang["{X7}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X8={Sort cột G (TYLE_NOITRU) cao xuống thấp và lấy thứ tự}; */
            var sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
            int position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThang.Add("X8", position.ToString());
            /* X9 ={tính toán: total cột F (TONG_LUOT_NOI) chia cho Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG=mã vùng của tỉnh báo cáo}; */
            bcThang.Add("{X9}", "0");
            so2 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung).Sum(row => row.Field<long>("tong_luot"));
            if (so2 != 0)
            {
                so1 = b02TQ.Where(row => row.Field<string>("ma_vung") == mavung).Sum(row => row.Field<long>("tong_luot_noi"));
                bcThang["{X9}"] = ((so1 / so2) * 100).ToString();
            }
            /* X10 ={đoạn văn tùy thuộc X5> hay < X9. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThang.Add("{X10}", "bằng");
            so1 = (double)dataTinhB02["tyle_noitru"];
            so2 = double.Parse(bcThang["{X9}"]); bcThang["{X9}"] = bcThang["{X9}"].ToString();
            if (so1 > so2) { bcThang["{X10}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThang["{X10}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X11= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort cột G (TYLE_NOITRU ) cao –thấp và lấy thứ tự} */
            sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                .OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThang.Add("{X11}", position.ToString());

            /* X12 = Ngày điều trị bình quân X12={Cột H NGAY_DTRI_BQ , dòng MA_TINH=10}; */
            bcThang.Add("{X12}", dataTinhB02["ngay_dtri_bq"].ToString());
            /* X13 = Nbình quân toàn quốc X13={cột H NGAY_DTRI_BQ, dòng MA_TINH=00}; */
            bcThang.Add("{X13}", dataTQB02["ngay_dtri_bq"].ToString());
            /* X14 = Số chênh lệch X14={đoạn văn tùy thuộc X12> hay < X13. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThang.Add("{X14}", "bằng");
            so1 = (double)dataTinhB02["ngay_dtri_bq"];
            so2 = (double)dataTQB02["ngay_dtri_bq"];
            if (so1 > so2) { bcThang["{X14}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThang["{X14}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X15 = xếp thứ so toàn quốc X15={Sort cột H (NGAY_DTRI_BQ) cao xuống thấp và lấy thứ tự}; */
            sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThang.Add("{X15}", position.ToString());
            /* X16 = Bình quân vùng X16 ={tính toán: A-Tổng ngày điều trị nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của cá tỉnh cùng mã vùng. A=Total(cột H (NGAY_DTRI_BQ) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
            bcThang.Add("{X16}", "0");
            so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<long>("tong_luot_noi"));
            if (so2 != 0)
            {
                so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => (r.Field<double>("ngay_dtri_bq") * r.Field<long>("tong_luot_noi")));
                bcThang["{X16}"] = (so1 / so2).ToString();
            }
            /* X17 = Số chênh lệch X17 ={đoạn văn tùy thuộc X12> hay < X16. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThang.Add("{X17}", "bằng");
            so1 = (double)dataTinhB02["ngay_dtri_bq"];
            so2 = double.Parse(bcThang["{X16}"]); bcThang["{X16}"] = bcThang["{X16}"].ToString();
            if (so1 > so2) { bcThang["{X17}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThang["{X17}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X18 = đứng thứ so với vùng X18 = {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột H (NGAY_DTRI_BQ) cao –thấp và lấy thứ tự} */
            sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == mavung)
                .OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThang.Add("{X18}", position.ToString());

            /* X19 = Chi bình quân chung X19={Cột I (CHI_BQ_CHUNG), dòng MA_TINH=10}; */
            tmpD = buildbcThangB02(19, "chi_bq_chung", "chi_bq_chung", "tong_luot", "tong_chi", mavung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X26 = Chi bình quân ngoại trú X26={Cột J (CHI_BQ_NGOAI), dòng MA_TINH=10}; */
            tmpD = buildbcThangB02(26, "chi_bq_ngoai", "chi_bq_chung", "tong_luot_ngoai", "tong_chi_ngoai", mavung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X33 = Chi bình quân nội trú X33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
            tmpD = buildbcThangB02(33, "chi_bq_noi", "chi_bq_chung", "tong_luot_noi", "tong_chi_noi", mavung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }

            /* ----- Dữ liệu X40 trở lên lọc dữ liệu tù B26 ------- */
            /* X40 = Bình quân xét nghiệm X40= {cột P (bq_xn) dòng có mã tỉnh = 10}; B26 */
            tmpD = buildbcThangB26(40, "bq_xn", "bq_xn_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X43 Bình quân CĐHA X43= {cột R(bq_cdha) dòng có mã tỉnh =10}; */
            tmpD = buildbcThangB26(43, "bq_cdha", "bq_cdha_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X46 Bình quân thuốc X46= {cột T(bq_thuoc) dòng có mã tỉnh =10}; */
            tmpD = buildbcThangB26(46, "bq_thuoc", "bq_thuoc_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X49 Bình quân chi phẫu thuật X49= {cột V(bq_pt) dòng có mã tỉnh =10}; */
            tmpD = buildbcThangB26(49, "bq_pt", "bq_pt_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X52 Bình quân chi thủ thuật X52= {cột X(bq_tt) dòng có mã tỉnh =10}; */
            tmpD = buildbcThangB26(52, "bq_tt", "bq_tt_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X55 Bình quân chi vật tư y tế X55= {cột Z(bq_vtyt) dòng có mã tỉnh =10}; */
            tmpD = buildbcThangB26(55, "bq_vtyt", "bq_vtyt_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X58 Bình quân chi tiền giường X58= {cột AB(bq_giuong) dòng có mã tỉnh =10}; */
            tmpD = buildbcThangB26(58, "bq_giuong", "bq_giuong_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }

            /* X61 Chỉ định xét nghiệm X61={cột AD, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
            tmpD = buildbcThang02B26(61, "chi_dinh_xn", "chi_dinh_xn_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }
            /* X64 =  Chỉ định CĐHA X64={cột AF, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
            tmpD = buildbcThang02B26(64, "chi_dinh_cdha", "chi_dinh_cdha_tang", dataTinhB26);
            foreach (var d in tmpD) { bcThang.Add(d.Key, d.Value); }

            /* X67 Công tác kiểm soát chi X67={lần đầu lập BC sẽ rỗng, người dùng tự trình bày văn bản, lưu lại ở bảng dữ liệu kết quả báo cáo, kỳ sau sẽ tự động lấy từ kỳ trước, để người dùng kế thừa, sửa và lưu dùng cho kỳ này và kỳ sau} */
            bcThang.Add("{X67}", x67);
            /* X68 Công tác thanh, quyết toán năm X68={tương tự X67} */
            bcThang.Add("{X68}", x68);
            /* X69 Phương hướng kỳ tiếp theo X69={tương tự X67} */
            bcThang.Add("{X69}", x69);
            /* X70 Khó khăn, vướng mắc, đề xuất (nếu có) X70={tương tự X67} */
            bcThang.Add("{X70}", x70);

            /* X71 = {cột S T_BHTT_NOI bảng B02_TOANQUOC } */
            bcThang.Add("{X71}", dataTinhB02["t_bhtt_noi"].ToString());
            /* X72 = {cột T T_BHTT_NGOAI bảng B02_TOANQUOC } */
            bcThang.Add("{X72}", dataTinhB02["t_bhtt_ngoai"].ToString());
            /* X73 Lấy tên tỉnh */
            tmp = $"{AppHelper.dbSqliteMain.getValue($"SELECT ten FROM dmTinh WHERE id='{maTinh.sqliteGetValueField()}'")}";
            bcThang.Add("{X73}", tmp);
            /* X74 Lấy ngày chọn báo cáo */
            bcThang.Add("{X74}", ngayTime.ToString("dd/MM/yyyy"));

            bcThang.Add("ma_tinh", maTinh);
            bcThang.Add("userid", idUser);
            bcThang.Add("ngay", ngayTime.toTimestamp().ToString());
            bcThang.Add("timecreate", timeCreate);
            /* Tự động cập nhật vào dữ tuyết giao */
            if (x3 != "0")
            {
                var item = new Dictionary<string, string>() {
                    { "namqd", $"{ngayTime.Year}" },
                    { "idtinh", maTinh },
                    { "idhuyen", "" },
                    { "so_kyhieu_qd", x2},
                    { "tong_dutoan", x3 },
                    { "iduser", idUser }
                };
                AppHelper.dbSqliteWork.Update("dutoangiao", item, "replace");
            }
            return bcThang;
        }

        private void createFilebcThangDocx(string idBaoCao, string idtinh, Dictionary<string, string> bcThang)
        {
            string pathFileTemplate = Path.Combine(AppHelper.pathAppData, "baocaothang.docx");
            if (System.IO.File.Exists(pathFileTemplate) == false) { throw new Exception("Không tìm thấy tập tin mẫu báo cáo 'baocaotuan.docx' trong thư mục App_Data"); }
            /*** 1.1 làm tròn đến triệu đồng (x1, x71, x72, x2, x3, x4) */
            bcThang["{X1}"] = bcThang["{X1}"].lamTronTrieuDong();
            bcThang["{X71}"] = bcThang["{X71}"].lamTronTrieuDong();
            bcThang["{X72}"] = bcThang["{X72}"].lamTronTrieuDong();
            bcThang["{X3}"] = bcThang["{X3}"].lamTronTrieuDong();

            /* Số tiền làm tròn đến đồng */
            var tronSo = new List<string>() { "{X19}", "{X20}", "{X23}", "{X26}", "{X27}", "{X30}", "{X33}", "{X34}", "{X37}", "{X40}", "{X43}", "{X46}", "{X49}", "{X52}", "{X55}", "{X58}" };
            foreach (var v in tronSo) { if (bcThang[v].Contains(".")) { bcThang[v] = Math.Round(double.Parse(bcThang[v]), 0).ToString(); } }
            var tmp = "";
            using (var fileStream = new FileStream(pathFileTemplate, FileMode.Open, FileAccess.Read))
            {
                var document = new NPOI.XWPF.UserModel.XWPFDocument(fileStream);
                foreach (var paragraph in document.Paragraphs)
                {
                    foreach (var run in paragraph.Runs)
                    {
                        tmp = run.ToString();
                        /* Sử dụng Regex để tìm tất cả các match */
                        MatchCollection matches = Regex.Matches(tmp, "{x[0-9]+}", RegexOptions.IgnoreCase);
                        foreach (System.Text.RegularExpressions.Match match in matches) { tmp = tmp.Replace(match.Value, bcThang.getValue(match.Value, "", true)); }
                        run.SetText(tmp, 0);
                    }
                }
                tmp = Path.Combine(AppHelper.pathAppData, "bcThang", $"tinh{idtinh}");
                if (Directory.Exists(tmp) == false) { Directory.CreateDirectory(tmp); }
                tmp = Path.Combine(tmp, $"bcThang_{idBaoCao}.docx");
                if (System.IO.File.Exists(tmp)) { System.IO.File.Delete(tmp); }
                using (FileStream stream = new FileStream(tmp, FileMode.Create, FileAccess.Write)) { document.Write(stream); }
                /*
                 * MemoryStream memoryStream = new MemoryStream();
                        document.Write(memoryStream);
                        memoryStream.Position = 0;
                        return File(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{data.Rows[0]["ma_tinh"]}_{thoigian}.docx");
                */
            }
        }

        public ActionResult Update()
        {
            var idtinh = $"{Session["idtinh"]}";
            if (idtinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            var id = Request.getValue("objectid");
            var tsql = "";
            ViewBag.id = id;
            try
            {
                var item = new Dictionary<string, string>();
                var dbBaoCao = BuildDatabase.getDataBaoCaoTuan(idtinh);
                if (Request.getValue("mode") == "update")
                {
                    var timeStart = DateTime.Now;
                    item = new Dictionary<string, string>() {
                        { "x2", Request.getValue("x2").sqliteGetValueField() },
                        { "x3", Request.getValue("x3").Trim() },
                        { "x67", Request.getValue("x67").sqliteGetValueField() },
                        { "x68", Request.getValue("x68").sqliteGetValueField() },
                        { "x69", Request.getValue("x69").sqliteGetValueField() },
                        { "x70", Request.getValue("x70").sqliteGetValueField() }
                    };
                    if (item["x3"].isNumberUSInt() == false) { return Content($"Tổng số tiền các dòng QĐ năm nay không đúng định dạng '{item["x3"]}'".BootstrapAlter("warning")); }
                    if (item["x3"] == "0") { return Content("Chưa điền Tổng số tiền các dòng QĐ năm nay".BootstrapAlter("warning")); }
                    tsql = $"UPDATE bcThangdocx SET x2='{item["x2"]}', x3='{item["x3"]}', x67='{item["x67"]}', x68='{item["x68"]}', x69='{item["x69"]}', x70='{item["x70"]}', x4=ROUND((x1/{item["x3"]})*100,2) WHERE id='{id.sqliteGetValueField()}'";
                    dbBaoCao.Execute(tsql);
                    tsql = $"SELECT * FROM bcThangdocx WHERE id='{id.sqliteGetValueField()}'";
                    var data = dbBaoCao.getDataTable(tsql);
                    dbBaoCao.Close();
                    if (data.Rows.Count == 0)
                    {
                        ViewBag.Error = $"Báo cáo tuần có ID '{id}' thuộc tỉnh có mã '{idtinh}' không tồn tại hoặc đã bị xoá khỏi hệ thống";
                        return View();
                    }
                    var bcThang = new Dictionary<string, string>();
                    foreach (DataColumn c in data.Columns) { bcThang.Add("{" + c.ColumnName.ToUpper() + "}", $"{data.Rows[0][c.ColumnName]}"); }
                    createFilebcThangDocx(id, idtinh, bcThang);
                    if (item["x3"] != bcThang["{X3}"])
                    {
                        var duToanGiao = new Dictionary<string, string>() {
                            { "namqd", bcThang["{X74}"].Substring(7) },
                            { "idtinh", idtinh },
                            { "idhuyen", "" },
                            { "so_kyhieu_qd", item["x2"]},
                            { "tong_dutoan", item["x3"] },
                            { "iduser", $"{Session["iduser"]}" }
                        };
                        AppHelper.dbSqliteWork.Update("dutoangiao", item, "replace");
                    }
                    return Content($"Lưu thành công ({timeStart.getTimeRun()})".BootstrapAlter());
                }
                tsql = $"SELECT * FROM bcThangdocx WHERE id='{id.sqliteGetValueField()}'";
                var d = dbBaoCao.getDataTable(tsql);
                dbBaoCao.Close();
                if (d.Rows.Count == 0)
                {
                    ViewBag.Error = $"Báo cáo tuần có ID '{id}' thuộc tỉnh có mã '{idtinh}' không tồn tại hoặc đã bị xoá khỏi hệ thống.";
                    return View();
                }
                foreach (DataColumn c in d.Columns) { item.Add($"{c.ColumnName}", $"{d.Rows[0][c.ColumnName]}"); }
                ViewBag.data = item;
            }
            catch (Exception ex) { ViewBag.Error = $"Lỗi: {ex.getErrorSave()}"; }
            return View();
        }

        public ActionResult TruyVan()
        {
            var matinh = $"{Session["idtinh"]}";
            if (matinh == "") { ViewBag.Error = "Bạn chưa cấp Mã tỉnh làm việc"; return View(); }
            /* Tài khoản system có thể xem được tất cả
             * Tài khoản admin tỉnh xem được toàn bộ của tỉnh được phân
             * Tải khoản người dùng chỉnh xem các báo cáo mình tạo ra
             */
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
                    var dbbcThang = BuildDatabase.getDataBaoCaoTuan(matinh);
                    var where = $"WHERE timecreate >= {time1.toTimestamp()} AND timecreate < {time2.AddDays(1).toTimestamp()}";
                    var tmp = $"{Session["nhom"]}";
                    if (tmp != "0" && tmp != "1") { where += $" AND userid='{Session["iduser"]}'"; }
                    var tsql = $"SELECT datetime(timecreate, 'auto', '+7 hour') AS ngayGM7,id,ma_tinh,x72,x74,userid FROM bcThangdocx {where} ORDER BY timecreate DESC";
                    ViewBag.data = dbbcThang.getDataTable(tsql);
                    dbbcThang.Close();
                    ViewBag.tsql = tsql;
                    return View();
                }
            }
            catch (Exception ex) { ViewBag.Error = ex.getLineHTML(); }
            return View();
        }

        public ActionResult Delete()
        {
            var timeStart = DateTime.Now;
            string ids = Request.getValue("id");
            var lid = new List<string>();
            string mode = Request.getValue("mode");
            try
            {
                if (string.IsNullOrEmpty(ids)) { return Content("Không có tham số".BootstrapAlter("warning")); }
                /* Kiểm tra danh sách nếu có */
                lid = ids.Split(new[] { '|', ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                ViewBag.data = string.Join(",", lid);
                if (mode == "force")
                {
                    foreach (string id in lid) { DeletebcThang(id, true); }
                    return Content($"Xoá thành công báo cáo có ID '{string.Join(", ", lid)}' ({timeStart.getTimeRun()})".BootstrapAlter());
                }
            }
            catch (Exception ex) { return Content(ex.getErrorSave().BootstrapAlter("warning")); }
            return View();
        }

        private void DeletebcThang(string id, bool throwEx = false)
        {
            /* ID: {yyyyMMddHHmmss}_{idtinh}_{Milisecon}*/
            var tmpl = id.Split('_');
            if (tmpl.Length != 3)
            {
                if (throwEx == false) { return; }
                throw new Exception("ID Báo cáo không đúng định dạng {yyyyMMddHHmmss}_{idtinh}_{Milisecon}: " + id);
            }
            string idtinh = tmpl[1];
            /* Xoá hết các file trong mục lưu trữ App_Data/bcThang */
            var folder = new DirectoryInfo(Path.Combine(AppHelper.pathApp, "App_Data", "bcThang", $"tinh{idtinh}"));
            if (folder.Exists)
            {
                foreach (var f in folder.GetFiles($"bcThang_{id}.*")) { try { f.Delete(); } catch { } }
                foreach (var f in folder.GetFiles($"bcThang_pl_{id}*.*")) { try { f.Delete(); } catch { } }
                foreach (var f in folder.GetFiles($"id{id}*.*")) { try { f.Delete(); } catch { } }
            }
            /* Xoá trong cơ sở dữ liệu */
            var db = BuildDatabase.getDataBaoCaoTuan(idtinh);
            try
            {
                var idBaoCao = id.sqliteGetValueField();
                db.Execute($@"DELETE FROM bcThangdocx WHERE id='{idBaoCao}';
                        DELETE FROM pl01 WHERE id_bc='{idBaoCao}';
                        DELETE FROM pl02 WHERE id_bc='{idBaoCao}';
                        DELETE FROM pl03 WHERE id_bc='{idBaoCao}';");
                db.Close();
                db = BuildDatabase.getDataImportBaoCaoTuan(idtinh);
                db.Execute($@"DELETE FROM b02 WHERE id_bc='{idBaoCao}';
                        DELETE FROM b04 WHERE id_bc='{idBaoCao}';
                        DELETE FROM b26 WHERE id_bc='{idBaoCao}';
                        DELETE FROM b02chitiet WHERE id_bc='{idBaoCao}';
                        DELETE FROM b04chitiet WHERE id_bc='{idBaoCao}';
                        DELETE FROM b26chitiet WHERE id_bc='{idBaoCao}';");
            }
            catch (Exception ex)
            {
                var msg = ex.getErrorSave();
                if (throwEx) { throw new Exception(msg); }
            }
            finally { db.Close(); }
        }
    }
}