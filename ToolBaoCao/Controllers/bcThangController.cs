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
    public class bcThangController : ControllerCheckLogin
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
            var id = $"{timeStart:yyMMddHHmmss}_{matinh}_{timeStart.Millisecond:000}";
            var timeUp = timeStart.toTimestamp().ToString();
            var folderTemp = Path.Combine(AppHelper.pathApp, "temp", "bcThang", $"{matinh}_{Session["iduser"]}".GetMd5Hash());
            var tmp = "";
            ViewBag.id = id;
            try
            {
                /* Xoá hết các File có trong thư mục */
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
                /* Khai báo dữ liệu tạm */
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                dbTemp.CreateImportBcThang();
                dbTemp.CreatePhucLucBcThang();
                dbTemp.CreateBcThang();
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
                bieus = bieus.Distinct().ToList();
                if (bieus.Count != 11) { throw new Exception($"Dư biểu hoặc thiếu biểu đầu vào. {string.Join(", ", bieus)}"); }
                if (bieus.Where(p => p.StartsWith("b01")).Count() != 3) { throw new Exception($"Dư biểu hoặc thiếu biểu đầu vào B01. {string.Join(", ", bieus)}"); }
                if (bieus.Where(p => p.StartsWith("b02")).Count() != 6) { throw new Exception($"Dư biểu hoặc thiếu biểu đầu vào B02. {string.Join(", ", bieus)}"); }
                if (bieus.Where(p => p.StartsWith("b04")).Count() != 2) { throw new Exception($"Dư biểu hoặc thiếu biểu đầu vào B04. {string.Join(", ", bieus)}"); }
                if (list.Count > 0) { throw new Exception(string.Join("<br />", list)); }
                /* Tạo Phục Lục 1 - Lấy từ nguồn cơ sở luỹ kế */
                tmp = $"{dbTemp.getValue($"SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='{matinh}' AND tu_thang=1 ORDER BY nam DESC LIMIT 1")}";
                var tsql = $@"INSERT INTO thangpl01 (id_bc
                    ,idtinh
                    ,ma_cskcb
                    ,ten_cskcb
                    ,dtgiao
                    ,tien_bhtt
                    ,tl_sudungdt
                    ,userid) SELECT '{id}' AS id_bc, '{matinh}' AS idtinh, ma_cskcb, ten_cskcb, 0 AS dtgiao, t_bhtt, 0 AS tl_sudungdt, '{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2='{tmp}';";
                dbTemp.Execute(tsql);
                /* Tạo Phục Lục 2a */
                /* Lấy dữ liệu từ biểu pl02a trong tháng (Từ tháng đến tháng = tháng báo cáo của toàn quốc nam1) */
                tmp = $"{dbTemp.getValue($"SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='00' AND tu_thang=den_thang ORDER BY nam DESC LIMIT 1")}";
                dbTemp.Execute($@"INSERT INTO thangpl02a (id_bc, idtinh
                ,ma_tinh
                ,ten_tinh
                ,ma_vung
                ,tyle_noitru
                ,ngay_dtri_bq
                ,chi_bq_chung
                ,chi_bq_ngoai
                ,chi_bq_noi, userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    ,ROUND(tyle_noitru, 2) AS tyle_noitru
                    ,ROUND(ngay_dtri_bq) AS ngay_dtri_bq
                    ,ROUND(chi_bq_chung) AS chi_bq_chung
                    ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai
                    ,ROUND(chi_bq_noi) AS chi_bq_noi
                    ,'{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2 = '{tmp}';");
                /* Tạo Phục Lục 2b */
                /* Lấy dữ liệu từ biểu b02 dành cho cả năm (từ tháng 1 đến tháng báo cáo) */
                tmp = $"{dbTemp.getValue($"SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='00' AND tu_thang=1 ORDER BY nam DESC LIMIT 1")}";
                dbTemp.Execute($@"INSERT INTO thangpl02b (id_bc, idtinh
                ,ma_tinh
                ,ten_tinh
                ,ma_vung
                ,tyle_noitru
                ,ngay_dtri_bq
                ,chi_bq_chung
                ,chi_bq_ngoai
                ,chi_bq_noi, userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    ,ROUND(tyle_noitru, 2) AS tyle_noitru
                    ,ROUND(ngay_dtri_bq) AS ngay_dtri_bq
                    ,ROUND(chi_bq_chung) AS chi_bq_chung
                    ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai
                    ,ROUND(chi_bq_noi) AS chi_bq_noi
                    ,'{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2 = '{tmp}';");
                /* Tạo Phục Lục 3a */
                /* Lấy dữ liệu từ biểu b02 csyt trong tháng */
                tmp = $"{dbTemp.getValue($"SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='{matinh}' AND tu_thang=den_thang ORDER BY nam DESC LIMIT 1")}";
                var data = dbTemp.getDataTable($@"SELECT id_bc, '{matinh}' as idtinh, ma_cskcb, ten_cskcb, ma_vung
                    ,ROUND(tyle_noitru, 2) AS tyle_noitru
                    ,ROUND(ngay_dtri_bq) AS ngay_dtri_bq
                    ,ROUND(chi_bq_chung) AS chi_bq_chung
                    ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai
                    ,ROUND(chi_bq_noi) AS chi_bq_noi
                    ,'' as tuyen_bv, '' as hang_bv,'{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2 = '{tmp}';");
                /* Lấy danh sách Ma_CSKCB */
                var dsCSYT = AppHelper.dbSqliteMain.getDataTable($"SELECT id, tuyencmkt, hangdv FROM dmcskcb WHERE ma_tinh ='{matinh}'");
                var dsCSKCB = dsCSYT.AsEnumerable().Select(x => new
                {
                    id = x.Field<string>("id"),
                    tuyen = string.IsNullOrEmpty(x.Field<string>("tuyencmkt")) ? "*" : x.Field<string>("tuyencmkt"),
                    hang = string.IsNullOrEmpty(x.Field<string>("hangdv")) ? "*" : x.Field<string>("hangdv")
                }).ToList();
                foreach (DataRow row in data.Rows)
                {
                    tmp = $"{row["ma_cskcb"]}";
                    var v = dsCSKCB.FirstOrDefault(x => x.id == tmp);
                    if (v == null) { row["tuyen_bv"] = "*"; row["hang_bv"] = "*"; }
                    else
                    {
                        row["tuyen_bv"] = v.tuyen;
                        row["hang_bv"] = v.hang.ToLower().StartsWith("h") ? v.hang : "*";
                    }
                }
                dbTemp.Insert("thangpl03a", data);
                /* Tạo phục lục 03b */
                /* Cách lập giống như Phụ lục 03 báo cáo tuần, nguồn dữ liệu lấy từ B02 từ tháng 1 đến tháng báo cáo */
                tmp = $"{dbTemp.getValue($"SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='{matinh}' AND tu_thang=den_thang ORDER BY nam DESC LIMIT 1")}";
                data = dbTemp.getDataTable($@"SELECT id_bc, '{matinh}' as idtinh, ma_cskcb, ten_cskcb, ma_vung
                    ,ROUND(tyle_noitru, 2) AS tyle_noitru
                    ,ROUND(ngay_dtri_bq) AS ngay_dtri_bq
                    ,ROUND(chi_bq_chung) AS chi_bq_chung
                    ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai
                    ,ROUND(chi_bq_noi) AS chi_bq_noi
                    ,'' as tuyen_bv, '' as hang_bv, '{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2 = '{tmp}';");
                foreach (DataRow row in data.Rows)
                {
                    tmp = $"{row["ma_cskcb"]}";
                    var v = dsCSKCB.FirstOrDefault(x => x.id == tmp);
                    if (v == null) { row["tuyen_bv"] = "*"; row["hang_bv"] = "*"; }
                    else
                    {
                        row["tuyen_bv"] = v.tuyen;
                        row["hang_bv"] = v.hang.ToLower().StartsWith("h") ? v.hang : "*";
                    }
                }
                dbTemp.Insert("thangpl03b", data);
                /* Tạo thangpl04a */
                /* Nguồn dữ liệu B04_00 từ tháng 1 đến tháng báo cáo. Giống như Phụ lục 2 của báo cáo tuần. */
                tmp = $"{dbTemp.getValue($"SELECT id FROM thangb04 WHERE id_bc='{id}' AND ma_tinh='00' AND tu_thang=1 ORDER BY nam DESC LIMIT 1")}";
                dbTemp.Execute($@"INSERT INTO thangpl04a (id_bc, idtinh, ma_tinh, ten_tinh, ma_vung, chi_bq_xn, chi_bq_cdha, chi_bq_thuoc, chi_bq_pttt, chi_bq_vtyt, chi_bq_giuong, ngay_ttbq, userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    ,ROUND(bq_xn) AS chi_bq_xn
                    ,ROUND(bq_cdha) AS chi_bq_cdha
                    ,ROUND(bq_thuoc) AS chi_bq_thuoc
                    ,ROUND(bq_ptt) AS chi_bq_pttt
                    ,ROUND(bq_vtyt) AS chi_bq_vtyt
                    ,ROUND(bq_giuong) AS chi_bq_giuong
                    ,ROUND(ngay_ttbq, 2) AS ngay_ttbq, '{idUser}' AS userid
                    FROM thangb04chitiet WHERE id_bc='{id}' AND id2='{tmp}';");
                /* Tạo thangpl04b */
                /* Nguồn dữ liệu B04_10 của tháng báo cáo. Giống như Phụ lục 2 của báo cáo tuần, nhưng chi tiết từng CSKCB và phân nhóm theo tuyến tỉnh huyện xã */
                tmp = $"{dbTemp.getValue($"SELECT id FROM thangb04 WHERE id_bc='{id}' AND ma_tinh='{matinh}' AND tu_thang=den_thang ORDER BY nam DESC LIMIT 1")}";
                tsql = $@"SELECT id_bc, '{matinh}' as idtinh, ma_cskcb, ten_cskcb, ma_vung
                    ,ROUND(bq_xn) AS chi_bq_xn
                    ,ROUND(bq_cdha) AS chi_bq_cdha
                    ,ROUND(bq_thuoc) AS chi_bq_thuoc
                    ,ROUND(bq_ptt) AS chi_bq_pttt
                    ,ROUND(bq_vtyt) AS chi_bq_vtyt
                    ,ROUND(bq_giuong) AS chi_bq_giuong
                    ,ROUND(ngay_ttbq, 2) AS ngay_ttbq
                    ,'' as tuyen_bv, '' as hang_bv, '{idUser}' AS userid
                    FROM thangb04chitiet WHERE id_bc='{id}' AND id2='{tmp}';";
                data = dbTemp.getDataTable(tsql);
                foreach (DataRow row in data.Rows)
                {
                    tmp = $"{row["ma_cskcb"]}";
                    var v = dsCSKCB.FirstOrDefault(x => x.id == tmp);
                    if (v == null) { row["tuyen_bv"] = "*"; row["hang_bv"] = "*"; }
                    else
                    {
                        row["tuyen_bv"] = v.tuyen;
                        row["hang_bv"] = v.hang.ToLower().StartsWith("h") ? v.hang : "*";
                    }
                }
                dbTemp.Insert("thangpl04b", data);
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

        private void createFilePhuLucbcThang(string idBaoCao, string matinh, dbSQLite dbBCThang = null, Dictionary<string, string> bcThang = null)
        {
            if (dbBCThang == null) { dbBCThang = BuildDatabase.getDataBCThang(matinh); }
            var fileName = $"bcThang_pl_{idBaoCao}.xlsx";
            var pathPLBCThang = Path.Combine(AppHelper.pathApp, "App_Data", "bcThang", $"tinh{matinh}", fileName);
            System.IO.File.Copy(Path.Combine(AppHelper.pathAppData, "plthang.xlsx"), pathPLBCThang, true);
            var idBaoCaoVauleField = idBaoCao.sqliteGetValueField();
            if (bcThang == null)
            {
                bcThang = new Dictionary<string, string>();
                var data = dbBCThang.getDataTable($"SELECT * FROM bcThangdocx WHERE id='{idBaoCaoVauleField}';");
                if (data.Rows.Count > 0)
                {
                    foreach (DataColumn c in data.Columns)
                    {
                        bcThang.Add("{" + c.ColumnName.ToUpper() + "}", $"{data.Rows[0][c.ColumnName]}");
                    }
                }
            }
            /* Tạo phụ lục báo cáo */
            dbBCThang.Execute($"UPDATE thangpl01 SET tl_sudungdt = 0 WHERE id_bc='{idBaoCaoVauleField}' AND dtgiao = 0;");
            dbBCThang.Execute($"UPDATE thangpl01 SET tl_sudungdt = ROUND(tien_bhtt/dtgiao, 2) WHERE id_bc='{idBaoCaoVauleField}' AND dtgiao > 0;");
            var PL01 = dbBCThang.getDataTable($"SELECT ma_cskcb, ten_cskcb, dtgiao, tien_bhtt, tl_sudungdt FROM thangpl01 WHERE id_bc='{idBaoCaoVauleField}' ORDER BY ma_cskcb;");
            PL01.TableName = "PL01";

            var pl = dbBCThang.getDataTable($"SELECT * FROM thangpl02a WHERE id_bc='{idBaoCaoVauleField}';");
            var phuluc02 = createPhuLuc02(pl, matinh);

            pl = dbBCThang.getDataTable($"SELECT * FROM thangpl03a WHERE id_bc='{idBaoCaoVauleField}';");
            var phuluc03 = createPhuLuc03(pl, matinh, PL01);

            var xlsx = exportPhuLucbcThang(PL01, phuluc02, phuluc03);

            var tmp = Path.Combine(AppHelper.pathApp, "App_Data", "bcThang", $"tinh{matinh}", fileName);
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
            var matinh = $"{Session["idtinh"]}".Trim();
            var listBieu = new List<string>();
            string bieu = "";
            string fileExtension = Path.GetExtension(inputFile.FileName);
            int sheetIndex = 0; int packetSize = 1000;
            int indexRow = 0; int indexColumn = 0; int maxRow = 0; int jIndex = 0;
            int fieldCount = 50; var tsql = new List<string>();
            var tmp = "";
            IWorkbook workbook = null;
            try
            {
                try { workbook = new XSSFWorkbook(inputFile.InputStream); }
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
                        if (tmp.StartsWith("b01")) { bieu = "b01"; /* 3 b01; b0100_nam1 b0100_nam2 b01cs_nam1 */ }
                        if (tmp.StartsWith("b02")) { bieu = "b02"; /* 6 b02: b0200_nam1 b0200_nam2 b0200_thang1 b0200_thang2 b02cs_nam1 b02cs_thang1 */ }
                        if (tmp.StartsWith("b04")) { bieu = "b04"; /* 2 b04: b0400_nam1 b04cs_thang1 */ }
                        if (tmp == "ma_tinh") { indexColumn = c.ColumnIndex; break; }
                    }
                    if (tmp == "ma_tinh") { break; }
                }
                /* Không xác định được biểu thì bỏ qua */
                if (bieu == "") { workbook.Close(); return ""; }
                if (indexRow >= maxRow) { throw new Exception("Không có dữ liệu"); }
                string pattern = "^20[0-9][0-9]$";
                int indexRegex = 3; int tmpInt = 0;
                /* Bắt đầu đọc dữ liệu
                 * - Đọc thông số biểu
                 * Biểu b01: ma_tinh    tu_thang    den_thang   nam         cs
                 * Biểu b02: ma_tinh	ma_loai_kcb	tu_thang	den_thang	nam	loai_bv	kieubv	loaick	hang_bv	tuyen   cs
                 * Biểu b04: ma_tinh	tu_thang	den_thang	nam	ma_loai_kcb	loai_bv	hang_bv	tuyen	kieubv	loaick	cs
                 */
                switch (bieu)
                {
                    /* Kiểm tra năm */
                    case "b01": fieldCount = 5; indexRegex = 3; pattern = "^20[0-9][0-9]$"; break;
                    /* Kiểm tra năm */
                    case "b02": fieldCount = 11; indexRegex = 4; pattern = "^20[0-9][0-9]$"; break;
                    /* Kiểm tra thoigian */
                    case "b04": fieldCount = 11; indexRegex = 3; pattern = "^20[0-9][0-9]$"; break;
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
                var idChiTiet = $""; /*NamThang1Thang2MaTinh*/
                switch (bieu)
                {
                    case "b01":
                        /* 3 b01; b0100_nam1 b0100_nam2 b01cs_nam1 */
                        /* ma_tinh    tu_thang    den_thang   nam         cs */
                        idChiTiet = $"{listValue[0]}_{listValue[3]}{(listValue[2].Length < 2 ? $"0{listValue[2]}" : listValue[2])}{(listValue[1].Length < 2 ? $"0{listValue[1]}" : listValue[1])}";
                        listBieu.Add($"b01{idChiTiet}");
                        if (listValue[1] != "1") { throw new Exception($"Biểu {bieu} yêu cầu từ tháng 1; Tháng từ của biểu là '{listValue[1]}'"); }
                        break;

                    case "b02":
                        /* 6 b02: b0200_nam1 b0200_nam2 b0200_thang1 b0200_thang2 b02cs_nam1 b02cs_thang1 */
                        /* ma_tinh	ma_loai_kcb	tu_thang	den_thang	nam	loai_bv	kieubv	loaick	hang_bv	tuyen   cs */
                        idChiTiet = $"{listValue[0]}_{listValue[4]}{(listValue[3].Length < 2 ? $"0{listValue[3]}" : listValue[3])}{(listValue[2].Length < 2 ? $"0{listValue[2]}" : listValue[2])}";
                        listBieu.Add($"b02{idChiTiet}");
                        if (listValue[2] != listValue[3])
                        {
                            if (listValue[2] != "1") { throw new Exception($"Biểu {bieu} yêu cầu từ tháng 1; Tháng từ của biểu là '{listValue[2]}'"); }
                        }
                        break;

                    case "b04":
                        /* 2 b04: b0400_nam1 b04cs_thang1 */
                        /* ma_tinh	tu_thang	den_thang	nam	ma_loai_kcb	loai_bv	hang_bv	tuyen	kieubv	loaick	cs */
                        idChiTiet = $"{listValue[0]}_{listValue[3]}{(listValue[2].Length < 2 ? $"0{listValue[2]}" : listValue[2])}{(listValue[1].Length < 2 ? $"0{listValue[1]}" : listValue[1])}";
                        listBieu.Add($"b04{idChiTiet}");
                        if (listValue[1] != listValue[2])
                        {
                            if (listValue[1] != "1") { throw new Exception($"Biểu {bieu} yêu cầu từ tháng 1; Tháng từ của biểu là '{listValue[1]}'"); }
                        }
                        break;

                    default: fieldCount = 11; break;
                }
                /* Có phải là cơ sở không? */
                tmpInt = (fieldCount - 1);
                listValue[tmpInt] = "1";
                if (listValue[0] == "00") { listValue[tmpInt] = "0"; cs = false; }

                tmp = string.Join(",", listValue);
                if (tmp.Contains(",,")) { throw new Exception($"Biểu {bieu} không đúng định dạng."); }
                /* Kiểm tra có đúng dữ liệu không */
                if (Regex.IsMatch(listValue[indexRegex], pattern) == false) { throw new Exception($"dữ liệu không đúng cấu trúc (năm, thời gian): {listValue[indexRegex]}"); }

                /* Lấy danh sách cột, bỏ cột ID */
                bieu = $"thang{bieu}";
                var allColumns = dbConnect.getColumns(bieu).Select(p => p.ColumnName).ToList();
                allColumns.RemoveAt(0);
                /* Thêm UserID */
                listValue.Add(userID);
                listValue.Add(timeUp);
                listValue.Add(idBaoCao);
                idChiTiet = idBaoCao + "_" + idChiTiet;
                listValue.Add(idChiTiet);
                tsql.Add($"INSERT INTO {bieu} ({string.Join(",", allColumns)}, id) VALUES ('{string.Join("','", listValue)}');");
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
                    /* Kiểm tra Nguồn trong năm */
                    case "thangb01":
                        fieldCount = 20; indexRegex = 3 + 1; pattern = @"^\d+$"; /* nguồn trong năm */
                        fieldNumbers = new List<int>() { 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20 };
                        break;
                    /* Kiểm tra Tổng cộng Số lượt KCB */
                    case "thangb02":
                        fieldCount = 20; indexRegex = 3 + 1; pattern = @"^\d+$";
                        fieldNumbers = new List<int>() { 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20 };
                        break;
                    /* Kiểm tra Chi lần KCB */
                    case "thangb04":
                        fieldCount = 11; indexRegex = 2 + 1; pattern = @"^\d+([.]\d+)?$";
                        fieldNumbers = new List<int>() { 3, 4, 5, 6, 7, 8, 9, 10 };
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

                    if (Regex.IsMatch(ma, "^V?[0-9]+$") == false) { continue; }
                    /* Xây dựng tsql VALUES */
                    listValue = new List<string>() { idChiTiet, ma.sqliteGetValueField() };
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
                /* System.IO.File.WriteAllText(Path.Combine(folderTemp, $"id{idBaoCao}_{listBieu[0]}_{matinhImport}.sql"), tmp); */
                dbConnect.Execute(tmp);
                if (tsql.Count < 2) { throw new Exception("Không có dữ liệu chi tiết"); }
                /* Lưu lại file */
                using (FileStream stream = new FileStream(Path.Combine(folderTemp, $"id{idBaoCao}_{listBieu[0]}{fileExtension}"), FileMode.Create, FileAccess.Write)) { workbook.Write(stream); }
            }
            catch (Exception ex2)
            {
                messageError = $"Lỗi trong quá trình đọc, nhập dữ liệu từ Excel '{inputFile.FileName}': {ex2.getLineHTML()}";
                AppHelper.saveError(tmp);
            }
            finally
            {
                /* Xoá luôn dữ liệu tạm của IIS */
                if (workbook != null) { workbook.Close(); workbook = null; }
            }
            if (messageError != "") { throw new Exception(messageError); }
            return listBieu[0];
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
                var pathDB = Path.Combine(folderTemp, "import.db");
                if (System.IO.File.Exists(pathDB) == false) { throw new Exception($"Dữ liệu tạo báo cáo tháng có ID '{idBaoCao}' đã bị huỷ hoặc không tồn tại trên hệ thống"); }
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                /* Tạo bcThang */
                var bcThang = createbcThang(dbTemp, idBaoCao, idtinh, iduser, Request.getValue("x1"), Request.getValue("x33"), Request.getValue("x34"), Request.getValue("x35"), Request.getValue("x36"), Request.getValue("x37"), Request.getValue("x38"));
                /* Tạo docx */
                createFilebcThangDocx(idBaoCao, idtinh, bcThang);
                /* Tạo dữ liệu để xuất phụ lục */
                string idBaoCaoVauleField = idBaoCao.sqliteGetValueField();
                var dbbcThang = BuildDatabase.getDataBCThang(idtinh);
                var dbImport = BuildDatabase.getDataImportBCThang(idtinh);
                /* Tạo phụ lục báo cáo */
                /* dmCSKCB */
                var dmCSKCB = AppHelper.dbSqliteMain.getDataTable($"SELECT id, ten, macaptren FROM dmckscb WHERE ma_tinh='{idtinh}'").AsEnumerable();
                /* Di chuyển tập tin Excel */
                foreach (var f in dirTemp.GetFiles("*.xls*")) { f.MoveTo(Path.Combine(folderSave, f.Name)); }

                /* Báo cáo tháng chuyển */
                dbbcThang.Update("bcThangdocx", bcThang);
                dbbcThang.Close();

                var data = new DataTable();
                list = new List<string>() { "thangpl01", "thangpl02a", "thangpl02b", "thangpl03a", "thangpl03b", "thangpl04a", "thangpl04b" };
                foreach (var v in list)
                {
                    data = dbTemp.getDataTable($"SELECT * FROM {v} WHERE id_bc='{idBaoCaoVauleField}';");
                    data.Columns.RemoveAt(0);
                    dbbcThang.Insert(v, data);
                }
                list = new List<string>() { "thangb01", "thangb02", "thangb04" };
                foreach (var v in list)
                {
                    data = dbTemp.getDataTable($"SELECT * FROM {v} WHERE id_bc='{idBaoCaoVauleField}';");
                    dbImport.Insert(v, data);
                }
                list = new List<string>() { "thangb01chitiet", "thangb02chitiet", "thangb04chitiet" };
                foreach (var v in list)
                {
                    data = dbTemp.getDataTable($"SELECT * FROM {v} WHERE id_bc='{idBaoCaoVauleField}';");
                    data.Columns.RemoveAt(0);
                    dbImport.Insert(v, data);
                }
                createFilePhuLucbcThang(idBaoCao, idtinh, dbbcThang, bcThang);

                dbTemp.Close();
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.getErrorSave();
                DeletebcThang(idtinh);
            }
            return View();
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

        private Dictionary<string, string> createbcThang(dbSQLite dbConnect, string idBaoCao, string maTinh, string idUser, string x1 = "", string x33 = "", string x34 = "", string x35 = "", string x36 = "", string x37 = "", string x38 = "")
        {
            var bcThang = new Dictionary<string, string>() { { "id", idBaoCao }, { "x1", x1 }, { "x33", x33 }, { "x34", x34 }, { "x35", x35 }, { "x36", x36 }, { "x37", x37 }, { "x38", x38 } };
            string tmp = AppHelper.dbSqliteMain.getValue($"SELECT ten FROM dmtinh WHERE id='{maTinh}';").ToString();
            string mavung = "";
            bcThang.Add("tentinh", tmp);
            var data = dbConnect.getDataTable($"SELECT den_thang, nam FROM thangb04 WHERE id_bc='{idBaoCao}' LIMIT 1;");
            if (data.Rows.Count == 0) { throw new Exception("[creatbcThang] Biểu 04 không có dữ liệu"); }
            bcThang.Add("nam1", $"{data.Rows[0]["nam"]}");
            bcThang.Add("nam2", (int.Parse($"{data.Rows[0]["nam"]}") - 1).ToString());
            bcThang.Add("thang", $"{data.Rows[0]["den_thang"]}");
            bcThang.Add("ngay2", $"01/{bcThang["thang"]}/{bcThang["nam1"]}");
            var time = new DateTime(int.Parse(bcThang["nam1"]), int.Parse(bcThang["thang"]), 1);
            time = time.AddMonths(1).AddDays(-1);
            bcThang["ngay1"] = $"{time:dd/MM/yyyy}";

            /* ,x2 real not null default 0 /* Dự toán giao {nam}
                ,x3 real not null default 0 /* Chi KCB toàn tỉnh
                ,x4 real not null default 0 /* Tỷ lệ % SD dự toán {nam} */
            tmp = $"{dbConnect.getValue($"SELECT id FROM thangb01 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND tu_thang=1 AND nam={bcThang["nam1"]} LIMIT 1;")}";
            var ldata = dbConnect.getDataTable($"SELECT * FROM thangb01chitiet WHERE id_bc='{idBaoCao}' AND id2='{tmp}' AND ma_tinh <> '00';").AsEnumerable().ToList();
            if (ldata.Count() == 0) { throw new Exception($"[creatbcThang] Biểu 01 Toàn quốc từ tháng 1 đến {bcThang["thang"]} năm {bcThang["nam1"]} không có dữ liệu"); }
            var item = ldata.FirstOrDefault(p => p.Field<string>("ma_tinh") == maTinh);
            if (item == null) { throw new Exception($"[creatbcThang] Biểu 01 Toàn quốc từ tháng 1 đến {bcThang["thang"]} năm {bcThang["nam1"]} không có dữ liệu của tỉnh {maTinh}"); }

            mavung = $"{item["ma_vung"]}";
            bcThang.Add("x2", $"{item["dtcsyt_trongnam"]}");
            bcThang.Add("x3", $"{item["dtcsyt_chikcb"]}");
            bcThang.Add("x4", $"{item["dtcsyt_tlsudungnam"]}");

            /* x5 integer not null default 0 /* xếp bn toàn quốc */
            tmp = getPosition("", maTinh, "dtcsyt_tlsudungnam", ldata);
            bcThang.Add("x5", tmp);
            /* x6 integer not null default 0 /* xếp thứ bao nhiêu so với vùng */
            tmp = getPosition(mavung, maTinh, "dtcsyt_tlsudungnam", ldata);
            bcThang.Add("x6", $"{data.Rows[0]["dtcsyt_trongnam"]}");
            /* x7 real not null default 0 /* Tỷ lệ % SD dự toán {nam2} */
            tmp = $"{dbConnect.getValue($"SELECT id FROM thangb01 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND tu_thang=1 AND nam={bcThang["nam2"]} LIMIT 1;")}";
            bcThang.Add("x7", $"{dbConnect.getValue($"SELECT IFNULL(dtcsyt_tlsudungnam, 0) AS X FROM thangb01chitiet WHERE id_bc='{idBaoCao}' AND id2='{tmp}' AND ma_tinh = '{maTinh}'")}");
            /* x8 real not null default 0 /* So cùng kỳ năm trước = 3-6 (x4 - x7) */
            bcThang.Add("x8", Math.Round((double.Parse(bcThang["x4"]) - double.Parse(bcThang["x7"])), 2).ToString());

            /* ,x9 real not null default 0 /* Tổng lượt = 2+3 (x10+x11)
                ,x10 real not null default 0 /* Lượt ngoại {nam1}
                ,x11 real not null default 0 /* Lượt nội {nam1}
            ,x21 real not null default 0 /* Tổng chi = 2+3 (x22+x23)
                ,x22 real not null default 0 /* Chi ngoại trú {nam1}
                ,x23 real not null default 0 /* Chi nội trú {nam1}  */
            tmp = $"{dbConnect.getValue($"SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND tu_thang=den_thang AND nam='{bcThang["nam1"]}' LIMIT 1")}";
            item = dbConnect.getDataTable($"SELECT * FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2='{tmp}' AND ma_tinh='{maTinh}' LIMIT 1").AsEnumerable().FirstOrDefault();
            if (item == null) { throw new Exception($"[creatbcThang] Biểu 02 Toàn quốc tháng {bcThang["thang"]} năm {bcThang["nam1"]} không có dữ liệu của tỉnh {maTinh}"); }
            bcThang.Add("x9", $"{item["tong_luot"]}");
            bcThang.Add("x10", $"{item["tong_luot_ngoai"]}");
            bcThang.Add("x11", $"{item["tong_luot_noi"]}");
            bcThang.Add("x21", $"{item["tong_chi"]}");
            bcThang.Add("x22", $"{item["tong_chi_ngoai"]}");
            bcThang.Add("x23", $"{item["tong_chi_noi"]}");
            /* ,x12 real not null default 0 /* Tổng lượt = 5+6 (x13+x14) Luỹ kế
                ,x13 real not null default 0 /* Lượt ngoại {nam1} luỹ kế
                ,x14 real not null default 0 /* Lượt nội {nam1} luỹ kế
            ,x24 real not null default 0 /* Tổng chi = 5+6 (x25+x26)
                ,x25 real not null default 0 /* Chi ngoại trú {nam1} luỹ kế
                ,x26 real not null default 0 /* Chi nội trú {nam1} luỹ kế */
            tmp = $"{dbConnect.getValue($"SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND tu_thang=1 AND nam='{bcThang["nam1"]}' LIMIT 1")}";
            item = dbConnect.getDataTable($"SELECT * FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2='{tmp}' AND ma_tinh='{maTinh}' LIMIT 1").AsEnumerable().FirstOrDefault();
            if (item == null) { throw new Exception($"[creatbcThang] Biểu 02 Toàn quốc từ tháng 1 đến {bcThang["thang"]} năm {bcThang["nam1"]} không có dữ liệu của tỉnh {maTinh}"); }
            bcThang.Add("x12", $"{item["tong_luot"]}");
            bcThang.Add("x13", $"{item["tong_luot_ngoai"]}");
            bcThang.Add("x14", $"{item["tong_luot_noi"]}");
            bcThang.Add("x24", $"{item["tong_chi"]}");
            bcThang.Add("x25", $"{item["tong_chi_ngoai"]}");
            bcThang.Add("x26", $"{item["tong_chi_noi"]}");

            /* ,x15 real not null default 0 /* Tổng lượt = 2+3 (x10+x11)
                ,x16 real not null default 0 /* Lượt ngoại {nam2}
                ,x17 real not null default 0 /* Lượt nội {nam2}
            ,x27 real not null default 0 /* Tổng chi = 2+3 (x22+x23)
                ,x28 real not null default 0 /* Chi ngoại trú {nam2}
                ,x29 real not null default 0 /* Chi nội trú {nam2}  */
            tmp = $"{dbConnect.getValue($"SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND tu_thang=den_thang AND nam='{bcThang["nam2"]}' LIMIT 1")}";
            item = dbConnect.getDataTable($"SELECT * FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2='{tmp}' AND ma_tinh='{maTinh}' LIMIT 1").AsEnumerable().FirstOrDefault();
            if (item == null) { throw new Exception($"[creatbcThang] Biểu 02 Toàn quốc tháng {bcThang["thang"]} năm {bcThang["nam2"]} không có dữ liệu của tỉnh {maTinh}"); }
            bcThang.Add("x15", $"{item["tong_luot"]}");
            bcThang.Add("x16", $"{item["tong_luot_ngoai"]}");
            bcThang.Add("x17", $"{item["tong_luot_noi"]}");
            bcThang.Add("x27", $"{item["tong_chi"]}");
            bcThang.Add("x28", $"{item["tong_chi_ngoai"]}");
            bcThang.Add("x29", $"{item["tong_chi_noi"]}");

            /* ,x18 real not null default 0 /* Tổng lượt = 5+6 (x13+x14) Luỹ kế
                ,x19 real not null default 0 /* Lượt ngoại {nam2} luỹ kế
                ,x20 real not null default 0 /* Lượt nội {nam2} luỹ kế
            ,x30 real not null default 0 /* Tổng chi = 5+6 (x25+x26)
                ,x31 real not null default 0 /* Chi ngoại trú {nam2} luỹ kế
                ,x32 real not null default 0 /* Chi nội trú {nam2} luỹ kế */
            tmp = $"{dbConnect.getValue($"SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND tu_thang=1 AND nam='{bcThang["nam2"]}' LIMIT 1")}";
            item = dbConnect.getDataTable($"SELECT * FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2='{tmp}' AND ma_tinh='{maTinh}' LIMIT 1").AsEnumerable().FirstOrDefault();
            if (item == null) { throw new Exception($"[creatbcThang] Biểu 02 Toàn quốc từ tháng 1 đến {bcThang["thang"]} năm {bcThang["nam2"]} không có dữ liệu của tỉnh {maTinh}"); }
            bcThang.Add("x18", $"{item["tong_luot"]}");
            bcThang.Add("x19", $"{item["tong_luot_ngoai"]}");
            bcThang.Add("x20", $"{item["tong_luot_noi"]}");
            bcThang.Add("x30", $"{item["tong_chi"]}");
            bcThang.Add("x31", $"{item["tong_chi_ngoai"]}");
            bcThang.Add("x32", $"{item["tong_chi_noi"]}");

            /* Tăng giảm so với cùng kỳ năm trước
             * ,m13lc13 real not null default 0 /* Tổng lượt = 2+3 (x15-x9)
                ,m13lc23 real not null default 0 /* Lượt ngoại = (x16-x10)
                ,m13lc33 real not null default 0 /* Lượt nội = (x17-x11)
                ,m13lc43 real not null default 0 /* Tổng lượt = 5+6 (x18-x12)
                ,m13lc53 real not null default 0 /* Lượt ngoại = (x19-x13)
                ,m13lc63 real not null default 0 /* Lượt nội = (x20-x14) */
            bcThang.Add("m13lc13", $"{(double.Parse(bcThang["x15"]) - double.Parse(bcThang["x9"]))}");
            bcThang.Add("m13lc23", $"{(double.Parse(bcThang["x16"]) - double.Parse(bcThang["x10"]))}");
            bcThang.Add("m13lc33", $"{(double.Parse(bcThang["x17"]) - double.Parse(bcThang["x11"]))}");
            bcThang.Add("m13lc43", $"{(double.Parse(bcThang["x18"]) - double.Parse(bcThang["x12"]))}");
            bcThang.Add("m13lc53", $"{(double.Parse(bcThang["x19"]) - double.Parse(bcThang["x13"]))}");
            bcThang.Add("m13lc63", $"{(double.Parse(bcThang["x20"]) - double.Parse(bcThang["x14"]))}");
            /* Tỷ lệ % tăng giảm
                ,m13lc14 real not null default 0 /* Tổng lượt = 2+3 ((m13lc13/x15)*100)
                ,m13lc24 real not null default 0 /* Lượt ngoại = (m13lc23/x16)*100/
                ,m13lc34 real not null default 0 /* Lượt nội = (m13lc33/x17)*100
                ,m13lc44 real not null default 0 /* Tổng lượt = 5+6 ((m13lc43/x18)*100)
                ,m13lc54 real not null default 0 /* Lượt ngoại = (m13lc53/x19)*100
                ,m13lc64 real not null default 0 /* Lượt nội = (m13lc63/x20)*100 */
            bcThang.Add("m13lc14", $"{Math.Round((double.Parse(bcThang["m13lc13"]) / double.Parse(bcThang["x15"])) * 100, 2)}");
            bcThang.Add("m13lc24", $"{Math.Round((double.Parse(bcThang["m13lc23"]) / double.Parse(bcThang["x16"])) * 100, 2)}");
            bcThang.Add("m13lc34", $"{Math.Round((double.Parse(bcThang["m13lc33"]) / double.Parse(bcThang["x17"])) * 100, 2)}");
            bcThang.Add("m13lc44", $"{Math.Round((double.Parse(bcThang["m13lc43"]) / double.Parse(bcThang["x18"])) * 100, 2)}");
            bcThang.Add("m13lc54", $"{Math.Round((double.Parse(bcThang["m13lc53"]) / double.Parse(bcThang["x19"])) * 100, 2)}");
            bcThang.Add("m13lc64", $"{Math.Round((double.Parse(bcThang["m13lc63"]) / double.Parse(bcThang["x20"])) * 100, 2)}");

            /* Tăng giảm so với cùng kỳ năm trước
             *  ,m13cc13 real not null default 0 /* Tổng lượt = 2+3 (x27-x21)
                ,m13cc23 real not null default 0 /* Chi ngoại trú = (x28-x22)
                ,m13cc33 real not null default 0 /* Chi nội trú = (x29-x23)
                ,m13cc43 real not null default 0 /* Tổng lượt = 5+6 (x30-x24)
                ,m13cc53 real not null default 0 /* Chi ngoại trú = (x31-x25)
                ,m13cc63 real not null default 0 /* Chi nội trú = (x32-x26) */
            bcThang.Add("m13cc13", $"{(double.Parse(bcThang["x27"]) - double.Parse(bcThang["x21"]))}");
            bcThang.Add("m13cc23", $"{(double.Parse(bcThang["x28"]) - double.Parse(bcThang["x22"]))}");
            bcThang.Add("m13cc33", $"{(double.Parse(bcThang["x29"]) - double.Parse(bcThang["x23"]))}");
            bcThang.Add("m13cc43", $"{(double.Parse(bcThang["x30"]) - double.Parse(bcThang["x24"]))}");
            bcThang.Add("m13cc53", $"{(double.Parse(bcThang["x31"]) - double.Parse(bcThang["x25"]))}");
            bcThang.Add("m13cc63", $"{(double.Parse(bcThang["x32"]) - double.Parse(bcThang["x26"]))}");

            /* Tỷ lệ % tăng giảm
                ,m13cc14 real not null default 0 /* Tổng lượt = 2+3 ((m13cc13/x27)*100)
                ,m13cc24 real not null default 0 /* Chi ngoại trú = (m13cc23/x28)*100
                ,m13cc34 real not null default 0 /* Chi nội trú = (m13cc33/x29)*100
                ,m13cc44 real not null default 0 /* Tổng lượt = 5+6 ((m13cc43/x30)*100)
                ,m13cc54 real not null default 0 /* Chi ngoại trú = (m13cc53/x31)*100
                ,m13cc64 real not null default 0 /* Chi nội trú = (m13cc63/x32)*100 */
            bcThang.Add("m13lc14", $"{Math.Round((double.Parse(bcThang["m13cc13"]) / double.Parse(bcThang["x27"])) * 100, 2)}");
            bcThang.Add("m13lc24", $"{Math.Round((double.Parse(bcThang["m13cc23"]) / double.Parse(bcThang["x28"])) * 100, 2)}");
            bcThang.Add("m13lc34", $"{Math.Round((double.Parse(bcThang["m13cc33"]) / double.Parse(bcThang["x29"])) * 100, 2)}");
            bcThang.Add("m13lc44", $"{Math.Round((double.Parse(bcThang["m13cc43"]) / double.Parse(bcThang["x30"])) * 100, 2)}");
            bcThang.Add("m13lc54", $"{Math.Round((double.Parse(bcThang["m13cc53"]) / double.Parse(bcThang["x31"])) * 100, 2)}");
            bcThang.Add("m13lc64", $"{Math.Round((double.Parse(bcThang["m13cc63"]) / double.Parse(bcThang["x32"])) * 100, 2)}");
            return bcThang;
        }

        private void createFilebcThangDocx(string idBaoCao, string idtinh, Dictionary<string, string> bcThang)
        {
            string pathFileTemplate = Path.Combine(AppHelper.pathAppData, "bcThang.docx");
            if (System.IO.File.Exists(pathFileTemplate) == false) { throw new Exception("Không tìm thấy tập tin mẫu báo cáo 'bcThang.docx' trong thư mục App_Data"); }

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