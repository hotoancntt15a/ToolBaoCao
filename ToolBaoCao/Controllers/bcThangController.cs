using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SharpCompress.Archives;
using SharpCompress.Archives.Rar;
using SharpCompress.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
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
            var tmp = ""; string nam = "";
            ViewBag.id = id;
            try
            {
                /* Xoá hết các File có trong thư mục */
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
                if (Request.Files.Count == 0) { throw new Exception("Không có tập tin nào được đẩy lên"); }
                var lsFile = new List<string>();
                var lsFileTarget = new List<string>();
                var list = new List<string>();
                if (Request.Files.Count == 1)
                {
                    var ext = Path.GetExtension(Request.Files[0].FileName).ToLower();
                    if (ext == ".xlsx")
                    {
                        list.Add($"{Request.Files[0].FileName} ({Request.Files[0].ContentLength.getFileSize()})");
                        /* Cập nhật dự toán được giao trong năm của csyt */
                        ViewBag.mode = "update";
                        string file = Path.Combine(folderTemp, $"{id}_pl01.xlsx");
                        Request.Files[0].SaveAs(file);
                        var dtgiaocsyt = zModules.NPOIExcel.XLSX.getDataFromExcel(new FileInfo(file));

                        /* Tìm năm trong 10 dòng đầu tiên */
                        int indexRow = -1;
                        for (int i = 0; i < (dtgiaocsyt.Rows.Count > 10 ? 10 : dtgiaocsyt.Rows.Count); i++)
                        {
                            indexRow++;
                            if (nam != "") { break; } /* Đã xác định được Năm */
                            if (Regex.IsMatch(dtgiaocsyt.Rows[0][0].ToString().Trim(), @"^[']?\d+$")) { break; } /* Đọc đến vùng dữ liệu */
                            for (int j = 0; j < dtgiaocsyt.Columns.Count; j++)
                            {
                                System.Text.RegularExpressions.Match match = Regex.Match(dtgiaocsyt.Rows[i][j].ToString(), @"\b(\d{4})\b");
                                if (match.Success) { nam = match.Value; break; }
                            }
                        }
                        var items = new DataTable("dtgiaocsyt");
                        items.Columns.Add("id");
                        items.Columns.Add("idtinh");
                        items.Columns.Add("nam");
                        items.Columns.Add("ma_cskcb");
                        items.Columns.Add("ten_cskcb");
                        items.Columns.Add("dutoangiao");
                        items.Columns.Add("userid");
                        items.Columns.Add("timeup");
                        var timeup = DateTime.Now.toTimestamp();
                        var listMaCSKCB = new List<string>();
                        var sotien = "";
                        indexRow++;
                        for (int i = indexRow; i < dtgiaocsyt.Rows.Count; i++)
                        {
                            tmp = dtgiaocsyt.Rows[0][0].ToString().Trim().Replace("'", "");
                            if (!Regex.IsMatch(tmp, @"^\d+$")) { break; }
                            sotien = dtgiaocsyt.Rows[i][2].ToString().Trim().Replace("'", "");
                            if (!Regex.IsMatch(sotien, @"^\d+(.\d+)?$")) { break; }
                            items.Rows.Add($"{matinh}.{nam}.{tmp}", matinh, nam, tmp, dtgiaocsyt.Rows[i][1].ToString().Trim(), sotien, idUser, timeup);
                            listMaCSKCB.Add(tmp);
                        }
                        if (items.Rows.Count == 0) { throw new Exception("Không có dữ liệu Dự Toán tạm giao CSYT"); }
                        var db = new dbSQLite(Path.Combine(AppHelper.pathAppData, $"BaoCaoThang{matinh}.db"));
                        db.CreateBcThang();
                        db.Insert("thangdtgiao", items, "replace");
                        ViewBag.Message = $"Đã cập nhật Dự toán tạm giao CSYT: {string.Join(",", listMaCSKCB)}";
                        return View();
                    }
                    else if (ext == ".zip")
                    {
                        /* Giải nén tập tin */
                        string fileName = Path.Combine(folderTemp, $"{id}.zip");
                        Request.Files[0].SaveAs(fileName);
                        using (ZipArchive archive = ZipFile.OpenRead(fileName))
                        {
                            int indexDBZip = 0;
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                indexDBZip++;
                                if (entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) == false) { continue; }
                                fileName = $"{id}_{indexDBZip}zip.xlsx";
                                entry.ExtractToFile(Path.Combine(folderTemp, fileName), overwrite: true);
                                lsFile.Add(fileName);
                                fileName = Path.GetFileName(entry.Name);
                                lsFileTarget.Add(fileName);
                                list.Add($"{fileName} ({entry.Length.getFileSize()})");
                            }
                        }
                    }
                    else if (ext == ".rar")
                    {
                        /* Giải nén tập tin */
                        string fileName = Path.Combine(folderTemp, $"{id}{ext}");
                        Request.Files[0].SaveAs(fileName);
                        using (var archive = RarArchive.Open(fileName))
                        {
                            int indexDBZip = 0;
                            foreach (var entry in archive.Entries)
                            {
                                if (entry.IsDirectory) { continue; }
                                if (entry.Key.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) == false) { continue; }
                                indexDBZip++;
                                fileName = $"{id}_{indexDBZip}zip.xlsx";
                                entry.WriteToFile(Path.Combine(folderTemp, fileName), new ExtractionOptions { ExtractFullPath = true, Overwrite = true });

                                lsFile.Add(fileName);
                                fileName = Path.GetFileName(entry.Key);
                                lsFileTarget.Add(fileName);
                                list.Add($"{fileName} ({entry.Size.getFileSize()})");
                            }
                        }
                    }
                }
                else
                {
                    string fileName = "";
                    for (int i = 0; i < Request.Files.Count; i++)
                    {
                        if (Path.GetExtension(Request.Files[i].FileName).ToLower() != ".xlsx") { continue; }
                        tmp = $"{id}_{i}.xlsx"; fileName = Request.Files[i].FileName;
                        list.Add($"{Request.Files[i].FileName} ({Request.Files[i].ContentLength.getFileSize()})");
                        Request.Files[i].SaveAs(Path.Combine(folderTemp, tmp));
                        lsFile.Add(tmp);
                        lsFileTarget.Add(fileName);
                    }
                }
                /* Khai báo dữ liệu tạm */
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                dbTemp.CreateImportBcThang();
                dbTemp.CreatePhucLucBcThang();
                dbTemp.CreateBcThang();
                /* Trường hợp cập nhật dữ liệu dự toán giao tại CSYT */
                /* Đọc và kiểm tra các tập tin */
                var bieus = new List<string>();
                for (int i = 0; i < lsFile.Count; i++)
                {
                    bieus.Add(readExcelbcThang(dbTemp, lsFile[i], Session, id, folderTemp, timeStart, lsFileTarget[i]));
                }
                ViewBag.files = list;
                list = new List<string>();
                bieus = bieus.Distinct().ToList();
                if (bieus.Where(p => p.StartsWith("b01")).Count() == 0) { throw new Exception($"Thiếu biểu đầu vào B01. {string.Join(", ", bieus)}"); }
                if (bieus.Where(p => p.StartsWith("b02")).Count() == 0) { throw new Exception($"Thiếu biểu đầu vào B02. {string.Join(", ", bieus)}"); }
                if (bieus.Where(p => p.StartsWith("b04")).Count() == 0) { throw new Exception($"Thiếu biểu đầu vào B04. {string.Join(", ", bieus)}"); }
                if (list.Count > 0) { throw new Exception(string.Join("<br />", list)); }
                /* Lấy năm, tháng báo cáo */
                nam = ""; string thang = "";
                var data = dbTemp.getDataTable($"SELECT den_thang, nam FROM thangb02 WHERE id_bc='{id}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
                if (nam == "")
                {
                    data = dbTemp.getDataTable($"SELECT den_thang, nam FROM thangb01 WHERE id_bc='{id}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                    if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
                }
                if (nam == "")
                {
                    data = dbTemp.getDataTable($"SELECT den_thang, nam FROM thangb04 WHERE id_bc='{id}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                    if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
                }
                if (nam == "") { throw new Exception("Không xác định được Năm, Tháng báo cáo"); }

                /* Tạo Phục Lục 1 - Lấy từ nguồn cơ sở luỹ kế - Chỉ lấy mã cấp trên */
                var tsql = $@"INSERT INTO thangpl01 (id_bc, idtinh, ma_cskcb, ten_cskcb, dtgiao, tien_bhtt, tl_sudungdt, userid)
                    SELECT '{id}' AS id_bc, '{matinh}' AS idtinh, ma_cskcb, ten_cskcb, ROUND(dtcsyt_trongnam, 0) AS dtgiao, ROUND(dtcsyt_chikcb, 0) AS t_bhtt, 0 AS tl_sudungdt, '{idUser}' AS userid
                    FROM thangb01chitiet WHERE id_bc='{id}' AND id2 IN (SELECT id FROM thangb01 WHERE id_bc='{id}' AND ma_tinh='{matinh}' AND nam={nam} AND ma_cskcb <> '' ORDER BY den_thang DESC LIMIT 1);";
                dbTemp.Execute(tsql);
                /* Tạo Phục Lục 2a */
                /* Lấy dữ liệu từ biểu pl02a trong tháng (Từ tháng đến tháng = tháng báo cáo của toàn quốc nam1) */
                dbTemp.Execute($@"INSERT INTO thangpl02a (id_bc, idtinh ,ma_tinh ,ten_tinh ,ma_vung
                ,tyle_noitru ,ngay_dtri_bq ,chi_bq_chung ,chi_bq_ngoai ,chi_bq_noi
                ,tong_luot, tong_luot_noi, tong_luot_ngoai
                ,tong_chi, tong_chi_noi, tong_chi_ngoai
                ,userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    ,ROUND(tyle_noitru, 2) AS tyle_noitru ,ROUND(ngay_dtri_bq, 2) AS ngay_dtri_bq ,ROUND(chi_bq_chung) AS chi_bq_chung ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai ,ROUND(chi_bq_noi) AS chi_bq_noi
                    ,tong_luot, tong_luot_noi, tong_luot_ngoai
                    ,tong_chi, tong_chi_noi, tong_chi_ngoai
                    ,'{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='00' AND nam={nam} AND tu_thang=den_thang AND tu_thang={thang} LIMIT 1);");
                /* Tạo Phục Lục 2b */
                /* Lấy dữ liệu từ biểu b02 dành cho cả năm (từ tháng 1 đến tháng báo cáo) */
                dbTemp.Execute($@"INSERT INTO thangpl02b (id_bc, idtinh ,ma_tinh ,ten_tinh ,ma_vung
                ,tyle_noitru ,ngay_dtri_bq ,chi_bq_chung ,chi_bq_ngoai ,chi_bq_noi
                ,tong_luot, tong_luot_noi, tong_luot_ngoai
                ,tong_chi, tong_chi_noi, tong_chi_ngoai
                ,userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    ,ROUND(tyle_noitru, 2) AS tyle_noitru ,ROUND(ngay_dtri_bq, 2) AS ngay_dtri_bq ,ROUND(chi_bq_chung) AS chi_bq_chung ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai ,ROUND(chi_bq_noi) AS chi_bq_noi
                    ,tong_luot, tong_luot_noi, tong_luot_ngoai
                    ,tong_chi, tong_chi_noi, tong_chi_ngoai
                    ,'{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='00' AND nam={nam} AND tu_thang=1 AND den_thang={thang} LIMIT 1);");
                /* Tạo Phục Lục 3a */
                /* Lấy dữ liệu từ biểu b02 csyt trong tháng */
                tsql = $@"SELECT p1.id_bc, '{matinh}' as idtinh, p1.ma_cskcb, p1.ten_cskcb, p1.ma_vung
                    ,ROUND(p1.tyle_noitru, 2) AS tyle_noitru ,ROUND(p1.ngay_dtri_bq, 2) AS ngay_dtri_bq
                    ,ROUND(p1.chi_bq_chung) AS chi_bq_chung ,ROUND(p1.chi_bq_ngoai) AS chi_bq_ngoai
                    ,ROUND(p1.chi_bq_noi) AS chi_bq_noi, p2.den_thang as thang, '' as tuyen_bv, '' as hang_bv,'{idUser}' AS userid
                    FROM thangb02chitiet p1 INNER JOIN thangb02 p2 ON p1.id2=p2.id WHERE p1.id_bc='{id}' AND p2.id_bc='{id}' AND p1.ma_cskcb <> '' AND p2.tu_thang=p2.den_thang";
                if (thang == "1") { tsql += $" AND p2.nam IN ({nam}, {(int.Parse(nam) - 1)}) AND p2.tu_thang IN (1, 12)"; }
                else { tsql += $" AND p2.nam = {nam} AND p2.tu_thang IN ({thang}, {(int.Parse(thang) - 1)})"; }
                data = dbTemp.getDataTable(tsql);
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
                /* Lấy dữ liệu từ biểu pl02a trong tháng (Từ tháng đến tháng = tháng báo cáo của toàn quốc nam2) */
                dbTemp.Execute($@"INSERT INTO thangpl03a2 (id_bc, idtinh ,ma_tinh ,ten_tinh ,ma_vung
                ,tyle_noitru ,ngay_dtri_bq ,chi_bq_chung ,chi_bq_ngoai ,chi_bq_noi
                ,tong_luot, tong_luot_noi, tong_luot_ngoai
                ,tong_chi, tong_chi_noi, tong_chi_ngoai
                ,userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    ,ROUND(tyle_noitru, 2) AS tyle_noitru ,ROUND(ngay_dtri_bq, 2) AS ngay_dtri_bq ,ROUND(chi_bq_chung) AS chi_bq_chung ,ROUND(chi_bq_ngoai) AS chi_bq_ngoai ,ROUND(chi_bq_noi) AS chi_bq_noi
                    ,tong_luot, tong_luot_noi, tong_luot_ngoai
                    ,tong_chi, tong_chi_noi, tong_chi_ngoai
                    ,'{idUser}' AS userid
                    FROM thangb02chitiet WHERE id_bc='{id}' AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{id}' AND ma_tinh='00' AND nam={(int.Parse(nam) - 1)} AND tu_thang=den_thang AND tu_thang={thang} LIMIT 1);");

                /* Tạo phục lục 03b */
                /* Cách lập giống như Phụ lục 03 báo cáo tuần, nguồn dữ liệu lấy từ B02 từ tháng 1 đến tháng báo cáo */
                data = dbTemp.getDataTable($@"SELECT p1.id_bc, '{matinh}' as idtinh, p1.ma_cskcb, p1.ten_cskcb, p1.ma_vung
                    ,ROUND(p1.tyle_noitru, 2) AS tyle_noitru ,ROUND(p1.ngay_dtri_bq, 2) AS ngay_dtri_bq
                    ,ROUND(p1.chi_bq_chung) AS chi_bq_chung ,ROUND(p1.chi_bq_ngoai) AS chi_bq_ngoai
                    ,ROUND(p1.chi_bq_noi) AS chi_bq_noi, p2.nam, '' as tuyen_bv, '' as hang_bv,'{idUser}' AS userid
                    FROM thangb02chitiet p1 INNER JOIN thangb02 p2 ON p1.id2=p2.id WHERE p1.id_bc='{id}' AND p2.id_bc='{id}' AND p1.ma_cskcb <> '' AND nam IN ({nam}, {(int.Parse(nam) - 1)}) AND p2.tu_thang=1 AND p2.den_thang={thang};");
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
                dbTemp.Execute($@"INSERT INTO thangpl04a (id_bc, idtinh, ma_tinh, ten_tinh, ma_vung, chi_bq_xn, chi_bq_cdha, chi_bq_thuoc, chi_bq_pttt, chi_bq_vtyt, chi_bq_giuong, ngay_ttbq, userid)
                    SELECT id_bc, '{matinh}' as idtinh, ma_tinh, ten_tinh, ma_vung
                    ,ROUND(bq_xn) AS chi_bq_xn ,ROUND(bq_cdha) AS chi_bq_cdha ,ROUND(bq_thuoc) AS chi_bq_thuoc ,ROUND(bq_ptt) AS chi_bq_pttt ,ROUND(bq_vtyt) AS chi_bq_vtyt ,ROUND(bq_giuong) AS chi_bq_giuong ,ROUND(ngay_ttbq, 2) AS ngay_ttbq
                    ,'{idUser}' AS userid
                    FROM thangb04chitiet WHERE id_bc='{id}' AND id2 IN (SELECT id FROM thangb04 WHERE id_bc='{id}' AND ma_tinh='00' AND nam={nam} AND tu_thang=1 AND den_thang={thang} LIMIT 1);");
                /* Tạo thangpl04b */
                /* Nguồn dữ liệu B04_10 của tháng báo cáo. Giống như Phụ lục 2 của báo cáo tuần, nhưng chi tiết từng CSKCB và phân nhóm theo tuyến tỉnh huyện xã */
                tsql = $@"SELECT p1.id_bc, '{matinh}' as idtinh, p1.ma_cskcb, p1.ten_cskcb, p1.ma_vung
                    , ROUND(p1.bq_xn) AS chi_bq_xn, ROUND(p1.bq_cdha) AS chi_bq_cdha, ROUND(p1.bq_thuoc) AS chi_bq_thuoc
                    , ROUND(p1.bq_ptt) AS chi_bq_pttt, ROUND(p1.bq_vtyt) AS chi_bq_vtyt, ROUND(p1.bq_giuong) AS chi_bq_giuong, ROUND(p1.ngay_ttbq, 2) AS ngay_ttbq
                    ,'' as tuyen_bv, '' as hang_bv, '{idUser}' AS userid, p2.den_thang as thang
                    FROM thangb04chitiet p1 INNER JOIN thangb04 p2 ON p1.id2=p2.id
                    WHERE p1.id_bc='{id}' AND p2.id_bc='{id}' AND p2.ma_tinh='{matinh}' AND p1.ma_cskcb <> ''";
                if (thang == "1") { tsql += $" AND p2.nam IN ({nam}, {(int.Parse(nam) - 1)}) AND p2.tu_thang = p2.den_thang AND p2.den_thang IN (1, 12);"; }
                else { tsql += $" AND p2.nam = {nam} AND p2.tu_thang = p2.den_thang AND p2.den_thang IN ({thang}, {(int.Parse(thang) - 1)});"; }
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
                if (thang == "1") { dbTemp.Execute($"UPDATE thangpl04b thang = 0 WHERE id_bc='{id}' AND thang=12"); }
                dbTemp.Close();
                /* Tìm các báo cáo trước để cập nhật x1, x33-x38 */
                var rsData = new Dictionary<string, string>() { { "x1", "" }, { "x33", "" }, { "x34", "" }, { "x35", "" }, { "x36", "" }, { "x37", "" }, { "x38", "" } };
                var dbBCThang = BuildDatabase.getDataBCThang(matinh);
                data = dbBCThang.getDataTable("SELECT x1, x33, x34, x35, x36, x37, x38 FROM bcthangdocx ORDER BY nam1 DESC, thang DESC LIMIT 5");
                if (data.Rows.Count > 0)
                {
                    foreach (DataRow row in data.Rows)
                    {
                        tmp = ""; for (int i = 0; i < data.Columns.Count; i++) { rsData[data.Columns[i].ColumnName] = $"{row[i]}"; tmp += $"{row[i]}"; }
                        if (tmp != "") { break; }
                    }
                }
                ViewBag.rsdata = rsData;
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.getLineHTML();
                var d = new System.IO.DirectoryInfo(folderTemp);
                foreach (var item in d.GetFiles()) { try { item.Delete(); } catch { } }
            }
            return View();
        }

        private void exportPhuLucbcThang(string idBC, string outFile, params DataTable[] par)
        {
            if (System.IO.File.Exists(outFile)) { System.IO.File.Delete(outFile); }
            string fileName = Path.Combine(AppHelper.pathTemp, "bcThang", $"bcthang{idBC}.xlsx");
            System.IO.File.Copy(Path.Combine(AppHelper.pathAppData, "bcThangPL.xlsx"), fileName);
            try
            {
                XSSFWorkbook workbook = new XSSFWorkbook(fileName);
                int i = 0; int rowIndex = 0;
                var names = new List<string>();
                string tmp = ""; bool isCreateSheet = false;
                var csContext = workbook.CreateCellStyleThin(getCache: false);
                var csContextR = workbook.CreateCellStyleThin(getCache: false, alignment: HorizontalAlignment.Right);
                var csContextB = workbook.CreateCellStyleThin(true, getCache: false);
                var csContextBR = workbook.CreateCellStyleThin(true, getCache: false, alignment: HorizontalAlignment.Right);
                foreach (DataTable dt in par)
                {
                    isCreateSheet = false;
                    ISheet sheet = workbook.GetSheet(dt.TableName);
                    if (sheet == null)
                    {
                        sheet = workbook.CreateSheet(dt.TableName);
                        isCreateSheet = true;
                    }
                    names.Add(dt.TableName);
                    var listColRight = new List<int>();
                    var listColWith = new List<int>();
                    switch (dt.TableName)
                    {
                        case "PL01":
                            listColRight = new List<int>() { 2, 3, 4 };
                            listColWith = new List<int>() { 11, 65, 25, 25, 13 };
                            break;

                        case "PL02a":
                            listColRight = new List<int>() { 2, 4, 6, 8, 10 };
                            listColWith = new List<int>() { 9, 18, 14, 14, 14, 14, 14, 14, 14, 14, 14 };
                            break;

                        case "PL02b":
                            listColRight = new List<int>() { 2, 4, 6, 8, 10 };
                            listColWith = new List<int>() { 9, 18, 14, 14, 14, 14, 14, 14, 14, 14, 14 };
                            break;

                        case "PL02c":
                            listColRight = new List<int>() { 3, 4, 5, 6, 7, 8 };
                            listColWith = new List<int>() { 11, 11, 36, 16, 16, 16, 16, 16, 16 };
                            break;

                        case "PL03a":
                            listColRight = new List<int>() { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18 };
                            listColWith = new List<int>() { 9, 9, 57, 13, 13, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14 };
                            break;

                        case "PL03b":
                            listColRight = new List<int>() { 3, 4, 5, 6, 7 };
                            listColWith = new List<int>() { 9, 9, 57, 13, 13, 14, 14, 14 };
                            break;

                        case "PL03c":
                            listColRight = new List<int>() { 3, 4, 5, 6, 7, 8 };
                            listColWith = new List<int>() { 11, 11, 36, 16, 16, 16, 16, 16, 16 };
                            break;

                        case "PL04a":
                            listColRight = new List<int>() { 2, 3, 4, 5, 6, 7, 8 };
                            listColWith = new List<int>() { 9, 18, 14, 14, 14, 14, 14, 14, 14 };
                            break;

                        case "PL04b":
                            listColRight = new List<int>() { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 };
                            listColWith = new List<int>() { 9, 9, 57, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14 };
                            break;

                        default: break;
                    }
                    for (int colIndex = 0; colIndex < listColWith.Count; colIndex++) { sheet.SetColumnWidth(colIndex, (listColWith[colIndex] * 256)); }
                    if (isCreateSheet)
                    {
                        /* Tạo tiêu đề */
                        rowIndex = 0;
                        var row = sheet.CreateRow(rowIndex);
                        i = -1;
                        var csHeader = workbook.CreateCellStyleThin(true, true, true, getCache: false);
                        foreach (DataColumn col in dt.Columns)
                        {
                            i++;
                            var cell = row.CreateCell(i, CellType.String);
                            cell.CellStyle = csHeader;
                            cell.SetCellValue(Regex.Replace(col.ColumnName, @"[ ][(]\d+[)]", ""));
                        }
                    }
                    else
                    {
                        /* Tìm việc trí bắt đầu đổ dữ liệu */
                        for (rowIndex = 0; rowIndex <= 10; rowIndex++)
                        {
                            var row = sheet.GetRow(rowIndex); if (row == null) { continue; }
                            var cell = row.GetCell(0); if (cell == null) { continue; }
                            tmp = $"{cell.GetValueAsString()}".Trim();
                            if (tmp == "{filldata}") { break; }
                        }
                    }
                    /* Đổ dữ liệu */
                    int indexColumn = -1; rowIndex--;
                    foreach (DataRow r in dt.Rows)
                    {
                        rowIndex++;
                        var row = sheet.CreateRow(rowIndex);
                        indexColumn = -1;
                        if ($"{r[0]}{r[1]}" == "")
                        {
                            foreach (DataColumn col in dt.Columns)
                            {
                                indexColumn++;
                                var cell = row.CreateCell(indexColumn, CellType.String);
                                cell.CellStyle = listColRight.Contains(indexColumn) ? csContextR : csContext;
                                cell.SetCellValue("");
                            }
                        }
                        else
                        {
                            foreach (DataColumn col in dt.Columns)
                            {
                                indexColumn++;
                                var cell = row.CreateCell(indexColumn, CellType.String);
                                tmp = $"{r[indexColumn]}";
                                if (tmp.StartsWith("<b>"))
                                {
                                    tmp = tmp.Substring(3);
                                    if (listColRight.Contains(indexColumn)) { cell.CellStyle = csContextBR; cell.SetCellValue(tmp.FormatCultureVN()); }
                                    else { cell.CellStyle = csContextB; cell.SetCellValue(tmp); }
                                }
                                else
                                {
                                    if (listColRight.Contains(indexColumn)) { cell.CellStyle = csContextR; cell.SetCellValue(tmp.FormatCultureVN()); }
                                    else { cell.CellStyle = csContext; cell.SetCellValue(tmp); }
                                }
                            }
                        }
                    }
                }
                using (FileStream stream = new FileStream(outFile, FileMode.Create, FileAccess.Write)) { workbook.Write(stream); }
                workbook.Close(); workbook.Clear();
            }
            catch (Exception ex) { throw new Exception(ex.getLineHTML()); }
            finally { try { System.IO.File.Delete(fileName); } catch { } }
            return;
        }

        private string readExcelbcThang(dbSQLite dbConnect, string inputFile, HttpSessionStateBase Session, string idBaoCao, string folderTemp, DateTime timeStart, string fileName)
        {
            string messageError = "";
            var timeUp = timeStart.toTimestamp().ToString();
            var userID = $"{Session["iduser"]}".Trim();
            var matinh = $"{Session["idtinh"]}".Trim();
            var listBieu = new List<string>();
            string bieu = "";
            string fileExtension = Path.GetExtension(inputFile);
            int sheetIndex = 0; int packetSize = 1000;
            int indexRow = 0; int indexColumn = 0; int maxRow = 0; int jIndex = 0;
            int fieldCount = 50; var tsql = new List<string>();
            var tmp = "";
            IWorkbook workbook = null;
            try
            {
                try { workbook = new XSSFWorkbook(Path.Combine(folderTemp, inputFile)); }
                catch (Exception ex) { throw new Exception($"Lỗi tập tin '{fileName}' sai định dạng : {ex.Message}"); }
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
                        if (tmp.StartsWith("b26")) { bieu = "b26"; /* 2 b26: b26_nam1 b26cs_thang1 */ }
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
                 * Biểu b26: ma_tinh	loai_kcb	thoi_gian	loai_bv	kieubv	loaick	hang_bv	tuyen	loai_so_sanh    cs
                 */
                switch (bieu)
                {
                    /* Kiểm tra năm */
                    case "b01": fieldCount = 5; indexRegex = 3; pattern = "^20[0-9][0-9]$"; break;
                    /* Kiểm tra năm */
                    case "b02": fieldCount = 11; indexRegex = 4; pattern = "^20[0-9][0-9]$"; break;
                    /* Kiểm tra thoigian */
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

                    case "b26":
                        /* Kiểm tra BQ chung trong kỳ */
                        /* Biểu b26: ma_tinh	loai_kcb	thoi_gian	loai_bv	kieubv	loaick	hang_bv	tuyen	loai_so_sanh     cs */
                        idChiTiet = $"{listValue[0]}-{listValue[2]}";
                        listBieu.Add($"b26{idChiTiet}");
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
                        fieldCount = 20; indexRegex = 3 + 1; pattern = @"^\d+([.]\d+)?$"; /* nguồn trong năm */
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
                    /* Kiểm tra BQ chung trong kỳ */
                    case "thangb26":
                        fieldCount = 34; indexRegex = 7 + 1; pattern = "^[0-9]+[.,][0-9]+$|^[0-9]+$";
                        fieldNumbers = new List<int>() { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33 };
                        break;

                    default: fieldCount = 11; break;
                }
                /* Bỏ qua dòng tiêu đề */
                indexRow++; int recordCount = 0;
                var tsqlVaues = new List<string>();
                for (; indexRow <= maxRow; indexRow++)
                {
                    if (tsqlVaues.Count > packetSize)
                    {
                        recordCount += tsqlVaues.Count;
                        tsql.Add($"INSERT INTO {bieu}chitiet ({string.Join(",", allColumns)}) VALUES {string.Join(",", tsqlVaues)};");
                        tsqlVaues = new List<string>();
                    }
                    /* Dòng không tồn tại */
                    row = sheet.GetRow(indexRow); if (row == null) { continue; }
                    /* Số cột ít hơn số trường cần lấy dữ liệu */
                    /* if ((int)row.LastCellNum < fieldCount) { continue; } */
                    /* Cột đầu tiên không phải là matinh dạng số */
                    string ma = row.GetCell(indexColumn).GetValueAsString().Trim();

                    if (Regex.IsMatch(ma, @"^([A-Z]+)?\d+([A-Z]+)?$") == false) { continue; }
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
                if (tsqlVaues.Count > 0)
                {
                    recordCount += tsqlVaues.Count;
                    tsql.Add($"INSERT INTO {bieu}chitiet ({string.Join(",", allColumns)}) VALUES {string.Join(",", tsqlVaues)};");
                    tsqlVaues = new List<string>();
                }
                tmp = string.Join(Environment.NewLine, tsql);
                /* System.IO.File.WriteAllText(Path.Combine(folderTemp, $"id{idBaoCao}_{listBieu[0]}_{matinhImport}.sql"), tmp); */
                dbConnect.Execute(tmp);
                if (tsql.Count < 2) { throw new Exception("Không có dữ liệu chi tiết"); }
                /* Lưu lại file */
                using (FileStream stream = new FileStream(Path.Combine(folderTemp, $"id{idBaoCao}_{listBieu[0]}{fileExtension}"), FileMode.Create, FileAccess.Write)) { workbook.Write(stream); }
            }
            catch (Exception ex2)
            {
                messageError = $"Lỗi trong quá trình đọc, nhập dữ liệu từ Excel '{fileName}': {ex2.getLineHTML()}";
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
                string fileZip = Path.Combine(d.FullName, $"bcThang_{id}.zip");
                if (System.IO.File.Exists(fileZip) == false)
                {
                    /* Tạo lại báo cáo */
                    var dbBaoCao = BuildDatabase.getDataBCThang(matinh);
                    tsql = $"SELECT * FROM bcThangdocx WHERE id='{id.sqliteGetValueField()}'";
                    var data = dbBaoCao.getDataTable(tsql);
                    dbBaoCao.Close();
                    if (data.Rows.Count == 0)
                    {
                        ViewBag.Error = $"Báo cáo tuần có ID '{id}' thuộc tỉnh có mã '{matinh}' không tồn tại hoặc đã bị xoá khỏi hệ thống";
                        return View();
                    }
                    var listFile = createFileBCThang(id, matinh, dbBaoCao);
                    dbBaoCao.Close();
                    AppHelper.zipAchive(fileZip, listFile);
                }
                if (System.IO.File.Exists(Path.Combine(d.FullName, $"bcThang_{id}.docx")) == false) { AppHelper.zipExtract(fileZip, d.FullName, ".docx"); }
                if (System.IO.File.Exists(Path.Combine(d.FullName, $"bcThang_{id}_pl.xlsx")) == false) { AppHelper.zipExtract(fileZip, d.FullName, ".xlsx"); }
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
                var data = new DataTable();
                var pathDB = Path.Combine(folderTemp, "import.db");
                if (System.IO.File.Exists(pathDB) == false) { throw new Exception($"Dữ liệu tạo báo cáo tháng có ID '{idBaoCao}' đã bị huỷ hoặc không tồn tại trên hệ thống"); }
                var dbTemp = new dbSQLite(Path.Combine(folderTemp, "import.db"));
                /* Tạo bcThang */
                var bcThang = createbcThang(dbTemp, idBaoCao, idtinh, iduser, Request.getValue("x1"), Request.getValue("x33"), Request.getValue("x34"), Request.getValue("x35"), Request.getValue("x36"), Request.getValue("x37"), Request.getValue("x38"));
                /* Tạo dữ liệu để xuất phụ lục */
                string idBaoCaoVauleField = idBaoCao.sqliteGetValueField();
                var dbBcThang = BuildDatabase.getDataBCThang(idtinh);
                var dbImport = BuildDatabase.getDataImportBCThang(idtinh);
                /* Tạo phụ lục báo cáo */
                /* dmCSKCB */
                var dmCSKCB = AppHelper.dbSqliteMain.getDataTable($"SELECT id, ten, macaptren FROM dmcskcb WHERE ma_tinh='{idtinh}'").AsEnumerable();
                /* Di chuyển tập tin Excel */
                foreach (var f in dirTemp.GetFiles($"id{idBaoCao}*.xls*")) { f.MoveTo(Path.Combine(folderSave, f.Name)); }

                /* Báo cáo tháng chuyển */
                data = dbTemp.getDataTable($"SELECT * FROM bcthangdocx WHERE id='{idBaoCao}'");
                dbBcThang.Insert("bcthangdocx", data, "repalce");
                data = dbTemp.getDataTable($"SELECT * FROM bcthangpldocx WHERE id='{idBaoCao}'");
                dbBcThang.Insert("bcthangpldocx", data, "repalce");
                dbBcThang.Close();

                list = new List<string>() { "thangpl01", "thangpl02a", "thangpl02b", "thangpl03a", "thangpl03a2", "thangpl03b", "thangpl04a", "thangpl04b" };
                foreach (var v in list)
                {
                    data = dbTemp.getDataTable($"SELECT * FROM {v} WHERE id_bc='{idBaoCaoVauleField}';");
                    data.Columns.RemoveAt(0);
                    dbBcThang.Insert(v, data);
                }
                list = new List<string>() { "thangb01", "thangb02", "thangb04", "thangb26" };
                foreach (var v in list)
                {
                    data = dbTemp.getDataTable($"SELECT * FROM {v} WHERE id_bc='{idBaoCaoVauleField}';");
                    dbImport.Insert(v, data);
                }
                list = new List<string>() { "thangb01chitiet", "thangb02chitiet", "thangb04chitiet", "thangb26chitiet" };
                foreach (var v in list)
                {
                    data = dbTemp.getDataTable($"SELECT * FROM {v} WHERE id_bc='{idBaoCaoVauleField}';");
                    data.Columns.RemoveAt(0);
                    dbImport.Insert(v, data);
                }
                /* Tạo docx */
                var listFile = createFileBCThang(idBaoCao, idtinh, dbTemp);

                dbTemp.Close();
                AppHelper.zipAchive(Path.Combine(AppHelper.pathAppData, "bcThang", $"tinh{idtinh}", $"bcThang_{idBaoCao}.zip"), listFile);
                /* Xoá tập tin ở thư mục tạm đi */
                foreach (var f in dirTemp.GetFiles($"*{idBaoCao}*.*")) { try { f.Delete(); } catch { } }
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.getErrorSave();
                DeleteBcThang(idtinh);
            }
            return View();
        }

        private DataTable sortDataTable(DataTable dt, string field1, string field2)
        {
            DataTable sortedTable = dt.Clone();
            var groups = dt.AsEnumerable().GroupBy(row => row.Field<string>(field1));
            foreach (var group in groups)
            {
                var matchingRow = group.FirstOrDefault(row => row.Field<string>(field2) == row.Field<string>(field1));
                if (matchingRow != null)
                {
                    sortedTable.ImportRow(matchingRow);
                }
                foreach (var row in group.Where(row => row.Field<string>(field2) != row.Field<string>(field1)))
                {
                    sortedTable.ImportRow(row);
                }
            }
            return sortedTable;
        }

        private DataTable addMaCapTren(DataTable dt, Dictionary<string, string> matchMaCapTren, int ColMaCSKCBIndex = 0)
        {
            DataTable rs = new DataTable(dt.TableName);
            rs.Columns.Add("MaCapTren");
            foreach (DataColumn c in dt.Columns) { rs.Columns.Add(c.ColumnName, c.DataType); }
            foreach (DataRow r in dt.Rows)
            {
                var objR = new List<object>() { matchMaCapTren.getValue($"{r[ColMaCSKCBIndex]}", $"{r[ColMaCSKCBIndex]}") };
                for (int i = 0; dt.Columns.Count > i; i++) { objR.Add(r[i]); }
                rs.Rows.Add(objR.ToArray());
            }
            /*
            string tmp = "";
            foreach (DataRow r in dt.Rows)
            {
                if (tmp == $"{r[0]}") { r[0] = "-"; }
                else { tmp = $"{r[0]}"; }
            }
            */
            return rs;
        }

        private DataTable createPL02c(dbSQLite db, string idBaoCao, string idTinh, long namBC, long tuThang, long denThang, Dictionary<string, string> matchMaCapTren)
        {
            /* So sánh lượt KCB và chi KCB năm nay với năm trước */
            /* Cột A- B02	Cột B-B02	 Cột D-B02-10-2024; từ tháng 1 đến tháng báo cáo	  Cột D-B02-10-2023; từ tháng 1 đến tháng báo cáo	năm trước - năm nay	 Cột R-B02-10-2024; từ tháng 1 đến tháng báo cáo	 Cột R-B02-10-2023; từ tháng 1 đến tháng báo cáo	năm trước- năm nay */
            var pl = db.getDataTable($"SELECT '' as macaptren, ma_cskcb, ten_cskcb, tong_luot as luot1, 0 as luot2, 0 as luot3, t_bhtt as chi1, t_bhtt as chi2, t_bhtt as chi3 FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND nam={namBC} AND ma_tinh='{idTinh}' AND tu_thang={tuThang} AND den_thang={denThang});");
            var dicCSKCB = new Dictionary<string, int>();
            for (int i = 0; i < pl.Rows.Count; i++)
            {
                dicCSKCB.Add($"{pl.Rows[i]["ma_cskcb"]}", i);
                pl.Rows[i]["chi2"] = 0;
                pl.Rows[i]["chi3"] = 0;
            }
            var dt = db.getDataTable($"SELECT '' as macaptren, ma_cskcb, ten_cskcb, tong_luot, t_bhtt FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND nam={(namBC - 1)} AND ma_tinh='{idTinh}' AND tu_thang={tuThang} AND den_thang={denThang});");
            string tmp = ""; int index = 0;
            long luot1 = 0, luot2 = 0;
            double chi1 = 0, chi2 = 0;
            foreach (DataRow dr in dt.Rows)
            {
                tmp = $"{dr["ma_cskcb"]}";
                if (dicCSKCB.Keys.Contains(tmp))
                {
                    /* Tính toán */
                    index = dicCSKCB[tmp];
                    dicCSKCB.Remove(tmp);
                    /* Lượt */
                    luot1 = (long)pl.Rows[index]["luot1"];
                    luot2 = (long)dr["tong_luot"];
                    pl.Rows[index]["luot2"] = luot2;
                    pl.Rows[index]["luot3"] = luot1 - luot2;
                    /* Chi */
                    chi1 = (double)pl.Rows[index]["chi1"];
                    chi2 = (double)dr["t_bhtt"];
                    pl.Rows[index]["chi2"] = chi2;
                }
                else
                {
                    /* Thêm vào phục lục */
                    luot2 = (long)dr["tong_luot"];
                    chi2 = (double)dr["t_bhtt"];
                    pl.Rows.Add("", dr["ma_cskcb"], dr["ten_cskcb"], long.Parse("0"), luot2, (0 - luot2), double.Parse("0"), chi2, 0);
                }
            }
            pl.TableName = "PL02c";
            /* Làm tròn triệu đồng */
            foreach (DataRow dr in pl.Rows)
            {
                dr[6] = double.Parse($"{dr[6]}").lamTronTrieuDong(true);
                dr[7] = double.Parse($"{dr[7]}").lamTronTrieuDong(true);
                dr[8] = double.Parse($"{dr[6]}") - double.Parse($"{dr[7]}");
                dr["macaptren"] = matchMaCapTren.getValue($"{dr[1]}", $"{dr[1]}");
            }
            pl = sortDataTable(pl, "macaptren", "ma_cskcb");
            foreach (DataRow dr in pl.Rows)
            {
                if (tmp == $"{dr[0]}") { dr[0] = "-"; } else { tmp = $"{dr[0]}"; }
            }
            return pl;
        }

        private DataTable createPL03c(dbSQLite db, string idBaoCao, string idTinh, string thang, Dictionary<string, string> matchMaCapTren)
        {
            /* So sánh lượt KCB và chi KCB năm nay với năm trước */
            /* Cột A- B02	Cột B-B02	 Cột D-B02-10-2024; từ tháng 1 đến tháng báo cáo	  Cột D-B02-10-2023; từ tháng 1 đến tháng báo cáo	năm trước - năm nay	 Cột R-B02-10-2024; từ tháng 1 đến tháng báo cáo	 Cột R-B02-10-2023; từ tháng 1 đến tháng báo cáo	năm trước- năm nay */
            var pl = db.getDataTable($"SELECT '' as macaptren, ma_cskcb, ten_cskcb, tong_luot as luot1, 0 as luot2, 0 as luot3, tong_chi as chi1, tong_chi as chi2, tong_chi as chi3 FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND ma_tinh='{idTinh}' AND tu_thang=den_thang AND den_thang={thang});");
            var dicCSKCB = new Dictionary<string, int>();
            for (int i = 0; i < pl.Rows.Count; i++)
            {
                dicCSKCB.Add($"{pl.Rows[i]["ma_cskcb"]}", i);
                pl.Rows[i]["chi2"] = 0;
                pl.Rows[i]["chi3"] = 0;
            }
            string thangTruoc = thang == "1" ? "12" : $"{(int.Parse(thang) - 1)}";
            var dt = db.getDataTable($"SELECT '' as macaptren, ma_cskcb, ten_cskcb, tong_luot, tong_chi FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND ma_tinh='{idTinh}' AND tu_thang=den_thang AND den_thang={thangTruoc});");
            string tmp = ""; int index = 0;
            long luot1 = 0, luot2 = 0;
            double chi1 = 0, chi2 = 0;
            foreach (DataRow dr in dt.Rows)
            {
                tmp = $"{dr["ma_cskcb"]}";
                if (dicCSKCB.Keys.Contains(tmp))
                {
                    /* Tính toán */
                    index = dicCSKCB[tmp];
                    dicCSKCB.Remove(tmp);
                    /* Lượt */
                    luot1 = (long)pl.Rows[index]["luot1"];
                    luot2 = (long)dr["tong_luot"];
                    pl.Rows[index]["luot2"] = luot2;
                    pl.Rows[index]["luot3"] = luot1 - luot2;
                    /* Chi */
                    chi1 = (double)pl.Rows[index]["chi1"];
                    chi2 = (double)dr["tong_chi"];
                    pl.Rows[index]["chi2"] = chi2;
                }
                else
                {
                    /* Thêm vào phục lục */
                    luot2 = (long)dr["tong_luot"];
                    chi2 = (double)dr["tong_chi"];
                    pl.Rows.Add("", dr["ma_cskcb"], dr["ten_cskcb"], long.Parse("0"), luot2, (0 - luot2), double.Parse("0"), chi2, 0);
                }
            }
            pl.TableName = "PL03c";
            /* Làm tròn triệu đồng */
            foreach (DataRow dr in pl.Rows)
            {
                dr[6] = double.Parse($"{dr[6]}").lamTronTrieuDong(true);
                dr[7] = double.Parse($"{dr[7]}").lamTronTrieuDong(true);
                dr[8] = double.Parse($"{dr[6]}") - double.Parse($"{dr[7]}");
                dr["macaptren"] = matchMaCapTren.getValue($"{dr[1]}", $"{dr[1]}");
            }
            pl = sortDataTable(pl, "macaptren", "ma_cskcb");
            tmp = "";
            foreach (DataRow dr in pl.Rows)
            {
                if (tmp == $"{dr[0]}") { dr[0] = "-"; } else { tmp = $"{dr[0]}"; }
            }
            return pl;
        }

        private DataTable createPL01(dbSQLite dbBCThang, string idBC, string matinh, string nam, string thang, Dictionary<string, string> matchMaCapTren, Dictionary<string, string> dsMaCapTren)
        {
            /* Trường hợp không có dữ liệu DTGiao thì lấy từ thangdtgiao */
            DataTable data = new DataTable();
            string tmp = $"{dbBCThang.getValue($"SELECT SUM(dtgiao) AS X FROM thangpl01 WHERE id_bc='{idBC}'")}";
            if (tmp == "0")
            {
                /* Cập nhật dự toán giao CSKCB */
                /* - Lấy danh sách ID */
                var listCSKCB = new List<string>();
                data = dbBCThang.getDataTable($"SELECT ma_cskcb FROM thangpl01 WHERE id_bc='{idBC}'");
                foreach (DataRow r in data.Rows) { listCSKCB.Add($"{r[0]}"); }
                var dbDTGiao = dbBCThang;
                tmp = Path.Combine(AppHelper.pathAppData, $"BaoCaoThang{matinh}.db");
                if (dbBCThang.getPathDataFile() != tmp) { dbDTGiao = new dbSQLite(tmp); dbDTGiao.CreateBcThang(); }
                /* - Lấy dự toán giao hiện tại để cập nhật */
                data = dbDTGiao.getDataTable($"SELECT ma_cskcb, dtgiao FROM thangdtgiao WHERE nam={nam} AND idtinh='{matinh}' AND ma_cskcb IN ('{string.Join("','", listCSKCB)}')");
                var tsql = new List<string>();
                foreach (DataRow r in data.Rows)
                {
                    tmp = $"{r[1]}"; if (tmp == "0") { continue; }
                    tsql.Add($"UPDATE thangpl01 SET dtgiao = '{r[1]}' WHERE id_bc='{idBC}' AND ma_cskcb='{r[0]}'; ");
                }
                if (tsql.Count > 0) { dbBCThang.Execute(string.Join(Environment.NewLine, tsql)); }
            }

            /* Cập nhật dự toán giao CSKCB */
            data = dbBCThang.getDataTable($"SELECT ma_cskcb, ten_cskcb, dtgiao, tien_bhtt FROM thangpl01 WHERE id_bc='{idBC}' ORDER BY ma_cskcb;");
            foreach (DataRow r in data.Rows)
            {
                tmp = $"{r[0]}";
                if (matchMaCapTren.Keys.Contains(tmp) == false) { continue; }
                tmp = matchMaCapTren[tmp];
                r[0] = tmp;
                if (dsMaCapTren.Keys.Contains(tmp)) { r[1] = dsMaCapTren[tmp]; }
            }
            var groupedData = from row in data.AsEnumerable()
                              group row by new
                              {
                                  ma_cskcb = row.Field<string>("ma_cskcb"),
                                  ten_cskcb = row.Field<string>("ten_cskcb")
                              }
                              into grp
                              select new
                              {
                                  ma_cskcb = grp.Key.ma_cskcb,
                                  ten_cskcb = grp.Key.ten_cskcb,
                                  sum_dtgiao = grp.Sum(r => r.Field<double>("dtgiao")),
                                  sum_tien_bhtt = grp.Sum(r => r.Field<double>("tien_bhtt"))
                              };
            var PL01 = new DataTable("PL01");
            PL01.Columns.Add("ma_cskcb");
            PL01.Columns.Add("ten_cskcb");
            PL01.Columns.Add("dtgiao");
            PL01.Columns.Add("tien_bhtt");
            PL01.Columns.Add("tyle", typeof(double));
            PL01.Columns.Add("tyle_sd");
            foreach (var grp in groupedData)
            {
                PL01.Rows.Add(grp.ma_cskcb, grp.ten_cskcb,
                    grp.sum_dtgiao.lamTronTrieuDong(true).ToString(),
                    grp.sum_tien_bhtt.lamTronTrieuDong(true).ToString(),
                    "0");
            }
            for (int i = 0; i < PL01.Rows.Count; i++)
            {
                tmp = $"{PL01.Rows[i][2]}"; if (tmp == "0") { continue; }
                PL01.Rows[i][4] = Math.Round(double.Parse($"{PL01.Rows[i][3]}") * 100 / double.Parse(tmp), 2);
                PL01.Rows[i][5] = $"{PL01.Rows[i][4]}%";
            }
            PL01.DefaultView.Sort = "tyle DESC"; PL01 = PL01.DefaultView.ToTable();
            PL01.Columns.RemoveAt(4);
            return PL01;
        }

        private DataTable createPL02(dbSQLite db, string idBaoCao, string idTinh, string nameSheet, Dictionary<string, string> dmVung)
        {
            var tsql = $"SELECT * FROM thang{nameSheet.ToLower()} WHERE id_bc='{idBaoCao}';";
            DataTable pl = db.getDataTable(tsql);
            if (pl.Rows.Count == 0) { return new DataTable(nameSheet); }
            /* Bỏ [ma tỉnh] - ở cột tên tỉnh */
            for (int i = 0; i < pl.Rows.Count; i++) { pl.Rows[i]["ten_tinh"] = Regex.Replace($"{pl.Rows[i]["ten_tinh"]}", @"^V?\d+[ -]+", ""); }
            var phuLuc = new DataTable(nameSheet);
            phuLuc.Columns.Add("Mã Tỉnh"); /* 0 */
            phuLuc.Columns.Add("Tên tỉnh"); /* 1 */
            phuLuc.Columns.Add("Tỷ lệ nội trú (%)"); /* 2 */
            phuLuc.Columns.Add("Tên tỉnh 1"); /* 3 */
            phuLuc.Columns.Add("Ngày điều trị BQ (ngày)"); /* 4 */
            phuLuc.Columns.Add("Tên tỉnh 2"); /* 5 */
            phuLuc.Columns.Add("Chi BQ chung (Đồng)"); /* 6 */
            phuLuc.Columns.Add("Tên tỉnh 3"); /* 7 */
            phuLuc.Columns.Add("Chi BQ nội trú (Đồng)"); /* 8 */
            phuLuc.Columns.Add("Tên tỉnh 4"); /* 9 */
            phuLuc.Columns.Add("Chi BQ ngoại trú"); /* 10 */
            /* Lấy dòng tỉnh */
            var plview = pl.AsEnumerable();
            var view = plview.Where(x => x.Field<string>("ma_tinh") != "00").OrderByDescending(x => x.Field<double>("tyle_noitru")).ToList();
            foreach (DataRow row in view)
            {
                string bold = row["ma_tinh"].ToString() == idTinh ? "<b>" : "";
                phuLuc.Rows.Add($"{bold}{row["ma_tinh"]}", $"{bold}{row["ten_tinh"]}", $"{bold}{row["tyle_noitru"]}"
                    , "", "", "", "", "", "", "", "");
            }
            var lsField = new List<string>() { "ngay_dtri_bq", "chi_bq_chung", "chi_bq_noi", "chi_bq_ngoai" };
            int indexCols = 0, indexRow = -1;
            foreach (string field in lsField)
            {
                indexCols++;
                indexRow = -1;
                view = plview.Where(x => x.Field<string>("ma_tinh") != "00").OrderByDescending(x => x.Field<double>(field)).ToList();
                foreach (DataRow row in view)
                {
                    indexRow++; int colIndex = (indexCols * 2) + 1;
                    string bold = row["ma_tinh"].ToString() == idTinh ? "<b>" : "";
                    phuLuc.Rows[indexRow][colIndex] = $"{bold}{row["ten_tinh"]}";
                    phuLuc.Rows[indexRow][(colIndex + 1)] = $"{bold}{row[field]}";
                }
            }
            /* Dòng trống */
            phuLuc.Rows.Add("", "", "", "", "", "", "", "", "");
            /* Toàn quốc */
            view = plview.Where(x => x.Field<string>("ma_tinh") == "00").ToList();
            if (view.Count == 0) { phuLuc.Rows.Add("00", "Toàn quốc", "0", "Toàn quốc", "0", "Toàn quốc", "0", "Toàn quốc", "0", "Toàn quốc", "0"); }
            else
            {
                phuLuc.Rows.Add("00", "Toàn quốc", $"{view[0]["tyle_noitru"]}"
                    , "Toàn quốc", $"{view[0]["ngay_dtri_bq"]}"
                    , "Toàn quốc", $"{view[0]["chi_bq_chung"]}"
                    , "Toàn quốc", $"{view[0]["chi_bq_noi"]}"
                    , "Toàn quốc", $"{view[0]["chi_bq_ngoai"]}");
            }
            DataRow row00 = phuLuc.Rows[phuLuc.Rows.Count - 1];

            /* Vùng */
            var mavung = plview.Where(x => x.Field<string>("ma_tinh") == idTinh).Select(x => x.Field<string>("ma_vung")).First();
            var itemVung = new KeyValuePair<string, string>("V" + (mavung.Length < 2 ? "0" + mavung : mavung), $"Vùng {mavung}");
            if (dmVung.Any(x => x.Key.EndsWith(mavung))) { itemVung = dmVung.FirstOrDefault(x => x.Key.EndsWith(mavung)); }
            indexRow = plview.Count(x => x.Field<string>("ma_vung") == mavung);
            var vung = plview.Where(x => x.Field<string>("ma_vung") == mavung)
                .GroupBy(x => x.Field<string>("ma_vung"))
                .Select(g => new
                {
                    luot = g.Sum(x => x.Field<long>("tong_luot")),
                    luotNoi = g.Sum(x => x.Field<long>("tong_luot_noi")),
                    luotNgoai = g.Sum(x => x.Field<long>("tong_luot_ngoai")),
                    ngaydtr = g.Sum(x => x.Field<double>("ngay_dtri_bq") * x.Field<long>("tong_luot_noi")),
                    chi = g.Sum(x => x.Field<double>("tong_chi")),
                    chiNoi = g.Sum(x => x.Field<double>("tong_chi_noi")),
                    chiNgoai = g.Sum(x => x.Field<double>("tong_chi_ngoai"))
                })
                .FirstOrDefault();
            if (vung == null) { phuLuc.Rows.Add(itemVung.Key, itemVung.Value, "0", itemVung.Value, "0", itemVung.Value, "0", itemVung.Value, "0", itemVung.Value, "0"); }
            else
            {
                phuLuc.Rows.Add(itemVung.Key,
                    itemVung.Value, ((double)vung.luotNoi * 100 / (double)vung.luot).ToString("0.##"), /* Tỷ lệ nội */
                    itemVung.Value, (vung.ngaydtr / (double)vung.luotNoi).ToString("0.##"), /* ngày điều trị*/
                    itemVung.Value, (vung.chi / (double)vung.luot).ToString("0"), /* chi bình quân */
                    itemVung.Value, (vung.chiNoi / (double)vung.luotNoi).ToString("0"), /* chi bình quân nội */
                    itemVung.Value, (vung.chiNgoai / (double)vung.luotNgoai).ToString("0")); /* chi bình quân ngoại */
            }
            DataRow rowVung = phuLuc.Rows[phuLuc.Rows.Count - 1];
            /* Tỉnh */
            view = plview.Where(x => x.Field<string>("ma_tinh") == idTinh).ToList().GetRange(0, 1);
            if (view.Count == 0) { phuLuc.Rows.Add(idTinh, idTinh, "0", idTinh, "0", idTinh, "0", idTinh, "0", idTinh, "0"); }
            else
            {
                phuLuc.Rows.Add(idTinh, view[0]["ten_tinh"], $"{view[0]["tyle_noitru"]}"
                    , view[0]["ten_tinh"], $"{view[0]["ngay_dtri_bq"]}"
                    , view[0]["ten_tinh"], $"{view[0]["chi_bq_chung"]}"
                    , view[0]["ten_tinh"], $"{view[0]["chi_bq_noi"]}"
                    , view[0]["ten_tinh"], $"{view[0]["chi_bq_ngoai"]}");
            }
            DataRow rowTinh = phuLuc.Rows[phuLuc.Rows.Count - 1];
            /* Chênh so toàn quốc */
            string title = "Chênh so toàn quốc";
            phuLuc.Rows.Add("", title, $"{(double.Parse($"{rowTinh[2]}") - double.Parse($"{row00[2]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[4]}") - double.Parse($"{row00[4]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{row00[6]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[8]}") - double.Parse($"{row00[8]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[10]}") - double.Parse($"{row00[10]}")).ToString("0.##")}");

            /* Chênh với Vùng */
            title = "Chênh so vùng";
            phuLuc.Rows.Add("", title, $"{(double.Parse($"{rowTinh[2]}") - double.Parse($"{rowVung[2]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[4]}") - double.Parse($"{rowVung[4]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{rowVung[6]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[8]}") - double.Parse($"{rowVung[8]}")).ToString("0.##")}",
                title, $"{(double.Parse($"{rowTinh[10]}") - double.Parse($"{rowVung[10]}")).ToString("0.##")}");
            return phuLuc;
        }

        private DataTable createPL03a(dbSQLite db, string idBaoCao, string nameSheet, string thang, DataTable PL02, DataTable PL03a2, Dictionary<string, string> dmVung, Dictionary<string, string> matchMaCapTren)
        {
            var tsql = $"SELECT * FROM thang{nameSheet.ToLower()} WHERE id_bc='{idBaoCao}' ORDER BY tuyen_bv, hang_bv";
            var data = db.getDataTable(tsql).AsEnumerable();
            if (data.Count() == 0) { throw new Exception($"Dữ liệu PL03a không có dữ liệu ID_BC: {idBaoCao}"); }
            var phuLuc = new DataTable(nameSheet);
            phuLuc.Columns.Add("Mã"); /* 0 */
            phuLuc.Columns.Add("Hạng BV/ Tên CSKCB"); /* 1 */
            phuLuc.Columns.Add("Tỷ lệ nội trú tháng này(%)"); /* 2 */
            phuLuc.Columns.Add("Tỷ lệ nội trú tháng năm trước (%)"); /* 3 */
            phuLuc.Columns.Add("Tỷ lệ nội trú tăng-giảm (%)"); /* 4 */
            phuLuc.Columns.Add("Ngày điều trị BQ (ngày) tháng này"); /* 5 */
            phuLuc.Columns.Add("Ngày điều trị BQ (ngày) tháng năm trước"); /* 6 */
            phuLuc.Columns.Add("Ngày điều trị BQ (ngày) tăng-giảm"); /* 7 */
            phuLuc.Columns.Add("Chi BQ chung (Đồng) tháng này"); /* 8 */
            phuLuc.Columns.Add("Chi BQ chung (Đồng) tháng năm trước"); /* 9 */
            phuLuc.Columns.Add("Chi BQ chung (Đồng) tăng-giảm"); /* 10 */
            phuLuc.Columns.Add("Chi BQ nội trú (Đồng) tháng này"); /* 11 */
            phuLuc.Columns.Add("Chi BQ nội trú (Đồng) tháng năm trước"); /* 12 */
            phuLuc.Columns.Add("Chi BQ nội trú (Đồng) tăng-giảm"); /* 13 */
            phuLuc.Columns.Add("Chi BQ ngoại trú tháng này"); /* 14 */
            phuLuc.Columns.Add("Chi BQ ngoại trú tháng năm trước"); /* 15 */
            phuLuc.Columns.Add("Chi BQ ngoại trú tăng-giảm"); /* 16 */

            /* 4 Dòng đầu copy của PL02a phần chênh lệnh */
            if (PL02.Rows.Count > 5)
            {
                int pl02Count = PL02.Rows.Count; int IndexPL02 = 0;
                for (int i = 5; i > 2; i--)
                {
                    IndexPL02 = pl02Count - i;
                    phuLuc.Rows.Add(PL02.Rows[IndexPL02][0], PL02.Rows[IndexPL02][1]
                        , PL02.Rows[IndexPL02][2], "0", "0" /* tyle_noitru */
                        , PL02.Rows[IndexPL02][4], "0", "0" /* ngay_dtri_bq */
                        , PL02.Rows[IndexPL02][6], "0", "0" /* chi_bq_chung */
                        , PL02.Rows[IndexPL02][8], "0", "0" /* chi_bq_noi */
                        , PL02.Rows[IndexPL02][10], "0", "0" /* chi_bq_ngoai */);
                }
                if (PL03a2.Rows.Count > 5)
                {
                    int PL03a2Count = PL03a2.Rows.Count; int IndexPL03a2 = 0; int index = 0;
                    for (int i = 5; i > 2; i--)
                    {
                        IndexPL03a2 = PL03a2Count - i;
                        /* tyle_noitru */
                        phuLuc.Rows[index][3] = $"{PL03a2.Rows[IndexPL03a2][2]}";
                        phuLuc.Rows[index][4] = Math.Round(double.Parse($"{phuLuc.Rows[index][2]}") - double.Parse($"{phuLuc.Rows[index][3]}"), 2).ToString();
                        /* ngay_dtri_bq */
                        phuLuc.Rows[index][6] = $"{PL03a2.Rows[IndexPL03a2][4]}";
                        phuLuc.Rows[index][7] = Math.Round(double.Parse($"{phuLuc.Rows[index][5]}") - double.Parse($"{phuLuc.Rows[index][6]}"), 2).ToString();
                        /* chi_bq_chung */
                        phuLuc.Rows[index][9] = $"{PL03a2.Rows[IndexPL03a2][6]}";
                        phuLuc.Rows[index][10] = Math.Round(double.Parse($"{phuLuc.Rows[index][8]}") - double.Parse($"{phuLuc.Rows[index][9]}"), 2).ToString();
                        /* chi_bq_noi */
                        phuLuc.Rows[index][12] = $"{PL03a2.Rows[IndexPL03a2][8]}";
                        phuLuc.Rows[index][13] = Math.Round(double.Parse($"{phuLuc.Rows[index][11]}") - double.Parse($"{phuLuc.Rows[index][12]}"), 2).ToString();
                        /* chi_bq_ngoai */
                        phuLuc.Rows[index][15] = $"{PL03a2.Rows[IndexPL03a2][10]}";
                        phuLuc.Rows[index][16] = Math.Round(double.Parse($"{phuLuc.Rows[index][14]}") - double.Parse($"{phuLuc.Rows[index][15]}"), 2).ToString();
                        index++;
                    }
                }
            }
            phuLuc.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            var listTuyen = new List<string>() { "*", "T", "H", "X" };
            string hang = "", macsyt = "", tmp = ""; int lr = -1;
            var matchIndex = new Dictionary<string, int>();
            foreach (string tuyen in listTuyen)
            {
                var view = new List<DataRow>();
                if (tuyen == "*")
                {
                    view = data.Where(x => x.Field<string>("tuyen_bv") == "")
                        .OrderBy(x => x.Field<string>("hang_bv"))
                        .ThenBy(x => x.Field<string>("ma_cskcb"))
                        .ThenByDescending(x => x.Field<long>("thang"))
                        .ToList();
                }
                else
                {
                    view = data.Where(x => x.Field<string>("tuyen_bv").StartsWith(tuyen))
                        .OrderBy(x => x.Field<string>("hang_bv"))
                        .ThenBy(x => x.Field<string>("ma_cskcb"))
                        .ThenByDescending(x => x.Field<long>("thang"))
                        .ToList();
                }
                if (view.Count() == 0) { continue; }
                string tenTuyen = "(*)";
                switch (tuyen)
                {
                    case "T": tenTuyen = "Tỉnh"; break;
                    case "H": tenTuyen = "Huyện"; break;
                    case "X": tenTuyen = "Xã"; break;
                    default: break;
                }
                phuLuc.Rows.Add("T" + (tuyen == "" ? "0" : tuyen), $"Tuyến {tenTuyen}", "", "", "", "", "");
                macsyt = "";
                foreach (DataRow row in view)
                {
                    tmp = $"{row["ma_cskcb"]}";
                    lr = -1;
                    if (matchIndex.Keys.Contains(tmp) == false) { matchIndex.Add(tmp, phuLuc.Rows.Count); lr = phuLuc.Rows.Count; }
                    if (lr > -1)
                    {
                        macsyt = $"{row["ma_cskcb"]}";
                        hang = $"{row["hang_bv"]}".Trim(); if (hang == "") { hang = "*"; }
                        if (thang == $"{row["thang"]}")
                        {
                            phuLuc.Rows.Add(macsyt, $"{hang}/ {row["ten_cskcb"]}"
                               , $"{row["tyle_noitru"]}", "0", $"{row["tyle_noitru"]}"
                               , $"{row["ngay_dtri_bq"]}", "0", $"{row["ngay_dtri_bq"]}"
                               , $"{row["chi_bq_chung"]}", "0", $"{row["chi_bq_chung"]}"
                               , $"{row["chi_bq_noi"]}", "0", $"{row["chi_bq_noi"]}"
                               , $"{row["chi_bq_ngoai"]}", "0", $"{row["chi_bq_ngoai"]}");
                        }
                        else
                        {
                            phuLuc.Rows.Add($"{row["ma_cskcb"]}", $"{hang}/ {row["ten_cskcb"]}"
                               , "0", $"{row["tyle_noitru"]}", $"-{row["tyle_noitru"]}"
                               , "0", $"{row["ngay_dtri_bq"]}", $"-{row["ngay_dtri_bq"]}"
                               , "0", $"{row["chi_bq_chung"]}", $"-{row["chi_bq_chung"]}"
                               , "0", $"{row["chi_bq_noi"]}", $"-{row["chi_bq_noi"]}"
                               , "0", $"{row["chi_bq_ngoai"]}", $"-{row["chi_bq_ngoai"]}");
                        }
                        continue;
                    }
                    lr = matchIndex[tmp];
                    if (thang == $"{row["thang"]}")
                    {
                        phuLuc.Rows[lr][2] = $"{row["tyle_noitru"]}";
                        phuLuc.Rows[lr][5] = $"{row["ngay_dtri_bq"]}";
                        phuLuc.Rows[lr][8] = $"{row["chi_bq_chung"]}";
                        phuLuc.Rows[lr][11] = $"{row["chi_bq_noi"]}";
                        phuLuc.Rows[lr][14] = $"{row["chi_bq_ngoai"]}";
                    }
                    else
                    {
                        phuLuc.Rows[lr][3] = $"{row["tyle_noitru"]}";
                        phuLuc.Rows[lr][6] = $"{row["ngay_dtri_bq"]}";
                        phuLuc.Rows[lr][9] = $"{row["chi_bq_chung"]}";
                        phuLuc.Rows[lr][12] = $"{row["chi_bq_noi"]}";
                        phuLuc.Rows[lr][15] = $"{row["chi_bq_ngoai"]}";
                    }
                    phuLuc.Rows[lr][4] = $"{Math.Round(double.Parse($"{phuLuc.Rows[lr][2]}") - double.Parse($"{phuLuc.Rows[lr][3]}"), 2)}";
                    phuLuc.Rows[lr][7] = $"{Math.Round(double.Parse($"{phuLuc.Rows[lr][5]}") - double.Parse($"{phuLuc.Rows[lr][6]}"), 2)}";
                    phuLuc.Rows[lr][10] = $"{(double.Parse($"{phuLuc.Rows[lr][8]}") - double.Parse($"{phuLuc.Rows[lr][9]}"))}";
                    phuLuc.Rows[lr][13] = $"{(double.Parse($"{phuLuc.Rows[lr][11]}") - double.Parse($"{phuLuc.Rows[lr][12]}"))}";
                    phuLuc.Rows[lr][16] = $"{(double.Parse($"{phuLuc.Rows[lr][14]}") - double.Parse($"{phuLuc.Rows[lr][15]}"))}";
                }
            }
            return addMaCapTren(phuLuc, matchMaCapTren);
        }

        private DataTable createPL03b(dbSQLite db, string idBaoCao, string nameSheet, DataTable PL02, long namBC, Dictionary<string, string> matchMaCapTren)
        {
            var data = db.getDataTable($"SELECT * FROM thang{nameSheet.ToLower()} WHERE id_bc='{idBaoCao}' AND nam={namBC} ORDER BY tuyen_bv, hang_bv").AsEnumerable();
            if (data.Count() == 0) { throw new Exception($"Dữ liệu PL03a không có dữ liệu ID_BC: {idBaoCao}"); }
            var phuLuc = new DataTable(nameSheet);
            phuLuc.Columns.Add("Mã"); /* 0 */
            phuLuc.Columns.Add("Hạng BV/ Tên CSKCB"); /* 1 */
            phuLuc.Columns.Add("Tỷ lệ nội trú (%)"); /* 2 */
            phuLuc.Columns.Add("Ngày điều trị BQ (ngày)"); /* 3 */
            phuLuc.Columns.Add("Chi BQ chung (Đồng)"); /* 4 */
            phuLuc.Columns.Add("Chi BQ nội trú (Đồng)"); /* 5 */
            phuLuc.Columns.Add("Chi BQ ngoại trú"); /* 6 */
            /* 4 Dòng đầu copy của PL02a, PL02b phần chênh lệnh */
            if (PL02.Rows.Count > 5)
            {
                int pl02Count = PL02.Rows.Count; int IndexPL02 = 0;
                for (int i = 5; i > 2; i--)
                {
                    IndexPL02 = pl02Count - i;
                    phuLuc.Rows.Add(PL02.Rows[IndexPL02][0], PL02.Rows[IndexPL02][1]
                        , PL02.Rows[IndexPL02][2] /* tyle_noitru */
                        , PL02.Rows[IndexPL02][4] /* ngay_dtri_bq */
                        , PL02.Rows[IndexPL02][6] /* chi_bq_chung */
                        , PL02.Rows[IndexPL02][8] /* chi_bq_noi */
                        , PL02.Rows[IndexPL02][10] /* chi_bq_ngoai */);
                }
            }

            phuLuc.Rows.Add("", "", "", "", "", "", "");
            var listTuyen = new List<string>() { "*", "T", "H", "X" };
            string hang = "";
            foreach (string tuyen in listTuyen)
            {
                var view = new List<DataRow>();
                if (tuyen == "*") { view = data.Where(x => x.Field<string>("tuyen_bv") == "").OrderBy(x => x.Field<string>("hang_bv")).ToList(); }
                else { view = data.Where(x => x.Field<string>("tuyen_bv").StartsWith(tuyen)).OrderBy(x => x.Field<string>("hang_bv")).ToList(); }
                if (view.Count() == 0) { continue; }
                string tenTuyen = "(*)";
                switch (tuyen)
                {
                    case "T": tenTuyen = "Tỉnh"; break;
                    case "H": tenTuyen = "Huyện"; break;
                    case "X": tenTuyen = "Xã"; break;
                    default: break;
                }
                phuLuc.Rows.Add("T" + (tuyen == "" ? "0" : tuyen), $"Tuyến {tenTuyen}", "", "", "", "", "");
                foreach (DataRow row in view)
                {
                    hang = $"{row["hang_bv"]}".Trim(); if (hang == "") { hang = "*"; }
                    phuLuc.Rows.Add($"{row["ma_cskcb"]}", $"{hang}/ {row["ten_cskcb"]}", $"{row["tyle_noitru"]}", $"{row["ngay_dtri_bq"]}", $"{row["chi_bq_chung"]}", $"{row["chi_bq_noi"]}", $"{row["chi_bq_ngoai"]}");
                }
            }
            return addMaCapTren(phuLuc, matchMaCapTren);
        }

        private DataTable createPL04a(dbSQLite db, string idBaoCao, string idTinh, Dictionary<string, string> dmVung)
        {
            /* Chỉ lấy danh sách trong vùng thống kế */
            string maVung = $"{db.getValue($"SELECT ma_vung FROM thangpl04a WHERE id_bc='{idBaoCao}' AND ma_tinh='{idTinh}' LIMIT 1;")}";
            var pl = db.getDataTable($"SELECT * FROM thangpl04a WHERE id_bc='{idBaoCao}' AND ma_vung='{maVung}';");

            /* Bỏ [ma tỉnh] - ở cột tên tỉnh */
            var phuLuc = new DataTable("PL04a");
            phuLuc.Columns.Add("Mã Tỉnh"); /* 0 */
            phuLuc.Columns.Add("Tên tỉnh"); /* 1 */
            phuLuc.Columns.Add("BQ_XN (đồng)"); /* 2 */
            phuLuc.Columns.Add("BQ_CĐHA (đồng)"); /* 3 */
            phuLuc.Columns.Add("BQ_THUOC (đồng)"); /* 4 */
            phuLuc.Columns.Add("BQ_PTTT (đồng)"); /* 5 */
            phuLuc.Columns.Add("BQ_VTYT (đồng)"); /* 6 */
            phuLuc.Columns.Add("BQ_GIUONG (đồng)"); /* 7 */
            phuLuc.Columns.Add("Ngày thanh toán BQ"); /* 8 */
            /* Lấy dòng tỉnh */
            var plview = pl.AsEnumerable();
            var view = plview.Where(x => !x.Field<string>("ma_tinh").StartsWith("V")).OrderByDescending(x => x.Field<double>("chi_bq_xn")).ToList();
            foreach (DataRow row in view)
            {
                string bold = row["ma_tinh"].ToString() == idTinh ? "<b>" : "";
                phuLuc.Rows.Add($"{bold}{row["ma_tinh"]}", $"{bold}{row["ten_tinh"]}"
                    , $"{bold}{row["chi_bq_xn"]}"
                    , $"{bold}{row["chi_bq_cdha"]}"
                    , $"{bold}{row["chi_bq_thuoc"]}"
                    , $"{bold}{row["chi_bq_pttt"]}"
                    , $"{bold}{row["chi_bq_vtyt"]}"
                    , $"{bold}{row["chi_bq_giuong"]}"
                    , $"{bold}{row["ngay_ttbq"]}");
            }
            /* Dòng trống */
            phuLuc.Rows.Add("", "", "", "", "", "", "", "", "");
            /* Toàn quốc */
            view = new List<DataRow>();
            var dt = db.getDataTable($"SELECT * FROM thangpl04a WHERE id_bc='{idBaoCao}' AND ma_tinh='00' LIMIT 1");
            if (dt.Rows.Count > 0) { view = new List<DataRow>() { dt.Rows[0] }; }
            if (view.Count == 0) { phuLuc.Rows.Add("00", "Toàn quốc", "0", "0", "0", "0", "0", "0", "0"); }
            else
            {
                phuLuc.Rows.Add("00", "Toàn quốc", $"{view[0]["chi_bq_xn"]}"
                    , $"{view[0]["chi_bq_cdha"]}"
                    , $"{view[0]["chi_bq_thuoc"]}"
                    , $"{view[0]["chi_bq_pttt"]}"
                    , $"{view[0]["chi_bq_vtyt"]}"
                    , $"{view[0]["chi_bq_giuong"]}"
                    , $"{view[0]["ngay_ttbq"]}");
            }
            DataRow row00 = phuLuc.Rows[phuLuc.Rows.Count - 1];
            /* Vùng */
            view = plview.Where(x => x.Field<string>("ma_tinh").StartsWith("V")).ToList();
            if (view.Count > 0)
            {
                phuLuc.Rows.Add(view[0]["ma_tinh"], view[0]["ten_tinh"], $"{view[0]["chi_bq_xn"]}"
                    , $"{view[0]["chi_bq_cdha"]}"
                    , $"{view[0]["chi_bq_thuoc"]}"
                    , $"{view[0]["chi_bq_pttt"]}"
                    , $"{view[0]["chi_bq_vtyt"]}"
                    , $"{view[0]["chi_bq_giuong"]}"
                    , $"{view[0]["ngay_ttbq"]}");
            }
            else
            {
                var itemVung = new KeyValuePair<string, string>("V" + (maVung.Length < 2 ? "0" + maVung : maVung), $"Vùng {maVung}");
                if (dmVung.Any(x => x.Key.EndsWith(maVung))) { itemVung = dmVung.FirstOrDefault(x => x.Key.EndsWith(maVung)); }

                int indexRow = plview.Count(x => !x.Field<string>("ma_tinh").StartsWith("V"));
                var vung = plview.Where(x => !x.Field<string>("ma_tinh").StartsWith("V"))
                    .GroupBy(x => x.Field<string>("ma_vung"))
                    .Select(g => new
                    {
                        chi_bq_xn = g.Sum(x => x.Field<double>("chi_bq_xn")) / indexRow,
                        chi_bq_cdha = g.Sum(x => x.Field<double>("chi_bq_cdha")) / indexRow,
                        chi_bq_thuoc = g.Sum(x => x.Field<double>("chi_bq_thuoc")) / indexRow,
                        chi_bq_pttt = g.Sum(x => x.Field<double>("chi_bq_pttt")) / indexRow,
                        chi_bq_vtyt = g.Sum(x => x.Field<double>("chi_bq_vtyt")) / indexRow,
                        chi_bq_giuong = g.Sum(x => x.Field<double>("chi_bq_giuong")) / indexRow,
                        ngay_ttbq = g.Sum(x => x.Field<double>("ngay_ttbq")) / indexRow
                    })
                    .FirstOrDefault();
                if (vung == null) { phuLuc.Rows.Add(itemVung.Key, itemVung.Value, "0", "0", "0", "0", "0", "0", "0"); }
                else
                {
                    phuLuc.Rows.Add(itemVung.Key, itemVung.Value
                        , $"{vung.chi_bq_xn.ToString("0.##")}"
                        , $"{vung.chi_bq_cdha.ToString("0.##")}"
                        , $"{vung.chi_bq_thuoc.ToString("0.##")}"
                        , $"{vung.chi_bq_pttt.ToString("0.##")}"
                        , $"{vung.chi_bq_vtyt.ToString("0.##")}"
                        , $"{vung.chi_bq_giuong.ToString("0.##")}"
                        , $"{vung.ngay_ttbq.ToString("0.##")}");
                }
            }

            DataRow rowVung = phuLuc.Rows[phuLuc.Rows.Count - 1];
            /* Tỉnh */
            view = plview.Where(x => x.Field<string>("ma_tinh") == idTinh).ToList().GetRange(0, 1);
            if (view.Count == 0) { phuLuc.Rows.Add(idTinh, idTinh, "0", "0", "0", "0", "0", "0", "0"); }
            else
            {
                phuLuc.Rows.Add(idTinh, view[0]["ten_tinh"]
                    , $"{view[0]["chi_bq_xn"]}"
                    , $"{view[0]["chi_bq_cdha"]}"
                    , $"{view[0]["chi_bq_thuoc"]}"
                    , $"{view[0]["chi_bq_pttt"]}"
                    , $"{view[0]["chi_bq_vtyt"]}"
                    , $"{view[0]["chi_bq_giuong"]}"
                    , $"{view[0]["ngay_ttbq"]}");
            }
            DataRow rowTinh = phuLuc.Rows[phuLuc.Rows.Count - 1];
            /* Chênh so toàn quốc */
            phuLuc.Rows.Add("", "Chênh so toàn quốc", $"{(double.Parse($"{rowTinh[2]}") - double.Parse($"{row00[2]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[3]}") - double.Parse($"{row00[3]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[4]}") - double.Parse($"{row00[4]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[5]}") - double.Parse($"{row00[5]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{row00[6]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[7]}") - double.Parse($"{row00[7]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[8]}") - double.Parse($"{row00[8]}")).ToString("0.##")}");

            /* Chênh với Vùng */
            phuLuc.Rows.Add("", "Chênh so vùng", $"{(double.Parse($"{rowTinh[2]}") - double.Parse($"{rowVung[2]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[3]}") - double.Parse($"{rowVung[3]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[4]}") - double.Parse($"{rowVung[4]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[5]}") - double.Parse($"{rowVung[5]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[6]}") - double.Parse($"{rowVung[6]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[7]}") - double.Parse($"{rowVung[7]}")).ToString("0.##")}"
                , $"{(double.Parse($"{rowTinh[8]}") - double.Parse($"{rowVung[8]}")).ToString("0.##")}");
            return phuLuc;
        }

        private DataTable createPL04b(dbSQLite db, string idBaoCao, string idTinh, string thang, Dictionary<string, string> matchMaCapTren)
        {
            var data = db.getDataTable($"SELECT * FROM thangpl04b WHERE id_bc='{idBaoCao}';").AsEnumerable();
            var phuLuc = new DataTable("PL04b");
            phuLuc.Columns.Add("Mã"); /* 0 */
            phuLuc.Columns.Add("Hạng BV/ Tên CSKCB"); /* 1 */

            phuLuc.Columns.Add("BQ_XN (đồng) tháng này"); /* 2 */
            phuLuc.Columns.Add("BQ_XN (đồng) tháng năm trước"); /* 3 */
            phuLuc.Columns.Add("BQ_XN (đồng) tăng giảm"); /* 4 */

            phuLuc.Columns.Add("BQ_CĐHA (đồng) tháng này"); /* 5 */
            phuLuc.Columns.Add("BQ_CĐHA (đồng) tháng năm trước"); /* 6 */
            phuLuc.Columns.Add("BQ_CĐHA (đồng) tăng giảm"); /* 7 */

            phuLuc.Columns.Add("BQ_THUOC (đồng) tháng này"); /* 8 */
            phuLuc.Columns.Add("BQ_THUOC (đồng) tháng năm trước"); /* 9 */
            phuLuc.Columns.Add("BQ_THUOC (đồng) tăng giảm"); /* 10 */

            phuLuc.Columns.Add("BQ_PTTT (đồng) tháng này"); /* 11 */
            phuLuc.Columns.Add("BQ_PTTT (đồng) tháng năm trước"); /* 12 */
            phuLuc.Columns.Add("BQ_PTTT (đồng) tăng giảm"); /* 13 */

            phuLuc.Columns.Add("BQ_VTYT (đồng) tháng này"); /* 14 */
            phuLuc.Columns.Add("BQ_VTYT (đồng) tháng năm trước"); /* 15 */
            phuLuc.Columns.Add("BQ_VTYT (đồng) tăng giảm"); /* 16 */

            phuLuc.Columns.Add("BQ_GIUONG (đồng) tháng này"); /* 17 */
            phuLuc.Columns.Add("BQ_GIUONG (đồng) tháng năm trước"); /* 18 */
            phuLuc.Columns.Add("BQ_GIUONG (đồng) tăng giảm"); /* 19 */

            phuLuc.Columns.Add("Ngày thanh toán BQ tháng này"); /* 20 */
            phuLuc.Columns.Add("Ngày thanh toán BQ tháng năm trước"); /* 21 */
            phuLuc.Columns.Add("Ngày thanh toán BQ tăng giảm"); /* 22 */

            var listTuyen = new List<string>() { "*", "T", "H", "X" };
            string hang = "", macsyt = "", tmp = ""; int lr = 0;
            var matchIndex = new Dictionary<string, int>();
            foreach (string tuyen in listTuyen)
            {
                var view = new List<DataRow>();
                if (tuyen == "*")
                {
                    view = data.Where(x => x.Field<string>("tuyen_bv") == "")
                                .OrderBy(x => x.Field<string>("hang_bv"))
                                .ThenBy(x => x.Field<string>("ma_cskcb"))
                                .ThenByDescending(x => x.Field<long>("thang"))
                                .ToList();
                }
                else
                {
                    view = data.Where(x => x.Field<string>("tuyen_bv").StartsWith(tuyen))
                                .OrderBy(x => x.Field<string>("hang_bv"))
                                .ThenBy(x => x.Field<string>("ma_cskcb"))
                                .ThenByDescending(x => x.Field<long>("thang"))
                                .ToList();
                }
                if (view.Count() == 0) { continue; }
                string tenTuyen = "(*)";
                switch (tuyen)
                {
                    case "T": tenTuyen = "Tỉnh"; break;
                    case "H": tenTuyen = "Huyện"; break;
                    case "X": tenTuyen = "Xã"; break;
                    default: break;
                }
                phuLuc.Rows.Add("T" + (tuyen == "" ? "0" : tuyen), $"Tuyến {tenTuyen}", "", "", "", "", "", "", "");
                foreach (DataRow row in view)
                {
                    tmp = $"{row["ma_cskcb"]}";
                    lr = -1;
                    if (matchIndex.Keys.Contains(tmp) == false) { matchIndex.Add(tmp, phuLuc.Rows.Count); lr = phuLuc.Rows.Count; }
                    if (lr > -1)
                    {
                        macsyt = $"{row["ma_cskcb"]}";
                        hang = $"{row["hang_bv"]}".Trim(); if (hang == "") { hang = "*"; }
                        if (thang == $"{row["thang"]}")
                        {
                            phuLuc.Rows.Add(macsyt, $"{hang}/ {row["ten_cskcb"]}"
                               , $"{row["chi_bq_xn"]}", "0", $"{row["chi_bq_xn"]}"
                               , $"{row["chi_bq_cdha"]}", "0", $"{row["chi_bq_cdha"]}"
                               , $"{row["chi_bq_thuoc"]}", "0", $"{row["chi_bq_thuoc"]}"
                               , $"{row["chi_bq_pttt"]}", "0", $"{row["chi_bq_pttt"]}"
                               , $"{row["chi_bq_vtyt"]}", "0", $"{row["chi_bq_vtyt"]}"
                               , $"{row["chi_bq_giuong"]}", "0", $"{row["chi_bq_giuong"]}"
                               , $"{row["ngay_ttbq"]}", "0", $"{row["ngay_ttbq"]}");
                        }
                        else
                        {
                            phuLuc.Rows.Add($"{row["ma_cskcb"]}", $"{hang}/ {row["ten_cskcb"]}"
                               , "0", $"{row["chi_bq_xn"]}", $"-{row["chi_bq_xn"]}"
                               , "0", $"{row["chi_bq_cdha"]}", $"-{row["chi_bq_cdha"]}"
                               , "0", $"{row["chi_bq_thuoc"]}", $"-{row["chi_bq_thuoc"]}"
                               , "0", $"{row["chi_bq_pttt"]}", $"-{row["chi_bq_pttt"]}"
                               , "0", $"{row["chi_bq_vtyt"]}", $"-{row["chi_bq_vtyt"]}"
                               , "0", $"{row["chi_bq_giuong"]}", $"-{row["chi_bq_giuong"]}"
                               , "0", $"{row["ngay_ttbq"]}", $"-{row["ngay_ttbq"]}");
                        }
                        continue;
                    }
                    lr = matchIndex[tmp];
                    if (thang == $"{row["thang"]}")
                    {
                        phuLuc.Rows[lr][2] = $"{row["chi_bq_xn"]}";
                        phuLuc.Rows[lr][5] = $"{row["chi_bq_cdha"]}";
                        phuLuc.Rows[lr][8] = $"{row["chi_bq_thuoc"]}";
                        phuLuc.Rows[lr][11] = $"{row["chi_bq_pttt"]}";
                        phuLuc.Rows[lr][14] = $"{row["chi_bq_vtyt"]}";
                        phuLuc.Rows[lr][17] = $"{row["chi_bq_giuong"]}";
                        phuLuc.Rows[lr][20] = $"{row["ngay_ttbq"]}";
                    }
                    else
                    {
                        phuLuc.Rows[lr][3] = $"{row["chi_bq_xn"]}";
                        phuLuc.Rows[lr][6] = $"{row["chi_bq_cdha"]}";
                        phuLuc.Rows[lr][9] = $"{row["chi_bq_thuoc"]}";
                        phuLuc.Rows[lr][12] = $"{row["chi_bq_pttt"]}";
                        phuLuc.Rows[lr][15] = $"{row["chi_bq_vtyt"]}";
                        phuLuc.Rows[lr][18] = $"{row["chi_bq_giuong"]}";
                        phuLuc.Rows[lr][21] = $"{row["ngay_ttbq"]}";
                    }
                    phuLuc.Rows[lr][4] = $"{(double.Parse($"{phuLuc.Rows[lr][2]}") - double.Parse($"{phuLuc.Rows[lr][3]}"))}";
                    phuLuc.Rows[lr][7] = $"{(double.Parse($"{phuLuc.Rows[lr][5]}") - double.Parse($"{phuLuc.Rows[lr][6]}"))}";
                    phuLuc.Rows[lr][10] = $"{(double.Parse($"{phuLuc.Rows[lr][8]}") - double.Parse($"{phuLuc.Rows[lr][9]}"))}";
                    phuLuc.Rows[lr][13] = $"{(double.Parse($"{phuLuc.Rows[lr][11]}") - double.Parse($"{phuLuc.Rows[lr][12]}"))}";
                    phuLuc.Rows[lr][16] = $"{(double.Parse($"{phuLuc.Rows[lr][14]}") - double.Parse($"{phuLuc.Rows[lr][15]}"))}";
                    phuLuc.Rows[lr][19] = $"{(double.Parse($"{phuLuc.Rows[lr][17]}") - double.Parse($"{phuLuc.Rows[lr][18]}"))}";
                    phuLuc.Rows[lr][22] = $"{Math.Round(double.Parse($"{phuLuc.Rows[lr][20]}") - double.Parse($"{phuLuc.Rows[lr][21]}"), 2)}";
                }
            }
            return addMaCapTren(phuLuc, matchMaCapTren);
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
            /* Lấy năm, tháng báo cáo */
            string nam = ""; string thang = "";
            var data = dbConnect.getDataTable($"SELECT den_thang, nam FROM thangb02 WHERE id_bc='{idBaoCao}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
            if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
            if (nam == "")
            {
                data = dbConnect.getDataTable($"SELECT den_thang, nam FROM thangb01 WHERE id_bc='{idBaoCao}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
            }
            if (nam == "")
            {
                data = dbConnect.getDataTable($"SELECT den_thang, nam FROM thangb04 WHERE id_bc='{idBaoCao}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
            }
            if (nam == "") { throw new Exception("Không xác định được Năm, Tháng báo cáo"); }

            bcThang.Add("nam1", nam);
            bcThang.Add("nam2", (int.Parse(nam) - 1).ToString());
            bcThang.Add("thang", thang);
            tmp = bcThang["thang"].Length < 2 ? "0" + bcThang["thang"] : bcThang["thang"];
            bcThang.Add("ngay2", $"01/{tmp}/{bcThang["nam1"]}");
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
            bcThang.Add("x2", $"{item["dtcsyt_trongnam"]}".lamTronTrieuDong(true));
            bcThang.Add("x3", $"{item["dtcsyt_chikcb"]}".lamTronTrieuDong(true));
            bcThang.Add("x4", $"{item["dtcsyt_tlsudungnam"]}");

            /* x5 integer not null default 0 /* xếp bn toàn quốc */
            tmp = getPosition("", maTinh, "dtcsyt_tlsudungnam", ldata);
            bcThang.Add("x5", tmp);
            /* x6 integer not null default 0 /* xếp thứ bao nhiêu so với vùng */
            tmp = getPosition(mavung, maTinh, "dtcsyt_trongnam", ldata);
            bcThang.Add("x6", $"{tmp}");
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
            bcThang.Add("x21", $"{item["t_bhtt"]}".lamTronTrieuDong(true));
            bcThang.Add("x22", $"{item["t_bhtt_ngoai"]}".lamTronTrieuDong(true));
            bcThang.Add("x23", $"{item["t_bhtt_noi"]}".lamTronTrieuDong(true));
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
            bcThang.Add("x24", $"{item["t_bhtt"]}".lamTronTrieuDong(true));
            bcThang.Add("x25", $"{item["t_bhtt_ngoai"]}".lamTronTrieuDong(true));
            bcThang.Add("x26", $"{item["t_bhtt_noi"]}".lamTronTrieuDong(true));

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
            bcThang.Add("x27", $"{item["t_bhtt"]}".lamTronTrieuDong(true));
            bcThang.Add("x28", $"{item["t_bhtt_ngoai"]}".lamTronTrieuDong(true));
            bcThang.Add("x29", $"{item["t_bhtt_noi"]}".lamTronTrieuDong(true));

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
            bcThang.Add("x30", $"{item["t_bhtt"]}".lamTronTrieuDong(true));
            bcThang.Add("x31", $"{item["t_bhtt_ngoai"]}".lamTronTrieuDong(true));
            bcThang.Add("x32", $"{item["t_bhtt_noi"]}".lamTronTrieuDong(true));

            /* Tăng giảm so với cùng kỳ năm trước
             * ,m13lc13 real not null default 0 /* Tổng lượt = 2+3 -(x15-x9)
                ,m13lc23 real not null default 0 /* Lượt ngoại = -(x16-x10)
                ,m13lc33 real not null default 0 /* Lượt nội = -(x17-x11)
                ,m13lc43 real not null default 0 /* Tổng lượt = 5+6 -(x18-x12)
                ,m13lc53 real not null default 0 /* Lượt ngoại = -(x19-x13)
                ,m13lc63 real not null default 0 /* Lượt nội = -(x20-x14) */
            bcThang.Add("m13lc13", $"{(double.Parse(bcThang["x9"]) - double.Parse(bcThang["x15"]))}");
            bcThang.Add("m13lc23", $"{(double.Parse(bcThang["x10"]) - double.Parse(bcThang["x16"]))}");
            bcThang.Add("m13lc33", $"{(double.Parse(bcThang["x11"]) - double.Parse(bcThang["x17"]))}");
            bcThang.Add("m13lc43", $"{(double.Parse(bcThang["x12"]) - double.Parse(bcThang["x18"]))}");
            bcThang.Add("m13lc53", $"{(double.Parse(bcThang["x13"]) - double.Parse(bcThang["x19"]))}");
            bcThang.Add("m13lc63", $"{(double.Parse(bcThang["x14"]) - double.Parse(bcThang["x20"]))}");
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
             *  ,m13cc13 real not null default 0 /* Tổng lượt = 2+3 -(x27-x21)
                ,m13cc23 real not null default 0 /* Chi ngoại trú = -(x28-x22)
                ,m13cc33 real not null default 0 /* Chi nội trú = -(x29-x23)
                ,m13cc43 real not null default 0 /* Tổng lượt = 5+6 -(x30-x24)
                ,m13cc53 real not null default 0 /* Chi ngoại trú = -(x31-x25)
                ,m13cc63 real not null default 0 /* Chi nội trú = -(x32-x26) */
            bcThang.Add("m13cc13", $"{(double.Parse(bcThang["x21"]) - double.Parse(bcThang["x27"]))}");
            bcThang.Add("m13cc23", $"{(double.Parse(bcThang["x22"]) - double.Parse(bcThang["x28"]))}");
            bcThang.Add("m13cc33", $"{(double.Parse(bcThang["x23"]) - double.Parse(bcThang["x29"]))}");
            bcThang.Add("m13cc43", $"{(double.Parse(bcThang["x24"]) - double.Parse(bcThang["x30"]))}");
            bcThang.Add("m13cc53", $"{(double.Parse(bcThang["x25"]) - double.Parse(bcThang["x31"]))}");
            bcThang.Add("m13cc63", $"{(double.Parse(bcThang["x26"]) - double.Parse(bcThang["x32"]))}");

            /* Tỷ lệ % tăng giảm
                ,m13cc14 real not null default 0 /* Tổng lượt = 2+3 ((m13cc13/x27)*100)
                ,m13cc24 real not null default 0 /* Chi ngoại trú = (m13cc23/x28)*100
                ,m13cc34 real not null default 0 /* Chi nội trú = (m13cc33/x29)*100
                ,m13cc44 real not null default 0 /* Tổng lượt = 5+6 ((m13cc43/x30)*100)
                ,m13cc54 real not null default 0 /* Chi ngoại trú = (m13cc53/x31)*100
                ,m13cc64 real not null default 0 /* Chi nội trú = (m13cc63/x32)*100 */
            bcThang.Add("m13cc14", $"{Math.Round((double.Parse(bcThang["m13cc13"]) / double.Parse(bcThang["x27"])) * 100, 2)}");
            bcThang.Add("m13cc24", $"{Math.Round((double.Parse(bcThang["m13cc23"]) / double.Parse(bcThang["x28"])) * 100, 2)}");
            bcThang.Add("m13cc34", $"{Math.Round((double.Parse(bcThang["m13cc33"]) / double.Parse(bcThang["x29"])) * 100, 2)}");
            bcThang.Add("m13cc44", $"{Math.Round((double.Parse(bcThang["m13cc43"]) / double.Parse(bcThang["x30"])) * 100, 2)}");
            bcThang.Add("m13cc54", $"{Math.Round((double.Parse(bcThang["m13cc53"]) / double.Parse(bcThang["x31"])) * 100, 2)}");
            bcThang.Add("m13cc64", $"{Math.Round((double.Parse(bcThang["m13cc63"]) / double.Parse(bcThang["x32"])) * 100, 2)}");

            bcThang["timecreate"] = DateTime.Now.toTimestamp().ToString();
            bcThang["userid"] = idUser;
            bcThang["ma_tinh"] = maTinh;
            var bcThangPL = createBCThangPLDocx(dbConnect, idBaoCao, maTinh, mavung, bcThang["nam1"], bcThang["thang"]);
            bcThangPL["id"] = idBaoCao;
            dbConnect.Update("bcthangpldocx", bcThangPL, "replace");
            dbConnect.Update("bcthangdocx", bcThang, "replace");
            foreach (var v in bcThangPL) { if (bcThang.ContainsKey(v.Key) == false) { bcThang.Add(v.Key, v.Value); } }
            return bcThang;
        }

        private Dictionary<string, string> createBCThangPLDocx(dbSQLite dbConnect, string idBaoCao, string maTinh, string maVung, string namBaoCao, string thang)
        {
            var bcThangPL = new Dictionary<string, string>() { { "id", idBaoCao } };

            double so1 = 0; double so2 = 0;
            var tmpD = new Dictionary<string, string>();
            string tsql = string.Empty;
            string tmp = string.Empty;

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
                FROM thangb02chitiet WHERE id_bc='{idBaoCao}' AND (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%')
                    AND id2 IN (SELECT id FROM thangb02 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND tu_thang=1 AND nam={namBaoCao} LIMIT 1);";
            var b02TQ = dbConnect.getDataTable(tsql).AsEnumerable().ToList();
            if (b02TQ.Count() == 0) { throw new Exception("B02 Toàn Quốc không có dữ liệu phù hợp truy vấn; " + tsql); }
            var dataTinhB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == maTinh).FirstOrDefault() ?? throw new Exception("B02 không có dữ liệu tỉnh phù hợp truy vấn");
            var dataTQB02 = b02TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault() ?? throw new Exception("B02 không có dữ liệu toàn quốc phù hợp truy vấn");
            b02TQ = b02TQ.Where(p => p.Field<string>("ma_tinh") != "00").ToList(); /* Bỏ Toàn quốc ra khỏi danh sách */

            /* t5 = {Cột tyle_noitru, dòng MA_TINH=10} bảng B02_TOANQUOC */
            bcThangPL.Add("{t5}", dataTinhB02["tyle_noitru"].ToString());
            /* t6 = {Cột tyle_noitru, dòng MA_TINH=00} bảng B02_TOANQUOC */
            bcThangPL.Add("{t6}", dataTQB02["tyle_noitru"].ToString());
            /* t7 = {đoạn văn tùy thuộc t5> hay < t6. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThangPL.Add("{t7}", "bằng");
            so1 = (double)dataTinhB02["tyle_noitru"];
            so2 = (double)dataTQB02["tyle_noitru"];
            if (so1 > so2) { bcThangPL["{t7}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThangPL["{t7}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* t8={Sort cột G (TYLE_NOITRU) cao xuống thấp và lấy thứ tự}; */
            var sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
            int position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThangPL.Add("t8", position.ToString());
            /* t9 ={tính toán: total cột F (TONG_LUOT_NOI) chia cho Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG=mã vùng của tỉnh báo cáo}; */
            bcThangPL.Add("{t9}", "0");
            so2 = b02TQ.Where(row => row.Field<string>("ma_vung") == maVung).Sum(row => row.Field<long>("tong_luot"));
            if (so2 != 0)
            {
                so1 = b02TQ.Where(row => row.Field<string>("ma_vung") == maVung).Sum(row => row.Field<long>("tong_luot_noi"));
                bcThangPL["{t9}"] = ((so1 / so2) * 100).ToString();
            }
            /* t10 ={đoạn văn tùy thuộc t5> hay < t9. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThangPL.Add("{t10}", "bằng");
            so1 = (double)dataTinhB02["tyle_noitru"];
            so2 = double.Parse(bcThangPL["{t9}"]); bcThangPL["{t9}"] = bcThangPL["{t9}"].ToString();
            if (so1 > so2) { bcThangPL["{t10}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThangPL["{t10}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* X11= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort cột G (TYLE_NOITRU ) cao –thấp và lấy thứ tự} */
            sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == maVung)
                .OrderByDescending(row => row.Field<double>("tyle_noitru")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThangPL.Add("{t11}", position.ToString());

            /* t12 = Ngày điều trị bình quân t12={Cột H NGAY_DTRI_BQ , dòng MA_TINH=10}; */
            bcThangPL.Add("{t12}", dataTinhB02["ngay_dtri_bq"].ToString());
            /* t13 = Nbình quân toàn quốc t13={cột H NGAY_DTRI_BQ, dòng MA_TINH=00}; */
            bcThangPL.Add("{t13}", dataTQB02["ngay_dtri_bq"].ToString());
            /* t14 = Số chênh lệch t14={đoạn văn tùy thuộc t12> hay < t13. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThangPL.Add("{t14}", "bằng");
            so1 = (double)dataTinhB02["ngay_dtri_bq"];
            so2 = (double)dataTQB02["ngay_dtri_bq"];
            if (so1 > so2) { bcThangPL["{t14}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThangPL["{t14}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* t15 = xếp thứ so toàn quốc X15={Sort cột H (NGAY_DTRI_BQ) cao xuống thấp và lấy thứ tự}; */
            sortedRows = b02TQ.OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThangPL.Add("{t15}", position.ToString());
            /* t16 = Bình quân vùng X16 ={tính toán: A-Tổng ngày điều trị nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của cá tỉnh cùng mã vùng. A=Total(cột H (NGAY_DTRI_BQ) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
            bcThangPL.Add("{t16}", "0");
            so2 = b02TQ.Where(r => r.Field<string>("ma_vung") == maVung).Sum(r => r.Field<long>("tong_luot_noi"));
            if (so2 != 0)
            {
                so1 = b02TQ.Where(r => r.Field<string>("ma_vung") == maVung).Sum(r => (r.Field<double>("ngay_dtri_bq") * r.Field<long>("tong_luot_noi")));
                bcThangPL["{t16}"] = (so1 / so2).ToString();
            }
            /* t17 = Số chênh lệch t17 ={đoạn văn tùy thuộc t12> hay < t16. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            bcThangPL.Add("{t17}", "bằng");
            so1 = (double)dataTinhB02["ngay_dtri_bq"];
            so2 = double.Parse(bcThangPL["{t16}"]); bcThangPL["{t16}"] = bcThangPL["{t16}"].ToString();
            if (so1 > so2) { bcThangPL["{t17}"] = $"cao hơn {(so1 - so2).FormatCultureVN()}"; }
            else { if (so1 < so2) { bcThangPL["{t17}"] = $"thấp hơn {(so2 - so1).FormatCultureVN()}"; } }
            /* t18 = đứng thứ so với vùng t18 = {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột H (NGAY_DTRI_BQ) cao –thấp và lấy thứ tự} */
            sortedRows = b02TQ.Where(r => r.Field<string>("ma_vung") == maVung)
                .OrderByDescending(row => row.Field<double>("ngay_dtri_bq")).ToList();
            position = sortedRows.FindIndex(row => row.Field<string>("ma_tinh") == maTinh) + 1;
            bcThangPL.Add("{t18}", position.ToString());

            /* t19 = Chi bình quân chung t19={Cột I (CHI_BQ_CHUNG), dòng MA_TINH=10}; */
            tmpD = buildBCThangB02(19, "chi_bq_chung", "chi_bq_chung", "tong_luot", "tong_chi", maVung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
            /* t26 = Chi bình quân ngoại trú t26={Cột J (CHI_BQ_NGOAI), dòng MA_TINH=10}; */
            tmpD = buildBCThangB02(26, "chi_bq_ngoai", "chi_bq_chung", "tong_luot_ngoai", "tong_chi_ngoai", maVung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
            /* t33 = Chi bình quân nội trú t33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
            tmpD = buildBCThangB02(33, "chi_bq_noi", "chi_bq_chung", "tong_luot_noi", "tong_chi_noi", maVung, maTinh, dataTinhB02, dataTQB02, b02TQ);
            foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }

            /* ----- Dữ liệu t40 trở lên lọc dữ liệu tù B26 ------- */
            /* Bỏ qua các vùng */
            tmp = namBaoCao;
            if (thang.Length > 1) { tmp += thang; } else { tmp += $"0{thang}"; }
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
                FROM thangb26chitiet WHERE id_bc='{idBaoCao}' AND (ma_tinh <> '' AND ma_tinh NOT LIKE 'V%') AND id2 IN (SELECT id FROM thangb26 WHERE id_bc='{idBaoCao}' AND ma_tinh='00' AND (thoigian > {tmp}00 AND thoigian < {tmp}33) LIMIT 1);";
            var b26TQ = dbConnect.getDataTable(tsql).AsEnumerable().ToList();
            if (b26TQ.Count() > 0)
            {
                var dataTinhB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == maTinh).FirstOrDefault();
                if (dataTinhB26 == null) { return bcThangPL; }
                var dataTQB26 = b26TQ.Where(r => r.Field<string>("ma_tinh") == "00").FirstOrDefault();
                if (dataTQB26 == null) { return bcThangPL; }
                b26TQ = b26TQ.Where(p => p.Field<string>("ma_tinh") != "00").ToList(); /* Bỏ Toàn quốc ra khỏi danh sách */

                /* t40 = Bình quân xét nghiệm t40= {cột P (bq_xn) dòng có mã tỉnh = 10}; B26 */
                tmpD = buildBCThangB26(40, "bq_xn", "bq_xn_tang", dataTinhB26);
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
                /* t43 Bình quân CĐHA t43= {cột R(bq_cdha) dòng có mã tỉnh =10}; */
                tmpD = buildBCThangB26(43, "bq_cdha", "bq_cdha_tang", dataTinhB26);
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
                /* t46 Bình quân thuốc t46= {cột T(bq_thuoc) dòng có mã tỉnh =10}; */
                tmpD = buildBCThangB26(46, "bq_thuoc", "bq_thuoc_tang", dataTinhB26);
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
                /* t49 Bình quân chi phẫu thuật t49= {cột V(bq_pt) dòng có mã tỉnh =10}; */
                tmpD = buildBCThangB26(49, "bq_pt", "bq_pt_tang", dataTinhB26);
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
                /* t52 Bình quân chi thủ thuật t52= {cột X(bq_tt) dòng có mã tỉnh =10}; */
                tmpD = buildBCThangB26(52, "bq_tt", "bq_tt_tang", dataTinhB26);
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
                /* t55 Bình quân chi vật tư y tế t55= {cột Z(bq_vtyt) dòng có mã tỉnh =10}; */
                tmpD = buildBCThangB26(55, "bq_vtyt", "bq_vtyt_tang", dataTinhB26);
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
                /* t58 Bình quân chi tiền giường t58= {cột AB(bq_giuong) dòng có mã tỉnh =10}; */
                tmpD = buildBCThangB26(58, "bq_giuong", "bq_giuong_tang", dataTinhB26);
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }

                /* t61 Chỉ định xét nghiệm t61={cột AD, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
                tmpD = buildBCThangB26(61, "chi_dinh_xn", "chi_dinh_xn_tang", dataTinhB26, "người");
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
                /* t64 =  Chỉ định CĐHA t64={cột AF, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
                tmpD = buildBCThangB26(64, "chi_dinh_cdha", "chi_dinh_cdha_tang", dataTinhB26, "người");
                foreach (var d in tmpD) { bcThangPL.Add(d.Key, d.Value); }
            }
            return bcThangPL;
        }

        private Dictionary<string, string> buildBCThangB26(int iKey, string field1, string field2, DataRow row, string dvt = "đồng")
        {
            var d = new Dictionary<string, string>();
            string key1 = "{t" + iKey.ToString() + "}", key2 = "{t" + (iKey + 1).ToString() + "}", key3 = "{t" + (iKey + 2).ToString() + "}";
            /* t46 Bình quân cột [x] dòng có mã tỉnh = 10}; */
            var x = (double)row[field1]; if (iKey == 61 || iKey == 64) { x = x * 100; }
            d.Add(key1, Math.Round(x, 0).ToString());
            /* t47 số tương đối t47={nếu cột [x+1] dòng có mã tỉnh=10 là số dương, “tăng “ & cột [x+1] & “%”, không thì “giảm “ & cột [x+1] %}; */
            d.Add(key2, "bằng");
            var x1 = (double)row[field2]; /* s */
            if (x1 > 0) { d[key2] = $"tăng {x1.FormatCultureVN()}%"; }
            else { if (x1 < 0) { d[key2] = $"giảm {Math.Abs(x1).FormatCultureVN()}%"; } }
            /* t48 số tuyệt đối t48={nếu cột [x+1] là dương, “tăng “ & [cột [x] - (cột [x] / (cột [x+1] +100) *100 )] & “ đồng”, không thì “giảm “ & [cột [x]- (cột [x] / (cột [x+1]+100) *100 )] & “ đồng”} */
            d.Add(key3, "bằng");
            if (x1 > 0) { d[key3] = "tăng " + Math.Round(Math.Abs(x - (x / (x1 + 100) * 100)), 0).FormatCultureVN() + " " + dvt; }
            else { if (x1 < 0) { d[key3] = "giảm " + Math.Round(Math.Abs(x - (x / (x1 + 100) * 100)), 0).FormatCultureVN() + " " + dvt; } }
            return d;
        }

        private Dictionary<string, string> buildBCThangB02(int iKey, string fieldChiBQ, string fieldChiBQChung, string fieldTongLuotVung, string fieldTongChiVung, string mavung, string matinh, DataRow rowTinh, DataRow rowTQ, List<DataRow> data)
        {
            var d = new Dictionary<string, string>();
            var keys = new List<string>();
            for (int i = iKey; i <= (iKey + 6); i++) { keys.Add("{t" + i.ToString() + "}"); }
            /* t33 = Chi bình quân nội trú t33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
            d.Add(keys[0], rowTinh[fieldChiBQ].ToString()); /* "chi_bq_noi" */
            /* t34 = bình quân toàn quốc t34={cột K (CHI_BQ_NOI), dòng MA_TINH=00}; */
            d.Add(keys[1], rowTQ[fieldChiBQ].ToString());
            /* t35 = Số chênh lệch t35={đoạn văn tùy thuộc X33> hay < X34. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            d.Add(keys[2], "bằng");
            var so1 = double.Parse(d[keys[0]]);
            var so2 = double.Parse(d[keys[1]]);
            if (so1 > so2) { d[keys[2]] = $"cao hơn {Math.Round(so1 - so2, 0).FormatCultureVN()}"; }
            else { if (so1 < so2) { d[keys[2]] = $"thấp hơn {Math.Round(so2 - so1, 0).FormatCultureVN()}"; } }
            /* t36= xếp thứ so toàn quốc t36={Sort cột K CHI_BQ_NOI cao xuống thấp và lấy thứ tự}; */
            d.Add(keys[3], getPosition("", matinh, fieldChiBQ, data));
            /*** Vùng
             = SUM(tong_chi)/SUM(tong_luot)
             */
            /* t37 = Bình quân vùng X37={tính toán: A-Tổng chi nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của các tỉnh cùng mã vùng. A=Total  (cột K (CHI_BQ_NOI) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
            d.Add(keys[4], "0");
            so2 = data.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<long>(fieldTongLuotVung));
            if (so2 != 0)
            {
                so1 = data.Where(r => r.Field<string>("ma_vung") == mavung).Sum(r => r.Field<double>(fieldTongChiVung));
                d[keys[4]] = Math.Round(so1 / so2, 0).ToString();
            }
            /* t38 = số chênh lệch t38 ={đoạn văn tùy thuộc t33 > hay < t37. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
            d.Add(keys[5], "bằng");
            so1 = double.Parse(d[keys[0]]);
            so2 = double.Parse(d[keys[4]]);
            if (so1 > so2) { d[keys[5]] = $"cao hơn {Math.Round(so1 - so2, 0).FormatCultureVN()}"; }
            else { if (so1 < so2) { d[keys[5]] = $"thấp hơn {Math.Round(so2 - so1, 0).FormatCultureVN()}"; } }
            /* t39 đứng thứ so với vùng t39= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột K (CHI_BQ_NOI) cao –thấp và lấy thứ tự} */
            d.Add(keys[6], getPosition(mavung, matinh, fieldChiBQ, data));
            return d;
        }

        private List<string> createFileBCThang(string idBaoCao, string matinh, dbSQLite dbBCThang = null)
        {
            if (dbBCThang == null) { dbBCThang = BuildDatabase.getDataBCThang(matinh); }
            var rs = new List<string>();
            string pathFileTemplate = Path.Combine(AppHelper.pathAppData, "bcThang.docx");
            if (!System.IO.File.Exists(pathFileTemplate)) { throw new Exception("Không tìm thấy tập tin mẫu báo cáo 'bcThang.docx' trong thư mục App_Data"); }
            var bcThangExport = new Dictionary<string, string>();
            var outputPath = Path.Combine(AppHelper.pathAppData, "bcThang", $"tinh{matinh}");
            if (!Directory.Exists(outputPath)) { Directory.CreateDirectory(outputPath); }
            var outputFile = Path.Combine(outputPath, $"bcThang_{idBaoCao}.docx");

            string valReplace = "", tmp = "";
            var data = dbBCThang.getDataTable($"SELECT * FROM bcthangdocx WHERE id='{idBaoCao}'");
            if (data.Rows.Count > 0)
            {
                DataRow r = data.Rows[0];
                foreach (DataColumn c in data.Columns) { bcThangExport.Add("{" + c.ColumnName + "}", $"{r[c.ColumnName]}"); }
            }
            data = dbBCThang.getDataTable($"SELECT * FROM bcthangpldocx WHERE id='{idBaoCao}';");
            if (data.Rows.Count > 0)
            {
                DataRow r = data.Rows[0];
                foreach (DataColumn c in data.Columns)
                {
                    valReplace = "{" + c.ColumnName + "}";
                    if (bcThangExport.ContainsKey(valReplace) == false) { bcThangExport.Add(valReplace, $"{r[c.ColumnName]}"); }
                }
            }
            rs.Add(outputFile);

            var dbImport = dbBCThang.getPathDataFile().StartsWith(AppHelper.pathTemp) ? dbBCThang : BuildDatabase.getDataImportBCThang(matinh);
            var outFile = Path.Combine(AppHelper.pathApp, "App_Data", "bcThang", $"tinh{matinh}", $"bcThang_{idBaoCao}_pl.xlsx");
            var idBC = idBaoCao.sqliteGetValueField();
            var dmVung = new Dictionary<string, string>();
            data = dbBCThang.getDataTable($"SELECT DISTINCT ma_tinh, ten_tinh FROM thangpl04a WHERE id_bc='{idBC}' AND ma_tinh LIKE 'V%'");
            foreach (DataRow r in data.Rows) { dmVung.Add($"{r[0]}", $"{r[1]}"); }
            /* Lấy năm tháng báo cáo */
            string nam = bcThangExport["{nam1}"]; string thang = bcThangExport["{thang}"];
            if (nam == "")
            {
                data = dbBCThang.getDataTable($"SELECT nam1, thang FROM bcthangdocx WHERE id = '{idBaoCao}'");
                if (data.Rows.Count > 0)
                {
                    nam = $"{data.Rows[0][0]}";
                    thang = $"{data.Rows[0][1]}";
                }
                else
                {
                    if (dbBCThang.tableExist("thangb02"))
                    {
                        data = dbBCThang.getDataTable($"SELECT den_thang, nam FROM thangb02 WHERE id_bc='{idBaoCao}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                        if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
                        if (nam == "")
                        {
                            data = dbBCThang.getDataTable($"SELECT den_thang, nam FROM thangb01 WHERE id_bc='{idBaoCao}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                            if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
                        }
                        if (nam == "")
                        {
                            data = dbBCThang.getDataTable($"SELECT den_thang, nam FROM thangb04 WHERE id_bc='{idBaoCao}' ORDER BY nam DESC, den_thang DESC LIMIT 1;");
                            if (data.Rows.Count > 0) { nam = $"{data.Rows[0][1]}"; thang = $"{data.Rows[0][0]}"; }
                        }
                    }
                    else { nam = "0"; thang = "0"; }
                }
            }
            /* Tạo phụ lục báo cáo */
            /* - Lấy danh sách mã cấp trên */
            var dataCSKCB = AppHelper.dbSqliteMain.getDataTable($"SELECT id, CASE WHEN macaptren = '' THEN id ELSE macaptren END AS macaptren, ten FROM dmcskcb WHERE ma_tinh='{matinh}';");
            var matchMaCapTren = new Dictionary<string, string>();
            var dsMaCapTren = new Dictionary<string, string>();
            for (int i = 0; i < dataCSKCB.Rows.Count; i++)
            {
                tmp = $"{dataCSKCB.Rows[i][0]}";
                matchMaCapTren.Add(tmp, $"{dataCSKCB.Rows[i][1]}");
                if (tmp == $"{dataCSKCB.Rows[i][1]}") { dsMaCapTren.Add(tmp, $"{dataCSKCB.Rows[i][2]}"); }
            }
            /* -- */
            var PL01 = createPL01(dbBCThang, idBaoCao, matinh, nam, thang, matchMaCapTren, dsMaCapTren);
            var PL02a = createPL02(dbBCThang, idBaoCao, matinh, "PL02a", dmVung);
            var PL02b = createPL02(dbBCThang, idBaoCao, matinh, "PL02b", dmVung);
            var PL02c = createPL02c(dbImport, idBaoCao, matinh, long.Parse(nam), 1, long.Parse(thang), matchMaCapTren);
            data = createPL02(dbBCThang, idBaoCao, matinh, "pl03a2", dmVung);
            var PL03a = createPL03a(dbBCThang, idBaoCao, "PL03a", thang, PL02a, data, dmVung, matchMaCapTren);
            var PL03b = createPL03b(dbBCThang, idBaoCao, "PL03b", PL02b, long.Parse(nam), matchMaCapTren);
            var PL03c = createPL03c(dbImport, idBaoCao, matinh, thang, matchMaCapTren);
            var PL04a = createPL04a(dbBCThang, idBaoCao, matinh, dmVung);
            var PL04b = createPL04b(dbBCThang, idBaoCao, matinh, thang, matchMaCapTren);
            exportPhuLucbcThang(idBaoCao, outFile, PL01, PL02a, PL02b, PL02c, PL03a, PL03b, PL03c, PL04a, PL04b);
            rs.Add(outFile);

            /* Cập nhật lại x39 */
            var lx39 = new List<string>();
            for (int i = 0; i < (PL01.Rows.Count > 5 ? 5 : PL01.Rows.Count); i++) { lx39.Add($"{PL01.Rows[i][1]}: {PL01.Rows[i][4]}"); }
            try { dbBCThang.Execute($"UPDATE bcthangdocx SET x39='{string.Join(", ", lx39)}' WHERE id='{idBC}';"); } catch { }
            bcThangExport["{x39}"] = string.Join(", ", lx39);
            /* Kiểm tra x1, x33 - x38 */
            if (bcThangExport["{x1}"] + bcThangExport["{x33}"] + bcThangExport["{x34}"] + bcThangExport["{x35}"] + bcThangExport["{x36}"] + bcThangExport["{x37}"] + bcThangExport["{x38}"] == "")
            {
                /* Tìm thông tin từ các báo cáo tháng trước có không để điền vào */
            }
            /* Export bcthang.docx */
            using (var fileStream = new FileStream(pathFileTemplate, FileMode.Open, FileAccess.Read))
            {
                var document = new NPOI.XWPF.UserModel.XWPFDocument(fileStream);
                foreach (var paragraph in document.Paragraphs)
                {
                    foreach (var run in paragraph.Runs)
                    {
                        tmp = run.ToString();
                        MatchCollection matches = Regex.Matches(tmp, "{[a-z0-9_]+}", RegexOptions.IgnoreCase);
                        foreach (System.Text.RegularExpressions.Match match in matches)
                        {
                            valReplace = bcThangExport.getValue(match.Value, "");
                            if (match.Value.StartsWith("{t") || match.Value.StartsWith("{x"))
                            {
                                if (valReplace.isNumberUS()) { valReplace = valReplace.FormatCultureVN(); }
                            }
                            tmp = tmp.Replace(match.Value, valReplace);
                        }
                        run.SetText(tmp, 0);
                    }
                }
                /* Thay thế trong các bảng */
                foreach (var table in document.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.GetTableCells())
                        {
                            foreach (var paragraph in cell.Paragraphs)
                            {
                                foreach (var run in paragraph.Runs)
                                {
                                    tmp = run.ToString();
                                    MatchCollection matches = Regex.Matches(tmp, "{[a-z0-9_]+}", RegexOptions.IgnoreCase);
                                    foreach (System.Text.RegularExpressions.Match match in matches)
                                    {
                                        valReplace = bcThangExport.getValue(match.Value, "");
                                        if (valReplace.isNumberUS()) { valReplace = valReplace.FormatCultureVN(); }
                                        tmp = tmp.Replace(match.Value, valReplace);
                                    }
                                    run.SetText(tmp, 0);
                                }
                            }
                        }
                    }
                }
                if (System.IO.File.Exists(outputFile)) { System.IO.File.Delete(outputFile); }
                using (var stream = new FileStream(outputFile, FileMode.Create, FileAccess.Write)) { document.Write(stream); }
            }
            return rs;
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
                var dbBaoCao = BuildDatabase.getDataBCThang(idtinh);
                if (Request.getValue("mode") == "update")
                {
                    var timeStart = DateTime.Now;
                    item = new Dictionary<string, string>() {
                        { "x1", Request.getValue("x1") },
                        { "x33", Request.getValue("x33") },
                        { "x34", Request.getValue("x34") },
                        { "x35", Request.getValue("x35") },
                        { "x36", Request.getValue("x36") },
                        { "x37", Request.getValue("x37") },
                        { "x38", Request.getValue("x38") }
                    };
                    dbBaoCao.Update("bcthangdocx", item, $"id='{id.sqliteGetValueField()}'");
                    tsql = $"SELECT * FROM bcthangdocx WHERE id='{id.sqliteGetValueField()}' LIMIT 1";
                    var data = dbBaoCao.getDataTable(tsql);
                    dbBaoCao.Close();
                    if (data.Rows.Count == 0)
                    {
                        ViewBag.Error = $"Báo cáo tuần có ID '{id}' thuộc tỉnh có mã '{idtinh}' không tồn tại hoặc đã bị xoá khỏi hệ thống";
                        return View();
                    }
                    var listFile = createFileBCThang(id, idtinh, dbBaoCao);
                    string fileZip = Path.Combine(AppHelper.pathAppData, "bcThang", $"tinh{idtinh}", $"bcThang_{id}.zip");
                    if (System.IO.File.Exists(fileZip)) { System.IO.File.Delete(fileZip); }
                    AppHelper.zipAchive(fileZip, listFile);
                    return Content($"Lưu thành công ({timeStart.getTimeRun()})".BootstrapAlter());
                }
                tsql = $"SELECT * FROM bcthangdocx WHERE id='{id.sqliteGetValueField()}'";
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
                    var dbBcThang = BuildDatabase.getDataBCThang(matinh);
                    var where = $"WHERE timecreate >= {time1.toTimestamp()} AND timecreate < {time2.AddDays(1).toTimestamp()}";
                    var tmp = $"{Session["nhom"]}";
                    if (tmp != "0" && tmp != "1") { where += $" AND userid='{Session["iduser"]}'"; }
                    var tsql = $"SELECT datetime(timecreate, 'auto', '+7 hour') AS ngayGM7, thang, nam1,id,ma_tinh,userid FROM bcthangdocx {where} ORDER BY timecreate DESC";
                    ViewBag.data = dbBcThang.getDataTable(tsql);
                    dbBcThang.Close();
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
                    foreach (string id in lid) { DeleteBcThang(id, true); }
                    return Content($"Xoá thành công báo cáo có ID '{string.Join(", ", lid)}' ({timeStart.getTimeRun()})".BootstrapAlter());
                }
            }
            catch (Exception ex) { return Content(ex.getErrorSave().BootstrapAlter("warning")); }
            return View();
        }

        private void DeleteBcThang(string id, bool throwEx = false)
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
                foreach (var f in folder.GetFiles($"bcThang_{id}*.*")) { try { f.Delete(); } catch { } }
                foreach (var f in folder.GetFiles($"id{id}*.*")) { try { f.Delete(); } catch { } }
            }
            /* Xoá trong cơ sở dữ liệu */
            var db = BuildDatabase.getDataBCThang(idtinh);
            try
            {
                var idBaoCao = id.sqliteGetValueField();
                var listTablePL = new List<string>() { "thangpl01", "thangpl02a", "thangpl02b", "thangpl03a", "thangpl03b", "thangpl04a", "thangpl04b" };
                var tsql = new List<string>() { $"DELETE FROM bcThangdocx WHERE id='{idBaoCao}';", $"DELETE FROM bcThangpldocx WHERE id='{idBaoCao}';" };
                foreach (var t in listTablePL) { tsql.Add($"DELETE FROM {t} WHERE id_bc='{idBaoCao}';"); }
                db.Execute(string.Join(" ", tsql));
                db.Close();
                db = BuildDatabase.getDataImportBCThang(idtinh);
                listTablePL = new List<string>() { "thangb01", "thangb02", "thangb04", "thangb21", "thangb26", "thangb01chitiet", "thangb02chitiet", "thangb04chitiet", "thangb21chitiet", "thangb26chitiet" };
                tsql = new List<string>();
                foreach (var t in listTablePL) { tsql.Add($"DELETE FROM {t} WHERE id_bc='{idBaoCao}';"); }
                db.Execute(string.Join(" ", tsql));
                db.Close();
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