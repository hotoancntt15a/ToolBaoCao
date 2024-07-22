using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ToolBaoCao
{
    public static class BuildDatabase
    {
        public static void buildData(this dbSQLite connect) {
            var tsqlInsert = new List<string>();
            var tsqlCreate = new List<string>();
            var tsql = "";
            var tables = connect.getAllTables();
            /**
             Nhóm quản lý web
             */
            if (tables.Contains("taikhoan") == false)
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS taikhoan(iduser TEXT NOT NULL Primary Key, mat_khau TEXT NOT NULL, ten_hien_thi TEXT NOT NULL, gioi_tinh TEXT NOT NULL DEFAULT '', ngay_sinh TEXT NOT NULL Default '', email TEXT NOT NULL Default '', dien_thoai TEXT NOT NULL Default '', dia_chi TEXT NOT NULL Default '', hinh_dai_dien TEXT NOT NULL Default '', idtinh text not null default '', ghi_chu TEXT NOT NULL DEFAULT '', nhom INTEGER NOT NULL default -1, locked INTEGER NOT NULL default 0, time_create double not null default 0, time_last_login double not null default 0);");
            }
            if (tables.Contains("dmtinh") == false)
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS dmtinh(id text primary key, ten text not null default '', tt integer not null default 999, ghichu text not null default '');");
            }
            if (tables.Contains("dmnhom") == false)
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS dmnhom(id INTEGER primary key, ten text not null default '', idwmenus text not null default '', ghichu text not null default '');");
            }
            if (tables.Contains("wmenu") == false)
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS wmenu(id INTEGER primary key, title text not null default '', link text not null default '', idfather integer not null default -1, path text not null default '', postion INTEGER NOT NULL DEFAULT 0, note text not null default '');");
            }
            /* Tạo cơ sở dữ liệu */
            try { tsql = string.Join(" ", tsqlCreate); connect.Execute(tsql); } catch (Exception ex) { ex.saveError(tsql); }
            /* if (tsqlCreate.Count > 0) { foreach (var v in tsqlCreate) { try { connect.Execute(v); } catch (Exception ex) { ex.saveError(v); } } } */
            /* Kiểm tra xem đã có dữ liệu chưa */
            try
            {
                var items = connect.getDataTable("SELECT iduser FROM taikhoan LIMIT 1");
                if (items.Rows.Count == 0)
                {
                    /* Thêm tài khoản admin mặc định */
                    tsqlInsert.Add($"INSERT INTO taikhoan (iduser, mat_khau, ten_hien_thi, gioi_tinh, ngay_sinh, email, dien_thoai, dia_chi, hinh_dai_dien, ghi_chu, time_create, nhom)VALUES ('admin','{"0914272795".GetMd5Hash()}','Quản trị hệ thống','','{DateTime.Now:dd/MM/yyyy}','hotoancntt15a@gmail.com','0914272795','TP Lào Cai, Tỉnh Lào Cai','','','{DateTime.Now.toTimestamp()}', 0);");
                }
                if (tsqlInsert.Count > 0) { foreach (var v in tsqlInsert) { try { connect.Execute(v); } catch (Exception ex) { ex.saveError(v); } } }
            }
            catch (Exception er) { er.saveError(); }
        }
        public static void buildDataCongViec(this dbSQLite connect)
        {
            var tsqlCreate = new List<string>();
            var tsql = "";
            var tables = connect.getAllTables();
            /**
             Cơ sở công việc
             */
            if (tables.Contains("sheetpl01") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS sheetpl01 (id INTEGER primary key AUTOINCREMENT, 
                id_bc text not null, /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                idtinh text not null, /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ma_tinh text not null default '', /* Mã tỉnh Cột A, B02 */            
                ten_tinh text not null default '', /* Tên tỉnh Cột B, B02 */
                mavung text not null default '', /* Mã vùng 0,1,2,3,4... cột C , B02 */
                tyle_noitru real not null default 0, /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú, B02 */
                ngay_dtri_bq real not null default 0, /*	Ngày điều trị BQ, vd 6,42, DVT: ngày; Lấy từ cột H: NGAY ĐT_BQ, B02 */
                chi_bq_chung real not null default 0, /* Chi bình quan chung lượt KCB ĐVT ( đồng)	Cột I, B02 */
                chi_bq_ngoai real not null default 0, /* Chi bình quân ngoại trú/lượt KCB ngoại trú (đồng); Cột J, B02 */
                chi_bq_noi real not null default 0, /* Như trên nhưng với nội trú	Cột K, B02 */
                user_id text not null default '', /* Lưu ID của người dùng	 */
                user_name text not null default '' /* Lưu tên đăng nhập của người dùng	 */
                );"
                );
            }
            if (tables.Contains("sheetpl02") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS sheetpl02 (id INTEGER primary key AUTOINCREMENT, 
                id_bc text not null, /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                idtinh text not null, /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ma_tinh text not null, /* Mã tỉnh Cột A, B02 */            
                ten_tinh text not null, /* Tên tỉnh Cột B, B02 */
                chi_bq_xn real not null default 0, /* chi BQ Xét nghiệm; đơn vị tính : đồng	Lấy từ B04 . Cột D */
                chi_bq_cdha real not null default 0, /* chi BQ Chẩn đoán hình ảnh; Lấy từ B04. Cột E */
                chi_bq_thuoc real not null default 0, /* chi BQ thuốc; Lấy từ B04. Cột F
                chi_bq_pttt real not null default 0, /* chi BQ phẫu thuật thủ thuật	Lấy từ B04. Cột G */
                chi_bq_vtyt real not null default 0, /* chi BQ vật tư y tế; Lấy từ B04. Cột H */
                chi_bq_giuong real not null default 0, /* chi BQ tiền giường; Lấy từ B04. Cột I */
                ngay_tt_bq text not null default '', /* Ngày thanh toán bình quân; Lấy từ B04. Cột J */
                user_id text not null default '', /* Lưu ID của người dùng	 */
                user_name text not null default '' /* Lưu tên đăng nhập của người dùng	 */
                );"
                );
            }
            if (tables.Contains("sheetpl03") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS sheetpl03 (id INTEGER primary key AUTOINCREMENT, 
                id_bc text not null, /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                idtinh text not null, /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ma_cskcb text not null, /* Mã cơ sơ KCB, có chứa cả mã toàn quốc:00, mã vùng V1, mã tỉnh 10 và mã CSKCB ví dụ 10061; Ngoài 3 dòng đầu lấy từ bảng lưu thông tin Sheet 1; Các dòng còn lại lấy từ các cột A Excel B02 */
                ten_cskcb text not null default '', /* Tên cskcb, ghép hạng BV vào đầu chuỗi tên CSKCB	Côt B */
                tyle_noitru real not null default 0, /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú */
                ngay_dtri_bq real not null default 0, /* Ngày điều trị BQ, vd 6,42, DVT: NGÀY; Lấy từ cột H: NGAY ĐT_BQ */
                chi_bq_chung real not null default 0, /* Chi bình quan chung lượt KCB ĐVT đồng; Cột I B02 */
                chi_bq_ngoai real not null default 0, /* Chi bình quân ngoại trú/lượt KCB ngoại trú	Cột J B02 */
                chi_bq_noi real not null default 0, /* Như trên nhưng với nội trú; Cột K B02 */
                user_id text not null default '', /* Lưu ID của người dùng	 */
                user_name text not null default '' /* Lưu tên đăng nhập của người dùng	 */
                );"
                );
            }
            /* B02. Thống kê KCB (Tháng) */
            if (tables.Contains("b02") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b02 (id INTEGER primary key AUTOINCREMENT, 
                ma_tinh text not null,
                ma_loai_kcb text not null,
                tu_thang integer not null default 0,
                den_thang integer not null default 0,
                nam integer not null default 0,
                loai_bv integer not null default 0,
                kieubv integer not null default 0,
                loaick integer not null default 0,
                hang_bv integer not null default 0,
                tuyen integer not null default 0,
                cs integer not null default 0,
                user_id text not null default ''
                );"
                );
            }
            if (tables.Contains("b02chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b02chitiet (id INTEGER primary key AUTOINCREMENT,
                id2 integer not null,
                ma_tinh text not null,
                ten_tinh text not null default '',
                ma_cskcb text not null default '',
                ten_cskcb text not null default '',
                ma_vung text not null default '',
                tong_luot integer not null default 0,
                tong_luot_ngoai integer not null default 0,
                tong_luot_noi integer not null default 0,
                tyle_noitru real not null default 0,
                ngay_dtri_bq real not null default 0, 
                chi_bq_chung real not null default 0,
                chi_bq_ngoai real not null default 0,
                chi_bq_noi real not null default 0,
                tong_chi real not null default 0,
                ty_trong real not null default 0,
                tong_chi_ngoai real not null default 0,
                ty_trong_kham real not null default 0,
                tong_chi_noi real not null default 0,
                ty_trong_giuong real not null default 0,
                t_bhtt real not null default 0,
                t_bhtt_noi real not null default 0,
                t_bhtt_ngoai real not null default 0                
                );"
                );
            }

            /* B04. Thống kê chi bình quân (Tháng) */
            if (tables.Contains("b04") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b04 (id INTEGER primary key AUTOINCREMENT, 
                ma_tinh text not null,
                tu_thang integer not null default 0,
                den_thang integer not null default 0,
                nam integer not null default 0,
                ma_loai_kcb integer not null default 0,
                loai_bv integer not null default 0,
                hang_bv integer not null default 0,
                tuyen integer not null default 0,
                kieubv integer not null default 0,
                loaick integer not null default 0,
                cs integer not null default 0,
                user_id text not null default ''
                );"
                );
            }
            if (tables.Contains("b04chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b04chitiet (id INTEGER primary key AUTOINCREMENT,
                id2 integer not null,
                ma_tinh text not null,
                ten_tinh text not null default '',
                ma_cskcb text not null default '',
                ten_cskcb text not null default '',
                chi_bq_luotkcb real not null default 0, 
                bq_xn real not null default 0,
                bq_cdha real not null default 0,
                bq_thuoc real not null default 0,
                bq_ptt real not null default 0,
                bq_vtyt real not null default 0,
                bq_giuong real not null default 0,
                ngay_ttbq real not null default 0,
                ma_vung text not null default ''
                );"
                );
            }

            /* B26. Thống kê gia tăng chi phí KCB BHYT theo NĐ75 (theo ngày nhận) */
            if (tables.Contains("b26") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b26 (id INTEGER primary key AUTOINCREMENT, 
                ma_tinh text not null,
                loai_kcb text not null default '',
                thoigian integer not null default 0,
                loai_bv integer not null default 0,
                kieubv integer not null default 0,
                loaick integer not null default 0,
                hang_bv integer not null default 0,
                tuyen integer not null default 0,
                loai_so_sanh text not null default '',
                cs integer not null default 0,
                user_id text not null default ''
                );"
                );
            }
            if (tables.Contains("b26chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b26chitiet (id INTEGER primary key AUTOINCREMENT,
                id2 integer not null,
                ma_tinh text not null,
                ten_tinh text not null default '',
                ma_cskcb text not null default '',
                ten_cskcb text not null default '',
                vitri_chibq integer not null default 0,
                vitri_tyle_noitru integer not null default 0,
                vitri_tlxn integer not null default 0,
                vitri_tlcdha integer not null default 0,
                tytrong real not null default 0,
                chi_bq_chung real not null default 0,
                chi_bq_chung_tang real not null default 0,
                tyle_noitru real not null default 0,
                tyle_noitru_tang real not null default 0,
                lan_kham_bq real not null default 0,
                lan_kham_bq_tang real not null default 0,
                ngay_dtri_bq real not null default 0,
                ngay_dtri_bq_tang real not null default 0,
                bq_xn real not null default 0,
                bq_xn_tang real not null default 0,
                bq_cdha real not null default 0,
                bq_cdha_tang real not null default 0,
                bq_thuoc real not null default 0,
                bq_thuoc_tang real not null default 0,
                bq_pt real not null default 0,
                bq_pt_tang real not null default 0,
                bq_tt real not null default 0,
                bq_tt_tang real not null default 0,
                bq_vtyt real not null default 0,
                bq_vtyt_tang real not null default 0,
                bq_giuong real not null default 0,
                bq_giuong_tang real not null default 0,
                chi_dinh_xn real not null default 0,
                chi_dinh_xn_tang real not null default 0,
                chi_dinh_cdha real not null default 0,
                chi_dinh_cdha_tang real not null default 0,
                ma_vung text not null default ''
                );"
                );
            }
            /* Tạo cơ sở dữ liệu */
            try { tsql = string.Join(" ", tsqlCreate); connect.Execute(tsql); } catch (Exception ex) { ex.saveError(tsql); }
            /* if (tsqlCreate.Count > 0) { foreach (var v in tsqlCreate) { try { connect.Execute(v); } catch (Exception ex) { ex.saveError(v); } } } */
        }
    }
}