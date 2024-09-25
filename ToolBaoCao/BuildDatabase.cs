using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ToolBaoCao
{
    public static class BuildDatabase
    {
        public static string sqliteGetValueField(this string value) => value.Replace("'", "''");
        public static string SQLiteLike(this string value, string fieldName)
        {
            if (string.IsNullOrEmpty(value)) { return ""; }
            if (Regex.IsMatch(value, "[*%_?]+") == false) { return $"{fieldName} = '{value.sqliteGetValueField()}'"; }
            if (value.Contains("*")) { value = value.Replace("*", "%"); }
            if (value.Contains("?")) { value = value.Replace("?", "_"); }
            value = Regex.Replace(value, "[%]+", "%");
            return $"{fieldName} LIKE '{value.sqliteGetValueField()}'";
        }

        public static string SQLiteLike(this string value, List<string> fieldNames)
        {
            if (string.IsNullOrEmpty(value)) { return ""; }
            var ls = new List<string>();
            if (Regex.IsMatch(value, "[*%_?]+") == false)
            {
                value = value.sqliteGetValueField();
                foreach (var v in fieldNames) ls.Add($"{v}='{value}'");
                return string.Join(" or ", ls);
            }
            if (value.Contains("*")) { value = value.Replace("*", "%"); }
            if (value.Contains("?")) { value = value.Replace("?", "_"); }
            value = (Regex.Replace(value, "[%]+", "%")).sqliteGetValueField();
            foreach (var v in fieldNames) ls.Add($"{v} LIKE '{value}'");
            return string.Join(" OR ", ls);
        }
        public static void buildDataMain(this dbSQLite connect)
        {
            var tsqlInsert = new List<string>();
            var tsqlCreate = new List<string>();
            var tsql = "";
            var tables = connect.getAllTables();
            /** Nhóm quản lý web */
            if (tables.Contains("taikhoan") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS taikhoan (
                    iduser text NOT NULL PRIMARY KEY,
                    mat_khau text NOT NULL,
                    ten_hien_thi text NOT NULL,
                    gioi_tinh text NOT NULL DEFAULT '',
                    ngay_sinh text NOT NULL DEFAULT '',
                    email text NOT NULL DEFAULT '',
                    dien_thoai text NOT NULL DEFAULT '',
                    dia_chi text NOT NULL DEFAULT '',
                    hinh_dai_dien text NOT NULL DEFAULT '',
                    idtinh text NOT NULL DEFAULT '',
                    ghi_chu text NOT NULL DEFAULT '',
                    vitrilamviec text NOT NULL DEFAULT '',
                    nhom INTEGER NOT NULL DEFAULT - 1,
                    locked INTEGER NOT NULL DEFAULT 0,
                    time_create double NOT NULL DEFAULT 0,
                    time_last_login double NOT NULL DEFAULT 0);");
            }
            if(tables.Contains("logintime"))
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS logintime (iduser text NOT NULL PRIMARY KEY, timelogin integer NOT NULL);");
            }
            if (tables.Contains("dmtinh") == false)
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS dmtinh(id text primary key, ten text not null default '', tt integer not null default 999, ghichu text not null default '');");
            }
            if (tables.Contains("dmnhom") == false)
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS dmnhom(id INTEGER primary key, ten text not null default '', idwmenu text not null default '', ghichu text not null default '');");
            }
            if (tables.Contains("wmenu") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS wmenu (
                    id integer PRIMARY KEY,
                    title text NOT NULL DEFAULT '',
                    link text NOT NULL DEFAULT '',
                    idfather integer NOT NULL DEFAULT - 1,
                    paths text NOT NULL DEFAULT '',
                    postion integer NOT NULL DEFAULT 0,
                    note text NOT NULL DEFAULT '',
                  css text NOT NULL DEFAULT '' 
                  );");
            }
            if(tables.Contains("dmcskcb") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS dmcskcb(
                  id text NOT NULL PRIMARY KEY,
                  ten text NOT NULL DEFAULT '',
                  tuyencmkt text NOT NULL DEFAULT '',
                  hangbv text NOT NULL DEFAULT '',
                  loaibv text NOT NULL DEFAULT '',
                  tenhuyen text NOT NULL DEFAULT '',
                  donvi text NOT NULL DEFAULT '',
                  madinhdanh text NOT NULL DEFAULT '',
                  macaptren text NOT NULL DEFAULT '',
                  diachi text NOT NULL DEFAULT '',
                  ttduyet text NOT NULL DEFAULT '',
                  hieuluc text NOT NULL DEFAULT '',
                  tuchu text NOT NULL DEFAULT '',
                  trangthai text NOT NULL DEFAULT '',
                  hangdv text NOT NULL DEFAULT '',
                  hangthuoc text NOT NULL DEFAULT '',
                  dangkykcb text NOT NULL DEFAULT '',
                  hinhthuctochuc text NOT NULL DEFAULT '',
                  hinhthucthanhtoan text NOT NULL DEFAULT '',
                  ngaycapma text NOT NULL DEFAULT '',
                  kcb text NOT NULL DEFAULT '',
                  ngayngunghd text NOT NULL DEFAULT '',
                  kt7 text NOT NULL DEFAULT '',
                  kcn text NOT NULL DEFAULT '',
                  knl text NOT NULL DEFAULT '',
                  cpdtt43 text NOT NULL DEFAULT '',
                  slthedacap integer NOT NULL DEFAULT 0,
                  donvichuquan text NOT NULL DEFAULT '',
                  mota text NOT NULL DEFAULT '',
                  loaichuyenkhoa text NOT NULL DEFAULT '',
                  ngaykyhopdong text NOT NULL DEFAULT '',
                  ngayhethieuluc text NOT NULL DEFAULT '',
                  ma_tinh text NOT NULL,
                  ma_huyen text NOT NULL DEFAULT '',
                  userid text NOT NULL DEFAULT '');");
            }
            tsqlCreate.Add("CREATE INDEX IF NOT EXISTS index_dmcskcb_ma_tinh ON dmcskcb (ma_tinh);");
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
        /* Cơ sở dữ liệu quản lý theo dõi user đang online */
        public static dbSQLite getDBUserOnline()
        {
            string pathData = AppHelper.pathApp + "App_Data\\useronline.db";
            dbSQLite db = new dbSQLite(pathData);
            if (File.Exists(pathData) == false)
            {
                try
                {
                    db.Execute(@"CREATE TABLE IF NOT EXISTS useronline (
                        userid text NOT NULL,
                        time1 INTEGER NOT NULL DEFAULT 0,
                        time2 INTEGER NOT NULL DEFAULT 0,
                        ten_hien_thi text NOT NULL DEFAULT '',
                        ip text NOT NULL DEFAULT '',
                        [local] text NOT NULL DEFAULT '', PRIMARY KEY (userid, ip));");
                }
                catch { }
            }
            return db;
        }

        /**
         * Dữ liệu làm việc chính 
         */
        public static void buildDataWork(this dbSQLite dbConnect)
        {
            var tables = dbConnect.getAllTables();
            /* Các bảng Import */
            dbConnect.CreateImportBcTuan(tables);
            /* Các bảng phục lục công việc */
            dbConnect.CreatePhucLucBcTuan(tables);
            dbConnect.CreateBcTuan(tables);
            dbConnect.Execute(@"CREATE TABLE IF NOT EXISTS dutoangiao (so_kyhieu_qd text not null default ''
                  ,tong_dutoan real not null default 0
                  ,iduser text not null default ''
                  ,idtinh text not null default ''
                  ,idhuyen text not null default ''
                  ,namqd integer not null default 0
                  ,PRIMARY KEY (namqd,idtinh,idhuyen));");
            /* Tạo bảng quản lý các tiến trình */
        }

        public static void CreateTableProcess(this dbSQLite dbConnect)
        {
            dbConnect.Execute(@"CREATE TABLE IF NOT EXISTS wprocess (id text NOT NULL PRIMARY KEY
                  ,name text NOT NULL DEFAULT ''
                  ,iduser text NOT NULL
                  ,args text NOT NULL DEFAULT ''
                  ,args2 text NOT NULL DEFAULT ''
                  ,pageindex integer NOT NULL DEFAULT 1
                  ,time1 integer NOT NULL DEFAULT 0
                  ,time2 integer NOT NULL DEFAULT 0);
                 CREATE INDEX IF NOT EXISTS index_wprocess_iduser ON wproccess (iduser);
            ");
        }
        /**
         * Dữ liệu báo cáo tuần
         * */
        public static dbSQLite getDataBaoCaoTuan(string matinh = "")
        {
            string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", $"BaoCaoTuan{matinh}.db");
            var db = new dbSQLite(pathDB);
            var tables = db.getAllTables();
            db.CreatePhucLucBcTuan(tables);
            db.CreateBcTuan(tables);
            return db;
        }

        public static dbSQLite getDataImportBaoCaoTuan(string matinh = "")
        {
            string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", $"ImportBaoCaoTuan{matinh}.db");
            var db = new dbSQLite(pathDB);
            var tables = db.getAllTables();
            db.CreateImportBcTuan(tables);
            return db;
        }

        public static void CreateBcTuan(this dbSQLite dbConnect, List<string> tables = null)
        {
            if (tables == null) { tables = dbConnect.getAllTables(); }
            var tsqlCreate = new List<string>();
            /* BaoCaoTuanDocx */
            tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS bctuandocx (
                    id text not null primary key /* Mã hóa rút gọn và gợi nhớ cho mỗi lần lập BC tuần Dùng trình bày danh sách báo cáo đã lập để tiện cho chọn và xử lý thao tác: Khóa và mở khóa báo cáo/ xóa báo cáo/ xem/in lại */
                    ,x1 real not null default 0 /* Tổng tiền các CSKCB đã đề nghị bảo hiểm thanh toán (T_BHTT): X1={cột R (T-BHTT) bảng B02_TOANQUOC }. Làm tròn đến triệu đồng */
                    ,x2 text not null default '' /* Số của Quyết định giao dự toán: X2={“ Nếu không tìm thấy dòng nào của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán thì “TW chưa giao dự toán, tạm lấy theo dự toán năm trước”, nếu thấy thì  lấy số ký hiệu các dòng QĐ của năm 2024 ở bảng hệ thống lưu thông tin quyết định giao dự toán} */
                    ,x3 real not null default 0 /* X3={Như trên, ko thấy thì lấy tổng tiền các dòng dự toán năm trước, thấy thì lấy tổng số tiền các dòng quyết định năm nay} */
                    ,x4 real not null default 0 /* So sánh với dự toán, tỉnh đã sử dụng X4={X1/X2%} */
                    ,x5 real not null default 0 /* Tỷ lệ điều trị nội trú X5={Cột G, dòng MA_TINH=10}; */
                    ,x6 real not null default 0 /* bình quân toàn quốc X6={cột G, dòng MA_TINH=00}; */
                    ,x7 text not null default '' /* Số chênh lệch X7={đoạn văn tùy thuộc X5> hay < X6. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x8 integer not null default 0 /* xếp thứ so với các tỉnh X8={Sort cột G (TYLE_NOITRU ) cao xuống thấp và lấy thứ tự}; */
                    ,x9 real not null default 0 /* Bình quân vùng X9 ={tính toán: total cột F (TONG_LUOT_NOI) chia cho Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG=mã vùng của tỉnh báo cáo}; */
                    ,x10 text not null default '' /* Số chênh lệch X10 ={đoạn văn tùy thuộc X5> hay < X9. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x11 integer not null default 0 /* đứng thứ so với vùng. X11= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort cột G (TYLE_NOITRU ) cao –thấp và lấy thứ tự} */
                    ,x12 real not null default 0 /* Ngày điều trị bình quân X12={Cột H, dòng MA_TINH=10}; */
                    ,x13 real not null default 0 /* bình quân toàn quốc X13={cột H, dòng MA_TINH=00}; */
                    ,x14 text not null default '' /* Số chênh lệch X14={đoạn văn tùy thuộc X12> hay < X13. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x15 integer not null default 0 /* xếp thứ so toàn quốc X15={Sort cột H (NGAY_DTRI_BQ) cao xuống thấp và lấy thứ tự}; */
                    ,x16 real not null default 0 /* Bình quân vùng X16 ={tính toán: A-Tổng ngày điều trị nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của cá tỉnh cùng mã vùng. A=Total(cột H (NGAY_DTRI_BQ) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                    ,x17 text not null default '' /* Số chênh lệch X17 ={đoạn văn tùy thuộc X12> hay < X16. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x18 integer not null default 0 /* đứng thứ so với vùng X18= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột H (NGAY_DTRI_BQ) cao –thấp và lấy thứ tự} */
                    ,x19 real not null default 0 /* Chi bình quân chung X19={Cột I (CHI_BQ_CHUNG), dòng MA_TINH=10}; */
                    ,x20 real not null default 0 /* bình quân toàn quốc X20={cột I, dòng MA_TINH=00}; */
                    ,x21 text not null default '' /* Số chênh lệch X21={đoạn văn tùy thuộc X19> hay < X20. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x22 integer not null default 0 /* xếp thứ so toàn quốc X22={Sort cột I cao xuống thấp và lấy thứ tự}; */
                    ,x23 real not null default 0 /* Bình quân vùng X23={tính toán: A-Tổng chi các tỉnh cùng mã vùng / B- Tổng lượt kcb của các tỉnh cùng mã vùng. A=Total  (cột I (CHI_BQ_CHUNG) * cột D (TONG_LUOT)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột D (TONG_LUOT) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                    ,x24 text not null default '' /* Số chênh lệch X24 ={đoạn văn tùy thuộc X19> hay < X23. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x25 integer not null default 0 /* đứng thứ so với vùng X25= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột I (CHI_BQ_CHUNG) cao –thấp và lấy thứ tự} */
                    ,x26 real not null default 0 /* Chi bình quân ngoại trú X26={Cột J (CHI_BQ_NGOAI), dòng MA_TINH=10}; */
                    ,x27 real not null default 0 /* bình quân toàn quốc X27={cột J, dòng MA_TINH=00}; */
                    ,x28 text not null default '' /* Số chênh lệch X28={đoạn văn tùy thuộc X26> hay < X27. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x29 integer not null default 0 /* xếp thứ so toàn quốc X29={Sort cột J cao xuống thấp và lấy thứ tự}; */
                    ,x30 real not null default 0 /* Bình quân vùng X30={tính toán: A-Tổng chi ngoại trú các tỉnh cùng mã vùng / B- Tổng lượt kcb ngoại trú của các tỉnh cùng mã vùng. A=Total  (cột J (CHI_BQ_NGOAI) * cột E (TONG_LUOT_NGOAI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột E (TONG_LUOT_NGOAI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                    ,x31 text not null default '' /* Số chênh lệch X31 ={đoạn văn tùy thuộc X19> hay < X30. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x32 integer not null default 0 /* đứng thứ so với vùng X32= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột J (CHI_BQ_NGOAI) cao –thấp và lấy thứ tự} */
                    ,x33 real not null default 0 /* Chi bình quân nội trú X33={Cột K (CHI_BQ_NOI), dòng MA_TINH=10}; */
                    ,x34 real not null default 0 /* bình quân toàn quốc X34={cột K, dòng MA_TINH=00}; */
                    ,x35 text not null default '' /* Số chênh lệch X35={đoạn văn tùy thuộc X33> hay < X34. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x36 integer not null default 0 /* xếp thứ so toàn quốc X36={Sort cột K cao xuống thấp và lấy thứ tự}; */
                    ,x37 real not null default 0 /* Bình quân vùng X37={tính toán: A-Tổng chi nội trú các tỉnh cùng mã vùng / B- Tổng lượt kcb nội trú của các tỉnh cùng mã vùng. A=Total  (cột K (CHI_BQ_NOI) * cột F (TONG_LUOT_NOI)) của tất cả các tỉnh cùng MA_VUNG với tỉnh báo cáo. B= Total cột F (TONG_LUOT_NOI) của các tỉnh có MA_VUNG cùng mã vùng của tỉnh báo cáo}; */
                    ,x38 text not null default '' /* Số chênh lệch X38 ={đoạn văn tùy thuộc X33> hay < X34. Nếu lớn hơn, lấy chuỗi “cao hơn”, không thì “thấp hơn” ghép với trị tuyệt đối của hiệu số }; */
                    ,x39 integer not null default 0 /* đứng thứ so với vùng X39= {lọc các dòng tỉnh có mã vùng trùng với mã vùng của tỉnh, sort Cột K (CHI_BQ_NOI) cao –thấp và lấy thứ tự} */
                    ,x40 real not null default 0 /* Bình quân xét nghiệm X40= {cột P dòng có mã tỉnh =10}; */
                    ,x41 real not null default 0 /* số tương đối so kỳ trước X41={nếu cột Q dòng có mã tỉnh=10 là số dương, “tăng “ & cột Q & “%”, không thì “giảm “ & cột Q %}; */
                    ,x42 text not null default '' /* số tuyệt đối  so kỳ trước X42={nếu cột Q là dương, “tăng “ & cột P trừ đi (cột P chia (cột Q +100) *100 ) & “ đồng”, không thì “giảm “ &  cột P trừ đi (cột P chia (cột Q +100) *100 ) & “ đồng”}; */
                    ,x43 real not null default 0 /* Bình quân CĐHA X43= {cột R dòng có mã tỉnh =10}; */
                    ,x44 text not null default '' /* số tương đối X44={nếu cột S dòng có mã tỉnh=10 là số dương, “tăng “ & cột S & “%”, không thì “giảm “ & cột S %}; */
                    ,x45 text not null default '' /* số tuyệt đối X45={nếu cột S là dương, “tăng “ & cột R trừ đi (cột R chia (cột S +100) *100 ) & “ đồng”, không thì “giảm “ &  cột R trừ đi (cột R chia (cột S +100) *100 ) & “ đồng”}; */
                    ,x46 real not null default 0 /* Bình quân thuốc X46= {cột T dòng có mã tỉnh =10}; */
                    ,x47 text not null default '' /* số tương đối X47={nếu cột U dòng có mã tỉnh=10 là số dương, “tăng “ & cột U & “%”, không thì “giảm “ & cột U %}; */
                    ,x48 text not null default '' /* số tuyệt đối X48={nếu cột U là dương, “tăng “ & cột T trừ đi (cột T chia (cột U +100) *100 ) & “ đồng”, không thì “giảm “ &  cột T trừ đi (cột T chia (cột U+100) *100 ) & “ đồng”} */
                    ,x49 real not null default 0 /* Bình quân chi phẫu thuật X49= {cột V dòng có mã tỉnh =10}; */
                    ,x50 text not null default '' /* số tương đối X50={nếu cột W dòng có mã tỉnh=10 là số dương, “tăng “ & cột W & “%”, không thì “giảm “ & cột W %}; */
                    ,x51 text not null default '' /* số tuyệt đối X51={nếu cột W là dương, “tăng “ & cột V trừ đi (cột V chia (cột W +100) *100 ) & “ đồng”, không thì “giảm “ &  cột V trừ đi (cột V chia (cột W+100) *100 ) & “ đồng”} */
                    ,x52 real not null default 0 /* Bình quân chi thủ thuật X52= {cột X dòng có mã tỉnh =10}; */
                    ,x53 text not null default '' /* số tương đối X53={nếu cột Y dòng có mã tỉnh=10 là số dương, “tăng “ & cột Y & “%”, không thì “giảm “ & cột Y %}; */
                    ,x54 text not null default '' /* số tuyệt đối X54={nếu cột Y là dương, “tăng “ & cột X trừ đi (cột X chia (cột Y +100) *100 ) & “ đồng”, không thì “giảm “ &  cột X trừ đi (cột X chia (cột Y+100) *100 ) & “ đồng”} */
                    ,x55 real not null default 0 /* Bình quân chi vật tư y tế X55= {cột Z dòng có mã tỉnh =10}; */
                    ,x56 text not null default '' /* số tương đối X56={nếu cột AA dòng có mã tỉnh=10 là số dương, “tăng “ & cột AA & “%”, không thì “giảm “ & cột AA %}; */
                    ,x57 text not null default '' /* số tuyệt đối X57={nếu cột AA là dương, “tăng “ & cột Z trừ đi (cột Z chia (cột AA +100) *100 ) & “ đồng”, không thì “giảm “ &  cột Z trừ đi (cột Z chia (cột AA+100) *100 ) & “ đồng”} */
                    ,x58 real not null default 0 /* Bình quân chi tiền giường X58= {cột AB dòng có mã tỉnh =10}; */
                    ,x59 text not null default '' /* số tương đối X59={nếu cột AC dòng có mã tỉnh=10 là số dương, “tăng “ & cột AC & “%”, không thì “giảm “ & cột AC %}; */
                    ,x60 text not null default '' /* số tuyệt đối X60={nếu cột AC là dương, “tăng “ & cột AB trừ đi (cột AB chia (cột AC +100) *100 ) & “ đồng”, không thì “giảm “ &  cột AB trừ đi (cột AB chia (cột AC+100) *100 ) & “ đồng”} */
                    ,x61 real not null default 0 /* Chỉ định xét nghiệm X61={cột AD, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
                    ,x62 text not null default '' /* số tương đối X62={cột AE dòng có mã tỉnh=10 & “%”}; */
                    ,x63 text not null default '' /* số tuyệt đối X63 {tính toán: X61 trừ đi (X61 chia (cột AE+100)*100) & “bệnh nhân”}  */
                    ,x64 real not null default 0 /* Chỉ định CĐHA X64={cột AF, dòng có mã tỉnh =10 nhân với 100 để ra số người}; */
                    ,x65 text not null default '' /* số tương đối X65={cột AG dòng có mã tỉnh=10 & “%”}; */
                    ,x66 text not null default '' /* số tuyệt đối X66 {tính toán: X64 trừ đi (X64 chia (cột AG+100)*100) & “bệnh nhân”} */
                    ,x67 text not null default '' /* Công tác kiểm soát chi X67={lần đầu lập BC sẽ rỗng, người dùng tự trình bày văn bản, lưu lại ở bảng dữ liệu kết quả báo cáo, kỳ sau sẽ tự động lấy từ kỳ trước, để người dùng kế thừa, sửa và lưu dùng cho kỳ này và kỳ sau} */
                    ,x68 text not null default '' /* Công tác thanh, quyết toán năm X68={tương tự X67} */
                    ,x69 text not null default '' /* Phương hướng kỳ tiếp theo X69={tương tự X67} */
                    ,x70 text not null default '' /* Khó khăn, vướng mắc, đề xuất (nếu có) X70={tương tự X67} */
                    ,x71 real not null default 0 /* Trong đó: Nội trú X71 = {cột S T_BHTT_NOI bảng B02_TOANQUOC }; */
                    ,x72 real not null default 0 /* Ngoại trú X72={cột T T_BHTT_NGOAI bảng B02_TOANQUOC } */
                    ,x73 text not null default '' /* Tên tỉnh/thành phố lập BC Lấy biến hệ thống khởi tạo khi User đăng nhập */
                    ,x74 text not null default '' /* THOI_GIAN_BC Chuỗi ký tự ngày lập BC. mặc định từ ô C3 biểu B26 khi khởi tạo 1 báo cáo	Có ô cho nhập, sửa */
                    ,userid text not null default '' /* Lưu ID của người dùng */
                    ,ma_tinh text not null default '' /* Lưu mã tỉnh làm báo cáo */
                    ,ngay integer not null default 0 /* Ngày làm báo cáo dạng timestamp */
                    ,timecreate integer not null default 0 /* Thời điểm tạo báo cáo */);");
            tsqlCreate.Add("CREATE INDEX IF NOT EXISTS bctuandocx_ma_tinh ON bctuandocx(ma_tinh);");
            tsqlCreate.Add("CREATE INDEX IF NOT EXISTS index_bctuandocx_timecreate ON bctuandocx(timecreate);");
            tsqlCreate.Add("CREATE INDEX IF NOT EXISTS index_bctuandocx_ngay ON bctuandocx(ngay);");
            if (tsqlCreate.Count > 0) { dbConnect.Execute(string.Join(Environment.NewLine, tsqlCreate)); }
        }

        public static void CreatePhucLucBcTuan(this dbSQLite dbConnect, List<string> tables = null)
        {
            if (tables == null) { tables = dbConnect.getAllTables(); }
            var tsqlCreate = new List<string>();
            if (tables.Contains("pl01") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS pl01 (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_tinh text not null default '' /* Mã tỉnh Cột A, B02 */
                ,ten_tinh text not null default '' /* Tên tỉnh Cột B, B02 */
                ,ma_vung text not null default '' /* Mã vùng 0,1,2,3,4... cột C , B02 */
                ,tyle_noitru real not null default 0 /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú, B02 */
                ,ngay_dtri_bq real not null default 0 /*	Ngày điều trị BQ, vd 6,42, DVT: ngày; Lấy từ cột H: NGAY ĐT_BQ, B02 */
                ,chi_bq_chung real not null default 0 /* Chi bình quan chung lượt KCB ĐVT ( đồng)	Cột I, B02 */
                ,chi_bq_ngoai real not null default 0 /* Chi bình quân ngoại trú/lượt KCB ngoại trú (đồng); Cột J, B02 */
                ,chi_bq_noi real not null default 0 /* Như trên nhưng với nội trú	Cột K, B02 */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                 CREATE INDEX IF NOT EXISTS index_pl01_id_bc ON pl01 (id_bc);");
            }
            if (tables.Contains("pl02") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS pl02 (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_tinh text not null default '' /* Mã tỉnh Cột A, B02 */
                ,ten_tinh text not null default '' /* Tên tỉnh Cột B, B02 */
                ,ma_vung text not null default '' /* Mã vùng */
                ,chi_bq_xn real not null default 0 /* chi BQ Xét nghiệm; đơn vị tính : đồng	Lấy từ B04 . Cột D */
                ,chi_bq_cdha real not null default 0 /* chi BQ Chẩn đoán hình ảnh; Lấy từ B04. Cột E */
                ,chi_bq_thuoc real not null default 0 /* chi BQ thuốc; Lấy từ B04. Cột F */
                ,chi_bq_pttt real not null default 0 /* chi BQ phẫu thuật thủ thuật	Lấy từ B04. Cột G */
                ,chi_bq_vtyt real not null default 0 /* chi BQ vật tư y tế; Lấy từ B04. Cột H */
                ,chi_bq_giuong real not null default 0 /* chi BQ tiền giường; Lấy từ B04. Cột I */
                ,ngay_ttbq real not null default 0 /* Ngày thanh toán bình quân; Lấy từ B04. Cột J */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                 CREATE INDEX IF NOT EXISTS index_pl02_id_bc ON pl02 (id_bc);");
            }
            if (tables.Contains("pl03") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS pl03 (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_cskcb text not null /* Mã cơ sơ KCB, có chứa cả mã toàn quốc:00, mã vùng V1, mã tỉnh 10 và mã CSKCB ví dụ 10061; Ngoài 3 dòng đầu lấy từ bảng lưu thông tin Sheet 1; Các dòng còn lại lấy từ các cột A Excel B02 */
                ,ten_cskcb text not null default '' /* Tên cskcb, ghép hạng BV vào đầu chuỗi tên CSKCB	Côt B */
                ,ma_vung text not null default '' /* Mã vùng */
                ,tyle_noitru real not null default 0 /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú */
                ,ngay_dtri_bq real not null default 0 /* Ngày điều trị BQ, vd 6,42, DVT: NGÀY; Lấy từ cột H: NGAY ĐT_BQ */
                ,chi_bq_chung real not null default 0 /* Chi bình quan chung lượt KCB ĐVT đồng; Cột I B02 */
                ,chi_bq_ngoai real not null default 0 /* Chi bình quân ngoại trú/lượt KCB ngoại trú	Cột J B02 */
                ,chi_bq_noi real not null default 0 /* Như trên nhưng với nội trú; Cột K B02 */
                ,tuyen_bv text not null default ''
                ,hang_bv text not null default ''
                ,userid text not null default '' /* Lưu ID của người dùng */);
                 CREATE INDEX IF NOT EXISTS index_pl03_id_bc ON pl03 (id_bc);");
            }
            if (tsqlCreate.Count > 0) { dbConnect.Execute(string.Join(Environment.NewLine, tsqlCreate)); }
        }

        public static void CreateImportBcTuan(this dbSQLite dbConnect, List<string> tables = null)
        {
            if (tables == null) { tables = dbConnect.getAllTables(); }
            var tsqlCreate = new List<string>();
            /* B02. Thống kê KCB (Tháng) */
            if (tables.Contains("b02") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b02 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,ma_loai_kcb text not null
                ,tu_thang integer not null default 0
                ,den_thang integer not null default 0
                ,nam integer not null default 0
                ,loai_bv integer not null default 0
                ,kieubv integer not null default 0
                ,loaick integer not null default 0
                ,hang_bv integer not null default 0
                ,tuyen integer not null default 0
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_b02_id_bc ON b02 (id_bc);");
            }
            if (tables.Contains("b02chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b02chitiet (id INTEGER primary key AUTOINCREMENT
                ,id2 integer not null default 0
                ,ma_tinh text not null default ''
                ,ten_tinh text not null default ''
                ,ma_cskcb text not null default ''
                ,ten_cskcb text not null default ''
                ,ma_vung text not null default ''
                ,tong_luot integer not null default 0
                ,tong_luot_ngoai integer not null default 0
                ,tong_luot_noi integer not null default 0
                ,tyle_noitru real not null default 0
                ,ngay_dtri_bq real not null default 0
                ,chi_bq_chung real not null default 0
                ,chi_bq_ngoai real not null default 0
                ,chi_bq_noi real not null default 0
                ,tong_chi real not null default 0
                ,ty_trong real not null default 0
                ,tong_chi_ngoai real not null default 0
                ,ty_trong_kham real not null default 0
                ,tong_chi_noi real not null default 0
                ,ty_trong_giuong real not null default 0
                ,t_bhtt real not null default 0
                ,t_bhtt_noi real not null default 0
                ,t_bhtt_ngoai real not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_b02chitiet_id_bc ON b02chitiet (id_bc);");
            }

            /* B04. Thống kê chi bình quân (Tháng) */
            if (tables.Contains("b04") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b04 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,tu_thang integer not null default 0
                ,den_thang integer not null default 0
                ,nam integer not null default 0
                ,ma_loai_kcb integer not null default 0
                ,loai_bv integer not null default 0
                ,hang_bv integer not null default 0
                ,tuyen integer not null default 0
                ,kieubv integer not null default 0
                ,loaick integer not null default 0
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_b04_id_bc ON b04 (id_bc);");
            }
            if (tables.Contains("b04chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b04chitiet (id INTEGER primary key AUTOINCREMENT
                ,id2 integer not null default 0
                ,ma_tinh text not null default ''
                ,ten_tinh text not null default ''
                ,ma_cskcb text not null default ''
                ,ten_cskcb text not null default ''
                ,chi_bq_luotkcb real not null default 0
                ,bq_xn real not null default 0
                ,bq_cdha real not null default 0
                ,bq_thuoc real not null default 0
                ,bq_ptt real not null default 0
                ,bq_vtyt real not null default 0
                ,bq_giuong real not null default 0
                ,ngay_ttbq real not null default 0
                ,ma_vung text not null default ''
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_b04chitiet_id_bc ON b04chitiet (id_bc);");
            }

            /* B26. Thống kê gia tăng chi phí KCB BHYT theo NĐ75 (theo ngày nhận) */
            if (tables.Contains("b26") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b26 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,loai_kcb text not null default ''
                ,thoigian integer not null default 0
                ,loai_bv integer not null default 0
                ,kieubv integer not null default 0
                ,loaick integer not null default 0
                ,hang_bv integer not null default 0
                ,tuyen integer not null default 0
                ,loai_so_sanh text not null default ''
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_b26_id_bc ON b26 (id_bc);");
            }
            if (tables.Contains("b26chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS b26chitiet (id INTEGER primary key AUTOINCREMENT,
                id2 integer not null default 0,
                ma_tinh text not null default '',
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
                ,id_bc text not null default '');
                CREATE INDEX IF NOT EXISTS index_b26chitiet_id_bc ON b26chitiet (id_bc);");
            }
            if (tsqlCreate.Count > 0) { dbConnect.Execute(string.Join(Environment.NewLine, tsqlCreate)); }
        }

        /**
         * Dữ liệu báo cáo tuần
         * */
        public static dbSQLite getDataBcThang(string matinh = "")
        {
            string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", $"BaoCaoThang{matinh}.db");
            var db = new dbSQLite(pathDB);
            var tables = db.getAllTables();
            db.CreatePhucLucBcThang(tables);
            db.CreateBcThang(tables);
            return db;
        }

        public static dbSQLite getDataImportBaoCaoThang(string matinh = "")
        {
            string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", $"ImportBaoCaoThang{matinh}.db");
            var db = new dbSQLite(pathDB);
            var tables = db.getAllTables();
            db.CreateImportBcThang(tables);
            return db;
        }

        public static void CreateBcThang(this dbSQLite dbConnect, List<string> tables = null)
        {
            if (tables == null) { tables = dbConnect.getAllTables(); }
            var tsqlCreate = new List<string>();
            /* BaoCaoTuanDocx */
            tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS bcthangdocx (id text not null primary key
                ,tentinh text not null default '' /* Tên tỉnh */
                ,ngay text not null default '' /* Ngày báo cáo (Chứa luôn cả ngày đầu tháng, và năm) */
                ,ngaydauthang text not null '' /* Ngày đầu tháng */
                ,thang integer not null '' /* Tháng báo cáo */
                ,nam1 integer not null default 0 /* Năm báo cáo */
                ,nam2 integer not null default 0 /* Năm trước báo cáo */
                ,x1 text not null default '' /* Công tác ký hợp đồng KCB BHYT */
                ,x33 text not null default '' /* Công tác kiểm soát chi KCB BHYT */
                ,x34 text not null default '' /* Công tác đấu thầu thuốc */
                ,x35 text not null default '' /* Công tác quyết toán chi KCB BHYT */
                ,x36 text not null default '' /* Công tác khác */
                ,x37 text not null default '' /* Phương hướng công tác tháng sau */
                ,x38 text not null default '' /* Khó khăn, vướng mắc, đề xuất (nếu có) */

                ,mi12c1 real not null default '' /* Dự toán giao {nam} */
                ,mi12c2 real not null default '' /* Chi KCB toàn tỉnh */
                ,mi12c3 real not null default 0 /* Tỷ lệ % SD dự toán {nam} */
                ,mi12c4 integer not null default 0 /* xếp bn toàn quốc */
                ,mi12c5 integer not null default 0 /* xếp thứ bao nhiêu so với vùng */
                ,mi12c6 real not null default 0 /* Tỷ lệ % SD dự toán {nam2} */
                ,mi12c7 real not null default 0 /* So cùng kỳ năm trước = 3-6 (mi12c3 - mi12c6) */

                ,m13lc11 real not null default 0 /* Tổng lượt = 2+3 (m13lc21+m13lc31) */
                ,m13lc21 real not null default 0 /* Lượt ngoại */
                ,m13lc31 real not null default 0 /* Lượt nội */
                ,m13lc41 real not null default 0 /* Tổng lượt = 5+6 (m13lc51+m13lc61) */
                ,m13lc51 real not null default 0 /* Lượt ngoại */
                ,m13lc61 real not null default 0 /* Lượt nội */

                ,m13lc12 real not null default 0 /* Tổng lượt = 2+3 (m13lc22+m13lc32) */
                ,m13lc22 real not null default 0 /* Lượt ngoại */
                ,m13lc32 real not null default 0 /* Lượt nội */
                ,m13lc42 real not null default 0 /* Tổng lượt = 5+6 (m13lc52+m13lc62) */
                ,m13lc52 real not null default 0 /* Lượt ngoại */
                ,m13lc62 real not null default 0 /* Lượt nội */

                ,m13lc13 real not null default 0 /* Tổng lượt = 2+3 (m13lc23+m13lc33) */
                ,m13lc23 real not null default 0 /* Lượt ngoại */
                ,m13lc33 real not null default 0 /* Lượt nội */
                ,m13lc43 real not null default 0 /* Tổng lượt = 5+6 (m13lc53+m13lc63) */
                ,m13lc53 real not null default 0 /* Lượt ngoại */
                ,m13lc63 real not null default 0 /* Lượt nội */

                ,m13lc14 real not null default 0 /* Tổng lượt = 2+3 (m13lc24+m13lc34) */
                ,m13lc24 real not null default 0 /* Lượt ngoại */
                ,m13lc34 real not null default 0 /* Lượt nội */
                ,m13lc44 real not null default 0 /* Tổng lượt = 5+6 (m13lc54+m13lc64) */
                ,m13lc54 real not null default 0 /* Lượt ngoại */
                ,m13lc64 real not null default 0 /* Lượt nội */

                ,mc13cc11 real not null default 0 /* Tổng lượt = 2+3 (mc13cc21+mc13cc31) */
                ,mc13cc21 real not null default 0 /* Chi ngoại trú */
                ,mc13cc31 real not null default 0 /* Chi nội trú */
                ,mc13cc41 real not null default 0 /* Tổng lượt = 5+6 (mc13cc51+mc13cc61) */
                ,mc13cc51 real not null default 0 /* Chi ngoại trú */
                ,mc13cc61 real not null default 0 /* Chi nội trú */

                ,mc13cc12 real not null default 0 /* Tổng lượt = 2+3 (mc13cc22+mc13cc32) */
                ,mc13cc22 real not null default 0 /* Chi ngoại trú */
                ,mc13cc32 real not null default 0 /* Chi nội trú */
                ,mc13cc42 real not null default 0 /* Tổng lượt = 5+6 (mc13cc52+mc13cc62) */
                ,mc13cc52 real not null default 0 /* Chi ngoại trú */
                ,mc13cc62 real not null default 0 /* Chi nội trú */

                ,mc13cc13 real not null default 0 /* Tổng lượt = 2+3 (mc13cc23+mc13cc33) */
                ,mc13cc23 real not null default 0 /* Chi ngoại trú */
                ,mc13cc33 real not null default 0 /* Chi nội trú */
                ,mc13cc43 real not null default 0 /* Tổng lượt = 5+6 (mc13cc53+mc13cc63) */
                ,mc13cc53 real not null default 0 /* Chi ngoại trú */
                ,mc13cc63 real not null default 0 /* Chi nội trú */

                ,mc13cc14 real not null default 0 /* Tổng lượt = 2+3 (mc13cc24+mc13cc34) */
                ,mc13cc24 real not null default 0 /* Chi ngoại trú */
                ,mc13cc34 real not null default 0 /* Chi nội trú */
                ,mc13cc44 real not null default 0 /* Tổng lượt = 5+6 (mc13cc54+mc13cc64) */
                ,mc13cc54 real not null default 0 /* Chi ngoại trú */
                ,mc13cc64 real not null default 0 /* Chi nội trú */       

                ,userid text not null default '' /* Lưu ID của người dùng */
                ,ma_tinh text not null default '' /* Lưu mã tỉnh làm báo cáo */
                ,timespan integer not null default 0 /* Ngày làm báo cáo dạng timestamp */
                ,timecreate integer not null default 0 /* Thời điểm tạo báo cáo */);");
            tsqlCreate.Add("CREATE INDEX IF NOT EXISTS bcthangdocx_ma_tinh ON bcthangdocx(ma_tinh);");
            tsqlCreate.Add("CREATE INDEX IF NOT EXISTS index_bcthangdocx_timecreate ON bcthangdocx(timecreate);");
            tsqlCreate.Add("CREATE INDEX IF NOT EXISTS index_bcthangdocx_ngay ON bcthangdocx(ngay);");
            if (tsqlCreate.Count > 0) { dbConnect.Execute(string.Join(Environment.NewLine, tsqlCreate)); }
        }

        public static void CreatePhucLucBcThang(this dbSQLite dbConnect, List<string> tables = null)
        {
            if (tables == null) { tables = dbConnect.getAllTables(); }
            var tsqlCreate = new List<string>();

            if (tables.Contains("thangpl01") == false)
            {
                /**
                 Yêu cầu nhập excel từ người dùng
                 */
                /* PHỤ LỤC 01. TÌNH HÌNH SỬ DỤNG DỰ TOÁN THEO HỢP ĐỒNG (luy kế năm của csyt) */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl01 (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_cskcb text not null /* Mã cơ sơ KCB */
                ,ten_cskcb text not null default '' /* Tên cskcb*/
                ,dutoangiao real not null default 0 /* Dự toán tạm giao */
                ,tien_bhtt real not null default 0 /* Tiền T- BHTT cột R-B02-10 */
                ,tl_sudungdt real not null default 0 /* Tỷ lệ sử dụng dự toán = (tien_bhtt/dutoangiao)*100  */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                CREATE INDEX IF NOT EXISTS index_thangpl01_id_bc ON thangpl01 (id_bc);");
            }

            if (tables.Contains("thangpl02a") == false)
            {
                /* Lấy dữ liệu từ biểu pl02a trong tháng (Từ tháng đến tháng = tháng báo cáo) */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl02a (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_tinh text not null default '' /* Mã tỉnh Cột A, B02 */
                ,ten_tinh text not null default '' /* Tên tỉnh Cột B, B02 */
                ,ma_vung text not null default '' /* Mã vùng 0,1,2,3,4... cột C , B02 */
                ,tyle_noitru real not null default 0 /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú, B02 */
                ,ngay_dtri_bq real not null default 0 /* Ngày điều trị BQ, vd 6,42, DVT: ngày; Lấy từ cột H: NGAY ĐT_BQ, B02 */
                ,chi_bq_chung real not null default 0 /* Chi bình quan chung lượt KCB ĐVT ( đồng)	Cột I, B02 */
                ,chi_bq_ngoai real not null default 0 /* Chi bình quân ngoại trú/lượt KCB ngoại trú (đồng); Cột J, B02 */
                ,chi_bq_noi real not null default 0 /* Như trên nhưng với nội trú	Cột K, B02 */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                 CREATE INDEX IF NOT EXISTS index_thangpl02a_id_bc ON thangpl02a (id_bc);");
            }
            if (tables.Contains("thangpl02b") == false)
            {
                /* Lấy dữ liệu từ biểu pl02b dành cho cả năm (từ tháng 1 đến tháng báo cáo) */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl02b (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_tinh text not null default '' /* Mã tỉnh Cột A, B02 */
                ,ten_tinh text not null default '' /* Tên tỉnh Cột B, B02 */
                ,ma_vung text not null default '' /* Mã vùng 0,1,2,3,4... cột C , B02 */
                ,tyle_noitru real not null default 0 /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú, B02 */
                ,ngay_dtri_bq real not null default 0 /* Ngày điều trị BQ, vd 6,42, DVT: ngày; Lấy từ cột H: NGAY ĐT_BQ, B02 */
                ,chi_bq_chung real not null default 0 /* Chi bình quan chung lượt KCB ĐVT ( đồng)	Cột I, B02 */
                ,chi_bq_ngoai real not null default 0 /* Chi bình quân ngoại trú/lượt KCB ngoại trú (đồng); Cột J, B02 */
                ,chi_bq_noi real not null default 0 /* Như trên nhưng với nội trú	Cột K, B02 */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                 CREATE INDEX IF NOT EXISTS index_thangpl02b_id_bc ON thangpl02b (id_bc);");
            }


            if (tables.Contains("thangpl03a") == false)
            {
                /* Lấy dữ liệu từ biểu pl03a csyt trong tháng */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl03a (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_cskcb text not null /* Mã cơ sơ KCB, có chứa cả mã toàn quốc:00, mã vùng V1, mã tỉnh 10 và mã CSKCB ví dụ 10061; Ngoài 3 dòng đầu lấy từ bảng lưu thông tin Sheet 1; Các dòng còn lại lấy từ các cột A Excel B02 */
                ,ten_cskcb text not null default '' /* Tên cskcb, ghép hạng BV vào đầu chuỗi tên CSKCB	Côt B */
                ,ma_vung text not null default '' /* Mã vùng */
                ,tyle_noitru real not null default 0 /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú */
                ,ngay_dtri_bq real not null default 0 /* Ngày điều trị BQ, vd 6,42, DVT: NGÀY; Lấy từ cột H: NGAY ĐT_BQ */
                ,chi_bq_chung real not null default 0 /* Chi bình quan chung lượt KCB ĐVT đồng; Cột I B02 */
                ,chi_bq_ngoai real not null default 0 /* Chi bình quân ngoại trú/lượt KCB ngoại trú	Cột J B02 */
                ,chi_bq_noi real not null default 0 /* Như trên nhưng với nội trú; Cột K B02 */
                ,tuyen_bv text not null default ''
                ,hang_bv text not null default ''
                ,userid text not null default '' /* Lưu ID của người dùng */);
                CREATE INDEX IF NOT EXISTS index_thangpl03a_id_bc ON thangpl03a (id_bc);");
            }
            if (tables.Contains("thangpl03b") == false)
            {
                /* Lấy dữ liệu từ biểu pl03a csyt luỹ kế của năm */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl03b (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_cskcb text not null /* Mã cơ sơ KCB, có chứa cả mã toàn quốc:00, mã vùng V1, mã tỉnh 10 và mã CSKCB ví dụ 10061; Ngoài 3 dòng đầu lấy từ bảng lưu thông tin Sheet 1; Các dòng còn lại lấy từ các cột A Excel B02 */
                ,ten_cskcb text not null default '' /* Tên cskcb, ghép hạng BV vào đầu chuỗi tên CSKCB	Côt B */
                ,ma_vung text not null default '' /* Mã vùng */
                ,tyle_noitru real not null default 0 /* Tỷ lệ nội trú, ví dụ 19,49%	Lấy từ cột G: TL_Nội trú */
                ,ngay_dtri_bq real not null default 0 /* Ngày điều trị BQ, vd 6,42, DVT: NGÀY; Lấy từ cột H: NGAY ĐT_BQ */
                ,chi_bq_chung real not null default 0 /* Chi bình quan chung lượt KCB ĐVT đồng; Cột I B02 */
                ,chi_bq_ngoai real not null default 0 /* Chi bình quân ngoại trú/lượt KCB ngoại trú	Cột J B02 */
                ,chi_bq_noi real not null default 0 /* Như trên nhưng với nội trú; Cột K B02 */
                ,tuyen_bv text not null default ''
                ,hang_bv text not null default ''
                ,userid text not null default '' /* Lưu ID của người dùng */);
                CREATE INDEX IF NOT EXISTS index_thangpl03b_id_bc ON thangpl03b (id_bc);");
            }

            if (tables.Contains("thangpl04a") == false)
            {
                /* Lấy dữ liệu từ biểu pl04a toàn quốc luỹ kế name */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl04a (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,mahanhchinh text not null default '' /*  */
                ,tenhanhchinh text not null default '' /*  */
                ,ma_vung text not null default '' /* Mã vùng */
                ,chi_bq_xn real not null default 0 /* chi BQ Xét nghiệm; đơn vị tính : đồng	Lấy từ B04 . Cột D */
                ,chi_bq_cdha real not null default 0 /* chi BQ Chẩn đoán hình ảnh; Lấy từ B04. Cột E */
                ,chi_bq_thuoc real not null default 0 /* chi BQ thuốc; Lấy từ B04. Cột F */
                ,chi_bq_pttt real not null default 0 /* chi BQ phẫu thuật thủ thuật	Lấy từ B04. Cột G */
                ,chi_bq_vtyt real not null default 0 /* chi BQ vật tư y tế; Lấy từ B04. Cột H */
                ,chi_bq_giuong real not null default 0 /* chi BQ tiền giường; Lấy từ B04. Cột I */
                ,ngay_ttbq real not null default 0 /* Ngày thanh toán bình quân; Lấy từ B04. Cột J */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                 CREATE INDEX IF NOT EXISTS index_thangpl04a_id_bc ON thangpl04a (id_bc);");
            }
            if (tables.Contains("thangpl04b") == false)
            {
                /* Lấy dữ liệu từ biểu pl04b của cơ sở trong tháng */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl04b (id INTEGER primary key AUTOINCREMENT
                ,id_bc text not null /* liên kết ID table lưu dữ liệu cho báo cáo docx. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,mahanhchinh text not null default '' /*  */
                ,tenhanhchinh text not null default '' /*  */
                ,ma_vung text not null default '' /* Mã vùng */
                ,chi_bq_xn real not null default 0 /* chi BQ Xét nghiệm; đơn vị tính : đồng	Lấy từ B04 . Cột D */
                ,chi_bq_cdha real not null default 0 /* chi BQ Chẩn đoán hình ảnh; Lấy từ B04. Cột E */
                ,chi_bq_thuoc real not null default 0 /* chi BQ thuốc; Lấy từ B04. Cột F */
                ,chi_bq_pttt real not null default 0 /* chi BQ phẫu thuật thủ thuật	Lấy từ B04. Cột G */
                ,chi_bq_vtyt real not null default 0 /* chi BQ vật tư y tế; Lấy từ B04. Cột H */
                ,chi_bq_giuong real not null default 0 /* chi BQ tiền giường; Lấy từ B04. Cột I */
                ,ngay_ttbq real not null default 0 /* Ngày thanh toán bình quân; Lấy từ B04. Cột J */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                 CREATE INDEX IF NOT EXISTS index_thangpl04b_id_bc ON thangpl04b (id_bc);");
            }
            if (tsqlCreate.Count > 0) { dbConnect.Execute(string.Join(Environment.NewLine, tsqlCreate)); }
        }

        public static void CreateImportBcThang(this dbSQLite dbConnect, List<string> tables = null)
        {
            if (tables == null) { tables = dbConnect.getAllTables(); }
            var tsqlCreate = new List<string>();
            if (tables.Contains("bcthang_dutoangiao") == false)
            {
                /**
                 Yêu cầu nhập excel từ người dùng
                 */
                /* PHỤ LỤC 01. TÌNH HÌNH SỬ DỤNG DỰ TOÁN THEO HỢP ĐỒNG (luy kế năm của csyt) */
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangpl01 (id INTEGER primary key AUTOINCREMENT
                ,thang text not null /* Tháng báo cáo tháng. */
                ,nam text not null /* Tháng báo cáo tháng. */
                ,idtinh text not null /* Mã tỉnh của người dùng, để chia dữ liệu riêng từng tỉnh cho các nhóm người dùng từng tỉnh. */
                ,ma_cskcb text not null /* Mã cơ sơ KCB */
                ,ten_cskcb text not null default '' /* Tên cskcb*/
                ,dutoangiao real not null default 0 /* Dự toán tạm giao */
                ,tien_bhtt real not null default 0 /* Tiền T- BHTT cột R-B02-10 */
                ,tl_sudungdt real not null default 0 /* Tỷ lệ sử dụng dự toán = (tien_bhtt/dutoangiao)*100  */
                ,userid text not null default '' /* Lưu ID của người dùng */);
                CREATE INDEX IF NOT EXISTS index_thangpl01_id_bc ON thangpl01 (id_bc);
                CREATE INDEX IF NOT EXISTS index_thangpl01_namthang ON thangpl01 (nam, thang);");
            }

            /* B01. Sử dụng dự toán chi KCB tại các tỉnh, TP */
            if (tables.Contains("thangb01") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb01 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,tu_thang integer not null default 0
                ,den_thang integer not null default 0
                ,nam integer not null default 0
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_thangb01_id_bc ON thangb01 (id_bc);");
            }
            if (tables.Contains("thangb01chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb01chitiet (id INTEGER primary key AUTOINCREMENT
                ,id2 integer not null default 0
                ,ma_tinh text not null default ''
                ,ten_tinh text not null default ''
                ,ma_cskcb text not null default ''
                ,ten_cskcb text not null default ''
                ,ma_vung text not null default ''
                ,dtcsyt_trongnam real not null default 0
                ,dtcsyt_conlai real not null default 0
                ,dtcsyt_chikcb real not null default 0
                ,dtcsyt_tlsudungthang real not null default 0
                ,dtcsyt_tlsudungnam real not null default 0
                ,dtnam_tongchikcb real not null default 0
                ,dtnam_dkbd real not null default 0
                ,dtnam_noitinh real not null default 0
                ,dtnam_ngoaitinh real not null default 0
                ,dtnam_tttt real not null default 0
                ,dtnam_ttho real not null default 0
                ,dtnam_cskcb real not null default 0
                ,dtnam_tongdt real not null default 0
                ,giamtru_tien real not null default 0
                ,giamtru_tl real not null default 0
                ,arv real not null default 0
                ,id_bc text not null default '');
                CREATE INDEX IF NOT EXISTS index_thangb01chitiet_id_bc ON thangb01chitiet (id_bc);");
            }
            /* B02. Thống kê KCB (Tháng) */
            if (tables.Contains("thangb02") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb02 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,ma_loai_kcb text not null
                ,tu_thang integer not null default 0
                ,den_thang integer not null default 0
                ,nam integer not null default 0
                ,loai_bv integer not null default 0
                ,kieubv integer not null default 0
                ,loaick integer not null default 0
                ,hang_bv integer not null default 0
                ,tuyen integer not null default 0
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_thangb02_id_bc ON thangb02 (id_bc);");
            }
            if (tables.Contains("thangb02chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb02chitiet (id INTEGER primary key AUTOINCREMENT
                ,id2 integer not null default 0
                ,ma_tinh text not null default ''
                ,ten_tinh text not null default ''
                ,ma_cskcb text not null default ''
                ,ten_cskcb text not null default ''
                ,ma_vung text not null default ''
                ,tong_luot integer not null default 0
                ,tong_luot_ngoai integer not null default 0
                ,tong_luot_noi integer not null default 0
                ,tyle_noitru real not null default 0
                ,ngay_dtri_bq real not null default 0
                ,chi_bq_chung real not null default 0
                ,chi_bq_ngoai real not null default 0
                ,chi_bq_noi real not null default 0
                ,tong_chi real not null default 0
                ,ty_trong real not null default 0
                ,tong_chi_ngoai real not null default 0
                ,ty_trong_kham real not null default 0
                ,tong_chi_noi real not null default 0
                ,ty_trong_giuong real not null default 0
                ,t_bhtt real not null default 0
                ,t_bhtt_noi real not null default 0
                ,t_bhtt_ngoai real not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_thangb02chitiet_id_bc ON thangb02chitiet (id_bc);");
            }

            /* B04. Thống kê chi bình quân (Tháng) */
            if (tables.Contains("thangb04") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb04 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,tu_thang integer not null default 0
                ,den_thang integer not null default 0
                ,nam integer not null default 0
                ,ma_loai_kcb integer not null default 0
                ,loai_bv integer not null default 0
                ,hang_bv integer not null default 0
                ,tuyen integer not null default 0
                ,kieubv integer not null default 0
                ,loaick integer not null default 0
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_thangb04_id_bc ON thangb04 (id_bc);");
            }
            if (tables.Contains("thangb04chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb04chitiet (id INTEGER primary key AUTOINCREMENT
                ,id2 integer not null default 0
                ,ma_tinh text not null default ''
                ,ten_tinh text not null default ''
                ,ma_cskcb text not null default ''
                ,ten_cskcb text not null default ''
                ,chi_bq_luotkcb real not null default 0
                ,bq_xn real not null default 0
                ,bq_cdha real not null default 0
                ,bq_thuoc real not null default 0
                ,bq_ptt real not null default 0
                ,bq_vtyt real not null default 0
                ,bq_giuong real not null default 0
                ,ngay_ttbq real not null default 0
                ,ma_vung text not null default ''
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_thangb04chitiet_id_bc ON thangb04chitiet (id_bc);");
            }

            /* B21. Theo dõi chỉ tiêu giám sát cơ bản */
            if (tables.Contains("thangb21") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb26 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,nam integer not null default 0
                ,tu_thang integer not null default 0
                ,den_thang integer not null default 0
                ,ma_lydo text not null default ''
                ,loai_bv text not null default ''
                ,hang_bv text not null default ''
                ,tuyen_bv text not null default ''
                ,kieu_bv text not null default ''
                ,loai_kc text not null default ''
                ,loai_kcb integer not null default 0
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_thangb21_id_bc ON thangb21 (id_bc);");
            }
            if (tables.Contains("thangb21chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb21chitiet (id INTEGER primary key AUTOINCREMENT
                ,id2 integer not null default 0
                ,ma_tinh text not null default ''
                ,ten_tinh text not null default ''
                ,ma_cskcb text not null default ''
                ,ten_cskcb text not null default ''
                ,ma_vung text not null default ''

                ,slkcb_trongky real not null default 0
                ,slkcb_tlkytruoc real not null default 0
                ,slkcb_kytruoc real not null default 0
                ,slkcb_trongky real not null default 0
                ,slkcb_tlnamtruoc real not null default 0
                ,slkcb_namtruoc real not null default 0

                ,tongchi_trongky real not null default 0
                ,tongchi_tlkytruoc real not null default 0
                ,tongchi_kytruoc real not null default 0
                ,tongchi_trongky real not null default 0
                ,tongchi_tlnamtruoc real not null default 0
                ,tongchi_namtruoc real not null default 0

                ,tienbhtt_trongky real not null default 0
                ,tienbhtt_tlkytruoc real not null default 0
                ,tienbhtt_kytruoc real not null default 0
                ,tienbhtt_trongky real not null default 0
                ,tienbhtt_tlnamtruoc real not null default 0
                ,tienbhtt_namtruoc real not null default 0

                ,chibq_trongky real not null default 0
                ,chibq_tlkytruoc real not null default 0
                ,chibq_kytruoc real not null default 0
                ,chibq_trongky real not null default 0
                ,chibq_tlnamtruoc real not null default 0
                ,chibq_namtruoc real not null default 0

                ,tlvvnoitru_trongky real not null default 0
                ,tlvvnoitru_tlkytruoc real not null default 0
                ,tlvvnoitru_kytruoc real not null default 0
                ,tlvvnoitru_trongky real not null default 0
                ,tlvvnoitru_tlnamtruoc real not null default 0
                ,tlvvnoitru_namtruoc real not null default 0

                ,ngaydtbq_trongky real not null default 0
                ,ngaydtbq_tlkytruoc real not null default 0
                ,ngaydtbq_kytruoc real not null default 0
                ,ngaydtbq_trongky real not null default 0
                ,ngaydtbq_tlnamtruoc real not null default 0
                ,ngaydtbq_namtruoc real not null default 0

                ,ngaygiuong_trongky real not null default 0
                ,ngaygiuong_tlkytruoc real not null default 0
                ,ngaygiuong_kytruoc real not null default 0
                ,ngaygiuong_trongky real not null default 0
                ,ngaygiuong_tlnamtruoc real not null default 0
                ,ngaygiuong_namtruoc real not null default 0

                ,id_bc text not null default '');
                CREATE INDEX IF NOT EXISTS index_thangb21chitiet_id_bc ON thangb21chitiet (id_bc);");
            }

            /* B26. Thống kê gia tăng chi phí KCB BHYT theo NĐ75 (theo ngày nhận) */
            if (tables.Contains("thangb26") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb26 (id INTEGER primary key AUTOINCREMENT
                ,ma_tinh text not null
                ,loai_kcb text not null default ''
                ,thoigian integer not null default 0
                ,loai_bv integer not null default 0
                ,kieubv integer not null default 0
                ,loaick integer not null default 0
                ,hang_bv integer not null default 0
                ,tuyen integer not null default 0
                ,loai_so_sanh text not null default ''
                ,cs integer not null default 0
                ,userid text not null default ''
                ,timeup integer not null default 0
                ,id_bc text not null default '');
                 CREATE INDEX IF NOT EXISTS index_thangb26_id_bc ON thangb26 (id_bc);");
            }
            if (tables.Contains("thangb26chitiet") == false)
            {
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS thangb26chitiet (id INTEGER primary key AUTOINCREMENT,
                id2 integer not null default 0,
                ma_tinh text not null default '',
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
                ,id_bc text not null default '');
                CREATE INDEX IF NOT EXISTS index_thangb26chitiet_id_bc ON thangb26chitiet (id_bc);");
            }
            if (tsqlCreate.Count > 0) { dbConnect.Execute(string.Join(Environment.NewLine, tsqlCreate)); }
        }
    }
}