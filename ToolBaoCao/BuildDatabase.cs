using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace ToolBaoCao
{
    public static class BuildDatabase
    {
        public static void buildDataMain(this dbSQLite connect)
        {
            var tsqlInsert = new List<string>();
            var tsqlCreate = new List<string>();
            var tsql = "";
            var tables = connect.getAllTables();
            /** Nhóm quản lý web */
            if (tables.Contains("taikhoan") == false)
            {
                tsqlCreate.Add("CREATE TABLE IF NOT EXISTS taikhoan(iduser TEXT NOT NULL Primary Key, mat_khau TEXT NOT NULL, ten_hien_thi TEXT NOT NULL, gioi_tinh TEXT NOT NULL DEFAULT '', ngay_sinh TEXT NOT NULL Default '', email TEXT NOT NULL Default '', dien_thoai TEXT NOT NULL Default '', dia_chi TEXT NOT NULL Default '', hinh_dai_dien TEXT NOT NULL Default '', idtinh text not null default '', ghi_chu TEXT NOT NULL DEFAULT '', vitrilamviec text not null default '', nhom INTEGER NOT NULL default -1, locked INTEGER NOT NULL default 0, time_create double not null default 0, time_last_login double not null default 0);");
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
                tsqlCreate.Add(@"CREATE TABLE IF NOT EXISTS wmenu(id integer PRIMARY KEY, title text NOT NULL DEFAULT '', link text NOT NULL DEFAULT '', idfather integer NOT NULL DEFAULT -1, paths text NOT NULL DEFAULT '', postion integer NOT NULL DEFAULT 0, note text NOT NULL DEFAULT '', css text NOT NULL DEFAULT '' );");
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

        public static dbSQLite getDBUserOnline()
        {
            string pathData = AppHelper.pathApp + "App_Data\\useronline.db";
            dbSQLite db = new dbSQLite(pathData);
            if (File.Exists(pathData) == false)
            {
                try
                {
                    db.Execute(@"CREATE TABLE IF NOT EXISTS useronline (
                        userid TEXT NOT NULL,
                        time1 INTEGER NOT NULL DEFAULT 0,
                        time2 INTEGER NOT NULL DEFAULT 0,
                        ten_hien_thi TEXT NOT NULL DEFAULT '',
                        ip TEXT NOT NULL DEFAULT '',
                        [local] TEXT NOT NULL DEFAULT '', PRIMARY KEY (userid, ip));");
                }
                catch { }
            }
            return db;
        }

        public static dbSQLite getDbSQLiteBaoCao(string matinh = "")
        {
            string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", $"BaoCao{matinh}.db");
            var db = new dbSQLite(pathDB);
            var tables = db.getAllTables();
            db.CreateTablePhucLucBaoCao(tables);
            db.CreateTableBaoCao(tables);
            return db;
        }

        public static dbSQLite getDataBaoCaoTuan(string matinh = "")
        {
            string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", $"BaoCaoTuan{matinh}.db");
            var db = new dbSQLite(pathDB);
            var tables = db.getAllTables();
            db.CreateTablePhucLucBaoCao(tables);
            db.CreateTableBaoCao(tables);
            return db;
        }

        public static dbSQLite getDataImportBaoCaoTuan(string matinh = "")
        {
            string pathDB = Path.Combine(AppHelper.pathApp, "App_Data", $"ImportBaoCaoTuan{matinh}.db");
            var db = new dbSQLite(pathDB);
            var tables = db.getAllTables();
            db.CreateTableImport(tables);
            return db;
        }

        public static dbSQLite getDbSQLiteImport(string matinhOrPath)
        {
            if (matinhOrPath == "") { matinhOrPath = Path.Combine(AppHelper.pathApp, "App_Data", "import.db"); }
            else
            {
                if (Regex.IsMatch(matinhOrPath, @"^\d+$")) { matinhOrPath = Path.Combine(AppHelper.pathApp, "App_Data", $"import{matinhOrPath}.db"); }
            }
            var db = new dbSQLite(matinhOrPath);
            db.CreateTableImport();
            return db;
        }

        public static void buildDataWork(this dbSQLite dbConnect)
        {
            var tables = dbConnect.getAllTables();
            /* Các bảng Import */
            dbConnect.CreateTableImport(tables);
            /* Các bảng phục lục công việc */
            dbConnect.CreateTablePhucLucBaoCao(tables);
            dbConnect.CreateTableBaoCao(tables);
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
            dbConnect.Execute(@"CREATE TABLE IF NOT EXISTS wprocess (
                  id text NOT NULL PRIMARY KEY
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

        public static void CreateTableBaoCao(this dbSQLite dbConnect, List<string> tables = null)
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
            if (tsqlCreate.Count > 0) { dbConnect.Execute(string.Join(Environment.NewLine, tsqlCreate)); }
        }

        public static void CreateTablePhucLucBaoCao(this dbSQLite dbConnect, List<string> tables = null)
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
                ,tong_luot real not null default 0
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

        public static void CreateTableImport(this dbSQLite dbConnect, List<string> tables = null)
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
    }
}