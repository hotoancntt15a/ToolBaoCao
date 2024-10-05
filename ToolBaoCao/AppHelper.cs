using NPOI.SS.UserModel;
using Org.BouncyCastle.Asn1.Ocsp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using ToolBaoCao.CaptchaImage;
using UAParser;

namespace ToolBaoCao
{
    public static class AppHelper
    {
        public static List<string> listKeyConfigCrypt = new List<string>() { "" };
        private static readonly string keyMD5 = typeof(AppHelper).Namespace;
        public static AppConfig appConfig = new AppConfig();
        public static readonly string pathApp = AppDomain.CurrentDomain.BaseDirectory;
        public static readonly string pathAppData = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data");
        public static readonly string pathTemp = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp");
        public static readonly string pathCache = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cache");
        public static readonly string pathCodeProject = Assembly.GetExecutingAssembly().GetPathCodeProject();
        public static readonly string projectTitle = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;
        public static readonly string projectName = typeof(AppHelper).Namespace;
        public static dbSQLite dbSqliteMain = new dbSQLite();
        public static dbSQLite dbSqliteWork = new dbSQLite();

        public static bool ExtractFileZip(string zipFilePath, string extractFolderPath, string ext = "", int allFileExt = 0)
        {
            bool dbFileFound = false;
            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                if (ext == "") { archive.ExtractToDirectory(extractFolderPath); dbFileFound = true; }
                else
                {
                    if (allFileExt > 0)
                    {
                        int i = 0;
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            if (entry.FullName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                            {
                                i++;
                                dbFileFound = true; 
                                entry.ExtractToFile(Path.Combine(extractFolderPath, entry.FullName), overwrite: true);
                                if (i > allFileExt) { break; }
                            }
                        }
                    }
                    else
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            if (entry.FullName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                            {
                                dbFileFound = true; 
                                entry.ExtractToFile(Path.Combine(extractFolderPath, entry.FullName), overwrite: true);
                            }
                        }
                    }
                }
            }
            return dbFileFound;
        }

        public static string SQLiteLike(this string field, string value) => dbSqliteMain.like(field, value);

        public static string GetUserIpAddress(this HttpContext http)
        {
            string ipAddress = http.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (string.IsNullOrEmpty(ipAddress)) { ipAddress = http.Request.ServerVariables["REMOTE_ADDR"]; }
            // Trường hợp có nhiều địa chỉ IP trong X-Forwarded-For, lấy địa chỉ đầu tiên
            if (!string.IsNullOrEmpty(ipAddress) && ipAddress.Contains(",")) { ipAddress = ipAddress.Split(',')[0].Trim(); }
            return ipAddress;
        }

        public static string GetUserBrowserInfo(this HttpContext http)
        {
            var uaParser = Parser.GetDefault();
            ClientInfo clientInfo = uaParser.Parse(http.Request.UserAgent);
            return $"{clientInfo.UA.Family} ({http.Request.UserAgent})";
        }

        public static string GetUserBrowser(this HttpContext http)
        {
            var uaParser = Parser.GetDefault();
            ClientInfo clientInfo = uaParser.Parse(http.Request.UserAgent);
            return $"{clientInfo.UA.Family} - {clientInfo.OS.Family} {clientInfo.OS.Major}";
        }

        public static string getMenuLeft(string nhom = "3")
        {
            if (Regex.IsMatch(nhom, @"^\d+$") == false) { nhom = "3"; }
            if (nhom == "0") { nhom = "1"; }
            /* mặc định nhóm người dùng */
            string fileCahce = Path.Combine(pathCache, $"menuleft_dmnhom_wmenu_{nhom}.tpl");
            if (File.Exists(fileCahce)) { return File.ReadAllText(fileCahce); }
            /* Lấy idmenu Father */
            var idFather = $"{dbSqliteMain.getValue($"SELECT idwmenu FROM dmnhom WHERE id={nhom}")}";
            if (Regex.IsMatch(idFather, @"^\d+$") == false)
            {
                return "<li class=\"nav-item\"> <hr class=\"sidebar-divider d-none d-md-block\" /> </li><li class=\"nav-item\"> <div class=\"sidebar-heading\"> Bạn chưa được cấp quyền </div> </li>";
            }
            var dataMenu = dbSqliteMain.getDataTable("SELECT * FROM wmenu ORDER BY postion");
            if (dataMenu.Rows.Count == 0)
            {
                return "<li class=\"nav-item\"> <hr class=\"sidebar-divider d-none d-md-block\" /> </li><li class=\"nav-item\"> <div class=\"sidebar-heading\"> Không có dữ liệu menu </div> </li>";
            }
            var tpl = getMenuLeft2(long.Parse(idFather), dataMenu);
            if (tpl != "") { tpl = $"<!-- {nhom} -->" + tpl; File.WriteAllText(fileCahce, tpl); }
            return tpl;
        }

        private static string getMenuLeft2(long idFather, DataTable dataMenu)
        {
            var menu = dataMenu.AsEnumerable().Where(r => r.Field<long>("idfather") == idFather).OrderBy(r => r.Field<long>("postion")).ToList();
            if (menu.Count == 0) { return ""; }
            var li = new List<string>();
            var dt = dataMenu.Clone();
            foreach (DataRow r in menu) { dt.ImportRow(r); }
            foreach (DataRow r in dt.Rows)
            {
                var link = $"{r["link"]}".Trim();
                var css = $"{r["css"]}".Trim(); if (css != "") { css = $"<i class=\"{css}\"></i> "; }
                if (link == "")
                {
                    li.Add("<li class=\"nav-item\"> <hr class=\"sidebar-divider d-none d-md-block\" /> </li>");
                    li.Add($"<li class=\"nav-item\"> <div class=\"sidebar-heading\"> {css}{r["title"]}</div></li>");
                }
                else
                {
                    li.Add($"<li class=\"nav-item\"> <a class=\"nav-link\" href=\"{link}\" title=\"{r["note"]}\"> {css}<span>{r["title"]}</span></a></li>");
                }
                li.Add(getMenuLeft2((long)r["id"], dataMenu));
            }
            return string.Join("", li);
        }

        /// <summary>
        /// Kiểm tra số định dạng US, không phân biệt số âm số dương
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static bool isNumberUS(this string numberUS)
        {
            return Regex.IsMatch(numberUS, @"^-?\d+(.\d+)?$");
        }

        /// <summary>
        /// Chỉ kiểm tra số dương định dạng US
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static bool isNumberUSDouble(this string numberUS)
        {
            return Regex.IsMatch(numberUS, @"^\d+(.\d+)?$");
        }

        /// <summary>
        /// Chỉ kiểm tra số nguyên dương định dạng US
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static bool isNumberUSInt(this string numberUS)
        {
            return Regex.IsMatch(numberUS, @"^\d+$");
        }

        /// <summary>
        /// Định dạng số US; không đúng định dạng trả lại numberUS; Số > triệu = Số tròn triệu đồng; Số > nghìn = Số tròn nghìn đồng; = Số tròn đồng
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static string lamTronTrieuDong(this string numberUS)
        {
            if (Regex.IsMatch(numberUS, @"^-?\d+(.\d+)?$") == false) { return numberUS; }
            if (numberUS.Contains(".")) { numberUS = numberUS.Split('.')[0]; }
            double so = double.Parse(numberUS);
            if (so > 1000000) { so = Math.Round(so / 1000000, 0); return $"{so}000000"; }
            if (so > 1000) { so = Math.Round(so / 1000, 0); return $"{so}000"; }
            return so.ToString();
        }

        /// <summary>
        /// Số > triệu = Số tròn triệu đồng; Số > nghìn = Số tròn nghìn đồng; = Số tròn đồng
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static double lamTronTrieuDong(this double numberUS)
        {
            if (numberUS > 1000000) { numberUS = Math.Round(numberUS / 1000000, 0); return double.Parse($"{numberUS}000000"); }
            if (numberUS > 1000) { numberUS = Math.Round(numberUS / 1000, 0); return double.Parse($"{numberUS}000"); }
            return Math.Round(numberUS, 0);
        }

        /// <summary>
        /// Định dạng số US; không đúng định dạng trả lại numberUS; Số > nghìn = Số tròn nghìn đồng; = Số tròn đồng
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static string lamTronNghinDong(this string numberUS)
        {
            if (Regex.IsMatch(numberUS, @"^-?\d+(.\d+)?$") == false) { return numberUS; }
            if (numberUS.Contains(".")) { numberUS = numberUS.Split('.')[0]; }
            double so = double.Parse(numberUS);
            if (so > 1000) { so = Math.Round(so / 1000, 0); return $"{so}000"; }
            return so.ToString();
        }

        /// <summary>
        /// Số > nghìn = Số tròn nghìn đồng; = Số tròn đồng
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static double lamTronNghinDong(this double numberUS)
        {
            if (numberUS > 1000) { numberUS = Math.Round(numberUS / 1000, 0); return double.Parse($"{numberUS}000"); }
            return Math.Round(numberUS, 0);
        }

        /// <summary>
        /// Trả về dạng {numberUS}{b/kb/mb/gb}
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static string getFileSize(this long size)
        {
            if (size > 1073741824) { return $"{(size / 1073741824):0.##}Gb"; }
            if (size > 1048576) { return $"{(size / 1048576):0.##}Mb"; }
            if (size > 1024) { return $"{(size / 1024):0.##}Kb"; }
            return $"{size}b";
        }

        /// <summary>
        /// Trả về dạng {numberUS}{b/kb/mb/gb}
        /// </summary>
        /// <param name="numberUS"></param>
        /// <returns></returns>
        public static string getFileSize(this int size)
        {
            if (size > 1073741824) { return $"{(size / 1073741824):0.##}Gb"; }
            if (size > 1048576) { return $"{(size / 1048576):0.##}Mb"; }
            if (size > 1024) { return $"{(size / 1024):0.##}Kb"; }
            return $"{size}b";
        }

        public static long toTimestamp(this DateTime time) => ((DateTimeOffset)time).ToUnixTimeSeconds();

        /* Việt Nam múi giờ GMT +7 */

        public static DateTime toDateTime(this long timestamp, int GMT = 7) => (DateTimeOffset.FromUnixTimeSeconds(timestamp).DateTime).AddHours(GMT);

        public static DateTime toDateTime(this string timestamp, int GMT = 7)
        {
            if (Regex.IsMatch(timestamp, "^[0-9]+$")) { return (long.Parse(timestamp).toDateTime(7)); }
            return new DateTime(1970, 1, 1);
        }

        public static string getTimeRun(this DateTime timeStart)
        {
            var t = DateTime.Now - timeStart;
            if (t.Days > 0) { return $"{t.Days} ngày {t.Hours}:{t.Minutes}:{t.Seconds}"; }
            if (t.Hours > 0) { return $"{t.Hours}:{t.Minutes}:{t.Seconds}"; }
            if (t.Minutes > 0) { return $"{t.Minutes}:{t.Seconds}"; }
            if (t.Seconds > 0) { return $"{t.Seconds},{t.Milliseconds.ToString().Substring(0, 2)} giây"; }
            return $"0,{t.Milliseconds} giây";
        }

        public static string FormatCultureVN(this string numberUS, int decimalDigits = 2)
        {
            if (Regex.IsMatch(numberUS, @"^-?\d+(\.\d+)?$") == false) { return numberUS; }
            CultureInfo vietnamCulture = new CultureInfo("vi-VN");
            NumberFormatInfo formatInfo = vietnamCulture.NumberFormat;
            if (Decimal.TryParse(numberUS, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal parsedNumber))
            {
                string formattedNumber = parsedNumber.ToString($"N{decimalDigits}", formatInfo);
                if (formattedNumber.Contains(","))
                {
                    string[] parts = formattedNumber.Split(',');
                    string decimalPart = parts[1].TrimEnd('0');
                    if (decimalPart.Length > 0) { return parts[0] + "," + decimalPart; }
                    return parts[0];
                }
                return formattedNumber;
            }
            return numberUS;
        }

        public static string FormatCultureVN(this long numberUS)
        {
            CultureInfo vietnamCulture = new CultureInfo("vi-VN");
            NumberFormatInfo formatInfo = vietnamCulture.NumberFormat;
            return numberUS.ToString($"N", formatInfo);
        }

        public static string FormatCultureVN(this int numberUS)
        {
            CultureInfo vietnamCulture = new CultureInfo("vi-VN");
            NumberFormatInfo formatInfo = vietnamCulture.NumberFormat;
            return numberUS.ToString($"N", formatInfo);
        }

        public static string FormatCultureVN(this double numberUS, int decimalDigits = 2)
        {
            CultureInfo vietnamCulture = new CultureInfo("vi-VN");
            NumberFormatInfo formatInfo = vietnamCulture.NumberFormat;
            string formattedNumber = numberUS.ToString($"N{decimalDigits}", formatInfo);
            if (formattedNumber.Contains(","))
            {
                string[] parts = formattedNumber.Split(',');
                string decimalPart = parts[1].TrimEnd('0');
                if (decimalPart.Length > 0) { return parts[0] + "," + decimalPart; }
                return parts[0];
            }
            return formattedNumber;
        }

        public static string FormatCultureVN(this decimal numberUS, int decimalDigits = 2)
        {
            CultureInfo vietnamCulture = new CultureInfo("vi-VN");
            NumberFormatInfo formatInfo = vietnamCulture.NumberFormat;
            return numberUS.ToString($"N{decimalDigits}", formatInfo);
        }

        public static List<string> GetTableNameFromTsql(string tsql)
        {
            var matches = Regex.Matches(tsql, @"\b(FROM|JOIN|UPDATE)\s+([a-zA-Z0-9_.\[\]]+)", RegexOptions.IgnoreCase);
            var tableNames = new List<string>();
            foreach (System.Text.RegularExpressions.Match match in matches) { tableNames.Add(match.Groups[2].Value); }
            return tableNames;
        }

        public static bool IsUpdateOrDelete(string sql) => Regex.IsMatch(sql, @"^\s*(UPDATE|DELETE)\s+", RegexOptions.IgnoreCase);

        public static string GetPathFileCacheQuery(string tsql, string dataName)
        {
            var tables = GetTableNameFromTsql(tsql);
            string fileCache = string.Empty;
            if (tables.Count == 1)
            {
                if (Regex.IsMatch(tables[0], @"^dm", RegexOptions.IgnoreCase)) { fileCache = GetMd5Hash(tsql); }
                else
                {
                    var tablesCache = new List<string> { "phanquyen", "nhomquyen", "w_menu", "wmenu", "taikhoan", "system_var" };
                    if (tablesCache.Contains(tables[0])) { fileCache = tsql.GetMd5Hash(); }
                }
                if (!string.IsNullOrEmpty(fileCache)) { fileCache = Path.Combine(pathCache, $"d{dataName}_{tables[0]}_query_{fileCache}.xml"); }
            }
            return fileCache;
        }

        public static void DeleteFileCacheQuery(string tsql, string dataName)
        {
            if (!IsUpdateOrDelete(tsql)) return;
            var tables = GetTableNameFromTsql(tsql);
            if (tables.Count == 0) return;
            DeleteCache(tables[0] + "_");
        }

        public static void DeleteCache(string likeName)
        {
            if (string.IsNullOrEmpty(likeName)) return;
            var files = Directory.GetFiles(pathCache, $"*{likeName}*.*");
            foreach (var file in files)
            {
                if (File.Exists(file)) { try { File.Delete(file); } catch { } }
            }
        }

        public static string GetValueAsString(this ICell cell, string formatDateTime = "yyyy-MM-dd H:mm:ss")
        {
            if (cell == null) { return ""; }
            switch (cell.CellType)
            {
                case CellType.Error: return FormulaError.ForInt(cell.ErrorCellValue).String;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell)) { return cell.DateCellValue?.ToString(formatDateTime); }
                    return cell.NumericCellValue.ToString().Replace(",", ".");

                case CellType.Formula:
                    // Lấy giá trị tính toán của công thức nếu cần
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.Numeric:
                            if (DateUtil.IsCellDateFormatted(cell)) { return cell.DateCellValue?.ToString(formatDateTime); }
                            return cell.NumericCellValue.ToString().Replace(",", ".");

                        case CellType.Error: return FormulaError.ForInt(cell.ErrorCellValue).String;
                        default: return $"{cell}";
                    }
                default: return $"{cell}";
            }
        }

        public static string getValueFieldTSQL(string valueField) => valueField.Replace("'", "''");

        public static string getConfig(string key, string valueDefault = "") => appConfig.Get(key, valueDefault);

        public static void LoadStart()
        {
            appConfig = new AppConfig(Path.Combine(pathApp, "config.json"));
            if (appConfig.Config.Settings.Count == 0)
            {
                appConfig.Set("App.Title", "Công cụ phân tích bảo hiểm y tế");
                appConfig.Set("App.PageSize", "50");
                appConfig.Set("App.PacketSize", "1000");
            }
            dbSqliteMain = new dbSQLite(Path.Combine(pathApp, @"App_Data\main.db"));
            dbSqliteMain.buildDataMain();
            dbSqliteWork = new dbSQLite(Path.Combine(pathApp, @"App_Data\data.db"));
            dbSqliteWork.buildDataWork();
            /* Check Folder Exists */
            var folders = new List<string>() {
                Path.Combine(pathApp, "cache")
                , Path.Combine(pathApp, "temp") };
            foreach (var pathFolder in folders) { if (Directory.Exists(pathFolder) == false) { Directory.CreateDirectory(pathFolder); } }
        }

        public static void SapXepNgauNhien(this List<string> arr)
        {
            Random rnd = new Random();
            int n = arr.Count;
            while (n > 1)
            {
                int k = rnd.Next(n--);
                (arr[k], arr[n]) = (arr[n], arr[k]);
            }
        }

        public static string CreateMenuTop(string url, string date, string content, string ClassbgColor = "bg-primary", bool fontBold = true)
        {
            return $"<a class=\"dropdown-item d-flex align-items-center\" href=\"{url}\">" +
                $"<div class=\"mr-3\"><div class=\"icon-circle {ClassbgColor}\"><i class=\"fas fa-file-alt text-white\"></i></div></div>" +
                $"<div><div class=\"small text-gray-500\">{date}</div>" + (fontBold ? $"<span class=\"font-weight-bold\">{content}</span>" : content) +
                "</div></a>";
        }

        public static bool CheckIsLogin()
        {
            var http = HttpContext.Current;
            if (http == null) return false;
            var tmp = $"{http.Session["app.isLogin"]}";
            if (tmp == "1") return true;
            if (http.Request.Cookies.AllKeys.Any(p => p == "idobject") == false) { return false; }
            tmp = $"{http.Request.Cookies["idobject"]?.Value}";
            /* IDUSER|PASS|DATETIME */
            tmp = tmp.MD5Decrypt();
            var idObject = tmp.Split('|');
            if (idObject.Length != 3) { return false; }
            tmp = setLogin(idObject[0], idObject[1], true);
            if (tmp == "") { return true; }
            return false;
        }

        public static string setLogin(string userName, string passWord, bool remember = false)
        {
            if (userName == "") { return "Tên đăng nhập để trống"; }
            if (passWord == "") { return "Mật khẩu để trống"; }
            string tsql = $"SELECT * FROM taikhoan WHERE iduser = @iduser AND mat_khau='{passWord.GetMd5Hash()}'";
            var http = HttpContext.Current;
            try
            {
                var items = dbSqliteMain.getDataTable(tsql, new KeyValuePair<string, string>("@iduser", userName));
                if (items.Rows.Count == 0)
                {
                    /* Không có tài khoản nào thì tạo tài khoản mặc định */
                    items = dbSqliteMain.getDataTable("SELECT * FROM taikhoan LIMIT 1");
                    if (items.Rows.Count == 0)
                    {
                        var time = DateTime.Now.toTimestamp();
                        dbSqliteMain.Execute($"INSERT INTO admins ([iduser] ,[mat_khau] ,[ten_hien_thi] ,[gioi_tinh] ,[ngay_sinh] ,[email] ,[dien_thoai] ,[dia_chi] ,[hinh_dai_dien] ,[ghi_chu] ,[time_create], nhom) VALUES ('admin', '{"admin123@".GetMd5Hash()}', 'Adminstrator', 'Nam', '{DateTime.Now:dd/MM/yyyy}', 'hotoancntt15a@gmail.com', '09140272795', 'Thành phố Lào Cai, Tỉnh Lào Cai', '', '', '{time}', 0);");
                    }
                    return $"Tài khoản '{userName}' không tồn tại hoặc mật khẩu không đúng";
                }
                if (http == null) { return keyMSG.HttpConnetNull; }
                http.Session.Clear();
                http.Request.Cookies.Clear();

                http.Session[keyMSG.SessionIPAddress] = http.GetUserIpAddress();
                http.Session[keyMSG.SessionBrowserInfo] = http.GetUserBrowser();

                http.Session.Add("app.isLogin", "1");
                foreach (DataColumn c in items.Columns) { http.Session.Add(c.ColumnName, $"{items.Rows[0][c.ColumnName]}"); }
                /* IDUSER|PASS|DATETIME */
                if (remember)
                {
                    HttpCookie c1 = new HttpCookie("idobject", $"{userName}|{passWord}|{DateTime.Now}".MD5Encrypt());
                    c1.Expires = DateTime.Now.AddMonths(1);
                    http.Response.Cookies.Add(c1);
                }
                try
                {
                    var item = new Dictionary<string, string>() { { "iduser", userName }, { "timelogin", DateTime.Now.toTimestamp().ToString() } };
                    dbSqliteMain.Update("logintime", item, "repalce");
                }
                catch { }
            }
            catch (Exception ex) { return $"Lỗi: {ex.Message} <br />Chi tiết: {ex.StackTrace}"; }
            var db = BuildDatabase.getDBUserOnline();
            var tmp = $"{http.Session["ten_hien_thi"]}".sqliteGetValueField();
            db.Execute($"INSERT OR IGNORE INTO useronline (userid, time1, time2, ip, ten_hien_thi, local) VALUES ('{http.Session["iduser"]}',{DateTime.Now.toTimestamp()},{DateTime.Now.toTimestamp()},'{http.Session[keyMSG.SessionIPAddress]}', '{tmp}', '{http.Session[keyMSG.SessionBrowserInfo]}'); UPDATE useronline SET time2={DateTime.Now.toTimestamp()} WHERE userid='{http.Session["iduser"]}' AND ip='{http.Session[keyMSG.SessionIPAddress]}';");
            return "";
        }

        public static string GetMd5Hash(this string input)
        {
            using (MD5 md5Hash = MD5.Create())
            {
                byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input));
                StringBuilder sBuilder = new StringBuilder();
                for (int i = 0; i < data.Length; i++) { sBuilder.Append(data[i].ToString("x2")); }
                return sBuilder.ToString();
            }
        }

        public static string MD5Encrypt(this string planText)
        {
            if (string.IsNullOrEmpty(planText)) return "";
            byte[] keyArray;
            byte[] toEndcry = Encoding.UTF8.GetBytes(planText);
            var md5 = new MD5CryptoServiceProvider();
            keyArray = md5.ComputeHash(Encoding.UTF8.GetBytes(keyMD5));
            var trip = new TripleDESCryptoServiceProvider { Key = keyArray, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 };
            ICryptoTransform tranform = trip.CreateEncryptor();
            byte[] resualArray = tranform.TransformFinalBlock(toEndcry, 0, toEndcry.Length);
            return Convert.ToBase64String(resualArray, 0, resualArray.Length);
        }

        public static string MD5Decrypt(this string cipherText)
        {
            if (string.IsNullOrEmpty(cipherText)) return "";
            cipherText = cipherText.Replace(" ", "+");
            try
            {
                byte[] keyArray;
                byte[] toEndArray = Convert.FromBase64String(cipherText);
                var md5 = new MD5CryptoServiceProvider();
                keyArray = md5.ComputeHash(Encoding.UTF8.GetBytes(keyMD5));
                var trip = new TripleDESCryptoServiceProvider { Key = keyArray, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7 };
                ICryptoTransform tranfrom = trip.CreateDecryptor();
                byte[] resualArray = tranfrom.TransformFinalBlock(toEndArray, 0, toEndArray.Length);
                return Encoding.UTF8.GetString(resualArray);
            }
            catch (Exception ex) { return $"Lỗi: {ex.Message}; Chuỗi nhập '{cipherText}'"; }
        }

        public static string RemoveHTMLTag(this object html)
        {
            /* Xóa các thẻ html */
            if (html is string)
            {
                Regex objRegEx = new Regex("<[^>]*>");
                return objRegEx.Replace((string)html, "");
            }
            return $"{html}";
        }

        public static string RemoveHTMLTagDecode(this object html)
        {
            /* Xóa các thẻ html */
            if (html is string)
            {
                Regex objRegEx = new Regex("<[^>]*>");
                return objRegEx.Replace((string)html, "").htmlDecode();
            }
            return $"{html}";
        }

        public static string RemoveHTMLTag(this string html)
        {
            /* Xóa các thẻ html */
            Regex objRegEx = new Regex("<[^>]*>");
            return objRegEx.Replace(html, "");
        }

        public static string htmlDecode(this string html) => System.Net.WebUtility.HtmlDecode(html);

        public static double toOADateFromVN(this string dateVN)
        {
            var format = "dd/MM/yyyy"; dateVN = dateVN.Trim();
            if (dateVN.Length > 10)
            {
                var tmp = dateVN.Split(':').ToList();
                if (tmp.Count == 2) { format += " HH:mm"; }
                else if (tmp.Count == 3) { format += " HH:mm:ss"; }
                else { return 0; }
            }
            var date = DateTime.Now;
            if (DateTime.TryParseExact(dateVN, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) { return date.ToOADate(); }
            return 0;
        }

        public static DateTime getFromDateVN(this string dateVN)
        {
            var format = "dd/MM/yyyy"; dateVN = dateVN.Trim();
            if (dateVN.Length > 10)
            {
                var tmp = dateVN.Split(':').ToList();
                if (tmp.Count == 2) { format += " HH:mm"; }
                else if (tmp.Count == 3) { format += " HH:mm:ss"; }
                else { return new DateTime(1970, 1, 1); }
            }
            var date = DateTime.Now;
            if (DateTime.TryParseExact(dateVN, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) { return date; }
            return new DateTime(1970, 1, 1);
        }

        public static Dictionary<string, object> toDictionary(this DataTable dataTable)
        {
            var items = new Dictionary<string, object>();
            if (dataTable.Rows.Count == 0) { return items; }
            if (dataTable.Columns.Count > 1) { foreach (DataRow r in dataTable.Rows) { items.Add($"{r[0]}", r[1]); } }
            else { foreach (DataRow r in dataTable.Rows) { items.Add($"{r[0]}", null); } }
            return items;
        }

        public static string getValue(this Dictionary<string, string> data, string key, string defaultValue = "", bool formatVN = false)
        {
            if (data.ContainsKey(key)) { return formatVN ? data[key].FormatCultureVN() : data[key]; }
            return defaultValue;
        }

        public static object getValue(this Dictionary<string, object> data, string key, object defaultValue = null)
        {
            if (data != null && data.ContainsKey(key)) { return data[key]; }
            return defaultValue;
        }

        public static bool isDateVN(this string timeVN)
        {
            return Regex.IsMatch(timeVN, "^[0-3][0-9]/[0-1][0-9]/[1-9][0-9]{3}$|^[0-3][0-9]/[0-1][0-9]/[1-9][0-9]{3} [0-2][0-9]:[0-5][0-9]$|^[0-3][0-9]/[0-1][0-9]/[1-9][0-9]{3} [0-2][0-9]:[0-5][0-9]:[0-5][0-9]$");
        }

        public static bool isDateVN(this string timeVN, out DateTime datetime)
        {
            var format = "dd/MM/yyyy"; timeVN = timeVN.Trim(); datetime = new DateTime(1970, 1, 1);
            if (timeVN.Length > 10)
            {
                var tmp = timeVN.Split(':').ToList();
                if (tmp.Count == 2) { format += " HH:mm"; }
                else if (tmp.Count == 3) { format += " HH:mm:ss"; }
                else { return false; }
            }
            if (DateTime.TryParseExact(timeVN, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out datetime)) { return true; }
            return false;
        }

        public static string connectString = "";

        public static string chuThuongDauChuoi(this string inputString) => inputString.First().ToString().ToLower() + inputString.Substring(1);

        public static void saveError(this Exception ex, string message = "")
        {
            try
            {
                using (var sw = new StreamWriter(HttpContext.Current.Server.MapPath("~/error.log"), true, Encoding.Unicode))
                {
                    try { sw.WriteLine($"{DateTime.Now:dd/MM/yyyy HH:mm:ss} {ex.Message} {message} {ex.StackTrace}"); sw.Flush(); } catch { }
                }
            }
            catch { }
        }

        public static void saveError(string message, string pathSave = "")
        {
            try
            {
                if (pathSave == "") { pathSave = HttpContext.Current.Server.MapPath("~/error.log"); }
                using (var sw = new StreamWriter(pathSave, true, Encoding.Unicode))
                {
                    try { sw.WriteLine($"{DateTime.Now:dd/MM/yyyy HH:mm:ss} {message}"); sw.Flush(); } catch { }
                }
            }
            catch { }
        }

        public static string getLineHTML(this Exception ex, string description = "")
        {
            var msg = ex.StackTrace.Split('\n').Where(p => Regex.IsMatch(p, @":line \d+")).ToList();
            var s2 = string.Join(" <br /> ", msg).Replace(pathCodeProject, "");
            if (string.IsNullOrEmpty(description)) return $"{ex.Message} <br /> {s2}";
            return $"Lỗi {description}: {ex.Message} <br /> {s2}";
        }

        public static string getErrorSave(this Exception ex, string newLineReturn = "<br />")
        {
            if (newLineReturn == "") { newLineReturn = Environment.NewLine; }
            /* Chỉ lấy dòng có chỉ số dòng */
            var stackTrace = ex.StackTrace.Replace(pathCodeProject, "");
            var msg = stackTrace.Split('\n').Where(p => Regex.IsMatch(p, @":line \d+")).ToList();
            try
            {
                using (var sw = new StreamWriter(HttpContext.Current.Server.MapPath("~/error.log"), true, Encoding.Unicode))
                {
                    try { sw.WriteLine($"{DateTime.Now:dd/MM/yyyy HH:mm:ss} {ex.Message}{Environment.NewLine}{string.Join(Environment.NewLine, msg)}"); sw.Flush(); } catch { }
                }
            }
            catch { }
            return $"Lỗi: {ex.Message}{newLineReturn}Chi tiết:{string.Join(newLineReturn, msg)}";
        }

        public static string getValue(this HttpRequestBase r, string key, string def = "")
        { return r[key] == null ? def : r[key].Trim(); }

        /* Request */

        public static string getIpAddress(this HttpRequestBase r) => r.ServerVariables["REMOTE_ADDR"];

        public static string getIpAddress(this HttpRequest r) => r.ServerVariables["REMOTE_ADDR"];

        public static string getBrowerName(this HttpRequest rq) => getBrowerName(rq.Browser.Browser, $"{rq.ServerVariables["HTTP_USER_AGENT"]}");

        public static string getBrowerName(this HttpRequestBase rq) => getBrowerName(rq.Browser.Browser, $"{rq.ServerVariables["HTTP_USER_AGENT"]}");

        public static string getBrowerName(string browser, string HTTP_USER_AGENT)
        {
            if (string.IsNullOrEmpty(HTTP_USER_AGENT)) { return browser; }
            /* var m = Regex.Matches(s, "[a-zA-Z_]+/[0-9]+") */
            if (HTTP_USER_AGENT.Contains("Edge/")) { return "Edge"; }
            if (HTTP_USER_AGENT.Contains("OPR/")) { return "Opera"; }
            if (HTTP_USER_AGENT.Contains("Firefox/")) { return "Firefox"; }
            if (HTTP_USER_AGENT.Contains("coc_coc_browser/")) { return "Cốc Cốc"; }
            if (HTTP_USER_AGENT.Contains("UCBrowser/")) { return "UCBrowser"; }
            if (HTTP_USER_AGENT.Contains("Falkon/")) { return "Falkon"; }
            if (HTTP_USER_AGENT.Contains("K-Meleon/")) { return "K-Meleon"; }
            if (HTTP_USER_AGENT.Contains("QupZilla/")) { return "QupZilla"; }
            if (HTTP_USER_AGENT.Contains("YaBrowser/")) { return "Yandex"; }
            if (HTTP_USER_AGENT.Contains("Iron ")) { return "Iron"; }
            if (HTTP_USER_AGENT.Contains("Maxthon/")) { return "Maxthon"; }
            return browser;
        }

        public static int getPageIndex(this HttpRequestBase rq, string key = "page")
        {
            if (rq[key] == null) return 1;
            int page = 1;
            try { page = int.Parse(rq[key].Trim()); } catch { page = 1; }
            if (page < 1) return 1;
            return page;
        }

        public static int getTotalPage(this int rowcount, int pagesize = 25)
        {
            int p = rowcount / pagesize;
            if (p * pagesize == rowcount) return p;
            return p + 1;
        }

        public static string BootstrapPage(this int rowcount, ref int pageindex, int pagesize, int numberShow = 3, string idform = "")
        {
            int totalpage = rowcount.getTotalPage(pagesize);
            if (pageindex > totalpage) { pageindex = totalpage; }
            if (pageindex < 1) { pageindex = 1; }
            return totalpage.BootstrapPage(pageindex, numberShow, idform);
        }

        public static string BootstrapPage(this int totalpage, int pageindex, int numberShow = 3, string idform = "")
        {
            if (numberShow < 3) { numberShow = 3; }
            string fc = "loadpage";
            /* <ul class="pagination"> <li><a href="#">1</a></li> <li class="active"><a href="#">2</a></li> </ul> */
            /* <span class="btn btn-primary btn-xs" onclick="loadPage('1');"><b> 1 </b></span> */
            var s = new List<string>();
            if (pageindex - numberShow > 1) s.Add($"<span class=\"btn btn-primary btn-xs\" onclick=\"{fc}('1','{idform}');\"><b> 1 </b></span> <span> [<b> .. </b>] </span");
            for (int i = pageindex - numberShow; i < pageindex; i++)
            {
                if (i < 1) continue;
                s.Add($"<span class=\"btn btn-primary btn-xs\" onclick=\"{fc}('{i}','{idform}');\"><b> {i} </b></span>");
            }
            s.Add($"<span> [<b> {pageindex} </b> ] </span>");
            for (int i = pageindex + 1; i <= pageindex + numberShow; i++)
            {
                if (i > totalpage) continue;
                s.Add($"<span class=\"btn btn-primary btn-xs\" onclick=\"{fc}('{i}','{idform}');\"><b> {i} </b></span>");
            }
            if (pageindex + numberShow < totalpage) s.Add($"<span> [<b> .. </b> ] </span> <span class=\"btn btn-primary btn-xs\" onclick=\"{fc}('{totalpage}','{idform}');\"><b> {totalpage} </b></span>");
            return string.Join(" ", s);
        }

        /// <summary>
        /// Alert: info, success, warning, danger
        /// </summary>
        /// <param name="s"></param>
        /// <param name="alert"></param>
        /// <returns></returns>

        public static string BootstrapAlter(this string s, string alert = "info") => $"<div class=\"alert alert-{alert}\">{s}</div>";

        public static void GetRequest(this Dictionary<string, string> ListRequest, HttpContext w = null)
        {
            if (w == null) w = HttpContext.Current;
            ListRequest.GetRequest(w.Request);
        }

        public static void GetRequest(this Dictionary<string, string> ListRequest, HttpRequest rq)
        {
            foreach (var v in rq.QueryString.AllKeys)
            {
                var getKey = v;
                if (v.Contains("$"))
                {
                    string[] keys = v.Split('$');
                    getKey = keys[keys.Length - 1];
                }
                ListRequest[getKey] = rq[v].Trim();
            }
            foreach (var key in rq.Form.AllKeys)
            {
                var getKey = key;
                if (key.Contains("$"))
                {
                    string[] keys = key.Split('$');
                    getKey = keys[keys.Length - 1];
                }
                ListRequest[getKey] = rq[key].Trim();
            }

            if (rq.Files.Count > 0)
            {
                for (int i = 0; i < rq.Files.Count; i++)
                {
                    var f = rq.Files[i];
                    if (f.ContentLength == 0) { continue; }
                    if (f.ContentLength < 1024) { ListRequest[f.FileName] = $"Size: {f.ContentLength}B"; }
                    if (f.ContentLength < 1048576) { ListRequest[f.FileName] = $"Size: {f.ContentLength}KB"; }
                    else { ListRequest[f.FileName] = $"Size: {(f.ContentLength / 1048576)}MB"; }
                }
            }
        }

        public static void GetRequest(this Dictionary<string, string> ListRequest, HttpRequestBase rq)
        {
            foreach (var v in rq.QueryString.AllKeys)
            {
                var getKey = v;
                if (v.Contains("$"))
                {
                    string[] keys = v.Split('$');
                    getKey = keys[keys.Length - 1];
                }
                ListRequest[getKey] = rq[v].Trim();
            }
            foreach (var key in rq.Form.AllKeys)
            {
                var getKey = key;
                if (key.Contains("$"))
                {
                    string[] keys = key.Split('$');
                    getKey = keys[keys.Length - 1];
                }
                ListRequest[getKey] = rq[key].Trim();
            }
            if (rq.Files.Count > 0)
            {
                for (int i = 0; i < rq.Files.Count; i++)
                {
                    var f = rq.Files[i];
                    if (f.ContentLength == 0) { continue; }
                    if (f.ContentLength < 1024) { ListRequest[f.FileName] = $"Size: {f.ContentLength}B"; }
                    if (f.ContentLength < 1048576) { ListRequest[f.FileName] = $"Size: {f.ContentLength}KB"; }
                    else { ListRequest[f.FileName] = $"Size: {(f.ContentLength / 1048576)}MB"; }
                }
            }
        }

        public static string toRequestString(this HttpRequestBase rq)
        {
            var ListRequest = new Dictionary<string, string>();
            ListRequest.GetRequest(rq);
            var s = new List<string>();
            foreach (var v in ListRequest) { s.Add($"{v.Key}: {v.Value}"); }
            return string.Join("; ", s);
        }

        public static void ShowMessage(string Message)
        {
            var w = HttpContext.Current;
            w.Response.Write($"<div>{Message}</div>");
        }

        public static string BuildElement(string Message, string element = "")
        {
            if (element == "") { element = "div"; }
            return $"<{element}>{Message}</{element}>";
        }

        /// <summary>
        /// Write Params by Method: POST & GET
        /// </summary>
        /// <param name="w"></param>
        public static void NotFoundMethod(HttpContext w = null)
        {
            if (w == null) { w = HttpContext.Current; }
            var rq = new Dictionary<string, string>();
            rq.GetRequest(w);
            var ls = new List<string>();
            ls.Add("Chưa khai báo phương thức thực thi. Thông tin tham số: ");
            foreach (var v in rq) ls.Add(v.Key + ": " + v.Value);
            w.Response.Write("<div>" + string.Join(" <br />", ls) + "</div>");
        }

        public static string ShowRequest(this HttpRequestBase request, string notStartWith = "", string element = "div")
        {
            var rq = new Dictionary<string, string>();
            rq.GetRequest(request);
            var ls = new List<string>();
            if (string.IsNullOrEmpty(notStartWith)) { foreach (var v in rq) { ls.Add(v.Key + ": " + v.Value); } }
            else
            {
                foreach (var v in rq)
                {
                    if (v.Key.StartsWith(notStartWith)) { continue; }
                    ls.Add(v.Key + ": " + v.Value);
                }
            }
            if (string.IsNullOrEmpty(element)) { return string.Join("; ", ls); }
            return $"<{element}>{string.Join(" <br />", ls)}</{element}>";
        }

        public static string ShowRequest(this HttpRequest request, string notStartWith = "", string element = "div")
        {
            if (string.IsNullOrEmpty(element)) { element = "div"; }
            var rq = new Dictionary<string, string>();
            rq.GetRequest(request);
            var ls = new List<string>();
            if (string.IsNullOrEmpty(notStartWith))
            {
                foreach (var v in rq) ls.Add(v.Key + ": " + v.Value);
            }
            else
            {
                foreach (var v in rq)
                {
                    if (v.Key.StartsWith(notStartWith)) continue;
                    ls.Add(v.Key + ": " + v.Value);
                }
            }
            if (string.IsNullOrEmpty(element)) return string.Join("; ", ls);
            return $"<{element}>{string.Join(" <br />", ls)}</{element}>";
        }

        public static string ShowRequest(this HttpContext w, string notStartWith = "", string element = "div") => w.Request.ShowRequest(notStartWith, element);

        public static string ShowRequest(string notStartWith = "", string element = "div") => HttpContext.Current.Request.ShowRequest(notStartWith, element);

        /// <summary>
        /// Return string HTML Class=jax_error
        /// </summary>
        /// <param name="sender">Exception; HttpContext; Element; Messager;</param>
        public static void ShowErrorHtml(params object[] sender)
        {
            var ls = new List<string>();
            var w = HttpContext.Current;
            Exception ex = null;
            foreach (object v in sender)
            {
                if (v is HttpContext) { w = v as HttpContext; continue; }
                if (v is Exception) { ex = v as Exception; continue; }
                ls.Add(v.ToString());
            }
            string description = "";
            if (ls.Count == 1) description = ls[0];
            string element = "div";
            if (ls.Count > 1)
            {
                if (ls[0] != "") element = ls[0];
                description = ls[1];
            }
            string msg = "Không xác định";
            if (ex != null)
            {
                var path = w.Server.MapPath("~");
                path = path.Substring(path.IndexOf(":") + 1);
                msg = ex.getLineHTML();
            }
            else if (ls.Count > 2) msg = ls[2];
            w.Response.Write($"<{element} class=\"jax_error\">Lỗi {element}: {msg}</{element}>");
        }

        public static string GetPathCodeProject(this Assembly a)
        {
            return Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(a.GetName().CodeBase))).Replace("file:\\", "");
        }

        /// <summary>
        /// Session: CaptchaImageText
        /// </summary>

        public static void CreatIamge(int MinLength = 3, int MaxLength = 9, int Width = 300, int Height = 75)
        {
            var w = HttpContext.Current;
            w.Response.Clear();
            w.Response.ContentType = "image/jpeg";
            w.Session["CaptchaImageText"] = Generate.RandomString(MinLength, MaxLength);
            var ci = new RandomImage(w.Session["CaptchaImageText"].ToString(), Width, Height);
            ci.Image.Save(w.Response.OutputStream, ImageFormat.Jpeg);
            ci.Dispose();
        }

        public static string RemoveTags(string source) => Regex.Replace(source, "<.*?>", string.Empty);
    }
}