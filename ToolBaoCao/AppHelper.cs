using Microsoft.Ajax.Utilities;
using NPOI.HSSF.Record.Chart;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using ToolBaoCao.CaptchaImage;

namespace ToolBaoCao
{
    public static class AppHelper
    {
        public static long toTimestamp(this DateTime time) => ((DateTimeOffset)time).ToUnixTimeSeconds();

        /* Việt Nam múi giờ GMT +7 */

        public static DateTime fromTimestamp(this long timestamp, int GMT = 7) => (DateTimeOffset.FromUnixTimeSeconds(timestamp).DateTime).AddHours(GMT);

        public static DateTime fromTimestamp(this string timestamp, int GMT = 7)
        {
            if (Regex.IsMatch(timestamp, "^[0-9]+$")) { return (long.Parse(timestamp).fromTimestamp(7)); }
            return new DateTime(1970, 1, 1);
        }

        public static List<string> listKeyConfigCrypt = new List<string>() { "" };
        private static readonly string keyMD5 = typeof(AppHelper).Namespace;
        public static AppConfig appConfig = new AppConfig();
        public static readonly string pathApp = AppDomain.CurrentDomain.BaseDirectory;
        public static readonly string projectTitle = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;
        public static readonly string projectName = typeof(AppHelper).Namespace;
        public static dbSQLite dbSqliteMain = new dbSQLite();
        public static dbSQLite dbSqliteWork = new dbSQLite();
        public static CultureInfo cul = CultureInfo.GetCultureInfo("vi-VN");
        public static string formatNumberVN(this string NumberUS, int Decimal = 3)
        {
            if (Regex.IsMatch(NumberUS, "^[0-9]+$|^[-][0-9]+$"))
            {
                long v = long.Parse(NumberUS);
                if (v < 0) { return "-" + Math.Abs(v).ToString("#,##0", cul.NumberFormat); }
                return v.ToString("#,##0", cul.NumberFormat);
            }
            if (Regex.IsMatch(NumberUS, "^[0-9][.][0-9]+$^[-][0-9][.][0-9]+$"))
            {
                var l = new List<string>();
                if (Decimal > 0) { for (int i = 0; i < Decimal; i++) { l.Add("0"); } }
                var f = l.Count == 0 ? "#,##0.#" : "#,##0." + string.Join("", l);
                double v = double.Parse(NumberUS);
                if (v < 0) { return "-" + Math.Abs(v).ToString(f, cul.NumberFormat); }
                return v.ToString(f, cul.NumberFormat);
            }
            return NumberUS;
        }
        public static string formatNumberVN(this object NumberUS, int Decimal = 3) => NumberUS.ToString().formatNumberVN(Decimal);
        public static List<string> GetTableNameFromTsql(string tsql)
        { 
            var matches = Regex.Matches(tsql, @"\b(FROM|JOIN|UPDATE)\s+([a-zA-Z0-9_.\[\]]+)", RegexOptions.IgnoreCase);
            var tableNames = new List<string>();
            foreach (System.Text.RegularExpressions.Match match in matches) { tableNames.Add(match.Groups[2].Value); }
            return tableNames;
        }

        public static bool IsUpdateOrDelete(string sql) => Regex.IsMatch(sql, @"^\s*(UPDATE|DELETE)\s+", RegexOptions.IgnoreCase);

        public static string GetPathFileCacheQuery(string tsql, string dataName = "")
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
                if (!string.IsNullOrEmpty(fileCache)) { fileCache = $"{pathApp}/cache/d{dataName}_{tables[0]}_query_{fileCache}.xml"; }
            }
            return fileCache;
        }

        public static void DeleteFileCacheQuery(string tsql, string dataName)
        {
            if (!IsUpdateOrDelete(tsql)) return;
            var tables = GetTableNameFromTsql(tsql);
            if (tables.Count == 0) return;
            DeleteCache(tables[0] + "_", dataName);
        }

        public static void DeleteCache(string nameStartWith, string dataName)
        {
            if (string.IsNullOrEmpty(nameStartWith)) return;
            var files = Directory.GetFiles($"{pathApp}/cache/", $"d{dataName}_{nameStartWith}*");
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
                appConfig.Set("App.Title", "Công cụ hỗ trợ báo cáo bảo hiểm");
                appConfig.Set("App.PageSize", "50");
                appConfig.Set("App.PacketSize", "1000");
            }
            dbSqliteMain = new dbSQLite(Path.Combine(pathApp, "App_Data\\main.db"));
            dbSqliteMain.buildData();
            dbSqliteWork = new dbSQLite(Path.Combine(pathApp, "App_Data\\data.db"));
            dbSqliteWork.buildDataCongViec();
            /* Check Folder Exists */
            if(Directory.Exists(pathApp + "cache") == false) { Directory.CreateDirectory(pathApp + "cache"); }
            if (Directory.Exists(pathApp + "temp") == false) { Directory.CreateDirectory(pathApp + "temp"); }
            if (Directory.Exists(pathApp + "temp\\data") == false) { Directory.CreateDirectory(pathApp + "temp\\data"); }
            if (Directory.Exists(pathApp + "temp\\excel") == false) { Directory.CreateDirectory(pathApp + "temp\\excel"); }
            getDBUserOnline();           
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
        public static dbSQLite getDBUserOnline()
        {
            string pathData = pathApp + "App_Data\\useronline.db";
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
        public static bool CheckIsLogin()
        {
            var http = HttpContext.Current;
            if (http == null) return false;
            var db = getDBUserOnline();
            int maxSeccondsOnline = 15 * 60;
            try { db.Execute($"DELETE useronline WHERE ({DateTime.Now.toTimestamp()} - time2) > {maxSeccondsOnline}"); } catch { }
            var tmp = $"{http.Session["app.isLogin"]}";
            if (tmp == "1") {
                db.Execute($"UPDATE useronline SET time2={DateTime.Now.toTimestamp()} WHERE userid='{http.Session["iduser"]}' AND ip='{http.Session[keyMSG.SessionIPAddress]}'");
                return true; 
            }
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
                if (items.Rows.Count == 0) { return $"Tài khoản '{userName}' không tồn tại hoặc mật khẩu không đúng"; }
                {
                    items = dbSqliteMain.getDataTable("SELECT * FROM taikhoan LIMIT 1");
                    if (items.Rows.Count == 0)
                    {
                        var time = DateTime.Now.toTimestamp();
                        dbSqliteMain.Execute($"INSERT INTO admins ([iduser] ,[mat_khau] ,[ten_hien_thi] ,[gioi_tinh] ,[ngay_sinh] ,[email] ,[dien_thoai] ,[dia_chi] ,[hinh_dai_dien] ,[ghi_chu] ,[time_create] ,[time_last_login], nhom) VALUES ('admin', '{"admin123@".GetMd5Hash()}', 'Adminstrator', 'Nam', '{DateTime.Now:dd/MM/yyyy}', 'hotoancntt15a@gmail.com', '09140272795', 'Thành phố Lào Cai, Tỉnh Lào Cai', '', '', '{time}', '0', 0);");
                        items = dbSqliteMain.getDataTable(tsql, new KeyValuePair<string, string>("@iduser", userName));
                        if (items.Rows.Count == 0) { return $"Tài khoản '{userName}' không tồn tại hoặc mật khẩu không đúng"; }
                    }
                }
                if (http == null) { return keyMSG.ErrorHttpConnetNull; }
                http.Session.Clear();
                http.Request.Cookies.Clear();
                http.Session.Add("app.isLogin", "1");
                foreach (DataColumn c in items.Columns) { http.Session.Add(c.ColumnName, $"{items.Rows[0][c.ColumnName]}"); }
                /* IDUSER|PASS|DATETIME */
                if (remember)
                {
                    HttpCookie c1 = new HttpCookie("idobject", $"{userName}|{passWord}|{DateTime.Now}".MD5Encrypt());
                    c1.Expires = DateTime.Now.AddMonths(1);
                    http.Response.Cookies.Add(c1);
                }
                try { dbSqliteMain.Execute($"UPDATE taikhoan SET time_last_login='{DateTime.Now.toTimestamp()}' WHERE iduser = @iduser", new KeyValuePair<string, string>("@iduser", userName)); } catch { }
            }
            catch (Exception ex) { return $"Lỗi: {ex.Message} <br />Chi tiết: {ex.StackTrace}"; }
            var db = getDBUserOnline();
            db.Execute($"INSERT OR IGNORE INTO useronline (userid, time1, time2, ip) VALUES ('{http.Session["iduser"]}',{DateTime.Now.toTimestamp()},{DateTime.Now.toTimestamp()},'{http.Session[keyMSG.SessionIPAddress]}')");
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
            catch (Exception ex) { return "Lỗi: " + ex.Message; }
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

        public static string getValue(this Dictionary<string, string> data, string key, string defaultValue = "")
        {
            if (data.ContainsKey(key)) { return data[key]; }
            return defaultValue;
        }

        public static object getValue(this Dictionary<string, object> data, string key, object defaultValue = null)
        {
            if (data != null && data.ContainsKey(key)) { return data[key]; }
            return defaultValue;
        }

        public static bool isDateVN(this string input)
        {
            return Regex.IsMatch(input, "^[0-3][0-9]/[0-1][0-9]/[1-9][0-9]{3}$|^[0-3][0-9]/[0-1][0-9]/[1-9][0-9]{3} [0-2][0-9]:[0-5][0-9]$|^[0-3][0-9]/[0-1][0-9]/[1-9][0-9]{3} [0-2][0-9]:[0-5][0-9]:[0-5][0-9]$");
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

        public static string getErrorSave(this Exception ex)
        {
            try
            {
                using (var sw = new StreamWriter(HttpContext.Current.Server.MapPath("~/error.log"), true, Encoding.Unicode))
                {
                    try { sw.WriteLine($"{DateTime.Now:dd/MM/yyyy HH:mm:ss} {ex.Message} {ex.StackTrace}"); sw.Flush(); } catch { }
                }
                return ex.getLineHTML();
            }
            catch { return $"Lỗi: {ex.Message}<br />Chi tiết:{ex.StackTrace}"; }
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

        public static string getPathCodeProject(this Assembly a)
        {
            var s = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(a.GetName().CodeBase)));
            s = s.Replace("file:\\", "");
            return s;
        }

        public static string getLineHTML(this Exception ex, string description = "", Assembly a = null)
        {
            var msg = ex.StackTrace.Split('\n').Where(p => Regex.IsMatch(p, @":line \d+")).ToList();
            var s = Assembly.GetExecutingAssembly().getPathCodeProject();
            var s2 = string.Join(" <br /> ", msg).Replace(s, "");
            if (a != null) s2 = s2.Replace(a.getPathCodeProject(), "");
            if (string.IsNullOrEmpty(description)) return $"{ex.Message} <br /> {s2}";
            return $"Lỗi {description}: {ex.Message} <br /> {s2}";
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

        public static string sqliteGetValueField(this string value) => value.Replace("'", "''");
    }
}