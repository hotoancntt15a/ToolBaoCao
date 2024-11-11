using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;

namespace ToolBaoCao
{
    public class AppConfig
    {
        private string timeLastWrite = "";
        private string filePath;
        public Config Config { get; private set; }

        public AppConfig(string configFilePath = "")
        {
            if (configFilePath == "") { filePath = Path.Combine(AppHelper.pathApp, "config.json"); }
            else { filePath = configFilePath; }
            Load();
        }

        /* Đọc cấu hình từ file */

        private void Load()
        {
            if (File.Exists(filePath))
            {
                timeLastWrite = File.GetLastWriteTime(filePath).ToString();
                string json = File.ReadAllText(filePath, System.Text.Encoding.Unicode);
                Config = JsonConvert.DeserializeObject<Config>(json) ?? new Config { Settings = new List<Setting>() };
            }
            else { Config = new Config { Settings = new List<Setting>() }; }
        }

        /* Lưu cấu hình vào file */

        private void Save()
        {
            string json = JsonConvert.SerializeObject(Config, Formatting.Indented);
            File.WriteAllText(filePath, json, System.Text.Encoding.Unicode);
            timeLastWrite = File.GetLastWriteTime(filePath).ToString();
        }

        /* Sửa và Thêm nếu chưa có */

        public void Set(string key, string newValue)
        {
            var setting = Config.Settings.Find(s => s.Key == key);
            if (setting != null) { setting.Value = newValue; }
            else { Config.Settings.Add(new Setting { Key = key, Value = newValue }); }
            Save();
        }

        /* Xóa phần tử */

        public void Remove(string key)
        {
            var setting = Config.Settings.Find(s => s.Key == key);
            if (setting != null)
            {
                Config.Settings.Remove(setting);
                Save();
            }
        }

        /* Lấy giá trị của phần tử */

        public string Get(string key, string valueDefault = "")
        {
            try { if (timeLastWrite != File.GetLastWriteTime(filePath).ToString()) { Load(); } } catch { }
            var setting = Config.Settings.Find(s => s.Key == key);
            if (setting != null) { return setting.Value; }
            if (key == "app.title" || key == "App.Title") { return "HIA-Tools"; }
            Config.Settings.Add(new Setting { Key = key, Value = valueDefault });
            Save();
            return valueDefault;
        }
    }

    public class Setting
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public class Config
    {
        public List<Setting> Settings { get; set; }
    }
}