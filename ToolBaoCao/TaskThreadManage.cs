using System;
using System.Collections.Concurrent;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;

namespace ToolBaoCao
{
    public class ItemTask
    {
        public ItemTask(string id, string name, string actionName = "", string args = "", long timeStart = 0)
        {
            ID = id;
            NameTask = name;
            ActionName = actionName;
            Args = args;
            if (timeStart == 0) { TimeStart = DateTime.Now; }
            else { TimeStart = timeStart.toDateTime(); }
        }

        public string ID { get; set; } = "";
        public string NameTask { get; set; } = "";
        public string ActionName { get; set; } = "";
        public string Args { get; set; } = "";
        public long Running { get; set; } = 0;
        public DateTime TimeStart { get; set; } = DateTime.Now;
    }

    public class TaskManage
    {
        private ConcurrentDictionary<string, ItemTask> _threads = new ConcurrentDictionary<string, ItemTask>();
        private Timer _timer;
        private dbSQLite dbTask = new dbSQLite(Path.Combine(AppHelper.pathAppData, "task.db"));
        public string IDRunning = "";

        public TaskManage()
        {
            Load();
            _timer = new Timer(_ => Call(), null, TimeSpan.Zero, TimeSpan.FromMinutes(30));
        }

        public void Load()
        {
            dbTask.Execute("CREATE TABLE IF NOT EXISTS task(id text not null primary key, nametask text not null default '', actionname text not null default '', args text not null default '', running integer not null default 0, timestart integer not null);");
            var data = dbTask.getDataTable("SELECT * FROM task ORDER BY timestart");
            foreach (DataRow row in data.Rows)
            {
                var item = new ItemTask(row["id"].ToString(), row["nametask"].ToString(), $"{row["actionname"]}", $"{row["args"]}", long.Parse($"{row["timestart"]}"));
                Add(item, false);
            }
        }

        public void Add(ItemTask item, bool callRun = true)
        {
            item.Running = 0;
            if (_threads.TryAdd(item.ID, item))
            {
                var tsql = $"INSERT OR IGNORE INTO task(id, nametask, actionname, args, timestart) VALUES ('{item.ID}', '{item.NameTask.sqliteGetValueField()}', '{item.ActionName.sqliteGetValueField()}', '{item.Args.sqliteGetValueField()}', '{item.TimeStart.toTimestamp()}')";
                try { dbTask.Execute(tsql); }
                catch (Exception ex)
                {
                    AppHelper.saveError($"Task({item.ID} - {item.ActionName} - {item.Args}): {tsql}{Environment.NewLine}{ex.Message}");
                    throw new Exception(ex.getLineHTML());
                }
                if (callRun) { Call(); }
            }
        }

        public void Delete(string ID)
        {
            if (_threads.TryGetValue(ID, out var item))
            {
                _threads.TryRemove(ID, out _);
                dbTask.Execute($"DELETE FROM task WHERE id='{item.ID}';");
            }
            Call();
        }

        public void Call()
        {
            if (IDRunning != "")
            {
                var obj = _threads.Values.FirstOrDefault(p => p.ID == IDRunning);
                if (obj != null) { AppHelper.saveError($"Running Task({obj.ID} - {obj.ActionName} - {obj.Args})"); return; }
                AppHelper.saveError($"Not Find ID Task '{IDRunning}'");
                IDRunning = "";
            }
            var item = _threads.Values.FirstOrDefault();
            if (item != null)
            {
                AppHelper.saveError($"Start Task({item.ID} - {item.ActionName} - {item.Args})");
                IDRunning = item.ID;
                try
                {
                    switch (item.ActionName.ToLower())
                    {
                        case "controller.xml":
                            AppHelper.saveError($"Wait run XMLThread({item.ID} - {item.ActionName} - {item.Args})");
                            Thread t = new Thread(new ThreadStart(() =>
                            {
                                AppHelper.saveError($"Run XMLThread({item.ID} - {item.ActionName} - {item.Args})");
                                /* try { XMLThread(item.Args); }
                                catch (Exception exT) { AppHelper.saveError($"Error XMLThread({item.ID} - {item.ActionName} - {item.Args}): {exT.Message}"); }
                                */
                            }));
                            t.Start();
                            break;

                        default: AppHelper.saveError($"Not Found XMLThread({item.ID} - {item.ActionName} - {item.Args})"); break;
                    }
                }
                catch (Exception ex) { AppHelper.saveError($"Task({item.ID} - {item.ActionName} - {item.Args}): {ex.Message}"); }
            }
        }

        public void setFinshThreadInAppStart()
        {
            Thread t = new Thread(new ThreadStart(() =>
            {
                var d = new DirectoryInfo(Path.Combine(AppHelper.pathAppData, "xml"));
                if ((d.Exists == false)) { d.Create(); return; }
                foreach (var f in d.GetFiles("*.db"))
                {
                    var db = new dbSQLite(f.FullName);
                    try
                    {
                        var tables = db.getAllTables();
                        if (tables.Contains("xml"))
                        {
                            db.Execute($"UPDATE xml SET title='Lỗi do do hệ thống bị ngắt đột ngột', time2={DateTime.Now.toTimestamp()} WHERE time2=0;");
                        }
                    }
                    catch { }
                    db.Close();
                }
            }));
            t.Start();
        }
    }
}