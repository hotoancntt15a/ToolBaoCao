using NPOI.HSSF.Record;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

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
        private readonly ConcurrentDictionary<string, ItemTask> _threads = new ConcurrentDictionary<string, ItemTask>();
        private readonly object _lock = new object();
        private Timer _timer;
        private dbSQLite dbTask = new dbSQLite(Path.Combine(AppHelper.pathAppData, "task.db"));

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
            if (_threads.TryAdd(item.ID, item))
            {
                var tsql = $"INSERT OR IGNORE INTO task(id, nametask, actionname, args, timestart) VALUES ('{item.ID}', '{item.NameTask.sqliteGetValueField()}', '{item.ActionName.sqliteGetValueField()}', '{item.Args.sqliteGetValueField()}', '{item.TimeStart.toTimestamp()}')";
                try
                {
                    dbTask.Execute(tsql);
                }
                catch (Exception ex)
                {
                    AppHelper.saveError(tsql + Environment.NewLine + ex.Message);
                    throw new Exception(ex.getLineHTML());
                }
                if (callRun) { Call(); }
            }
        }

        public void Delete(string ID)
        {
            if (_threads.TryGetValue(ID, out var item) && item.Running == 0)
            {
                _threads.TryRemove(ID, out _);
                dbTask.Execute($"DELETE FROM task WHERE id='{ID}';");
            }
        }

        public void Call()
        {
            lock (_lock)
            {
                // Find the first thread item that is not running
                var item = _threads.Values.FirstOrDefault(t => t.Running == 0);
                if (item != null)
                {
                    item.Running = 1;
                    Task.Run(() =>
                    {
                        try
                        {
                            switch (item.ActionName)
                            {
                                case "Controller.XML":
                                    SQLiteCopy.threadCopyXML(item.Args);
                                    break;

                                default: break;
                            }
                            Thread.Sleep(2000);
                        }
                        finally
                        {
                            Delete(item.ID);
                            Call();
                        }
                    });
                }
            }
        }
    }
}