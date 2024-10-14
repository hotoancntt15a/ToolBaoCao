using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ToolBaoCao
{
    public class ItemThread
    {
        public ItemThread(string id, string name, string actionName = "", string args = "", bool run = false)
        {
            ID = id;
            Name = name;
            ActionName = actionName;
            Args = args;
        }

        public string ID { get; set; } = "";
        public string Name { get; set; } = "";
        public string ActionName { get; set; } = "";
        public string Args { get; set; } = "";
        public bool Run { get; set; } = false;
    }

    public static class ThreadManage
    {
        private static List<ItemThread> ListThread = new List<ItemThread>();
        public static void Add(ItemThread item)
        {
            if (ListThread.Any(p => p.ID == item.ID) == false) { return; }

        }
        public static void Call()
        {

        }
    }
}