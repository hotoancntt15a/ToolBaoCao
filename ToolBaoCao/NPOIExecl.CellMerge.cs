using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace zModules.NPOIExcel
{
    public class CellMerge
    {
        public int RowIndex { get; set; } = 0;
        public int ColumnIndex { get; set; } = 0;
        public int MergeRowCount { get; set; } = 0;
        public int MergeColumnCount { get; set; } = 0;
        public string Value { get; set; }

        public CellMerge()
        {
        }

        public CellMerge(int RowIndex, int ColumnIndex, string Value, int MergeRowCount = 0, int MergeColumnCount = 0)
        {
            this.RowIndex = RowIndex;
            this.ColumnIndex = ColumnIndex;
            this.Value = Value;
            this.MergeRowCount = MergeRowCount;
            this.MergeColumnCount = MergeColumnCount;
        }
    }
}
