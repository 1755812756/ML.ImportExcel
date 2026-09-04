using System.Collections.Generic;

namespace ML.ImportExcel.Core.Excel
{
    /// <summary>Excel 数据表中的一行数据。</summary>
    public sealed class SheetDataRow
    {
        /// <summary>Excel 实际行号（从 1 开始计，含表头行）。</summary>
        public int RowNumber { get; set; }

        /// <summary>该行各列的值，键为列序号（从 0 开始）。</summary>
        public Dictionary<int, object> Values { get; set; } = new Dictionary<int, object>();

        /// <summary>是否整行为空。</summary>
        public bool IsEmpty => Values == null || Values.Count == 0;
    }
}
