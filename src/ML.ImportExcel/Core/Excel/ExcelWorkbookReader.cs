using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using NPOI.SS.UserModel;
using ML.ImportExcel.Core.Models;

namespace ML.ImportExcel.Core.Excel
{
    /// <summary>
    /// 基于 NPOI 的 Excel 读取封装：打开工作簿、解析工作表、读取列头与数据行。
    /// 同时支持 .xls(HSSF) 与 .xlsx(XSSF)。
    /// </summary>
    public static class ExcelWorkbookReader
    {
        private static int _encRegistered;

        /// <summary>注册 CodePages 编码（读取老式 .xls 中文内容需要）。</summary>
        public static void EnsureCodePagesRegistered()
        {
            if (Interlocked.Exchange(ref _encRegistered, 1) == 0)
            {
                try
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                }
                catch
                {
                    // 宿主可能已注册过或环境不支持
                }
            }
        }

        /// <summary>从流中打开工作簿（自动识别 .xls/.xlsx）。</summary>
        public static IWorkbook Open(Stream stream)
        {
            EnsureCodePagesRegistered();
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (stream.CanSeek) stream.Position = 0;
            return WorkbookFactory.Create(stream);
        }

        /// <summary>获取全部工作表名称。</summary>
        public static IList<string> GetSheetNames(IWorkbook wb)
        {
            var names = new List<string>();
            if (wb == null) return names;
            for (var i = 0; i < wb.NumberOfSheets; i++)
                names.Add(wb.GetSheetName(i));
            return names;
        }

        /// <summary>
        /// 解析出要读取的工作表：优先按名称，其次按索引，最后取第一个表。
        /// </summary>
        public static bool ResolveSheet(IWorkbook wb, string preferredSheetName, int preferredSheetIndex,
            out ISheet sheet, out int actualIndex, out string actualName)
        {
            sheet = null;
            actualIndex = 0;
            actualName = "";

            if (wb == null) return false;

            if (!string.IsNullOrWhiteSpace(preferredSheetName))
            {
                var i = wb.GetSheetIndex(preferredSheetName);
                if (i >= 0)
                {
                    sheet = wb.GetSheetAt(i);
                    actualIndex = i;
                    actualName = wb.GetSheetName(i);
                    return true;
                }
            }

            var idx = Math.Max(0, preferredSheetIndex);
            if (idx < wb.NumberOfSheets)
            {
                sheet = wb.GetSheetAt(idx);
                actualIndex = idx;
                actualName = wb.GetSheetName(idx);
                return true;
            }

            if (wb.NumberOfSheets > 0)
            {
                sheet = wb.GetSheetAt(0);
                actualIndex = 0;
                actualName = wb.GetSheetName(0);
                return true;
            }

            return false;
        }

        /// <summary>读取列头行（只返回存在非空文本的列）。</summary>
        public static List<HeaderColumn> ReadHeaders(ISheet sheet, int headerRowIndex)
        {
            var list = new List<HeaderColumn>();
            if (sheet == null) return list;

            var hr = Math.Max(0, headerRowIndex);
            var row = sheet.GetRow(hr);
            if (row == null) return list;

            foreach (var cell in row.Cells)
            {
                if (cell == null) continue;
                var text = ExcelValueConverters.GetText(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                list.Add(new HeaderColumn { Index = cell.ColumnIndex, Text = text });
            }

            list.Sort((a, b) => a.Index.CompareTo(b.Index));
            return list;
        }

        /// <summary>返回列头下前 maxRows 行非空数据的示例（每行与列头对齐）。</summary>
        public static List<object[]> ReadSampleRows(ISheet sheet, int headerRowIndex, int maxRows, int width)
        {
            var result = new List<object[]>();
            if (sheet == null || width <= 0) return result;

            var count = 0;
            for (var r = Math.Max(0, headerRowIndex) + 1; r <= sheet.LastRowNum && count < maxRows; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                var arr = new object[width];
                var any = false;
                foreach (var cell in row.Cells)
                {
                    if (cell == null) continue;
                    if (cell.ColumnIndex < 0 || cell.ColumnIndex >= width) continue;
                    var v = ExcelValueConverters.GetTypedValue(cell);
                    if (ExcelValueConverters.IsEmptyValue(v)) continue;
                    arr[cell.ColumnIndex] = v;
                    any = true;
                }
                if (!any) continue;

                result.Add(arr);
                count++;
            }
            return result;
        }

        /// <summary>读取全部数据行（跳过空行）。maxRows 为 null 表示不限制。</summary>
        public static List<SheetDataRow> ReadRows(ISheet sheet, int headerRowIndex, int? maxRows)
        {
            var rows = new List<SheetDataRow>();
            if (sheet == null) return rows;

            var first = Math.Max(sheet.FirstRowNum, Math.Max(0, headerRowIndex) + 1);
            var last = sheet.LastRowNum;
            var dataCount = 0;

            for (var r = first; r <= last; r++)
            {
                if (maxRows.HasValue && dataCount >= maxRows.Value) break;

                var row = sheet.GetRow(r);
                if (row == null) continue;

                var values = new Dictionary<int, object>();
                var any = false;
                foreach (var cell in row.Cells)
                {
                    if (cell == null) continue;
                    var v = ExcelValueConverters.GetTypedValue(cell);
                    if (ExcelValueConverters.IsEmptyValue(v)) continue;
                    values[cell.ColumnIndex] = v;
                    any = true;
                }
                if (!any) continue;

                rows.Add(new SheetDataRow { RowNumber = r + 1, Values = values });
                dataCount++;
            }
            return rows;
        }

        /// <summary>粗略统计表头之下的行数（含空行，供预览展示）。</summary>
        public static int CountDataRows(ISheet sheet, int headerRowIndex)
        {
            if (sheet == null) return 0;
            var hr = Math.Max(0, headerRowIndex);
            if (sheet.LastRowNum <= hr) return 0;
            return sheet.LastRowNum - hr;
        }
    }
}
