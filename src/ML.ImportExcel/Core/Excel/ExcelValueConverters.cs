using System;
using System.Globalization;
using NPOI.SS.UserModel;

namespace ML.ImportExcel.Core.Excel
{
    /// <summary>
    /// Excel 单元格值转换工具：负责把 NPOI 单元格转成对象、文本，并按目标数据类型做校验转换。
    /// </summary>
    public static class ExcelValueConverters
    {
        private static readonly string[] DateFormats =
        {
            "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-dd",
            "yyyy/MM/dd HH:mm:ss", "yyyy/MM/dd HH:mm", "yyyy/MM/dd",
            "yyyy-M-d H:mm:ss", "yyyy-M-d H:mm", "yyyy-M-d",
            "yyyy年M月d日 HH:mm:ss", "yyyy年M月d日 HH:mm", "yyyy年M月d日",
            "yyyy-MM-ddTHH:mm:ss"
        };

        /// <summary>整型化处理：整数值存为 long，否则保留 double。</summary>
        public static object NormalizeNumber(double d)
        {
            if (double.IsNaN(d) || double.IsInfinity(d)) return null;
            if (d == Math.Floor(d) && d >= long.MinValue && d <= long.MaxValue)
                return (long)d;
            return d;
        }

        /// <summary>判断值是否为空（null / 空白字符串）。</summary>
        public static bool IsEmptyValue(object v)
        {
            if (v == null) return true;
            if (v is string s) return string.IsNullOrWhiteSpace(s);
            return false;
        }

        /// <summary>把对象转成展示文本。</summary>
        public static string ToText(object v)
        {
            if (v == null) return "";
            if (v is DateTime dt) return FormatDate(dt);
            if (v is string s) return s.Trim();
            if (v is bool b) return b ? "true" : "false";
            if (v is long l) return l.ToString(CultureInfo.InvariantCulture);
            if (v is int i) return i.ToString(CultureInfo.InvariantCulture);
            if (v is double d) return d.ToString(CultureInfo.InvariantCulture);
            if (v is decimal m) return m.ToString(CultureInfo.InvariantCulture);
            if (v is float f) return f.ToString(CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
        }

        public static string FormatDate(DateTime dt)
        {
            return dt.TimeOfDay == TimeSpan.Zero
                ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 按声明的目标数据类型转换单元格值。
        /// dataType 为空表示“自动”，原样返回原始类型。
        /// 失败时 error 返回原因，result 为 null。
        /// </summary>
        public static object ConvertToFieldType(object raw, string dataType, out string error)
        {
            error = null;
            var type = string.IsNullOrWhiteSpace(dataType) ? "" : dataType.Trim().ToLowerInvariant();
            if (type.Length == 0 || raw == null) return raw;
            if (raw is string s && string.IsNullOrWhiteSpace(s)) return null;

            switch (type)
            {
                case "string":
                    return ToText(raw);

                case "int":
                    if (TryToInt64(raw, out var l) && l >= int.MinValue && l <= int.MaxValue)
                        return (int)l;
                    error = "type.invalidInt";
                    return null;

                case "long":
                    if (TryToInt64(raw, out var l2)) return l2;
                    error = "type.invalidLong";
                    return null;

                case "double":
                    if (TryToDouble(raw, out var dbl)) return dbl;
                    error = "type.invalidDouble";
                    return null;

                case "decimal":
                    if (raw is decimal dec) return dec;
                    if (raw is long lngD) return (decimal)lngD;
                    if (raw is double dd)
                    {
                        if (double.IsNaN(dd) || double.IsInfinity(dd)) { error = "type.invalidNumber"; return null; }
                        return (decimal)dd;
                    }
                    if (raw is string sd &&
                        decimal.TryParse(sd.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out dec))
                        return dec;
                    error = "type.invalidDecimal";
                    return null;

                case "datetime":
                    return ConvertToDateTime(raw, out error);

                case "bool":
                    return ConvertToBool(raw, out error);

                default:
                    return raw;
            }
        }

        private static object ConvertToDateTime(object raw, out string error)
        {
            error = null;
            if (raw is DateTime dt) return dt;
            if (raw is long lng) return ExcelSerialToDate(lng);
            if (raw is double d) return ExcelSerialToDate(d);
            if (raw is int i) return ExcelSerialToDate(i);
            if (raw is string str)
            {
                var t = str.Trim();
                if (DateTime.TryParseExact(t, DateFormats, CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var exact))
                    return exact;
                if (DateTime.TryParse(t, CultureInfo.CurrentCulture, DateTimeStyles.None, out var cur))
                    return cur;
                if (DateTime.TryParse(t, CultureInfo.InvariantCulture, DateTimeStyles.None, out var inv))
                    return inv;
            }
            error = "type.invalidDateTime";
            return null;
        }

        private static object ConvertToBool(object raw, out string error)
        {
            error = null;
            if (raw is bool b) return b;
            if (raw is long lng) return lng != 0;
            if (raw is int i) return i != 0;
            if (raw is double d) return d != 0;
            if (raw is string s)
            {
                switch (s.Trim().ToLowerInvariant())
                {
                    case "true": case "1": case "是": case "yes": case "y": case "√": case "对":
                        return true;
                    case "false": case "0": case "否": case "no": case "n": case "×": case "错":
                        return false;
                }
            }
            error = "type.invalidBool";
            return null;
        }

        private static bool TryToInt64(object raw, out long result)
        {
            result = 0;
            if (raw is long l) { result = l; return true; }
            if (raw is int i) { result = i; return true; }
            if (raw is double d)
            {
                if (double.IsNaN(d) || double.IsInfinity(d)) return false;
                if (d == Math.Floor(d) && d >= long.MinValue && d <= long.MaxValue)
                {
                    result = (long)d;
                    return true;
                }
                return false;
            }
            if (raw is string s)
                return long.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
            return false;
        }

        private static bool TryToDouble(object raw, out double result)
        {
            result = 0;
            if (raw is double d) { result = d; return true; }
            if (raw is long l) { result = l; return true; }
            if (raw is int i) { result = i; return true; }
            if (raw is decimal m) { result = (double)m; return true; }
            if (raw is float f) { result = f; return true; }
            if (raw is string s)
                return double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out result);
            return false;
        }

        /// <summary>把 Excel 序列化日期数字转成 DateTime（1900 日期系统）。</summary>
        public static DateTime ExcelSerialToDate(double serial)
        {
            // Excel 1900 日期系统：1899-12-30 为 0 基准（修正 1900 闰年 bug）
            try
            {
                return new DateTime(1899, 12, 30).AddDays(serial);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>获取单元格的“原始类型”对象值（数值/日期/布尔/文本）。</summary>
        public static object GetTypedValue(ICell cell)
        {
            if (cell == null) return null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.String:
                {
                    var s = cell.StringCellValue;
                    return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
                }
                case CellType.Numeric:
                {
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try { return cell.DateCellValue; }
                        catch { return null; }
                    }
                    return NormalizeNumber(cell.NumericCellValue);
                }
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Formula:
                    return EvaluateFormula(cell);
                case CellType.Error:
                    return null;
                default:
                    return null;
            }
        }

        /// <summary>获取单元格显示文本（列头、预览用）。</summary>
        public static string GetText(ICell cell)
        {
            if (cell == null) return "";
            var v = GetTypedValue(cell);
            return v == null ? "" : ToText(v);
        }

        private static object EvaluateFormula(ICell cell)
        {
            try
            {
                var evaluator = cell.Sheet?.Workbook?.GetCreationHelper()?.CreateFormulaEvaluator();
                if (evaluator == null) return null;
                var cv = evaluator.Evaluate(cell);
                if (cv == null) return null;
                switch (cv.CellType)
                {
                    case CellType.Blank:
                        return null;
                    case CellType.String:
                    {
                        var s = cv.StringValue;
                        return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
                    }
                    case CellType.Numeric:
                        return NormalizeNumber(cv.NumberValue);
                    case CellType.Boolean:
                        return cv.BooleanValue;
                    default:
                        return null;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
