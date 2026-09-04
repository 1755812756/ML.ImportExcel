using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ML.ImportExcel.Core.Models
{
    /// <summary>字段数据类型常量。</summary>
    public static class ExcelFieldDataTypes
    {
        public const string String = "string";
        public const string Int = "int";
        public const string Long = "long";
        public const string Double = "double";
        public const string Decimal = "decimal";
        public const string DateTime = "datetime";
        public const string Bool = "bool";

        /// <summary>全部支持的类型（用于网页下拉选项）。</summary>
        public static readonly string[] All = { String, Int, Long, Double, Decimal, DateTime, Bool };

        /// <summary>是否为受支持类型。</summary>
        public static bool IsKnown(string t)
        {
            if (string.IsNullOrEmpty(t)) return false;
            return All.Any(x => string.Equals(x, t, StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>“Excel 列头 → 实体字段”绑定配置。</summary>
    public sealed class ImportFieldMapping
    {
        public long Id { get; set; }

        /// <summary>所属配置 Id。</summary>
        public long ConfigId { get; set; }

        /// <summary>Excel 中的列序号（从 0 开始）。</summary>
        public int ColumnIndex { get; set; }

        /// <summary>Excel 列头文本。</summary>
        public string Header { get; set; } = "";

        /// <summary>绑定的实体字段名称（即实体类的属性名）。</summary>
        public string Field { get; set; } = "";

        /// <summary>是否必填。</summary>
        public bool Required { get; set; }

        /// <summary>期望数据类型（见 ExcelFieldDataTypes）。为空表示自动。</summary>
        public string DataType { get; set; } = "";
    }

    /// <summary>
    /// 配置文件夹（树形结构，用于按“功能/业务”分组管理导入配置）。
    /// 每个文件夹可有父文件夹（ParentId 为空表示根目录）。
    /// </summary>
    public sealed class ImportFolder
    {
        public long Id { get; set; }

        /// <summary>文件夹名称（同级唯一）。</summary>
        public string Name { get; set; } = "";

        /// <summary>父文件夹 Id；为空/null 表示位于根目录。</summary>
        public long? ParentId { get; set; }
    }

    /// <summary>一份 Excel 导入配置（对应一个实体 / 一种导入模板）。</summary>
    public sealed class ImportConfig
    {
        public long Id { get; set; }

        /// <summary>所属文件夹 Id；为空/null 表示根目录。</summary>
        public long? FolderId { get; set; }

        /// <summary>配置名称（唯一，用于标识一个实体模板）。</summary>
        public string Name { get; set; } = "";

        /// <summary>实体名称（业务实体/类名，用于结果 JSON 中标识）。</summary>
        public string EntityName { get; set; } = "";

        /// <summary>配置描述。</summary>
        public string Description { get; set; } = "";

        /// <summary>工作表索引（从 0 开始），默认 0。</summary>
        public int SheetIndex { get; set; } = 0;

        /// <summary>工作表名称（可选；若存在则优先按名称取表）。</summary>
        public string SheetName { get; set; } = "";

        /// <summary>列头所在行的索引（从 0 开始），默认 0。</summary>
        public int HeaderRowIndex { get; set; } = 0;

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public DateTime UpdatedAt { get; set; } = DateTime.Now;

        /// <summary>列头 → 字段 的绑定列表。</summary>
        public List<ImportFieldMapping> Fields { get; set; } = new List<ImportFieldMapping>();

        /// <summary>当前配置中已绑定有效字段的映射数量（供前端/匹配使用）。</summary>
        [JsonIgnore]
        public int BoundFieldCount => Fields == null ? 0 : Fields.Count(f =>
            !string.IsNullOrWhiteSpace(f.Field) && !string.IsNullOrWhiteSpace(f.Header));
    }

    /// <summary>Excel 表头列信息。</summary>
    public sealed class HeaderColumn
    {
        public int Index { get; set; }

        public string Text { get; set; } = "";
    }

    /// <summary>网页加载 Excel 后返回的预览信息（列头 + 示例数据）。</summary>
    public sealed class TemplatePreview
    {
        /// <summary>加载过程提示（如无可用工作表等）。</summary>
        public string Message { get; set; } = "";

        public string[] SheetNames { get; set; } = Array.Empty<string>();

        public int SheetIndex { get; set; }

        public string SheetName { get; set; } = "";

        /// <summary>读取到的列头行号（0 基）。</summary>
        public int HeaderRowIndex { get; set; }

        public int ColumnCount { get; set; }

        /// <summary>列头以下的数据行数（粗略）。</summary>
        public int DataRowCount { get; set; }

        public List<HeaderColumn> Columns { get; set; } = new List<HeaderColumn>();

        /// <summary>示例数据行（每行与 Columns 对齐）。</summary>
        public List<object[]> SampleRows { get; set; } = new List<object[]>();
    }

    /// <summary>Excel 解析结果中的单行错误。</summary>
    public sealed class ParseRowError
    {
        /// <summary>Excel 中的实际行号（从 1 开始计，含表头）。</summary>
        public int RowNumber { get; set; }

        /// <summary>出错所在工作表名称。</summary>
        public string Sheet { get; set; } = "";

        /// <summary>Excel 单元格坐标（如 E2，列字母 + 行号）。</summary>
        public string Cell { get; set; } = "";

        /// <summary>出错列的序号（从 0 开始）。</summary>
        public int? ColumnIndex { get; set; }

        /// <summary>该列对应的 Excel 原始列表头（如“月薪”）。</summary>
        public string ColumnHeader { get; set; } = "";

        /// <summary>绑定的实体字段名。</summary>
        public string Field { get; set; } = "";

        /// <summary>Excel 单元格原始值（显示文本）。</summary>
        public string RawValue { get; set; } = "";

        public string Message { get; set; } = "";
    }

    /// <summary>Excel 解析结果。</summary>
    public sealed class ExcelParseResult
    {
        public bool Success { get; set; }

        public string Message { get; set; } = "";

        public long? ConfigId { get; set; }

        /// <summary>命中的配置名称。</summary>
        public string ConfigName { get; set; } = "";

        /// <summary>命中的实体名称。</summary>
        public string EntityName { get; set; } = "";

        /// <summary>自动匹配的列头命中率。</summary>
        public double MatchScore { get; set; }

        public int TotalRows { get; set; }

        public int SuccessRows { get; set; }

        public int ErrorRows { get; set; }

        /// <summary>校验错误明细。</summary>
        public List<ParseRowError> Errors { get; set; } = new List<ParseRowError>();

        /// <summary>解析出的数据行（每个字段名 → 值）。</summary>
        public List<Dictionary<string, object>> Rows { get; set; } = new List<Dictionary<string, object>>();

        /// <summary>将结果序列化为 JSON 字符串（成功时包含实体名、命中率与数据行）。</summary>
        public string ToJson()
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };
            return JsonSerializer.Serialize(this, options);
        }

        /// <summary>仅提取数据行部分为 JSON 数组字符串（字段名 → 值）。</summary>
        public string RowsToJson()
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };
            return JsonSerializer.Serialize(Rows, options);
        }
    }
}
