using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ML.ImportExcel.Core.Excel;
using ML.ImportExcel.Core.Localization;
using ML.ImportExcel.Core.Matching;
using ML.ImportExcel.Core.Models;
using ML.ImportExcel.Core.Options;
using ML.ImportExcel.Core.Storage;
using NPOI.SS.UserModel;

namespace ML.ImportExcel.Core
{
    /// <summary>
    /// Excel 导入解析服务门面：
    ///  - 导入配置管理（CRUD，SQLite 持久化）；
    ///  - 加载模板 Excel 列头（供网页绑定实体字段）；
    ///  - 传入 Excel 自动匹配配置/实体并解析数据，输出 JSON。
    /// </summary>
    public interface IExcelImportService
    {
        // ---------- 配置管理 ----------
        IReadOnlyList<ImportConfig> ListConfigs();
        ImportConfig GetConfig(long id);
        ImportConfig GetConfigByNameOrEntity(string nameOrEntity);
        ImportConfig SaveConfig(ImportConfig config);
        void DeleteConfig(long id);

        // ---------- 文件夹（树形分组）管理 ----------
        IReadOnlyList<ImportFolder> ListFolders();
        ImportFolder GetFolder(long id);
        /// <summary>新增或更新文件夹（Id&lt;=0 视为新增；ParentId 为空表示根目录）。</summary>
        ImportFolder SaveFolder(ImportFolder folder);
        /// <summary>删除文件夹及其子文件夹，其下配置移动到根目录。</summary>
        void DeleteFolder(long id);

        // ---------- 模板列头加载（网页“导入 Excel→加载列头→绑定字段”）----------
        /// <summary>加载模板列头（使用默认语系）。</summary>
        TemplatePreview LoadTemplate(Stream excel, int? sheetIndex = null, int? headerRowIndex = null, int? maxSampleRows = null);

        /// <summary>加载模板列头（指定语系；未传时默认，找不到回退英文）。</summary>
        TemplatePreview LoadTemplate(Stream excel, int? sheetIndex, int? headerRowIndex, string language, int? maxSampleRows = null);

        // ---------- 解析 ----------
        /// <summary>传入 Excel，自动匹配配置/实体并解析，返回结果对象（使用默认语系）。</summary>
        ExcelParseResult Parse(Stream excel);

        /// <summary>传入 Excel 解析；configName 可空表示自动匹配，language 指定返回消息语系。</summary>
        ExcelParseResult Parse(Stream excel, string language, string configName="");

        /// <summary>仅指定语系解析（自动匹配实体）。</summary>
        ExcelParseResult ParseByLanguage(Stream excel, string language);

        /// <summary>传入 Excel，自动匹配配置并返回 JSON 字符串（使用默认语系）。</summary>
        string ParseToJson(Stream excel);

        /// <summary>传入 Excel 解析并返回 JSON；configName 可空表示自动匹配，language 指定语系。</summary>
        string ParseToJson(Stream excel,  string language, string configName = "");

        /// <summary>仅指定语系解析并返回 JSON（自动匹配实体）。</summary>
        string ParseToJsonByLanguage(Stream excel, string language);
    }

    public sealed class ExcelImportService : IExcelImportService
    {
        private readonly SqliteImportStore _store;
        private readonly ExcelImportOptions _options;
        private readonly MessageCatalog _catalog;
        private const int MaxDetailErrors = 1000;

        /// <summary>底层 SQLite 存储（可访问 DatabasePath 等）。</summary>
        public SqliteImportStore Store => _store;

        /// <summary>当前使用的消息目录（多语言包）。</summary>
        public MessageCatalog Catalog => _catalog;

        public ExcelImportService(SqliteImportStore store, ExcelImportOptions options)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _options = options ?? new ExcelImportOptions();
            _catalog = new MessageCatalog(_options.LanguageDirectory);
        }

        public ExcelImportService(string databasePath, ExcelImportOptions options = null)
            : this(new SqliteImportStore(databasePath), options ?? new ExcelImportOptions())
        {
        }

        // ================= 配置管理 =================

        public IReadOnlyList<ImportConfig> ListConfigs() => _store.ListConfigs();

        public IReadOnlyList<ImportFolder> ListFolders() => _store.ListFolders();

        public ImportFolder GetFolder(long id) => _store.GetFolder(id);

        public ImportFolder SaveFolder(ImportFolder folder)
        {
            if (folder == null) throw new ArgumentNullException(nameof(folder));
            return _store.SaveFolder(folder);
        }

        public void DeleteFolder(long id) => _store.DeleteFolderTree(id);

        public ImportConfig GetConfig(long id) => _store.GetConfig(id);

        public ImportConfig GetConfigByNameOrEntity(string nameOrEntity) => _store.GetByNameOrEntity(nameOrEntity);

        public ImportConfig SaveConfig(ImportConfig config)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));
            config.Fields = config.Fields ?? new List<ImportFieldMapping>();
            // 保证映射的列序号合法
            foreach (var f in config.Fields)
            {
                if (f.ColumnIndex < 0) f.ColumnIndex = 0;
                f.ConfigId = config.Id;
            }
            return _store.SaveConfig(config);
        }

        public void DeleteConfig(long id) => _store.DeleteConfig(id);

        // ================= 模板列头加载 =================

        public TemplatePreview LoadTemplate(Stream excel, int? sheetIndex = null, int? headerRowIndex = null, int? maxSampleRows = null)
            => LoadTemplate(excel, sheetIndex, headerRowIndex, null, maxSampleRows);

        public TemplatePreview LoadTemplate(Stream excel, int? sheetIndex, int? headerRowIndex, string language, int? maxSampleRows = null)
        {
            if (excel == null) throw new ArgumentNullException(nameof(excel));

            var lang = _catalog.ResolveLanguage(string.IsNullOrWhiteSpace(language)
                ? _options.DefaultLanguage : language);

            var maxSample = maxSampleRows ?? (_options != null ? _options.MaxSampleRows : 6);
            var hr = headerRowIndex ?? 0;
            var sheetIdx = sheetIndex ?? 0;

            var preview = new TemplatePreview { HeaderRowIndex = hr, SheetIndex = sheetIdx };

            var wb = ExcelWorkbookReader.Open(excel);
            try
            {
                preview.SheetNames = ExcelWorkbookReader.GetSheetNames(wb).ToArray();
                if (!ExcelWorkbookReader.ResolveSheet(wb, null, sheetIdx, out var sheet, out var actualIdx, out var actualName))
                {
                    preview.Message = _catalog.Get(lang, "template.noSheet");
                    return preview;
                }

                preview.SheetIndex = actualIdx;
                preview.SheetName = actualName;

                var columns = ExcelWorkbookReader.ReadHeaders(sheet, hr);
                preview.Columns = columns;
                preview.ColumnCount = columns.Count > 0 ? columns.Max(c => c.Index) + 1 : 0;

                var width = preview.ColumnCount > 0 ? preview.ColumnCount : 0;
                preview.SampleRows = ExcelWorkbookReader.ReadSampleRows(sheet, hr, maxSample, width);

                // 示例值友好化：DateTime → 可读字符串
                foreach (var row in preview.SampleRows)
                {
                    for (var j = 0; j < row.Length; j++)
                        row[j] = Friendly(row[j]);
                }

                preview.DataRowCount = ExcelWorkbookReader.CountDataRows(sheet, hr);
                return preview;
            }
            finally
            {
                if (wb is IDisposable d) d.Dispose();
            }
        }

        // ================= 解析 =================

        public ExcelParseResult Parse(Stream excel) => Parse(excel, null, null);


        public ExcelParseResult ParseByLanguage(Stream excel, string language) => Parse(excel, null, language);

        public ExcelParseResult Parse(Stream excel, string language, string configName="")
        {
            if (excel == null) throw new ArgumentNullException(nameof(excel));

            // 解析返回消息的语系：未传/未知时回退默认（en）
            var lang = _catalog.ResolveLanguage(string.IsNullOrWhiteSpace(language)
                ? _options.DefaultLanguage : language);
            string Tr(string key, params object[] args) => _catalog.Get(lang, key, args);

            var result = new ExcelParseResult { Success = false, Message = Tr("parse.failed") };

            var wb = ExcelWorkbookReader.Open(excel);
            try
            {
                var allConfigs = _store.ListConfigs();
                if (allConfigs.Count == 0)
                {
                    result.Message = Tr("parse.noConfig");
                    return result;
                }

                // 1) 确定命中的配置
                ImportConfig config;
                MatchCandidate candidate = null;
                ISheet parseSheet;
                int headerRow;

                if (!string.IsNullOrWhiteSpace(configName))
                {
                    config = _store.GetByNameOrEntity(configName);
                    if (config == null)
                    {
                        result.Message = Tr("parse.configNotFound", configName);
                        return result;
                    }
                    parseSheet = ResolveParseSheet(wb, config, out headerRow);
                    if (parseSheet == null)
                    {
                        result.Message = Tr("parse.sheetNotFound");
                        return result;
                    }
                }
                else
                {
                    // 自动匹配：对每个配置按其自身 sheet/表头行设置读取列头并计分
                    candidate = AutoMatch(wb, allConfigs);
                    if (candidate == null || candidate.Config == null)
                    {
                        result.Message = "无法匹配任何导入配置：请先在配置页中为实体绑定字段。";
                        return result;
                    }

                    config = candidate.Config;
                    result.MatchScore = candidate.Score;
                    if (candidate.Score < _options.MatchThreshold)
                    {
                        result.Message = Tr("parse.lowMatch", candidate.Matched, candidate.Expected, candidate.Score, _options.MatchThreshold);
                        result.ConfigName = config.Name;
                        result.EntityName = config.EntityName;
                        result.ConfigId = config.Id;
                        return result;
                    }

                    parseSheet = ResolveParseSheet(wb, config, out headerRow);
                    if (parseSheet == null)
                    {
                        result.Message = Tr("parse.sheetNotFound");
                        result.ConfigName = config.Name;
                        return result;
                    }
                }

                // 2) 组装列映射
                result.ConfigId = config.Id;
                result.ConfigName = config.Name;
                result.EntityName = config.EntityName;
                result.Success = true;

                var fileHeaders = ExcelWorkbookReader.ReadHeaders(parseSheet, headerRow);
                var headerByNorm = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var headerByIndex = new Dictionary<int, string>();
                foreach (var h in fileHeaders)
                {
                    var norm = ConfigMatcher.Normalize(h.Text);
                    if (norm.Length > 0 && !headerByNorm.ContainsKey(norm))
                        headerByNorm[norm] = h.Index;
                    if (!headerByIndex.ContainsKey(h.Index))
                        headerByIndex[h.Index] = h.Text;
                }
                var parseSheetName = parseSheet?.SheetName ?? "";

                var targets = new List<MappedColumn>();
                foreach (var f in config.Fields)
                {
                    if (string.IsNullOrWhiteSpace(f.Field)) continue; // 未绑定实体字段的列不导入
                    if (string.IsNullOrWhiteSpace(f.Header)) continue;

                    var col = ResolveColumn(f, fileHeaders, headerByNorm);
                    targets.Add(new MappedColumn { Mapping = f, ColumnIndex = col });
                }

                // 3) 逐行读取并转换
                var maxRows = _options.MaxDataRows;
                var dataRows = ExcelWorkbookReader.ReadRows(parseSheet, headerRow, maxRows);

                result.TotalRows = dataRows.Count;
                foreach (var dataRow in dataRows)
                {
                    var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    var rowHasError = false;

                    foreach (var t in targets)
                    {
                        var f = t.Mapping;
                        object raw = null;
                        if (t.ColumnIndex >= 0)
                            dataRow.Values.TryGetValue(t.ColumnIndex, out raw);

                        if (ExcelValueConverters.IsEmptyValue(raw))
                        {
                            if (f.Required)
                            {
                                rowHasError = true;
                                AddRowError(result, dataRow, t, raw, Tr("field.required"), headerByIndex, parseSheetName);
                            }
                            dict[f.Field] = null;
                            continue;
                        }

                        var val = ExcelValueConverters.ConvertToFieldType(raw, f.DataType, out var errMsg);
                        if (val == null && !string.IsNullOrEmpty(errMsg))
                        {
                            rowHasError = true;
                            AddRowError(result, dataRow, t, raw, Tr("field.formatError", Tr(errMsg)), headerByIndex, parseSheetName);
                        }
                        else if (f.Required && ExcelValueConverters.IsEmptyValue(val))
                        {
                            rowHasError = true;
                            AddRowError(result, dataRow, t, raw, Tr("field.required"), headerByIndex, parseSheetName);
                        }

                        dict[f.Field] = Friendly(val);
                    }

                    if (rowHasError) result.ErrorRows++;
                    else result.SuccessRows++;
                    result.Rows.Add(dict);
                }

                if (result.ErrorRows == 0)
                    result.Message = Tr("parse.completed", result.Rows.Count, result.SuccessRows, 0);
                else if (result.Errors.Count < result.ErrorRows)
                    result.Message = Tr("parse.completedTooMany", result.TotalRows, result.SuccessRows, result.ErrorRows, result.Errors.Count);
                else
                    result.Message = Tr("parse.completed", result.TotalRows, result.SuccessRows, result.ErrorRows);
                return result;
            }
            finally
            {
                if (wb is IDisposable d) d.Dispose();
            }
        }

        public string ParseToJson(Stream excel) => ParseToJson(excel, null, null);

        public string ParseToJson(Stream excel, string language, string configName="") => Parse(excel, language, configName).ToJson();

        public string ParseToJsonByLanguage(Stream excel, string language) => ParseByLanguage(excel, language).ToJson();

        // ================= 私有辅助 =================

        private ISheet ResolveParseSheet(IWorkbook wb, ImportConfig config, out int headerRow)
        {
            headerRow = Math.Max(0, config.HeaderRowIndex);
            ExcelWorkbookReader.ResolveSheet(wb, config.SheetName, config.SheetIndex, out var sheet, out _, out _);
            return sheet;
        }

        private static MatchCandidate AutoMatch(IWorkbook wb, IReadOnlyList<ImportConfig> allConfigs)
        {
            MatchCandidate best = null;
            foreach (var cfg in allConfigs)
            {
                if (cfg.BoundFieldCount == 0) continue;
                ExcelWorkbookReader.ResolveSheet(wb, cfg.SheetName, cfg.SheetIndex, out var sheet, out _, out _);
                if (sheet == null) continue;

                var hr = Math.Max(0, cfg.HeaderRowIndex);
                var headers = ExcelWorkbookReader.ReadHeaders(sheet, hr);
                var texts = headers.Select(h => h.Text).ToList();

                var cand = ConfigMatcher.BestMatch(new List<ImportConfig> { cfg }, texts);
                if (cand == null) continue;

                if (best == null ||
                    cand.Score > best.Score ||
                    (Math.Abs(cand.Score - best.Score) < 1e-9 && cand.Matched > best.Matched) ||
                    (Math.Abs(cand.Score - best.Score) < 1e-9 && cand.Matched == best.Matched && cand.BoundMatched > best.BoundMatched) ||
                    (Math.Abs(cand.Score - best.Score) < 1e-9 && cand.Matched == best.Matched && cand.BoundMatched == best.BoundMatched && cand.OrderMatch && !best.OrderMatch) ||
                    (Math.Abs(cand.Score - best.Score) < 1e-9 && cand.Matched == best.Matched && cand.BoundMatched == best.BoundMatched && cand.OrderMatch == best.OrderMatch && cand.Config.Id < best.Config.Id))
                {
                    best = cand;
                }
            }
            return best;
        }

        /// <summary>决定某绑定映射实际应读取的列序号；无法确定时返回 -1。</summary>
        private static int ResolveColumn(ImportFieldMapping f, IList<HeaderColumn> fileHeaders, IDictionary<string, int> headerByNorm)
        {
            var norm = ConfigMatcher.Normalize(f.Header);
            if (norm.Length > 0 && headerByNorm.TryGetValue(norm, out var idx))
                return idx;

            // 位置兜底：仅当该位置表头为空或与绑定列头一致时使用
            var at = fileHeaders.FirstOrDefault(h => h.Index == f.ColumnIndex);
            if (at == null) return f.ColumnIndex; // 该位置无表头 → 仍按位置取数
            if (string.IsNullOrWhiteSpace(at.Text)) return f.ColumnIndex;
            var atNorm = ConfigMatcher.Normalize(at.Text);
            if (string.Equals(atNorm, norm, StringComparison.Ordinal)) return f.ColumnIndex;

            return -1; // 表头既不匹配也不为空，说明列可能已变动
        }

        /// <summary>追加一条行级错误（携带 Excel 列表头/单元格坐标/原始值等定位信息）。</summary>
        private static void AddRowError(ExcelParseResult result, SheetDataRow row, MappedColumn target, object raw,
            string message, IDictionary<int, string> headerByIndex, string sheetName)
        {
            if (result.Errors.Count >= MaxDetailErrors) return;

            var mapping = target?.Mapping;
            var col = target?.ColumnIndex ?? -1;

            string header = null;
            if (col >= 0 && headerByIndex.TryGetValue(col, out var ht) && !string.IsNullOrWhiteSpace(ht))
                header = ht;
            if (string.IsNullOrEmpty(header) && mapping != null && !string.IsNullOrWhiteSpace(mapping.Header))
                header = mapping.Header;

            string rawText = null;
            if (raw != null && !ExcelValueConverters.IsEmptyValue(raw))
                rawText = ExcelValueConverters.ToText(raw);

            result.Errors.Add(new ParseRowError
            {
                RowNumber = row.RowNumber,
                Sheet = string.IsNullOrWhiteSpace(sheetName) ? null : sheetName,
                Cell = col >= 0 ? ToExcelCell(col, row.RowNumber) : null,
                ColumnIndex = col >= 0 ? (int?)col : null,
                ColumnHeader = string.IsNullOrWhiteSpace(header) ? null : header,
                Field = mapping?.Field ?? "",
                RawValue = rawText,
                Message = message
            });
        }

        /// <summary>把 0 基列序号 + Excel 物理行号转为单元格坐标，如 (4,3) → "E3"。</summary>
        private static string ToExcelCell(int columnIndex, int rowNumber)
        {
            var letters = "";
            var n = columnIndex + 1; // 1 基
            while (n > 0)
            {
                var rem = (n - 1) % 26;
                letters = (char)('A' + rem) + letters;
                n = (n - 1) / 26;
            }
            return letters + rowNumber;
        }

        private static string rowErrorsSummary(ExcelParseResult result)
        {
            if (result.ErrorRows == 0) return null;
            if (result.Errors.Count < result.ErrorRows)
                return $"解析完成，共 {result.TotalRows} 行，成功 {result.SuccessRows} 行，失败 {result.ErrorRows} 行（错误过多，仅展示前 {result.Errors.Count} 条）。";
            return $"解析完成，共 {result.TotalRows} 行，成功 {result.SuccessRows} 行，失败 {result.ErrorRows} 行。";
        }

        /// <summary>把 DateTime 转为可读字符串，便于输出 JSON；其余类型原样返回。</summary>
        private static object Friendly(object v)
        {
            if (v is DateTime dt) return ExcelValueConverters.FormatDate(dt);
            return v;
        }

        private sealed class MappedColumn
        {
            public ImportFieldMapping Mapping { get; set; }
            public int ColumnIndex { get; set; }
        }
    }
}
