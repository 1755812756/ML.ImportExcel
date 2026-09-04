using System;
using System.IO;

namespace ML.ImportExcel.Core.Options
{
    /// <summary>
    /// Excel 导入组件全局配置。
    /// 说明：
    ///  - 当以 ASP.NET Core 中间件方式注入时，可直接通过 app.UseExcelImport(o => ...) 配置；
    ///  - 在非 Web（控制台等）环境下使用 IExcelImportService 时，可在
    ///    services.AddExcelImport(o => ...) 中配置，或在构造时传入。
    /// </summary>
    public sealed class ExcelImportOptions
    {
        /// <summary>配置网页挂载的相对路径前缀，默认 /excelimport。</summary>
        public string MountPath { get; set; } = "/excelimport";

        /// <summary>
        /// SQLite 数据库文件路径。
        /// 为 null 时：自动生成到“当前运行目录”（即程序集所在目录 AppContext.BaseDirectory，
        /// 开发时位于 bin 输出目录，发布后位于发布目录）下的 excelimport.db。
        /// 可通过本属性显式指定其它路径覆盖。
        /// </summary>
        public string DatabasePath { get; set; }

        /// <summary>自动匹配时列头命中率阈值(0~1)，默认 0.6。</summary>
        public double MatchThreshold { get; set; } = 0.6;

        /// <summary>网页“加载列头”预览时最多展示的示例数据行数，默认 6。</summary>
        public int MaxSampleRows { get; set; } = 6;

        /// <summary>单次解析最多处理的数据行数，null 表示不限，默认 200000。</summary>
        public int? MaxDataRows { get; set; } = 200000;

        /// <summary>是否在解析结果中包含每一行的原始 Excel 行号信息，默认 true。</summary>
        public bool IncludeRowNumberInResult { get; set; } = true;

        /// <summary>网页标题，默认 “Excel 导入配置中心”。</summary>
        public string PageTitle { get; set; } = "Excel 导入配置中心";

        /// <summary>
        /// 多语言包目录（可选）。为 null 时自动使用“运行目录\ExcelLang”（程序集所在目录下，
        /// 开发时即 bin 输出目录），并把内置 en/zh 词条生成到该目录，方便直接修改/新增语系。
        /// 也可指定其它目录（同样会生成内置词条）。目录内形如 {culture}.json 会覆盖内置文案。
        /// </summary>
        public string LanguageDirectory { get; set; }

        /// <summary>解析等接口未传 language 时使用的默认语系；未知语系亦回退到此值（默认 en）。</summary>
        public string DefaultLanguage { get; set; } = ML.ImportExcel.Core.Localization.MessageCatalog.DefaultLanguage;

        /// <summary>返回默认数据库文件路径（生成在当前运行目录下，即程序集所在目录）。</summary>
        internal string ResolveDefaultDatabasePath()
        {
            var dir = AppContext.BaseDirectory;
            if (string.IsNullOrWhiteSpace(dir)) dir = Environment.CurrentDirectory;
            return Path.Combine(dir, "excelimport.db");
        }

        /// <summary>深度复制一份配置，供各服务持有独立实例。</summary>
        internal ExcelImportOptions Clone()
        {
            return new ExcelImportOptions
            {
                MountPath = string.IsNullOrEmpty(MountPath) ? "/excelimport" : MountPath,
                DatabasePath = DatabasePath,
                MatchThreshold = MatchThreshold,
                MaxSampleRows = MaxSampleRows,
                MaxDataRows = MaxDataRows,
                IncludeRowNumberInResult = IncludeRowNumberInResult,
                PageTitle = PageTitle,
                LanguageDirectory = LanguageDirectory,
                DefaultLanguage = string.IsNullOrWhiteSpace(DefaultLanguage)
                    ? ML.ImportExcel.Core.Localization.MessageCatalog.DefaultLanguage
                    : DefaultLanguage
            };
        }
    }
}
