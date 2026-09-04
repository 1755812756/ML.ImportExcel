using System;
using Microsoft.Extensions.DependencyInjection;
using ML.ImportExcel.Core;
using ML.ImportExcel.Core.Options;
using ML.ImportExcel.Core.Storage;

namespace ML.ImportExcel.AspNetCore
{
    /// <summary>
    /// 在 DI 中注册 Excel 导入组件（配置选项 + 解析服务 + SQLite 存储）。
    /// 典型用法：services.AddExcelImport(o => o.DatabasePath = "d:/data/excelimport.db");
    /// 未指定 DatabasePath 时，库文件生成在当前运行目录（程序集所在目录）下。
    /// </summary>
    public static class ExcelImportServiceCollectionExtensions
    {
        public static IServiceCollection AddExcelImport(this IServiceCollection services,
            Action<ExcelImportOptions> configure = null)
        {
            if (services == null) throw new ArgumentNullException(nameof(services));

            var options = new ExcelImportOptions();
            configure?.Invoke(options);
            services.AddSingleton(options);

            services.AddSingleton<IExcelImportService>(sp =>
            {
                var opt = sp.GetRequiredService<ExcelImportOptions>();

                // 未配置路径时：生成到“当前运行目录”（程序集所在目录，开发时即 bin 输出目录）
                var db = string.IsNullOrWhiteSpace(opt.DatabasePath)
                    ? opt.ResolveDefaultDatabasePath()
                    : opt.DatabasePath;

                var store = new SqliteImportStore(db);
                return new ExcelImportService(store, opt);
            });

            return services;
        }
    }
}
