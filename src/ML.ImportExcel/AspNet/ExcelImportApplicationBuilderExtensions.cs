using System;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using ML.ImportExcel.Core.Options;

namespace ML.ImportExcel.AspNetCore
{
    /// <summary>
    /// 挂载 Excel 导入配置网页与 API。
    /// 典型用法：
    ///   services.AddExcelImport();                       // ConfigureServices
    ///   app.UseExcelImport();                            // Configure，默认路径 /excelimport
    /// 之后访问：http://host/excelimport  即可打开配置页面。
    /// </summary>
    public static class ExcelImportApplicationBuilderExtensions
    {
        public static IApplicationBuilder UseExcelImport(this IApplicationBuilder app,
            Action<ExcelImportOptions> configure = null)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            var options = app.ApplicationServices.GetService<ExcelImportOptions>();
            if (options == null)
                throw new InvalidOperationException(
                    "未找到 Excel 导入组件注册。请先在 ConfigureServices 中调用 services.AddExcelImport()，再调用 app.UseExcelImport()。");

            configure?.Invoke(options);

            // 规范化挂载路径
            var mount = (options.MountPath ?? "").Trim();
            if (!mount.StartsWith("/", StringComparison.Ordinal)) mount = "/" + mount;
            options.MountPath = mount.TrimEnd('/');
            if (options.MountPath.Length == 0) options.MountPath = "/";
            if (options.MatchThreshold <= 0) options.MatchThreshold = 0.01;
            if (options.MatchThreshold > 1) options.MatchThreshold = 1;

            app.UseMiddleware<ExcelImportMiddleware>();
            return app;
        }
    }
}
