using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using ML.ImportExcel.Core;
using ML.ImportExcel.Core.Models;
using ML.ImportExcel.Core.Options;

namespace ML.ImportExcel.AspNetCore
{
    /// <summary>
    /// Excel 导入配置页/API 中间件。
    ///  - {mount}                 GET   配置网页（树形文件夹分组）
    ///  - {mount}/api/folders      GET   文件夹列表 / POST 新增或更新文件夹(JSON)
    ///  - {mount}/api/folders/{id} GET   文件夹详情 / DELETE 删除文件夹(含子文件夹，配置移回根目录)
    ///  - {mount}/api/configs      GET   配置列表 / POST 保存配置(JSON，含 folderId)
    ///  - {mount}/api/configs/{id} GET   详情 / DELETE 删除
    ///  - {mount}/api/templates/preview  POST multipart 上传模板 Excel，返回列头+示例
    ///  - {mount}/api/parse              POST multipart 上传待解析 Excel（可带 configName）
    /// </summary>
    public sealed class ExcelImportMiddleware
    {
        private const string PageResource = "ML.ImportExcel.Web.excelimport-page.html";
        private readonly RequestDelegate _next;

        private static readonly Lazy<string> PageHtml = new Lazy<string>(LoadPage, true);

        private static readonly JsonSerializerOptions JsonOpt = new JsonSerializerOptions
        {
            WriteIndented = false,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        public ExcelImportMiddleware(RequestDelegate next)
        {
            _next = next ?? throw new ArgumentNullException(nameof(next));
        }

        private static string LoadPage()
        {
            var asm = typeof(ExcelImportMiddleware).Assembly;
            using (var s = asm.GetManifestResourceStream(PageResource))
            {
                if (s == null) return "<!DOCTYPE html><html><body><h3>嵌入页面资源缺失：ML.ImportExcel.Web.excelimport-page.html</h3></body></html>";
                using (var r = new StreamReader(s, Encoding.UTF8))
                    return r.ReadToEnd();
            }
        }

        public async Task InvokeAsync(HttpContext context)
        {
            var options = context.RequestServices.GetService<ExcelImportOptions>();
            if (options == null)
            {
                await _next(context);
                return;
            }

            var mount = options.MountPath ?? "/excelimport";
            var path = context.Request.Path.Value ?? "/";

            if (!PathStartsWith(path, mount))
            {
                await _next(context);
                return;
            }

            var relative = path.Length > mount.Length ? path.Substring(mount.Length) : "";
            relative = relative.TrimStart('/');

            // ---------- 页面 ----------
            if (relative.Length == 0 || relative.Equals("index.html", StringComparison.OrdinalIgnoreCase))
            {
                if (!HttpMethods.IsGet(context.Request.Method))
                {
                    context.Response.StatusCode = StatusCodes.Status405MethodNotAllowed;
                    return;
                }
                var html = PageHtml.Value
                    .Replace("{{TITLE}}", string.IsNullOrWhiteSpace(options.PageTitle) ? "Excel 导入配置中心" : options.PageTitle)
                    .Replace("{{MOUNT}}", mount.Length == 0 ? "/" : mount);
                context.Response.ContentType = "text/html; charset=utf-8";
                await context.Response.WriteAsync(html, Encoding.UTF8);
                return;
            }

            // ---------- API ----------
            if (!relative.StartsWith("api/", StringComparison.OrdinalIgnoreCase) && !relative.Equals("api", StringComparison.OrdinalIgnoreCase))
            {
                await _next(context);
                return;
            }

            var segments = relative.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            // segments[0] == "api"
            await RouteApiAsync(context, options, segments.Skip(1).ToArray());
        }

        private static bool PathStartsWith(string path, string mount)
        {
            if (string.IsNullOrEmpty(mount) || mount == "/")
                return path.StartsWith("/", StringComparison.OrdinalIgnoreCase);

            var m = mount.TrimEnd('/');
            return path.Equals(m, StringComparison.OrdinalIgnoreCase)
                || path.StartsWith(m + "/", StringComparison.OrdinalIgnoreCase);
        }

        private async Task RouteApiAsync(HttpContext context, ExcelImportOptions options, string[] seg)
        {
            try
            {
                var svc = context.RequestServices.GetRequiredService<IExcelImportService>();
                if (seg.Length == 0 || (seg[0].Equals("help", StringComparison.OrdinalIgnoreCase) && HttpMethods.IsGet(context.Request.Method)))
                {
                    await WriteJsonAsync(context, new
                    {
                        name = "Excel 导入组件",
                        endpoints = new[]
                        {
                            options.MountPath + "  (GET) 配置网页（树形文件夹分组）",
                            options.MountPath + "/api/folders  (GET/POST)",
                            options.MountPath + "/api/folders/{id}  (GET/DELETE)",
                            options.MountPath + "/api/configs  (GET/POST)",
                            options.MountPath + "/api/configs/{id}  (GET/DELETE)",
                            options.MountPath + "/api/templates/preview  (POST multipart: file, headerRow?)",
                            options.MountPath + "/api/parse  (POST multipart: file, configName?)"
                        }
                    });
                    return;
                }

                if (seg[0].Equals("configs", StringComparison.OrdinalIgnoreCase))
                {
                    await RouteConfigsAsync(context, svc, seg.Skip(1).ToArray());
                    return;
                }

                if (seg[0].Equals("folders", StringComparison.OrdinalIgnoreCase))
                {
                    await RouteFoldersAsync(context, svc, seg.Skip(1).ToArray());
                    return;
                }

                if (seg[0].Equals("templates", StringComparison.OrdinalIgnoreCase))
                {
                    await RouteTemplatesAsync(context, svc, seg.Skip(1).ToArray());
                    return;
                }

                if (seg[0].Equals("parse", StringComparison.OrdinalIgnoreCase))
                {
                    await RouteParseAsync(context, svc, seg.Skip(1).ToArray());
                    return;
                }

                context.Response.StatusCode = StatusCodes.Status404NotFound;
                await WriteJsonAsync(context, new { success = false, message = "未知接口路径。" });
            }
            catch (ArgumentException ex)
            {
                context.Response.StatusCode = StatusCodes.Status400BadRequest;
                await WriteJsonAsync(context, new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                context.Response.StatusCode = StatusCodes.Status500InternalServerError;
                await WriteJsonAsync(context, new { success = false, message = "服务端错误：" + ex.Message });
            }
        }

        // ---------------- 配置管理 ----------------

        private static async Task RouteConfigsAsync(HttpContext context, IExcelImportService svc, string[] seg)
        {
            if (seg.Length == 0)
            {
                if (HttpMethods.IsGet(context.Request.Method))
                {
                    await WriteJsonAsync(context, svc.ListConfigs());
                    return;
                }
                if (HttpMethods.IsPost(context.Request.Method))
                {
                    ImportConfig cfg;
                    try
                    {
                        cfg = await JsonSerializer.DeserializeAsync<ImportConfig>(context.Request.Body, JsonOpt);
                    }
                    catch (Exception ex)
                    {
                        context.Response.StatusCode = StatusCodes.Status400BadRequest;
                        await WriteJsonAsync(context, new { success = false, message = "配置 JSON 解析失败：" + ex.Message });
                        return;
                    }

                    if (cfg == null)
                    {
                        context.Response.StatusCode = StatusCodes.Status400BadRequest;
                        await WriteJsonAsync(context, new { success = false, message = "请求体为空。" });
                        return;
                    }

                    var saved = svc.SaveConfig(cfg);
                    context.Response.StatusCode = StatusCodes.Status200OK;
                    await WriteJsonAsync(context, saved);
                    return;
                }

                context.Response.StatusCode = StatusCodes.Status405MethodNotAllowed;
                return;
            }

            if (long.TryParse(seg[0], out var id))
            {
                if (HttpMethods.IsGet(context.Request.Method))
                {
                    var cfg = svc.GetConfig(id);
                    if (cfg == null)
                    {
                        context.Response.StatusCode = StatusCodes.Status404NotFound;
                        await WriteJsonAsync(context, new { success = false, message = $"配置 {id} 不存在。" });
                        return;
                    }
                    await WriteJsonAsync(context, cfg);
                    return;
                }

                if (HttpMethods.IsDelete(context.Request.Method))
                {
                    svc.DeleteConfig(id);
                    context.Response.StatusCode = StatusCodes.Status200OK;
                    await WriteJsonAsync(context, new { success = true, message = "已删除。" });
                    return;
                }
            }

            context.Response.StatusCode = StatusCodes.Status405MethodNotAllowed;
        }

        // ---------------- 文件夹管理 ----------------

        private static async Task RouteFoldersAsync(HttpContext context, IExcelImportService svc, string[] seg)
        {
            if (seg.Length == 0)
            {
                if (HttpMethods.IsGet(context.Request.Method))
                {
                    await WriteJsonAsync(context, svc.ListFolders());
                    return;
                }
                if (HttpMethods.IsPost(context.Request.Method))
                {
                    ImportFolder folder;
                    try
                    {
                        folder = await JsonSerializer.DeserializeAsync<ImportFolder>(context.Request.Body, JsonOpt);
                    }
                    catch (Exception ex)
                    {
                        context.Response.StatusCode = StatusCodes.Status400BadRequest;
                        await WriteJsonAsync(context, new { success = false, message = "文件夹 JSON 解析失败：" + ex.Message });
                        return;
                    }
                    if (folder == null)
                    {
                        context.Response.StatusCode = StatusCodes.Status400BadRequest;
                        await WriteJsonAsync(context, new { success = false, message = "请求体为空。" });
                        return;
                    }

                    var saved = svc.SaveFolder(folder);
                    context.Response.StatusCode = StatusCodes.Status200OK;
                    await WriteJsonAsync(context, saved);
                    return;
                }

                context.Response.StatusCode = StatusCodes.Status405MethodNotAllowed;
                return;
            }

            if (seg.Length >= 1 && long.TryParse(seg[0], out var id))
            {
                if (HttpMethods.IsGet(context.Request.Method))
                {
                    var folder = svc.GetFolder(id);
                    if (folder == null)
                    {
                        context.Response.StatusCode = StatusCodes.Status404NotFound;
                        await WriteJsonAsync(context, new { success = false, message = $"文件夹 {id} 不存在。" });
                        return;
                    }
                    await WriteJsonAsync(context, folder);
                    return;
                }

                if (HttpMethods.IsDelete(context.Request.Method))
                {
                    svc.DeleteFolder(id);
                    context.Response.StatusCode = StatusCodes.Status200OK;
                    await WriteJsonAsync(context, new { success = true, message = "文件夹已删除，其下配置已移至根目录。" });
                    return;
                }
            }

            context.Response.StatusCode = StatusCodes.Status405MethodNotAllowed;
        }

        // ---------------- 模板列头预览 ----------------

        private static async Task RouteTemplatesAsync(HttpContext context, IExcelImportService svc, string[] seg)
        {
            if (seg.Length == 1 && seg[0].Equals("preview", StringComparison.OrdinalIgnoreCase)
                && HttpMethods.IsPost(context.Request.Method))
            {
                var upload = await ReadUploadAsync(context);
                if (upload.Stream == null)
                {
                    context.Response.StatusCode = StatusCodes.Status400BadRequest;
                    await WriteJsonAsync(context, new { success = false, message = "缺少上传文件(file)。" });
                    return;
                }

                using (upload.Stream)
                {
                    var preview = svc.LoadTemplate(upload.Stream, upload.SheetIndex, upload.HeaderRow, upload.Language);
                    await WriteJsonAsync(context, preview);
                }
                return;
            }

            context.Response.StatusCode = StatusCodes.Status404NotFound;
            await WriteJsonAsync(context, new { success = false, message = "未知接口。" });
        }

        // ---------------- 解析 ----------------

        private static async Task RouteParseAsync(HttpContext context, IExcelImportService svc, string[] seg)
        {
            if (!HttpMethods.IsPost(context.Request.Method))
            {
                context.Response.StatusCode = StatusCodes.Status405MethodNotAllowed;
                return;
            }

            var upload = await ReadUploadAsync(context);
            if (upload.Stream == null)
            {
                context.Response.StatusCode = StatusCodes.Status400BadRequest;
                await WriteJsonAsync(context, new { success = false, message = "缺少上传文件(file)。" });
                return;
            }

            using (upload.Stream)
            {
                var result = svc.Parse(upload.Stream, upload.ConfigName, upload.Language);
                await WriteJsonAsync(context, result);
            }
        }

        // ---------------- 工具 ----------------

        /// <summary>上传文件解析结果。</summary>
        private sealed class UploadInfo
        {
            public MemoryStream Stream { get; set; }
            public int? SheetIndex { get; set; }
            public int? HeaderRow { get; set; }
            public string ConfigName { get; set; }
            public string Language { get; set; }
        }

        /// <summary>读取 multipart 上传文件与可选的 sheetIndex/headerRow/configName/language 参数。</summary>
        private static async Task<UploadInfo> ReadUploadAsync(HttpContext context)
        {
            var info = new UploadInfo();
            if (context.Request.Query.ContainsKey("configName"))
                info.ConfigName = context.Request.Query["configName"].ToString();
            if (context.Request.Query.ContainsKey("language"))
                info.Language = context.Request.Query["language"].ToString();

            if (!context.Request.HasFormContentType)
                return info;

            var form = await context.Request.ReadFormAsync();
            IFormFile file = null;
            if (form.Files != null)
                file = form.Files.GetFile("file");

            if (form.ContainsKey("sheetIndex") && int.TryParse(form["sheetIndex"], out var si)) info.SheetIndex = si;
            if (form.ContainsKey("headerRow") && int.TryParse(form["headerRow"], out var hr)) info.HeaderRow = hr;
            if (form.ContainsKey("configName")) info.ConfigName = form["configName"].ToString();
            if (form.ContainsKey("language")) info.Language = form["language"].ToString();

            if (file == null || file.Length == 0) return info;

            var ms = new MemoryStream();
            using (var src = file.OpenReadStream())
                await src.CopyToAsync(ms);
            ms.Position = 0;
            info.Stream = ms;
            return info;
        }

        private static async Task WriteJsonAsync(HttpContext context, object value)
        {
            context.Response.ContentType = "application/json; charset=utf-8";
            await JsonSerializer.SerializeAsync(context.Response.Body, value, value?.GetType() ?? typeof(object), JsonOpt);
        }
    }
}
