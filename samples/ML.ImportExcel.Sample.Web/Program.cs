using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ML.ImportExcel.AspNetCore;
using ML.ImportExcel.Core;
using ML.ImportExcel.Core.Models;
using ML.ImportExcel.Core.Options;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

var builder = WebApplication.CreateBuilder(args);

// ============ ① 注册 Excel 导入组件 ============
builder.Services.AddExcelImport(o =>
{
    o.MountPath = "/excelimport";          // 配置网页挂载路径（默认值）
    o.PageTitle = "Excel 导入配置中心";
    o.MatchThreshold = 0.6;
    // 多语言：默认自动生成到“运行目录\ExcelLang”（内置 en/zh），发布后可直接修改该目录下的
    // json 或新增语系文件（如 ja.json）即生效；也可通过 o.LanguageDirectory 指定其它目录
    // o.DatabasePath = "d:/data/excelimport.db";   // 可选：自定义 SQLite 路径
});

var app = builder.Build();

// 首次运行时写入两份示例导入配置（仅当库为空）
SeedSampleConfigs(app.Services);

// ============ ② 注入配置网页 / API（默认 /excelimport）============
app.UseExcelImport();

// 简易首页：提供链接与示例 Excel 下载（用于演示“加载列头→绑定→试解析”）
app.MapGet("/", async ctx =>
{
    var html = new StringBuilder();
    html.AppendLine("<html><head><meta charset='utf-8'><title>ML.ImportExcel 示例</title>");
    html.AppendLine("<style>body{font-family:'Segoe UI','Microsoft YaHei';margin:40px;line-height:1.8}" +
                    "a{color:#2563eb;margin-right:10px}</style></head><body>");
    html.AppendLine("<h2>📊 ML.ImportExcel · Excel 导入解析示例站点</h2>");
    html.AppendLine("<p><a href='/excelimport' target='_blank'>➜ 打开 Excel 导入配置页（/excelimport）</a></p>");
    html.AppendLine("<p>示例 Excel 下载：</p><ul>");
    html.AppendLine("<li><a href='/demo/student.xlsx'>学生信息.xlsx</a>（列头：学号/姓名/性别/年龄/班级/入学日期）</li>");
    html.AppendLine("<li><a href='/demo/employee.xlsx'>员工信息.xlsx</a>（列头：工号/姓名/部门/入职日期/月薪/在职）</li>");
    html.AppendLine("</ul>");
    html.AppendLine("<p class='hint' style='color:#64748b'>使用流程：打开配置页 → 选择已有配置或上传示例 Excel 载入列头 → 绑定实体字段 → 保存 → 再上传示例 Excel 点“试解析”，即可看到自动匹配实体后的 JSON。</p>");
    html.AppendLine("</body></html>");
    ctx.Response.ContentType = "text/html; charset=utf-8";
    await ctx.Response.WriteAsync(html.ToString());
});

// 生成并下载演示 Excel
app.MapGet("/demo/student.xlsx", async ctx => { await DownloadAsync(ctx, BuildStudentExcel(), "学生信息.xlsx"); });
app.MapGet("/demo/employee.xlsx", async ctx => { await DownloadAsync(ctx, BuildEmployeeExcel(), "员工信息.xlsx"); });
app.MapGet("/demo/employee-errors.xlsx", async ctx => { await DownloadAsync(ctx, BuildEmployeeErrorExcel(), "员工错误示例.xlsx"); });

app.Run();

// ==================== 种子示例配置 ====================
static void SeedSampleConfigs(IServiceProvider sp)
{
    using var scope = sp.CreateScope();
    var svc = scope.ServiceProvider.GetRequiredService<IExcelImportService>();
    if (svc.ListConfigs().Count > 0) return;

    // 文件夹分组演示：教学管理/学生、人事管理
    var teachFolder = svc.SaveFolder(new ImportFolder { Name = "教学管理" });
    var stuFolder = svc.SaveFolder(new ImportFolder { Name = "学生", ParentId = teachFolder.Id });
    var hrFolder = svc.SaveFolder(new ImportFolder { Name = "人事管理" });

    // 学生信息（entity: Student）
    svc.SaveConfig(new ImportConfig
    {
        Name = "学生信息",
        EntityName = "Student",
        FolderId = stuFolder.Id,
        Description = "学生基础信息导入模板",
        SheetName = "学生",
        HeaderRowIndex = 0,
        Fields =
        {
            new ImportFieldMapping { ColumnIndex = 0, Header = "学号",   Field = "StudentNo",  Required = true,  DataType = ExcelFieldDataTypes.String },
            new ImportFieldMapping { ColumnIndex = 1, Header = "姓名",   Field = "Name",       Required = true,  DataType = ExcelFieldDataTypes.String },
            new ImportFieldMapping { ColumnIndex = 2, Header = "性别",   Field = "Gender",     Required = false, DataType = ExcelFieldDataTypes.String },
            new ImportFieldMapping { ColumnIndex = 3, Header = "年龄",   Field = "Age",        Required = false, DataType = ExcelFieldDataTypes.Int },
            new ImportFieldMapping { ColumnIndex = 4, Header = "班级",   Field = "ClassName",  Required = false, DataType = ExcelFieldDataTypes.String },
            new ImportFieldMapping { ColumnIndex = 5, Header = "入学日期", Field = "EnrollDate", Required = false, DataType = ExcelFieldDataTypes.DateTime }
        }
    });

    // 员工信息（entity: Employee）
    svc.SaveConfig(new ImportConfig
    {
        Name = "员工信息",
        EntityName = "Employee",
        FolderId = hrFolder.Id,
        Description = "员工花名册导入模板",
        HeaderRowIndex = 0,
        Fields =
        {
            new ImportFieldMapping { ColumnIndex = 0, Header = "工号",   Field = "EmployeeNo", Required = true,  DataType = ExcelFieldDataTypes.String },
            new ImportFieldMapping { ColumnIndex = 1, Header = "姓名",   Field = "Name",       Required = true,  DataType = ExcelFieldDataTypes.String },
            new ImportFieldMapping { ColumnIndex = 2, Header = "部门",   Field = "Department", Required = false, DataType = ExcelFieldDataTypes.String },
            new ImportFieldMapping { ColumnIndex = 3, Header = "入职日期", Field = "HireDate",  Required = false, DataType = ExcelFieldDataTypes.DateTime },
            new ImportFieldMapping { ColumnIndex = 4, Header = "月薪",   Field = "Salary",     Required = false, DataType = ExcelFieldDataTypes.Decimal },
            new ImportFieldMapping { ColumnIndex = 5, Header = "在职",   Field = "IsActive",   Required = false, DataType = ExcelFieldDataTypes.Bool }
        }
    });
}

// ==================== 生成示例工作簿 ====================
static byte[] BuildStudentExcel()
{
    using var wb = new XSSFWorkbook();
    var dateStyle = wb.CreateCellStyle();
    dateStyle.DataFormat = wb.CreateDataFormat().GetFormat("yyyy-mm-dd");
    var sheet = wb.CreateSheet("学生");
    string[] headers = { "学号", "姓名", "性别", "年龄", "班级", "入学日期" };
    var headerRow = sheet.CreateRow(0);
    for (var i = 0; i < headers.Length; i++) headerRow.CreateCell(i).SetCellValue(headers[i]);

    var data = new object[][]
    {
        new object[] { "2023001", "张三", "男", 18, "高一(1)班", new DateTime(2023, 9, 1) },
        new object[] { "2023002", "李四", "女", 17, "高一(1)班", new DateTime(2023, 9, 1) },
        new object[] { "2023003", "王五", "男", 18, "高一(2)班", new DateTime(2023, 9, 1) },
        new object[] { "2023004", "赵六", "", 19, "高一(2)班", new DateTime(2022, 9, 1) }
    };
    for (var r = 0; r < data.Length; r++)
    {
        var row = sheet.CreateRow(r + 1);
        for (var c = 0; c < data[r].Length; c++)
        {
            var cell = row.CreateCell(c);
            var v = data[r][c];
            switch (v)
            {
                case DateTime dt: cell.SetCellValue(dt); cell.CellStyle = dateStyle; break;
                case int i2: cell.SetCellValue(i2); break;
                case double d2: cell.SetCellValue(d2); break;
                case bool b2: cell.SetCellValue(b2); break;
                default: cell.SetCellValue(Convert.ToString(v) ?? ""); break;
            }
        }
    }
    for (var i = 0; i < headers.Length; i++) sheet.AutoSizeColumn(i);

    using var ms = new MemoryStream();
    wb.Write(ms, true);
    return ms.ToArray();
}

static byte[] BuildEmployeeExcel()
{
    using var wb = new XSSFWorkbook();
    var dateStyle = wb.CreateCellStyle();
    dateStyle.DataFormat = wb.CreateDataFormat().GetFormat("yyyy-mm-dd");
    var sheet = wb.CreateSheet("员工");
    string[] headers = { "工号", "姓名", "部门", "入职日期", "月薪", "在职" };
    var headerRow = sheet.CreateRow(0);
    for (var i = 0; i < headers.Length; i++) headerRow.CreateCell(i).SetCellValue(headers[i]);

    var data = new object[][]
    {
        new object[] { "E1001", "陈一", "研发部", new DateTime(2020, 3, 12), 15000m, true },
        new object[] { "E1002", "林二", "市场部", new DateTime(2021, 7, 1), 12000m, true },
        new object[] { "E1003", "周三", "研发部", new DateTime(2019, 11, 20), 18000m, false }
    };
    for (var r = 0; r < data.Length; r++)
    {
        var row = sheet.CreateRow(r + 1);
        for (var c = 0; c < data[r].Length; c++)
        {
            var cell = row.CreateCell(c);
            var v = data[r][c];
            switch (v)
            {
                case DateTime dt: cell.SetCellValue(dt); cell.CellStyle = dateStyle; break;
                case int i2: cell.SetCellValue(i2); break;
                case decimal m2: cell.SetCellValue((double)m2); break;
                case bool b2: cell.SetCellValue(b2); break;
                default: cell.SetCellValue(Convert.ToString(v) ?? ""); break;
            }
        }
    }
    for (var i = 0; i < headers.Length; i++) sheet.AutoSizeColumn(i);

    using var ms = new MemoryStream();
    wb.Write(ms, true);
    return ms.ToArray();
}

static async Task DownloadAsync(HttpContext ctx, byte[] bytes, string fileName)
{
    ctx.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    ctx.Response.Headers["Content-Disposition"] = "attachment; filename*=UTF-8''" + Uri.EscapeDataString(fileName);
    await ctx.Response.Body.WriteAsync(bytes, 0, bytes.Length);
}

/// <summary>构造一份“含有错误数据”的员工表，便于演示解析错误定位（月薪非数字、工号为空等）。</summary>
static byte[] BuildEmployeeErrorExcel()
{
    using var wb = new XSSFWorkbook();
    var dateStyle = wb.CreateCellStyle();
    dateStyle.DataFormat = wb.CreateDataFormat().GetFormat("yyyy-mm-dd");
    var sheet = wb.CreateSheet("员工");
    string[] headers = { "工号", "姓名", "部门", "入职日期", "月薪", "在职" };
    var headerRow = sheet.CreateRow(0);
    for (var i = 0; i < headers.Length; i++) headerRow.CreateCell(i).SetCellValue(headers[i]);

    var data = new object[][]
    {
        new object[] { "E2001", "钱七", "财务部", new DateTime(2021, 5, 10), 9000m, true },
        new object[] { "E2002", "孙八", "财务部", new DateTime(2022, 6, 1), "不是数字", true },
        new object[] { "",      "周五", "行政部", new DateTime(2020, 1, 15), 8000m, false },
        new object[] { "E2004", "吴九", "行政部", "2023-02-30", 6000m, true }
    };
    for (var r = 0; r < data.Length; r++)
    {
        var row = sheet.CreateRow(r + 1);
        for (var c = 0; c < data[r].Length; c++)
        {
            var cell = row.CreateCell(c);
            var v = data[r][c];
            if (v is string s && s.Length == 0) { row.CreateCell(c); continue; }
            switch (v)
            {
                case DateTime dt: cell.SetCellValue(dt); cell.CellStyle = dateStyle; break;
                case int i2: cell.SetCellValue(i2); break;
                case decimal m2: cell.SetCellValue((double)m2); break;
                case bool b2: cell.SetCellValue(b2); break;
                default: cell.SetCellValue(Convert.ToString(v) ?? ""); break;
            }
        }
    }
    for (var i = 0; i < headers.Length; i++) sheet.AutoSizeColumn(i);

    using var ms = new MemoryStream();
    wb.Write(ms, true);
    return ms.ToArray();
}
