# ML.ImportExcel · Excel 导入解析组件

一个面向 **.NET 5 / .NET 6 / .NET 7 / .NET 8 / .NET 9（ASP.NET Core）** 的 NuGet 组件，
用于实现 **Excel 导入配置 → 列头加载 → 实体字段绑定 → 数据解析输出 JSON** 的完整闭环。

在 ASP.NET Core 应用中添加本包后，可**自动注入一个网页**（默认 `/excelimport`）：
用户上传模板 Excel 即可加载列头，把“Excel 列头”绑定到“实体字段名称”并存到 **SQLite**；
之后再传入任意 Excel，组件会依据列头**自动匹配出对应的实体/配置**并解析数据，返回 JSON 字符串。

---

## 一、功能总览

| 能力 | 说明 |
| --- | --- |
| 多目标 | `net5.0`/`net6.0`/`net7.0`/`net8.0`/`net9.0`（均含 ASP.NET Core 网页注入；.NET 10 在 .NET 10 SDK 构建时自动追加 net10.0） |
| Excel 读取 | 基于 NPOI，同时支持 `.xls` 与 `.xlsx` |
| 网页注入 | `app.UseExcelImport()` 即可，默认路径 `/excelimport` |
| 树形分组 | 配置以**文件夹树**按功能/业务分组；支持增/改/删文件夹、拖入不同功能目录 |
| 列头绑定 | 网页上传模板 → 读取列头与示例 → 逐列绑定“实体字段名/类型/必填” |
| 存储 | Microsoft.Data.Sqlite，SQLite 库文件（默认 `excelimport.db`，路径可配置） |
| 自动匹配 | 解析时按列头命中率自动匹配已保存的实体/配置 |
| 配置备份/迁移 | 编辑器一键“导出配置 / 导入配置”：导出当前编辑的配置为 JSON；导入载入编辑器预览后保存（同名覆盖更新） |
| 输出 | `ParseToJson(...)` 返回 JSON 字符串（含命中的实体名、命中率、数据行、校验错误） |

---

## 二、目录结构

```
ML.ImportExcel.sln
├─ src/ML.ImportExcel/                  # NuGet 包源码（多目标 net5.0;net6.0;net7.0;net8.0;net9.0；net10 需 .NET10 SDK）
│  ├─ Core/                        # 核心引擎（纯逻辑，无 Web 依赖）
│  │  ├─ ExcelImportService.cs     # 门面服务：配置管理/模板列头/解析/JSON
│  │  ├─ Models/                   # ImportConfig / ImportFieldMapping / TemplatePreview / ExcelParseResult
│  │  ├─ Options/                  # ExcelImportOptions
│  │  ├─ Storage/                  # SqliteImportStore（建表 + CRUD）
│  │  ├─ Excel/                    # NPOI 封装：读列头/读数据/类型转换
│  │  └─ Matching/                 # 列头自动匹配算法
│  ├─ AspNet/                      # AddExcelImport / UseExcelImport / 中间件
│  └─ Web/excelimport-page.html    # 嵌入的配置网页
└─ samples/ML.ImportExcel.Sample.Web/   # 可直接运行的示例站点
```

---

## 三、在 ASP.NET Core 中使用（含网页注入）

### 1. 引用包

```xml
<PackageReference Include="ML.ImportExcel" Version="1.0.0" />
```

### 2. 注册并挂载（Program.cs）

```csharp
using ML.ImportExcel.AspNetCore;
using ML.ImportExcel.Core.Options;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddExcelImport(o =>
{
    o.MountPath = "/excelimport";   // 网页挂载路径（默认）
    o.MatchThreshold = 0.6;         // 自动匹配命中率阈值
    // o.DatabasePath = "d:/data/excelimport.db"; // 可选：不配置时自动生成到“当前运行目录”（bin/发布目录）excelimport.db
});
var app = builder.Build();
app.UseExcelImport();
app.Run();
```

### 3. 打开配置网页

浏览器访问：

```
http://localhost:5000/excelimport
```

页面使用流程：

1. **左侧配置树**：按文件夹分组展示配置（如“教学管理/学生”“人事管理”）。
   - 点“＋文件夹”在当前文件夹下新建子文件夹；选中文件夹后可“改名 / 删除夹”
     （删除文件夹会连同子文件夹删除，其中的配置自动移回根目录）。
   - 点“＋配置”在当前文件夹下新建导入配置。
2. **点击配置节点**：右侧回填该配置并可修改（所属文件夹可用“所属文件夹”下拉调整后保存即完成移动）。
3. **编辑导入配置**：填写“配置名称（唯一）”“实体名称（如 `Student`）”。
4. **上传模板 Excel** → 点“① 载入 Excel 列头”，系统读取列头与示例值。
5. **绑定实体字段**：逐列填写“绑定实体字段名”（即实体属性名），可指定类型与是否必填。
6. **保存配置**：映射连同所属文件夹写入 SQLite。
7. **试解析**：上传待解析的数据 Excel → 点“② 开始解析”，自动匹配实体并返回 JSON。
8. **配置备份 / 迁移（可选）**：“保存配置”右侧的“⬇ 导出配置”可把**当前编辑的配置**下载为 JSON；在另一环境（如正式）打开本页点“⬆ 导入配置”载入到编辑器，预览后点“保存配置”写入（已存在同名配置即覆盖更新）。

### 4. 网页提供的 API（同一路径前缀）

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET  | `/excelimport` | 配置网页（树形文件夹分组） |
| GET  | `/excelimport/api/folders` | 文件夹列表 |
| POST | `/excelimport/api/folders` | 新增/更新文件夹（JSON：`name`、可选 `parentId`；`id>0` 为改名/移动） |
| GET/DELETE | `/excelimport/api/folders/{id}` | 文件夹详情 / 删除（级联子文件夹，配置移回根目录） |
| GET  | `/excelimport/api/configs` | 配置列表（含 `folderId`） |
| POST | `/excelimport/api/configs` | 新增/更新配置（JSON，含 `folderId`/字段映射） |
| GET/DELETE | `/excelimport/api/configs/{id}` | 详情 / 删除 |
| POST | `/excelimport/api/templates/preview` | multipart 上传模板 Excel（`file`、可选 `headerRow`/`sheetIndex`），返回列头+示例 |
| POST | `/excelimport/api/parse` | multipart 上传数据 Excel（`file`、可选 `configName`），返回解析 JSON |

---

## 五、在代码中调用解析（返回 JSON 字符串）

```csharp
using ML.ImportExcel.Core;
using ML.ImportExcel.Core.Options;
using ML.ImportExcel.Core.Storage;

// 非 Web 宿主（控制台等）也可直接构造服务（底层同样是 SQLite 存储）
IExcelImportService svc =
    new ExcelImportService(new SqliteImportStore(@"d:/data/excelimport.db"), new ExcelImportOptions());

using var fs = File.OpenRead(@"d:/tmp/学生导入.xlsx");

// ① 自动匹配实体/配置并返回 JSON
string json = svc.ParseToJson(fs);

// ② 显式指定配置：签名 ParseToJson(stream, language, configName)，language 传空串则用默认语系
string json2 = svc.ParseToJson(fs, "", "学生信息");
```

在 ASP.NET Core 中通过依赖注入获取即可：

```csharp
app.MapPost("/my/import", (HttpContext ctx, IExcelImportService svc) =>
{
    // 保存 IFormFile 到 MemoryStream 后调用 svc.ParseToJson(ms)
});
```

### 返回 JSON 示例

```json
{
  "success": true,
  "message": "解析完成，共 4 行，成功 4 行，失败 0 行。",
  "configId": 1,
  "configName": "学生信息",
  "entityName": "Student",
  "matchScore": 1.0,
  "totalRows": 4,
  "successRows": 4,
  "errorRows": 0,
  "errors": [],
  "rows": [
    { "StudentNo": "2023001", "Name": "张三", "Gender": "男", "Age": 18, "ClassName": "高一(1)班", "EnrollDate": "2023-09-01" }
  ]
}
```

> `rows` 数组每个元素是 `字段名 → 值` 的对象，可直接反序列化为 `List<Dictionary<string,object>>`
> 或结合实体定义（entityName）继续映射为强类型对象。

> **错误明细**：`errors[]` 中每条除 `rowNumber`、`field`、`message` 外，还包含用于定位 Excel 原格的信息：
>
> ```json
> {
>   "rowNumber": 3,
>   "sheet": "员工",
>   "cell": "E3",
>   "columnIndex": 4,
>   "columnHeader": "月薪",
>   "field": "Salary",
>   "rawValue": "不是数字",
>   "message": "字段格式错误：不是有效的数值(Decimal)"
> }
> ```
> 其中 `cell` 为 Excel 单元格坐标（列字母+行号，如 `E3`），`columnHeader` 为该列在 Excel 中的原始列表头，`rawValue` 为单元格原始显示值。

### 指定返回语系（多语言）

所有返回文案均来自 **JSON 语言包**，不再硬编码中文。调用解析方法时可传入语系：

```csharp
// 签名：ParseToJson(stream, language, configName) —— language 在前；configName 为空表示自动匹配实体
string jsonZh = svc.ParseToJson(fs, "zh-CN", "学生信息"); // 简体中文 + 指定配置
string jsonEn = svc.ParseToJson(fs, "en", "学生信息");    // English + 指定配置
string jsonDefault = svc.ParseToJson(fs, "", "学生信息");  // 指定配置 + 默认语系(English)
string jsonAuto = svc.ParseToJson(fs, "zh-CN");            // 自动匹配实体 + 简体中文
string jsonJa  = svc.ParseToJsonByLanguage(fs, "ja");      // 仅按语系（自动匹配，需外部语言包）
```

- 内置语言：`en`（English）、`zh` / `zh-CN`（简体中文）。**未知语系或未传语系默认回退 English**，语言包缺失的词条也自动回退英文。
- 语言包自动生成到**运行目录\ExcelLang**：首次运行会把内置 `en.json`、`zh.json` 自动写入该目录，
  发布后无需重新编译，直接改这里的 json 即可更新词条；往该目录放一个 `{语系}.json`（如 `ja.json`）
  即新增语系（同 key 覆盖内置文案）。也可用 `ExcelImportOptions.LanguageDirectory` 指定其它目录。
- HTTP 接口：`/excelimport/api/parse` 支持 multipart 字段 `language`；配置页“试解析”也提供“返回语系”下拉。
- 语言包 JSON 结构：`{ "key": "带{0}占位符的模板" }`，完整词条参见 `src/ML.ImportExcel/Lang/en.json` 与 `zh.json`。

---

## 六、配置项（ExcelImportOptions）

| 属性 | 默认值 | 说明 |
| --- | --- | --- |
| `MountPath` | `/excelimport` | 网页挂载路径 |
| `DatabasePath` | 空 | SQLite 文件路径；为空时自动生成到**当前运行目录**（程序集所在目录 `AppContext.BaseDirectory`，开发时即 `bin` 输出目录）下的 `excelimport.db` |
| `MatchThreshold` | `0.6` | 自动匹配列头命中率阈值(0~1) |
| `MaxSampleRows` | `6` | 网页列头预览的示例行数 |
| `MaxDataRows` | `200000` | 单次解析最多行数 |
| `PageTitle` | `Excel 导入配置中心` | 网页标题 |
| `LanguageDirectory` | 空 | 多语言目录；为空时自动生成/读取到**运行目录\ExcelLang**（内置 en/zh 首次自动写入）；可指定其它目录 |
| `DefaultLanguage` | `en` | 未传语系时的默认语系；未知语系也回退至此 |

---

## 七、Excel 约定

- 默认在第 **1 行（行号 0）** 放列头；如模板前有标题行，可在配置中调整 `HeaderRowIndex`。
- 支持多个工作表，可指定 `SheetName`（优先）或 `SheetIndex`。
- 支持的数据类型（字段绑定时可选）：`string`（文本）、`int`、`long`、`double`、
  `decimal`（金额）、`datetime`（日期时间，Excel 日期序列或 `yyyy-MM-dd` 等文本均可）、`bool`。
- 日期单元格会读取为日期；普通数字按整数/小数处理；`必填` 字段为空或格式错误会记录到 `errors`。

---

## 八、SQLite 存储

默认数据库文件 `excelimport.db` 自动生成在**当前运行目录**（程序集所在目录，`AppContext.BaseDirectory`；
开发时位于 `bin` 输出目录，发布后位于发布目录）。也可通过 `ExcelImportOptions.DatabasePath` 指定其它路径。
共三张表：

- `import_folders`：文件夹树（`Id/ParentId/Name`，同级名称唯一）
- `import_configs`：配置主表（名称/实体名/`FolderId`/工作表/列头行号/时间戳）
- `import_config_fields`：列头→字段映射（列序号/Excel 列头/实体字段/必填/类型）

外键：删除配置会级联删除其字段映射；删除文件夹会把其下配置移回根目录（`FolderId` 置空）。

> 从旧版升级：`SqliteImportStore` 建库时会自动执行幂等迁移——新增 `import_folders` 表、
> 给 `import_configs` 补充 `FolderId` 列，无需手工处理已有数据库。

---

## 九、构建与打包

```bash
# 构建
dotnet build ML.ImportExcel.sln -c Release

# 生成 NuGet 包（src/ML.ImportExcel/bin/Release/ML.ImportExcel.1.0.0.nupkg）
dotnet pack src/ML.ImportExcel/ML.ImportExcel.csproj -c Release -o artifacts

# 运行示例站点（.NET 8）
dotnet run --project samples/ML.ImportExcel.Sample.Web
# 打开 http://localhost:5000/excelimport

# .NET 6 WebApi 集成测试（独立文件夹，通过本地 NuGet 包引入，见 ML.ImportExcel.Test.Net6）
dotnet run --project ML.ImportExcel.Test.Net6 --urls http://localhost:5198
# 打开 http://localhost:5198/excelimport
```

> 兼容性说明：
> - `net5.0` 资产面向 .NET 5（已停止官方支持，仅为既有 .NET 5 应用提供兼容资产）；`net6.0` / `net7.0` / `net8.0` / `net9.0` 资产面向 ASP.NET Core（各版本应用按最近目标框架选用），自带配置网页注入。
> - `net10.0` 资产：使用 .NET 10 SDK 构建本包时自动追加（无需改代码）。
> - 本地包引入示例：仓库根目录独立项目 `ML.ImportExcel.Test.Net6`（单一项目/文件夹），其 `nuget.config` 把 `..\artifacts` 注册为
>   包源，以 `PackageReference Include="ML.ImportExcel" Version="1.0.0"` 方式验证包内容与注入效果。
