# ML.ImportExcel

**Excel 导入解析组件（支持多语言方案 + 自定义配置）**

面向 .NET 5 / .NET 6 / .NET 7 / .NET 8 / .NET 9（.NET 10 在 .NET 10 SDK 构建时自动追加）的 Excel 导入中间件与引擎：

- 上传模板 Excel → 自动读取**列头与示例** → 逐列绑定“实体字段/类型/必填” → 保存映射；
- 解析时传入 Excel，**按列头自动匹配实体/配置**（对 A~D / A~E / A~F 这类“超集报表”也能精确区分），或显式指定配置；
- 支持 `.xls` / `.xlsx`，数据行输出为 `字段名 → 值` 的 JSON；
- 错误明细携带 **Excel 单元格坐标 / 原始列表头 / 原始值**，方便定位脏数据。

## 功能亮点

| 特性 | 说明 |
| --- | --- |
| 多目标 | `net5.0`~`net9.0`（均含 ASP.NET Core 网页注入，net10 随 .NET10 SDK 自动追加；net5 已停止官方支持） |
| 网页配置注入 | ASP.NET Core 下 `app.UseExcelImport()` 即可，默认 `/excelimport` |
| 树形分组 | 配置按文件夹树组织（可增/改/删文件夹），点配置即可回填修改 |
| 自动/指定匹配 | 传 Excel 自动匹配实体；也可显式指定配置名解析 |
| **配置同步** | 配置页“⬇ 导出配置 / ⬆ 导入配置”（保存配置右侧）：导出**当前编辑的配置**为 JSON；导入载入编辑器预览后保存（同名覆盖更新），便于开发→正式环境迁移 |
| **多语言方案** | 内置 `en`/`zh`，首次运行自动生成到**运行目录 `ExcelLang`** 的 JSON，直接改词条或放 `{语系}.json` 即新增语系；解析方法支持 `language` 参数，未知语系默认回退英文 |
| **自定义配置** | 库文件路径、网页挂载路径、命中率阈值、多语言目录、默认语系、每行/总数上限等均可配置 |
| 存储 | Microsoft.Data.Sqlite（SQLite），默认库文件生成到运行目录，老库自动幂等迁移 |
| 字段自定义 | 列头→实体字段、数据类型（string/int/long/double/decimal/datetime/bool）、必填校验均可在网页配置 |

## ASP.NET Core 快速开始（网页注入）

```csharp
// Program.cs
builder.Services.AddExcelImport(o =>
{
    // o.DatabasePath = "d:/data/excelimport.db";      // 自定义 SQLite
    // o.MountPath = "/excelimport";                   // 网页路径（默认）
    // o.LanguageDirectory = "d:/langs";               // 自定义多语言目录（默认运行目录/ExcelLang）
    // o.DefaultLanguage = "en";                       // 缺省语系（默认 en）
});
var app = builder.Build();
app.UseExcelImport();   // 打开 http://localhost:5000/excelimport 进行配置
```

解析并返回 JSON（多语言）：

```csharp
using ML.ImportExcel.Core;

var svc = ...; // DI: IExcelImportService
using var fs = File.OpenRead("report.xlsx");

string json = svc.ParseToJson(fs);                    // 自动匹配实体（默认 English）
string json2 = svc.ParseToJson(fs, "zh-CN", "报表2"); // 指定配置 + 简体中文
string json3 = svc.ParseToJsonByLanguage(fs, "ja");   // 仅按语系（自动匹配）
```

HTTP 导入（multipart）：

```
POST /api/excel/import
  file       必填  .xls/.xlsx
  configName 可选  指定配置（为空自动匹配）
  language   可选  zh-CN / en / ja …
```

## 多语言 JSON 语言包

语言包结构：`{ "key": "含{0}占位符的模板" }`。

首次运行会自动把内置词条生成到 **`运行目录/ExcelLang`**（`en.json`、`zh.json`），
发布后直接修改即可更新语系；新增语系 = 在该目录放一个 `{语系}.json`（如 `ja.json`），无需重新编译。
缺失词条自动回退英文。

> 参考完整词条见包源码 `src/ML.ImportExcel/Lang/en.json`、`zh.json`。
