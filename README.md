# Excel2Object

[![install from nuget](http://img.shields.io/nuget/v/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)
[![release](https://img.shields.io/github/release/chsword/Excel2Object.svg?style=flat-square)](https://github.com/chsword/Excel2Object/releases)
[![.NET CI](https://github.com/chsword/Excel2Object/actions/workflows/dotnet-ci.yml/badge.svg)](https://github.com/chsword/Excel2Object/actions/workflows/dotnet-ci.yml)
[![CodeFactor](https://www.codefactor.io/repository/github/chsword/excel2object/badge)](https://www.codefactor.io/repository/github/chsword/excel2object)

Excel 与 .NET 对象互相转换 / Excel convert to .NET Object and vice versa.

[English](README_EN.md) | 中文

- [Top](#excel2object)
    - [安装 NuGet](#安装-nuget)
    - [发布说明和路线图](#发布说明和路线图)
    - [示例代码](#示例代码)
    - [文档](#文档)
    - [开发和发布](#开发和发布)
    - [贡献者](#贡献者)
    - [参考](#参考)

## 平台支持

[![.NET 4.7.2 +](https://img.shields.io/badge/-4.7.2%2B-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.0](https://img.shields.io/badge/-standard2.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.1](https://img.shields.io/badge/-standard2.1-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 6.0](https://img.shields.io/badge/-6.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 8.0](https://img.shields.io/badge/-8.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)

## 安装 NuGet

``` powershell
PM> Install-Package Chsword.Excel2Object
```

或使用 .NET CLI:
``` bash
dotnet add package Chsword.Excel2Object
```

### 命令行工具 excel2obj

``` bash
dotnet tool install -g Chsword.Excel2Object.Cli

excel2obj convert orders.xlsx --output orders.json --typed   # Excel -> JSON
excel2obj convert orders.json --output orders.xlsx           # JSON -> Excel
excel2obj generate-model orders.xlsx --class Order           # 由表头生成带 [ExcelTitle] 的 C# 模型类
```

详见 [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)。

## 发布说明和路线图

### 暂不支持的特性

- [x] CLI 工具 ✅ **v2.2.1 新增** - `dotnet tool install -g Chsword.Excel2Object.Cli`，查看 [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)
- [x] 支持自动列宽 ✅ **v2.0.4 新增**
- [x] 支持 Excel 日期/日期时间/时间格式 ✅ **v2.0.4 新增**，导出为真正的日期单元格 ✅ **v2.4.0 新增** - 查看 [DateTimeFormats.md](DateTimeFormats.md)
- [x] 公式列引用同一工作簿的其他 sheet ✅ **v2.1.0 新增** - 查看 [ExcelFunctions.md](ExcelFunctions.md)
- [x] 公式内置函数库 ✅ **v2.3.0 新增** - 334 个 Excel 函数，10 个类别 - 查看 [ExcelFunctions.md](ExcelFunctions.md)

### 发布说明

* **2026.09.11** - v2.4.0
- [x] ✨ **新增:** 导出时 `DateTime` / `DateTime?` 列写成真正的日期单元格（序列号 + 日期格式），Excel 可以排序、筛选、参与计算，`null` 写成空单元格；`Dictionary<string, object>` 与 `DataTable` 导出中的 `DateTime` 值同样处理 - 查看 [DateTimeFormats.md](DateTimeFormats.md)
- [x] ✨ **新增:** `[ExcelColumn(Format = ...)]` 的 .NET 日期格式串自动翻译成对应的 Excel 数字格式（`yyyy-MM-dd HH:mm:ss` → `yyyy-mm-dd hh:mm:ss`，`yyyy年MM月dd日` → `yyyy"年"mm"月"dd"日"`），未指定时为 `yyyy-mm-dd hh:mm:ss`；已经是 Excel 拼写的格式（`m/d/yy`、`[$-409]d-mmm-yy`、`yyyy/m/d;@`）原样使用
- [x] ✨ **新增:** `ExcelExporterOptions.DateTimeAsText`，恢复 v2.3 及之前把日期写成文本的导出方式；`DataTable` 导出新增接受 options 的重载 `ObjectToExcelBytes(DataTable, Action<ExcelExporterOptions>)`
- [x] 🐛 修复公式列接管带自定义 `Format` 的 `DateTime` 列时公式被丢弃、改写成模型值文本的问题：现在保留公式并套用日期格式（公式若返回数字而非日期，用 `FormulaResultType` 声明）
- [x] 🐛 修复日期单元格读进 `string` 属性或 `Dictionary<string, object>` 时得到 `46276.6` 这类序列号的问题：现按单元格格式渲染成 `yyyy-MM-dd`、`HH:mm:ss` 或两者，`[h]:mm` 这类经过时间仍按数字读；数值属性（`double`、`decimal` 等）始终取原始数字
- [x] ⚠️ **行为变更:** Excel 日历不覆盖 1900 年之前（1904 日期系统的工作簿为 1904 年之前），这类日期（尤其 `default(DateTime)`）仍按 `Format` 渲染成文本写入；`hh` 不带 `tt` 时显示 24 小时制，Excel 没有不带 AM/PM 标记的 12 小时制

* **2026.09.11** - v2.3.0
- [x] ✨ **新增:** 公式内置函数库从 34 个扩充到 **334 个** Excel 函数，分 10 类：数学与三角（73）、统计（80）、逻辑（11）、查找与引用（33）、日期时间（25）、文本（39）、信息（20）、财务（21）、工程（20）、数据库（12）。信息 / 财务 / 工程 / 数据库为全新分类，入口为 `ExcelFunctions.Information`、`.Financial`、`.Engineering`、`.Database` - 查看 [ExcelFunctions.md](ExcelFunctions.md)
- [x] 🐛 修复 Excel 2007 之后新增的函数写入后打开显示 `#NAME?` 的问题：xlsx 格式要求这类函数存成 `_xlfn.IFS`、`_xlfn.STDEV.P`（`SORT`/`FILTER` 为 `_xlfn._xlws.`），现已自动加前缀；此前已有的 `DAYS` 即受此影响
- [x] 🐛 修复公式中的数字按当前区域性格式化的问题：`de-DE` 等区域下 `1.5` 会写成 `1,5` 导致公式无法解析，现统一使用不变区域性
- [x] 🐛 修复文本字面量中的双引号未按 Excel 规则转义（`"` 需写成 `""`）的问题
- [x] ✨ **新增:** 公式中可直接书写更多普通 C#：`%`（MOD）、`^`（乘幂）、`&&` / `||` / `!`（AND / OR / NOT）、三元表达式（IF）、`Math.*`、`string` 成员（`ToUpper`、`Substring`、`Length`、`Contains` 等；语义与 Excel 不完全一致的 `Math.Round`、`string.Trim` 等不翻译，需显式调用对应的 `ExcelFunctions` 函数）、`DateTime` 成员与 `AddDays` / `AddMonths` / `AddYears`、可空属性的 `.HasValue`
- [x] ✨ **新增:** 公式中可引用 lambda 捕获的局部变量与 `new DateTime(...)` 等常量表达式，自动求值后写成字面量（此前会生成无效公式）
- [x] ✨ **改进:** 区间（`c.Matrix(...)`、`c.Columns(...)`）可隐式用于接受单值的参数位置，`SUM` 等函数可混合传入单元格与区间
- [x] 🐛 修复 `[ExcelColumn]` 的单元格样式被静默忽略的问题：样式缓存键与分支判断不匹配，`CellBold` / `CellFontColor` / `CellAlignment` 等此前完全没有写进单元格；另修复一列声明 `CellAlignment` 会改到工作簿默认样式、导致其他列跟着变的外溢问题（#47）
- [x] 🐛 修复 `FormulaResultType = typeof(DateTime)` 的公式列没有日期格式、在 Excel 里显示成 `46282.6` 这类序列号的问题（#47）
- [x] ⚠️ **不兼容变更:** `IStatisticsFunction.Sum` 由 `params ColumnMatrix[]` 改为 `params ColumnValue[]`，`IMathFunction` 多个方法的返回值由 `int` / `double` 统一为 `ColumnValue`。源码级兼容（现有公式无需修改即可编译），但二进制不兼容，升级后需重新编译

* **2026.09.11** - v2.2.1
- [x] 🔧 修复命令行工具包体超过 NuGet 250 MB 上限导致 2.2.0 未能发布的问题：CLI 仅保留 net8.0 并剔除原生 `.pdb`（库代码无变更）

* **2026.09.11** - v2.2.0
- [x] ✨ **新增:** 命令行工具 `excel2obj`（`dotnet tool install -g Chsword.Excel2Object.Cli`）：Excel ↔ JSON 转换、批量转换、由表头生成带 `[ExcelTitle]` 的模型类 - 查看 [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)（#9）
- [x] ✨ **新增:** 公式列可通过模型属性引用列：`options.FormulaColumns.Add<Order>("Total", (c, m) => m.Price * m.Qty)`，编译期检查属性名，引用未导出的属性会抛出异常 - 查看 [ExcelFunctions.md](ExcelFunctions.md)（#22）
- [x] ✨ **改进:** 导出时数值列写成数字单元格、bool 列写成布尔单元格、null/DBNull 写成空白单元格（此前一律写成文本）；字符串列仍为文本，前导零不会丢失
- [x] 🐛 修复公式中运算符优先级导致缺少括号的问题，如 `c["A"] * (c["B"] + c["C"])` 此前生成 `A2*B2+C2`
- [x] ✨ **改进:** 公式中引用当前 sheet 不存在的列标题时抛出 `Excel2ObjectException`，不再静默回退

* **2026.09.08** - v2.1.0
- [x] ✨ **新增:** 公式列支持跨 sheet 引用：`c.Sheet("Products")["Price", 2]`、`.Matrix(...)`、`.Columns(...)`，可直接用于 `VLOOKUP` 等函数；另新增当前 sheet 的整列引用 `c.Columns("A列", "D列")` - 查看 [ExcelFunctions.md](ExcelFunctions.md)
- [x] ✨ **改进:** 公式中引用不存在的列或 sheet 时抛出 `Excel2ObjectException`（附公式列名），不再生成无效公式
- [x] ✨ **更新:** NPOI 到 2.8.0（NPOI 2.8 起附带 OSMF EULA，本库已在 csproj 中声明接受；源代码仍为 Apache-2.0）
- [x] 🔒 修复传递依赖 System.Security.Cryptography.Xml 的安全漏洞（钉到 8.0.4）
- [x] 🧹 移除未使用的 SixLabors.ImageSharp 依赖；补全 `FormulaColumnsCollection.CopyTo`；`ColumnValue.Equals/GetHashCode` 不再抛异常
- [x] 🔧 CI 与测试迁移到 .NET 10；修复发布流程中 GitHub Release 创建失败的问题

* **2025.10.11** - v2.0.5
- [x] 🔧 `release.ps1` 支持自动递增补丁版本号与 `-Help`（库代码无变更）

* **2025.10.11** - v2.0.4
- [x] ℹ️ 自 2.0.0.211 之后首个发布到 NuGet 的 2.0.x 版本（v2.0.1–v2.0.3 未曾发布），包含以下全部改动：
- [x] ✨ **新增:** 基于内容的自动列宽调整
  - 自动计算最优列宽
  - 支持最小和最大宽度限制
  - 正确处理中文/Unicode 字符
  - 可通过 `ExcelExporterOptions` 配置
- [x] ✨ **新增:** 全面的日期/时间格式支持（56 种格式）- 查看 [DateTimeFormats.md](DateTimeFormats.md)
  - ISO 8601 格式（支持时区和毫秒）
  - 多种日期分隔符（横线、斜线、点号）
  - 12 小时制和 24 小时制时间格式
  - 仅时间格式
  - 区域格式支持（美国、欧洲等）
  - 向后兼容现有的中文日期格式（年月日）
- [x] ✨ **更新:** NPOI 到 2.7.5
- [x] ✨ **更新:** SixLabors.ImageSharp 到 3.1.11（修复安全漏洞）
- [x] 🔧 GitHub Actions CI 取代 AppVeyor；新增 tag 触发的自动发布流程与 `release.ps1` 一键发布脚本

* **2024.10.21**
- [x] 更新 SixLabors.ImageSharp 到 2.1.9
- [x] 测试 .NET 8.0

* **2024.05.10**
- [x] 支持 .NET 8.0 / .NET 6.0 / .NET Standard 2.1 / .NET Standard 2.0 / .NET Framework 4.7.2
- [x] 清理已弃用的库

* **2023.11.02**
- [x] 支持列标题映射 - [Issue39DynamicMappingTitle.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue39DynamicMappingTitle.cs)

* **2023.07.31**
- [x] 支持 DateTime 和 Nullable<DateTime> 格式，如 `[ExcelColumn("Title",Format="yyyy-MM-dd HH:mm:ss")]`

* **2023.03.26**
- [x] 支持列标题中的特殊符号 #37 - [Issue37SpecialCharTest.cs](https://github.com/chsword/Excel2Object/commit/273122275e724367bb6154e03df61702fcec81b3#diff-5f0f5f7558bf7d4207cfa752a4506c4df89d9b491e2501e4862aff0c2276bd61)

* **2023.02.20**
- [x] 支持平台：.NET Standard 2.0/2.1、.NET 6.0、.NET Framework 4.7.2

* **2022.03.19**
- [x] 支持 ExcelImporterOptions，跳过行 - [Issue32SkipLineImport.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue32SkipLineImport.cs)
- [x] 修复超类属性 bug - [Issue31SuperClass.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue31SuperClass.cs)

* **2021.11.4**
- [x] 多 sheet 支持 - [Pr28MultipleSheetTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr28MultipleSheetTest.cs)

* **2021.10.23**
- [x] 修复 Nullable DateTime bug @SunBrook

* **2021.10.22**
- [x] 支持 Nullable 类型 - [Pr24NullableTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr24NullableTest.cs) @SunBrook

* **2021.5.28**
- [x] 支持表头和单元格样式，新增列的 [ExcelColumnAttribute] 
- [x] 支持公式 - [ExcelFunctions.md](./ExcelFunctions.md)

```C#
var list = new List<Pr20Model>
{
        new Pr20Model
        {
            Fullname = "AAA", Mobile = "123456798123"
        },
        new Pr20Model
        {
            Fullname = "BBB", Mobile = "234"
        }
};
var bytes = ExcelHelper.ObjectToExcelBytes(list, ExcelType.Xlsx);
// model
[ExcelTitle("SheetX")]
public class Pr20Model
{
    [ExcelColumn("Full name", CellFontColor = ExcelStyleColor.Red)]
    public string Fullname { get; set; }

    [ExcelColumn("Phone Number",
        HeaderFontFamily = "Normal",
        HeaderBold = true,
        HeaderFontHeight = 30,
        HeaderItalic = true,
        HeaderFontColor = ExcelStyleColor.Blue,
        HeaderUnderline = true,
        HeaderAlignment = HorizontalAlignment.Right,
        //cell
        CellAlignment = HorizontalAlignment.Justify
    )]
    public string Mobile { get; set; }
}
```

* **v2.0.0.113**
```
convert project to netstandard2.0 and .net452
fixbug #12 #13
```

* **v1.0.0.80**

- [x] support simple formula
- [x] support standard excel model
  - [x] excel & JSON convert
  - [x] excel & Dictionary<string,object> convert

```
Support Uri to a hyperlink cell
And also support text cell to Uri Type
```

* **v1.0.0.43**
```
Support xlsx [thanks Soar360]
Support complex Boolean type
```

* **v1.0.0.36**
```
Add ExcelToObject<T>(bytes)
```


## 示例代码

### 定义模型

``` csharp
public class ReportModel
{
    [Excel("My Title", Order=1)]
    public string Title { get; set; }
    
    [Excel("User Name", Order=2)]
    public string Name { get; set; }
}
```

### 创建模型列表

``` csharp
var models = new List<ReportModel>
{
    new ReportModel{Name="a", Title="b"},
    new ReportModel{Name="c", Title="d"},
    new ReportModel{Name="f", Title="e"}
};
```

### 对象转 Excel 文件

``` csharp
var exporter = new ExcelExporter();
var bytes = exporter.ObjectToExcelBytes(models);
File.WriteAllBytes("C:\\demo.xls", bytes);
```

### Excel 文件转对象

``` csharp
var importer = new ExcelImporter();
IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>("c:\\demo.xls");

// 也可以直接使用字节数组
// IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>(bytes);
```

### 自动列宽（新特性）

``` csharp
// 启用自动列宽调整
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.ExcelType = ExcelType.Xlsx;
    options.AutoColumnWidth = true;        // 启用自动宽度
    options.MinColumnWidth = 8;            // 最小宽度（字符）
    options.MaxColumnWidth = 50;           // 最大宽度（字符）
    options.DefaultColumnWidth = 16;       // 禁用自动时的默认宽度
});
```

### 冻结首行与自动筛选

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.ExcelType = ExcelType.Xlsx;
    options.FreezeHeader = true;           // 滚动时表头始终可见
    options.AutoFilter = true;             // 表头带上筛选下拉
});
```

两个选项默认关闭，`.xls` 与 `.xlsx` 都支持。筛选范围覆盖表头及本次写入的数据行；`AppendObjectToExcelBytes` 追加的 sheet 各自独立，不影响已有 sheet。

### 数据验证（下拉列表）

``` csharp
public class OrderModel
{
    [ExcelTitle("订单号")] public string No { get; set; }

    // 固定取值写在特性上
    [ExcelColumn("状态", Dropdown = new[] {"启用", "停用", "待审"})]
    public string Status { get; set; }

    [ExcelTitle("城市")] public string City { get; set; }
}

var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    // 运行时才知道的取值走 options，优先于特性
    options.Dropdowns["城市"] = cities.Select(c => c.Name).ToArray();
});
```

Excel 会把取值显示为下拉，并拒绝其他输入。列表总长超过 255 字符、或取值里含逗号/引号时（Excel 行内列表放不下），自动改写到一张隐藏 sheet 上再由下拉引用，用法不变；`.xls` 与 `.xlsx` 都支持。导出空列表时下拉仍会挂在第一行，方便做填写模板。

### 在 ASP.NET MVC 中使用

在 ASP.NET MVC 模型中，`DisplayAttribute` 可以像 `ExcelTitleAttribute` 一样被支持。

## 文档

- [ExcelFunctions.md](ExcelFunctions.md) - 公式列：内置函数、按列标题/模型属性引用、跨 sheet 引用
- [DateTimeFormats.md](DateTimeFormats.md) - 支持的日期时间格式与解析规则
- [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md) - 命令行工具 `excel2obj` 用法
- [ROADMAP.md](ROADMAP.md) - 路线图

更多信息请访问：http://www.cnblogs.com/chsword/p/excel2object.html

## 开发和发布

- **[自动化发布脚本 `release.ps1`](release.ps1)** - 一键完成版本更新、提交和 Tag 创建的 PowerShell 脚本
- **[自动化发布系统说明](RELEASE_AUTOMATION.md)** - 包含 Copilot 指令、版本管理和自动发布流程的完整说明
- **[版本管理规范](.github/VERSIONING.md)** - 语义化版本规范和版本号递增规则
- **[发布流程指南](.github/RELEASE_GUIDE.md)** - 详细的发布步骤和故障排查指南
- **[Copilot 使用说明](.github/copilot-instructions.md)** - GitHub Copilot 开发指导和项目规范

### 快速发布新版本

使用自动化脚本一键发布：

```powershell
# Windows：不带参数则自动递增修订号，也可用 -Version 2.3.0 指定
.\release.ps1

# Linux/macOS (需要 PowerShell Core)
pwsh ./release.ps1
```

详细说明请参阅 [RELEASE_AUTOMATION.md](RELEASE_AUTOMATION.md)。

## 贡献者

[![Contributors](https://contrib.rocks/image?repo=chsword/Excel2Object)](https://github.com/chsword/Excel2Object/graphs/contributors)

## 参考

- https://github.com/tonyqus/npoi
- https://github.com/chsword/ctrc

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。
