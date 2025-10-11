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

## 发布说明和路线图

### 暂不支持的特性

- [ ] CLI 工具
- [x] 支持自动列宽 ✅ **v2.0.1 新增**
- [x] 支持 Excel 日期/日期时间/时间格式 ✅ **v2.0.2 新增** - 查看 [DateTimeFormats.md](DateTimeFormats.md)

### 发布说明

* **2025.10.11** - v2.0.2
- [x] ✨ **新增:** 全面的日期/时间格式支持（56 种格式）- 查看 [DateTimeFormats.md](DateTimeFormats.md)
  - ISO 8601 格式（支持时区和毫秒）
  - 多种日期分隔符（横线、斜线、点号）
  - 12 小时制和 24 小时制时间格式
  - 仅时间格式
  - 区域格式支持（美国、欧洲等）
  - 向后兼容现有的中文日期格式（年月日）
- [x] ✨ **更新:** NPOI 到 2.7.5
- [x] ✨ **更新:** SixLabors.ImageSharp 到 3.1.11（修复安全漏洞）

* **2025.07.23** - v2.0.1
- [x] ✨ **新增:** 基于内容的自动列宽调整
  - 自动计算最优列宽
  - 支持最小和最大宽度限制
  - 正确处理中文/Unicode 字符
  - 可通过 `ExcelExporterOptions` 配置

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

### 在 ASP.NET MVC 中使用

在 ASP.NET MVC 模型中，`DisplayAttribute` 可以像 `ExcelTitleAttribute` 一样被支持。

## 文档

更多信息请访问：http://www.cnblogs.com/chsword/p/excel2object.html

## 开发和发布

- **[自动化发布系统说明](RELEASE_AUTOMATION.md)** - 包含 Copilot 指令、版本管理和自动发布流程的完整说明
- **[版本管理规范](.github/VERSIONING.md)** - 语义化版本规范和版本号递增规则
- **[发布流程指南](.github/RELEASE_GUIDE.md)** - 详细的发布步骤和故障排查指南
- **[Copilot 使用说明](.github/copilot-instructions.md)** - GitHub Copilot 开发指导和项目规范

## 贡献者

[![Contributors](https://contrib.rocks/image?repo=chsword/Excel2Object)](https://github.com/chsword/Excel2Object/graphs/contributors)

## 参考

- https://github.com/tonyqus/npoi
- https://github.com/chsword/ctrc

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。
