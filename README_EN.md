# Excel2Object

[![install from nuget](http://img.shields.io/nuget/v/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)
[![release](https://img.shields.io/github/release/chsword/Excel2Object.svg?style=flat-square)](https://github.com/chsword/Excel2Object/releases)
[![Build status](https://ci.appveyor.com/api/projects/status/4po2h27j7yg4bph5/branch/master?svg=true)](https://ci.appveyor.com/project/chsword/excel2object)
[![CodeFactor](https://www.codefactor.io/repository/github/chsword/excel2object/badge)](https://www.codefactor.io/repository/github/chsword/excel2object)

Excel convert to .NET Object / .NET Object convert to Excel.

- [Top](#excel2object)
    - [NuGet install](#nuget-install)
    - [Release notes and roadmap](#release-notes-and-roadmap)
    - [Demo code](#demo-code)
    - [Document](#document)
    - [Reference](#reference)

## Platform Support

[![.NET 4.7.2 +](https://img.shields.io/badge/-4.7.2%2B-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.0](https://img.shields.io/badge/-standard2.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.1](https://img.shields.io/badge/-standard2.1-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 6.0](https://img.shields.io/badge/-6.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 8.0](https://img.shields.io/badge/-8.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)

## NuGet Install

``` powershell
PM> Install-Package Chsword.Excel2Object
```

Or using .NET CLI:
``` bash
dotnet add package Chsword.Excel2Object
```

## Release Notes and Roadmap

### Features Not Yet Supported

- [ ] CLI tool
- [x] Support auto width column ✅ **New in v2.0.1**
- [ ] Support date/datetime/time formats in Excel

### Release Notes

* **2025.01.XX** - v2.0.2
- [x] ✨ **Updated:** NPOI to 2.7.5
- [x] ✨ **Updated:** SixLabors.ImageSharp to 3.1.11 (Fixed security vulnerability)

* **2025.07.23** - v2.0.1
- [x] ✨ **NEW:** Auto column width adjustment based on content
  - Automatically calculates optimal column widths
  - Supports minimum and maximum width constraints
  - Handles Chinese/Unicode characters properly
  - Configurable through `ExcelExporterOptions`

* **2024.10.21**
- [x] Updated SixLabors.ImageSharp to 2.1.9
- [x] Tested for .NET 8.0

* **2024.05.10**
- [x] Support .NET 8.0 / .NET 6.0 / .NET Standard 2.1 / .NET Standard 2.0 / .NET Framework 4.7.2
- [x] Cleared deprecated libraries

* **2023.11.02**
- [x] Support column title mapping - [Issue39DynamicMappingTitle.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue39DynamicMappingTitle.cs)

* **2023.07.31**
- [x] Support DateTime and Nullable<DateTime> format, such as `[ExcelColumn("Title",Format="yyyy-MM-dd HH:mm:ss")]`

* **2023.03.26**
- [x] Support special symbols in column titles #37 - [Issue37SpecialCharTest.cs](https://github.com/chsword/Excel2Object/commit/273122275e724367bb6154e03df61702fcec81b3#diff-5f0f5f7558bf7d4207cfa752a4506c4df89d9b491e2501e4862aff0c2276bd61)

* **2023.02.20**
- [x] Support platforms: .NET Standard 2.0/2.1, .NET 6.0, .NET Framework 4.7.2

* **2022.03.19**
- [x] Support ExcelImporterOptions, Skipline - [Issue32SkipLineImport.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue32SkipLineImport.cs)
- [x] Fixed superclass property bug - [Issue31SuperClass.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue31SuperClass.cs)

* **2021.11.4**
- [x] Multiple sheet support - [Pr28MultipleSheetTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr28MultipleSheetTest.cs)

* **2021.10.23**
- [x] Fixed Nullable DateTime bug @SunBrook

* **2021.10.22**
- [x] Support Nullable types - [Pr24NullableTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr24NullableTest.cs) @SunBrook

* **2021.5.28**
- [x] Support styling for headers & cells, new [ExcelColumnAttribute] for columns
- [x] Support Functions - [ExcelFunctions.md](./ExcelFunctions.md)

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

// Model definition
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
Converted project to .NET Standard 2.0 and .NET Framework 4.5.2
Fixed bugs #12 #13
```

* **v1.0.0.80**
- [x] Support simple formulas
- [x] Support standard Excel model
  - [x] Excel & JSON conversion
  - [x] Excel & Dictionary<string,object> conversion

```
Support Uri to hyperlink cell
Also support text cell to Uri type
```

* **v1.0.0.43**
```
Support xlsx format [thanks Soar360]
Support complex Boolean type
```

* **v1.0.0.36**
```
Add ExcelToObject<T>(bytes)
```

## Demo Code

### Define Model

``` csharp
public class ReportModel
{
    [Excel("My Title", Order=1)]
    public string Title { get; set; }
    
    [Excel("User Name", Order=2)]
    public string Name { get; set; }
}
```

### Create Model List

``` csharp
var models = new List<ReportModel>
{
    new ReportModel{Name="a", Title="b"},
    new ReportModel{Name="c", Title="d"},
    new ReportModel{Name="f", Title="e"}
};
```

### Convert Object to Excel File

``` csharp
var exporter = new ExcelExporter();
var bytes = exporter.ObjectToExcelBytes(models);
File.WriteAllBytes("C:\\demo.xls", bytes);
```

### Convert Excel File to Object

``` csharp
var importer = new ExcelImporter();
IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>("c:\\demo.xls");

// You can also use bytes directly
// IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>(bytes);
```

### Auto Column Width (New Feature)

``` csharp
// Enable auto column width adjustment
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.ExcelType = ExcelType.Xlsx;
    options.AutoColumnWidth = true;        // Enable auto width
    options.MinColumnWidth = 8;            // Minimum width in characters
    options.MaxColumnWidth = 50;           // Maximum width in characters
    options.DefaultColumnWidth = 16;       // Default width when auto is disabled
});
```

### Use with ASP.NET MVC

In ASP.NET MVC models, the `DisplayAttribute` can be supported like `ExcelTitleAttribute`.

## Document

For more information, please visit: http://www.cnblogs.com/chsword/p/excel2object.html

## Contributors

[![Contributors](https://contrib.rocks/image?repo=chsword/Excel2Object)](https://github.com/chsword/Excel2Object/graphs/contributors)

## Reference

- https://github.com/tonyqus/npoi
- https://github.com/chsword/ctrc

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
