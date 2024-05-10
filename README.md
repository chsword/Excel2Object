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

Platform

[![.NET 4.7.2 +](https://img.shields.io/badge/-4.7.2%2B-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.0](https://img.shields.io/badge/-standard2.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.1](https://img.shields.io/badge/-standard2.1-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 6.0](https://img.shields.io/badge/-6.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 8.0](https://img.shields.io/badge/-8.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)

### NuGet Install
``` powershell
PM> Install-Package Chsword.Excel2Object
```

### Release Notes and roadmap

#### Features not supported

- [ ] cli tool
- [ ] support auto width column
- [ ] 1. support date datetime time in excel\

#### Release Notes

* **2024.05.10**
- [x] support .net8.0 / .net6.0 / .netstandard2.1 / .netstandard2.0 / .net4.7.2
- [x] clear deprecated library
* **2023.11.02**
- [x] support column title mapping Issue39DynamicMappingTitle.cs
* **2023.07.31**
- [x] support DateTime and Nullable<DateTime> format ,such as `[ExcelColumn("Title",Format="yyyy-MM-dd HH:mm:ss")]`
* **2023.03.26**
- [x] support special symbol in columns title #37 [Issue37SpecialCharTest.cs](https://github.com/chsword/Excel2Object/commit/273122275e724367bb6154e03df61702fcec81b3#diff-5f0f5f7558bf7d4207cfa752a4506c4df89d9b491e2501e4862aff0c2276bd61)
* **2023.02.20**
- [x] support platform netstandard2.0/netstandard2.1/.net6.0/.netframework4.7.2
* **2022.03.19**
- [x] support ExcelImporterOptions , Skipline [Issue32SkipLineImport.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue32SkipLineImport.cs)
- [x] fixed super class prop bug [Issue31SuperClass.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue31SuperClass.cs)
* **2021.11.4**
- [x] multiple sheet , demo & test file : [Pr28MultipleSheetTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr28MultipleSheetTest.cs)
* **2021.10.23**
- [x] Nullable DateTime bugfixed @SunBrook 
* **2021.10.22**
- [x] support Nullable, test file :[Pr24NullableTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr24NullableTest.cs) @SunBrook 
* **2021.5.28**
- [x] support style for header & cell, new [ExcelColumnAttribute] for column.
- [x] support Functions [./ExcelFunctions.md](./ExcelFunctions.md)

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

### Demo Code

Model
``` csharp
    public class ReportModel
    {
        [Excel("My Title",Order=1)]
        public string Title { get; set; }
        [Excel("User Name",Order=2)]
        public string Name { get; set; }
    }
```

Model List
``` csharp
      var models = new List<ReportModel>
            {
                new ReportModel{Name="a",Title="b"},
                new ReportModel{Name="c",Title="d"},
                new ReportModel{Name="f",Title="e"}
            };
```

Convert Object to Excel file.
``` csharp
      var exporter = new ExcelExporter();
      var bytes = exporter.ObjectToExcelBytes(models);
      File.WriteAllBytes("C:\\demo.xls", bytes);
```

Convert Excel file to Object
``` csharp
      var importer = new ExcelImporter();
      IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>("c:\\demo.xls");
      // also can use bytes
      //IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>(bytes);
```

With ASP.NET MVC
      In ASP.NET MVC Model, DisplayAttribute can be supported like ExcelTitleAttribute.

### Document

http://www.cnblogs.com/chsword/p/excel2object.html

### Reference

https://github.com/tonyqus/npoi

https://github.com/chsword/ctrc
