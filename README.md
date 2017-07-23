# Excel2Object

[![install from nuget](http://img.shields.io/nuget/v/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)
[![release](https://img.shields.io/github/release/chsword/Excel2Object.svg?style=flat-square)](https://github.com/chsword/Excel2Object/releases)
[![Build status](https://ci.appveyor.com/api/projects/status/4po2h27j7yg4bph5/branch/master?svg=true)](https://ci.appveyor.com/project/chsword/excel2object/branch/master)
[![CodeFactor](https://www.codefactor.io/repository/github/chsword/excel2object/badge)](https://www.codefactor.io/repository/github/chsword/excel2object)

Excel convert to .NET Object / .NET Object convert to Excel.

- [Top](#excel2object)
    - [NuGet install](#nuget-install)
    - [Release notes](#release-notes)
    - [Demo code](#demo-code)
    - [Document](#document)
    - [Reference](#reference)
          
### NuGet Install
``` powershell
PM> Install-Package Chsword.Excel2Object
```

### Release Notes and roadmap

- [ ] 1. support auto width column
- [ ] 1. support date datetime time in excel
- [x] 1. support Uri type to a Hyperlink
- [x] 1. support xls / xlsx
- [x] 1. support complex Boolean type
- [x] 1. support convert List<Model> to excel file and Excel file to List<Model>
- [x] 1. support to specify the order of fields to export
- [x] 1. support convert List<Model> to excel bytes and Excel bytes to List<Model>

* v1.0.0.80
```
Support Uri to a hyperlink cell
And also support text cell to Uri Type
```

* v1.0.0.43
```
Support xlsx [thanks Soar360]
Support complex Boolean type
```

* v1.0.0.36
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
``` cs
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
