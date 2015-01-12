Excel2Object
============
[![install from nuget](http://img.shields.io/nuget/v/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)
[![downloads](http://img.shields.io/nuget/dt/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)
[![release](https://img.shields.io/github/release/chsword/Chsword.Excel2Object.svg?style=flat-square)](https://github.com/chsword/Chsword.Excel2Object/releases)

Excel convert to .NET Object


### NuGet Install
```powershell
PM> Install-Package Chsword.Excel2Object
```
### Demo Code
Model
``` csharp
    public class ReportModel
    {
        [Excel("标题",Order=1)]
        public string Title { get; set; }
        [Excel("用户",Order=2)]
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
```
With ASP.NET MVC
      In ASP.NET MVC Model, DisplayAttribute can be supported like ExcelAttribute.

### Document
http://www.cnblogs.com/chsword/p/excel2object.html

### Reference
NPOI
https://github.com/tonyqus/npoi

