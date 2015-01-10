Excel2Object
============
[![install from nuget](http://img.shields.io/nuget/v/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)[![downloads](http://img.shields.io/nuget/dt/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)



Excel 与 Object 互相转换

使用的NPOI
https://github.com/tonyqus/npoi

### NuGet Install
```powershell
PM> Install-Package Chsword.Excel2Object
```
### Demo Code
前提
准备一个Model
``` csharp
    public class ReportModel
    {
        [Excel("标题",Order=1)]
        public string Title { get; set; }
        [Excel("用户",Order=2)]
        public string Name { get; set; }
    }
```
准备一个List
``` cs
      var models = new List<ReportModel>
            {
                new ReportModel{Name="a",Title="b"},
                new ReportModel{Name="c",Title="d"},
                new ReportModel{Name="f",Title="e"}
            };
```
由Object转为Excel
``` csharp
      var exporter = new ExcelExporter();
      var bytes = exporter.ObjectToExcelBytes(models);
      File.WriteAllBytes("C:\\demo.xls", bytes);
```
由Excel转为Object
``` csharp
      var importer = new ExcelImporter();
      IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>("c:\\demo.xls");
```
与ASP.NET MVC结合使用
      由于ASP.NET MVC中Model上会使用DisplayAttribute所以Excel2Object除了支持ExcelAttribute外，也支持DisplayAttribute。

博客说明
http://www.cnblogs.com/chsword/p/excel2object.html
