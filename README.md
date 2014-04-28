Excel2Object
============

Excel 与 Object 互相转换


Demo Code
前提
准备一个Model

    public class ReportModel
    {
        [Excel("标题",Order=1)]
        public string Title { get; set; }
        [Excel("用户",Order=2)]
        public string Name { get; set; }
    }

准备一个List

      var models = new List<ReportModel>
            {
                new ReportModel{Name="a",Title="b"},
                new ReportModel{Name="c",Title="d"},
                new ReportModel{Name="f",Title="e"}
            };

由Object转为Excel

      var exporter = new ExcelExporter();
      var bytes = exporter.ObjectToExcelBytes(models);
      File.WriteAllBytes("C:\\demo.xls", bytes);

由Excel转为Object

      var importer = new ExcelImporter();
      IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>("c:\\demo.xls");
      
与ASP.NET MVC结合使用
      由于ASP.NET MVC中Model上会使用DisplayAttribute所以Excel2Object除了支持ExcelAttribute外，也支持DisplayAttribute。
            
.NET 项目中使用
  使用NuGet安装即可，命令行安装
    
    Install-Package Chsword.Excel2Object
    
  或搜索包
    
    Chsword.Excel2Object

博客说明

    [http://www.cnblogs.com/chsword/p/excel2object.html]
    
    
    
