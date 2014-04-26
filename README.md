Excel2Object
============

Excel 与 Object 互相转换


Demo Code
前提，准备一个List
      List<ReportModel> GetModels()
        {
            return new List<ReportModel>
            {
                new ReportModel{Name="a",Title="b"},
                new ReportModel{Name="c",Title="d"},
                new ReportModel{Name="f",Title="e"}
            };
        }

由Excel转为Object

