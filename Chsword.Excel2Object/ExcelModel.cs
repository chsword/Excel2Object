using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object
{
    internal class ExcelModel
    {
        public List<SheetModel> Sheets { get; set; }
    }

    internal class SheetModel
    {
        public List<ExcelColumn> Columns { get; set; }
        public List<Dictionary<string, ExcelCell>> Rows { get; set; }
        public int Index { get; set; }
        public string Title { get; set; }

        public static SheetModel Create(string title)
        {
            return new SheetModel
            {
                Title = title,
                Columns = new List<ExcelColumn>(),
                Rows = new List<Dictionary<string, ExcelCell>>()
            };
        }
    }

    internal class ExcelCell
    {
        public ExcelCell()
        {
            
        }

        public ExcelCell(object val)
        {
            Value = val;
        }

        public ExcelCell(object val,CellType type):this(val)
        {
            CellType = type;
        }
        public CellType CellType { get; set; } = CellType.Unknown;
        public object Value { get; set; }
    }

    internal class ExcelColumn
    {
        public int Order { get; set; }
        public string Title { get; set; }
        public Type Type { get; set; }

    }
}