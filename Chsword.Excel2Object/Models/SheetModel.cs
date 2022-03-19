using System.Collections.Generic;

namespace Chsword.Excel2Object
{
    internal class SheetModel
    {
        public List<ExcelColumn> Columns { get; set; }
        public int Index { get; set; }
        public List<Dictionary<string, object>> Rows { get; set; }
        public string Title { get; set; }

        public static SheetModel Create(string title)
        {
            return new SheetModel()
            {
                Title = title,
                Columns = new List<ExcelColumn>(),
                Rows = new List<Dictionary<string, object>>()
            };
        }
    }
}