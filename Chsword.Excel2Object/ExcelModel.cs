using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Chsword.Excel2Object
{
    internal class ExcelModel
    {
        public List<SheetModel> Sheets { get; set; }
    }

    internal class SheetModel
    {
        public List<ExcelColumn> Columns { get; set; }
        public List<Dictionary<string, object>> Rows { get; set; }
        public int Index { get; set; }
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
    internal class ExcelColumn
    {
        public int Order { get; set; }
        public string Title { get; set; }
        public Type Type { get; set; }
        /// <summary>
        /// 当且仅当 Type = Expression时有效
        /// </summary>
        public Type ResultType { get;set;}
        public Expression<Func<Dictionary<string, object>, object>> Formula { get; set; }
    }
}