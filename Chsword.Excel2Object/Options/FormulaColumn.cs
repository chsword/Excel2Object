using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Chsword.Excel2Object
{
    public class FormulaColumn
    {
        public string Title { get; set; }
        public Expression<Func<Dictionary<string, object>, object>> Formula { get; set; }
        public string AfterColumnTitle { get; set; }

        public Type FormulaResultType { get;set; }
    }
}