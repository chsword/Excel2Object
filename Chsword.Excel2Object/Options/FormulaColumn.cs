using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object.Options;

public class FormulaColumn
{
    public string? AfterColumnTitle { get; set; }
    public Expression<Func<ColumnCellDictionary, object>>? Formula { get; set; }

    public Type? FormulaResultType { get; set; }
    public string? Title { get; set; }
}