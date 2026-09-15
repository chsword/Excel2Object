using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object;

internal class ExcelColumn
{
    public IExcelCellStyle? CellStyle { get; set; }

    public LambdaExpression? Formula { get; set; }

    public IExcelHeaderStyle? HeaderStyle { get; set; }
    public int Order { get; set; }

    /// <summary>仅当 <see cref="Type" /> 为 Expression 时有效：公式算出的值属于哪种类型。</summary>
    public Type? ResultType { get; set; }

    /// <summary>列标题。各处构造时均会给出，公式列的标题由 FormulaColumnsCollection 在加入时校验。</summary>
    public string Title { get; set; } = null!;

    /// <summary>该列的值的类型；公式列为 Expression。</summary>
    public Type Type { get; set; } = null!;

    /// <summary>The values the column's cells are limited to, shown as a dropdown.</summary>
    public string[]? Dropdown { get; set; }
}