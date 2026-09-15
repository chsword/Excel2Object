using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object;

internal class ExcelColumn
{
    /// <summary>
    ///     标题与类型经构造函数给出，而非以 <c>null!</c> 作默认值：后者仍允许造出违反其契约的对象，
    ///     漏填时会一路静默传递下去，正是可空检查本该拦住的那类问题。
    /// </summary>
    public ExcelColumn(string title, Type type)
    {
        Title = title;
        Type = type;
    }

    public IExcelCellStyle? CellStyle { get; set; }

    public LambdaExpression? Formula { get; set; }

    public IExcelHeaderStyle? HeaderStyle { get; set; }
    public int Order { get; set; }

    /// <summary>仅当 <see cref="Type" /> 为 Expression 时有效：公式算出的值属于哪种类型。</summary>
    public Type? ResultType { get; set; }

    /// <summary>列标题，构造时即须给出：它是表头上的名字，也是各处按列定位的依据。</summary>
    public string Title { get; set; }

    /// <summary>该列的值的类型；公式列为 Expression。</summary>
    public Type Type { get; set; }

    /// <summary>The values the column's cells are limited to, shown as a dropdown.</summary>
    public string[]? Dropdown { get; set; }
}