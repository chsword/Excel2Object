using Chsword.Excel2Object.Functions;
using Chsword.Excel2Object.Styles;
using System.Linq.Expressions;

namespace Chsword.Excel2Object;

internal class ExcelColumn
{
    public IExcelCellStyle? CellStyle { get; set; }

    public Expression<Func<ColumnCellDictionary, object>>? Formula { get; set; }

    public IExcelHeaderStyle? HeaderStyle { get; set; }
    public int Order { get; set; }

    /// <summary>
    ///     ���ҽ��� Type = Expression ʱ��Ч
    /// </summary>
    public Type? ResultType { get; set; }

    public string? Title { get; set; }
    public Type? Type { get; set; }
}