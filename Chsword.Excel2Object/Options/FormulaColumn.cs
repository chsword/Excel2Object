using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object.Options;

public class FormulaColumn
{
    public string? AfterColumnTitle { get; set; }

    /// <summary>
    ///     Formula that refers to columns by title, e.g. <c>c => c["Price"] * c["Qty"]</c>.
    /// </summary>
    public Expression<Func<ColumnCellDictionary, object>>? Formula { get; set; }

    /// <summary>
    ///     Formula that also receives the exported model, so columns can be referred to by property,
    ///     e.g. <c>(c, m) => m.Price * m.Qty</c>. Takes precedence over <see cref="Formula" />.
    ///     Usually set through <see cref="FormulaColumnsCollection.Add{TModel}" />.
    /// </summary>
    public LambdaExpression? ModelFormula { get; set; }

    public Type? FormulaResultType { get; set; }
    public string? Title { get; set; }
}