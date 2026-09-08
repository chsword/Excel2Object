using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object.Internal;

internal class ExpressionConvert
{
    private static readonly Type[] CallMethodTypes =
    {
        typeof(IMathFunction),
        typeof(IStatisticsFunction),
        typeof(IConditionFunction),
        typeof(IReferenceFunction),
        typeof(IDateTimeFunction),
        typeof(ITextFunction),
        typeof(IAllFunction)
    };

    public ExpressionConvert(string[] columns, int rowIndex)
        : this(columns, rowIndex, null)
    {
    }

    /// <param name="columns">Column titles of the sheet the formula is written to, in column order.</param>
    /// <param name="rowIndex">0-based row index of the formula cell.</param>
    /// <param name="sheetColumnsResolver">
    ///     Returns the column titles of another sheet in the workbook by sheet title, or null if the
    ///     sheet's layout is unknown. Used to translate <see cref="ColumnCellDictionary.Sheet" /> references.
    /// </param>
    public ExpressionConvert(string[] columns, int rowIndex, Func<string, string[]?>? sheetColumnsResolver)
    {
        Columns = columns;
        RowIndex = rowIndex;
        SheetColumnsResolver = sheetColumnsResolver;
    }

    private static Dictionary<ExpressionType, string> BinarySymbolDictionary { get; } =
        new()
        {
            [ExpressionType.Add] = "+",
            [ExpressionType.Subtract] = "-",
            [ExpressionType.Multiply] = "*",
            [ExpressionType.Divide] = "/",
            [ExpressionType.Equal] = "=",
            [ExpressionType.NotEqual] = "<>",
            [ExpressionType.GreaterThan] = ">",
            [ExpressionType.LessThan] = "<",
            [ExpressionType.GreaterThanOrEqual] = ">=",
            [ExpressionType.LessThanOrEqual] = "<=",
            [ExpressionType.And] = "&"
        };

    private string[] Columns { get; }
    private int RowIndex { get; }
    private Func<string, string[]?>? SheetColumnsResolver { get; }

    public string Convert(Expression? expression)
    {
        if (expression == null) return string.Empty;
        return expression.NodeType == ExpressionType.Lambda
            ? InternalConvert((expression as LambdaExpression)?.Body)
            : string.Empty;
    }

    private static string ConvertConstant(Expression expression)
    {
        var exp = expression as ConstantExpression;
        return (exp?.Type == typeof(bool) ? exp.ToString().ToUpper() : exp?.ToString()) ?? string.Empty;
    }

    private string ConvertBinaryExpression(Expression expression)
    {
        if (!(expression is BinaryExpression binary)) return "null";
        var symbol = $"unsupported binary symbol:{binary.NodeType}";
        if (BinarySymbolDictionary.TryGetValue(binary.NodeType, out var value)) symbol = value;

        return $"{InternalConvert(binary.Left)}{symbol}{InternalConvert(binary.Right)}";
    }

    private string ConvertCall(Expression expression)
    {
        if (!(expression is MethodCallExpression exp) || exp.Object == null) return "null";
        if (exp.Method.Name == "get_Item" &&
            (exp.Object.Type == typeof(ColumnCellDictionary)
             || exp.Object.Type == typeof(Dictionary<string, ColumnValue>)
            )
           )
            return exp.Arguments.Count == 2
                ? $"{GetColumn(exp.Arguments[0])}{InternalConvert(exp.Arguments[1])}"
                : $"{GetColumn(exp.Arguments[0])}{RowIndex + 1}";

        if (exp.Object.Type == typeof(ColumnCellDictionary) &&
            exp.Method.Name == nameof(ColumnCellDictionary.Matrix))
            return
                $"{GetColumn(exp.Arguments[0])}{exp.Arguments[1]}:{GetColumn(exp.Arguments[2])}{exp.Arguments[3]}";

        if (exp.Object.Type == typeof(ColumnCellDictionary) &&
            exp.Method.Name == nameof(ColumnCellDictionary.Columns))
            return $"{GetColumn(exp.Arguments[0])}:{GetColumn(exp.Arguments[1])}";

        if (exp.Object.Type == typeof(SheetCellDictionary))
            return ConvertSheetCall(exp);

        if (exp.Method.DeclaringType == typeof(DateTime))
        {
            if (exp.Method.Name == nameof(DateTime.AddMonths))
                return $"EDATE({InternalConvert(exp.Object)},{InternalConvert(exp.Arguments[0])})";
        }
        else if (CallMethodTypes.Contains(exp.Method.DeclaringType))
        {
            return
                $"{exp.Method.Name.ToUpper()}({string.Join(",", exp.Arguments.Select(c => InternalConvert(c)))})";
        }

        return $"unspport call type={exp.Method.DeclaringType} name={exp.Method.Name}";
    }

    private string ConvertMemberAccess(Expression expression)
    {
        var exp = expression as MemberExpression;
        var member = exp?.Member;
        if (member == null) return string.Empty;
        if (member.DeclaringType != typeof(DateTime))
            return $"unspport member access type={member.DeclaringType} name={member.Name}";
        switch (member.Name)
        {
            case "Now":
                return "NOW()";
            case "Year":
                return $"YEAR({InternalConvert(exp.Expression)})";
            case "Month":
                return $"MONTH({InternalConvert(exp.Expression)})";
            case "Day":
                return $"DAY({InternalConvert(exp.Expression)})";
            default:
                return $"unsupported member access type={member.DeclaringType} name={member.Name}";
        }
    }

    private string ConvertUnaryExpression(Expression expression)
    {
        if (!(expression is UnaryExpression unary)) return "null";
        var symbol = unary.NodeType == ExpressionType.Negate ? "-" : "unsupported unary symbol";
        return $"{symbol}{InternalConvert(unary.Operand)}";
    }

    /// <summary>
    ///     Translates a call on <see cref="SheetCellDictionary" />; <paramref name="exp" />.Object is the
    ///     <see cref="ColumnCellDictionary.Sheet" /> call that names the sheet.
    /// </summary>
    private string ConvertSheetCall(MethodCallExpression exp)
    {
        if (exp.Object is not MethodCallExpression sheetCall ||
            sheetCall.Method.Name != nameof(ColumnCellDictionary.Sheet) ||
            GetConstantValue(sheetCall.Arguments[0])?.ToString() is not { } sheetTitle)
            return "ERROR sheet reference must be c.Sheet(\"title\")";

        var prefix = QuoteSheetTitle(sheetTitle) + "!";
        var args = exp.Arguments;
        switch (exp.Method.Name)
        {
            case "get_Item":
                return args.Count == 2
                    ? $"{prefix}{GetColumn(args[0], sheetTitle)}{InternalConvert(args[1])}"
                    : $"{prefix}{GetColumn(args[0], sheetTitle)}{RowIndex + 1}";
            case nameof(SheetCellDictionary.Matrix):
                return
                    $"{prefix}{GetColumn(args[0], sheetTitle)}{args[1]}:{GetColumn(args[2], sheetTitle)}{args[3]}";
            case nameof(SheetCellDictionary.Columns):
                return $"{prefix}{GetColumn(args[0], sheetTitle)}:{GetColumn(args[1], sheetTitle)}";
            default:
                return $"unspport call type={exp.Method.DeclaringType} name={exp.Method.Name}";
        }
    }

    /// <summary>
    ///     Excel only requires quoting for titles with spaces or punctuation, but quoting is always valid,
    ///     so every title is quoted and embedded apostrophes are doubled.
    /// </summary>
    private static string QuoteSheetTitle(string title)
    {
        return "'" + title.Replace("'", "''") + "'";
    }

    /// <summary>
    ///     Reads the value of a literal, or of a local variable captured by the lambda (which the compiler
    ///     emits as a field access on a closure constant).
    /// </summary>
    private static object? GetConstantValue(Expression exp)
    {
        switch (exp)
        {
            case ConstantExpression constant:
                return constant.Value;
            case MemberExpression { Expression: ConstantExpression closure, Member: System.Reflection.FieldInfo field }:
                return field.GetValue(closure.Value);
            default:
                return null;
        }
    }

    private string GetColumn(Expression exp)
    {
        if (exp is not ConstantExpression constant) return "null";
        var key = constant.Value?.ToString();
        var columnIndex = Array.IndexOf(Columns, key);
        if (columnIndex == -1)
            throw new Excel2ObjectException($"refers to column [{key}], which is not a column of this sheet.");
        return ExcelColumnNameParser.Parse(columnIndex);
    }

    /// <summary>
    ///     Resolves a column on another sheet. When the sheet's titles are unknown, or the key is not one
    ///     of them, a key that already looks like a column letter (A, BC, ...) is used verbatim so sheets
    ///     the exporter did not write can still be referenced; anything else is rejected.
    /// </summary>
    private string GetColumn(Expression exp, string sheetTitle)
    {
        if (GetConstantValue(exp)?.ToString() is not { } key) return "null";
        var columns = SheetColumnsResolver?.Invoke(sheetTitle);
        var columnIndex = columns == null ? -1 : Array.IndexOf(columns, key);
        if (columnIndex != -1) return ExcelColumnNameParser.Parse(columnIndex);
        if (IsColumnLetters(key)) return key;
        throw new Excel2ObjectException(
            $"refers to column [{key}], which is neither a column title on sheet [{sheetTitle}] nor a column letter.");
    }

    private static bool IsColumnLetters(string key)
    {
        return key.Length is >= 1 and <= 3 && key.All(ch => ch is >= 'A' and <= 'Z');
    }

    private string InternalConvert(params Expression?[] expressions)
    {
        var expression = expressions[0];
        if (expression == null) return "";
        switch (expression.NodeType)
        {
            case ExpressionType.Convert:
                return InternalConvert((expression as UnaryExpression)?.Operand);
            case ExpressionType.Call:
                return ConvertCall(expression);
            case ExpressionType.MemberAccess:
                return ConvertMemberAccess(expression);
            case ExpressionType.Constant:
                return ConvertConstant(expression);
        }

        switch (expression)
        {
            case BinaryExpression _:
                return ConvertBinaryExpression(expression);
            case UnaryExpression _:
                return ConvertUnaryExpression(expression);
        }

        if (expression.NodeType != ExpressionType.NewArrayInit) return $"unsupported type {expressions[0]?.NodeType}";
        if (expression is not NewArrayExpression exp) return "null";
        return string.Join(",", exp.Expressions.Select(c => InternalConvert(c)));
    }
}