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
    {
        Columns = columns;
        RowIndex = rowIndex;
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

    public string Convert(Expression? expression)
    {
        if (expression == null) return string.Empty;
        
        if (expression.NodeType == ExpressionType.Lambda)
        {
            var body = (expression as LambdaExpression)?.Body;
            if (body == null) return string.Empty;
            
            // 使用缓存机制优化性能
            return ExpressionCache.GetOrAdd(body, Columns, RowIndex, () => InternalConvert(body));
        }
        
        return string.Empty;
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

    private string GetColumn(Expression exp)
    {
        if (exp is not ConstantExpression constant) return "null";
        var key = constant.Value?.ToString();
        
        // 如果key是完整的单元格引用（如 "A1", "B5"），直接返回
        if (!string.IsNullOrEmpty(key) && IsExcelCellReference(key))
        {
            return key;
        }
        
        // 如果key是直接的Excel列名（如 "A", "B", "C"），直接返回
        if (!string.IsNullOrEmpty(key) && IsExcelColumnName(key))
        {
            return key;
        }
        
        // 否则查找在Columns数组中的位置
        var columnIndex = Array.IndexOf(Columns, key);
        return columnIndex == -1 ? $"ERROR key:{key}" : ExcelColumnNameParser.Parse(columnIndex);
    }

    private bool IsExcelColumnName(string key)
    {
        return !string.IsNullOrEmpty(key) && key.All(c => c >= 'A' && c <= 'Z');
    }

    private bool IsExcelCellReference(string key)
    {
        if (string.IsNullOrEmpty(key)) return false;
        
        // 检查是否符合Excel单元格引用格式（如A1, B5, AA123）
        int i = 0;
        // 检查列部分（字母）
        while (i < key.Length && key[i] >= 'A' && key[i] <= 'Z')
        {
            i++;
        }
        
        // 必须有至少一个字母，且后面跟着数字
        if (i == 0 || i == key.Length) return false;
        
        // 检查行部分（数字）
        while (i < key.Length && char.IsDigit(key[i]))
        {
            i++;
        }
        
        // 整个字符串必须被消费完
        return i == key.Length;
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