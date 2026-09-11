using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object.Internal;

internal class ExpressionConvert
{
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

    /// <summary>
    ///     Compiled evaluators for the sheet-independent sub-expressions of a formula, keyed by the node
    ///     they came from. A formula column holds one expression tree and is converted once per row, so
    ///     without this the same lambda would be compiled again for every row of the sheet.
    /// </summary>
    private static readonly ConditionalWeakTable<Expression, Func<object?>> Evaluators = new();

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
            [ExpressionType.And] = "&",
            [ExpressionType.ExclusiveOr] = "^"
        };

    /// <summary>
    ///     <see cref="System.Math" /> methods that Excel spells the same way, only in upper case.
    /// </summary>
    private static HashSet<string> DirectMathMethods { get; } = new(StringComparer.Ordinal)
    {
        "Abs", "Sqrt", "Exp", "Sign", "Sin", "Cos", "Tan", "Asin", "Acos", "Atan",
        "Sinh", "Cosh", "Tanh", "Max", "Min", "Log10"
    };

    private string[] Columns { get; }
    private int RowIndex { get; }
    private Func<string, string[]?>? SheetColumnsResolver { get; }

    /// <summary>
    ///     The lambda parameter that stands for the exported model when the formula was added through
    ///     <c>FormulaColumnsCollection.Add&lt;TModel&gt;</c>; null for title-only formulas.
    /// </summary>
    private ParameterExpression? ModelParameter { get; set; }

    /// <summary>Column title of each exported property of the model, by property name.</summary>
    private Dictionary<string, string> ModelPropertyTitles { get; set; } = new();

    public string Convert(Expression? expression)
    {
        if (expression is not LambdaExpression lambda) return string.Empty;
        if (lambda.Parameters.Count >= 2)
        {
            ModelParameter = lambda.Parameters[1];
            ModelPropertyTitles = ExcelUtil.GetPropertiesAttributesDict(ModelParameter.Type)
                .ToDictionary(c => c.Key.Name, c => c.Value.Title!);
        }

        return InternalConvert(lambda.Body);
    }

    private static string ConvertConstant(Expression expression)
    {
        return expression is ConstantExpression exp ? FormatValue(exp.Value) : string.Empty;
    }

    /// <summary>
    ///     Writes a .NET value as formula text. Numbers are always written with an invariant decimal
    ///     point, since a formula that carries a locale's comma would not parse in Excel.
    /// </summary>
    private static string FormatValue(object? value)
    {
        switch (value)
        {
            case null:
                return string.Empty;
            case bool b:
                return b ? "TRUE" : "FALSE";
            case string s:
                return QuoteText(s);
            case char ch:
                return QuoteText(ch.ToString());
            case DateTime dt:
                return FormatDateTime(dt);
            case Enum e:
                return System.Convert.ToInt64(e, CultureInfo.InvariantCulture)
                    .ToString(CultureInfo.InvariantCulture);
            case IFormattable formattable:
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            default:
                return value.ToString() ?? string.Empty;
        }
    }

    /// <summary>Excel text literals are double quoted, and an embedded quote is doubled.</summary>
    private static string QuoteText(string text)
    {
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    private static string FormatDateTime(DateTime value)
    {
        var date = $"DATE({value.Year},{value.Month},{value.Day})";
        if (value.TimeOfDay == TimeSpan.Zero) return date;
        var time = $"TIME({value.Hour},{value.Minute},{value.Second})";
        // TIME only takes whole seconds, so anything finer is added as its fraction of a day
        var subSecond = value.Ticks % TimeSpan.TicksPerSecond;
        if (subSecond == 0) return $"{date}+{time}";
        var fraction = (double) subSecond / TimeSpan.TicksPerDay;
        return $"{date}+{time}+{fraction.ToString("R", CultureInfo.InvariantCulture)}";
    }

    private string ConvertBinaryExpression(Expression expression)
    {
        if (expression is not BinaryExpression binary) return "null";
        if (FunctionOfBinary(binary) is { } function)
            return $"{function}({InternalConvert(binary.Left)},{InternalConvert(binary.Right)})";
        var symbol = BinarySymbol(binary);
        var precedence = Precedence(binary);
        var left = InternalConvert(binary.Left);
        var right = InternalConvert(binary.Right);
        // Excel formulas are plain text, so the tree's grouping has to be restored with parentheses
        // wherever an operand binds less tightly than its parent (or equally tightly on the right of
        // a non-associative operator, as in a-(b-c)).
        if (Precedence(binary.Left) < precedence) left = $"({left})";
        var rightPrecedence = Precedence(binary.Right);
        if (rightPrecedence < precedence ||
            (rightPrecedence == precedence && binary.NodeType is ExpressionType.Subtract
                 or ExpressionType.Divide or ExpressionType.ExclusiveOr))
            right = $"({right})";

        return $"{left}{symbol}{right}";
    }

    /// <summary>
    ///     The Excel function a binary node is written as, for the operators Excel has no symbol for;
    ///     null when the node is written with an operator instead.
    /// </summary>
    private static string? FunctionOfBinary(BinaryExpression binary)
    {
        // a lifted operator on nullable operands is typed bool?/int?, so compare the underlying type
        var type = Nullable.GetUnderlyingType(binary.Type) ?? binary.Type;
        var isBoolean = type == typeof(bool);
        var isInteger = type == typeof(int) || type == typeof(long) ||
                        type == typeof(short) || type == typeof(byte) ||
                        type == typeof(uint) || type == typeof(ulong);
        switch (binary.NodeType)
        {
            case ExpressionType.Modulo:
                return "MOD";
            case ExpressionType.AndAlso:
                return "AND";
            case ExpressionType.OrElse:
                return "OR";
            case ExpressionType.And:
                return isBoolean ? "AND" : isInteger ? Future("BITAND") : null;
            case ExpressionType.Or:
                return isInteger ? Future("BITOR") : "OR";
            case ExpressionType.ExclusiveOr:
                // ^ on a ColumnValue is Excel's power operator; on bools and integers it really is a xor
                return isBoolean ? Future("XOR") : isInteger ? Future("BITXOR") : null;
            default:
                return null;
        }
    }

    private static string BinarySymbol(BinaryExpression binary)
    {
        // string concatenation compiles to Add with String.Concat; Excel spells it &
        if (binary.NodeType == ExpressionType.Add && binary.Method?.DeclaringType == typeof(string)) return "&";
        return BinarySymbolDictionary.TryGetValue(binary.NodeType, out var value)
            ? value
            : $"unsupported binary symbol:{binary.NodeType}";
    }

    /// <summary>
    ///     Excel operator precedence of the operator an expression will be rendered with; atoms (cells,
    ///     constants, function calls) rank highest so they never get parenthesized.
    /// </summary>
    private static int Precedence(Expression expression)
    {
        while (expression is UnaryExpression {NodeType: ExpressionType.Convert} convert)
            expression = convert.Operand;
        if (expression is not BinaryExpression binary) return int.MaxValue;
        if (FunctionOfBinary(binary) != null) return int.MaxValue;
        switch (BinarySymbol(binary))
        {
            case "^":
                return 5;
            case "*":
            case "/":
                return 4;
            case "+":
            case "-":
                return 3;
            case "&":
                return 2;
            default:
                return 1; // comparisons
        }
    }

    private string ConvertCall(Expression expression)
    {
        if (expression is not MethodCallExpression exp) return "null";
        if (exp.Object != null)
        {
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
                    $"{GetColumn(exp.Arguments[0])}{ConvertRowNumber(exp.Arguments[1])}:{GetColumn(exp.Arguments[2])}{ConvertRowNumber(exp.Arguments[3])}";

            if (exp.Object.Type == typeof(ColumnCellDictionary) &&
                exp.Method.Name == nameof(ColumnCellDictionary.Columns))
                return $"{GetColumn(exp.Arguments[0])}:{GetColumn(exp.Arguments[1])}";

            if (exp.Object.Type == typeof(SheetCellDictionary))
                return ConvertSheetCall(exp);
        }

        if (IsExcelFunction(exp.Method.DeclaringType))
            return $"{ExcelFunctionName(exp.Method)}({JoinArguments(exp.Arguments)})";

        if (exp.Method.DeclaringType == typeof(DateTime) && ConvertDateTimeCall(exp) is { } dateCall)
            return dateCall;

        if (exp.Method.DeclaringType == typeof(string) && ConvertStringCall(exp) is { } stringCall)
            return stringCall;

        if (exp.Method.DeclaringType == typeof(System.Math) && ConvertMathCall(exp) is { } mathCall)
            return mathCall;

        // A call that uses nothing from the sheet is just a value, e.g. a helper the formula was built with.
        if (TryEvaluate(exp, out var value)) return FormatValue(value);

        return $"unspport call type={exp.Method.DeclaringType} name={exp.Method.Name}";
    }

    /// <summary>
    ///     The name a function interface method is written under, from its attribute or its method name.
    ///     Functions Excel gained after 2007 carry the <c>_xlfn.</c> prefix the file format stores them with.
    /// </summary>
    private static string ExcelFunctionName(MethodInfo method)
    {
        var attribute = method.GetCustomAttribute<ExcelFunctionNameAttribute>();
        return attribute?.StoredName ?? method.Name.ToUpperInvariant();
    }

    /// <summary>A function Excel gained after 2007, spelled the way the file format stores it.</summary>
    private static string Future(string name)
    {
        return ExcelFunctionNameAttribute.FuturePrefix + name;
    }

    private static bool IsExcelFunction(Type? declaringType)
    {
        return declaringType != null && typeof(IExcelFunction).IsAssignableFrom(declaringType);
    }

    private string JoinArguments(IEnumerable<Expression> arguments)
    {
        return string.Join(",", arguments.Select(c => InternalConvert(c)));
    }

    private string? ConvertDateTimeCall(MethodCallExpression exp)
    {
        if (exp.Object == null) return null;
        var target = InternalConvert(exp.Object);
        switch (exp.Method.Name)
        {
            case nameof(DateTime.AddMonths):
                return $"EDATE({target},{InternalConvert(exp.Arguments[0])})";
            case nameof(DateTime.AddYears):
                return $"EDATE({target},({InternalConvert(exp.Arguments[0])})*12)";
            case nameof(DateTime.AddDays):
                // a date plus a number of days; parenthesized so it composes like any other operand
                return $"({target}+{InternalConvert(exp.Arguments[0])})";
            default:
                return null;
        }
    }

    private string? ConvertStringCall(MethodCallExpression exp)
    {
        var args = exp.Arguments;
        if (exp.Object == null)
            switch (exp.Method.Name)
            {
                case nameof(string.Concat):
                    return $"{Future("CONCAT")}({JoinArguments(args)})";
                case nameof(string.Join):
                    return
                        // ignore_empty FALSE: string.Join keeps empty entries
                        $"{Future("TEXTJOIN")}({InternalConvert(args[0])},FALSE,{JoinArguments(args.Skip(1))})";
                case nameof(string.IsNullOrEmpty):
                    return $"({InternalConvert(args[0])}=\"\")";
                // IsNullOrWhiteSpace has no equivalent: it counts tabs and the Unicode spaces, while
                // Excel's TRIM only strips the ASCII space.
                default:
                    return null;
            }

        // Every case below pins its argument count, which is what keeps the overloads Excel cannot
        // express - the ones taking a StringComparison, a CultureInfo, a start index or a char[] -
        // off the translated path. Dropping such an argument would change what the formula means.
        // Trim() is absent for the same reason: Excel's TRIM also collapses runs of spaces inside
        // the text, so it does not mean what .NET's Trim() means.
        var target = InternalConvert(exp.Object);
        switch (exp.Method.Name)
        {
            case nameof(string.ToUpper) when args.Count == 0:
                return $"UPPER({target})";
            case nameof(string.ToLower) when args.Count == 0:
                return $"LOWER({target})";
            case nameof(string.Substring) when args.Count == 1:
                return $"MID({target},{OneBased(args[0])},LEN({target}))";
            case nameof(string.Substring) when args.Count == 2:
                return $"MID({target},{OneBased(args[0])},{InternalConvert(args[1])})";
            case nameof(string.Replace) when args.Count == 2:
                return $"SUBSTITUTE({target},{InternalConvert(args[0])},{InternalConvert(args[1])})";
            case nameof(string.Contains) when args.Count == 1:
                // FIND, not SEARCH: the .NET methods are case sensitive and take no wildcards
                return $"ISNUMBER(FIND({InternalConvert(args[0])},{target}))";
            case nameof(string.StartsWith) when args.Count == 1:
                // EXACT, not =: Excel's = ignores case
                return $"EXACT(LEFT({target},LEN({InternalConvert(args[0])})),{InternalConvert(args[0])})";
            case nameof(string.EndsWith) when args.Count == 1:
                return $"EXACT(RIGHT({target},LEN({InternalConvert(args[0])})),{InternalConvert(args[0])})";
            case nameof(string.IndexOf) when args.Count == 1:
                // Excel counts from 1 and errors when the text is absent, .NET counts from 0 and returns -1
                return $"IFERROR(FIND({InternalConvert(args[0])},{target})-1,-1)";
            default:
                return null;
        }
    }

    private string? ConvertMathCall(MethodCallExpression exp)
    {
        var args = exp.Arguments;
        var name = exp.Method.Name;
        // Math.Round is deliberately absent: it rounds halves to even where Excel's ROUND rounds them
        // away from zero, so there is no formula that means the same thing. Overloads taking a
        // MidpointRounding are rejected outright rather than falling through to a near equivalent.
        if (args.Any(a => a.Type == typeof(MidpointRounding))) return null;
        if (DirectMathMethods.Contains(name))
            return $"{name.ToUpperInvariant()}({JoinArguments(args)})";
        switch (name)
        {
            case nameof(System.Math.Pow):
                return $"POWER({JoinArguments(args)})";
            case nameof(System.Math.Atan2):
                // .NET takes (y, x), Excel takes (x_num, y_num)
                return $"ATAN2({InternalConvert(args[1])},{InternalConvert(args[0])})";
            case nameof(System.Math.Log):
                return args.Count == 1 ? $"LN({InternalConvert(args[0])})" : $"LOG({JoinArguments(args)})";
            case nameof(System.Math.Ceiling):
                return $"{Future("CEILING.MATH")}({InternalConvert(args[0])})";
            case nameof(System.Math.Floor):
                return $"{Future("FLOOR.MATH")}({InternalConvert(args[0])})";
            case nameof(System.Math.Truncate):
                return $"TRUNC({InternalConvert(args[0])})";
            default:
                return null;
        }
    }

    /// <summary>Converts a .NET 0-based offset to the 1-based position Excel's text functions take.</summary>
    private string OneBased(Expression expression)
    {
        if (TryEvaluate(expression, out var value) && value is int index) return (index + 1).ToString();
        var text = InternalConvert(expression);
        return Precedence(expression) < 3 ? $"({text})+1" : $"{text}+1";
    }

    private string ConvertConditional(Expression expression)
    {
        if (expression is not ConditionalExpression exp) return "null";
        return
            $"IF({InternalConvert(exp.Test)},{InternalConvert(exp.IfTrue)},{InternalConvert(exp.IfFalse)})";
    }

    private string ConvertMemberAccess(Expression expression)
    {
        var exp = expression as MemberExpression;
        var member = exp?.Member;
        if (member == null) return string.Empty;
        if (ModelParameter != null && exp!.Expression == ModelParameter)
            return ConvertModelProperty(member);
        // m.Price.Value on a nullable property refers to the same cell
        var nullableUnderlyingType = member.DeclaringType == null
            ? null
            : Nullable.GetUnderlyingType(member.DeclaringType);
        if (nullableUnderlyingType != null)
        {
            if (member.Name == "Value") return InternalConvert(exp!.Expression);
            if (member.Name == "HasValue") return $"NOT(ISBLANK({InternalConvert(exp!.Expression)}))";
        }

        if (member.DeclaringType == typeof(DateTime) && ConvertDateTimeMember(exp!, member) is { } dateMember)
            return dateMember;
        if (member.DeclaringType == typeof(string) && member.Name == nameof(string.Length))
            return $"LEN({InternalConvert(exp!.Expression)})";
        // A member that uses nothing from the sheet is just a value, e.g. a variable the lambda captured.
        if (TryEvaluate(exp!, out var value)) return FormatValue(value);
        return $"unspport member access type={member.DeclaringType} name={member.Name}";
    }

    private string? ConvertDateTimeMember(MemberExpression exp, MemberInfo member)
    {
        switch (member.Name)
        {
            case nameof(DateTime.Now):
                return "NOW()";
            case nameof(DateTime.Today):
                return "TODAY()";
            case nameof(DateTime.Year):
                return $"YEAR({InternalConvert(exp.Expression)})";
            case nameof(DateTime.Month):
                return $"MONTH({InternalConvert(exp.Expression)})";
            case nameof(DateTime.Day):
                return $"DAY({InternalConvert(exp.Expression)})";
            case nameof(DateTime.Hour):
                return $"HOUR({InternalConvert(exp.Expression)})";
            case nameof(DateTime.Minute):
                return $"MINUTE({InternalConvert(exp.Expression)})";
            case nameof(DateTime.Second):
                return $"SECOND({InternalConvert(exp.Expression)})";
            case nameof(DateTime.Date):
                return $"INT({InternalConvert(exp.Expression)})";
            default:
                return null;
        }
    }

    /// <summary>
    ///     Translates <c>m.Property</c> to the cell of that property's column on the current row.
    /// </summary>
    private string ConvertModelProperty(MemberInfo property)
    {
        if (!ModelPropertyTitles.TryGetValue(property.Name, out var title))
            throw new Excel2ObjectException(
                $"refers to property [{property.Name}] of {ModelParameter!.Type.Name}, which is not an exported column (no [ExcelTitle] or [Display] attribute).");
        return $"{GetColumnByTitle(title)}{RowIndex + 1}";
    }

    private string ConvertUnaryExpression(Expression expression)
    {
        if (expression is not UnaryExpression unary) return "null";
        // ExpressionType.Not covers both !x and ~x; Excel has no bitwise complement
        if (unary.NodeType == ExpressionType.Not)
        {
            var type = Nullable.GetUnderlyingType(unary.Type) ?? unary.Type;
            if (type == typeof(bool) || type == typeof(ColumnValue))
                return $"NOT({InternalConvert(unary.Operand)})";
            return $"unsupported unary symbol:~ on {type.Name}";
        }
        if (unary.NodeType == ExpressionType.UnaryPlus)
            return InternalConvert(unary.Operand);
        var symbol = unary.NodeType == ExpressionType.Negate ? "-" : "unsupported unary symbol";
        var operand = InternalConvert(unary.Operand);
        if (Precedence(unary.Operand) != int.MaxValue) operand = $"({operand})";
        return $"{symbol}{operand}";
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
                    $"{prefix}{GetColumn(args[0], sheetTitle)}{ConvertRowNumber(args[1])}:{GetColumn(args[2], sheetTitle)}{ConvertRowNumber(args[3])}";
            case nameof(SheetCellDictionary.Columns):
                return $"{prefix}{GetColumn(args[0], sheetTitle)}:{GetColumn(args[1], sheetTitle)}";
            default:
                return $"unspport call type={exp.Method.DeclaringType} name={exp.Method.Name}";
        }
    }

    /// <summary>Row numbers of a range are plain integers, whether written inline or held in a variable.</summary>
    private string ConvertRowNumber(Expression expression)
    {
        return TryEvaluate(expression, out var value) ? FormatValue(value) : InternalConvert(expression);
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
            case MemberExpression {Expression: ConstantExpression closure, Member: FieldInfo field}:
                return field.GetValue(closure.Value);
            default:
                return TryEvaluate(exp, out var value) ? value : null;
        }
    }

    /// <summary>
    ///     Evaluates a sub-expression that refers to nothing in the sheet - a captured variable, a
    ///     constant folded call - so it can be written into the formula as a literal.
    /// </summary>
    /// <remarks>
    ///     A converter is built per row, so this runs once per row per literal. The two shapes that
    ///     cover almost everything - a literal, and the field access the compiler emits for a captured
    ///     variable - are read directly; only the rest falls back to compiling, and that compilation is
    ///     cached against the expression node, which the formula column reuses for every row.
    /// </remarks>
    private static bool TryEvaluate(Expression expression, out object? value)
    {
        switch (expression)
        {
            case ConstantExpression constant:
                value = constant.Value;
                return true;
            case MemberExpression {Expression: ConstantExpression owner} member:
                switch (member.Member)
                {
                    case FieldInfo field:
                        value = field.GetValue(owner.Value);
                        return true;
                    case PropertyInfo property when property.GetIndexParameters().Length == 0:
                        value = property.GetValue(owner.Value);
                        return true;
                }

                break;
        }

        value = null;
        if (ContainsParameter(expression)) return false;
        try
        {
            value = Evaluators.GetValue(expression, Compile)();
            return true;
        }
        catch
        {
            // Anything that only makes sense as formula text (the ColumnValue members all throw) is
            // not a value, and is translated by the callers instead.
            return false;
        }
    }

    private static Func<object?> Compile(Expression expression)
    {
        return Expression.Lambda<Func<object?>>(Expression.Convert(expression, typeof(object))).Compile();
    }

    private static bool ContainsParameter(Expression expression)
    {
        var finder = new ParameterFinder();
        finder.Visit(expression);
        return finder.Found;
    }

    private string GetColumn(Expression exp)
    {
        if (GetConstantValue(exp)?.ToString() is not { } title) return "null";
        return GetColumnByTitle(title);
    }

    private string GetColumnByTitle(string? title)
    {
        var columnIndex = Array.IndexOf(Columns, title);
        if (columnIndex == -1)
            throw new Excel2ObjectException($"refers to column [{title}], which is not a column of this sheet.");
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
            case ExpressionType.Conditional:
                return ConvertConditional(expression);
        }

        switch (expression)
        {
            case BinaryExpression _:
                return ConvertBinaryExpression(expression);
            case UnaryExpression _:
                return ConvertUnaryExpression(expression);
        }

        // params arrays are spread into the argument list of the function that takes them
        if (expression.NodeType == ExpressionType.NewArrayInit && expression is NewArrayExpression exp)
            return string.Join(",", exp.Expressions.Select(c => InternalConvert(c)));
        // anything else that does not touch the sheet, e.g. new DateTime(2024, 3, 1), is a literal
        if (TryEvaluate(expression, out var value)) return FormatValue(value);
        return $"unsupported type {expression.NodeType}";
    }

    /// <summary>Tells whether an expression depends on a lambda parameter, i.e. on the sheet.</summary>
    private class ParameterFinder : ExpressionVisitor
    {
        public bool Found { get; private set; }

        protected override Expression VisitParameter(ParameterExpression node)
        {
            Found = true;
            return node;
        }
    }
}
