using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Chsword.Excel2Object.Functions;

namespace Chsword.Excel2Object.Internal
{
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
            new Dictionary<ExpressionType, string>
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

        public string Convert(Expression expression)
        {
            if (expression == null) return string.Empty;
            return expression.NodeType == ExpressionType.Lambda
                ? InternalConvert((expression as LambdaExpression)?.Body)
                : string.Empty;
        }

        private static string ConvertConstant(Expression expression)
        {
            var exp = expression as ConstantExpression;
            return exp?.Type == typeof(bool) ? exp.ToString().ToUpper() : exp?.ToString();
        }

        private string ConvertBinaryExpression(Expression expression)
        {
            if (!(expression is BinaryExpression binary)) return "null";
            string symbol = $"unsupport binary symbol:{binary.NodeType}";
            if (BinarySymbolDictionary.ContainsKey(binary.NodeType))
            {
                symbol = BinarySymbolDictionary[binary.NodeType];
            }

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
            {
                return exp.Arguments.Count == 2
                    ? $"{GetColumn(exp.Arguments[0])}{InternalConvert(exp.Arguments[1])}"
                    : $"{GetColumn(exp.Arguments[0])}{RowIndex + 1}";
            }

            if (exp.Object.Type == typeof(ColumnCellDictionary) &&
                exp.Method.Name == nameof(ColumnCellDictionary.Matrix))
            {
                return
                    $"{GetColumn(exp.Arguments[0])}{exp.Arguments[1]}:{GetColumn(exp.Arguments[2])}{exp.Arguments[3]}";
            }

            if (exp.Method.DeclaringType == typeof(DateTime))
            {
                if (exp.Method.Name == nameof(DateTime.AddMonths))
                {
                    return $"EDATE({InternalConvert(exp.Object)},{InternalConvert(exp.Arguments[0])})";
                }
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
                    return $"unspport member access type={member.DeclaringType} name={member.Name}";
            }
        }

        private string ConvertUnaryExpression(Expression expression)
        {
            if (!(expression is UnaryExpression unary)) return "null";
            var symbol = unary.NodeType == ExpressionType.Negate ? "-" : "unsupport unary symbol";
            return $"{symbol}{InternalConvert(unary.Operand)}";
        }

        string GetColumn(Expression exp)
        {
            if (!(exp is ConstantExpression constant)) return "null";
            var key = constant.Value.ToString();
            var columnIndex = Array.IndexOf(Columns, key);
            return columnIndex == -1 ? $"ERROR key:{key}" : ExcelColumnNameParser.Parse(columnIndex);
        }

        private string InternalConvert(params Expression[] expressions)
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

            if (expression.NodeType != ExpressionType.NewArrayInit) return $"unsupport type {expressions[0].NodeType}";
            if (!(expression is NewArrayExpression exp)) return "null";
            return string.Join(",", exp.Expressions.Select(c => InternalConvert(c)));
        }
    }
}