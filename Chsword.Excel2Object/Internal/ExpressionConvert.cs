using Chsword.Excel2Object.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace Chsword.Excel2Object.Internal
{
    internal class ExpressionConvert
    {
        public string[] Columns { get; set; }
        public int RowIndex { get; set; }

        public ExpressionConvert(string[] columns, int rowIndex)
        {
            Columns = columns;
            RowIndex = rowIndex;
        }
        public string Convert(Expression expression)
        {
            if (expression == null) return string.Empty;
            if (expression.NodeType == ExpressionType.Lambda)
            {
                return InternalConvert((expression as LambdaExpression)?.Body);
            }

            return string.Empty;
        }

        string GetColumn(Expression exp)
        {
            var constant = exp as ConstantExpression;
            var key = constant.Value.ToString();
            var columnIndex = Array.IndexOf(Columns, key);
            if (columnIndex == -1)
            {
                return $"ERROR key:{key}";
            }

            return ExcelColumnNameParser.Parse(columnIndex);
        }

        private string InternalConvert(params Expression[] expressions)
        {
            var expression = expressions[0];
            if (expression == null) return "";
            if (expression.NodeType == ExpressionType.Convert)
            {
                return InternalConvert((expression as UnaryExpression)?.Operand);
            }
            if(expression.NodeType == ExpressionType.Call)
            {
                var exp = expression as MethodCallExpression;
                if (exp.Method.Name == "get_Item" && 
                    (exp.Object.Type == typeof(ColumnCellDictionary)
                     || exp.Object.Type == typeof(Dictionary<string,ColumnValue>)
                    )
                    )
                {
                    return $"{GetColumn(exp.Arguments[0])}{RowIndex + 1}";
                }

                if (exp.Object.Type == typeof(ColumnCellDictionary))
                {
                    if (exp.Method.Name == nameof(ColumnCellDictionary.Matrix))
                    {
                        return $"{GetColumn(exp.Arguments[0])}{exp.Arguments[1]}:{GetColumn(exp.Arguments[2])}{exp.Arguments[3]}";
                    }
                }
                if (exp.Method.DeclaringType == typeof(DateTime)) { 
                    if (exp.Method.Name == nameof(DateTime.AddMonths))
                    {
                        return $"EDATE({InternalConvert(exp.Object)},{InternalConvert(exp.Arguments[0])})";
                    }
                }else if (exp.Method.DeclaringType == typeof(IMathFunction) ||
                          exp.Method.DeclaringType == typeof(IStatisticsFunction) ||
                          exp.Method.DeclaringType == typeof(IAllFunction))
                {
                    return $"{exp.Method.Name.ToUpper()}({string.Join(",",exp.Arguments.Select(c => InternalConvert(c)))})";
                }
                
                return $"unspport call type={exp.Method.DeclaringType} name={exp.Method.Name}";
            }
            if (expression.NodeType == ExpressionType.MemberAccess)
            {
                var exp= (expression as MemberExpression);
                var member = exp?.Member;
                if (member == null) return string.Empty;
                if (member.DeclaringType == typeof(DateTime))
                {
                    if (member.Name == "Now")
                    {
                        return "NOW()";
                    }
                    if (member.Name == "Year")
                    {
                        return $"YEAR({InternalConvert(exp.Expression)})";
                    }
                    if (member.Name == "Month")
                    {
                        return $"MONTH({InternalConvert(exp.Expression)})";
                    }
                    if (member.Name == "Day")
                    {
                        return $"DAY({InternalConvert(exp.Expression)})";
                    }
                   
                }


                return $"unspport member access type={member.DeclaringType } name={member.Name}";
            }

            if (expression.NodeType == ExpressionType.Constant)
            {
                return (expression as ConstantExpression)?.ToString();
            }
 
            if (expression is BinaryExpression)
            { 
                var binary = expression as BinaryExpression;
                string symbol = $"unsupport binary symbol:{binary.NodeType}";
                if (BinarySymbolDictionary.ContainsKey(binary.NodeType))
                {
                    symbol = BinarySymbolDictionary[binary.NodeType];
                }
                return $"{InternalConvert(binary.Left)}{symbol}{InternalConvert(binary.Right)}";
            }
            if (expression is UnaryExpression)
            {
                var unary = expression as UnaryExpression;
                string symbol = "unsupport unary symbol";
                if (unary.NodeType == ExpressionType.Negate)
                {
                    symbol = "-";
                }
                return $"{symbol}{InternalConvert(unary.Operand)}";
            }

            if (expression.NodeType == ExpressionType.NewArrayInit)
            {
                var exp = expression as NewArrayExpression;
                return string.Join(",", exp.Expressions.Select(c => InternalConvert(c)));
            }

            return $"unsupport type {expressions[0].NodeType}";
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
    }
}