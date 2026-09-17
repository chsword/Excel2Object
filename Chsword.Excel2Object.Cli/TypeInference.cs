using System.Globalization;

namespace Chsword.Excel2Object.Cli;

/// <summary>The CLR type a column of cell texts fits into, from most to least specific.</summary>
public enum InferredType
{
    Bool,
    Int,
    Long,
    Decimal,
    DateTime,
    String
}

/// <summary>
///     Infers a column type from the cell texts the importer returns. Every non-empty value has to fit the type;
///     a column that mixes kinds (or has no values at all) is a string column.
/// </summary>
public static class TypeInference
{
    private static readonly string[] DateFormats =
    {
        "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssK", "yyyy/MM/dd",
        "yyyy/MM/dd HH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy/MM/dd HH:mm"
    };

    public static InferredType Infer(IEnumerable<string> values)
    {
        var inference = new Inference();
        foreach (var value in values) inference.Observe(value);
        return inference.Result;
    }

    /// <summary>
    ///     一列的推断过程：一个值一个值地喂进来，不必先把整列攒在内存里。命令行要处理的文件可能有
    ///     几十万行，逐行读出便是为此。
    /// </summary>
    public sealed class Inference
    {
        /// <summary>由窄到宽，取第一个仍然成立的。静态存放：这个取值会被逐格问到。</summary>
        private static readonly InferredType[] Order =
        {
            InferredType.Bool, InferredType.Int, InferredType.Long, InferredType.Decimal, InferredType.DateTime
        };

        private readonly HashSet<InferredType> _candidates = new()
            {InferredType.Bool, InferredType.Int, InferredType.Long, InferredType.Decimal, InferredType.DateTime};

        private bool _seen;

        /// <summary>这一列上是否出现过空值——空值本身不参与类型判断，却决定该属性是否可空。</summary>
        public bool HasBlank { get; private set; }

        /// <summary>是否有过任何一行。一行都没有的列按字符串处理，且算作可空。</summary>
        public bool Any { get; private set; }

        public InferredType Result
        {
            get
            {
                if (!_seen) return InferredType.String;
                foreach (var candidate in Order)
                    if (_candidates.Contains(candidate))
                        return candidate;

                return InferredType.String;
            }
        }

        public void Observe(string value)
        {
            Any = true;
            if (string.IsNullOrWhiteSpace(value))
            {
                HasBlank = true;
                return;
            }

            _seen = true;
            if (_candidates.Count > 0) _candidates.RemoveWhere(candidate => !Fits(candidate, value));
        }
    }

    public static bool Fits(InferredType type, string value)
    {
        switch (type)
        {
            case InferredType.Bool:
                return TryParseBool(value, out _);
            case InferredType.Int:
                return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _);
            case InferredType.Long:
                return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _);
            case InferredType.Decimal:
                return decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
            case InferredType.DateTime:
                return TryParseDateTime(value, out _);
            default:
                return true;
        }
    }

    public static bool TryParseBool(string value, out bool result)
    {
        switch (value.Trim().ToUpperInvariant())
        {
            case "TRUE":
                result = true;
                return true;
            case "FALSE":
                result = false;
                return true;
            default:
                result = false;
                return false;
        }
    }

    public static bool TryParseDateTime(string value, out DateTime result)
    {
        return DateTime.TryParseExact(value.Trim(), DateFormats, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out result);
    }

    /// <summary>C# keyword for the type, nullable when the column contains empty cells.</summary>
    public static string ToCSharp(InferredType type, bool nullable)
    {
        var name = type switch
        {
            InferredType.Bool => "bool",
            InferredType.Int => "int",
            InferredType.Long => "long",
            InferredType.Decimal => "decimal",
            InferredType.DateTime => "DateTime",
            _ => "string"
        };
        return nullable || type == InferredType.String ? name + "?" : name;
    }
}
