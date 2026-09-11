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
        var seen = false;
        var candidates = new HashSet<InferredType>
            {InferredType.Bool, InferredType.Int, InferredType.Long, InferredType.Decimal, InferredType.DateTime};
        foreach (var value in values)
        {
            if (string.IsNullOrWhiteSpace(value)) continue;
            seen = true;
            candidates.RemoveWhere(candidate => !Fits(candidate, value));
            if (candidates.Count == 0) break;
        }

        if (!seen) return InferredType.String;
        foreach (var candidate in new[]
                 {
                     InferredType.Bool, InferredType.Int, InferredType.Long, InferredType.Decimal,
                     InferredType.DateTime
                 })
            if (candidates.Contains(candidate))
                return candidate;
        return InferredType.String;
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
