using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     One side of a cell's border, written the way CSS writes it: <c>1px solid #D0D0D0</c>, <c>2px
///     dashed red</c>, <c>none</c>.
/// </summary>
/// <remarks>
///     Excel has a fixed set of line styles rather than a width in pixels, so the width picks one of them:
///     1px is thin, 2px medium, 3px and up thick. The style word chooses between solid, dashed, dotted and
///     double; a dashed or dotted 2px line becomes Excel's medium variant of it.
/// </remarks>
internal readonly struct ExcelBorder : IEquatable<ExcelBorder>
{
    public BorderStyle Line { get; }
    public StyleColor? Color { get; }

    private ExcelBorder(BorderStyle line, StyleColor? color)
    {
        Line = line;
        Color = color;
    }

    public static ExcelBorder Parse(string border)
    {
        if (string.IsNullOrWhiteSpace(border))
            throw new Excel2ObjectException("A border cannot be empty; write one such as \"1px solid #D0D0D0\".");

        var width = 1;
        var style = "solid";
        StyleColor? color = null;

        foreach (var part in border.Split(new[] {' ', '\t'}, StringSplitOptions.RemoveEmptyEntries))
        {
            if (part.EndsWith("px", StringComparison.OrdinalIgnoreCase) &&
                int.TryParse(part.Substring(0, part.Length - 2), out var pixels))
                width = pixels;
            else if (Keyword(part) is { } named)
                width = named;
            else if (IsStyleWord(part))
                style = part.ToLowerInvariant();
            else
                try
                {
                    color = StyleColor.Parse(part);
                }
                catch (Excel2ObjectException e)
                {
                    throw new Excel2ObjectException(
                        $"[{border}] is no border: {e.Message} Write one as CSS writes it, e.g. " +
                        "\"1px solid #D0D0D0\", \"medium dashed red\" or \"none\".", e);
                }
        }

        return new ExcelBorder(Line2(style, width), color);
    }

    /// <summary>CSS names three widths as well as measuring them in pixels.</summary>
    private static int? Keyword(string part)
    {
        switch (part.ToLowerInvariant())
        {
            case "thin": return 1;
            case "medium": return 2;
            case "thick": return 3;
            default: return null;
        }
    }

    private static bool IsStyleWord(string part)
    {
        switch (part.ToLowerInvariant())
        {
            case "none":
            case "hidden":
            case "solid":
            case "dashed":
            case "dotted":
            case "double":
                return true;
            default:
                return false;
        }
    }

    private static BorderStyle Line2(string style, int width)
    {
        if (width <= 0) return BorderStyle.None;

        switch (style)
        {
            case "none":
            case "hidden":
                return BorderStyle.None;
            case "dashed":
                return width >= 2 ? BorderStyle.MediumDashed : BorderStyle.Dashed;
            case "dotted":
                return width >= 2 ? BorderStyle.MediumDashDot : BorderStyle.Dotted;
            case "double":
                return BorderStyle.Double;
            default:
                return width >= 3 ? BorderStyle.Thick : width == 2 ? BorderStyle.Medium : BorderStyle.Thin;
        }
    }

    public bool Equals(ExcelBorder other)
    {
        return Line == other.Line && Nullable.Equals(Color, other.Color);
    }

    public override bool Equals(object? obj)
    {
        return obj is ExcelBorder other && Equals(other);
    }

    public override int GetHashCode()
    {
        return ((int) Line * 397) ^ (Color?.GetHashCode() ?? 0);
    }

    public override string ToString()
    {
        return $"{Line}:{Color}";
    }
}
