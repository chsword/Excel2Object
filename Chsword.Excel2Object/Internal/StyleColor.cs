using System.Globalization;
using Chsword.Excel2Object.Styles;
using NPOI.HSSF.Util;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     A colour a style asks for, written the way CSS writes one (<c>#4472C4</c>, <c>#FFF</c>) or picked
///     from <see cref="ExcelStyleColor" />.
/// </summary>
/// <remarks>
///     .xlsx stores the colour itself, so it comes out exactly as written. .xls has only the 56 colours of
///     its palette, so a colour written as hex is matched to the nearest one it holds.
/// </remarks>
internal readonly struct StyleColor : IEquatable<StyleColor>
{
    public byte R { get; }
    public byte G { get; }
    public byte B { get; }

    /// <summary>The palette index this came from, or 0 when it was written as hex.</summary>
    public short Indexed { get; }

    /// <summary>
    ///     True when the colour was picked from <see cref="ExcelStyleColor" />. Such a colour is written as
    ///     its palette index in both formats - that is what it has always meant - while a colour written as
    ///     hex is stored as itself where the format can hold it.
    /// </summary>
    public bool IsFromPalette => Indexed > 0;

    private StyleColor(byte r, byte g, byte b, short indexed)
    {
        R = r;
        G = g;
        B = b;
        Indexed = indexed;
    }

    public static StyleColor FromPalette(ExcelStyleColor color)
    {
        var index = (short) color;
        var rgb = HSSFColor.GetIndexHash().TryGetValue(index, out var known) ? known.RGB : new byte[] {0, 0, 0};
        return new StyleColor(rgb[0], rgb[1], rgb[2], index);
    }

    /// <summary>The colour names CSS started with, which a stylesheet may as well accept.</summary>
    private static readonly Dictionary<string, string> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000", ["silver"] = "C0C0C0", ["gray"] = "808080", ["grey"] = "808080",
        ["white"] = "FFFFFF", ["maroon"] = "800000", ["red"] = "FF0000", ["purple"] = "800080",
        ["fuchsia"] = "FF00FF", ["green"] = "008000", ["lime"] = "00FF00", ["olive"] = "808000",
        ["yellow"] = "FFFF00", ["navy"] = "000080", ["blue"] = "0000FF", ["teal"] = "008080",
        ["aqua"] = "00FFFF", ["cyan"] = "00FFFF", ["orange"] = "FFA500", ["pink"] = "FFC0CB"
    };

    /// <summary>Reads <c>#RGB</c>, <c>#RRGGBB</c>, the bare digits of either, or a CSS colour name.</summary>
    public static StyleColor Parse(string color)
    {
        if (string.IsNullOrWhiteSpace(color))
            throw new Excel2ObjectException("A colour cannot be empty; write one such as \"#4472C4\".");

        var text = color.Trim();
        if (Named.TryGetValue(text, out var named)) text = named;
        text = text.TrimStart('#');
        if (text.Length == 3)
            text = new string(new[] {text[0], text[0], text[1], text[1], text[2], text[2]});

        if (text.Length != 6 || !ushort.TryParse(text.Substring(0, 2), NumberStyles.HexNumber,
                CultureInfo.InvariantCulture, out var r)
                             || !ushort.TryParse(text.Substring(2, 2), NumberStyles.HexNumber,
                                 CultureInfo.InvariantCulture, out var g)
                             || !ushort.TryParse(text.Substring(4, 2), NumberStyles.HexNumber,
                                 CultureInfo.InvariantCulture, out var b))
            throw new Excel2ObjectException(
                $"[{color}] is no colour. Write one as #RRGGBB or #RGB, e.g. \"#4472C4\", name one of " +
                $"{string.Join(", ", Named.Keys)}, or pick an {nameof(ExcelStyleColor)}.");

        return new StyleColor((byte) r, (byte) g, (byte) b, 0);
    }

    public byte[] Rgb()
    {
        return new[] {R, G, B};
    }

    /// <summary>
    ///     The palette index .xls has to use: the one this colour came from, or the nearest the palette
    ///     holds. Nearest is measured in CIELAB, where a distance matches what the eye sees - in plain RGB
    ///     a light grey comes out nearer to a pale lavender than to a grey.
    /// </summary>
    public short PaletteIndex()
    {
        if (Indexed > 0) return Indexed;

        var target = Lab(R, G, B);
        var best = (short) 8;
        var bestDistance = double.MaxValue;
        foreach (var entry in HSSFColor.GetIndexHash())
        {
            var rgb = entry.Value.RGB;
            var lab = Lab(rgb[0], rgb[1], rgb[2]);
            var distance = ((lab[0] - target[0]) * (lab[0] - target[0])) +
                           ((lab[1] - target[1]) * (lab[1] - target[1])) +
                           ((lab[2] - target[2]) * (lab[2] - target[2]));
            if (distance >= bestDistance) continue;
            bestDistance = distance;
            best = (short) entry.Key;
        }

        return best;
    }

    /// <summary>sRGB to CIELAB, through XYZ with the D65 white point.</summary>
    private static double[] Lab(byte r, byte g, byte b)
    {
        var red = Linear(r);
        var green = Linear(g);
        var blue = Linear(b);

        var x = ((red * 0.4124) + (green * 0.3576) + (blue * 0.1805)) / 0.95047;
        var y = (red * 0.2126) + (green * 0.7152) + (blue * 0.0722);
        var z = ((red * 0.0193) + (green * 0.1192) + (blue * 0.9505)) / 1.08883;

        return new[] {(116 * F(y)) - 16, 500 * (F(x) - F(y)), 200 * (F(y) - F(z))};
    }

    private static double Linear(byte channel)
    {
        var value = channel / 255.0;
        return value <= 0.04045 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);
    }

    private static double F(double value)
    {
        return value > 0.008856 ? Math.Pow(value, 1.0 / 3) : (7.787 * value) + (16.0 / 116);
    }

    public bool Equals(StyleColor other)
    {
        return R == other.R && G == other.G && B == other.B && Indexed == other.Indexed;
    }

    public override bool Equals(object? obj)
    {
        return obj is StyleColor other && Equals(other);
    }

    public override int GetHashCode()
    {
        return (R << 24) | (G << 16) | (B << 8) | (byte) Indexed;
    }

    public override string ToString()
    {
        return Indexed > 0 ? $"#{R:X2}{G:X2}{B:X2}@{Indexed}" : $"#{R:X2}{G:X2}{B:X2}";
    }
}
