using System.Collections.Concurrent;
using System.Text;
using NPOI.HSSF.UserModel;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     Turns a .NET date/time format string into the Excel number format that displays a date cell the
///     same way, so an <c>[ExcelColumn(Format = "yyyy-MM-dd HH:mm:ss")]</c> written for the old text
///     export keeps working now that the column is a real date cell.
/// </summary>
/// <remarks>
///     The two languages are close: <c>d</c>, <c>y</c>, <c>s</c> and the separators mean the same, while
///     <c>M</c>/<c>H</c> become <c>m</c>/<c>h</c>, <c>tt</c> becomes <c>AM/PM</c> and <c>fff</c> becomes
///     <c>000</c>. Anything Excel cannot show - time zones, eras - is dropped. Literal text that .NET copies
///     through unchanged (<c>年</c>, <c>T</c>) is quoted, because Excel gives some letters a meaning of
///     their own. A format Excel already knows, such as its builtin <c>m/d/yy</c> or the Excel spelling
///     <c>yyyy-mm-dd</c>, comes out unchanged; so do the bracketed tags Excel formats carry
///     (<c>[$-409]</c>, <c>[Red]</c>, elapsed <c>[h]</c>).
/// </remarks>
internal static class ExcelDateFormat
{
    /// <summary>The format a date column gets when its [ExcelColumn] names none.</summary>
    public const string Default = "yyyy-mm-dd hh:mm:ss";

    /// <summary>The .NET spelling of what <see cref="Default" /> displays.</summary>
    public const string IsoDateTime = "yyyy-MM-dd HH:mm:ss";

    public const string IsoDate = "yyyy-MM-dd";
    public const string IsoTime = "HH:mm:ss";

    /// <summary>
    ///     .NET's single-letter standard formats. Their meaning depends on the current culture, so Excel's
    ///     locale-aware builtins are used where one exists and the invariant patterns otherwise.
    /// </summary>
    private static readonly Dictionary<char, string> StandardFormats = new()
    {
        ['d'] = "m/d/yy",
        ['D'] = "dddd, mmmm d, yyyy",
        ['f'] = "dddd, mmmm d, yyyy h:mm AM/PM",
        ['F'] = "dddd, mmmm d, yyyy h:mm:ss AM/PM",
        ['g'] = "m/d/yy h:mm",
        ['G'] = "m/d/yy h:mm:ss",
        ['m'] = "mmmm d",
        ['M'] = "mmmm d",
        ['o'] = "yyyy-mm-dd\"T\"hh:mm:ss.000",
        ['O'] = "yyyy-mm-dd\"T\"hh:mm:ss.000",
        ['r'] = "ddd, dd mmm yyyy hh:mm:ss \"GMT\"",
        ['R'] = "ddd, dd mmm yyyy hh:mm:ss \"GMT\"",
        ['s'] = "yyyy-mm-dd\"T\"hh:mm:ss",
        ['t'] = "h:mm",
        ['T'] = "h:mm:ss",
        ['u'] = "yyyy-mm-dd hh:mm:ss\"Z\"",
        ['y'] = "mmmm yyyy",
        ['Y'] = "mmmm yyyy"
    };

    /// <summary>
    ///     Characters Excel takes literally inside a number format, so they need no quoting (digits too, so
    ///     an Excel-style <c>ss.0</c> survives). The section separator and text placeholder are included
    ///     for formats copied from Excel's Format Cells dialog, which end in <c>;@</c>.
    /// </summary>
    private const string LiteralChars = " -/:,.()$;@";

    /// <summary>
    ///     Formats come from attributes, so there are few distinct ones, while the translation is asked for
    ///     once per cell.
    /// </summary>
    private static readonly ConcurrentDictionary<string, string> Translated = new();

    public static string ToExcel(string? format)
    {
        if (string.IsNullOrEmpty(format)) return Default;
        return Translated.GetOrAdd(format!, Translate);
    }

    /// <summary>
    ///     True when the format is already in Excel's spelling and so has no .NET rendering: .NET would
    ///     read its <c>mm</c> as minutes and its <c>hh</c> as the 12-hour clock. Where the two languages
    ///     spell a format identically (<c>yyyy</c>, <c>dd</c>) the answer is true as well, and nothing is
    ///     lost by treating it as Excel's.
    /// </summary>
    public static bool IsExcelSpelling(string format)
    {
        return ToExcel(format) == format;
    }

    private static string Translate(string format)
    {
        if (HSSFDataFormat.GetBuiltinFormat(format) != -1) return format;
        if (format.Length == 1)
            return StandardFormats.TryGetValue(format[0], out var standard) ? standard : Quote(format);

        var sb = new StringBuilder(format.Length + 8);
        var literal = new StringBuilder();
        var i = 0;
        while (i < format.Length)
        {
            var c = format[i];
            var run = RunLength(format, i);

            // Excel's own AM/PM markers pass through, otherwise the M in them would turn into a month
            var markerLength = StartsWithIgnoreCase(format, i, "AM/PM") ? 5 :
                StartsWithIgnoreCase(format, i, "A/P") ? 3 : 0;
            if (markerLength > 0)
            {
                Flush(sb, literal).Append(format, i, markerLength);
                i += markerLength;
                continue;
            }

            switch (c)
            {
                case 'd':
                case 'D':
                    Flush(sb, literal).Append('d', Math.Min(run, 4));
                    break;
                case 'M':
                    // .NET's month never goes past the full name; a fifth m would be Excel's initial
                    Flush(sb, literal).Append('m', Math.Min(run, 4));
                    break;
                case 'y':
                case 'Y':
                    // .NET "y" is the year without a leading zero, which Excel has no equivalent of
                    Flush(sb, literal).Append(run <= 2 ? "yy" : "yyyy");
                    break;
                case 'H':
                case 'h':
                    Flush(sb, literal).Append('h', Math.Min(run, 2));
                    break;
                case 'm':
                    // minutes in .NET, never more than two; a longer run is Excel's own month spelling
                    Flush(sb, literal).Append('m', Math.Min(run, 5));
                    break;
                case 's':
                case 'S':
                    Flush(sb, literal).Append('s', Math.Min(run, 2));
                    break;
                case 'f':
                case 'F':
                    // Excel shows at most three decimals of a second
                    Flush(sb, literal).Append('0', Math.Min(run, 3));
                    break;
                case 't':
                    Flush(sb, literal).Append(run == 1 ? "A/P" : "AM/PM");
                    break;
                case 'z':
                case 'K':
                case 'g':
                    // time zone offsets and eras have no Excel counterpart
                    break;
                case '%':
                    // marks a lone custom specifier, e.g. "%d"; nothing to keep
                    run = 1;
                    break;
                case '\\':
                    // an escaped character is a literal in both languages
                    if (i + 1 < format.Length)
                        literal.Append(format[i + 1]);
                    run = 2;
                    break;
                case '\'':
                case '"':
                {
                    var end = format.IndexOf(c, i + 1);
                    if (end < 0) end = format.Length;
                    literal.Append(format, i + 1, end - i - 1);
                    run = end - i + 1;
                    break;
                }
                case '[':
                {
                    // a tag of Excel's own - locale, colour, elapsed time - copied as it is
                    var end = format.IndexOf(']', i + 1);
                    if (end < 0) end = format.Length - 1;
                    Flush(sb, literal).Append(format, i, end - i + 1);
                    run = end - i + 1;
                    break;
                }
                default:
                    if (LiteralChars.IndexOf(c) >= 0 || char.IsDigit(c))
                        Flush(sb, literal).Append(c, run);
                    else
                        literal.Append(c, run);
                    break;
            }

            i += run;
        }

        Flush(sb, literal);
        // a dropped time zone or era leaves its separator behind
        return sb.ToString().Trim();
    }

    /// <summary>Which of date and time an Excel date format displays.</summary>
    [Flags]
    public enum Parts
    {
        /// <summary>Nothing a calendar date could stand for, e.g. elapsed time <c>[h]:mm</c>.</summary>
        None = 0,
        Date = 1,
        Time = 2
    }

    /// <summary>
    ///     Reads an Excel date format for what it shows, so a date cell can be rendered as text at the
    ///     right length: a date, a time of day, or both.
    /// </summary>
    public static Parts PartsShown(string excelFormat)
    {
        var date = false;
        var time = false;
        var month = false;
        var i = 0;
        while (i < excelFormat.Length)
        {
            var c = excelFormat[i];
            switch (c)
            {
                case '"':
                {
                    var end = excelFormat.IndexOf('"', i + 1);
                    i = end < 0 ? excelFormat.Length : end + 1;
                    continue;
                }
                case '\\':
                    i += 2;
                    continue;
                case '[':
                {
                    var end = excelFormat.IndexOf(']', i + 1);
                    if (end < 0) end = excelFormat.Length - 1;
                    // [h], [mm], [ss] count hours, minutes or seconds elapsed - a number, not a date
                    var elapsed = end > i + 1;
                    for (var j = i + 1; j < end && elapsed; j++)
                        elapsed = "hms".IndexOf(char.ToLowerInvariant(excelFormat[j])) >= 0;
                    if (elapsed) return Parts.None;
                    i = end + 1;
                    continue;
                }
                case ';':
                    // only the first section applies to a positive number, and a date is one
                    i = excelFormat.Length;
                    continue;
            }

            switch (char.ToLowerInvariant(c))
            {
                case 'y':
                case 'd':
                    date = true;
                    break;
                case 'h':
                case 's':
                    time = true;
                    break;
                case 'm':
                    month = true;
                    break;
            }

            i++;
        }

        // an m is a month unless it keeps company with hours or seconds
        if (month && !time) date = true;
        return (date ? Parts.Date : Parts.None) | (time ? Parts.Time : Parts.None);
    }

    private static int RunLength(string format, int start)
    {
        var c = format[start];
        var end = start + 1;
        while (end < format.Length && format[end] == c) end++;
        return end - start;
    }

    private static bool StartsWithIgnoreCase(string s, int index, string value)
    {
        return index + value.Length <= s.Length &&
               string.Compare(s, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    /// <summary>Writes any pending literal text, quoted, before the next format token.</summary>
    private static StringBuilder Flush(StringBuilder sb, StringBuilder literal)
    {
        if (literal.Length == 0) return sb;
        sb.Append(Quote(literal.ToString()));
        literal.Clear();
        return sb;
    }

    private static string Quote(string text)
    {
        // Excel has no way to escape a quote inside quoted text, but does accept a backslash-escaped one
        return "\"" + text.Replace("\"", "\"\\\"\"") + "\"";
    }
}
