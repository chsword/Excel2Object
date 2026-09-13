using Chsword.Excel2Object.Internal;

namespace Chsword.Excel2Object.Styles;

/// <summary>Where the text sits in the cell from top to bottom.</summary>
public enum ExcelVerticalAlignment
{
    Default = 0,
    Top,
    Middle,
    Bottom
}

/// <summary>
///     How cells look, written the way a stylesheet writes it: colours as <c>#4472C4</c>, borders as
///     <c>1px solid #D0D0D0</c>. Only what you set is applied, so styles layer over each other - a
///     sheet-wide rule, then the column's own.
/// </summary>
/// <example>
///     <code>
/// options.Styles.Header(s => s.Bold().Background("#4472C4").Color("#FFF").Center());
/// options.Styles.Column("金额", s => s.Format("#,##0.00").Right());
/// options.Styles.EvenRows(s => s.Background("#F2F2F2"));
/// </code>
/// </example>
public class ExcelStyle
{
    internal StyleColor? TextColor { get; private set; }
    internal StyleColor? FillColor { get; private set; }
    internal bool? IsBold { get; private set; }
    internal bool? IsItalic { get; private set; }
    internal bool? IsUnderline { get; private set; }
    internal bool? IsStrikeout { get; private set; }
    internal string? FontName { get; private set; }
    internal double? FontPoints { get; private set; }
    internal HorizontalAlignment? Horizontal { get; private set; }
    internal ExcelVerticalAlignment? Vertical { get; private set; }
    internal bool? IsWrapped { get; private set; }
    internal string? NumberFormat { get; private set; }
    internal ExcelBorder? TopLine { get; private set; }
    internal ExcelBorder? RightLine { get; private set; }
    internal ExcelBorder? BottomLine { get; private set; }
    internal ExcelBorder? LeftLine { get; private set; }

    /// <summary>True when nothing was asked for, so the cell can keep the workbook's own look.</summary>
    internal bool IsEmpty =>
        TextColor == null && FillColor == null && IsBold == null && IsItalic == null && IsUnderline == null &&
        IsStrikeout == null && FontName == null && FontPoints == null && Horizontal == null && Vertical == null &&
        IsWrapped == null && NumberFormat == null && TopLine == null && RightLine == null && BottomLine == null && LeftLine == null;

    /// <summary>The colour of the text, as <c>#RRGGBB</c>, <c>#RGB</c>, or a palette colour.</summary>
    public ExcelStyle Color(string color)
    {
        TextColor = StyleColor.Parse(color);
        return this;
    }

    public ExcelStyle Color(ExcelStyleColor color)
    {
        TextColor = StyleColor.FromPalette(color);
        return this;
    }

    /// <summary>The colour the cell is filled with.</summary>
    public ExcelStyle Background(string color)
    {
        FillColor = StyleColor.Parse(color);
        return this;
    }

    public ExcelStyle Background(ExcelStyleColor color)
    {
        FillColor = StyleColor.FromPalette(color);
        return this;
    }

    public ExcelStyle Bold(bool on = true)
    {
        IsBold = on;
        return this;
    }

    public ExcelStyle Italic(bool on = true)
    {
        IsItalic = on;
        return this;
    }

    public ExcelStyle Underline(bool on = true)
    {
        IsUnderline = on;
        return this;
    }

    public ExcelStyle Strikeout(bool on = true)
    {
        IsStrikeout = on;
        return this;
    }

    public ExcelStyle FontFamily(string name)
    {
        FontName = name;
        return this;
    }

    /// <summary>The size of the text in points.</summary>
    public ExcelStyle FontSize(double points)
    {
        FontPoints = points;
        return this;
    }

    public ExcelStyle Align(HorizontalAlignment alignment)
    {
        Horizontal = alignment;
        return this;
    }

    public ExcelStyle Left()
    {
        return Align(HorizontalAlignment.Left);
    }

    public ExcelStyle Center()
    {
        return Align(HorizontalAlignment.Center);
    }

    public ExcelStyle Right()
    {
        return Align(HorizontalAlignment.Right);
    }

    public ExcelStyle VerticalAlign(ExcelVerticalAlignment alignment)
    {
        Vertical = alignment;
        return this;
    }

    /// <summary>Let the text wrap onto more lines instead of running past the cell.</summary>
    public ExcelStyle Wrap(bool on = true)
    {
        IsWrapped = on;
        return this;
    }

    /// <summary>
    ///     The number format, e.g. <c>#,##0.00</c>. On a date column a .NET format string such as
    ///     <c>yyyy-MM-dd</c> is translated into the Excel format that shows the same thing.
    /// </summary>
    public ExcelStyle Format(string format)
    {
        NumberFormat = format;
        return this;
    }

    /// <summary>All four sides, written as CSS writes them: <c>1px solid #D0D0D0</c>.</summary>
    public ExcelStyle Border(string border)
    {
        var parsed = ExcelBorder.Parse(border);
        TopLine = RightLine = BottomLine = LeftLine = parsed;
        return this;
    }

    public ExcelStyle BorderTop(string border)
    {
        TopLine = ExcelBorder.Parse(border);
        return this;
    }

    public ExcelStyle BorderRight(string border)
    {
        RightLine = ExcelBorder.Parse(border);
        return this;
    }

    public ExcelStyle BorderBottom(string border)
    {
        BottomLine = ExcelBorder.Parse(border);
        return this;
    }

    public ExcelStyle BorderLeft(string border)
    {
        LeftLine = ExcelBorder.Parse(border);
        return this;
    }

    /// <summary>
    ///     This style laid over <paramref name="under" />: everything this one asks for wins, and what it
    ///     says nothing about is left as the one underneath had it.
    /// </summary>
    internal ExcelStyle Over(ExcelStyle? under)
    {
        if (under == null || under.IsEmpty) return this;
        if (IsEmpty) return under;

        return new ExcelStyle
        {
            TextColor = TextColor ?? under.TextColor,
            FillColor = FillColor ?? under.FillColor,
            IsBold = IsBold ?? under.IsBold,
            IsItalic = IsItalic ?? under.IsItalic,
            IsUnderline = IsUnderline ?? under.IsUnderline,
            IsStrikeout = IsStrikeout ?? under.IsStrikeout,
            FontName = FontName ?? under.FontName,
            FontPoints = FontPoints ?? under.FontPoints,
            Horizontal = Horizontal ?? under.Horizontal,
            Vertical = Vertical ?? under.Vertical,
            IsWrapped = IsWrapped ?? under.IsWrapped,
            NumberFormat = NumberFormat ?? under.NumberFormat,
            TopLine = TopLine ?? under.TopLine,
            RightLine = RightLine ?? under.RightLine,
            BottomLine = BottomLine ?? under.BottomLine,
            LeftLine = LeftLine ?? under.LeftLine
        };
    }

    /// <summary>Everything this style asks for, as one string, so equal styles share one cell style.</summary>
    /// <remarks>
    ///     写成插值而非 <c>string.Join(string, params object[])</c>：.NET Framework 的该重载在首个元素
    ///     为 null 时直接返回空字符串，未设字体色的样式便会得到相同的键，进而共用同一个单元格样式。
    /// </remarks>
    internal string Key()
    {
        return $"{TextColor}|{FillColor}|{IsBold}|{IsItalic}|{IsUnderline}|{IsStrikeout}|{FontName}|" +
               $"{FontPoints}|{Horizontal}|{Vertical}|{IsWrapped}|{NumberFormat}|{TopLine}|{RightLine}|" +
               $"{BottomLine}|{LeftLine}";
    }
}
