using Chsword.Excel2Object.Styles;
using NPOI.OOXML.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using HorizontalAlignment = Chsword.Excel2Object.Styles.HorizontalAlignment;

namespace Chsword.Excel2Object.Internal;

/// <summary>
///     Turns the styles an export asks for into the workbook's own cell styles, one per distinct look.
/// </summary>
/// <remarks>
///     A workbook holds a limited number of styles and fonts - .xls stops at 4000 and 512 - so a striped
///     table of a hundred thousand rows must not create a style per cell. Everything here is keyed by what
///     the style asks for, so cells that look alike share one.
/// </remarks>
internal sealed class CellStyleFactory
{
    private readonly IWorkbook _workbook;
    private readonly Dictionary<string, ICellStyle> _styles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, IFont> _fonts = new(StringComparer.Ordinal);

    public CellStyleFactory(IWorkbook workbook)
    {
        _workbook = workbook;
    }

    /// <summary>
    ///     The cell style for a look, or null when there is nothing to apply and the cell can keep the
    ///     workbook's own.
    /// </summary>
    /// <param name="style">What the export asked for, if anything.</param>
    /// <param name="format">
    ///     The number format the cell needs whatever the style says - <c>@</c> for text, the translated
    ///     format of a date - which the style's own <c>Format</c> overrides when it has one.
    /// </param>
    public ICellStyle? Get(ExcelStyle? style, string? format = null)
    {
        if ((style == null || style.IsEmpty) && format == null) return null;

        var key = (style?.Key() ?? "") + "" + format;
        if (_styles.TryGetValue(key, out var cached)) return cached;

        var cellStyle = _workbook.CreateCellStyle();
        if (style != null && !style.IsEmpty) Apply(cellStyle, style);

        // the caller's format is the cell's own need - "@" for text, a date's translated format - and wins
        var numberFormat = format ?? style?.NumberFormat;
        if (numberFormat != null)
            // GetFormat hands back the builtin index when there is one and registers the format otherwise
            cellStyle.DataFormat = _workbook.CreateDataFormat().GetFormat(numberFormat);

        _styles[key] = cellStyle;
        return cellStyle;
    }

    private void Apply(ICellStyle cellStyle, ExcelStyle style)
    {
        var font = Font(style);
        if (font != null) cellStyle.SetFont(font);

        if (style.FillColor != null)
        {
            SetFill(cellStyle, style.FillColor.Value);
            cellStyle.FillPattern = FillPattern.SolidForeground;
        }

        if (style.Horizontal != null && style.Horizontal != HorizontalAlignment.General)
            cellStyle.Alignment = (NPOI.SS.UserModel.HorizontalAlignment) style.Horizontal;
        if (style.Vertical != null && style.Vertical != ExcelVerticalAlignment.Default)
            cellStyle.VerticalAlignment = style.Vertical switch
            {
                ExcelVerticalAlignment.Top => VerticalAlignment.Top,
                ExcelVerticalAlignment.Middle => VerticalAlignment.Center,
                _ => VerticalAlignment.Bottom
            };
        if (style.IsWrapped != null) cellStyle.WrapText = style.IsWrapped.Value;

        SetBorder(cellStyle, style.TopLine, BorderSide.Top);
        SetBorder(cellStyle, style.RightLine, BorderSide.Right);
        SetBorder(cellStyle, style.BottomLine, BorderSide.Bottom);
        SetBorder(cellStyle, style.LeftLine, BorderSide.Left);
    }

    private IFont? Font(ExcelStyle style)
    {
        if (style.TextColor == null && style.IsBold == null && style.IsItalic == null && style.IsUnderline == null &&
            style.IsStrikeout == null && style.FontName == null && style.FontPoints == null)
            return null;

        var key = string.Join("|", style.TextColor, style.IsBold, style.IsItalic, style.IsUnderline,
            style.IsStrikeout, style.FontName, style.FontPoints);
        if (_fonts.TryGetValue(key, out var cached)) return cached;

        var font = _workbook.CreateFont();
        if (style.FontName != null) font.FontName = style.FontName;
        // Leave the height alone unless asked for one: forcing a default here would shrink every column
        // that only asks for a colour.
        if (style.FontPoints != null) font.FontHeightInPoints = style.FontPoints.Value;
        if (style.IsBold != null) font.IsBold = style.IsBold.Value;
        if (style.IsItalic != null) font.IsItalic = style.IsItalic.Value;
        if (style.IsStrikeout != null) font.IsStrikeout = style.IsStrikeout.Value;
        if (style.IsUnderline != null)
            font.Underline = style.IsUnderline.Value ? FontUnderlineType.Single : FontUnderlineType.None;
        if (style.TextColor != null) SetFontColor(font, style.TextColor.Value);

        _fonts[key] = font;
        return font;
    }

    private static void SetFontColor(IFont font, StyleColor color)
    {
        if (font is XSSFFont xssf && !color.IsFromPalette)
            xssf.SetColor(Color(color));
        else
            font.Color = color.PaletteIndex();
    }

    private static void SetFill(ICellStyle cellStyle, StyleColor color)
    {
        if (cellStyle is XSSFCellStyle xssf && !color.IsFromPalette)
            xssf.SetFillForegroundColor(Color(color));
        else
            cellStyle.FillForegroundColor = color.PaletteIndex();
    }

    private enum BorderSide
    {
        Top,
        Right,
        Bottom,
        Left
    }

    private static void SetBorder(ICellStyle cellStyle, ExcelBorder? border, BorderSide side)
    {
        if (border == null) return;
        var line = border.Value.Line;
        var color = border.Value.Color;
        // a colour picked from the palette keeps its index, which is what it meant before .xlsx could
        // hold the colour itself
        var xssf = color?.IsFromPalette == true ? null : cellStyle as XSSFCellStyle;

        switch (side)
        {
            case BorderSide.Top:
                cellStyle.BorderTop = line;
                if (color == null) break;
                if (xssf != null) xssf.SetTopBorderColor(Color(color.Value));
                else cellStyle.TopBorderColor = color.Value.PaletteIndex();
                break;
            case BorderSide.Right:
                cellStyle.BorderRight = line;
                if (color == null) break;
                if (xssf != null) xssf.SetRightBorderColor(Color(color.Value));
                else cellStyle.RightBorderColor = color.Value.PaletteIndex();
                break;
            case BorderSide.Bottom:
                cellStyle.BorderBottom = line;
                if (color == null) break;
                if (xssf != null) xssf.SetBottomBorderColor(Color(color.Value));
                else cellStyle.BottomBorderColor = color.Value.PaletteIndex();
                break;
            default:
                cellStyle.BorderLeft = line;
                if (color == null) break;
                if (xssf != null) xssf.SetLeftBorderColor(Color(color.Value));
                else cellStyle.LeftBorderColor = color.Value.PaletteIndex();
                break;
        }
    }

    /// <summary>The colour itself, which .xlsx stores as written rather than as a palette index.</summary>
    private static XSSFColor Color(StyleColor color)
    {
        var value = new XSSFColor(new DefaultIndexedColorMap());
        value.SetRgb(color.Rgb());
        return value;
    }
}
