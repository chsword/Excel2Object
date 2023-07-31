namespace Chsword.Excel2Object.Styles;

public interface IExcelCellStyle
{
    HorizontalAlignment CellAlignment { get; set; }
    bool CellBold { get; set; }
    ExcelStyleColor CellFontColor { get; set; }
    string? CellFontFamily { get; set; }
    double CellFontHeight { get; set; }
    bool CellItalic { get; set; }
    bool CellStrikeout { get; set; }
    bool CellUnderline { get; set; }

    string? Format { get; set; }
}