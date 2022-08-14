using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : ExcelTitleAttribute, IExcelHeaderStyle, IExcelCellStyle
{
    public ExcelColumnAttribute(string title) : base(title)
    {
    }

    // Cell

    public HorizontalAlignment CellAlignment { get; set; }
    public bool CellBold { get; set; }
    public ExcelStyleColor CellFontColor { get; set; }
    public string? CellFontFamily { get; set; }
    public double CellFontHeight { get; set; }
    public bool CellItalic { get; set; }
    public bool CellStrikeout { get; set; }
    public bool CellUnderline { get; set; }


    // Header
    public HorizontalAlignment HeaderAlignment { get; set; }
    public bool HeaderBold { get; set; }
    public ExcelStyleColor HeaderFontColor { get; set; }
    public string? HeaderFontFamily { get; set; }
    public double HeaderFontHeight { get; set; }
    public bool HeaderItalic { get; set; }
    public bool HeaderStrikeout { get; set; }
    public bool HeaderUnderline { get; set; }
}