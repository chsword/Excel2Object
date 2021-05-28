using Chsword.Excel2Object.Styles;

namespace Chsword.Excel2Object
{
    public interface IExcelHeaderStyle
    { 
        HorizontalAlignment HeaderAlignment { get; set; }
        string HeaderFontFamily { get; set; }
        double HeaderFontHeight { get; set; }
        ExcelStyleColor HeaderFontColor { get; set; }
        bool HeaderBold { get; set; }
        bool HeaderItalic { get; set; }
        bool HeaderStrikeout { get; set; }
        bool HeaderUnderline { get; set; }
    }
}