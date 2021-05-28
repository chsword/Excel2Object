namespace Chsword.Excel2Object.Styles
{
    public interface IExcelHeaderStyle
    {
        HorizontalAlignment HeaderAlignment { get; set; }
        bool HeaderBold { get; set; }
        ExcelStyleColor HeaderFontColor { get; set; }
        string HeaderFontFamily { get; set; }
        double HeaderFontHeight { get; set; }
        bool HeaderItalic { get; set; }
        bool HeaderStrikeout { get; set; }
        bool HeaderUnderline { get; set; }
    }
}