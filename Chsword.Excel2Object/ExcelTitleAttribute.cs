using System;
using System.ComponentModel;

namespace Chsword.Excel2Object
{
    /// <summary>
    /// On property ,it will be a column title
    /// On class ,it will be a sheet title
    /// </summary>
    public class ExcelTitleAttribute : Attribute
    {
        public ExcelTitleAttribute(string title)
        {
            Title = title;
        }

        public int Order { get; set; }
        public string Title { get; set; }

    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : ExcelTitleAttribute, IExcelHeaderStyle, IExcelCellStyle
    {
        public ExcelColumnAttribute(string title) : base(title)
        {

        }


        // Header
        [EditorBrowsable(EditorBrowsableState.Never)]
        public int HeaderAlignment { get; set; }
        public string HeaderFontFamily { get; set; }
        public double HeaderFontHeight { get; set; }
        public ExcelStyleColor HeaderFontColor { get; set; }
        public bool HeaderBold { get; set; }
        public bool HeaderItalic { get; set; }
        public bool HeaderStrikeout { get; set; }
        public bool HeaderUnderline { get; set; }

        // Cell
        [EditorBrowsable(EditorBrowsableState.Never)]
        public int CellAlignment { get; set; }
        public string CellFontFamily { get; set; }
        public double CellFontHeight { get; set; }
        public ExcelStyleColor CellFontColor { get; set; }
        public bool CellBold { get; set; }
        public bool CellItalic { get; set; }
        public bool CellStrikeout { get; set; }
        public bool CellUnderline { get; set; }
    }

    public interface IExcelHeaderStyle
    {
         int HeaderAlignment { get; set; }
         string HeaderFontFamily { get; set; }
         double HeaderFontHeight { get; set; }
         ExcelStyleColor HeaderFontColor { get; set; }
         bool HeaderBold { get; set; }
         bool HeaderItalic { get; set; }
         bool HeaderStrikeout { get; set; }
         bool HeaderUnderline { get; set; }
    }
    public interface IExcelCellStyle
    {
        int CellAlignment { get; set; }
        string CellFontFamily { get; set; }
        double CellFontHeight { get; set; }
        ExcelStyleColor CellFontColor { get; set; }
        bool CellBold { get; set; }
        bool CellItalic { get; set; }
        bool CellStrikeout { get; set; }
        bool CellUnderline { get; set; }
    }
    public enum ExcelStyleColor : short
    {
        //ColorNormal = 32767,
        Black = 8,
        Brown = 60,
        OliveGreen = 59,
        DarkGreen = 58,
        DarkTeal = 56,
        DarkBlue = 18,
        Indigo = 62,
        Grey80Percent = 63,
        Orange = 53,
        DarkYellow = 19,
        Green = 17,
        Teal = 21,
        Blue = 12,
        BlueGrey = 54,
        Grey50Percent = 23,
        Red = 10,
        LightOrange = 52,
        Lime = 50,
        SeaGreen = 57,
        Aqua = 49,
        LightBlue = 48,
        Violet = 20,
        Grey40Percent = 55,
        Pink = 14,
        Gold = 51,
        Yellow = 13,
        BrightGreen = 11,
        Turquoise = 15,
        DarkRed = 16,
        SkyBlue = 40,
        Plum = 61,
        Grey25Percent = 22,
        Rose = 45,
        LightYellow = 43,
        LightGreen = 42,
        LightTurquoise = 41,
        PaleBlue = 44,
        Lavender = 46,
        White = 9,
        CornflowerBlue = 24,
        LemonChiffon = 26,
        Maroon = 25,
        Orchid = 28,
        Coral = 29,
        RoyalBlue = 30,
        LightCornflowerBlue = 31,
        Tan = 47,
    }
}