using System;

namespace Chsword.Excel2Object
{
    [Obsolete("", true)]
    internal class Font
    {
        public Font(string fontName, double fontHeightInPoints, short color, bool isBold)
        {
            FontName = fontName;
            FontHeightInPoints = fontHeightInPoints;
            Color = color;
            IsBold = IsBold;
        }

        public short Color { get; set; }
        public double FontHeightInPoints { get; set; }

        public string FontName { get; set; }
        public bool IsBold { get; set; }
    }
}