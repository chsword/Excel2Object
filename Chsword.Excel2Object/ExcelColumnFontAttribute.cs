using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chsword.Excel2Object
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnFontAttribute : Attribute
    {
        public ExcelColumnFontAttribute(string fontName, double fontHeightInPoints, short color)
        {
            FontName = fontName;
            FontHeightInPoints = fontHeightInPoints;
            Color = color;
        }

        public ExcelColumnFontAttribute(string fontName)
        {
            FontName = fontName;
        }

        public string FontName { get; set; }
        public double FontHeightInPoints { get; set; } = 16;
        /// <summary>
        /// BLACK=8
        /// </summary>
        public short Color { get; set; } = 8;
        public bool IsBold { get; set; } = false;
    }
}
