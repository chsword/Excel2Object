using System;

namespace Chsword.Excel2Object
{
    public class ExcelTitleAttribute : Attribute
    {
        public ExcelTitleAttribute(string title)
        {
            Title = title;
        }

        public int Order { get; set; }
        public string Title { get; set; }
    }
}