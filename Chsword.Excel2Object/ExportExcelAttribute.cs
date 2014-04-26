using System;

namespace Chsword.Excel2Object
{
    public class ExcelAttribute : Attribute
    {
        public ExcelAttribute(string name)
        {
            Title = name;
        }

        public int Order { get; set; }
        public string Title { get; set; }
    }
}