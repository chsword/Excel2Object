using System;

namespace Chsword.Excel2Object
{
	[Obsolete(message:"instand of ExcelTitleAttribute")]
    public class ExcelAttribute : ExcelTitleAttribute
	{
        public ExcelAttribute(string name) :base(name)        {         }
    }
}