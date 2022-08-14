namespace Chsword.Excel2Object;

[Obsolete("instand of ExcelTitleAttribute", true)]
public class ExcelAttribute : ExcelTitleAttribute
{
    public ExcelAttribute(string name) : base(name)
    {
    }
}