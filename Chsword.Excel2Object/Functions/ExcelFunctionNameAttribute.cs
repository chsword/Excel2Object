namespace Chsword.Excel2Object.Functions;

/// <summary>
///     How a function interface method is spelled inside a formula, when the method name alone
///     cannot say it.
/// </summary>
/// <remarks>
///     Two things need spelling out. A name Excel writes with dots, such as <c>STDEV.P</c>, cannot be a
///     C# method name. And every function Excel gained after the 2007 file format was fixed is *stored*
///     under a <c>_xlfn.</c> prefix - Excel shows <c>IFS(...)</c> but the sheet holds
///     <c>_xlfn.IFS(...)</c>, and a workbook that stores the bare name shows <c>#NAME?</c> instead of a
///     result. <see cref="Future" /> marks those.
/// </remarks>
[AttributeUsage(AttributeTargets.Method)]
public sealed class ExcelFunctionNameAttribute : Attribute
{
    /// <summary>Prefix xlsx stores functions added after Excel 2007 under.</summary>
    public const string FuturePrefix = "_xlfn.";

    /// <summary>Prefix for the ones of those that only work on a worksheet, i.e. SORT and FILTER.</summary>
    public const string FutureWorksheetPrefix = "_xlfn._xlws.";

    public ExcelFunctionNameAttribute(string name)
    {
        Name = name;
    }

    /// <summary>The function name as Excel shows it, e.g. <c>STDEV.P</c>.</summary>
    public string Name { get; }

    /// <summary>Whether Excel gained this function after 2007, so that it is stored with a prefix.</summary>
    public bool Future { get; set; }

    /// <summary>Whether the function is one of the worksheet-only future functions (SORT, FILTER).</summary>
    public bool WorksheetOnly { get; set; }

    /// <summary>The name to write into the formula, prefix included.</summary>
    public string StoredName =>
        Future ? (WorksheetOnly ? FutureWorksheetPrefix : FuturePrefix) + Name : Name;
}
