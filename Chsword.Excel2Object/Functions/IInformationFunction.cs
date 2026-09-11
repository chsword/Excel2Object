namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel information functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface IInformationFunction : IExcelFunction
{
    /// <summary>ISBLANK - TRUE when the cell is empty.</summary>
    [ExcelFunctionName("ISBLANK")]
    ColumnValue IsBlank(ColumnValue val);

    /// <summary>ISERROR - TRUE for any error value.</summary>
    [ExcelFunctionName("ISERROR")]
    ColumnValue IsError(ColumnValue val);

    /// <summary>ISERR - TRUE for any error value except #N/A.</summary>
    [ExcelFunctionName("ISERR")]
    ColumnValue IsErr(ColumnValue val);

    /// <summary>ISNA - TRUE for the #N/A error.</summary>
    [ExcelFunctionName("ISNA")]
    ColumnValue IsNa(ColumnValue val);

    /// <summary>ISNUMBER - TRUE when the value is a number.</summary>
    [ExcelFunctionName("ISNUMBER")]
    ColumnValue IsNumber(ColumnValue val);

    /// <summary>ISTEXT - TRUE when the value is text.</summary>
    [ExcelFunctionName("ISTEXT")]
    ColumnValue IsText(ColumnValue val);

    /// <summary>ISNONTEXT - TRUE when the value is not text.</summary>
    [ExcelFunctionName("ISNONTEXT")]
    ColumnValue IsNonText(ColumnValue val);

    /// <summary>ISLOGICAL - TRUE when the value is a logical value.</summary>
    [ExcelFunctionName("ISLOGICAL")]
    ColumnValue IsLogical(ColumnValue val);

    /// <summary>ISREF - TRUE when the value is a reference.</summary>
    [ExcelFunctionName("ISREF")]
    ColumnValue IsRef(ColumnValue val);

    /// <summary>ISFORMULA - TRUE when the cell contains a formula.</summary>
    [ExcelFunctionName("ISFORMULA", Future = true)]
    ColumnValue IsFormula(ColumnValue reference);

    /// <summary>ISEVEN - TRUE when the number is even.</summary>
    [ExcelFunctionName("ISEVEN")]
    ColumnValue IsEven(ColumnValue val);

    /// <summary>ISODD - TRUE when the number is odd.</summary>
    [ExcelFunctionName("ISODD")]
    ColumnValue IsOdd(ColumnValue val);

    /// <summary>NA - the #N/A error value.</summary>
    [ExcelFunctionName("NA")]
    ColumnValue Na();

    /// <summary>N - the value converted to a number.</summary>
    [ExcelFunctionName("N")]
    ColumnValue N(ColumnValue val);

    /// <summary>TYPE - a number identifying the type of a value.</summary>
    ColumnValue Type(ColumnValue val);

    /// <summary>ERROR.TYPE - a number identifying an error value.</summary>
    [ExcelFunctionName("ERROR.TYPE")]
    ColumnValue ErrorType(ColumnValue val);

    /// <summary>CELL - information about the formatting, location or contents of a cell.</summary>
    ColumnValue Cell(ColumnValue infoType);

    /// <summary>CELL - information about a given cell.</summary>
    ColumnValue Cell(ColumnValue infoType, ColumnValue reference);

    /// <summary>SHEET - the number of the sheet of a reference.</summary>
    [ExcelFunctionName("SHEET", Future = true)]
    ColumnValue Sheet();

    /// <summary>SHEET - the number of the sheet of a reference.</summary>
    [ExcelFunctionName("SHEET", Future = true)]
    ColumnValue Sheet(ColumnValue reference);

    /// <summary>SHEETS - the number of sheets of a reference.</summary>
    [ExcelFunctionName("SHEETS", Future = true)]
    ColumnValue Sheets();

    /// <summary>INFO - information about the current operating environment.</summary>
    ColumnValue Info(ColumnValue typeText);
}
