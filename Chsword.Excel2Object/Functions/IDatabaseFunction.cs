namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel database functions. Every one of them takes the table, the field to work on and a range
///     holding the criteria. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface IDatabaseFunction : IExcelFunction
{
    /// <summary>DSUM - sums the values of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DSUM")]
    ColumnValue DSum(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DAVERAGE - averages the values of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DAVERAGE")]
    ColumnValue DAverage(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DCOUNT - counts the numeric cells of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DCOUNT")]
    ColumnValue DCount(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DCOUNTA - counts the non empty cells of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DCOUNTA")]
    ColumnValue DCountA(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DGET - the single value of a field for the one row that meets the criteria.</summary>
    [ExcelFunctionName("DGET")]
    ColumnValue DGet(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DMAX - the largest value of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DMAX")]
    ColumnValue DMax(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DMIN - the smallest value of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DMIN")]
    ColumnValue DMin(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DPRODUCT - multiplies the values of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DPRODUCT")]
    ColumnValue DProduct(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DSTDEV - the sample standard deviation of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DSTDEV")]
    ColumnValue DStDev(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DSTDEVP - the population standard deviation of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DSTDEVP")]
    ColumnValue DStDevP(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DVAR - the sample variance of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DVAR")]
    ColumnValue DVar(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);

    /// <summary>DVARP - the population variance of a field for the rows that meet the criteria.</summary>
    [ExcelFunctionName("DVARP")]
    ColumnValue DVarP(ColumnMatrix database, ColumnValue field, ColumnMatrix criteria);
}
