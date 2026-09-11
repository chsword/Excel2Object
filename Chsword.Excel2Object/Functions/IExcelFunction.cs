namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Marker for the interfaces whose methods are translated to Excel functions.
/// </summary>
/// <remarks>
///     Members of these interfaces are never executed: they exist only to be captured in a formula
///     expression tree, where each call becomes <c>NAME(arg,arg,...)</c>. The Excel name is the method
///     name in upper case unless <see cref="ExcelFunctionNameAttribute" /> says otherwise.
/// </remarks>
public interface IExcelFunction;
