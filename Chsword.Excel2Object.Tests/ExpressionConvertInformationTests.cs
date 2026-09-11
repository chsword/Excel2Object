using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertInformationTests : BaseFunctionTest
{
    [TestMethod]
    public void TypeChecks()
    {
        TestFunction(c => ExcelFunctions.Information.IsBlank(c["One"]), "ISBLANK(A4)");
        TestFunction(c => ExcelFunctions.Information.IsNumber(c["One"]), "ISNUMBER(A4)");
        TestFunction(c => ExcelFunctions.Information.IsText(c["One"]), "ISTEXT(A4)");
        TestFunction(c => ExcelFunctions.Information.IsEven(c["One"]), "ISEVEN(A4)");
        TestFunction(c => ExcelFunctions.Information.IsOdd(c["One"]), "ISODD(A4)");
    }

    [TestMethod]
    public void ErrorChecks()
    {
        TestFunction(c => ExcelFunctions.Information.IsError(c["One"]), "ISERROR(A4)");
        TestFunction(c => ExcelFunctions.Information.IsNa(c["One"]), "ISNA(A4)");
        TestFunction(c => ExcelFunctions.Information.ErrorType(c["One"]), "ERROR.TYPE(A4)");
        TestFunction(c => ExcelFunctions.Information.Na(), "NA()");
    }

    [TestMethod]
    public void CombinedWithIf()
    {
        TestFunction(
            c => ExcelFunctions.Condition.If(ExcelFunctions.Information.IsBlank(c["One"]), 0, c["One"]),
            "IF(ISBLANK(A4),0,A4)");
    }
}
