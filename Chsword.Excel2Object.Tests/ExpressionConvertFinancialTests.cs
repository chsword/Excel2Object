using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertFinancialTests : BaseFunctionTest
{
    [TestMethod]
    public void Annuities()
    {
        TestFunction(c => ExcelFunctions.Financial.Pmt(c["One"], c["Two"], c["Three"]), "PMT(A4,B4,C4)");
        TestFunction(c => ExcelFunctions.Financial.Pv(c["One"], c["Two"], c["Three"]), "PV(A4,B4,C4)");
        TestFunction(c => ExcelFunctions.Financial.Fv(c["One"], c["Two"], c["Three"]), "FV(A4,B4,C4)");
        TestFunction(c => ExcelFunctions.Financial.NPer(c["One"], c["Two"], c["Three"]), "NPER(A4,B4,C4)");
    }

    [TestMethod]
    public void CashFlows()
    {
        TestFunction(c => ExcelFunctions.Financial.Npv(c["One"], c.Matrix("Two", 2, "Two", 9)),
            "NPV(A4,B2:B9)");
        TestFunction(c => ExcelFunctions.Financial.Irr(c.Matrix("Two", 2, "Two", 9)), "IRR(B2:B9)");
        TestFunction(
            c => ExcelFunctions.Financial.XIrr(c.Matrix("Two", 2, "Two", 9), c.Matrix("Three", 2, "Three", 9)),
            "XIRR(B2:B9,C2:C9)");
    }

    [TestMethod]
    public void Depreciation()
    {
        TestFunction(c => ExcelFunctions.Financial.Sln(c["One"], c["Two"], c["Three"]), "SLN(A4,B4,C4)");
        TestFunction(c => ExcelFunctions.Financial.Ddb(c["One"], c["Two"], c["Three"], 1), "DDB(A4,B4,C4,1)");
    }
}
