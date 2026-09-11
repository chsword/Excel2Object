using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class ExpressionConvertEngineeringTests : BaseFunctionTest
{
    [TestMethod]
    public void BaseConversions()
    {
        TestFunction(c => ExcelFunctions.Engineering.Dec2Hex(c["One"]), "DEC2HEX(A4)");
        TestFunction(c => ExcelFunctions.Engineering.Hex2Dec(c["One"]), "HEX2DEC(A4)");
        TestFunction(c => ExcelFunctions.Engineering.Bin2Dec(c["One"]), "BIN2DEC(A4)");
    }

    [TestMethod]
    public void BitwiseAndUnits()
    {
        TestFunction(c => ExcelFunctions.Engineering.BitAnd(c["One"], 6), "_xlfn.BITAND(A4,6)");
        TestFunction(c => ExcelFunctions.Engineering.Convert(c["One"], "km", "mi"), "CONVERT(A4,\"km\",\"mi\")");
        TestFunction(c => ExcelFunctions.Engineering.Delta(c["One"], c["Two"]), "DELTA(A4,B4)");
    }
}
