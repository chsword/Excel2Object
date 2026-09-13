using System.Collections.Generic;
using System.IO;
#if !NETFRAMEWORK
using System.IO.Compression;
#endif
using System.Linq;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>临时诊断：比较各目标框架下写出的 styles.xml 与读回的结果。用完即删。</summary>
[TestClass]
public class TempDiagnosticTest
{
    public class Model
    {
        [ExcelTitle("城市")] public string City { get; set; } = "北京";
    }

    [TestMethod]
    public void DumpStyles()
    {
        // 完全复刻 RowsStripe：只设偶数行底色
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Model> {new(), new(), new()}, o =>
        {
            o.ExcelType = ExcelType.Xlsx;
            o.Styles.EvenRows(s => s.Background("#F2F2F2"));
        })!;

        var workbook = (XSSFWorkbook) WorkbookFactory.Create(new MemoryStream(bytes));
        var sheet = workbook.GetSheetAt(0);
        var dump = "";
        for (var r = 0; r <= 3; r++)
        {
            var style = (XSSFCellStyle) sheet.GetRow(r).GetCell(0).CellStyle;
            var rgb = style.FillForegroundColorColor?.RGB;
            dump += $"row{r}:idx={style.Index},pattern={style.FillPattern}," +
                    $"fill={(rgb == null ? "null" : string.Join("-", rgb))},fmt={style.GetDataFormatString()} ";
        }

        Assert.Fail($"DIAG styleCount={workbook.NumCellStyles} {dump}");
    }

    private static string Fragment(string xml, string tag)
    {
        var i = xml.IndexOf(tag, System.StringComparison.Ordinal);
        if (i < 0) return "(无)";
        return xml.Substring(i, System.Math.Min(320, xml.Length - i)).Replace("\n", "");
    }
}
