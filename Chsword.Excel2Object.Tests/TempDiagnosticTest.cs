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
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Model> {new(), new()}, o =>
        {
            o.ExcelType = ExcelType.Xlsx;
            o.Styles.Cells(s => s.Background("#F2F2F2").Italic());
        })!;

        var xml = "(skipped)";
#if !NETFRAMEWORK
        using (var zip = new ZipArchive(new MemoryStream(bytes)))
        {
            var entry = zip.Entries.First(e => e.FullName.EndsWith("styles.xml"));
            xml = new StreamReader(entry.Open()).ReadToEnd();
        }
#endif

        var workbook = (XSSFWorkbook) WorkbookFactory.Create(new MemoryStream(bytes));
        var cell = workbook.GetSheetAt(0).GetRow(1).GetCell(0);
        var style = (XSSFCellStyle) cell.CellStyle;
        var font = (XSSFFont) style.GetFont(workbook);
        var rgb = style.FillForegroundColorColor?.RGB;

        Assert.Fail(
            $"DIAG hasF2F2F2={xml.Contains("F2F2F2")} fills={Fragment(xml, "<fills")} " +
            $"readFill={(rgb == null ? "null" : string.Join(",", rgb))} pattern={style.FillPattern} " +
            $"italic={font.IsItalic} styleCount={workbook.NumCellStyles} styleIndex={style.Index} " +
            $"styleType={style.GetType().Name} fontType={font.GetType().Name}");
    }

    private static string Fragment(string xml, string tag)
    {
        var i = xml.IndexOf(tag, System.StringComparison.Ordinal);
        if (i < 0) return "(无)";
        return xml.Substring(i, System.Math.Min(320, xml.Length - i)).Replace("\n", "");
    }
}
