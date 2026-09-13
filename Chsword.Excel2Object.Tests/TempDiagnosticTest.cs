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
        // 绕开导出流程，直接看 DSL 与解析器的中间结果
        var sheet = new Chsword.Excel2Object.Styles.ExcelStyleSheet();
        sheet.EvenRows(s => s.Background("#F2F2F2"));

        var even = sheet.EvenRowsStyle;
        var parsed = Chsword.Excel2Object.Internal.StyleColor.Parse("#F2F2F2");

        var column = new Chsword.Excel2Object.ExcelColumn {Title = "城市", Type = typeof(string)};
        var odd = Chsword.Excel2Object.Internal.StyleResolver.Resolve(column, sheet, 1);
        var evenResolved = Chsword.Excel2Object.Internal.StyleResolver.Resolve(column, sheet, 2);

        Assert.Fail(
            $"DIAG parse={parsed} evenNull={even == null} evenIsEmpty={even?.IsEmpty} " +
            $"evenFill={even?.FillColor?.ToString() ?? "null"} evenKey=[{even?.Key()}] " +
            $"oddResolvedNull={odd.Style == null} evenResolvedNull={evenResolved.Style == null} " +
            $"evenResolvedKey=[{evenResolved.Style?.Key()}]");
    }

    private static string Fragment(string xml, string tag)
    {
        var i = xml.IndexOf(tag, System.StringComparison.Ordinal);
        if (i < 0) return "(无)";
        return xml.Substring(i, System.Math.Min(320, xml.Length - i)).Replace("\n", "");
    }
}
