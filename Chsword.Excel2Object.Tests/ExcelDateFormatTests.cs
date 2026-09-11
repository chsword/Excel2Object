using System.Collections.Generic;
using Chsword.Excel2Object.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     [ExcelColumn(Format = ...)] takes a .NET date format string; the exporter needs the Excel number
///     format that displays the same thing.
/// </summary>
[TestClass]
public class ExcelDateFormatTests
{
    [DataTestMethod]
    [DataRow("yyyy-MM-dd HH:mm:ss", "yyyy-mm-dd hh:mm:ss")]
    [DataRow("yyyy-MM-dd", "yyyy-mm-dd")]
    [DataRow("yyyy/M/d", "yyyy/m/d")]
    [DataRow("dd/MM/yyyy", "dd/mm/yyyy")]
    [DataRow("HH:mm", "hh:mm")]
    [DataRow("HH:mm:ss.fff", "hh:mm:ss.000")]
    [DataRow("HH:mm:ss.fffffff", "hh:mm:ss.000")]
    [DataRow("hh:mm tt", "hh:mm AM/PM")]
    [DataRow("h:mm t", "h:mm A/P")]
    [DataRow("ddd, dd MMM yyyy", "ddd, dd mmm yyyy")]
    [DataRow("dddd MMMM", "dddd mmmm")]
    [DataRow("MMMMM", "mmmm")]
    [DataRow("YYYY-MM-DD", "yyyy-mm-dd")]
    [DataRow("yy", "yy")]
    [DataRow("yyy", "yyyy")]
    [DataRow("yyyyy", "yyyy")]
    public void DotNetTokensBecomeTheirExcelCounterparts(string dotNet, string excel)
    {
        Assert.AreEqual(excel, ExcelDateFormat.ToExcel(dotNet));
    }

    /// <summary>
    ///     .NET copies characters it does not recognise into the output; Excel gives some letters a
    ///     meaning of their own, so literal text is quoted.
    /// </summary>
    [DataTestMethod]
    [DataRow("yyyy年MM月dd日", "yyyy\"年\"mm\"月\"dd\"日\"")]
    [DataRow("yyyy年M月d日 HH:mm", "yyyy\"年\"m\"月\"d\"日\" hh:mm")]
    [DataRow("yyyy-MM-ddTHH:mm:ss", "yyyy-mm-dd\"T\"hh:mm:ss")]
    [DataRow("yyyy-MM-ddTHH:mm:ssZ", "yyyy-mm-dd\"T\"hh:mm:ss\"Z\"")]
    [DataRow("'week' dd", "\"week\" dd")]
    [DataRow("\"at\" HH:mm", "\"at\" hh:mm")]
    [DataRow(@"yyyy\-MM", "yyyy\"-\"mm")]
    [DataRow("%d", "d")]
    public void LiteralTextIsQuoted(string dotNet, string excel)
    {
        Assert.AreEqual(excel, ExcelDateFormat.ToExcel(dotNet));
    }

    [DataTestMethod]
    [DataRow("yyyy-MM-dd HH:mm zzz", "yyyy-mm-dd hh:mm")]
    [DataRow("yyyy-MM-ddTHH:mm:ssK", "yyyy-mm-dd\"T\"hh:mm:ss")]
    [DataRow("gg yyyy", "yyyy")]
    public void WhatExcelCannotShowIsDropped(string dotNet, string excel)
    {
        Assert.AreEqual(excel, ExcelDateFormat.ToExcel(dotNet));
    }

    /// <summary>A format Excel already understands comes out unchanged.</summary>
    [DataTestMethod]
    [DataRow("m/d/yy")]
    [DataRow("h:mm AM/PM")]
    [DataRow("[h]:mm:ss")]
    [DataRow("yyyy-mm-dd")]
    [DataRow("yyyy-mm-dd hh:mm:ss")]
    [DataRow("hh:mm am/pm")]
    [DataRow("yyyy\"年\"m\"月\"d\"日\"")]
    [DataRow("[$-409]d-mmm-yy")]
    [DataRow("[Red]yyyy")]
    [DataRow("yyyy/m/d;@")]
    [DataRow("ss.0")]
    public void ExcelFormatsPassThrough(string excel)
    {
        Assert.AreEqual(excel, ExcelDateFormat.ToExcel(excel));
        Assert.IsTrue(ExcelDateFormat.IsExcelSpelling(excel));
    }

    [DataTestMethod]
    [DataRow("yyyy-MM-dd")]
    [DataRow("d")]
    [DataRow("HH:mm")]
    public void ADotNetFormatIsNotExcelsSpelling(string dotNet)
    {
        Assert.IsFalse(ExcelDateFormat.IsExcelSpelling(dotNet));
    }

    /// <summary>The importer renders a date cell as text at the length its format shows.</summary>
    [DataTestMethod]
    [DataRow("yyyy-mm-dd hh:mm:ss", "date time")]
    [DataRow("m/d/yy h:mm", "date time")]
    [DataRow("yyyy-mm-dd", "date")]
    [DataRow("mmm-yy", "date")]
    [DataRow("mmmm", "date")]
    [DataRow("yyyy\"年\"m\"月\"d\"日\"", "date")]
    [DataRow("[$-409]d-mmm-yy;@", "date")]
    [DataRow("h:mm", "time")]
    [DataRow("hh:mm:ss AM/PM", "time")]
    [DataRow("mm:ss.0", "time")]
    [DataRow("[h]:mm", "")]
    [DataRow("[mm]:ss", "")]
    public void AFormatSaysWhetherItShowsADateOrATime(string excel, string expected)
    {
        var parts = ExcelDateFormat.PartsShown(excel);
        var shown = new List<string>();
        if ((parts & ExcelDateFormat.Parts.Date) != 0) shown.Add("date");
        if ((parts & ExcelDateFormat.Parts.Time) != 0) shown.Add("time");
        Assert.AreEqual(expected, string.Join(" ", shown));
    }

    [DataTestMethod]
    [DataRow("d", "m/d/yy")]
    [DataRow("g", "m/d/yy h:mm")]
    [DataRow("t", "h:mm")]
    [DataRow("T", "h:mm:ss")]
    [DataRow("s", "yyyy-mm-dd\"T\"hh:mm:ss")]
    [DataRow("o", "yyyy-mm-dd\"T\"hh:mm:ss.000")]
    [DataRow("y", "mmmm yyyy")]
    public void StandardFormatsHaveAFixedSpelling(string dotNet, string excel)
    {
        Assert.AreEqual(excel, ExcelDateFormat.ToExcel(dotNet));
    }

    [TestMethod]
    public void NoFormatMeansDateAndTime()
    {
        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", ExcelDateFormat.ToExcel(null));
        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", ExcelDateFormat.ToExcel(""));
    }
}
