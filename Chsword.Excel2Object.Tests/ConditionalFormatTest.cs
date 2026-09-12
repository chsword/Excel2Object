using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     A rule restyles the cells whose value meets a condition, and Excel keeps re-evaluating it as the
///     sheet is edited.
/// </summary>
[TestClass]
public class ConditionalFormatTest : BaseExcelTest
{
    public class Model
    {
        [ExcelTitle("城市")] public string City { get; set; } = "北京";
        [ExcelTitle("金额")] public decimal Amount { get; set; } = 1000;
        [ExcelTitle("日期")] public DateTime When { get; set; } = new(2026, 9, 12);
    }

    private static readonly List<Model> Rows = new() {new(), new() {Amount = 20000}};

    private static ISheet Export(ExcelType excelType, Action<ExcelExporterOptions> configure,
        List<Model>? rows = null)
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(rows ?? Rows, options =>
        {
            options.ExcelType = excelType;
            configure(options);
        });
        Assert.IsNotNull(bytes);
        return WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0);
    }

    private static IConditionalFormattingRule SingleRule(ISheet sheet, out CellRangeAddress[] ranges)
    {
        var formatting = sheet.SheetConditionalFormatting;
        Assert.AreEqual(1, formatting.NumConditionalFormattings);
        var applied = formatting.GetConditionalFormattingAt(0);
        ranges = applied.GetFormattingRanges();
        return applied.GetRule(0);
    }

    [TestMethod]
    public void AComparisonColoursItsOwnColumnInBothFileFormats()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var sheet = Export(excelType, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
            {
                Operator = ConditionalOperator.GreaterThan,
                Value = 10000,
                FontColor = ExcelStyleColor.Red,
                Bold = true,
                BackgroundColor = ExcelStyleColor.Yellow
            }));

            var rule = SingleRule(sheet, out var ranges);
            Assert.AreEqual(ConditionType.CellValueIs, rule.ConditionType, excelType.ToString());
            Assert.AreEqual(ComparisonOperator.GreaterThan, rule.ComparisonOperation, excelType.ToString());
            Assert.AreEqual("10000", rule.Formula1, excelType.ToString());
            Assert.AreEqual((short) ExcelStyleColor.Red, rule.FontFormatting.FontColorIndex, excelType.ToString());
            Assert.IsTrue(rule.FontFormatting.IsBold, excelType.ToString());
            Assert.AreEqual((short) ExcelStyleColor.Yellow, rule.PatternFormatting.FillBackgroundColor,
                excelType.ToString());
            // the header keeps its own look; the rule belongs to the data rows of that column
            Assert.AreEqual("B2:B3", ranges[0].FormatAsString(), excelType.ToString());
        }
    }

    /// <summary>A value is written as Excel's literal for its type, so a string is quoted.</summary>
    [DataTestMethod]
    [DataRow("急件", "\"急件\"")]
    [DataRow(1500.5, "1500.5")]
    public void TheValueBecomesAnExcelLiteral(object value, string literal)
    {
        var sheet = Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("城市")
        {
            Operator = ConditionalOperator.Equal,
            Value = value,
            Bold = true
        }));

        Assert.AreEqual(literal, SingleRule(sheet, out _).Formula1);
    }

    [TestMethod]
    public void ADateIsComparedAsADate()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("日期")
        {
            Operator = ConditionalOperator.LessThan,
            Value = new DateTime(2026, 1, 31),
            FontColor = ExcelStyleColor.Grey50Percent
        }));

        Assert.AreEqual("DATE(2026,1,31)", SingleRule(sheet, out _).Formula1);
    }

    [TestMethod]
    public void BetweenTakesBothEnds()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
        {
            Operator = ConditionalOperator.Between,
            Value = 1000,
            Value2 = 5000,
            Italic = true
        }));

        var rule = SingleRule(sheet, out _);
        Assert.AreEqual(ComparisonOperator.Between, rule.ComparisonOperation);
        Assert.AreEqual("1000", rule.Formula1);
        Assert.AreEqual("5000", rule.Formula2);
        Assert.IsTrue(rule.FontFormatting.IsItalic);
    }

    /// <summary>A rule of your own is written as it is, against the first data row.</summary>
    [TestMethod]
    public void AFormulaRuleIsWrittenAsGiven()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
        {
            Formula = "$B2>$B3",
            BackgroundColor = ExcelStyleColor.LightGreen
        }));

        var rule = SingleRule(sheet, out _);
        Assert.AreEqual(ConditionType.Formula, rule.ConditionType);
        Assert.AreEqual("$B2>$B3", rule.Formula1);
    }

    /// <summary>
    ///     The whole row can take the colour instead. Excel compares the cell a rule is written on, so the
    ///     comparison becomes an expression anchored on the column it watches.
    /// </summary>
    [TestMethod]
    public void WholeRowColoursEveryColumn()
    {
        foreach (var excelType in new[] {ExcelType.Xlsx, ExcelType.Xls})
        {
            var sheet = Export(excelType, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
            {
                Operator = ConditionalOperator.GreaterThan,
                Value = 10000,
                WholeRow = true,
                BackgroundColor = ExcelStyleColor.Yellow
            }));

            var rule = SingleRule(sheet, out var ranges);
            Assert.AreEqual(ConditionType.Formula, rule.ConditionType, excelType.ToString());
            Assert.AreEqual("$B2>10000", rule.Formula1, excelType.ToString());
            Assert.AreEqual("A2:C3", ranges[0].FormatAsString(), excelType.ToString());
        }
    }

    [TestMethod]
    public void SeveralRulesCanWatchTheSameColumn()
    {
        var sheet = Export(ExcelType.Xlsx, options =>
        {
            options.ConditionalFormats.Add(new ConditionalFormat("金额")
                {Operator = ConditionalOperator.GreaterThan, Value = 10000, FontColor = ExcelStyleColor.Red});
            options.ConditionalFormats.Add(new ConditionalFormat("金额")
                {Operator = ConditionalOperator.LessThan, Value = 100, FontColor = ExcelStyleColor.Blue});
        });

        Assert.AreEqual(2, sheet.SheetConditionalFormatting.NumConditionalFormattings);
    }

    /// <summary>An empty export is a template, and the rule has to outlive the rows it was written with.</summary>
    [TestMethod]
    public void AnEmptySheetStillGetsTheRule()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
            {Operator = ConditionalOperator.GreaterThan, Value = 1, Bold = true}), new List<Model>());

        SingleRule(sheet, out var ranges);
        Assert.AreEqual("B2", ranges[0].FormatAsString());
    }

    [TestMethod]
    public void AnUnknownColumnSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("不存在")
                {Operator = ConditionalOperator.GreaterThan, Value = 1})));

        StringAssert.Contains(e.Message, "不存在");
    }

    [TestMethod]
    public void AComparisonWithoutAValueSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
                {Operator = ConditionalOperator.GreaterThan})));

        StringAssert.Contains(e.Message, "Value");
        // which of the rules is at fault, when there are several
        StringAssert.Contains(e.Message, "金额");
        StringAssert.Contains(e.Message, "GreaterThan");
    }

    /// <summary>Between without its other end would reach NPOI as half a rule.</summary>
    [DataTestMethod]
    [DataRow(ConditionalOperator.Between)]
    [DataRow(ConditionalOperator.NotBetween)]
    public void BetweenWithoutTheOtherEndSaysSo(ConditionalOperator op)
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
                {Operator = op, Value = 1000})));

        StringAssert.Contains(e.Message, "Value2");
        StringAssert.Contains(e.Message, "金额");
    }

    [TestMethod]
    public void ARuleWithNeitherOperatorNorFormulaSaysSo()
    {
        var e = Assert.ThrowsException<Excel2ObjectException>(() =>
            Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
                {Bold = true})));

        StringAssert.Contains(e.Message, "Formula");
        StringAssert.Contains(e.Message, "金额");
    }

    /// <summary>A whole-row Between becomes an expression over both ends.</summary>
    [TestMethod]
    public void WholeRowBetweenSpellsOutBothEnds()
    {
        var sheet = Export(ExcelType.Xlsx, options => options.ConditionalFormats.Add(new ConditionalFormat("金额")
        {
            Operator = ConditionalOperator.Between,
            Value = 1000,
            Value2 = 5000,
            WholeRow = true,
            BackgroundColor = ExcelStyleColor.Yellow
        }));

        Assert.AreEqual("AND($B2>=1000,$B2<=5000)", SingleRule(sheet, out _).Formula1);
    }
}
