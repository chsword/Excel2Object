using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     A formula column may take over a column that already exists on the model (they are matched by
///     title), and then inherits that column's [ExcelColumn] style - Format included. Up to v2.3 a
///     Format Excel had no builtin for made the exporter drop the formula and write the model's own
///     value as text instead; now the Format becomes the cell's number format and the formula is kept.
/// </summary>
[TestClass]
public class FormulaColumnFormatTest : BaseExcelTest
{
    private const string CustomFormat = "yyyy-MM-dd HH:mm:ss";
    private static readonly DateTime Birthday = new(2026, 9, 11, 14, 30, 45);

    public class Model
    {
        [ExcelTitle("姓名")] public string Name { get; set; } = "";

        [ExcelColumn("生日", Format = CustomFormat, CellBold = true)]
        public DateTime Birthday { get; set; }

        [ExcelColumn("备注", Format = CustomFormat)]
        public string Note { get; set; } = "";

        [ExcelColumn("年份", CellBold = true)] public int Year { get; set; }
    }

    private static IRow Export()
    {
        var list = new List<Model>
        {
            new() {Name = "test", Birthday = Birthday, Note = "n/a", Year = 2026}
        };
        var bytes = new ExcelExporter().ObjectToExcelBytes(list, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.FormulaColumns.Add("生日", c => c["姓名"]);
            options.FormulaColumns.Add("备注", c => c["姓名"]);
            options.FormulaColumns.Add("年份", c => c["姓名"]);
        });
        Assert.IsNotNull(bytes);
        return WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0).GetRow(1);
    }

    /// <summary>
    ///     The formula is written, and the date column it took over lends it its date format - the
    ///     property being a DateTime is enough, no FormulaResultType needed.
    /// </summary>
    [TestMethod]
    public void CustomFormatOnADateColumnBecomesTheFormulaCellsFormat()
    {
        var cell = Export().GetCell(1);
        Assert.AreEqual(CellType.Formula, cell.CellType);
        Assert.AreEqual("A2", cell.CellFormula);
        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", cell.CellStyle.GetDataFormatString());
    }

    /// <summary>
    ///     The column's font and alignment reach the formula cell too - for a date column on the same
    ///     style as its format, for any other column on a style of its own.
    /// </summary>
    [TestMethod]
    public void TheColumnsLookReachesTheFormulaCell()
    {
        var row = Export();
        Assert.IsTrue(row.GetCell(1).CellStyle.GetFont(row.Sheet.Workbook).IsBold);
        Assert.IsTrue(row.GetCell(3).CellStyle.GetFont(row.Sheet.Workbook).IsBold);
    }

    /// <summary>
    ///     Format means nothing on a string column, so a formula taking one over gets no date format.
    /// </summary>
    [TestMethod]
    public void FormatOnAStringColumnIsIgnoredAndTheFormulaIsWritten()
    {
        var cell = Export().GetCell(2);
        Assert.AreEqual(CellType.Formula, cell.CellType);
        Assert.AreEqual("A2", cell.CellFormula);
        Assert.AreEqual(0, cell.CellStyle.DataFormat);
    }

    /// <summary>A nullable FormulaResultType says the same thing as its underlying type.</summary>
    [TestMethod]
    public void ANullableFormulaResultTypeStillGetsTheDateFormat()
    {
        var bytes = new ExcelExporter().ObjectToExcelBytes(new List<Model> {new()}, options =>
        {
            options.ExcelType = ExcelType.Xlsx;
            options.FormulaColumns.Add(new FormulaColumn
            {
                Title = "创建",
                Formula = c => c["姓名"],
                FormulaResultType = typeof(DateTime?)
            });
        });
        Assert.IsNotNull(bytes);

        var cell = WorkbookFactory.Create(new MemoryStream(bytes)).GetSheetAt(0).GetRow(1).GetCell(4);
        Assert.AreEqual(CellType.Formula, cell.CellType);
        Assert.AreEqual("yyyy-mm-dd hh:mm:ss", cell.CellStyle.GetDataFormatString());
    }

    [TestMethod]
    public void WithoutACustomFormatTheFormulaIsWritten()
    {
        var cell = Export().GetCell(3);
        Assert.AreEqual(CellType.Formula, cell.CellType);
        Assert.AreEqual("A2", cell.CellFormula);
    }
}
