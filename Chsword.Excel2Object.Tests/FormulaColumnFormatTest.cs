using System;
using System.Collections.Generic;
using System.IO;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     A formula column may take over a column that already exists on the model (they are matched by
///     title), and then inherits that column's [ExcelColumn] style. When that style carries a Format
///     Excel has no builtin number format for, the exporter cannot express it as a cell format, so it
///     falls back to writing the model's own value as text and drops the formula.
///     <para>
///         These tests record that fallback, they do not endorse it: asking for a formula and silently
///         getting a static value instead is a known defect. They exist so the behaviour cannot change
///         unnoticed, and so a future fix has to update them deliberately.
///     </para>
/// </summary>
[TestClass]
public class FormulaColumnFormatTest : BaseExcelTest
{
    private const string CustomFormat = "yyyy-MM-dd HH:mm:ss";
    private static readonly DateTime Birthday = new(2026, 9, 11, 14, 30, 45);

    public class Model
    {
        [ExcelTitle("姓名")] public string Name { get; set; } = "";

        [ExcelColumn("生日", Format = CustomFormat)]
        public DateTime Birthday { get; set; }

        [ExcelColumn("备注", Format = CustomFormat)]
        public string Note { get; set; } = "";

        [ExcelTitle("年份")] public int Year { get; set; }
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

    [TestMethod]
    public void CustomFormatOnADateColumnSilentlyReplacesTheFormula()
    {
        // known defect: the requested formula is dropped, the model's own value is written instead
        var cell = Export().GetCell(1);
        Assert.AreEqual(CellType.String, cell.CellType);
        Assert.AreEqual(Birthday.ToString(CustomFormat), cell.StringCellValue);
    }

    [TestMethod]
    public void CustomFormatOnAValueThatIsNotADateWritesItVerbatim()
    {
        // same defect, for a value that does not parse as a date
        var cell = Export().GetCell(2);
        Assert.AreEqual(CellType.String, cell.CellType);
        Assert.AreEqual("n/a", cell.StringCellValue);
    }

    [TestMethod]
    public void WithoutACustomFormatTheFormulaIsWritten()
    {
        var cell = Export().GetCell(3);
        Assert.AreEqual(CellType.Formula, cell.CellType);
        Assert.AreEqual("A2", cell.CellFormula);
    }
}
