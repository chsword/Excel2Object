using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Chsword.Excel2Object.Cli;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Tests;

/// <summary>The excel2obj command-line tool (issue #9), driven through Excel2ObjCli.Run with captured output.</summary>
[TestClass]
public class CliTest
{
    public class Order
    {
        [ExcelTitle("Product")] public string Product { get; set; } = "";
        [ExcelTitle("Qty")] public int Qty { get; set; }
        [ExcelTitle("Unit Price")] public decimal? UnitPrice { get; set; }
        [ExcelTitle("Active")] public bool Active { get; set; }
    }

    private static readonly List<Order> Orders = new()
    {
        new() {Product = "Apple", Qty = 4, UnitPrice = 3.5m, Active = true},
        new() {Product = "梨", Qty = 2, UnitPrice = null, Active = false}
    };

    private string _dir = "";

    [TestInitialize]
    public void CreateWorkDir()
    {
        _dir = Path.Combine(Path.GetTempPath(), "excel2obj-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
    }

    [TestCleanup]
    public void RemoveWorkDir()
    {
        Directory.Delete(_dir, true);
    }

    private string WriteWorkbook(string name = "orders.xlsx", string? sheet = null)
    {
        var path = Path.Combine(_dir, name);
        var bytes = new ExcelExporter().ObjectToExcelBytes(Orders, o =>
        {
            o.ExcelType = name.EndsWith(".xls") ? ExcelType.Xls : ExcelType.Xlsx;
            o.SheetTitle = sheet;
        });
        File.WriteAllBytes(path, bytes!);
        return path;
    }

    private static (int code, string stdout, string stderr) Run(params string[] args)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Excel2ObjCli.Run(args, stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }

    [TestMethod]
    public void ExcelToJsonWritesCellTextByDefault()
    {
        var (code, stdout, _) = Run("convert", WriteWorkbook());
        Assert.AreEqual(Excel2ObjCli.Ok, code);

        using var doc = JsonDocument.Parse(stdout);
        var rows = doc.RootElement.EnumerateArray().ToList();
        Assert.AreEqual(2, rows.Count);
        Assert.AreEqual("Apple", rows[0].GetProperty("Product").GetString());
        Assert.AreEqual("4", rows[0].GetProperty("Qty").GetString());
        Assert.AreEqual("3.5", rows[0].GetProperty("Unit Price").GetString());
        Assert.AreEqual("TRUE", rows[0].GetProperty("Active").GetString());
        Assert.AreEqual("梨", rows[1].GetProperty("Product").GetString());
        Assert.AreEqual("", rows[1].GetProperty("Unit Price").GetString());
        StringAssert.Contains(stdout, "梨"); // not escaped as \uXXXX
    }

    [TestMethod]
    public void TypedExcelToJsonInfersNumbersBoolsAndNulls()
    {
        var output = Path.Combine(_dir, "orders.json");
        var (code, stdout, stderr) = Run("convert", WriteWorkbook(), "--typed", "--output", output);
        Assert.AreEqual(Excel2ObjCli.Ok, code, stderr);
        Assert.AreEqual("", stdout);

        using var doc = JsonDocument.Parse(File.ReadAllText(output));
        var rows = doc.RootElement.EnumerateArray().ToList();
        Assert.AreEqual(4, rows[0].GetProperty("Qty").GetInt32());
        Assert.AreEqual(3.5, rows[0].GetProperty("Unit Price").GetDouble());
        Assert.IsTrue(rows[0].GetProperty("Active").GetBoolean());
        Assert.AreEqual(JsonValueKind.Null, rows[1].GetProperty("Unit Price").ValueKind);
        Assert.IsFalse(rows[1].GetProperty("Active").GetBoolean());
    }

    [TestMethod]
    public void JsonToExcelProducesTypedCellsAndRoundTrips()
    {
        var json = Path.Combine(_dir, "orders.json");
        File.WriteAllText(json, """
            [
              {"Product": "Apple", "Qty": 4, "Unit Price": 3.5, "Active": true},
              {"Product": "梨", "Qty": 2, "Unit Price": null, "Active": false, "Extra": {"a": 1}}
            ]
            """);
        var xlsx = Path.Combine(_dir, "orders.xlsx");
        var (code, _, stderr) = Run("convert", json, "--output", xlsx, "--sheet", "Orders");
        Assert.AreEqual(Excel2ObjCli.Ok, code, stderr);

        using (var stream = File.OpenRead(xlsx))
        {
            var sheet = WorkbookFactory.Create(stream).GetSheet("Orders");
            Assert.IsNotNull(sheet);
            var row = sheet.GetRow(1);
            Assert.AreEqual(CellType.String, row.GetCell(0).CellType);
            Assert.AreEqual(CellType.Numeric, row.GetCell(1).CellType);
            Assert.AreEqual(4d, row.GetCell(1).NumericCellValue);
            Assert.AreEqual(CellType.Boolean, row.GetCell(3).CellType);
            Assert.AreEqual(CellType.Blank, sheet.GetRow(2).GetCell(2).CellType);
            Assert.AreEqual("Extra", sheet.GetRow(0).GetCell(4).StringCellValue);
            Assert.AreEqual("{\"a\":1}", sheet.GetRow(2).GetCell(4).StringCellValue);
        }

        var orders = ExcelHelper.ExcelToObject<Order>(File.ReadAllBytes(xlsx), "Orders").ToList();
        Assert.AreEqual("Apple", orders[0].Product);
        Assert.AreEqual(4, orders[0].Qty);
        Assert.AreEqual(3.5m, orders[0].UnitPrice);
        Assert.IsTrue(orders[0].Active);
        Assert.IsNull(orders[1].UnitPrice);
    }

    [TestMethod]
    public void SeveralInputsGoToAnOutputDirectory()
    {
        var a = WriteWorkbook("a.xlsx");
        var b = WriteWorkbook("b.xls");
        var outDir = Path.Combine(_dir, "out");
        var (code, _, stderr) = Run("convert", a, b, "--output", outDir);
        Assert.AreEqual(Excel2ObjCli.Ok, code, stderr);
        Assert.IsTrue(File.Exists(Path.Combine(outDir, "a.json")));
        Assert.IsTrue(File.Exists(Path.Combine(outDir, "b.json")));

        Assert.AreEqual(Excel2ObjCli.UsageError, Run("convert", a, b).code);
    }

    [TestMethod]
    public void GenerateModelInfersTypesAndSanitizesTitles()
    {
        var (code, stdout, stderr) = Run("generate-model", WriteWorkbook(sheet: "Order List"), "--namespace", "My.App");
        Assert.AreEqual(Excel2ObjCli.Ok, code, stderr);

        StringAssert.Contains(stdout, "using Chsword.Excel2Object;");
        StringAssert.Contains(stdout, "namespace My.App;");
        StringAssert.Contains(stdout, "public class OrderList"); // class name from the sheet title
        StringAssert.Contains(stdout, "[ExcelTitle(\"Product\")]\n    public string? Product { get; set; }");
        StringAssert.Contains(stdout, "[ExcelTitle(\"Qty\")]\n    public int Qty { get; set; }");
        StringAssert.Contains(stdout, "[ExcelTitle(\"Unit Price\")]\n    public decimal? UnitPrice { get; set; }");
        StringAssert.Contains(stdout, "[ExcelTitle(\"Active\")]\n    public bool Active { get; set; }");
    }

    [TestMethod]
    public void GenerateModelHonoursClassAndOutput()
    {
        var output = Path.Combine(_dir, "Order.cs");
        var (code, stdout, _) = Run("generate-model", WriteWorkbook(), "--class=Order", "--output", output);
        Assert.AreEqual(Excel2ObjCli.Ok, code);
        Assert.AreEqual("", stdout);
        StringAssert.Contains(File.ReadAllText(output), "public class Order\n{");
    }

    [TestMethod]
    public void IdentifiersFromTitles()
    {
        Assert.AreEqual("UnitPrice", GenerateModelCommand.ToIdentifier("unit price", "X"));
        Assert.AreEqual("_2024Total", GenerateModelCommand.ToIdentifier("2024 total", "X"));
        Assert.AreEqual("姓名", GenerateModelCommand.ToIdentifier("姓名", "X"));
        Assert.AreEqual("Column3", GenerateModelCommand.ToIdentifier("***", "Column3"));
    }

    [TestMethod]
    public void TypeInferenceRules()
    {
        Assert.AreEqual(InferredType.Int, TypeInference.Infer(new[] {"1", "", "-3"}));
        Assert.AreEqual(InferredType.Long, TypeInference.Infer(new[] {"1", "9999999999"}));
        Assert.AreEqual(InferredType.Decimal, TypeInference.Infer(new[] {"1", "2.5"}));
        Assert.AreEqual(InferredType.Bool, TypeInference.Infer(new[] {"TRUE", "false"}));
        Assert.AreEqual(InferredType.DateTime, TypeInference.Infer(new[] {"2026-09-08", "2026/09/09 10:00:00"}));
        Assert.AreEqual(InferredType.String, TypeInference.Infer(new[] {"1", "x"}));
        Assert.AreEqual(InferredType.String, TypeInference.Infer(new[] {"", ""}));
    }

    [TestMethod]
    public void UsageErrorsAndHelp()
    {
        var (code, stdout, _) = Run();
        Assert.AreEqual(Excel2ObjCli.Ok, code);
        StringAssert.Contains(stdout, "Usage:");

        var (unknownCode, _, unknownErr) = Run("frobnicate");
        Assert.AreEqual(Excel2ObjCli.UsageError, unknownCode);
        StringAssert.Contains(unknownErr, "unknown command 'frobnicate'");

        var (missingCode, _, missingErr) = Run("convert", Path.Combine(_dir, "nope.xlsx"));
        Assert.AreEqual(Excel2ObjCli.CommandFailed, missingCode);
        StringAssert.Contains(missingErr, "nope.xlsx");

        var (sheetCode, _, sheetErr) = Run("convert", WriteWorkbook(), "--sheet", "Nowhere");
        Assert.AreEqual(Excel2ObjCli.CommandFailed, sheetCode);
        StringAssert.Contains(sheetErr, "[Nowhere]");
    }
}
