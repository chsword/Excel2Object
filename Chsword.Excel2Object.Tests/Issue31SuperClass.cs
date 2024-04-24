using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class Issue31SuperClass
{
    [TestMethod]
    public void CheckModelA()
    {
        var excel = TypeConvert.ConvertObjectToExcelModel(GetExcel<SubClassA>()!,
            new ExcelExporterOptions());

        Assert.IsNotNull(excel);

        Assert.AreEqual(1, excel.Sheets?.Count);
        Assert.IsNotNull(excel.Sheets);
        Assert.AreEqual("SuperClass", excel.Sheets[0].Title);
        Console.WriteLine(JsonConvert.SerializeObject(excel));

        Assert.IsTrue(excel.Sheets[0].Columns.Any(c => c.Title == "IdA"));
        Assert.IsTrue(excel.Sheets[0].Columns.Any(c => c.Title == "P1"));
    }

    [TestMethod]
    public void CheckModelB()
    {
        var excel = TypeConvert.ConvertObjectToExcelModel(
            GetExcel<SubClassB>(),
            new ExcelExporterOptions());
        Assert.IsNotNull(excel);
        Assert.IsNotNull(excel.Sheets);
        Assert.AreEqual(1, excel.Sheets.Count);
        Assert.AreEqual("SubClassB", excel.Sheets[0].Title);
        Console.WriteLine(JsonConvert.SerializeObject(excel));

        Assert.IsTrue(excel.Sheets[0].Columns
            .Any(c => c.Title == "IdB"));
        Assert.IsTrue(excel.Sheets[0]
            .Columns.Any(c => c.Title == "P1"));
    }

    [TestMethod]
    public void SuperClassTest()
    {
        var export = new ExcelExporter();
        var bytes = export.ObjectToExcelBytes(GetExcel<SubClassA>());
        var importer = new ExcelImporter();
        Assert.IsNotNull(bytes);
        Console.WriteLine(bytes.Length);
        Assert.IsNotNull(importer);
        // var model = ExcelImporter.
    }

    private List<T> GetExcel<T>() where T : SuperClass
    {
        if (typeof(T).Name == "SubClassA")
        {
            var list = new List<SubClassA>
            {
                new() {Id = 1, P = "x"},
                new() {Id = 2, P = "x"}
            };
            return (list as List<T>)!;
        }

        if (typeof(T).Name == "SubClassB")
        {
            var list = new List<SubClassB>
            {
                new() {Id = 11, P = "x"},
                new() {Id = 12, P = "x"}
            };
            return (list as List<T>)!;
        }
        throw new Exception("not support");
    }

    [ExcelTitle("SuperClass")]
    public abstract class SuperClass
    {
        [ExcelColumn("Id1")] public int Id { get; set; }

        // ReSharper disable once NotNullOrRequiredMemberIsNotInitialized
        [ExcelColumn("P1")] public string P { get; set; }
    }

    public class SubClassA : SuperClass
    {
        [ExcelColumn("IdA")] public new int Id { get; set; }
    }

    [ExcelTitle("SubClassB")]
    public class SubClassB : SuperClass
    {
        [ExcelColumn("IdB")] public new int Id { get; set; }
    }
}