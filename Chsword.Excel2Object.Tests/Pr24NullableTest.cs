using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class Pr24NullableTest : BaseExcelTest
{
    [TestMethod]
    public void ImportExcelNullableType()
    {
        var path = GetLocalFilePath("test.person.xlsx");
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(path)!.ToList();
        Assert.AreEqual(2, result.Count);
        Console.WriteLine(JsonConvert.SerializeObject(result));
    }

    [TestMethod]
    public void ImportExcelUnNullableTypeException()
    {
        var path = GetLocalFilePath("test.person.xlsx");
        var importer = new ExcelImporter();
        Assert.ThrowsException<FormatException>(() =>
        {
            var result = importer.ExcelToObject<TestModelStrictPerson>(path)!.ToList();
            Console.WriteLine(JsonConvert.SerializeObject(result));
        });
    }

    [TestMethod]
    public void ImportExcelUnNullableType()
    {
        var path = GetLocalFilePath("test.person.unnullable.xlsx");
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelStrictPerson>(path)!.ToList();
        Assert.AreEqual(2, result.Count);
        Console.WriteLine(JsonConvert.SerializeObject(result));
    }

    private List<TestModelPerson> GetPersonList()
    {
        return new List<TestModelPerson>
        {
            new() {Name = "张三", Age = 18, Birthday = null},
            new() {Name = "李四", Age = null, Birthday = new DateTime(2021, 10, 10)}
        };
    }

    [TestMethod]
    public void ExportExcelNullableType()
    {
        var personList = GetPersonList();
        var bytes = ExcelHelper.ObjectToExcelBytes(personList);
        Assert.IsNotNull(bytes);
        Assert.IsTrue(bytes.Length > 0);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();
        Console.WriteLine(JsonConvert.SerializeObject(result));
    }
}