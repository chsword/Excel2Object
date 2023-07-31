using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.Linq;

namespace Chsword.Excel2Object.Tests;
[TestClass]
public class Issue37SpecialCharTest : BaseExcelTest
{
    [TestMethod]
    public void SpecialCharTest()
    {
        var path = GetLocalFilePath("test.person.special-char.xlsx");
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPersonSpecialChar>(path)!.ToList();
        Assert.AreEqual(2, result.Count);
        Assert.AreEqual("100", result[0].Money);
        Assert.AreEqual("200", result[1].Money);
        Console.WriteLine(JsonConvert.SerializeObject(result));
    }


}
[ExcelTitle("Test Person")]
public class TestModelPersonSpecialChar
{
    [ExcelTitle("姓名$")] public string Name { get; set; } = null!;
    [ExcelTitle("$年龄")] public int? Age { get; set; }
    [ExcelTitle("出生日期#")] public DateTime? Birthday { get; set; }
    [ExcelTitle("金额$")] public string Money { get; set; } = null!;
}