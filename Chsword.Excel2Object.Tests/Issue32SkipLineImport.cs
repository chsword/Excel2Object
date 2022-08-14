using System;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class Issue32SkipLineImport : BaseExcelTest
{
    [TestMethod]
    public void SkipLineImport()
    {
        var path = GetLocalFilePath("test-issue32-skipline.xlsx");
        var importer = new ExcelImporter();
        var result =
            importer.ExcelToObject<TestModelPerson>(
                    path, options => { options.TitleSkipLine = 3; })
                .ToList();
        Assert.AreEqual(2, result.Count);
        Console.WriteLine(JsonConvert.SerializeObject(result));
    }
}