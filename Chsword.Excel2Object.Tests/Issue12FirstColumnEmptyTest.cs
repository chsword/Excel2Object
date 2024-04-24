using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace Chsword.Excel2Object.Tests;

//https://github.com/chsword/Excel2Object/issues/12
[TestClass]
public class Issue12FirstColumnEmptyTest
{
    [TestMethod]
    public void EmptyFirstProperty()
    {
        var models = GetModels();
        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        Assert.IsNotNull(bytes);
        Assert.IsTrue(bytes.Length > 0);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<ReportModel>(bytes).ToList();
        Console.WriteLine(result.FirstOrDefault());
        Assert.AreEqual(models.Count, result.Count);
        models.AreEqual(result);
    }

    private ReportModelCollection GetModels()
    {
        return new ReportModelCollection
        {
            new()
            {
                Name = "x", Title = "", Enabled = true
            },
            new()
            {
                Name = "y", Title = "", Enabled = false
            },
            new()
            {
                Name = "z", Title = "e", Uri = new Uri("http://chsword.cnblogs.com")
            }
        };
    }
}