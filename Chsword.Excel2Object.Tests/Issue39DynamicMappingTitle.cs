using System;
using System.Collections.Generic;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Newtonsoft.Json;

namespace Chsword.Excel2Object.Tests;
[TestClass]
public class Issue39DynamicMappingTitleTest
{
    [TestMethod]
    public void MappingTitle()
    {
        var models = GetModels();
        var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
        {
            options.MappingColumnAction = (title,_) =>
            {
                if (title == "姓名")
                {
                    return "n c name";
                }
                return title;
            };
        });
        Assert.IsNotNull(bytes);

        Assert.IsTrue(bytes.Length > 0);

        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<Dictionary<string,object>>(bytes).ToList();
        Console.WriteLine(JsonConvert.SerializeObject(result));
 
    }

    private IEnumerable<TestModelDatePerson> GetModels()
    {
        var list =  new List<TestModelDatePerson>
        {
            new()
            {
                Name = "Three Zhang",
                Age = 18,
                Birthday = new DateTime(1990, 1, 1), 
                Birthday2 = null
            },
            new()
            {
                Name = "Four Lee",
                Age = 18,
                Birthday = new DateTime(1990, 1, 1), 
                Birthday2 = null
            },
        };
        return list;
    }
}