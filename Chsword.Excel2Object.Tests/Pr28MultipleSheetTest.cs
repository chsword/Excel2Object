using System;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class Pr28MultipleSheetTest : BaseExcelTest
    {
        [TestMethod]
        public void ImportMultipleSheet()
        {
            var path = GetLocalFilePath("test-pr28-multiples-heet.xlsx");
            var importer = new ExcelImporter();
            var sheetName = "Flat3Door";
            var result = importer.ExcelToObject<TestModelPerson>(path, sheetName).ToList();
            Assert.AreEqual(3,result.Count);
            Assert.AreEqual("陈皮", result[0].Name);

            Console.WriteLine(JsonConvert.SerializeObject(result));
        }

        [TestMethod]
        public void ImportMultipleSheetException()
        {
            Assert.ThrowsException<Excel2ObjectException>(() =>
            {
                var path = GetLocalFilePath("test-pr28-multiples-heet.xlsx");
                var importer = new ExcelImporter();
                var sheetName = "xxxxxxxxxxxxxxxxxxxx3Door";
                var result = importer.ExcelToObject<TestModelPerson>(path, sheetName).ToList();
                Assert.AreEqual(3, result.Count);
                Assert.AreEqual("陈皮", result[0].Name);
            });
        }
    }
}