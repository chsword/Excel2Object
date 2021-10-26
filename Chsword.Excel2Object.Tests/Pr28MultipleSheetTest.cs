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
        // Import

        [TestMethod]
        public void ImportMultipleSheet()
        {
            var path = GetLocalFilePath("test-pr28-multiples-heet.xlsx");
            var importer = new ExcelImporter();
            var resultFlat = importer.ExcelToObject<TestModelPerson>(path, "Flat3Door").ToList();
            Assert.AreEqual(3,resultFlat.Count);
            Assert.AreEqual("陈皮", resultFlat[0].Name);
            Console.WriteLine(JsonConvert.SerializeObject(resultFlat));
            var resultUp = importer.ExcelToObject<TestModelPerson>(path, "Up3Door").ToList();
            Assert.AreEqual(3, resultUp.Count);
            Assert.AreEqual("张启山", resultUp[0].Name);
            Console.WriteLine(JsonConvert.SerializeObject(resultUp));
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

        // Export

        [TestMethod]
        public void ExportMultipleSheet()
        {
            // read source
      
            var path = GetLocalFilePath("test-pr28-multiples-heet.xlsx");
            var flatModel = ExcelHelper.ExcelToObject<TestModelPerson>(path, "Flat3Door").ToList();
            var upModel = ExcelHelper.ExcelToObject<TestModelPerson>(path, "Up3Door").ToList();
            var downModel = ExcelHelper.ExcelToObject<TestModelPerson>(path, "Down3Door").ToList();
            // write bytes
            var bytes = ExcelHelper.ObjectToExcelBytes(upModel, sheetTitle: "Up3Door");
            // append
            bytes = ExcelHelper.AppendObjectToExcelBytes(bytes, flatModel, "Flat3Door");
            bytes = ExcelHelper.AppendObjectToExcelBytes(bytes, downModel, "Down3Door");
            // check
            var importer = new ExcelImporter();
            var resultFlat = importer.ExcelToObject<TestModelPerson>(bytes, "Flat3Door").ToList();
            Assert.AreEqual(3, resultFlat.Count);
            Assert.AreEqual("陈皮", resultFlat[0].Name);
            Console.WriteLine(JsonConvert.SerializeObject(resultFlat));
            var resultUp = importer.ExcelToObject<TestModelPerson>(bytes, "Up3Door").ToList();
            Assert.AreEqual(3, resultUp.Count);
            Assert.AreEqual("张启山", resultUp[0].Name);
            Console.WriteLine(JsonConvert.SerializeObject(resultUp));
        }
    }
}