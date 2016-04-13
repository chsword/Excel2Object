using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExcelTest
    {
        private ReportModelCollection GetModels()
        {
            return new ReportModelCollection
            {
                new ReportModel {Name = "a", Title = "b", Enabled = true},
                new ReportModel {Name = "c", Title = "d", Enabled = false},
                new ReportModel {Name = "f", Title = "e"}
            };
        }

        private string GetFilePath(string file)
        {
            return Path.Combine(Environment.CurrentDirectory, file);
        }

        [TestMethod]
        public void ConvertTest()
        {
            var models = GetModels();
            var bytes = ExcelHelper.ObjectToExcelBytes(models);
            var path = GetFilePath("text.xls");
            File.WriteAllBytes(path, bytes);
            Assert.IsTrue(File.Exists(path));
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(path);
            Assert.AreEqual(models.Count, result.Count());
        }

        [TestMethod]
        public void ConvertTest1()
        {
            var models = GetModels();
            var bytes = ExcelHelper.ObjectToExcelBytes(models);


            Assert.IsTrue(bytes.Length > 0);
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(bytes).ToList();
            models.AreEqual(result);
        }
        [TestMethod]
        public void XlsxTest()
        {
            var models = this.GetModels();
            byte[] array = ExcelHelper.ObjectToExcelBytes<ReportModel>(models, ExcelType.Xlsx);
            Assert.IsTrue(array.Length != 0);
            ExcelImporter excelImporter = new ExcelImporter();
            var result = excelImporter.ExcelToObject<ReportModel>(array).ToList();
            models.AreEqual(result);
        }
    }
}