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
        List<ReportModel> GetModels()
        {
            return new List<ReportModel>
            {
                new ReportModel{Name="a",Title="b"},
                new ReportModel{Name="c",Title="d"},
                new ReportModel{Name="f",Title="e"}
            };
        }

        string GetFilePath(string file)
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
            var result = importer.ExcelToObject<ReportModel>(bytes);
            Assert.AreEqual(models.Count, result.Count());

        }
    }
}
