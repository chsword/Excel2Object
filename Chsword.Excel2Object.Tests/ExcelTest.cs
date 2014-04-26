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
        public void Test()
        {
            
            var models = GetModels();
            var exporter = new ExcelExporter();
            var bytes = exporter.ObjectToExcelBytes(models);
            var path = GetFilePath("text.xls");
            Console.WriteLine(path);
            File.WriteAllBytes(path, bytes);
            Assert.IsTrue(File.Exists(path));
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(path);
            Assert.AreEqual(models.Count, result.Count());

        }
    }
}
