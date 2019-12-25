using System;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests
{
    //https://github.com/chsword/Excel2Object/issues/12
    [TestClass]
    public class ExcelIssue12Test
    {
        private ReportModelCollection GetModels()
        {
            return new ReportModelCollection
            {
                new ReportModel
                {
                    Name = "x", Title = "", Enabled = true
                },
                new ReportModel
                {
                    Name = "y", Title = null, Enabled = false
                },
                new ReportModel
                {
                    Name = "z", Title = "e", Uri = new Uri("http://chsword.cnblogs.com")
                }
            };
        }
        [TestMethod]
        public void EmptyFirstProperty()
        {
            var models = GetModels();
            var bytes = ExcelHelper.ObjectToExcelBytes(models);
            Assert.IsNotNull(bytes);
            Assert.IsTrue(bytes.Length>0);
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(bytes).ToList();
            Console.WriteLine(result.FirstOrDefault());
            Assert.AreEqual(models.Count, result.Count());
            models.AreEqual(result);
        }
    }
}