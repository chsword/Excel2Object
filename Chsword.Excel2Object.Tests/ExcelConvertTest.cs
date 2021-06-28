using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace Chsword.Excel2Object.Tests
{
    [TestClass]
    public class ExcelConvertTest : BaseExcelTest
    {
        [TestMethod]
        public void ConvertXlsBytesTest()
        {
            var models = GetModels();
            var bytes = ExcelHelper.ObjectToExcelBytes(models);
            Assert.IsTrue(bytes.Length > 0);
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(bytes).ToList();
            models.AreEqual(result);
        }


        [TestMethod]
        public void ConvertXlsFileTest()
        {
            var models = GetModels();
            var bytes = ExcelHelper.ObjectToExcelBytes(models);
            var path = GetFilePath("test.xls");
            File.WriteAllBytes(path, bytes);
            Assert.IsTrue(File.Exists(path));
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(path).ToList();
            Assert.AreEqual(models.Count, result.Count());
            models.AreEqual(result);
        }

        [TestMethod]
        public void ConvertXlsFileUseObjectToExcelTest()
        {
            var models = GetModels();
            var path = GetFilePath("test.xls");
            ExcelHelper.ObjectToExcel(models, path);
            Assert.IsTrue(File.Exists(path));
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(path).ToList();
            Assert.AreEqual(models.Count, result.Count);
            models.AreEqual(result);
        }

        [TestMethod]
        public void ConvertXlsFromDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("姓名", typeof(string)));
            dt.Columns.Add(new DataColumn("Age", typeof(int)));
            DataRow dr = dt.NewRow();
            dr["姓名"] = "吴老狗";
            dr["Age"] = 19;
            dt.Rows.Add(dr);
            var bytes = ExcelHelper.ObjectToExcelBytes(dt, ExcelType.Xls);
            var path = GetFilePath("test.xls");
            File.WriteAllBytes(path, bytes);
            Assert.IsTrue(File.Exists(path));
        }

        [TestMethod]
        public void ConvertXlsxBytesTest()
        {
            var models = GetModels();
            var array = ExcelHelper.ObjectToExcelBytes(models, ExcelType.Xlsx);
            Assert.IsTrue(array.Length != 0);
            var excelImporter = new ExcelImporter();
            var result = excelImporter.ExcelToObject<ReportModel>(array).ToList();
            models.AreEqual(result);
        }

        [TestMethod]
        public void ConvertXlsxFileTest()
        {
            var models = GetModels();
            var bytes = ExcelHelper.ObjectToExcelBytes(models, ExcelType.Xlsx);
            var path = GetFilePath("test.xlsx");
            File.WriteAllBytes(path, bytes);
            Assert.IsTrue(File.Exists(path));
            var importer = new ExcelImporter();
            var result = importer.ExcelToObject<ReportModel>(path).ToList();
            Assert.AreEqual(models.Count, result.Count);
            models.AreEqual(result);
        }

        [TestMethod]
        public void ConvertXlsxWithDictionary()
        {
            var list = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object> {["姓名"] = "吴老狗", ["Age"] = "19"},
                new Dictionary<string, object> {["姓名"] = "老林", ["Age"] = "50"}
            };
            var bytes = ExcelHelper.ObjectToExcelBytes(list, ExcelType.Xlsx);
            var path = GetFilePath("test.xlsx");
            File.WriteAllBytes(path, bytes);
            var result = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes).ToList();
            Assert.AreEqual(
                JsonConvert.SerializeObject(list),
                JsonConvert.SerializeObject(result)
            );
        }

        private ReportModelCollection GetModels()
        {
            return new ReportModelCollection
            {
                new ReportModel
                {
                    Name = "a", Title = "b", Enabled = true
                },
                new ReportModel
                {
                    Name = "c", Title = "d", Enabled = false
                },
                new ReportModel
                {
                    Name = "f", Title = "e", Uri = new Uri("http://chsword.cnblogs.com")
                }
            };
        }
    }
}