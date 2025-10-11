using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Tests.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

[TestClass]
public class DateTimeFormatTest : BaseExcelTest
{
    [TestMethod]
    public void TestISO8601DateTimeFormat()
    {
        // Test ISO 8601 format: yyyy-MM-ddTHH:mm:ss
        var testDate = new DateTime(2023, 7, 31, 14, 30, 45);
        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Test Person",
                Age = 25,
                Birthday = testDate
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(1, result.Count);
        Assert.IsNotNull(result[0].Birthday);
        Assert.AreEqual(testDate.Date, result[0].Birthday!.Value.Date);
    }

    [TestMethod]
    public void TestCommonDateFormats()
    {
        var testDate = new DateTime(2023, 12, 25);
        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Test Person",
                Age = 30,
                Birthday = testDate
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(1, result.Count);
        Assert.IsNotNull(result[0].Birthday);
        Assert.AreEqual(testDate.Date, result[0].Birthday!.Value.Date);
    }

    [TestMethod]
    public void TestDateTimeWithSeconds()
    {
        var testDate = new DateTime(2023, 6, 15, 9, 45, 30);
        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Test Person",
                Age = 35,
                Birthday = testDate,
                CreateTime = testDate
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(1, result.Count);
        Assert.IsNotNull(result[0].Birthday);
        Assert.AreEqual(testDate.Date, result[0].Birthday!.Value.Date);
    }

    [TestMethod]
    public void TestNullableDateTimeFormat()
    {
        var testDate = new DateTime(2023, 3, 10, 18, 20, 0);
        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Test Person 1",
                Age = 28,
                Birthday = testDate
            },
            new()
            {
                Name = "Test Person 2",
                Age = 32,
                Birthday = null
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(2, result.Count);
        Assert.IsNotNull(result[0].Birthday);
        Assert.AreEqual(testDate.Date, result[0].Birthday!.Value.Date);
        Assert.IsNull(result[1].Birthday);
    }

    [TestMethod]
    public void TestDateTimeWithCustomFormat()
    {
        var testDate = new DateTime(2023, 11, 5, 14, 30, 0);
        var models = new List<TestModelDatePerson>
        {
            new()
            {
                Name = "Test Person",
                Age = 40,
                Birthday = testDate,
                Birthday2 = testDate
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelDatePerson>(bytes).ToList();

        Assert.AreEqual(1, result.Count);
        Assert.AreEqual(testDate.Date, result[0].Birthday.Date);
        Assert.IsNotNull(result[0].Birthday2);
        Assert.AreEqual(testDate.Date, result[0].Birthday2!.Value.Date);
    }

    [TestMethod]
    public void TestMultipleDateFormatsInSameFile()
    {
        var testDate1 = new DateTime(2023, 1, 15);
        var testDate2 = new DateTime(2023, 8, 20, 16, 45, 0);
        var testDate3 = new DateTime(2023, 12, 31, 23, 59, 59);

        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Person 1",
                Age = 25,
                Birthday = testDate1,
                CreateTime = testDate1
            },
            new()
            {
                Name = "Person 2",
                Age = 30,
                Birthday = testDate2,
                CreateTime = testDate2
            },
            new()
            {
                Name = "Person 3",
                Age = 35,
                Birthday = testDate3,
                CreateTime = testDate3
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(3, result.Count);
        Assert.AreEqual(testDate1.Date, result[0].Birthday!.Value.Date);
        Assert.AreEqual(testDate2.Date, result[1].Birthday!.Value.Date);
        Assert.AreEqual(testDate3.Date, result[2].Birthday!.Value.Date);
    }

    [TestMethod]
    public void TestDateOnlyFormat()
    {
        // Test date without time component
        var testDate = new DateTime(2023, 5, 20);
        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Test Person",
                Age = 27,
                Birthday = testDate
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(1, result.Count);
        Assert.IsNotNull(result[0].Birthday);
        Assert.AreEqual(testDate.Date, result[0].Birthday!.Value.Date);
    }

    [TestMethod]
    public void TestMinMaxDateValues()
    {
        // Test edge cases with min and max reasonable date values
        var minDate = new DateTime(1900, 1, 1);
        var maxDate = new DateTime(2099, 12, 31);

        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Min Date Person",
                Age = 25,
                Birthday = minDate
            },
            new()
            {
                Name = "Max Date Person",
                Age = 30,
                Birthday = maxDate
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(2, result.Count);
        Assert.AreEqual(minDate.Date, result[0].Birthday!.Value.Date);
        Assert.AreEqual(maxDate.Date, result[1].Birthday!.Value.Date);
    }

    [TestMethod]
    public void TestLeapYearDate()
    {
        // Test leap year date: February 29, 2024
        var leapDate = new DateTime(2024, 2, 29);
        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Leap Year Person",
                Age = 29,
                Birthday = leapDate
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(1, result.Count);
        Assert.AreEqual(leapDate.Date, result[0].Birthday!.Value.Date);
    }

    [TestMethod]
    public void TestMidnightAndNoonTimes()
    {
        // Test midnight and noon times
        var midnight = new DateTime(2023, 7, 15, 0, 0, 0);
        var noon = new DateTime(2023, 7, 15, 12, 0, 0);

        var models = new List<TestModelPerson>
        {
            new()
            {
                Name = "Midnight Person",
                Age = 25,
                Birthday = midnight,
                CreateTime = midnight
            },
            new()
            {
                Name = "Noon Person",
                Age = 30,
                Birthday = noon,
                CreateTime = noon
            }
        };

        var bytes = ExcelHelper.ObjectToExcelBytes(models);
        var importer = new ExcelImporter();
        var result = importer.ExcelToObject<TestModelPerson>(bytes).ToList();

        Assert.AreEqual(2, result.Count);
        Assert.AreEqual(midnight.Date, result[0].Birthday!.Value.Date);
        Assert.AreEqual(noon.Date, result[1].Birthday!.Value.Date);
    }
}
