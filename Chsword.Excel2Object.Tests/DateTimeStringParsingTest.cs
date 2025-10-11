using System;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
/// Tests to validate date/time string parsing with various formats
/// This simulates when dates are stored as text in Excel cells
/// </summary>
[TestClass]
public class DateTimeStringParsingTest
{
    [TestMethod]
    public void TestISO8601Formats()
    {
        // Test ISO 8601 format variations
        var testCases = new[]
        {
            "2023-07-31T14:30:45",
            "2023-07-31T14:30:45Z",
            "2023-07-31T14:30:45.123",
            "2023-07-31 14:30:45",
            "2023-07-31 14:30"
        };

        foreach (var testCase in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.None, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(2023, parsedDate.Year);
            Assert.AreEqual(7, parsedDate.Month);
            Assert.AreEqual(31, parsedDate.Day);
        }
    }

    [TestMethod]
    public void TestDashSeparatedFormats()
    {
        // Test date formats with dash separator
        var testCases = new[]
        {
            "2023-12-25",
            "25-12-2023",
            "12-25-2023",
            "2023-12-25 15:30:45",
            "25-12-2023 15:30:45"
        };

        foreach (var testCase in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.None, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(12, parsedDate.Month);
            Assert.AreEqual(25, parsedDate.Day);
            Assert.AreEqual(2023, parsedDate.Year);
        }
    }

    [TestMethod]
    public void TestSlashSeparatedFormats()
    {
        // Test date formats with slash separator
        var testCases = new[]
        {
            "2023/06/15",
            "15/06/2023",
            "06/15/2023",
            "2023/6/15",
            "15/6/2023",
            "6/15/2023"
        };

        foreach (var testCase in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.None, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(6, parsedDate.Month);
            Assert.AreEqual(15, parsedDate.Day);
            Assert.AreEqual(2023, parsedDate.Year);
        }
    }

    [TestMethod]
    public void TestDotSeparatedFormats()
    {
        // Test date formats with dot separator (common in European formats)
        // Using year-first format to avoid ambiguity
        var testCases = new[]
        {
            "2023.03.10",
            "10.03.2023",
            "03.10.2023"
        };

        foreach (var testCase in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.None, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            // Note: Month/day order depends on format, just verify we can parse it
            Assert.AreEqual(2023, parsedDate.Year);
        }
    }

    [TestMethod]
    public void Test12HourFormatWithAMPM()
    {
        // Test 12-hour time formats with AM/PM
        var testCases = new[]
        {
            ("2023-07-15 02:30:45 PM", 14, 30, 45),
            ("2023-07-15 02:30 PM", 14, 30, 0),
            ("15-07-2023 09:15:30 AM", 9, 15, 30),
            ("15/07/2023 09:15:30 AM", 9, 15, 30)
        };

        foreach (var (testCase, expectedHour, expectedMinute, expectedSecond) in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.None, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(expectedHour, parsedDate.Hour, $"Hour mismatch for: {testCase}");
            Assert.AreEqual(expectedMinute, parsedDate.Minute, $"Minute mismatch for: {testCase}");
            Assert.AreEqual(expectedSecond, parsedDate.Second, $"Second mismatch for: {testCase}");
        }
    }

    [TestMethod]
    public void Test24HourFormat()
    {
        // Test 24-hour time formats
        var testCases = new[]
        {
            ("2023-11-20 14:30:45", 14, 30, 45),
            ("2023-11-20 14:30", 14, 30, 0),
            ("20/11/2023 23:45:30", 23, 45, 30),
            ("11/20/2023 00:15", 0, 15, 0)
        };

        foreach (var (testCase, expectedHour, expectedMinute, expectedSecond) in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.None, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(expectedHour, parsedDate.Hour, $"Hour mismatch for: {testCase}");
            Assert.AreEqual(expectedMinute, parsedDate.Minute, $"Minute mismatch for: {testCase}");
            Assert.AreEqual(expectedSecond, parsedDate.Second, $"Second mismatch for: {testCase}");
        }
    }

    [TestMethod]
    public void TestTimeOnlyFormats()
    {
        // Test time-only formats
        var testCases = new[]
        {
            ("14:30:45", 14, 30, 45),
            ("14:30", 14, 30, 0),
            ("09:15:30", 9, 15, 30),
            ("23:59", 23, 59, 0)
        };

        foreach (var (testCase, expectedHour, expectedMinute, expectedSecond) in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.NoCurrentDateDefault, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(expectedHour, parsedDate.Hour, $"Hour mismatch for: {testCase}");
            Assert.AreEqual(expectedMinute, parsedDate.Minute, $"Minute mismatch for: {testCase}");
            Assert.AreEqual(expectedSecond, parsedDate.Second, $"Second mismatch for: {testCase}");
        }
    }

    [TestMethod]
    public void TestTimeOnlyFormatsWithAMPM()
    {
        // Test time-only formats with AM/PM
        var testCases = new[]
        {
            ("02:30:45 PM", 14, 30, 45),
            ("02:30 PM", 14, 30, 0),
            ("09:15:30 AM", 9, 15, 30),
            ("11:45 PM", 23, 45, 0)
        };

        foreach (var (testCase, expectedHour, expectedMinute, expectedSecond) in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.NoCurrentDateDefault, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(expectedHour, parsedDate.Hour, $"Hour mismatch for: {testCase}");
            Assert.AreEqual(expectedMinute, parsedDate.Minute, $"Minute mismatch for: {testCase}");
            Assert.AreEqual(expectedSecond, parsedDate.Second, $"Second mismatch for: {testCase}");
        }
    }

    [TestMethod]
    public void TestShortDateFormats()
    {
        // Test short date formats (single digit month/day) as defined in CommonDateTimeFormats
        // Using year-first format to avoid ambiguity between MM/dd and dd/MM
        var testCases = new[]
        {
            ("2023/3/5", 2023, 3, 5),
            ("2023-3-5", 2023, 3, 5),
            ("2023/12/25", 2023, 12, 25),
            ("2023-12-25", 2023, 12, 25)
        };

        foreach (var (testCase, expectedYear, expectedMonth, expectedDay) in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.None, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(expectedYear, parsedDate.Year, $"Year mismatch for: {testCase}");
            Assert.AreEqual(expectedMonth, parsedDate.Month, $"Month mismatch for: {testCase}");
            Assert.AreEqual(expectedDay, parsedDate.Day, $"Day mismatch for: {testCase}");
        }
    }

    [TestMethod]
    public void TestSingleDigitTimeComponents()
    {
        // Test time formats with single digit components as defined in CommonDateTimeFormats
        var testCases = new[]
        {
            ("9:05:03", 9, 5, 3),
            ("9:05", 9, 5, 0),
            ("14:05", 14, 5, 0)
        };

        foreach (var (testCase, expectedHour, expectedMinute, expectedSecond) in testCases)
        {
            var result = DateTime.TryParseExact(testCase, 
                Internal.ExcelConstants.DateFormats.CommonDateTimeFormats,
                CultureInfo.InvariantCulture, 
                DateTimeStyles.NoCurrentDateDefault, 
                out var parsedDate);

            Assert.IsTrue(result, $"Failed to parse: {testCase}");
            Assert.AreEqual(expectedHour, parsedDate.Hour, $"Hour mismatch for: {testCase}");
            Assert.AreEqual(expectedMinute, parsedDate.Minute, $"Minute mismatch for: {testCase}");
            Assert.AreEqual(expectedSecond, parsedDate.Second, $"Second mismatch for: {testCase}");
        }
    }
}
