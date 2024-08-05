using System;

// ReSharper disable All

namespace Chsword.Excel2Object.Tests.Models;

[ExcelTitle("Test Person")]
public class TestModelPerson
{
    [ExcelTitle("姓名")] public string Name { get; set; }
    [ExcelTitle("年龄")] public int? Age { get; set; }
    [ExcelTitle("出生日期")] public DateTime? Birthday { get; set; }
}

[ExcelTitle("Test Strict Person")]
public class TestModelStrictPerson
{
    [ExcelTitle("姓名")] public string Name { get; set; }
    [ExcelTitle("年龄")] public int Age { get; set; }
    [ExcelTitle("出生日期")] public DateTime Birthday { get; set; }
}

[ExcelTitle("Test Date Person")]
public class TestModelDatePerson
{
    [ExcelTitle("姓名")] public string Name { get; set; }
    [ExcelTitle("年龄")] public int Age { get; set; }

    [ExcelColumn("出生日期", Format = "yyyy-MM-dd HH:mm:ss")]
    public DateTime Birthday { get; set; }

    [ExcelColumn("出生日期2", Format = "yyyy-MM-dd HH:mm:ss")]
    public DateTime? Birthday2 { get; set; }
}