using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Chsword.Excel2Object.Functions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Chsword.Excel2Object.Tests;

/// <summary>
///     Every method of a function interface has to spell a name Excel would accept, since the
///     translator writes the method name (or its attribute) straight into the formula.
/// </summary>
[TestClass]
public class ExcelFunctionNameTests
{
    private static readonly Regex ValidName =
        new("^(_xlfn\\.(_xlws\\.)?)?[A-Z][A-Z0-9]*(\\.[A-Z][A-Z0-9]*)*$");

    private static MethodInfo[] FunctionMethods()
    {
        return typeof(IExcelFunction).Assembly.GetTypes()
            .Where(t => t.IsInterface && t != typeof(IExcelFunction) && typeof(IExcelFunction).IsAssignableFrom(t))
            .SelectMany(t => t.GetMethods())
            .ToArray();
    }

    private static string NameOf(MethodInfo method)
    {
        return method.GetCustomAttribute<ExcelFunctionNameAttribute>()?.StoredName ??
               method.Name.ToUpperInvariant();
    }

    [TestMethod]
    public void EveryFunctionNameLooksLikeAnExcelFunction()
    {
        foreach (var method in FunctionMethods())
            Assert.IsTrue(ValidName.IsMatch(NameOf(method)),
                $"{method.DeclaringType?.Name}.{method.Name} maps to \"{NameOf(method)}\"");
    }

    [TestMethod]
    public void MethodsWhoseNameCannotSpellTheFunctionCarryTheAttribute()
    {
        foreach (var method in FunctionMethods().Where(m => NameOf(m).Contains('.')))
            Assert.IsNotNull(method.GetCustomAttribute<ExcelFunctionNameAttribute>());
    }

    [TestMethod]
    public void FunctionsAddedAfterExcel2007AreStoredWithTheirPrefix()
    {
        // A workbook that stores the bare name of one of these shows #NAME? instead of a result.
        var stored = FunctionMethods()
            .GroupBy(m => $"{m.DeclaringType?.Name}.{m.Name}/{m.GetParameters().Length}")
            .ToDictionary(g => g.Key, g => NameOf(g.First()));
        Assert.AreEqual("_xlfn.IFS", stored["IConditionFunction.Ifs/1"]);
        Assert.AreEqual("_xlfn.STDEV.P", stored["IStatisticsFunction.StDevP/1"]);
        Assert.AreEqual("_xlfn.XLOOKUP", stored["IReferenceFunction.XLookup/3"]);
        Assert.AreEqual("_xlfn._xlws.FILTER", stored["IReferenceFunction.Filter/2"]);
        Assert.AreEqual("_xlfn.TEXTJOIN", stored["ITextFunction.TextJoin/3"]);
        Assert.AreEqual("_xlfn.DAYS", stored["IDateTimeFunction.Days/2"]);
        // ... while the functions Excel 2007 already had keep their bare name.
        Assert.AreEqual("SUMIF", stored["IMathFunction.SumIf/2"]);
        Assert.AreEqual("VLOOKUP", stored["IReferenceFunction.VLookup/4"]);
        Assert.AreEqual("STDEV", stored["IStatisticsFunction.StDev/1"]);
    }

    [TestMethod]
    public void AllCategoriesAreReachableFromExcelFunctions()
    {
        var exposed = typeof(ExcelFunctions).GetProperties()
            .Select(p => p.PropertyType)
            .ToArray();
        foreach (var type in typeof(IExcelFunction).Assembly.GetTypes()
                     .Where(t => t.IsInterface && t != typeof(IExcelFunction) &&
                                 typeof(IExcelFunction).IsAssignableFrom(t)))
            Assert.IsTrue(exposed.Contains(type), $"{type.Name} is not exposed on ExcelFunctions");
    }

    [TestMethod]
    public void AllFunctionInterfaceCoversEveryCategory()
    {
        foreach (var type in typeof(IExcelFunction).Assembly.GetTypes()
                     .Where(t => t.IsInterface && t != typeof(IExcelFunction) && t != typeof(IAllFunction) &&
                                 typeof(IExcelFunction).IsAssignableFrom(t)))
            Assert.IsTrue(type.IsAssignableFrom(typeof(IAllFunction)),
                $"IAllFunction does not include {type.Name}");
    }
}
