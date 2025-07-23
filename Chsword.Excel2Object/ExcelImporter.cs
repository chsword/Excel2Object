using System.Collections;
using System.Globalization;
using System.Reflection;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object;

public class ExcelImporter
{
    private static readonly Dictionary<Type, Func<IRow, int, object>> SpecialConvertDict =
        new()
        {
            [typeof(DateTime)] = GetCellDateTime,
            [typeof(bool)] = GetCellBoolean,
            [typeof(Uri)] = GetCellUri
        };

    public IEnumerable<TModel>? ExcelToObject<TModel>(string path, string? sheetTitle)
        where TModel : class, new()
    {
        return ExcelToObject<TModel>(path, options => { options.SheetTitle = sheetTitle; });
    }

    public IEnumerable<TModel>? ExcelToObject<TModel>(string path,
        Action<ExcelImporterOptions>? optionAction = null)
        where TModel : class, new()
    {
        if (string.IsNullOrWhiteSpace(path))
            return null;
        var bytes = File.ReadAllBytes(path);
        return ExcelToObject<TModel>(bytes, optionAction);
    }

    public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes,
        Action<ExcelImporterOptions>? optionAction = null)
        where TModel : class, new()
    {
        var options = new ExcelImporterOptions();
        optionAction?.Invoke(options);
        var result = GetDataRows(bytes, options);
        if (typeof(TModel) == typeof(Dictionary<string, object>))
            return (InternalExcelToDictionary(result) as IEnumerable<TModel>)!;

        var list = InternalExcelToObject<TModel>(result);
        return list;
    }

    public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes, string? sheetTitle)
        where TModel : class, new()
    {
        return ExcelToObject<TModel>(bytes, options => { options.SheetTitle = sheetTitle; });
    }

    private static IEnumerable<Dictionary<string, object>> InternalExcelToDictionary(IEnumerator? result)
    {
        var list = new List<Dictionary<string, object>>();

        if (result == null)
            return list;
        var rows = result;
        var titleRow = (IRow) rows.Current;
        if (titleRow == null) return list;
        var columns = titleRow.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);

        while (rows.MoveNext())
        {
            var row = (IRow) rows.Current;
            if (row == null || row.Cells?.Count == 0)
                continue;

            var model = new Dictionary<string, object>();

            foreach (var column in columns) model[column.Key] = GetCellValue(row, column.Value);

            list.Add(model);
        }

        return list;
    }

    private static IEnumerable<TModel> InternalExcelToObject<TModel>(IEnumerator? result)
        where TModel : class, new()
    {
        if (result == null)
            yield break;
            
        var dictColumns = BuildColumnMappings<TModel>(result);

        while (result.MoveNext())
        {
            var row = (IRow) result.Current;

            if (row == null || row.Cells?.Count == 0)
                continue;

            var model = new TModel();
            PopulateModelFromRow(model, row, dictColumns);
            yield return model;
        }
    }

    private static Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>> BuildColumnMappings<TModel>(IEnumerator result)
        where TModel : class, new()
    {
        var dict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
        var dictColumns = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>>();
        var titleRow = (IRow) result.Current;
        
        if (titleRow != null)
            foreach (var cell in titleRow.Cells)
            {
                var prop = dict.FirstOrDefault(c => cell.StringCellValue == c.Value.Title);
                if (prop.Key != null && !dictColumns.ContainsKey(cell.ColumnIndex))
                    dictColumns.Add(cell.ColumnIndex, prop);
            }
            
        return dictColumns;
    }

    private static void PopulateModelFromRow<TModel>(TModel model, IRow row, 
        Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>> dictColumns)
        where TModel : class, new()
    {
        foreach (var pair in dictColumns)
        {
            var propType = pair.Value.Key.PropertyType;
            var type = TypeUtil.GetUnNullableType(propType);
            
            object? value = type.IsEnum 
                ? GetEnum(row, pair.Key, type)
                : GetCellValueByType(row, pair.Key, propType, type);
                
            pair.Value.Key.SetValue(model, value, null);
        }
    }

    private static object? GetCellValueByType(IRow row, int columnIndex, Type propType, Type type)
    {
        if (SpecialConvertDict.ContainsKey(type))
        {
            return SpecialConvertDict[type](row, columnIndex);
        }

        var cellValue = GetCellValue(row, columnIndex);
        if (string.IsNullOrEmpty(cellValue)
            && propType != typeof(string)
            && propType.IsGenericType
            && propType.GetGenericTypeDefinition() == typeof(Nullable<>))
            return null;
            
        return Convert.ChangeType(cellValue, type);
    }

    private static object? GetCellBoolean(IRow row, int key)
    {
        var cellValue = GetCellValue(row, key);
        if (string.IsNullOrEmpty(cellValue)) return null;
        if (bool.TryParse(cellValue, out var value)) return value;
        
        var lowerValue = cellValue.ToLower();
        if (ExcelConstants.BooleanValues.TrueValues.Any(v => v.Equals(lowerValue, StringComparison.OrdinalIgnoreCase)))
            return true;
        if (ExcelConstants.BooleanValues.FalseValues.Any(v => v.Equals(lowerValue, StringComparison.OrdinalIgnoreCase)))
            return false;
            
        return Convert.ToBoolean(cellValue);
    }

    private static object? GetCellDateTime(IRow row, int index)
    {
        DateTime? result = null;
        try
        {
            var cell = row.GetCell(index);

            var cellValue = GetCellValue(cell);
            if (string.IsNullOrEmpty(cellValue)) return null;

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    try
                    {
                        result = cell.DateCellValue;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }

                    break;
                case CellType.String:
                    var str = cell.StringCellValue;
                    result = GetDateTimeFromString(str);
                    break;
                case CellType.Blank:
                    break;
                case CellType.Unknown:
                    break;
                case CellType.Formula:
                    break;
                case CellType.Boolean:
                    break;
                case CellType.Error:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
        }

        return result;
    }

    private static object? GetCellUri(IRow row, int key)
    {
        var cellValue = GetCellValue(row, key);
        return string.IsNullOrEmpty(cellValue) ? null : new Uri(cellValue);
    }

    private static string GetCellValue(ICell? cell)
    {
        var result = string.Empty;
        if (cell == null) return result;
        try
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    result = cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    break;
                case CellType.String:
                    result = cell.StringCellValue;
                    break;
                case CellType.Blank:
                    result = string.Empty;
                    break;
                case CellType.Formula:

                    var e = WorkbookFactory.CreateFormulaEvaluator(cell.Sheet.Workbook);
                    result = GetCellValue(e.EvaluateInCell(cell));
                    //result = e.EvaluateInCell(row.GetCell(index)).StringCellValue;
                    break;
                //case CellType.Boolean:
                //    result = row.GetCell(index).NumericCellValue.ToString();
                //    break;
                //case CellType.Error:
                //    result = row.GetCell(index).NumericCellValue.ToString();
                //    break;
                //case CellType.Unknown:
                //    result = row.GetCell(index).NumericCellValue.ToString();
                //    break;
                default:
                    result = cell.ToString();
                    break;
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
        }

        return (result ?? "").Trim();
    }

    private static string GetCellValue(IRow row, int index)
    {
        return GetCellValue(row.GetCell(index));
    }

    private static IEnumerator? GetDataRows(byte[]? bytes, ExcelImporterOptions options)
    {
        if (bytes == null || bytes.Length == 0)
            return null;
        IWorkbook workbook;
        try
        {
            using var memoryStream = new MemoryStream(bytes);
            workbook = WorkbookFactory.Create(memoryStream);
        }
        catch
        {
            return null;
        }

        ISheet sheet;
        if (string.IsNullOrEmpty(options.SheetTitle))
        {
            sheet = workbook.GetSheetAt(0);
        }
        else
        {
            sheet = workbook.GetSheet(options.SheetTitle);
            if (sheet == null)
                throw new Excel2ObjectException($"The specified sheet:[{options.SheetTitle}] does not exist");
        }

        var rows = sheet.GetRowEnumerator();
        rows.MoveNext();
        for (var i = 0; i < options.TitleSkipLine; i++) rows.MoveNext();
        return rows;
    }

    private static DateTime? GetDateTimeFromString(string str)
    {
        DateTime dt;
        if (str.EndsWith(ExcelConstants.DateFormats.YearSuffix))
        {
            if (DateTime.TryParse((str + ExcelConstants.DateFormats.DefaultYearMonthSuffix).Replace(ExcelConstants.DateFormats.YearSuffix, ""), out dt))
                return dt;
        }
        else if (str.EndsWith(ExcelConstants.DateFormats.MonthSuffix))
        {
            if (DateTime.TryParse((str + ExcelConstants.DateFormats.DefaultDaySuffix).Replace(ExcelConstants.DateFormats.YearSuffix, "").Replace(ExcelConstants.DateFormats.MonthSuffix, ""), out dt))
                return dt;
        }
        else if (!str.Contains(ExcelConstants.DateFormats.YearSuffix) && !str.Contains(ExcelConstants.DateFormats.MonthSuffix) && !str.Contains(ExcelConstants.DateFormats.DaySuffix))
        {
            if (DateTime.TryParse(str, out dt))
                return dt;
            if (DateTime.TryParse((str + ExcelConstants.DateFormats.DefaultYearMonthSuffix).Replace(ExcelConstants.DateFormats.YearSuffix, "").Replace(ExcelConstants.DateFormats.MonthSuffix, ""), out dt))
                return dt;
        }
        else
        {
            if (DateTime.TryParse(str.Replace(ExcelConstants.DateFormats.YearSuffix, "").Replace(ExcelConstants.DateFormats.MonthSuffix, ""), out dt))
                return dt;
        }

        return null;
    }

    private static object? GetEnum(IRow row, int key, Type enumType)
    {
        var cellValue = GetCellValue(row, key);
        if (string.IsNullOrEmpty(cellValue)) return null;
        if (Enum.GetNames(enumType).Contains(cellValue)) return Enum.Parse(enumType, cellValue);

        return Enum.Parse(enumType, "0");
    }
}