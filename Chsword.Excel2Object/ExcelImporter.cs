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
            
        if (!File.Exists(path))
            throw new FileNotFoundException($"Excel file not found: {path}");
            
        try
        {
            var bytes = File.ReadAllBytes(path);
            return ExcelToObject<TModel>(bytes, optionAction);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new Excel2ObjectException($"Access denied reading file: {path}", ex);
        }
        catch (IOException ex)
        {
            throw new Excel2ObjectException($"IO error reading file: {path}", ex);
        }
    }

    public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes,
        Action<ExcelImporterOptions>? optionAction = null)
        where TModel : class, new()
    {
        if (bytes == null)
            throw new ArgumentNullException(nameof(bytes));
            
        if (bytes.Length == 0)
            throw new ArgumentException("Byte array cannot be empty.", nameof(bytes));
            
        if (!ValidateExcelFileFormat(bytes))
            throw new Excel2ObjectException("Invalid Excel file format. Only .xls and .xlsx files are supported.");
        
        var options = new ExcelImporterOptions();
        optionAction?.Invoke(options);
        var result = GetDataRows(bytes, options);
        if (typeof(TModel) == typeof(Dictionary<string, object>))
            return (InternalExcelToDictionary(result) as IEnumerable<TModel>)!;;

        var list = InternalExcelToObject<TModel>(result);
        return list;
    }

    public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes, string? sheetTitle)
        where TModel : class, new()
    {
        return ExcelToObject<TModel>(bytes, options => { options.SheetTitle = sheetTitle; });
    }

    /// <summary>
    /// Validates if the specified file is a valid Excel file
    /// </summary>
    /// <param name="filePath">Path to the file to validate</param>
    /// <returns>True if the file is a valid Excel file, false otherwise</returns>
    public static bool IsValidExcelFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            return false;
            
        try
        {
            var bytes = File.ReadAllBytes(filePath);
            return IsValidExcelFile(bytes);
        }
        catch
        {
            return false;
        }
    }
    
    /// <summary>
    /// Validates if the byte array represents a valid Excel file
    /// </summary>
    /// <param name="bytes">The byte array to validate</param>
    /// <returns>True if the data appears to be a valid Excel file, false otherwise</returns>
    public static bool IsValidExcelFile(byte[] bytes)
    {
        return ValidateExcelFileFormat(bytes);
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
        
        // Use StringComparison.OrdinalIgnoreCase instead of ToLower() for better performance
        if (ExcelConstants.BooleanValues.TrueValues.Any(v => v.Equals(cellValue, StringComparison.OrdinalIgnoreCase)))
            return true;
        if (ExcelConstants.BooleanValues.FalseValues.Any(v => v.Equals(cellValue, StringComparison.OrdinalIgnoreCase)))
            return false;
            
        return Convert.ToBoolean(cellValue);
    }

    private static object? GetCellDateTime(IRow row, int index)
    {
        try
        {
            var cell = row.GetCell(index);
            var cellValue = GetCellValue(cell);
            if (string.IsNullOrEmpty(cellValue)) return null;

            return cell.CellType switch
            {
                CellType.Numeric => TryGetDateCellValue(cell),
                CellType.String => GetDateTimeFromString(cell.StringCellValue),
                CellType.Blank or CellType.Unknown or CellType.Formula or CellType.Boolean or CellType.Error => null,
                _ => null
            };
        }
        catch (InvalidOperationException)
        {
            // Cell contains invalid date format
            return null;
        }
        catch (FormatException)
        {
            // Cell value cannot be converted to DateTime
            return null;
        }
        catch (Exception)
        {
            // Any other unexpected error - return null for robustness
            return null;
        }
    }

    private static DateTime? TryGetDateCellValue(ICell cell)
    {
        try
        {
            return cell.DateCellValue;
        }
        catch (Exception)
        {
            // If it's not a valid date, return null
            return null;
        }
    }

    private static object? GetCellUri(IRow row, int key)
    {
        var cellValue = GetCellValue(row, key);
        return string.IsNullOrEmpty(cellValue) ? null : new Uri(cellValue);
    }

    private static string GetCellValue(ICell? cell)
    {
        if (cell == null) return string.Empty;
        
        try
        {
            return cell.CellType switch
            {
                CellType.Numeric => cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
                CellType.String => cell.StringCellValue,
                CellType.Blank => string.Empty,
                CellType.Formula => EvaluateFormula(cell),
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Error => ExcelConstants.CellTypes.Text, // Return placeholder for error cells
                _ => cell.ToString() ?? string.Empty
            };
        }
        catch (InvalidOperationException)
        {
            // Cell type mismatch or invalid operation
            return cell.ToString() ?? string.Empty;
        }
        catch (InvalidCastException)
        {
            // Type conversion error
            return cell.ToString() ?? string.Empty;
        }
        catch (Exception)
        {
            // Any other error - return empty string for robustness
            return string.Empty;
        }
    }

    private static string EvaluateFormula(ICell cell)
    {
        try
        {
            var evaluator = WorkbookFactory.CreateFormulaEvaluator(cell.Sheet.Workbook);
            var evaluatedCell = evaluator.EvaluateInCell(cell);
            return GetCellValue(evaluatedCell).Trim();
        }
        catch (Exception)
        {
            // If formula evaluation fails, return the raw formula
            return cell.CellFormula ?? string.Empty;
        }
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
        
        try
        {
            // Try exact match first
            if (Enum.GetNames(enumType).Contains(cellValue))
                return Enum.Parse(enumType, cellValue);
            
            // Try case-insensitive match using Enum.Parse with ignoreCase parameter
            return Enum.Parse(enumType, cellValue, true);
        }
        catch (ArgumentException)
        {
            // Try parsing as integer value if string parsing fails
            if (int.TryParse(cellValue, out var intValue) && Enum.IsDefined(enumType, intValue))
                return Enum.ToObject(enumType, intValue);
            
            // Default to first enum value (0) if no match found
            var enumValues = Enum.GetValues(enumType);
            return enumValues.Length > 0 ? enumValues.GetValue(0) : null;
        }
    }
    
    /// <summary>
    /// Validates if the byte array represents a valid Excel file (.xls or .xlsx)
    /// </summary>
    /// <param name="bytes">The byte array to validate</param>
    /// <returns>True if the file appears to be a valid Excel file, false otherwise</returns>
    private static bool ValidateExcelFileFormat(byte[] bytes)
    {
        if (bytes.Length < 8) return false;
        
        // Check for Excel file signatures
        // XLSX (Office Open XML) starts with PK (ZIP file)
        if (bytes[0] == 0x50 && bytes[1] == 0x4B)
        {
            return true; // XLSX file
        }
        
        // XLS (BIFF8) file signature
        if (bytes.Length >= 512 && 
            bytes[0] == 0xD0 && bytes[1] == 0xCF && 
            bytes[2] == 0x11 && bytes[3] == 0xE0 && 
            bytes[4] == 0xA1 && bytes[5] == 0xB1 && 
            bytes[6] == 0x1A && bytes[7] == 0xE1)
        {
            return true; // XLS file
        }
        
        return false;
    }
}