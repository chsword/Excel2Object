using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object;

public class ExcelImporter
{
    private static readonly Dictionary<Type, Func<IRow, int, ImportContext, object?>> SpecialConvertDict =
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
        var context = new ImportContext(options);
        var result = GetDataRows(bytes, options);
        if (typeof(TModel) == typeof(Dictionary<string, object>))
            return (InternalExcelToDictionary(result, context) as IEnumerable<TModel>)!;

        var list = InternalExcelToObject<TModel>(result, context);
        return list;
    }

    public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes, string? sheetTitle)
        where TModel : class, new()
    {
        return ExcelToObject<TModel>(bytes, options => { options.SheetTitle = sheetTitle; });
    }

    private static IEnumerable<Dictionary<string, object>> InternalExcelToDictionary(IEnumerator<IRow>? result,
        ImportContext context)
    {
        var list = new List<Dictionary<string, object>>();

        if (result == null)
            return list;
        var rows = result;
        var titleRow = rows.Current;
        if (titleRow == null) return list;
        var columns = titleRow.Cells.ToDictionary(c => c.StringCellValue, c => c.ColumnIndex);

        while (rows.MoveNext())
        {
            var row = rows.Current;
            if (row == null || row.Cells?.Count == 0)
                continue;

            var model = new Dictionary<string, object>();

            foreach (var column in columns)
                model[column.Key] = GetCellValue(row.GetCell(column.Value), context, datesAsText: true) ?? "";

            list.Add(model);
        }

        return list;
    }

    private static IEnumerable<TModel> InternalExcelToObject<TModel>(IEnumerator<IRow>? result,
        ImportContext context)
        where TModel : class, new()
    {
        if (result == null)
            yield break;
            
        var dictColumns = BuildColumnMappings<TModel>(result);

        while (result.MoveNext())
        {
            var row = result.Current;

            if (row == null || row.Cells?.Count == 0)
                continue;

            var model = new TModel();
            PopulateModelFromRow(model, row, dictColumns, context);
            yield return model;
        }
    }

    private static Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>> BuildColumnMappings<TModel>(IEnumerator<IRow> result)
        where TModel : class, new()
    {
        var dict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
        var dictColumns = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>>();
        var titleRow = result.Current;
        
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
        Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>> dictColumns, ImportContext context)
        where TModel : class, new()
    {
        foreach (var pair in dictColumns)
        {
            var propType = pair.Value.Key.PropertyType;
            var type = TypeUtil.GetUnNullableType(propType);
            
            object? value = type.IsEnum
                ? GetEnum(row, pair.Key, type, context)
                : GetCellValueByType(row, pair.Key, propType, type, context);
                
            pair.Value.Key.SetValue(model, value, null);
        }
    }

    private static object? GetCellValueByType(IRow row, int columnIndex, Type propType, Type type,
        ImportContext context)
    {
        if (SpecialConvertDict.TryGetValue(type, out var special))
        {
            return special(row, columnIndex, context);
        }

        // a date cell reads as the date it shows into a string, and as the serial number Excel stores
        // into anything numeric
        var cellValue = GetCellValue(row.GetCell(columnIndex), context, type == typeof(string));

        // 读取失败（已上报）与读到空值不同：前者取该类型的默认值，不再尝试转换，否则会以
        // Convert.ChangeType("") 抛出的 FormatException 掩盖真正的原因，并把同一格上报两次
        if (cellValue == null) return DefaultOf(propType, type);

        if (cellValue.Length == 0
            && propType != typeof(string)
            && propType.IsGenericType
            && propType.GetGenericTypeDefinition() == typeof(Nullable<>))
            return null;

        try
        {
            return Convert.ChangeType(cellValue, type);
        }
        catch (Exception e)
        {
            // 转换失败照旧向外抛出并中止导入，上报只是让调用方知道是哪一个单元格
            context.Report(row, columnIndex, e);
            throw;
        }
    }

    /// <summary>该属性类型在读取失败时取的值：可空与引用类型取 null，其余取其默认值。</summary>
    private static object? DefaultOf(Type propType, Type type)
    {
        if (!propType.IsValueType) return null;
        if (propType.IsGenericType && propType.GetGenericTypeDefinition() == typeof(Nullable<>)) return null;
        return Activator.CreateInstance(type);
    }

    private static object? GetCellBoolean(IRow row, int key, ImportContext context)
    {
        var cellValue = GetCellValue(row.GetCell(key), context);
        if (string.IsNullOrEmpty(cellValue)) return null;
        if (bool.TryParse(cellValue, out var value)) return value;
        
        var lowerValue = cellValue.ToLower();
        if (ExcelConstants.BooleanValues.TrueValues.Any(v => v.Equals(lowerValue, StringComparison.OrdinalIgnoreCase)))
            return true;
        if (ExcelConstants.BooleanValues.FalseValues.Any(v => v.Equals(lowerValue, StringComparison.OrdinalIgnoreCase)))
            return false;
            
        try
        {
            return Convert.ToBoolean(cellValue);
        }
        catch (Exception e)
        {
            context.Report(row, key, e);
            throw;
        }
    }

    private static object? GetCellDateTime(IRow row, int index, ImportContext context)
    {
        var cell = row.GetCell(index);
        // 取文本的这一步自行上报失败，放在 try 之外，同一次失败才不会被上报两次
        if (string.IsNullOrEmpty(GetCellValue(cell, context))) return null;

        try
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    // 序列号超出 Excel 日历时 DateCellValue 会抛出，此时该单元格取 null
                    return cell.DateCellValue;
                case CellType.String:
                    var text = cell.StringCellValue;
                    var parsed = GetDateTimeFromString(text);
                    if (parsed == null && !string.IsNullOrWhiteSpace(text))
                        // 文本不是日期是最常见的导入失败，同样要让调用方知道
                        context.Report(cell, new FormatException($"[{text}] 不是可识别的日期。"));
                    return parsed;
                default:
                    return null;
            }
        }
        catch (Exception e)
        {
            context.Report(cell, e);
            return null;
        }
    }

    private static object? GetCellUri(IRow row, int key, ImportContext context)
    {
        var cellValue = GetCellValue(row.GetCell(key), context);
        if (string.IsNullOrEmpty(cellValue)) return null;

        try
        {
            return new Uri(cellValue);
        }
        catch (Exception e)
        {
            context.Report(row, key, e);
            throw;
        }
    }

    /// <param name="cell">The cell to read.</param>
    /// <param name="datesAsText">
    ///     Whether a date cell reads as the date it shows rather than as the serial number Excel stores;
    ///     what a string property or a dictionary wants, and what a numeric property cannot parse.
    /// </param>
    /// <returns>
    ///     单元格的文本；<c>null</c> 表示读取失败——失败已经上报，调用方据此取默认值即可，不应再
    ///     按空字符串继续转换。
    /// </returns>
    private static string? GetCellValue(ICell? cell, ImportContext context, bool datesAsText = false)
    {
        var result = string.Empty;
        if (cell == null) return result;
        try
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    result = (datesAsText ? DateCellText(cell) : null)
                             ?? cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    break;
                case CellType.String:
                    result = cell.StringCellValue;
                    break;
                case CellType.Blank:
                    result = string.Empty;
                    break;
                case CellType.Formula:
                    // 求值器按工作簿创建一次，不再逐单元格创建
                    result = GetCellValue(context.Evaluator(cell.Sheet.Workbook).EvaluateInCell(cell), context,
                        datesAsText);
                    break;
                default:
                    // Boolean ("TRUE"/"FALSE"), Error and _None all render acceptably through ToString.
                    result = cell.ToString();
                    break;
            }
        }
        catch (Exception e)
        {
            context.Report(cell, e);
            return null;
        }

        return (result ?? "").Trim();
    }

    /// <summary>
    ///     What each number format seen so far displays, so NPOI's format parsing runs once per format
    ///     rather than once per cell. <see cref="ExcelDateFormat.Parts.None" /> marks a format that is no
    ///     date at all.
    /// </summary>
    private static readonly ConcurrentDictionary<string, ExcelDateFormat.Parts> FormatParts = new();

    /// <summary>
    ///     A date cell is a number with a date format. Read as text - into a string property or a
    ///     dictionary - it comes out as the date it shows, not as the serial number Excel stores: an ISO
    ///     date, a time of day or both, whichever its format displays. Null when the cell is no date;
    ///     elapsed time (<c>[h]:mm</c>) counts as none, the number being the best that can be offered then.
    /// </summary>
    private static string? DateCellText(ICell cell)
    {
        var style = cell.CellStyle;
        var format = style?.GetDataFormatString();
        if (format == null) return null;

        if (!FormatParts.TryGetValue(format, out var parts))
        {
            parts = DateUtil.IsADateFormat(style!.DataFormat, format)
                ? ExcelDateFormat.PartsShown(format)
                : ExcelDateFormat.Parts.None;
            FormatParts[format] = parts;
        }

        if (parts == ExcelDateFormat.Parts.None || !DateUtil.IsValidExcelDate(cell.NumericCellValue))
            return null;
        var date = cell.DateCellValue;
        if (date == null) return null;

        var pattern = parts == ExcelDateFormat.Parts.Date ? ExcelDateFormat.IsoDate :
            parts == ExcelDateFormat.Parts.Time ? ExcelDateFormat.IsoTime : ExcelDateFormat.IsoDateTime;
        return date.Value.ToString(pattern, CultureInfo.InvariantCulture);
    }

    private static IEnumerator<IRow>? GetDataRows(byte[]? bytes, ExcelImporterOptions options)
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

        var rows = sheet.GetEnumerator();
        rows.MoveNext();
        for (var i = 0; i < options.TitleSkipLine; i++) rows.MoveNext();
        return rows;
    }

    private static DateTime? GetDateTimeFromString(string str)
    {
        DateTime dt;
        
        // Handle Chinese date formats (年月日)
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
            // Try standard parsing first
            if (DateTime.TryParse(str, out dt))
                return dt;
            
            // Try parsing with specific formats
            if (DateTime.TryParseExact(str, ExcelConstants.DateFormats.CommonDateTimeFormats, 
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt;
            
            // Try parsing with current culture
            if (DateTime.TryParseExact(str, ExcelConstants.DateFormats.CommonDateTimeFormats, 
                CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
                return dt;
            
            // Handle time-only formats - combine with today's date
            if (DateTime.TryParseExact(str, ExcelConstants.DateFormats.CommonDateTimeFormats, 
                CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out dt))
            {
                // If only time is provided, combine with today's date
                if (dt.Date == DateTime.MinValue.Date)
                {
                    return DateTime.Today.Add(dt.TimeOfDay);
                }
                return dt;
            }
            
            // Fallback for partial dates
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

    private static object? GetEnum(IRow row, int key, Type enumType, ImportContext context)
    {
        var cellValue = GetCellValue(row.GetCell(key), context);
        if (string.IsNullOrEmpty(cellValue)) return null;
        if (Enum.GetNames(enumType).Contains(cellValue)) return Enum.Parse(enumType, cellValue);

        // 取值不在枚举中：沿用既有行为取 0，但不再悄无声息
        context.Report(row, key,
            new FormatException($"[{cellValue}] 不是 {enumType.Name} 的取值，按 0 处理。"));
        return Enum.ToObject(enumType, 0);
    }
}