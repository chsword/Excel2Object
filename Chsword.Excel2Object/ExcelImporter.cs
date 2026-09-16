using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using NPOI.SS.UserModel;

namespace Chsword.Excel2Object;

public class ExcelImporter
{
    private static readonly Dictionary<Type, Func<IImportRow, int, ImportContext, object?>> SpecialConvertDict =
        new()
        {
            [typeof(DateTime)] = GetCellDateTime,
            [typeof(bool)] = GetCellBoolean,
            [typeof(Uri)] = GetCellUri
        };

    /// <summary>
    ///     What each number format seen so far displays, so NPOI's format parsing runs once per format
    ///     rather than once per cell. <see cref="ExcelDateFormat.Parts.None" /> marks a format that is no
    ///     date at all.
    /// </summary>
    private static readonly ConcurrentDictionary<string, ExcelDateFormat.Parts> FormatParts = new();

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
        var source = GetDataRows(bytes, options, context);
        return ToModels<TModel>(source == null ? null : AtHeader(source, options), context);
    }

    public IEnumerable<TModel> ExcelToObject<TModel>(byte[] bytes, string? sheetTitle)
        where TModel : class, new()
    {
        return ExcelToObject<TModel>(bytes, options => { options.SheetTitle = sheetTitle; });
    }

    /// <summary>
    ///     逐行读出 <paramref name="input" /> 中的一张工作表，不把整个工作簿建进内存，适用于行数多到
    ///     不宜整份读入的导入。
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         取到的序列是惰性的：取一行、转一行，读过的行随即可以回收。留在内存里的只有共享字符串表
    ///         与样式表——前者是文件中所有不重复的字符串，后者通常只有几十条。
    ///     </para>
    ///     <para>
    ///         与整份读入相比有一处不同：<strong>公式格读的是文件中存着的上一次计算结果</strong>，而非
    ///         当场求值。Excel 存盘时会写下这个结果，本库导出的文件则没有（要等 Excel 打开时才算出），
    ///         此时该格读作空白，与其他空白格一样。
    ///     </para>
    ///     <para>
    ///         只有 <c>.xlsx</c> 能够逐行读出；传入 <c>.xls</c> 时会照旧整份读入内存，结果一致。
    ///         <paramref name="input" /> 由本方法读取，遍历结束（或中途放弃）时关闭其上的工作簿；与
    ///         <see cref="File.ReadLines(string)" /> 一样，取到的序列须被遍历。
    ///     </para>
    /// </remarks>
    /// <example>
    ///     <code>
    /// using var file = File.OpenRead("orders.xlsx");
    /// foreach (var order in new ExcelImporter().ExcelStreamToObject&lt;Order&gt;(file))
    ///     Handle(order);
    ///     </code>
    /// </example>
    public IEnumerable<TModel> ExcelStreamToObject<TModel>(Stream input,
        Action<ExcelImporterOptions>? optionAction = null)
        where TModel : class, new()
    {
        // .xls 的格式无从逐行读出，整份读入后结果与既有方式一致。此处先于建立选项判断，
        // optionAction 才不会被调用两次——它可能是有状态的
        if (!LooksLikeXlsx(input)) return ExcelToObject<TModel>(ReadAll(input), optionAction);

        var options = new ExcelImporterOptions();
        optionAction?.Invoke(options);
        var context = new ImportContext(options);
        return ToModels<TModel>(AtHeader(XlsxRowReader.Rows(input, options, context), options), context);
    }

    /// <summary>
    ///     只读出表头：表名与各列的标题，按列的先后。用于在导入之前核对列，或据表头生成模型。
    /// </summary>
    /// <remarks>
    ///     只读到表头那一行为止，后面有多少行数据都不影响其开销。<c>.xls</c> 仍须整份读入——该格式
    ///     的数据并非顺序存放。传入的流由本方法读取，返回前即已读完。
    /// </remarks>
    /// <example>
    ///     <code>
    /// using var file = File.OpenRead("orders.xlsx");
    /// var header = ExcelHelper.ReadHeader(file);
    /// Console.WriteLine($"{header.SheetTitle}：{string.Join("、", header.Columns)}");
    ///     </code>
    /// </example>
    public ExcelSheetHeader ReadHeader(Stream input, Action<ExcelImporterOptions>? optionAction = null)
    {
        var options = new ExcelImporterOptions();
        optionAction?.Invoke(options);
        var context = new ImportContext(options);

        var source = LooksLikeXlsx(input)
            ? XlsxRowReader.Rows(input, options, context)
            : GetDataRows(ReadAll(input), options, context);
        if (source == null) return new ExcelSheetHeader(null, new List<string>());

        using var rows = AtHeader(source, options);
        var titleRow = rows.Current;
        var columns = new List<string>();
        if (titleRow != null)
            foreach (var cell in titleRow.Cells)
                columns.Add(TextOf(cell.Value) ?? string.Empty);

        return new ExcelSheetHeader(source.Title, columns);
    }

    /// <summary>
    ///     <c>.xlsx</c> 是个 zip，头两个字节为 <c>PK</c>。流不可定位时无从先看一眼，此时按 .xlsx 处理，
    ///     真不是的话打开那一步会说清楚。
    /// </summary>
    private static bool LooksLikeXlsx(Stream input)
    {
        if (!input.CanSeek) return true;

        var position = input.Position;
        var head = new byte[2];
        var read = input.Read(head, 0, 2);
        input.Position = position;
        return read == 2 && head[0] == (byte) 'P' && head[1] == (byte) 'K';
    }

    private static byte[] ReadAll(Stream input)
    {
        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>行从哪里来并不影响其后的转换：字典与模型两条路都只认 <see cref="IImportRow" />。</summary>
    private static IEnumerable<TModel> ToModels<TModel>(IEnumerator<IImportRow>? rows, ImportContext context)
        where TModel : class, new()
    {
        if (typeof(TModel) == typeof(Dictionary<string, object>))
            return (InternalExcelToDictionary(rows, context) as IEnumerable<TModel>)!;

        return InternalExcelToObject<TModel>(rows, context);
    }

    private static IEnumerable<Dictionary<string, object>> InternalExcelToDictionary(IEnumerator<IImportRow>? result,
        ImportContext context)
    {
        if (result == null) yield break;

        // 取行的枚举器在此释放：流式导入由它持有着打开的工作簿，中途放弃（Take、break）时也须关上
        using (result)
        {
            var titleRow = result.Current;
            if (titleRow == null) yield break;

            // 同名的标题以最左边那一列为准，与模型列的对应方式一致
            var columns = new Dictionary<string, int>();
            foreach (var cell in titleRow.Cells)
            {
                var title = TextOf(cell.Value);
                if (title != null && !columns.ContainsKey(title)) columns[title] = cell.Key;
            }

            while (result.MoveNext())
            {
                var row = result.Current;
                if (row == null || row.CellCount == 0)
                    continue;

                var model = new Dictionary<string, object>();

                foreach (var column in columns)
                    model[column.Key] = TextOf(row.Cell(column.Value), true) ?? "";

                yield return model;
            }
        }
    }

    private static IEnumerable<TModel> InternalExcelToObject<TModel>(IEnumerator<IImportRow>? result,
        ImportContext context)
        where TModel : class, new()
    {
        if (result == null)
            yield break;

        // 取行的枚举器在此释放：流式导入由它持有着打开的工作簿，中途放弃（Take、break）时也须关上
        using (result)
        {
            var dictColumns = BuildColumnMappings<TModel>(result, context);

            while (result.MoveNext())
            {
                var row = result.Current;

                if (row == null || row.CellCount == 0)
                    continue;

                var model = new TModel();
                PopulateModelFromRow(model, row, dictColumns, context);
                yield return model;
            }
        }
    }

    private static Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>> BuildColumnMappings<TModel>(
        IEnumerator<IImportRow> result, ImportContext context)
        where TModel : class, new()
    {
        var dict = ExcelUtil.GetPropertiesAttributesDict<TModel>();
        var dictColumns = new Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>>();
        var titleRow = result.Current;
        if (titleRow == null) return dictColumns;

        var headerTitles = new List<string>();
        foreach (var cell in titleRow.Cells)
        {
            var title = TextOf(cell.Value) ?? string.Empty;
            headerTitles.Add(title);
            var prop = dict.FirstOrDefault(c => title == c.Value.Title);
            if (prop.Key != null && !dictColumns.ContainsKey(cell.Key))
                dictColumns.Add(cell.Key, prop);
        }

        // 模型上写着、表头里却没有的标题：那一列不会被填上，整列都是默认值，此处上报
        var mapped = new HashSet<string>(dictColumns.Values.Select(c => c.Value.Title), StringComparer.Ordinal);
        foreach (var pair in dict)
            if (!mapped.Contains(pair.Value.Title))
                context.ReportMissingColumn(pair.Value.Title, pair.Key.Name, titleRow.SheetTitle, headerTitles);

        return dictColumns;
    }

    private static void PopulateModelFromRow<TModel>(TModel model, IImportRow row,
        Dictionary<int, KeyValuePair<PropertyInfo, ExcelTitleAttribute>> dictColumns, ImportContext context)
        where TModel : class, new()
    {
        foreach (var pair in dictColumns)
        {
            var propType = pair.Value.Key.PropertyType;
            var type = TypeUtil.GetUnNullableType(propType);

            var value = type.IsEnum
                ? GetEnum(row, pair.Key, type, context)
                : GetCellValueByType(row, pair.Key, propType, type, context);

            pair.Value.Key.SetValue(model, value, null);
        }
    }

    private static object? GetCellValueByType(IImportRow row, int columnIndex, Type propType, Type type,
        ImportContext context)
    {
        if (SpecialConvertDict.TryGetValue(type, out var special)) return special(row, columnIndex, context);

        // a date cell reads as the date it shows into a string, and as the serial number Excel stores
        // into anything numeric
        var cellValue = TextOf(row.Cell(columnIndex), type == typeof(string));

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
            return System.Convert.ChangeType(cellValue, type);
        }
        catch (Exception e)
        {
            // 转换失败照旧向外抛出并中止导入，上报只是让调用方知道是哪一个单元格
            context.Report(row.SheetTitle, row.RowIndex, columnIndex, e);
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

    private static object? GetCellBoolean(IImportRow row, int key, ImportContext context)
    {
        var cellValue = TextOf(row.Cell(key));
        // 显式判空而非 string.IsNullOrEmpty：后者在较早的目标框架上没有 NotNullWhen 标注
        if (cellValue == null || cellValue.Length == 0) return null;
        if (bool.TryParse(cellValue, out var value)) return value;

        var lowerValue = cellValue.ToLower();
        if (ExcelConstants.BooleanValues.TrueValues.Any(v => v.Equals(lowerValue, StringComparison.OrdinalIgnoreCase)))
            return true;
        if (ExcelConstants.BooleanValues.FalseValues.Any(v => v.Equals(lowerValue, StringComparison.OrdinalIgnoreCase)))
            return false;

        try
        {
            return System.Convert.ToBoolean(cellValue);
        }
        catch (Exception e)
        {
            context.Report(row.SheetTitle, row.RowIndex, key, e);
            throw;
        }
    }

    private static object? GetCellDateTime(IImportRow row, int index, ImportContext context)
    {
        var cell = row.Cell(index);
        if (!cell.Exists) return null;

        // 取文本的这一步自行上报失败，放在其外，同一次失败才不会被上报两次
        var cellText = TextOf(cell);
        if (cellText == null || cellText.Length == 0) return null;

        switch (cell.Kind)
        {
            case CellType.Numeric:
                var date = cell.Date();
                if (date == null)
                    // 序列号超出 Excel 的日历，此时该单元格取 null
                    context.Report(row.SheetTitle, row.RowIndex, index,
                        new ArgumentException($"[{cell.NumberText()}] 不在 Excel 的日历范围内，无从读作日期。"));
                return date;
            case CellType.String:
                var text = cell.Text;
                var parsed = GetDateTimeFromString(text);
                if (parsed == null && !string.IsNullOrWhiteSpace(text))
                    // 文本不是日期是最常见的导入失败，同样要让调用方知道
                    context.Report(row.SheetTitle, row.RowIndex, index,
                        new FormatException($"[{text}] 不是可识别的日期。"));
                return parsed;
            default:
                return null;
        }
    }

    private static object? GetCellUri(IImportRow row, int key, ImportContext context)
    {
        var cellValue = TextOf(row.Cell(key));
        if (cellValue == null || cellValue.Length == 0) return null;

        try
        {
            return new Uri(cellValue);
        }
        catch (Exception e)
        {
            context.Report(row.SheetTitle, row.RowIndex, key, e);
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
    private static string? TextOf(CellData cell, bool datesAsText = false)
    {
        if (cell.Failed) return null;
        if (!cell.Exists) return string.Empty;

        var result = cell.Kind switch
        {
            CellType.Numeric => (datesAsText ? DateCellText(cell) : null) ?? cell.NumberText(),
            CellType.String => cell.Text,
            CellType.Blank => string.Empty,
            // Boolean ("TRUE"/"FALSE") and Error both read as the text they carry
            _ => cell.Text
        };

        return (result ?? "").Trim();
    }

    /// <summary>
    ///     A date cell is a number with a date format. Read as text - into a string property or a
    ///     dictionary - it comes out as the date it shows, not as the serial number Excel stores: an ISO
    ///     date, a time of day or both, whichever its format displays. Null when the cell is no date;
    ///     elapsed time (<c>[h]:mm</c>) counts as none, the number being the best that can be offered then.
    /// </summary>
    private static string? DateCellText(CellData cell)
    {
        var format = cell.Format;
        if (format == null) return null;

        if (!FormatParts.TryGetValue(format, out var parts))
        {
            parts = DateUtil.IsADateFormat(cell.FormatIndex, format)
                ? ExcelDateFormat.PartsShown(format)
                : ExcelDateFormat.Parts.None;
            FormatParts[format] = parts;
        }

        if (parts == ExcelDateFormat.Parts.None) return null;
        var date = cell.Date();
        if (date == null) return null;

        var pattern = parts == ExcelDateFormat.Parts.Date ? ExcelDateFormat.IsoDate :
            parts == ExcelDateFormat.Parts.Time ? ExcelDateFormat.IsoTime : ExcelDateFormat.IsoDateTime;
        return date.Value.ToString(pattern, CultureInfo.InvariantCulture);
    }

    private static SheetSource? GetDataRows(byte[]? bytes, ExcelImporterOptions options,
        ImportContext context)
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

        return new SheetSource(sheet.SheetName, NpoiRows(sheet, context));
    }

    /// <summary>取到停在表头那一行的枚举器：表头之上还可以有若干行说明文字。</summary>
    private static IEnumerator<IImportRow> AtHeader(SheetSource source, ExcelImporterOptions options)
    {
        var rows = source.Rows.GetEnumerator();
        rows.MoveNext();
        for (var i = 0; i < options.TitleSkipLine; i++) rows.MoveNext();
        return rows;
    }

    /// <summary>把整份读入内存的工作表的各行，包成与来源无关的行。</summary>
    private static IEnumerable<IImportRow> NpoiRows(ISheet sheet, ImportContext context)
    {
        // 1904 日期系统是整个工作簿的属性，逐格去问一遍并无意义
        var date1904 = sheet.Workbook.IsDate1904();
        foreach (IRow row in sheet)
            if (row != null)
                yield return new NpoiRow(row, context, date1904);
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

    private static object? GetEnum(IImportRow row, int key, Type enumType, ImportContext context)
    {
        var cellValue = TextOf(row.Cell(key));
        if (cellValue == null || cellValue.Length == 0) return null;
        if (Enum.GetNames(enumType).Contains(cellValue)) return Enum.Parse(enumType, cellValue);

        // 取值不在枚举中：沿用既有行为取 0，但不再悄无声息
        context.Report(row.SheetTitle, row.RowIndex, key,
            new FormatException($"[{cellValue}] 不是 {enumType.Name} 的取值，按 0 处理。"));
        return Enum.ToObject(enumType, 0);
    }
}
