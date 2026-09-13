using System.Data;
using System.Globalization;
using System.Linq.Expressions;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Styles;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using HorizontalAlignment = Chsword.Excel2Object.Styles.HorizontalAlignment;

namespace Chsword.Excel2Object;

public class ExcelExporter
{
    public byte[]? AppendObjectToExcelBytes<TModel>(byte[] sourceExcelBytes, IEnumerable<TModel> data,
        string sheetTitle)
    {
        return AppendObjectToExcelBytes(sourceExcelBytes, data, options => options.SheetTitle = sheetTitle);
    }

    public byte[]? AppendObjectToExcelBytes<TModel>(byte[] sourceExcelBytes, IEnumerable<TModel> data,
        Action<ExcelExporterOptions> optionsAction)
    {
        return ObjectToExcelBytes(data, options =>
        {
            optionsAction(options);
            options.SourceExcelBytes = sourceExcelBytes;
        });
    }

    /// <summary>
    ///     Export a excel file from a List of T generic list
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <param name="data"></param>
    /// <param name="excelType"></param>
    /// <param name="sheetTitle"></param>
    /// <returns></returns>
    public byte[]? ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, ExcelType excelType = ExcelType.Xls,
        string? sheetTitle = null)
    {
        return ObjectToExcelBytes(data, options =>
        {
            options.ExcelType = excelType;
            options.SheetTitle = sheetTitle;
        });
    }

    public byte[]? ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, Action<ExcelExporterOptions> optionsAction)
    {
        var options = new ExcelExporterOptions();
        optionsAction(options);
        ExcelModel excel;
        if (data is IEnumerable<Dictionary<string, object>> models)
            excel = TypeConvert.ConvertDictionaryToExcelModel(models, options);
        else
            excel = TypeConvert.ConvertObjectToExcelModel(data, options);

        return ObjectToExcelBytes(excel, options);
    }

    public byte[]? ObjectToExcelBytes(DataTable dt, ExcelType excelType, string? sheetTitle = null)
    {
        return ObjectToExcelBytes(dt, options =>
        {
            options.ExcelType = excelType;
            options.SheetTitle = sheetTitle;
        });
    }

    public byte[]? ObjectToExcelBytes(DataTable dt, Action<ExcelExporterOptions> optionsAction)
    {
        var options = new ExcelExporterOptions();
        optionsAction(options);
        var excel = TypeConvert.ConvertDataSetToExcelModel(dt, options);
        return ObjectToExcelBytes(excel, options);
    }

    private byte[]? ObjectToExcelBytes(ExcelModel excel, ExcelExporterOptions options)
    {
        var excelType = options.ExcelType;

        IWorkbook workbook;
        if (options.SourceExcelBytes == null)
            workbook = Workbook(excelType);
        else
            // read work book
            try
            {
                using var memoryStream = new MemoryStream(options.SourceExcelBytes);
                workbook = WorkbookFactory.Create(memoryStream);
            }
            catch
            {
                return null;
            }

        if (options.MappingColumnAction == null) options.MappingColumnAction = (s, _) => s;
        var styleFactory = new CellStyleFactory(workbook);
        
        if (excel.Sheets != null)
            foreach (var excelSheet in excel.Sheets)
            {
                var sheet = string.IsNullOrWhiteSpace(excelSheet.Title)
                    ? workbook.CreateSheet()
                    : workbook.CreateSheet(excelSheet.Title);
                sheet.ForceFormulaRecalculation = true;
                var columns = excelSheet.Columns.OrderBy(c => c.Order).ToArray();

                // Calculate column widths
                if (options.AutoColumnWidth)
                {
                    var columnWidths = CalculateColumnWidths(columns, excelSheet.Rows, options, options.MappingColumnAction);
                    for (var i = 0; i < columns.Length; i++)
                    {
                        sheet.SetColumnWidth(i, columnWidths[i] * ExcelConstants.DefaultColumnWidthMultiplier);
                    }
                }
                else
                {
                    for (var i = 0; i < columns.Length; i++)
                    {
                        sheet.SetColumnWidth(i, options.DefaultColumnWidth * ExcelConstants.DefaultColumnWidthMultiplier);
                    }
                }
                
                var headerRow = sheet.CreateRow(ExcelConstants.DefaultHeaderRowIndex);
                for (var i = 0; i < columns.Length; i++)
                {
                    var cell = headerRow.CreateCell(i);
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(options.MappingColumnAction(columns[i].Title, columns[i].Type));
                    var headerStyle = styleFactory.Get(StyleResolver.Header(columns[i], options.Styles));
                    if (headerStyle != null) cell.CellStyle = headerStyle;
                }

                var columnTitles = columns.Select(c => c.Title).ToArray();
                var sheetColumnsResolver = BuildSheetColumnsResolver(workbook, sheet, columnTitles!);
                var rowNumber = ExcelConstants.DefaultDataStartRowIndex;
                var data = excelSheet.Rows;
                foreach (var item in data)
                {
                    var row = sheet.CreateRow(rowNumber++);
                    for (var i = 0; i < columns.Length; i++)
                    {
                        var column = columns[i];
                        var cell = row.CreateCell(i);
                        var raw = column.Title != null && item.TryGetValue(column.Title, out var value)
                            ? value
                            : null;
                        if (raw is DBNull) raw = null;
                        var val = raw?.ToString() ?? "";
                        var style = StyleResolver.Cell(column, options.Styles, row.RowNum);
                        SetCellValue(options, column, style, cell, raw, val, columnTitles, sheetColumnsResolver,
                            styleFactory);
                    }
                }

                ApplyHeaderView(sheet, options, columns.Length, rowNumber - 1);
                DropdownValidation.Apply(sheet, columns, rowNumber - 1);
                ConditionalFormatting.Apply(sheet, columns, rowNumber - 1, options.ConditionalFormats);
            }

        return ToBytes(workbook);
    }

    /// <summary>
    ///     Keeps the header usable on a long sheet: frozen in view while scrolling, and carrying Excel's
    ///     filter dropdowns.
    /// </summary>
    private static void ApplyHeaderView(ISheet sheet, ExcelExporterOptions options, int columnCount, int lastRowIndex)
    {
        if (options.FreezeHeader)
            sheet.CreateFreezePane(0, ExcelConstants.DefaultDataStartRowIndex);

        if (!options.AutoFilter || columnCount == 0) return;
        // the range covers the header even when no row followed it, so the dropdowns are there either way
        sheet.SetAutoFilter(new CellRangeAddress(ExcelConstants.DefaultHeaderRowIndex,
            Math.Max(lastRowIndex, ExcelConstants.DefaultHeaderRowIndex), 0, columnCount - 1));
    }

    private static bool IsNumeric(Type type)
    {
        return type == typeof(int) || type == typeof(long) || type == typeof(double) || type == typeof(decimal) ||
               type == typeof(float) || type == typeof(short) || type == typeof(byte) || type == typeof(uint) ||
               type == typeof(ulong) || type == typeof(ushort) || type == typeof(sbyte);
    }

    private static T Switch<T>(ExcelType excelType, Func<T> funcXlsHssf, Func<T> funcXlsxXssf)
    {
        var obj = excelType switch
        {
            ExcelType.Xls => funcXlsHssf(),
            ExcelType.Xlsx => funcXlsxXssf(),
            _ => throw new ArgumentOutOfRangeException(nameof(excelType))
        };

        return obj;
    }

    private static byte[] ToBytes(IWorkbook workbook)
    {
        using var output = new MemoryStream();
        workbook.Write(output, true);
        var bytes = output.ToArray();
        return bytes;
    }

    private static IWorkbook Workbook(ExcelType excelType)
    {
        IWorkbook workbook = excelType switch
        {
            ExcelType.Xls => new HSSFWorkbook(),
            ExcelType.Xlsx => new XSSFWorkbook(),
            _ => throw new ArgumentOutOfRangeException(nameof(excelType))
        };

        return workbook;
    }

    /// <summary>
    ///     Lets formulas on <paramref name="currentSheet" /> refer to other sheets by title. Sheets already in
    ///     the workbook (e.g. from <see cref="AppendObjectToExcelBytes{TModel}" />'s source bytes) have their
    ///     header row read as column titles; the current sheet resolves against its own titles.
    /// </summary>
    private static Func<string, string[]?> BuildSheetColumnsResolver(IWorkbook workbook, ISheet currentSheet,
        string[] currentColumnTitles)
    {
        var cache = new Dictionary<string, string[]?>(StringComparer.Ordinal);
        return sheetTitle =>
        {
            if (cache.TryGetValue(sheetTitle, out var cached)) return cached;

            string[]? titles = null;
            if (sheetTitle == currentSheet.SheetName)
            {
                titles = currentColumnTitles;
            }
            else
            {
                // NPOI happily writes a formula against a sheet that does not exist (Excel then shows
                // #REF!), so fail here instead, where the sheet title is still known.
                var other = workbook.GetSheet(sheetTitle) ?? throw new Excel2ObjectException(
                    $"refers to sheet [{sheetTitle}], which is not in the workbook. " +
                    "Write that sheet first, e.g. via AppendObjectToExcelBytes.");
                var header = other.GetRow(ExcelConstants.DefaultHeaderRowIndex);
                if (header != null)
                    // GetCell(i) rather than Cells: the latter skips blank cells and would shift indexes.
                    titles = Enumerable.Range(0, header.LastCellNum)
                        .Select(i => header.GetCell(i)?.ToString() ?? string.Empty)
                        .ToArray();
            }

            cache[sheetTitle] = titles;
            return titles;
        };
    }

    private void SetCellValue(ExcelExporterOptions options, ExcelColumn column, ExcelStyle? style, ICell cell,
        object? raw, string val, string[] columnTitles, Func<string, string[]?> sheetColumnsResolver,
        CellStyleFactory styleFactory)
    {
        var valueType = column.Type == null ? null : TypeUtil.GetUnNullableType(column.Type);
        if (valueType != null && valueType != typeof(Expression) && valueType != typeof(string) && val.Length == 0)
        {
            // null / missing values stay blank so formulas treat them as 0 instead of failing on "" - but
            // they still take the look of the rows around them, or a striped table would have holes in it
            cell.SetBlank();
            SetStyle(cell, styleFactory.Get(style));
            return;
        }

        if (valueType != null && IsNumeric(valueType))
        {
            // typed numbers become numeric cells (Excel would otherwise flag "number stored as text")
            var number = raw != null && IsNumeric(raw.GetType())
                ? Convert.ToDouble(raw, CultureInfo.InvariantCulture)
                : double.TryParse(val, NumberStyles.Any, CultureInfo.CurrentCulture, out var parsed)
                    ? parsed
                    : double.NaN;
            if (!double.IsNaN(number))
            {
                cell.SetCellValue(number);
                SetStyle(cell, styleFactory.Get(style));
                return;
            }
        }
        else if (valueType == typeof(bool) && bool.TryParse(val, out var flag))
        {
            cell.SetCellValue(flag);
            SetStyle(cell, styleFactory.Get(style));
            return;
        }

        if (column.Type == typeof(Uri))
        {
            cell.Hyperlink = Switch<IHyperlink>(
                options.ExcelType,
                () => new HSSFHyperlink(HyperlinkType.Url)
                {
                    Address = val
                },
                () => new XSSFHyperlink(HyperlinkType.Url)
                {
                    Address = val
                }
            );
            // A hyperlink cell holds text, so it takes the column's look - but nothing is forced on it.
            // Unlike a plain string column there is no leading zero to protect, so giving every link the
            // text format would only change the format of existing exports.
            SetStyle(cell, styleFactory.Get(style, style?.NumberFormat));
        }
        else if (column.Type == typeof(Expression))
        {
            var convert = new ExpressionConvert(columnTitles, cell.RowIndex, sheetColumnsResolver);

            string formula;
            try
            {
                formula = convert.Convert(column.Formula);
                cell.SetCellFormula(formula);
            }
            catch (Excel2ObjectException e)
            {
                throw new Excel2ObjectException($"Formula column [{column.Title}] {e.Message}", e);
            }
            catch (FormulaParseException e)
            {
                throw new Excel2ObjectException(
                    $"Formula column [{column.Title}] produced an invalid formula: {e.Message}", e);
            }
            // a formula that yields a date needs a date format, or Excel shows the serial number
            SetStyle(cell, column.ResultType == typeof(DateTime)
                ? styleFactory.Get(style, ExcelDateFormat.ToExcel(StyleResolver.DateFormat(style, column)))
                : styleFactory.Get(style));

            return;
        }
        else if (raw is DateTime date)
        {
            // decided on the value, not the column type: a dictionary export types every column string
            SetDateTimeCellValue(options, column, style, cell, date, val, styleFactory);
            return;
        }
        else if (column.Type == typeof(string))
        {
            cell.SetCellType(CellType.String);
            // text keeps what it holds verbatim - a leading zero, an identifier Excel would read as a number
            SetStyle(cell, styleFactory.Get(style, style?.NumberFormat ?? ExcelConstants.CellFormats.Text));
        }

        cell.SetCellValue(val);
    }

    /// <summary>
    ///     Dates become real date cells - a number with a date format - so Excel can sort, filter and
    ///     calculate with them. The column's Format, a .NET format string, is translated into the Excel
    ///     format that shows the same thing; <see cref="ExcelExporterOptions.DateTimeAsText" /> restores
    ///     the text export of earlier versions.
    /// </summary>
    private void SetDateTimeCellValue(ExcelExporterOptions options, ExcelColumn column, ExcelStyle? style,
        ICell cell, DateTime date, string val, CellStyleFactory styleFactory)
    {
        var asText = options.DateTimeAsText;
        var format = StyleResolver.DateFormat(style, column);

        // Excel's calendar starts at 1900-01-01 (1904 in a workbook on the 1904 date system), so anything
        // earlier - default(DateTime) above all - can only be kept as text
        var workbook = cell.Sheet.Workbook;
        if (!asText && DateUtil.IsValidExcelDate(DateUtil.GetExcelDate(date, workbook.IsDate1904())))
        {
            cell.SetCellValue(date);
            SetStyle(cell, styleFactory.Get(style, ExcelDateFormat.ToExcel(format)));
            return;
        }

        cell.SetCellType(CellType.String);
        if (format != null)
        {
            cell.SetCellValue(DateToText(date, format));
            // the Format went into the text, but the column's font and alignment still apply
            SetStyle(cell, styleFactory.Get(style));
        }
        else
        {
            cell.SetCellValue(val);
            SetStyle(cell, styleFactory.Get(style, ExcelConstants.CellFormats.Text));
        }
    }

    private static void SetStyle(ICell cell, ICellStyle? style)
    {
        if (style != null) cell.CellStyle = style;
    }

    /// <summary>
    /// Calculate optimal column widths based on content
    /// </summary>
    private int[] CalculateColumnWidths(ExcelColumn[] columns, IEnumerable<Dictionary<string, object>> data, 
        ExcelExporterOptions options, Func<string, Type, string> mappingColumnAction)
    {
        var columnWidths = new int[columns.Length];
        
        // Initialize with header widths
        for (var i = 0; i < columns.Length; i++)
        {
            var headerText = mappingColumnAction(columns[i].Title ?? "", columns[i].Type);
            columnWidths[i] = CalculateTextWidth(headerText);
        }

        // Calculate widths based on data content
        foreach (var item in data)
        {
            for (var i = 0; i < columns.Length; i++)
            {
                var column = columns[i];
                if (column.Title != null && item.TryGetValue(column.Title, out var value))
                {
                    // a cell shows its column's format, which is what the width has to fit
                    var format = StyleResolver.ColumnFormat(column, options.Styles);
                    var cellText = value is DateTime date
                        ? DateToText(date, format)
                        : NumberToText(value, format) ?? (value ?? "").ToString() ?? "";
                    var textWidth = CalculateTextWidth(cellText);
                    if (textWidth > columnWidths[i])
                    {
                        columnWidths[i] = textWidth;
                    }
                }
            }
        }

        // Apply min/max constraints
        for (var i = 0; i < columnWidths.Length; i++)
        {
            columnWidths[i] = Math.Max(options.MinColumnWidth, 
                Math.Min(options.MaxColumnWidth, columnWidths[i]));
        }

        return columnWidths;
    }

    /// <summary>
    ///     What a number shows under its format, for the width of an auto-sized column. Excel's number
    ///     formats and .NET's agree on the part people write - digits, a thousands separator, decimals, a
    ///     percent sign - so that much is rendered; anything else Excel alone understands (colours,
    ///     conditions, literals) is left to be measured as the plain number.
    /// </summary>
    private static string? NumberToText(object? value, string? format)
    {
        if (format == null || value == null || !IsNumeric(value.GetType())) return null;
        foreach (var c in format)
            if (!char.IsDigit(c) && c != '#' && c != '0' && c != '.' && c != ',' && c != '%' && c != ' ')
                return null;

        try
        {
            return Convert.ToDecimal(value, CultureInfo.InvariantCulture)
                .ToString(format, CultureInfo.InvariantCulture);
        }
        catch (Exception e) when (e is FormatException or OverflowException or InvalidCastException)
        {
            return null;
        }
    }

    /// <summary>
    ///     What a date shows under its Format, as .NET renders it: the text it is written as when it cannot
    ///     be a date cell, and what an auto-sized column is measured against. A Format already in Excel's
    ///     spelling has no .NET rendering (.NET would read its mm as minutes), so it falls back to the ISO
    ///     form the default format shows too.
    /// </summary>
    private static string DateToText(DateTime date, string? format)
    {
        if (format != null && !ExcelDateFormat.IsExcelSpelling(format))
            try
            {
                return date.ToString(format);
            }
            catch (FormatException)
            {
                // an unbalanced quote; Excel is more forgiving of those than .NET is
            }

        // an Excel-spelled format still says whether it shows a date, a time or both
        var parts = ExcelDateFormat.PartsShown(ExcelDateFormat.ToExcel(format));
        var pattern = parts == ExcelDateFormat.Parts.Date ? ExcelDateFormat.IsoDate :
            parts == ExcelDateFormat.Parts.Time ? ExcelDateFormat.IsoTime : ExcelDateFormat.IsoDateTime;
        return date.ToString(pattern, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Calculate text width in characters (approximation)
    /// </summary>
    private static int CalculateTextWidth(string text)
    {
        if (string.IsNullOrEmpty(text))
            return 1;

        // Basic character width calculation
        // This is a simple approximation - could be enhanced with font metrics
        double width = 0;
        foreach (var c in text)
        {
            if (char.IsControl(c))
                continue;
                
            // Wide characters (like Chinese) count as 2, others as 1
            if (c > 127)
                width += 2;
            else if (char.IsUpper(c) || "MWmw".Contains(c))
                width += 1.2; // Slightly wider for uppercase and wide letters
            else
                width += 1;
        }

        return (int)Math.Ceiling(width) + 2; // Add padding
    }
}