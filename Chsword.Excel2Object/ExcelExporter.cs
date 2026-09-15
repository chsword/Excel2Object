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
using NPOI.XSSF.Streaming;
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

    /// <summary>
    ///     逐行取数据并写出，内存中只保留 <see cref="ExcelExporterOptions.StreamingRowWindow" /> 指定的
    ///     若干行，适用于行数多到不宜先在内存中建好整个工作簿的导出。
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         省下的是内存，不是等待时间：写过的行随即离开内存，落到临时文件，而最终的 <c>.xlsx</c>
    ///         包要等数据取完才一次写入 <paramref name="output" />。因此调用方的流在导出过程中并不会
    ///         陆续收到内容。
    ///     </para>
    ///     <para>
    ///         仅 <see cref="ExcelType.Xlsx" /> 能够流式写入；<see cref="ExcelType.Xls" /> 的格式决定了
    ///         必须先在内存中建好再写出，此时本方法只是把结果写进 <paramref name="output" />，并不省内存
    ///         （该格式至多 65,536 行，本也不适合大文件）。
    ///     </para>
    ///     <para>
    ///         流式写入依赖 NPOI 的 SXSSF，后者在把行刷出内存时要测量默认字符宽度，为此需要 SkiaSharp，
    ///         而 NPOI 把该依赖标为不随包传递。缺少时本方法会在写入开始前抛出
    ///         <see cref="Excel2ObjectException" /> 并说明需要引用哪些包。
    ///     </para>
    ///     <para><paramref name="output" /> 由调用方负责关闭，本方法不会关闭它。</para>
    /// </remarks>
    /// <example>
    ///     <code>
    /// using var file = File.Create("orders.xlsx");
    /// new ExcelExporter().ObjectToExcelStream(ReadOrders(), file, options =>
    /// {
    ///     options.ExcelType = ExcelType.Xlsx;
    ///     options.FreezeHeader = true;
    /// });
    ///     </code>
    /// </example>
    public void ObjectToExcelStream<TModel>(IEnumerable<TModel> data, Stream output,
        Action<ExcelExporterOptions>? optionsAction = null)
    {
        var options = new ExcelExporterOptions();
        optionsAction?.Invoke(options);
        var excel = data is IEnumerable<Dictionary<string, object>> models
            ? TypeConvert.ConvertDictionaryToExcelModel(models, options)
            : TypeConvert.ConvertObjectToExcelModel(data, options);

        ObjectToExcelStream(excel, options, output);
    }

    /// <inheritdoc cref="ObjectToExcelStream{TModel}(IEnumerable{TModel}, Stream, Action{ExcelExporterOptions})" />
    public void ObjectToExcelStream(DataTable dt, Stream output,
        Action<ExcelExporterOptions>? optionsAction = null)
    {
        var options = new ExcelExporterOptions();
        optionsAction?.Invoke(options);
        ObjectToExcelStream(TypeConvert.ConvertDataSetToExcelModel(dt, options), options, output);
    }

    private void ObjectToExcelStream(ExcelModel excel, ExcelExporterOptions options, Stream output)
    {
        // 字节数组那两个入口在源工作簿读不出来时返回 null，此处没有 null 可返回，故如实报错
        if (!Write(excel, options, output, true))
            throw new Excel2ObjectException("SourceExcelBytes 不是能够打开的工作簿。");
    }

    private byte[]? ObjectToExcelBytes(ExcelModel excel, ExcelExporterOptions options)
    {
        using var output = new MemoryStream();
        return Write(excel, options, output, false) ? output.ToArray() : null;
    }

    /// <summary>写出整个工作簿；源工作簿读不出来时返回 false。</summary>
    private bool Write(ExcelModel excel, ExcelExporterOptions options, Stream output, bool streaming)
    {
        IWorkbook? workbook = null;
        try
        {
            // 建立工作簿这一步本身也会失败（选项不合法、源工作簿读不出来、缺少 SkiaSharp），
            // 故一并放在 try 之内：字典入口为取列名已先读了一行，那个枚举器无论如何都要释放
            var window = streaming ? Window(options) : 0;
            workbook = OpenWorkbook(options, streaming, window);
            if (workbook == null) return false;

            var mapColumn = options.MappingColumnAction ?? ((title, _) => title);
            var styleFactory = new CellStyleFactory(workbook);

            if (excel.Sheets != null)
                foreach (var excelSheet in excel.Sheets)
                    WriteSheet(workbook, excelSheet, options, styleFactory, mapColumn);

            // 调用方给的流由调用方关闭
            workbook.Write(output, true);
            return true;
        }
        finally
        {
            // SXSSF 把刷出内存的行写在临时文件里，不释放则留在磁盘上
            (workbook as SXSSFWorkbook)?.Dispose();
            // 数据尚未取完即告失败时，取数据的枚举器也要释放
            excel.RowSource?.Dispose();
        }
    }

    /// <summary>
    ///     建立或读入工作簿。流式导出用 SXSSF：它只在内存中保留最近的 <paramref name="window" /> 行，
    ///     其余写入临时文件。
    /// </summary>
    private static IWorkbook? OpenWorkbook(ExcelExporterOptions options, bool streaming, int window)
    {
        IWorkbook workbook;
        if (options.SourceExcelBytes == null)
        {
            workbook = streaming && options.ExcelType == ExcelType.Xlsx
                ? new SXSSFWorkbook(window)
                : Workbook(options.ExcelType);
        }
        else
        {
            try
            {
                using var memoryStream = new MemoryStream(options.SourceExcelBytes);
                workbook = WorkbookFactory.Create(memoryStream);
            }
            catch
            {
                return null;
            }

            // .xls 无从流式写入，续写时就按原样在内存中完成
            if (streaming && workbook is XSSFWorkbook xssf) workbook = new SXSSFWorkbook(xssf, window);
        }

        if (workbook is SXSSFWorkbook streamed) StreamingSupport.Ensure(streamed);
        return workbook;
    }

    /// <summary>
    ///     内存中保留的行数。无论哪种格式都要校验：`.xls` 那条路上虽用不着这个值，取值不合法却同样
    ///     说明调用方想要的与得到的并不一致。
    /// </summary>
    private static int Window(ExcelExporterOptions options)
    {
        if (options.StreamingRowWindow > 0) return options.StreamingRowWindow;

        throw new Excel2ObjectException(
            $"StreamingRowWindow 为 {options.StreamingRowWindow}：内存中保留的行数须为正数。");
    }

    private static void WriteSheet(IWorkbook workbook, SheetModel excelSheet, ExcelExporterOptions options,
        CellStyleFactory styleFactory, Func<string, Type, string> mapColumn)
    {
        var sheet = string.IsNullOrWhiteSpace(excelSheet.Title)
            ? workbook.CreateSheet()
            : workbook.CreateSheet(excelSheet.Title);
        sheet.ForceFormulaRecalculation = true;
        var columns = excelSheet.Columns.OrderBy(c => c.Order).ToArray();

        // 要合并哪些列在此处即解析：写入尚未开始，列名写错也就不会留下写了一半的工作表
        var merges = new MergedRegions(columns, options);
        var layout = new SheetLayout(columns, options);

        WriteHeader(sheet, layout, options, styleFactory, mapColumn);
        var lastRowIndex = WriteRows(sheet, excelSheet.Rows, layout, options, styleFactory,
            BuildSheetColumnsResolver(workbook, sheet, layout.Titles), merges);
        ApplyColumnWidths(sheet, layout, options);

        ApplyHeaderView(sheet, options, columns.Length, lastRowIndex);
        DropdownValidation.Apply(sheet, columns, lastRowIndex);
        ConditionalFormatting.Apply(sheet, columns, lastRowIndex, options.ConditionalFormats);
        merges.Apply(sheet, lastRowIndex);
    }

    private static void WriteHeader(ISheet sheet, SheetLayout layout, ExcelExporterOptions options,
        CellStyleFactory styleFactory, Func<string, Type, string> mapColumn)
    {
        var row = sheet.CreateRow(ExcelConstants.DefaultHeaderRowIndex);
        for (var i = 0; i < layout.Columns.Length; i++)
        {
            var title = mapColumn(layout.Columns[i].Title, layout.Columns[i].Type);
            var cell = row.CreateCell(i);
            cell.SetCellType(CellType.String);
            cell.SetCellValue(title);
            var style = styleFactory.Get(StyleResolver.Header(layout.Columns[i], options.Styles));
            if (style != null) cell.CellStyle = style;
            layout.Measure(i, title);
        }
    }

    /// <summary>逐行写入，返回最后一行的行号；行只经过一遍，写过便可离开内存。</summary>
    private static int WriteRows(ISheet sheet, IEnumerable<Dictionary<string, object>> rows, SheetLayout layout,
        ExcelExporterOptions options, CellStyleFactory styleFactory, Func<string, string[]?> sheetColumnsResolver,
        MergedRegions merges)
    {
        var tracking = merges.TracksAnyColumn;
        var rowNumber = ExcelConstants.DefaultDataStartRowIndex;
        foreach (var item in rows)
        {
            var row = sheet.CreateRow(rowNumber++);
            var rowStyles = (row.RowNum - ExcelConstants.DefaultDataStartRowIndex) % 2 == 0
                ? layout.OddStyles
                : layout.EvenStyles;
            for (var i = 0; i < layout.Columns.Length; i++)
            {
                var column = layout.Columns[i];
                var cell = row.CreateCell(i);
                var raw = item.TryGetValue(column.Title, out var value) ? value : null;
                if (raw is DBNull) raw = null;
                var val = raw?.ToString() ?? "";
                SetCellValue(options, column, rowStyles[i], cell, raw, val, layout.Titles,
                    sheetColumnsResolver, styleFactory);

                // 一格量一次宽：单元格显示成什么样由所在列的样式决定，隔行底色于此无关，故一律按
                // 奇数行那一套来量，与逐列解析一次的做法一致
                layout.Measure(i, CellText(raw, layout.OddStyles[i]));

                if (tracking && merges.Tracks(i)) merges.Observe(i, row.RowNum, cell);
            }
        }

        return rowNumber - 1;
    }

    private static void ApplyColumnWidths(ISheet sheet, SheetLayout layout, ExcelExporterOptions options)
    {
        for (var i = 0; i < layout.Columns.Length; i++)
        {
            var width = layout.Widths == null
                ? options.DefaultColumnWidth
                : Math.Max(options.MinColumnWidth, Math.Min(options.MaxColumnWidth, layout.Widths[i]));
            sheet.SetColumnWidth(i, width * ExcelConstants.DefaultColumnWidthMultiplier);
        }
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
                var header = HeaderRow(workbook, sheetTitle, other);
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

    /// <summary>
    ///     另一张表的表头行。流式导出时 SXSSF 只是在工作簿外面套了一层：随源工作簿带进来的表，其行
    ///     并不在这一层里，须向底下的 XSSF 去取。
    /// </summary>
    private static IRow? HeaderRow(IWorkbook workbook, string sheetTitle, ISheet sheet)
    {
        var header = sheet.GetRow(ExcelConstants.DefaultHeaderRowIndex);
        if (header != null || workbook is not SXSSFWorkbook streamed) return header;

        return streamed.XssfWorkbook.GetSheet(sheetTitle)?.GetRow(ExcelConstants.DefaultHeaderRowIndex);
    }

    private static void SetCellValue(ExcelExporterOptions options, ExcelColumn column, ResolvedColumnStyle resolved,
        ICell cell, object? raw, string val, string[] columnTitles, Func<string, string[]?> sheetColumnsResolver,
        CellStyleFactory styleFactory)
    {
        var style = resolved.Style;
        var valueType = column.Type == null ? null : TypeUtil.GetUnNullableType(column.Type);
        if (valueType != null && valueType != typeof(Expression) && valueType != typeof(string) && val.Length == 0)
        {
            // null / missing values stay blank so formulas treat them as 0 instead of failing on "" - but
            // they still take the look of the rows around them, or a striped table would have holes in it
            cell.SetBlank();
            SetStyle(cell, styleFactory.Get(style, resolved.ValueFormat));
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
                SetStyle(cell, styleFactory.Get(style, resolved.ValueFormat));
                return;
            }
        }
        else if (valueType == typeof(bool) && bool.TryParse(val, out var flag))
        {
            cell.SetCellValue(flag);
            SetStyle(cell, styleFactory.Get(style, resolved.ValueFormat));
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
            SetStyle(cell, styleFactory.Get(style, resolved.ValueFormat));
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
                ? styleFactory.Get(style, ExcelDateFormat.ToExcel(resolved.DateFormat))
                : styleFactory.Get(style, resolved.ValueFormat));

            return;
        }
        else if (raw is DateTime date)
        {
            // decided on the value, not the column type: a dictionary export types every column string
            SetDateTimeCellValue(options, resolved, cell, date, val, styleFactory);
            return;
        }
        else if (column.Type == typeof(string))
        {
            cell.SetCellType(CellType.String);
            // text keeps what it holds verbatim - a leading zero, an identifier Excel would read as a number
            SetStyle(cell, styleFactory.Get(style, resolved.TextFormat ?? ExcelConstants.CellFormats.Text));
        }
        else
        {
            // an enum, a Guid, a TimeSpan: written as the text it renders to, and styled like any other cell
            SetStyle(cell, styleFactory.Get(style, resolved.ValueFormat));
        }

        cell.SetCellValue(val);
    }

    /// <summary>
    ///     Dates become real date cells - a number with a date format - so Excel can sort, filter and
    ///     calculate with them. The column's Format, a .NET format string, is translated into the Excel
    ///     format that shows the same thing; <see cref="ExcelExporterOptions.DateTimeAsText" /> restores
    ///     the text export of earlier versions.
    /// </summary>
    private static void SetDateTimeCellValue(ExcelExporterOptions options, ResolvedColumnStyle resolved, ICell cell,
        DateTime date, string val, CellStyleFactory styleFactory)
    {
        var asText = options.DateTimeAsText;
        var style = resolved.Style;
        var format = resolved.DateFormat;

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
            SetStyle(cell, styleFactory.Get(style, resolved.TextFormat ?? ExcelConstants.CellFormats.Text));
        }
    }

    private static void SetStyle(ICell cell, ICellStyle? style)
    {
        if (style != null) cell.CellStyle = style;
    }

    /// <summary>
    ///     The text a cell shows, which is what the width of an auto-sized column has to fit. A date
    ///     column given a number format - which only a format written for that very column can do - shows
    ///     the serial number Excel stores it as, so that is what it is measured as.
    /// </summary>
    private static string CellText(object? value, ResolvedColumnStyle resolved)
    {
        if (value is not DateTime date)
            return NumberToText(value, resolved.ValueFormat) ?? (value ?? "").ToString() ?? "";

        return NumberToText(DateUtil.GetExcelDate(date), resolved.DateFormat) ??
               DateToText(date, resolved.DateFormat);
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

    /// <summary>
    ///     写一张表时一路带着的东西：各列、各列在奇偶行上的样式，以及正在量的列宽。
    /// </summary>
    private sealed class SheetLayout
    {
        public SheetLayout(ExcelColumn[] columns, ExcelExporterOptions options)
        {
            Columns = columns;
            Titles = columns.Select(c => c.Title).ToArray();
            // one look per column per kind of row, rather than one worked out per cell
            OddStyles = columns
                .Select(c => StyleResolver.Resolve(c, options.Styles, ExcelConstants.DefaultDataStartRowIndex))
                .ToArray();
            EvenStyles = columns
                .Select(c => StyleResolver.Resolve(c, options.Styles, ExcelConstants.DefaultDataStartRowIndex + 1))
                .ToArray();
            // 自动列宽按内容而定，故边写边量，量完再设；数据只经过一遍，流式导出也就无需回看
            Widths = options.AutoColumnWidth ? new int[columns.Length] : null;
        }

        public ExcelColumn[] Columns { get; }

        public string[] Titles { get; }

        public ResolvedColumnStyle[] OddStyles { get; }

        public ResolvedColumnStyle[] EvenStyles { get; }

        /// <summary>各列已量到的宽度；未开启自动列宽时为 null。</summary>
        public int[]? Widths { get; }

        public void Measure(int column, string text)
        {
            if (Widths == null) return;

            var width = CalculateTextWidth(text);
            if (width > Widths[column]) Widths[column] = width;
        }
    }
}
