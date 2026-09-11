using System.Data;
using System.Globalization;
using System.Linq.Expressions;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Styles;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using HorizontalAlignment = Chsword.Excel2Object.Styles.HorizontalAlignment;

namespace Chsword.Excel2Object;

public class ExcelExporter
{
    public byte[]? AppendObjectToExcelBytes<TModel>(byte[] sourceExcelBytes, IEnumerable<TModel> data,
        string sheetTitle)
    {
        return ObjectToExcelBytes(data, options =>
        {
            options.SheetTitle = sheetTitle;
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
        var cellStyleDict = new Dictionary<string, ICellStyle>();
        
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
                    cell.CellStyle = cell.Sheet.Workbook.CreateCellStyle();
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(options.MappingColumnAction(columns[i].Title, columns[i].Type));
                    SetHeaderStyle(cell, columns[i].HeaderStyle);
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
                        SetCellValue(options, column, cell, raw, val, columnTitles, sheetColumnsResolver, cellStyleDict);
                    }
                }
            }

        return ToBytes(workbook);
    }    private static bool IsNumeric(Type type)
    {
        return type == typeof(int) || type == typeof(long) || type == typeof(double) || type == typeof(decimal) ||
               type == typeof(float) || type == typeof(short) || type == typeof(byte) || type == typeof(uint) ||
               type == typeof(ulong) || type == typeof(ushort) || type == typeof(sbyte);
    }

    private static void SetHeaderStyle(ICell cell, IExcelHeaderStyle? style)
    {
        if (style == null)
            return;
        var font = cell.Sheet.Workbook.CreateFont();
        cell.CellStyle.SetFont(font);
        if (!string.IsNullOrWhiteSpace(style.HeaderFontFamily))
            font.FontName = style.HeaderFontFamily;
        if (style.HeaderFontHeight > 0)
            font.FontHeightInPoints = style.HeaderFontHeight;
        else
            font.FontHeightInPoints = ExcelConstants.DefaultFontHeightInPoints;

        if (style.HeaderFontColor > 0)
            font.Color = (short) style.HeaderFontColor;
        if (style.HeaderBold)
            font.IsBold = true;
        if (style.HeaderItalic)
            font.IsItalic = true;
        if (style.HeaderStrikeout)
            font.IsStrikeout = true;
        if (style.HeaderUnderline)
            font.Underline = FontUnderlineType.Single; //暂不考虑等情况 Double
        if (style.HeaderAlignment != HorizontalAlignment.General)
            cell.CellStyle.Alignment = (NPOI.SS.UserModel.HorizontalAlignment) style.HeaderAlignment;
    }

    /// <summary>
    ///     Whether the column asked for anything that changes how a cell looks. [ExcelColumn] is handed
    ///     to the column as its CellStyle even when it only carries a title or header settings, so a
    ///     null check is not enough to tell "styled" from "not styled".
    /// </summary>
    private static bool DeclaresAppearance(IExcelCellStyle? style)
    {
        return style != null &&
               (!string.IsNullOrWhiteSpace(style.CellFontFamily) || style.CellFontHeight > 0 ||
                style.CellFontColor > 0 || style.CellBold || style.CellItalic || style.CellStrikeout ||
                style.CellUnderline || style.CellAlignment != HorizontalAlignment.General);
    }

    private static IFont? StyleToFont(IWorkbook workbook, IExcelCellStyle? style)
    {
        if (style == null || !DeclaresAppearance(style)) return null;
        var font = workbook.CreateFont();
        if (!string.IsNullOrWhiteSpace(style.CellFontFamily))
            font.FontName = style.CellFontFamily;
        // Leave the height alone unless asked for one: this font is now really applied to the cell, and
        // forcing a default here would shrink every [ExcelColumn] column that only sets a title.
        if (style.CellFontHeight > 0)
            font.FontHeightInPoints = style.CellFontHeight;

        if (style.CellFontColor > 0)
            font.Color = (short) style.CellFontColor;
        if (style.CellBold)
            font.IsBold = true;
        if (style.CellItalic)
            font.IsItalic = true;
        if (style.CellStrikeout)
            font.IsStrikeout = true;
        if (style.CellUnderline)
            font.Underline = FontUnderlineType.Single;

        return font;
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
    ///     Builds the cell style for a cell type, or returns one already built during this export. The
    ///     cache key fingerprints the requested style too, so columns asking for the same look share one
    ///     <see cref="ICellStyle" /> - a workbook can only hold a limited number of them.
    /// </summary>
    /// <remarks>
    ///     The cache is local to each export run, so styles are only reused within the workbook currently
    ///     being written.
    /// </remarks>
    private ICellStyle? CreateStyle(string type, ICell cell, IExcelCellStyle? style,
        IDictionary<string, ICellStyle> cellStyleDict)
    {
        if (type != ExcelConstants.CellTypes.Text && type != ExcelConstants.CellTypes.DateTime &&
            type != ExcelConstants.CellTypes.Appearance)
            return null;

        // Text and dates need their format either way; an Appearance cell has none of its own, so a
        // column that asked for no look has nothing to apply and keeps the workbook default. Decided
        // before the cache key is built, because this is the common case and the key costs an
        // allocation.
        if (type == ExcelConstants.CellTypes.Appearance && !DeclaresAppearance(style))
            return null;

        var key = GetKey(type, style);
        if (cellStyleDict.TryGetValue(key, out var cached)) return cached;

        var format = type == ExcelConstants.CellTypes.Text ? "@" :
            type == ExcelConstants.CellTypes.DateTime ? ExcelDateFormat.ToExcel(style?.Format) : null;
        var workbook = cell.Sheet.Workbook;
        var cellStyle = workbook.CreateCellStyle();
        var font = StyleToFont(workbook, style);
        if (font != null)
            cellStyle.SetFont(font);
        if (format != null)
            // GetFormat hands back the builtin index when there is one and registers the format otherwise
            cellStyle.DataFormat = workbook.CreateDataFormat().GetFormat(format);
        if (style != null && style.CellAlignment != HorizontalAlignment.General)
            cellStyle.Alignment = (NPOI.SS.UserModel.HorizontalAlignment) style.CellAlignment;

        cellStyleDict[key] = cellStyle;
        return cellStyle;
    }

    private void ApplyStyle(ICell cell, string type, IExcelCellStyle? style,
        IDictionary<string, ICellStyle> cellStyleDict)
    {
        var cellStyle = CreateStyle(type, cell, style, cellStyleDict);
        if (cellStyle != null)
            cell.CellStyle = cellStyle;
    }

    private string GetKey(string type, IExcelCellStyle? style)
    {
        if (style == null) return type;
        var arr = new List<string?>
        {
            type, style.CellFontFamily, style.CellAlignment.ToString(),
            style.CellBold.ToString(), style.CellFontColor.ToString(),
            style.CellFontHeight.ToString(CultureInfo.InvariantCulture),
            style.CellItalic.ToString(),
            style.CellStrikeout.ToString(),
            style.CellUnderline.ToString()
        };
        if (type == ExcelConstants.CellTypes.DateTime)
            arr.Add(style.Format);
        return string.Join("|", arr);
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

    private void SetCellValue(ExcelExporterOptions options, ExcelColumn column, ICell cell, object? raw, string val,
        string[] columnTitles, Func<string, string[]?> sheetColumnsResolver,
        IDictionary<string, ICellStyle> cellStyleDict)
    {
        var valueType = column.Type == null ? null : TypeUtil.GetUnNullableType(column.Type);
        if (valueType != null && valueType != typeof(Expression) && valueType != typeof(string) && val.Length == 0)
        {
            // null / missing values stay blank so formulas treat them as 0 instead of failing on ""
            cell.SetBlank();
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
                ApplyStyle(cell, ExcelConstants.CellTypes.Appearance, column.CellStyle, cellStyleDict);
                return;
            }
        }
        else if (valueType == typeof(bool) && bool.TryParse(val, out var flag))
        {
            cell.SetCellValue(flag);
            ApplyStyle(cell, ExcelConstants.CellTypes.Appearance, column.CellStyle, cellStyleDict);
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
            // A hyperlink cell holds text, so it takes the text style - but only when the column asked
            // for one. Unlike a plain string column there is nothing to protect here (no leading zeros
            // to keep), so styling every link would only change the format of existing exports.
            if (DeclaresAppearance(column.CellStyle))
                ApplyStyle(cell, ExcelConstants.CellTypes.Text, column.CellStyle, cellStyleDict);
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
            if (column.ResultType == typeof(DateTime))
                cell.CellStyle = CreateStyle(ExcelConstants.CellTypes.DateTime, cell, column.CellStyle,
                    cellStyleDict);
            else
                ApplyStyle(cell, ExcelConstants.CellTypes.Appearance, column.CellStyle, cellStyleDict);

            return;
        }
        else if (raw is DateTime date)
        {
            // decided on the value, not the column type: a dictionary export types every column string
            SetDateTimeCellValue(column, cell, date, val, options.DateTimeAsText, cellStyleDict);
            return;
        }
        else if (column.Type == typeof(string))
        {
            cell.SetCellType(CellType.String);
            cell.CellStyle = CreateStyle(ExcelConstants.CellTypes.Text, cell, column.CellStyle, cellStyleDict);
        }

        cell.SetCellValue(val);
    }

    /// <summary>
    ///     Dates become real date cells - a number with a date format - so Excel can sort, filter and
    ///     calculate with them. The column's Format, a .NET format string, is translated into the Excel
    ///     format that shows the same thing; <see cref="ExcelExporterOptions.DateTimeAsText" /> restores
    ///     the text export of earlier versions.
    /// </summary>
    private void SetDateTimeCellValue(ExcelColumn column, ICell cell, DateTime date, string val, bool asText,
        IDictionary<string, ICellStyle> cellStyleDict)
    {
        var format = column.CellStyle?.Format;

        // Excel's calendar starts at 1900-01-01 (1904 in a workbook on the 1904 date system), so anything
        // earlier - default(DateTime) above all - can only be kept as text
        var workbook = cell.Sheet.Workbook;
        if (!asText && DateUtil.IsValidExcelDate(DateUtil.GetExcelDate(date, workbook.IsDate1904())))
        {
            cell.SetCellValue(date);
            cell.CellStyle = CreateStyle(ExcelConstants.CellTypes.DateTime, cell, column.CellStyle, cellStyleDict);
            return;
        }

        cell.SetCellType(CellType.String);
        if (format != null)
        {
            cell.SetCellValue(DateToText(date, format));
            // the Format went into the text, but the column's font and alignment still apply
            ApplyStyle(cell, ExcelConstants.CellTypes.Appearance, column.CellStyle, cellStyleDict);
        }
        else
        {
            cell.SetCellValue(val);
            cell.CellStyle = CreateStyle(ExcelConstants.CellTypes.Text, cell, column.CellStyle, cellStyleDict);
        }
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
                    // a date cell shows its column's format, which is what the width has to fit
                    var cellText = value is DateTime date
                        ? DateToText(date, column.CellStyle?.Format)
                        : (value ?? "").ToString() ?? "";
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