using System.Collections.Concurrent;
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
    private readonly ConcurrentDictionary<string, ICellStyle> _cellStyleDict = new();

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
        var options = new ExcelExporterOptions
        {
            ExcelType = excelType,
            SheetTitle = sheetTitle
        };
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
        
        // Clear style cache for each new workbook
        _cellStyleDict.Clear();
        
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
                        SetCellValue(excelType, column, cell, raw, val, columnTitles, sheetColumnsResolver);
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

    private static IFont? StyleToFont(IWorkbook workbook, IExcelCellStyle? style)
    {
        if (style == null) return null;
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
    ///     Builds (and caches per workbook) the cell style for a cell type. The cache key fingerprints the
    ///     requested style too, so columns asking for the same look share one <see cref="ICellStyle" /> -
    ///     a workbook can only hold a limited number of them.
    /// </summary>
    private ICellStyle? CreateStyle(string type, ICell cell, IExcelCellStyle? style)
    {
        var key = GetKey(type, style);
        if (_cellStyleDict.TryGetValue(key, out var cached)) return cached;

        string? format;
        if (type == ExcelConstants.CellTypes.Text)
            format = "text";
        else if (type == ExcelConstants.CellTypes.DateTime)
            format = style?.Format ?? "m/d/yy";
        else if (type == ExcelConstants.CellTypes.Number || type == ExcelConstants.CellTypes.Boolean)
        {
            // Numbers and booleans carry no format of their own, so a column that declared no style has
            // nothing to apply and keeps the workbook default rather than gaining an empty style.
            if (style == null) return null;
            format = type == ExcelConstants.CellTypes.Number && style.Format != null &&
                     HSSFDataFormat.GetBuiltinFormats().Contains(style.Format)
                ? style.Format
                : null;
        }
        else
            return null;

        var workbook = cell.Sheet.Workbook;
        var cellStyle = workbook.CreateCellStyle();
        var font = StyleToFont(workbook, style);
        if (font != null)
            cellStyle.SetFont(font);
        if (format != null)
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat(format);
        if (style != null && style.CellAlignment != HorizontalAlignment.General)
            cellStyle.Alignment = (NPOI.SS.UserModel.HorizontalAlignment) style.CellAlignment;

        _cellStyleDict.AddOrUpdate(key, cellStyle, (_, _) => cellStyle);
        return cellStyle;
    }

    private void ApplyStyle(ICell cell, string type, IExcelCellStyle? style)
    {
        var cellStyle = CreateStyle(type, cell, style);
        if (cellStyle != null)
            cell.CellStyle = cellStyle;
    }

    private string GetKey(string type, IExcelCellStyle? style)
    {
        if (style == null) return type;
        var arr = new[]
        {
            type, style.CellFontFamily, style.CellAlignment.ToString(),
            style.CellBold.ToString(), style.CellFontColor.ToString(),
            style.CellFontHeight.ToString(CultureInfo.InvariantCulture),
            style.CellItalic.ToString(),
            style.CellStrikeout.ToString(),
            style.CellUnderline.ToString(),
            style.Format
        };
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

    private void SetCellValue(ExcelType excelType, ExcelColumn column, ICell cell, object? raw, string val,
        string[] columnTitles, Func<string, string[]?> sheetColumnsResolver)
    {
        var valueType = column.Type == null ? null : Nullable.GetUnderlyingType(column.Type) ?? column.Type;
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
                ApplyStyle(cell, ExcelConstants.CellTypes.Number, column.CellStyle);
                return;
            }
        }
        else if (valueType == typeof(bool) && bool.TryParse(val, out var flag))
        {
            cell.SetCellValue(flag);
            ApplyStyle(cell, ExcelConstants.CellTypes.Boolean, column.CellStyle);
            return;
        }

        if (column.Type == typeof(Uri))
        {
            cell.Hyperlink = Switch<IHyperlink>(
                excelType,
                () => new HSSFHyperlink(HyperlinkType.Url)
                {
                    Address = val
                },
                () => new XSSFHyperlink(HyperlinkType.Url)
                {
                    Address = val
                }
            );
        }
        else if (column.Type == typeof(Expression))
        {
            var convert = new ExpressionConvert(columnTitles, cell.RowIndex, sheetColumnsResolver);

            if (column.CellStyle?.Format != null &&
                !HSSFDataFormat.GetBuiltinFormats().Contains(column.CellStyle.Format))
            {
                if (DateTime.TryParse(val, out var dt))
                {
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(dt.ToString(column.CellStyle.Format));
                }
                else
                {
                    cell.SetCellValue(val);
                }

                return;
            }

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
            if (column.ResultType != null)
                if (column.ResultType == typeof(DateTime))
                    cell.CellStyle = CreateStyle(ExcelConstants.CellTypes.DateTime, cell, column.CellStyle);

            return;
        }
        else if (column.Type == typeof(string))
        {
            cell.SetCellType(CellType.String);
            cell.CellStyle = CreateStyle(ExcelConstants.CellTypes.Text, cell, column.CellStyle);
        }
        else if (column.Type == typeof(DateTime) || column.Type == typeof(DateTime?))
        {
            if (column.CellStyle?.Format != null &&
                !HSSFDataFormat.GetBuiltinFormats().Contains(column.CellStyle.Format))
            {
                if (DateTime.TryParse(val, out var dt))
                {
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(dt.ToString(column.CellStyle.Format));
                }
                else
                {
                    cell.SetCellValue(val);
                }

                return;
            }

            cell.SetCellType(CellType.String);
            cell.CellStyle = CreateStyle(ExcelConstants.CellTypes.Text, cell, column.CellStyle);
        }

        cell.SetCellValue(val);
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
                    var cellText = (value ?? "").ToString() ?? "";
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