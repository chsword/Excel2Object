﻿using System.Collections.Concurrent;
using System.Data;
using System.Globalization;
using System.Linq.Expressions;
using Chsword.Excel2Object.Internal;
using Chsword.Excel2Object.Options;
using Chsword.Excel2Object.Styles;
using NPOI.HSSF.UserModel;
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

        CheckExcelModel(excel);
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
                var rowNumber = ExcelConstants.DefaultDataStartRowIndex;
                var data = excelSheet.Rows;
                foreach (var item in data)
                {
                    var row = sheet.CreateRow(rowNumber++);
                    for (var i = 0; i < columns.Length; i++)
                    {
                        var column = columns[i];
                        var cell = row.CreateCell(i);
                        var val = column.Title != null && item.TryGetValue(column.Title, out var value)
                            ? (value ?? "").ToString()
                            : "";
                        SetCellValue(excelType, column, cell, val!, columnTitles);
                    }
                }
            }

        return ToBytes(workbook);
    }    private static void SetHeaderStyle(ICell cell, IExcelHeaderStyle? style)
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
        //NPOI.SS.UserModel.FontColor.Red
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

    private static IFont? StyleToFont(ICell cell, IExcelCellStyle? style)
    {
        if (style == null) return null;
        var font = cell.Sheet.Workbook.CreateFont();
        if (!string.IsNullOrWhiteSpace(style.CellFontFamily))
            font.FontName = style.CellFontFamily;
        if (style.CellFontHeight > 0)
            font.FontHeightInPoints = style.CellFontHeight;
        else
            font.FontHeightInPoints = 10;

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
        if (style.CellAlignment != HorizontalAlignment.General)
            cell.CellStyle.Alignment = (NPOI.SS.UserModel.HorizontalAlignment) style.CellAlignment;

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

    // ReSharper disable once UnusedParameter.Local
    private void CheckExcelModel(ExcelModel excel)
    {
        //todo validate
    }

    private ICellStyle? CreateStyle(string type, ICell cell, IExcelCellStyle? style)
    {
        var key = GetKey(type, style);
        if (_cellStyleDict.TryGetValue(key, out var val)) return val;


        var font = StyleToFont(cell, style);

        if (key == "text")
        {
            var s1 = cell.Sheet.Workbook.CreateCellStyle();
            if (font != null)
                s1.SetFont(font);
            s1.DataFormat = HSSFDataFormat.GetBuiltinFormat("text");
            _cellStyleDict.AddOrUpdate(key, s1, (_, _) => s1);
            return s1;
        }

        if (key == "datatime")
        {
            var s1 = cell.Sheet.Workbook.CreateCellStyle();
            if (font != null) s1.SetFont(font);
            s1.DataFormat = HSSFDataFormat.GetBuiltinFormat(style?.Format ?? "m/d/yy");
            _cellStyleDict.AddOrUpdate(key, s1, (_, _) => s1);
            return s1;
        }

        return null;
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
            ((int) style.CellAlignment).ToString()
        };
        return string.Join("|", arr);
    }

    private void SetCellValue(ExcelType excelType, ExcelColumn column, ICell cell, string val,
        string[] columnTitles)
    {
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
            var convert = new ExpressionConvert(columnTitles, cell.RowIndex);

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

            var formula = convert.Convert(column.Formula);
            cell.SetCellFormula(formula);
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