using NPOI.SS.UserModel;

namespace Chsword.Excel2Object.Cli;

/// <summary>A sheet as the importer sees it: the header titles in column order and one string per cell.</summary>
public sealed class SheetData
{
    public SheetData(string sheetTitle, List<string> columns, List<Dictionary<string, object>> rows)
    {
        SheetTitle = sheetTitle;
        Columns = columns;
        Rows = rows;
    }

    public string SheetTitle { get; }
    public List<string> Columns { get; }
    public List<Dictionary<string, object>> Rows { get; }

    public static SheetData Load(string path, string? sheetTitle)
    {
        if (!File.Exists(path)) throw new FileNotFoundException($"input file not found: {path}", path);
        var bytes = File.ReadAllBytes(path);
        var rows = ExcelHelper.ExcelToObject<Dictionary<string, object>>(bytes, sheetTitle).ToList();
        var (title, columns) = ReadHeader(bytes, sheetTitle);
        return new SheetData(title, columns, rows);
    }

    /// <summary>The header row straight from the workbook, so an empty sheet still yields its columns.</summary>
    private static (string title, List<string> columns) ReadHeader(byte[] bytes, string? sheetTitle)
    {
        using var stream = new MemoryStream(bytes);
        var workbook = WorkbookFactory.Create(stream);
        var sheet = string.IsNullOrEmpty(sheetTitle) ? workbook.GetSheetAt(0) : workbook.GetSheet(sheetTitle);
        if (sheet == null) throw new Excel2ObjectException($"The specified sheet:[{sheetTitle}] does not exist");
        var header = sheet.GetRow(sheet.FirstRowNum);
        // untrimmed on purpose: the importer keys rows by the exact header text
        var columns = header?.Cells.Select(cell => cell.ToString() ?? "").ToList() ?? new List<string>();
        return (sheet.SheetName, columns);
    }

    public IEnumerable<string> ColumnValues(string column)
    {
        return Rows.Select(row => row.TryGetValue(column, out var value) ? value?.ToString() ?? "" : "");
    }

    public static bool IsExcel(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xls", StringComparison.OrdinalIgnoreCase);
    }
}
