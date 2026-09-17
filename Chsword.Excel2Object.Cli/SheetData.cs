using Chsword.Excel2Object.Options;

namespace Chsword.Excel2Object.Cli;

/// <summary>
///     A sheet as the importer sees it: the header titles in column order, and its rows on demand.
/// </summary>
/// <remarks>
///     行不再一次性读进内存：<see cref="Rows" /> 每次遍历都逐行读出，命令行因而能处理远大于内存的
///     文件。需要过两遍数据的地方（<c>--typed</c> 先推断类型再写出）就遍历两次，各自的开销是一遍
///     顺序读。
/// </remarks>
public sealed class SheetData
{
    private readonly string _path;
    private readonly string? _sheetTitle;
    private readonly bool _whole;

    private SheetData(string path, string? sheetTitle, bool whole, string title, List<string> columns)
    {
        _path = path;
        _sheetTitle = sheetTitle;
        _whole = whole;
        SheetTitle = title;
        Columns = columns;
    }

    public string SheetTitle { get; }

    public List<string> Columns { get; }

    /// <param name="whole">
    ///     整份读入工作簿，公式当场求值；否则逐行读出，公式取文件里存着的上一次计算结果。
    /// </param>
    public static SheetData Load(string path, string? sheetTitle, bool whole = false)
    {
        if (!File.Exists(path)) throw new FileNotFoundException($"input file not found: {path}", path);

        // 只读到表头那一行为止，后面有多少行数据都不影响这一步的开销
        using var input = File.OpenRead(path);
        var header = ExcelHelper.ReadHeader(input, options => options.SheetTitle = sheetTitle);
        return new SheetData(path, sheetTitle, whole, header.SheetTitle ?? "", header.Columns.ToList());
    }

    /// <summary>读出该表的各行。每次遍历都重新读一遍文件。</summary>
    public IEnumerable<Dictionary<string, object>> Rows()
    {
        if (_whole)
            return ExcelHelper.ExcelToObject<Dictionary<string, object>>(File.ReadAllBytes(_path), _sheetTitle);

        return Streamed();
    }

    private IEnumerable<Dictionary<string, object>> Streamed()
    {
        using var input = File.OpenRead(_path);
        foreach (var row in ExcelHelper.ExcelStreamToObject<Dictionary<string, object>>(input,
                     options => options.SheetTitle = _sheetTitle))
            yield return row;
    }

    /// <summary>
    ///     各列的类型推断，一遍读完，按列的先后存放——表头允许有重名的列，故不以标题为键。
    /// </summary>
    public TypeInference.Inference[] Infer()
    {
        var inferences = Columns.Select(_ => new TypeInference.Inference()).ToArray();
        foreach (var row in Rows())
            for (var i = 0; i < Columns.Count; i++)
                inferences[i].Observe(Text(row, Columns[i]));

        return inferences;
    }

    public static string Text(IReadOnlyDictionary<string, object> row, string column)
    {
        return row.TryGetValue(column, out var value) ? value?.ToString() ?? "" : "";
    }

    public static bool IsExcel(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xls", StringComparison.OrdinalIgnoreCase);
    }
}
