using System.Data;
using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Chsword.Excel2Object.Cli;

/// <summary>excel2obj convert: Excel -> JSON or JSON -> Excel, decided per input by file extension.</summary>
public static class ConvertCommand
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public static int Run(Arguments args, TextWriter output, TextWriter error)
    {
        if (args.Positional.Count == 0) throw new UsageException("convert needs at least one input file");
        var inputs = args.Positional;
        var target = args.Get("output");
        var sheet = args.Get("sheet");
        var typed = args.Has("typed");
        var xls = args.Has("xls");

        if (inputs.Count > 1)
        {
            if (target == null) throw new UsageException("--output <dir> is required when converting several files");
            Directory.CreateDirectory(target);
        }

        foreach (var input in inputs)
        {
            if (!File.Exists(input)) throw new FileNotFoundException($"input file not found: {input}", input);
            string? destination;
            if (inputs.Count > 1)
                destination = Path.Combine(target!, Path.GetFileNameWithoutExtension(input) +
                                                    (SheetData.IsExcel(input) ? ".json" : xls ? ".xls" : ".xlsx"));
            else if (target != null && Directory.Exists(target))
                destination = Path.Combine(target, Path.GetFileNameWithoutExtension(input) +
                                                   (SheetData.IsExcel(input) ? ".json" : xls ? ".xls" : ".xlsx"));
            else
                destination = target;

            if (SheetData.IsExcel(input))
            {
                var json = ExcelToJson(input, sheet, typed);
                if (destination == null)
                {
                    output.WriteLine(json);
                }
                else
                {
                    File.WriteAllText(destination, json);
                    error.WriteLine($"{input} -> {destination}");
                }
            }
            else
            {
                if (destination == null)
                    throw new UsageException("--output <file.xlsx> is required when converting JSON to Excel");
                var excelType = xls || destination.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
                    ? ExcelType.Xls
                    : ExcelType.Xlsx;
                File.WriteAllBytes(destination, JsonToExcel(input, sheet, excelType));
                error.WriteLine($"{input} -> {destination}");
            }
        }

        return Excel2ObjCli.Ok;
    }

    public static string ExcelToJson(string path, string? sheet, bool typed)
    {
        var data = SheetData.Load(path, sheet);
        var types = typed
            ? data.Columns.ToDictionary(c => c, c => TypeInference.Infer(data.ColumnValues(c)))
            : null;

        var array = new JsonArray();
        foreach (var row in data.Rows)
        {
            var item = new JsonObject();
            foreach (var column in data.Columns)
            {
                var text = row.TryGetValue(column, out var value) ? value?.ToString() ?? "" : "";
                item[column] = types == null ? JsonValue.Create(text) : ToJsonValue(text, types[column]);
            }

            array.Add(item);
        }

        return array.ToJsonString(JsonOptions);
    }

    private static JsonNode? ToJsonValue(string text, InferredType type)
    {
        if (type != InferredType.String && string.IsNullOrWhiteSpace(text)) return null;
        switch (type)
        {
            case InferredType.Bool:
                TypeInference.TryParseBool(text, out var flag);
                return JsonValue.Create(flag);
            case InferredType.Int:
                return JsonValue.Create(int.Parse(text, CultureInfo.InvariantCulture));
            case InferredType.Long:
                return JsonValue.Create(long.Parse(text, CultureInfo.InvariantCulture));
            case InferredType.Decimal:
                return JsonValue.Create(decimal.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture));
            case InferredType.DateTime:
                TypeInference.TryParseDateTime(text, out var date);
                return JsonValue.Create(date.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture));
            default:
                return JsonValue.Create(text);
        }
    }

    /// <summary>
    ///     A JSON array of flat objects becomes one sheet; the column set is the union of all keys in first-seen
    ///     order, and a column whose values are all numbers or all booleans gets typed cells.
    /// </summary>
    public static byte[] JsonToExcel(string path, string? sheetTitle, ExcelType excelType)
    {
        JsonNode? root;
        using (var stream = File.OpenRead(path))
        {
            root = JsonNode.Parse(stream);
        }

        if (root is not JsonArray array)
            throw new Excel2ObjectException($"{path}: expected a JSON array of objects at the top level");

        var columns = new List<string>();
        var kinds = new Dictionary<string, JsonValueKind?>();
        var objects = new List<JsonObject>();
        foreach (var node in array)
        {
            if (node is not JsonObject obj)
                throw new Excel2ObjectException($"{path}: every array element must be a JSON object");
            objects.Add(obj);
            foreach (var property in obj)
            {
                if (!kinds.ContainsKey(property.Key))
                {
                    columns.Add(property.Key);
                    kinds[property.Key] = null;
                }

                var kind = Kind(property.Value);
                if (kind == JsonValueKind.Null) continue;
                kinds[property.Key] = kinds[property.Key] == null || kinds[property.Key] == kind
                    ? kind
                    : JsonValueKind.String;
            }
        }

        var table = new DataTable(sheetTitle ?? Path.GetFileNameWithoutExtension(path));
        foreach (var column in columns)
            table.Columns.Add(column, kinds[column] switch
            {
                JsonValueKind.Number => typeof(double),
                JsonValueKind.True or JsonValueKind.False => typeof(bool),
                _ => typeof(string)
            });

        foreach (var obj in objects)
        {
            var row = table.NewRow();
            foreach (var column in columns)
            {
                var node = obj.TryGetPropertyValue(column, out var value) ? value : null;
                row[column] = node == null
                    ? DBNull.Value
                    : table.Columns[column]!.DataType == typeof(double) ? node.GetValue<double>()
                    : table.Columns[column]!.DataType == typeof(bool) ? node.GetValue<bool>()
                    : node is JsonValue ? node.ToString()
                    : node.ToJsonString();
            }

            table.Rows.Add(row);
        }

        return ExcelHelper.ObjectToExcelBytes(table, excelType, table.TableName)
               ?? throw new Excel2ObjectException($"{path}: nothing to write");
    }

    private static JsonValueKind Kind(JsonNode? node)
    {
        if (node == null) return JsonValueKind.Null;
        if (node is JsonValue value)
        {
            var element = value.GetValue<JsonElement>();
            return element.ValueKind == JsonValueKind.True ? JsonValueKind.False : element.ValueKind;
        }

        return JsonValueKind.String; // nested objects and arrays are written as their JSON text
    }
}
