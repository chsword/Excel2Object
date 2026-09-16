using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Chsword.Excel2Object.Cli;

/// <summary>excel2obj convert: Excel -> JSON or JSON -> Excel, decided per input by file extension.</summary>
public static class ConvertCommand
{
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
                if (destination == null)
                {
                    output.WriteLine(ExcelToJson(input, sheet, typed));
                }
                else
                {
                    using (var file = File.Create(destination)) WriteJson(input, sheet, typed, file);
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

    /// <summary>写到标准输出时才用得到：那里本就要把整段文本拿在手上。</summary>
    public static string ExcelToJson(string path, string? sheet, bool typed)
    {
        using var buffer = new MemoryStream();
        WriteJson(path, sheet, typed, buffer);
        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    /// <summary>
    ///     一行读出、一行写出，中途不把整份数据攒在内存里。<c>--typed</c> 要先知道每列是什么类型，
    ///     故先过一遍推断，再过一遍写出——两遍各是一次顺序读。
    /// </summary>
    public static void WriteJson(string path, string? sheet, bool typed, Stream destination)
    {
        var data = SheetData.Load(path, sheet);
        var types = typed ? data.Infer() : null;

        using var writer = new Utf8JsonWriter(destination,
            new JsonWriterOptions {Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping});
        writer.WriteStartArray();
        foreach (var row in data.Rows())
        {
            writer.WriteStartObject();
            foreach (var column in data.Columns)
            {
                writer.WritePropertyName(column);
                var text = SheetData.Text(row, column);
                if (types == null)
                    writer.WriteStringValue(text);
                else
                    WriteTypedValue(writer, text, types[column].Result);
            }

            writer.WriteEndObject();
        }

        writer.WriteEndArray();
    }

    private static void WriteTypedValue(Utf8JsonWriter writer, string text, InferredType type)
    {
        if (type != InferredType.String && string.IsNullOrWhiteSpace(text))
        {
            writer.WriteNullValue();
            return;
        }

        switch (type)
        {
            case InferredType.Bool:
                TypeInference.TryParseBool(text, out var flag);
                writer.WriteBooleanValue(flag);
                break;
            case InferredType.Int:
                writer.WriteNumberValue(int.Parse(text, CultureInfo.InvariantCulture));
                break;
            case InferredType.Long:
                writer.WriteNumberValue(long.Parse(text, CultureInfo.InvariantCulture));
                break;
            case InferredType.Decimal:
                writer.WriteNumberValue(decimal.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture));
                break;
            case InferredType.DateTime:
                TypeInference.TryParseDateTime(text, out var date);
                writer.WriteStringValue(date.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture));
                break;
            default:
                writer.WriteStringValue(text);
                break;
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
