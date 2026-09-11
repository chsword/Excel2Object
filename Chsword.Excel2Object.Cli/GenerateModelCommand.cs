using System.Globalization;
using System.Text;

namespace Chsword.Excel2Object.Cli;

/// <summary>excel2obj generate-model: a C# class with [ExcelTitle] attributes inferred from a sheet.</summary>
public static class GenerateModelCommand
{
    public static int Run(Arguments args, TextWriter output, TextWriter error)
    {
        if (args.Positional.Count != 1) throw new UsageException("generate-model needs exactly one input file");
        var input = args.Positional[0];
        var sheet = args.Get("sheet");
        var data = SheetData.Load(input, sheet);
        var className = args.Get("class") ?? ToIdentifier(data.SheetTitle, "Model");
        var code = Generate(data, className, args.Get("namespace"));

        var destination = args.Get("output");
        if (destination == null)
        {
            output.Write(code);
        }
        else
        {
            File.WriteAllText(destination, code, new UTF8Encoding(false));
            error.WriteLine($"{input} -> {destination}");
        }

        return Excel2ObjCli.Ok;
    }

    public static string Generate(SheetData data, string className, string? ns)
    {
        var sb = new StringBuilder();
        sb.AppendLine("using System;");
        sb.AppendLine("using Chsword.Excel2Object;");
        sb.AppendLine();
        if (!string.IsNullOrWhiteSpace(ns))
        {
            sb.Append("namespace ").Append(ns).AppendLine(";");
            sb.AppendLine();
        }

        sb.Append("public class ").AppendLine(className);
        sb.AppendLine("{");
        var used = new HashSet<string>(StringComparer.Ordinal) {className};
        for (var i = 0; i < data.Columns.Count; i++)
        {
            var title = data.Columns[i];
            var values = data.ColumnValues(title).ToList();
            var type = TypeInference.Infer(values);
            var nullable = values.Count == 0 || values.Any(string.IsNullOrWhiteSpace);
            var name = Unique(ToIdentifier(title, $"Column{i + 1}"), used);

            if (i > 0) sb.AppendLine();
            sb.Append("    [ExcelTitle(").Append(Quote(title)).AppendLine(")]");
            sb.Append("    public ").Append(TypeInference.ToCSharp(type, nullable)).Append(' ').Append(name)
                .AppendLine(" { get; set; }");
        }

        sb.AppendLine("}");
        // generated code uses \n on every OS so the output is byte-for-byte stable
        return sb.ToString().Replace("\r\n", "\n");
    }

    /// <summary>
    ///     Turns a column title into a PascalCase identifier: letters (any script) and digits survive, everything
    ///     else splits words; a leading digit gets an underscore; an empty result falls back to <paramref name="fallback" />.
    /// </summary>
    public static string ToIdentifier(string title, string fallback)
    {
        var sb = new StringBuilder();
        var upperNext = true;
        foreach (var ch in title)
        {
            if (char.IsLetterOrDigit(ch) || ch == '_')
            {
                sb.Append(upperNext && char.IsLetter(ch) ? char.ToUpper(ch, CultureInfo.InvariantCulture) : ch);
                upperNext = false;
            }
            else
            {
                upperNext = true;
            }
        }

        if (sb.Length == 0) return fallback;
        if (char.IsDigit(sb[0])) sb.Insert(0, '_');
        return sb.ToString(); // PascalCase never collides with a (lowercase) C# keyword
    }

    private static string Unique(string name, HashSet<string> used)
    {
        var candidate = name;
        for (var i = 2; !used.Add(candidate); i++) candidate = name + i;
        return candidate;
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }
}
