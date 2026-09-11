namespace Chsword.Excel2Object.Cli;

/// <summary>
///     Entry point shared by Program.cs and the tests: parses the command line and dispatches to a command.
///     Returns the process exit code (0 on success, 1 on a usage error, 2 when the command itself failed).
/// </summary>
public static class Excel2ObjCli
{
    public const int Ok = 0;
    public const int UsageError = 1;
    public const int CommandFailed = 2;

    public static int Run(string[] args, TextWriter output, TextWriter error)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help" or "help")
        {
            output.Write(Help);
            return Ok;
        }

        if (args[0] is "-v" or "--version")
        {
            output.WriteLine(typeof(Excel2ObjCli).Assembly.GetName().Version?.ToString(3));
            return Ok;
        }

        Arguments arguments;
        try
        {
            arguments = Arguments.Parse(args.Skip(1));
        }
        catch (UsageException e)
        {
            error.WriteLine($"error: {e.Message}");
            return UsageError;
        }

        try
        {
            switch (args[0])
            {
                case "convert":
                    return ConvertCommand.Run(arguments, output, error);
                case "generate-model":
                    return GenerateModelCommand.Run(arguments, output, error);
                default:
                    error.WriteLine($"error: unknown command '{args[0]}'");
                    error.Write(Help);
                    return UsageError;
            }
        }
        catch (UsageException e)
        {
            error.WriteLine($"error: {e.Message}");
            return UsageError;
        }
        catch (Exception e)
        {
            error.WriteLine($"error: {e.Message}");
            return CommandFailed;
        }
    }

    public const string Help = """
        excel2obj - Excel <-> JSON conversion and C# model generation (https://github.com/chsword/Excel2Object)

        Usage:
          excel2obj convert <input>... [--output <file|dir>] [--sheet <title>] [--typed] [--xls]
          excel2obj generate-model <input.xlsx> [--output <file.cs>] [--sheet <title>] [--class <Name>] [--namespace <Ns>]
          excel2obj --version | --help

        convert
          Excel (.xlsx/.xls) -> JSON array of objects keyed by the header row, or JSON -> Excel.
          The direction follows the file extensions. Without --output a single JSON result goes to stdout;
          with several inputs --output must be a directory.
          --sheet <title>   Sheet to read (default: first sheet). When writing Excel, the sheet title.
          --typed           Excel -> JSON: emit numbers, booleans and dates as JSON numbers/booleans/ISO strings
                            instead of the raw cell text.
          --xls             JSON -> Excel: write the legacy .xls format (default .xlsx).

        generate-model
          Reads the header row and sample data of a sheet and writes a C# class with [ExcelTitle] attributes,
          inferring int / decimal / bool / DateTime / string (nullable when the column has empty cells).
          --class <Name>    Class name (default: derived from the sheet title).
          --namespace <Ns>  Wrap the class in a file-scoped namespace.
          --output <file>   Destination (default: stdout).

        """;
}

/// <summary>Positional arguments plus --name value / --name=value options and --flag switches.</summary>
public sealed class Arguments
{
    private readonly Dictionary<string, string?> _options = new(StringComparer.Ordinal);

    public List<string> Positional { get; } = new();

    public static Arguments Parse(IEnumerable<string> args)
    {
        var result = new Arguments();
        var queue = new Queue<string>(args);
        while (queue.Count > 0)
        {
            var arg = queue.Dequeue();
            if (!arg.StartsWith("--", StringComparison.Ordinal))
            {
                result.Positional.Add(arg);
                continue;
            }

            var name = arg.Substring(2);
            string? value = null;
            var eq = name.IndexOf('=');
            if (eq >= 0)
            {
                value = name.Substring(eq + 1);
                name = name.Substring(0, eq);
            }
            else if (queue.Count > 0 && !queue.Peek().StartsWith("--", StringComparison.Ordinal) &&
                     !Switches.Contains(name))
            {
                value = queue.Dequeue();
            }

            if (name.Length == 0) throw new UsageException("empty option name");
            result._options[name] = value;
        }

        return result;
    }

    /// <summary>Options that never take a value, so a following positional argument is not swallowed.</summary>
    private static readonly HashSet<string> Switches = new(StringComparer.Ordinal) {"typed", "xls"};

    public bool Has(string name)
    {
        return _options.ContainsKey(name);
    }

    public string? Get(string name)
    {
        return _options.TryGetValue(name, out var value) ? value : null;
    }

    public string Require(string name)
    {
        return Get(name) ?? throw new UsageException($"--{name} requires a value");
    }
}

public sealed class UsageException : Exception
{
    public UsageException(string message) : base(message)
    {
    }
}
