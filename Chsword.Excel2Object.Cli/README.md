# excel2obj

Command-line companion to [Excel2Object](https://github.com/chsword/Excel2Object): convert Excel workbooks to JSON
and back, and generate C# model classes from a sheet.

```bash
dotnet tool install -g Chsword.Excel2Object.Cli
```

## Excel -> JSON

```bash
excel2obj convert orders.xlsx                       # prints a JSON array to stdout
excel2obj convert orders.xlsx --output orders.json  # writes the file
excel2obj convert orders.xlsx --sheet Orders --typed
excel2obj convert a.xlsx b.xls --output ./json/     # several inputs need an output directory
excel2obj convert orders.xlsx --whole               # read the workbook whole and evaluate formulas
```

Each row becomes an object keyed by the header row. Values are the cell text; with `--typed`, a column whose
non-empty values are all integers, decimals, `TRUE`/`FALSE` or ISO dates is emitted as JSON numbers, booleans or
`yyyy-MM-ddTHH:mm:ss` strings, and empty cells become `null`.

The sheet is read row by row, so a file far larger than memory converts fine. One consequence: a **formula cell
reads the result stored in the file** rather than being evaluated. Excel writes that result when it saves, so
files saved by Excel read as before; a file written by this library carries formulas with no stored result, and
those cells read as blank. Pass `--whole` to read the workbook whole and evaluate formulas instead - the memory
it needs then grows with the file. `--whole` works with `generate-model` as well.

## JSON -> Excel

```bash
excel2obj convert orders.json --output orders.xlsx
excel2obj convert orders.json --output orders.xls --sheet Orders
```

The input must be a JSON array of flat objects. Columns are the union of all keys in first-seen order; a column
whose values are all numbers or all booleans is written as numeric / boolean cells, `null` as blank cells, nested
objects and arrays as their JSON text.

## Generate a model class

```bash
excel2obj generate-model orders.xlsx                                  # prints the class
excel2obj generate-model orders.xlsx --output Order.cs --class Order --namespace My.App
```

produces something like

```csharp
using System;
using Chsword.Excel2Object;

namespace My.App;

public class Order
{
    [ExcelTitle("Product")]
    public string? Product { get; set; }

    [ExcelTitle("Qty")]
    public int Qty { get; set; }

    [ExcelTitle("Unit Price")]
    public decimal? UnitPrice { get; set; }
}
```

Types are inferred from the data (`bool`, `int`, `long`, `decimal`, `DateTime`, otherwise `string`) and become
nullable when the column has empty cells. Titles are turned into PascalCase identifiers; any Unicode letters are kept.

## Exit codes

`0` success, `1` usage error (message on stderr followed by help), `2` the conversion itself failed.
