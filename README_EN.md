# Excel2Object

[![install from nuget](http://img.shields.io/nuget/v/Chsword.Excel2Object.svg?style=flat-square)](https://www.nuget.org/packages/Chsword.Excel2Object)
[![release](https://img.shields.io/github/release/chsword/Excel2Object.svg?style=flat-square)](https://github.com/chsword/Excel2Object/releases)
[![.NET CI](https://github.com/chsword/Excel2Object/actions/workflows/dotnet-ci.yml/badge.svg)](https://github.com/chsword/Excel2Object/actions/workflows/dotnet-ci.yml)
[![CodeFactor](https://www.codefactor.io/repository/github/chsword/excel2object/badge)](https://www.codefactor.io/repository/github/chsword/excel2object)

Excel convert to .NET Object / .NET Object convert to Excel.

- [Top](#excel2object)
    - [NuGet install](#nuget-install)
    - [Release notes and roadmap](#release-notes-and-roadmap)
    - [Demo code](#demo-code)
    - [Document](#document)
    - [Development and Release](#development-and-release)
    - [Contributors](#contributors)
    - [Reference](#reference)

## Platform Support

[![.NET 4.7.2 +](https://img.shields.io/badge/-4.7.2%2B-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.0](https://img.shields.io/badge/-standard2.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET Standard 2.1](https://img.shields.io/badge/-standard2.1-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 6.0](https://img.shields.io/badge/-6.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)
[![.NET 8.0](https://img.shields.io/badge/-8.0-brightgreen?logo=dotnet&style=for-the-badge&color=blue)](#)

## NuGet Install

``` powershell
PM> Install-Package Chsword.Excel2Object
```

Or using .NET CLI:
``` bash
dotnet add package Chsword.Excel2Object
```

### Command-line tool excel2obj

``` bash
dotnet tool install -g Chsword.Excel2Object.Cli

excel2obj convert orders.xlsx --output orders.json --typed   # Excel -> JSON
excel2obj convert orders.json --output orders.xlsx           # JSON -> Excel
excel2obj generate-model orders.xlsx --class Order           # C# model class with [ExcelTitle] from the header row
```

See [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md).

## Release Notes and Roadmap

### Features Not Yet Supported

- [x] CLI tool ✅ **New in v2.2.1** - `dotnet tool install -g Chsword.Excel2Object.Cli`, see [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md)
- [x] Support auto width column ✅ **New in v2.0.4**
- [x] Support date/datetime/time formats in Excel ✅ **New in v2.0.4**, exported as real date cells ✅ **New in v2.4.0** - See [DateTimeFormats.md](DateTimeFormats.md)
- [x] Formula columns referencing other sheets of the same workbook ✅ **New in v2.1.0** - See [ExcelFunctions.md](ExcelFunctions.md)
- [x] Built-in formula function library ✅ **New in v2.3.0** - 334 Excel functions in 10 categories - See [ExcelFunctions.md](ExcelFunctions.md)

### Release Notes

* **2026.09.11** - v2.4.0
- [x] ✨ **NEW:** `DateTime` / `DateTime?` columns are exported as real date cells (a serial number with a date format), so Excel can sort, filter and calculate with them; `null` becomes a blank cell. `DateTime` values in `Dictionary<string, object>` and `DataTable` exports get the same treatment - See [DateTimeFormats.md](DateTimeFormats.md)
- [x] ✨ **NEW:** The .NET date format string of `[ExcelColumn(Format = ...)]` is translated into the Excel number format that shows the same thing (`yyyy-MM-dd HH:mm:ss` → `yyyy-mm-dd hh:mm:ss`, `yyyy年MM月dd日` → `yyyy"年"mm"月"dd"日"`); without one the column shows `yyyy-mm-dd hh:mm:ss`. A format already in Excel's spelling (`m/d/yy`, `[$-409]d-mmm-yy`, `yyyy/m/d;@`) is taken as is
- [x] ✨ **NEW:** `ExcelExporterOptions.DateTimeAsText` restores the text export of v2.3 and earlier; `DataTable` exports gained an options overload, `ObjectToExcelBytes(DataTable, Action<ExcelExporterOptions>)`
- [x] 🐛 Fixed a formula column taking over a `DateTime` column with a custom `Format` dropping the formula and writing the model's own value as text; the formula is kept and the date format applied (declare `FormulaResultType` when the formula yields a number rather than a date)
- [x] 🐛 Fixed date cells read into a `string` property or a `Dictionary<string, object>` coming out as the serial number Excel stores (`46276.6`): they are rendered as `yyyy-MM-dd`, `HH:mm:ss` or both, whichever the cell's format displays, while elapsed time such as `[h]:mm` stays a number. Numeric properties (`double`, `decimal`, ...) always receive the number the cell stores
- [x] ⚠️ **BEHAVIOR CHANGE:** Excel's calendar does not reach before 1900 (before 1904 in a workbook on the 1904 date system), so such dates - `default(DateTime)` above all - are still written as text rendered with the `Format`; and `hh` without `tt` shows the 24-hour clock, Excel having no 12-hour clock without an AM/PM marker

* **2026.09.11** - v2.3.0
- [x] ✨ **NEW:** The built-in formula function library grew from 34 to **334** Excel functions in 10 categories: math and trigonometry (73), statistical (80), logical (11), lookup and reference (33), date and time (25), text (39), information (20), financial (21), engineering (20) and database (12). Information, financial, engineering and database are new categories, reached through `ExcelFunctions.Information`, `.Financial`, `.Engineering` and `.Database` - See [ExcelFunctions.md](ExcelFunctions.md)
- [x] 🐛 Fixed functions Excel gained after 2007 showing `#NAME?` once the workbook was opened: xlsx stores them under a prefix (`_xlfn.IFS`, `_xlfn.STDEV.P`, and `_xlfn._xlws.` for `SORT` / `FILTER`), which is now added automatically; the already supported `DAYS` was affected
- [x] 🐛 Fixed numbers in formulas being formatted with the current culture, so that `1.5` became `1,5` under `de-DE` and the formula no longer parsed; the invariant culture is used now
- [x] 🐛 Fixed double quotes inside text literals not being escaped the way Excel wants them (`"` written as `""`)
- [x] ✨ **NEW:** More plain C# can be written straight into a formula: `%` (MOD), `^` (power), `&&` / `||` / `!` (AND / OR / NOT), the conditional operator (IF), `Math.*`, `string` members (`ToUpper`, `Substring`, `Length`, `Contains`, ...; members whose Excel counterpart computes something else, such as `Math.Round` and `string.Trim`, are left unsupported - call the `ExcelFunctions` function explicitly instead), `DateTime` members and `AddDays` / `AddMonths` / `AddYears`, and `.HasValue` on nullable properties
- [x] ✨ **NEW:** Formulas can use local variables the lambda captured and constant expressions such as `new DateTime(...)`; they are evaluated and written as literals instead of producing an invalid formula
- [x] ✨ **IMPROVED:** Ranges (`c.Matrix(...)`, `c.Columns(...)`) convert implicitly wherever a single value is expected, so `SUM` and friends take cells and ranges in one call
- [x] 🐛 Fixed cell styles from `[ExcelColumn]` being silently ignored: the style cache key never matched the branch that reads it, so `CellBold` / `CellFontColor` / `CellAlignment` and friends never reached the cell. Also fixed one column's `CellAlignment` being written to the workbook's default style, which dragged unrelated columns along with it (#47)
- [x] 🐛 Fixed formula columns with `FormulaResultType = typeof(DateTime)` getting no date format, so Excel showed a serial number such as `46282.6` (#47)
- [x] ⚠️ **BREAKING:** `IStatisticsFunction.Sum` takes `params ColumnValue[]` instead of `params ColumnMatrix[]`, and several `IMathFunction` members return `ColumnValue` instead of `int` / `double`. Source compatible - existing formulas still compile - but binary breaking, so recompile after upgrading

* **2026.09.11** - v2.2.1
- [x] 🔧 Fixed the command-line tool package exceeding NuGet's 250 MB limit, which kept 2.2.0 of the CLI from being published: net8.0 only and native `.pdb` files dropped (no library changes)

* **2026.09.11** - v2.2.0
- [x] ✨ **NEW:** Command-line tool `excel2obj` (`dotnet tool install -g Chsword.Excel2Object.Cli`): Excel ↔ JSON conversion, batch conversion, and model class generation with `[ExcelTitle]` from the header row - See [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md) (#9)
- [x] ✨ **NEW:** Formula columns can refer to columns through model properties: `options.FormulaColumns.Add<Order>("Total", (c, m) => m.Price * m.Qty)`, checked at compile time; referring to a property that is not exported throws - See [ExcelFunctions.md](ExcelFunctions.md) (#22)
- [x] ✨ **Improved:** Numeric columns are exported as numeric cells, bool columns as boolean cells and null/DBNull as blank cells (previously everything was text); string columns stay text so leading zeros survive
- [x] 🐛 Fixed missing parentheses from operator precedence in formulas, e.g. `c["A"] * (c["B"] + c["C"])` used to produce `A2*B2+C2`
- [x] ✨ **Improved:** Referring to a column title that does not exist on the current sheet now throws `Excel2ObjectException` instead of silently falling back

* **2026.09.08** - v2.1.0
- [x] ✨ **NEW:** Cross-sheet references in formula columns: `c.Sheet("Products")["Price", 2]`, `.Matrix(...)`, `.Columns(...)`, usable directly in `VLOOKUP` etc.; plus whole-column ranges on the current sheet via `c.Columns("A", "D")` - See [ExcelFunctions.md](ExcelFunctions.md)
- [x] ✨ **Improved:** Referencing an unknown column or sheet in a formula now throws `Excel2ObjectException` (naming the formula column) instead of emitting an invalid formula
- [x] ✨ **Updated:** NPOI to 2.8.0 (NPOI 2.8+ ships with the OSMF EULA, which this library accepts in its csproj; the source remains Apache-2.0)
- [x] 🔒 Fixed vulnerable transitive dependency System.Security.Cryptography.Xml (pinned to 8.0.4)
- [x] 🧹 Removed the unused SixLabors.ImageSharp dependency; implemented `FormulaColumnsCollection.CopyTo`; `ColumnValue.Equals/GetHashCode` no longer throw
- [x] 🔧 CI and tests moved to .NET 10; fixed the release workflow failing to create GitHub Releases

* **2025.10.11** - v2.0.5
- [x] 🔧 `release.ps1` auto-increments the patch version and supports `-Help` (no library changes)

* **2025.10.11** - v2.0.4
- [x] ℹ️ First 2.0.x version published to NuGet since 2.0.0.211 (v2.0.1–v2.0.3 were never published); includes all of the following:
- [x] ✨ **NEW:** Auto column width adjustment based on content
  - Automatically calculates optimal column widths
  - Supports minimum and maximum width constraints
  - Handles Chinese/Unicode characters properly
  - Configurable through `ExcelExporterOptions`
- [x] ✨ **NEW:** Comprehensive date/time format support (56 formats) - See [DateTimeFormats.md](DateTimeFormats.md)
  - ISO 8601 formats with/without timezone and milliseconds
  - Multiple date separators (dash, slash, dot)
  - 12-hour and 24-hour time formats
  - Time-only formats
  - Regional format support (US, European, etc.)
  - Backward compatible with existing Chinese date formats (年月日)
- [x] ✨ **Updated:** NPOI to 2.7.5
- [x] ✨ **Updated:** SixLabors.ImageSharp to 3.1.11 (Fixed security vulnerability)
- [x] 🔧 GitHub Actions CI replaces AppVeyor; tag-triggered release workflow and `release.ps1` one-command release script

* **2024.10.21**
- [x] Updated SixLabors.ImageSharp to 2.1.9
- [x] Tested for .NET 8.0

* **2024.05.10**
- [x] Support .NET 8.0 / .NET 6.0 / .NET Standard 2.1 / .NET Standard 2.0 / .NET Framework 4.7.2
- [x] Cleared deprecated libraries

* **2023.11.02**
- [x] Support column title mapping - [Issue39DynamicMappingTitle.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue39DynamicMappingTitle.cs)

* **2023.07.31**
- [x] Support DateTime and Nullable<DateTime> format, such as `[ExcelColumn("Title",Format="yyyy-MM-dd HH:mm:ss")]`

* **2023.03.26**
- [x] Support special symbols in column titles #37 - [Issue37SpecialCharTest.cs](https://github.com/chsword/Excel2Object/commit/273122275e724367bb6154e03df61702fcec81b3#diff-5f0f5f7558bf7d4207cfa752a4506c4df89d9b491e2501e4862aff0c2276bd61)

* **2023.02.20**
- [x] Support platforms: .NET Standard 2.0/2.1, .NET 6.0, .NET Framework 4.7.2

* **2022.03.19**
- [x] Support ExcelImporterOptions, Skipline - [Issue32SkipLineImport.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue32SkipLineImport.cs)
- [x] Fixed superclass property bug - [Issue31SuperClass.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Issue31SuperClass.cs)

* **2021.11.4**
- [x] Multiple sheet support - [Pr28MultipleSheetTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr28MultipleSheetTest.cs)

* **2021.10.23**
- [x] Fixed Nullable DateTime bug @SunBrook

* **2021.10.22**
- [x] Support Nullable types - [Pr24NullableTest.cs](https://github.com/chsword/Excel2Object/blob/main/Chsword.Excel2Object.Tests/Pr24NullableTest.cs) @SunBrook

* **2021.5.28**
- [x] Support styling for headers & cells, new [ExcelColumnAttribute] for columns
- [x] Support Functions - [ExcelFunctions.md](./ExcelFunctions.md)

```C#
var list = new List<Pr20Model>
{
    new Pr20Model
    {
        Fullname = "AAA", Mobile = "123456798123"
    },
    new Pr20Model
    {
        Fullname = "BBB", Mobile = "234"
    }
};
var bytes = ExcelHelper.ObjectToExcelBytes(list, ExcelType.Xlsx);

// Model definition
[ExcelTitle("SheetX")]
public class Pr20Model
{
    [ExcelColumn("Full name", CellFontColor = ExcelStyleColor.Red)]
    public string Fullname { get; set; }

    [ExcelColumn("Phone Number",
        HeaderFontFamily = "Normal",
        HeaderBold = true,
        HeaderFontHeight = 30,
        HeaderItalic = true,
        HeaderFontColor = ExcelStyleColor.Blue,
        HeaderUnderline = true,
        HeaderAlignment = HorizontalAlignment.Right,
        //cell
        CellAlignment = HorizontalAlignment.Justify
    )]
    public string Mobile { get; set; }
}
```

* **v2.0.0.113**
```
Converted project to .NET Standard 2.0 and .NET Framework 4.5.2
Fixed bugs #12 #13
```

* **v1.0.0.80**
- [x] Support simple formulas
- [x] Support standard Excel model
  - [x] Excel & JSON conversion
  - [x] Excel & Dictionary<string,object> conversion

```
Support Uri to hyperlink cell
Also support text cell to Uri type
```

* **v1.0.0.43**
```
Support xlsx format [thanks Soar360]
Support complex Boolean type
```

* **v1.0.0.36**
```
Add ExcelToObject<T>(bytes)
```

## Demo Code

### Define Model

``` csharp
public class ReportModel
{
    [Excel("My Title", Order=1)]
    public string Title { get; set; }
    
    [Excel("User Name", Order=2)]
    public string Name { get; set; }
}
```

### Create Model List

``` csharp
var models = new List<ReportModel>
{
    new ReportModel{Name="a", Title="b"},
    new ReportModel{Name="c", Title="d"},
    new ReportModel{Name="f", Title="e"}
};
```

### Convert Object to Excel File

``` csharp
var exporter = new ExcelExporter();
var bytes = exporter.ObjectToExcelBytes(models);
File.WriteAllBytes("C:\\demo.xls", bytes);
```

### Convert Excel File to Object

``` csharp
var importer = new ExcelImporter();
IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>("c:\\demo.xls");

// You can also use bytes directly
// IEnumerable<ReportModel> result = importer.ExcelToObject<ReportModel>(bytes);
```

### Auto Column Width (New Feature)

``` csharp
// Enable auto column width adjustment
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.ExcelType = ExcelType.Xlsx;
    options.AutoColumnWidth = true;        // Enable auto width
    options.MinColumnWidth = 8;            // Minimum width in characters
    options.MaxColumnWidth = 50;           // Maximum width in characters
    options.DefaultColumnWidth = 16;       // Default width when auto is disabled
});
```

### Use with ASP.NET MVC

In ASP.NET MVC models, the `DisplayAttribute` can be supported like `ExcelTitleAttribute`.

## Document

- [ExcelFunctions.md](ExcelFunctions.md) - Formula columns: built-in functions, references by column title or model property, cross-sheet references
- [DateTimeFormats.md](DateTimeFormats.md) - Supported date/time formats and parsing rules
- [Chsword.Excel2Object.Cli/README.md](Chsword.Excel2Object.Cli/README.md) - The `excel2obj` command-line tool
- [ROADMAP.md](ROADMAP.md) - Roadmap

For more information, please visit: http://www.cnblogs.com/chsword/p/excel2object.html

## Development and Release

- **[Automated Release Script `release.ps1`](release.ps1)** - PowerShell script for one-click version update, commit, and tag creation
- **[Automated Release System Documentation](RELEASE_AUTOMATION.md)** - Complete guide including Copilot instructions, version management, and automated release workflow
- **[Versioning Guidelines](.github/VERSIONING.md)** - Semantic versioning specification and version increment rules
- **[Release Process Guide](.github/RELEASE_GUIDE.md)** - Detailed release steps and troubleshooting guide
- **[Copilot Instructions](.github/copilot-instructions.md)** - GitHub Copilot development guidelines and project conventions

### Quick Release

Use the automation script for one-click release:

```powershell
# Windows: without arguments it auto-increments the patch version; pass -Version 2.3.0 to pick one
.\release.ps1

# Linux/macOS (PowerShell Core required)
pwsh ./release.ps1
```

See [RELEASE_AUTOMATION.md](RELEASE_AUTOMATION.md) for detailed instructions.

## Contributors

[![Contributors](https://contrib.rocks/image?repo=chsword/Excel2Object)](https://github.com/chsword/Excel2Object/graphs/contributors)

## Reference

- https://github.com/tonyqus/npoi
- https://github.com/chsword/ctrc

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
