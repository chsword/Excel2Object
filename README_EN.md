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
- [x] Frozen header row and filter dropdowns ✅ **New in v2.5.0**
- [x] Data validation (dropdown lists) ✅ **New in v2.5.0**
- [x] Conditional formatting ✅ **New in v2.6.0**
- [x] A CSS-flavoured stylesheet: backgrounds, borders, striped rows, header styles ✅ **New in v2.7.0**
- [x] Support date/datetime/time formats in Excel ✅ **New in v2.0.4**, exported as real date cells ✅ **New in v2.4.0** - See [DateTimeFormats.md](DateTimeFormats.md)
- [x] Formula columns referencing other sheets of the same workbook ✅ **New in v2.1.0** - See [ExcelFunctions.md](ExcelFunctions.md)
- [x] Built-in formula function library ✅ **New in v2.3.0** - 334 Excel functions in 10 categories - See [ExcelFunctions.md](ExcelFunctions.md)

### Release Notes

* **2026.09.13** - v2.8.1
- [x] 🐛 Fixed styles overwriting one another on .NET Framework: `string.Join(string, params object[])` returns an empty string there when the first element is `null`, and both the cell-style and font cache keys start with the font colour, so every style that set none shared one cell style. Since v2.7.0 this left striped rows, borders, bold, underline and the `[ExcelColumn]` cell styles without effect on .NET Framework; on .NET (Core) they have always been correct
- [x] 🐛 Fixed numbers losing precision on import under .NET Framework, where the default numeric format is `G15` and `"R"` does not round-trip for some values, so a date serial read into a `double` property differed from the value stored. The text is now verified to parse back to the same number, falling back to 17 significant digits only where it does not
- [x] 🔧 The tests now run on several target frameworks - `net10.0` and `net8.0` everywhere, `net472` on Windows - which is how both defects above were found

* **2026.09.13** - v2.8.0
- [x] 🐛 Fixed how a failed cell read is handled on import: `ExcelImporter` used to catch the exception in three places, write it to standard output and carry on, leaving the caller with an empty value and no explanation. Nothing is written to standard output any more; failures are reported through the `ExcelImporterOptions.OnCellError` callback, which receives the sheet name, the row and column indexes, the cell reference and the original exception. A cell that cannot be read (a date serial outside Excel's calendar, say) takes its default value and the import continues, as before, and throwing from the callback stops it; a value that cannot be converted to the property type (the text `abc` into an `int`, say) is reported and then still throws, aborting the import as it always has. Text that is no date and a value outside its enum used to pass silently; both are now reported
- [x] ✨ **IMPROVED:** The formula evaluator is created once per workbook and reused throughout the import, instead of once per formula cell

* **2026.09.13** - v2.7.1
- [x] 🐛 Fixed which cells a stylesheet `Format` reaches: one written for a column now outranks the attribute's own (a format written for every cell used to win instead, dropping the time of day of `[ExcelColumn(Format = "yyyy-MM-dd HH:mm:ss")]`), and one written for every cell or for the row stripes only reaches the columns it can speak for - a date format no longer shows `12.5` as `1900-01-12`, and a number format no longer takes the `@` that keeps a leading zero away from a text column
- [x] ✨ **IMPROVED:** A border width can be named `thin` / `medium` / `thick` as CSS names it, and a border that cannot be read says which part of it was not understood
- [x] ✨ **IMPROVED:** The look of a column is resolved once per kind of row rather than once per cell, taking millions of style merges and key builds out of a large export

* **2026.09.13** - v2.7.0
- [x] ✨ **NEW:** A stylesheet, `options.Styles`: styles written by what they apply to rather than repeated on every `[ExcelColumn]` - `Header` / `Cells` / `Column("title")` / `OddRows` / `EvenRows`, layering `Cells → row stripes → [ExcelColumn] → Column`, with only the properties a style sets taking part
- [x] ✨ **NEW:** Cell background and borders, which had no support at all: `Background("#4472C4")`, `Border("1px solid #D0D0D0")`, along with `Wrap`, `VerticalAlign` and `FontSize`
- [x] ✨ **NEW:** Colours can be written as hex (`#RRGGBB` / `#RGB`) instead of only the 56-colour enum. `.xlsx` stores them as they are; `.xls` picks the nearest palette colour by perceived distance (CIELAB), so a light grey stays grey; a colour picked from `ExcelStyleColor` is still written as its palette index in both formats
- [x] ✨ **IMPROVED:** A style written for a column applies to any column type, so `Format("#,##0.00")` reaches a number column (the attribute's `Format` still applies to date columns only, as it always has)
- [x] ✨ **IMPROVED:** Cells that look alike share one cell style, so striped rows no longer push against the 4000-style limit of `.xls`

* **2026.09.12** - v2.6.0
- [x] ✨ **NEW:** Conditional formatting through `options.ConditionalFormats`: a rule per column (`Operator` + `Value`, `Value2` for `Between`, or a `Formula` of your own) restyles the cells it matches - font colour, bold, italic, background fill - and `WholeRow = true` colours the whole row. Excel evaluates the rules, so the colours follow the data as it is edited; a column can carry several, and both `.xls` and `.xlsx` support them

* **2026.09.12** - v2.5.0
- [x] ✨ **NEW:** `ExcelExporterOptions.FreezeHeader` freezes the header row so it stays in view while scrolling, and `ExcelExporterOptions.AutoFilter` puts Excel's filter dropdowns on it over the rows this export wrote. Both are off by default and work in `.xls` as well as `.xlsx`
- [x] ✨ **NEW:** `AppendObjectToExcelBytes` gained an options overload, so an appended sheet can be frozen, filtered and given dropdowns too
- [x] ✨ **NEW:** Data validation (dropdown lists): a fixed list goes on the attribute, `[ExcelColumn("Status", Dropdown = new[] {"Open", "Closed"})]`, and values only known at runtime go through `options.Dropdowns["Column title"] = values`, which wins over the attribute. Excel rejects anything outside the list; a list over 255 characters, or one whose values carry a comma or a quote, is written to a hidden sheet the dropdown reads from (such a list used to throw outright on `.xls`), and an export with no rows still gets the dropdown on its first row so it works as a template

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

### Frozen Header and Filter Dropdowns

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.ExcelType = ExcelType.Xlsx;
    options.FreezeHeader = true;           // the header stays in view while scrolling
    options.AutoFilter = true;             // the header carries Excel's filter dropdowns
});
```

Both are off by default and work in `.xls` as well as `.xlsx`. The filter covers the header and the rows written in this export; a sheet appended with `AppendObjectToExcelBytes` gets its own, leaving the sheets already in the workbook alone.

### Data Validation (Dropdown Lists)

``` csharp
public class OrderModel
{
    [ExcelTitle("No")] public string No { get; set; }

    // a fixed list belongs on the attribute
    [ExcelColumn("Status", Dropdown = new[] {"Open", "Closed", "Pending"})]
    public string Status { get; set; }

    [ExcelTitle("City")] public string City { get; set; }
}

var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    // values only known at runtime go through the options and win over the attribute
    options.Dropdowns["City"] = cities.Select(c => c.Name).ToArray();
});
```

Excel shows the values as a dropdown and rejects anything else. `ExcelHelper.AppendObjectToExcelBytes(bytes, models, options => ...)` takes the same options when appending a sheet. A list Excel cannot hold inline - over 255 characters in total, or a value carrying a comma or a quote - is written to a hidden sheet that a defined name points at, which the dropdown then reads (a `.xls` validation cannot reference another sheet directly); nothing about how it is used changes, and both `.xls` and `.xlsx` support it. An export with no rows still gets the dropdown on its first row, so it works as a template to fill in.

### Stylesheet (CSS-flavoured)

Styles need not be repeated on every `[ExcelColumn]`; write them by what they apply to:

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    options.Styles
        .Header(s => s.Bold().Background("#4472C4").Color("#FFF").Center())
        .Cells(s => s.Border("1px solid #D0D0D0"))
        .EvenRows(s => s.Background("#F2F2F2"))              // striped rows
        .Column("Amount", s => s.Format("#,##0.00").Right())
        .Column("Note", s => s.Wrap().VerticalAlign(ExcelVerticalAlignment.Middle));
});
```

What a style can set: `Color` / `Background` (`#RRGGBB`, `#RGB` or an `ExcelStyleColor`), `Bold` / `Italic` / `Underline` / `Strikeout`, `FontFamily` / `FontSize`, `Left` / `Center` / `Right` / `Align` / `VerticalAlign`, `Wrap`, `Format`, and `Border` with `BorderTop` / `BorderRight` / `BorderBottom` / `BorderLeft` (written as CSS writes them: `1px solid #D0D0D0`, `medium dashed red`, `none`; the width in pixels or as `thin` / `medium` / `thick`, colours as hex or one of the basic CSS names such as `red`, `gray`, `navy`).

**They layer** from the widest to the narrowest, and only the properties a style sets take part, so the layers add up rather than replace one another:

    Cells → OddRows / EvenRows → [ExcelColumn] attribute → Column("title")

`Header(...)` sits under the attribute's `Header…` properties the same way. Once a header style is written the whole header row shares one size, rather than keeping the 10pt every `[ExcelColumn]` header used to carry.

Where a `Format` is written decides which cells it reaches. Written for one column - `Column("title")` - it is a deliberate statement about that column and is used as it is, on numbers, text and dates alike, outranking the attribute's own `Format`. Written for every cell, or for the row stripes, it is a default that only reaches the columns it can speak for: a number format does not turn a date into `46278.00`, a date format does not turn `12.5` into `1900-01-12`, and neither takes the `@` that keeps a leading zero away from a text column.

**Colours**: `.xlsx` stores a hex colour as it is. `.xls` has only its 56-colour palette, so the nearest one is picked by what the eye sees (CIELAB) - a light grey comes out grey rather than lavender - though a very pale colour such as `#F2F2F2` lands on white, so striping an `.xls` wants a deeper grey such as `#C0C0C0`. A colour picked from `ExcelStyleColor` is still written as its palette index in both formats.

Cells that look alike share one cell style (`.xls` stops at 4000), so striping a long sheet costs a handful of styles rather than one per row.

### Conditional Formatting

``` csharp
var bytes = ExcelHelper.ObjectToExcelBytes(models, options =>
{
    // amounts over 10,000 turn red and bold
    options.ConditionalFormats.Add(new ConditionalFormat("Amount")
    {
        Operator = ConditionalOperator.GreaterThan,
        Value = 10000,
        FontColor = ExcelStyleColor.Red,
        Bold = true
    });

    // the whole row turns yellow when the status says so
    options.ConditionalFormats.Add(new ConditionalFormat("Status")
    {
        Operator = ConditionalOperator.Equal,
        Value = "Urgent",
        WholeRow = true,
        BackgroundColor = ExcelStyleColor.Yellow
    });

    // or write the Excel condition yourself, against the first data row (row 2), anchoring the column
    options.ConditionalFormats.Add(new ConditionalFormat("Amount")
    {
        Formula = "$B2>$C2",
        BackgroundColor = ExcelStyleColor.LightGreen
    });
});
```

Excel evaluates the rules itself, so the colours follow the data as it is edited. `Value` is written as the Excel literal for its type (a number as it is, a string quoted, a `DateTime` as `DATE(y,m,d)`), and `Between` / `NotBetween` take the other end in `Value2`. A column can carry several rules, and both `.xls` and `.xlsx` support them.

### Use with ASP.NET MVC

In ASP.NET MVC models, the `DisplayAttribute` can be supported like `ExcelTitleAttribute`.

## Document

- [docs/](docs/README.md) - Documentation index: **per-version feature notes** (in Chinese) and topic guides
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
