# Date/Time Format Support

Excel2Object now supports a comprehensive set of date and time formats for Excel import operations. This document lists all supported formats and provides usage examples.

## Supported Date/Time Formats

### ISO 8601 Formats
- `yyyy-MM-ddTHH:mm:ss` - Standard ISO 8601 format (e.g., "2023-07-31T14:30:45")
- `yyyy-MM-ddTHH:mm:ssZ` - ISO 8601 with UTC timezone (e.g., "2023-07-31T14:30:45Z")
- `yyyy-MM-ddTHH:mm:ss.fff` - ISO 8601 with milliseconds (e.g., "2023-07-31T14:30:45.123")
- `yyyy-MM-ddTHH:mm:ss.fffZ` - ISO 8601 with milliseconds and UTC (e.g., "2023-07-31T14:30:45.123Z")
- `yyyy-MM-dd HH:mm:ss` - ISO 8601 with space separator (e.g., "2023-07-31 14:30:45")
- `yyyy-MM-dd HH:mm` - ISO 8601 without seconds (e.g., "2023-07-31 14:30")

### Date Formats with Various Separators

#### Dash Separator (-)
- `yyyy-MM-dd` - Year-first format (e.g., "2023-12-25")
- `dd-MM-yyyy` - Day-first format (European style, e.g., "25-12-2023")
- `MM-dd-yyyy` - Month-first format (US style, e.g., "12-25-2023")

#### Slash Separator (/)
- `yyyy/MM/dd` - Year-first format (e.g., "2023/12/25")
- `dd/MM/yyyy` - Day-first format (European style, e.g., "25/12/2023")
- `MM/dd/yyyy` - Month-first format (US style, e.g., "12/25/2023")

#### Dot Separator (.)
- `yyyy.MM.dd` - Year-first format (e.g., "2023.12.25")
- `dd.MM.yyyy` - Day-first format (European style, e.g., "25.12.2023")
- `MM.dd.yyyy` - Month-first format (US style, e.g., "12.25.2023")

### Date with Time (12-Hour Format with AM/PM)
- `yyyy-MM-dd hh:mm:ss tt` - Full date/time with AM/PM (e.g., "2023-07-31 02:30:45 PM")
- `yyyy-MM-dd hh:mm tt` - Date/time without seconds (e.g., "2023-07-31 02:30 PM")
- `dd-MM-yyyy hh:mm:ss tt` - European date with AM/PM (e.g., "31-07-2023 02:30:45 PM")
- `dd/MM/yyyy hh:mm:ss tt` - European date with slash and AM/PM (e.g., "31/07/2023 02:30:45 PM")
- `MM-dd-yyyy hh:mm:ss tt` - US date with AM/PM (e.g., "07-31-2023 02:30:45 PM")
- `MM/dd/yyyy hh:mm:ss tt` - US date with slash and AM/PM (e.g., "07/31/2023 02:30:45 PM")

### Date with Time (24-Hour Format)
- `dd-MM-yyyy HH:mm:ss` - European date with 24-hour time (e.g., "31-07-2023 14:30:45")
- `dd/MM/yyyy HH:mm:ss` - European date with slash and 24-hour time (e.g., "31/07/2023 14:30:45")
- `MM-dd-yyyy HH:mm:ss` - US date with 24-hour time (e.g., "07-31-2023 14:30:45")
- `MM/dd/yyyy HH:mm:ss` - US date with slash and 24-hour time (e.g., "07/31/2023 14:30:45")
- `dd-MM-yyyy HH:mm` - European date without seconds (e.g., "31-07-2023 14:30")
- `dd/MM/yyyy HH:mm` - European date with slash, no seconds (e.g., "31/07/2023 14:30")
- `MM-dd-yyyy HH:mm` - US date without seconds (e.g., "07-31-2023 14:30")
- `MM/dd/yyyy HH:mm` - US date with slash, no seconds (e.g., "07/31/2023 14:30")

### Short Date Formats (Single-Digit Month/Day)
- `yyyy/M/d` - Year-first with single digits (e.g., "2023/3/5")
- `yyyy-M-d` - Year-first with dash and single digits (e.g., "2023-3-5")
- `d/M/yyyy` - Day-first with single digits (e.g., "5/3/2023")
- `d-M-yyyy` - Day-first with dash and single digits (e.g., "5-3-2023")
- `M/d/yyyy` - Month-first with single digits (e.g., "3/5/2023")
- `M-d-yyyy` - Month-first with dash and single digits (e.g., "3-5-2023")

### Time-Only Formats (24-Hour)
When only time is provided, it will be combined with today's date.
- `HH:mm:ss` - 24-hour time with seconds (e.g., "14:30:45")
- `HH:mm` - 24-hour time without seconds (e.g., "14:30")
- `H:mm:ss` - Single-digit hour allowed (e.g., "9:30:45")
- `H:mm` - Single-digit hour without seconds (e.g., "9:30")

### Time-Only Formats (12-Hour with AM/PM)
When only time is provided, it will be combined with today's date.
- `hh:mm:ss tt` - 12-hour time with seconds and AM/PM (e.g., "02:30:45 PM")
- `hh:mm tt` - 12-hour time without seconds (e.g., "02:30 PM")
- `h:mm:ss tt` - Single-digit hour with AM/PM (e.g., "2:30:45 PM")
- `h:mm tt` - Single-digit hour without seconds (e.g., "2:30 PM")

### Chinese Date Formats (Legacy Support)
Excel2Object continues to support traditional Chinese date formats:
- `yyyy年MM月dd日` - Full Chinese date (e.g., "2023年07月31日")
- `yyyy年MM月` - Year and month only (e.g., "2023年07月")
- `yyyy年` - Year only (e.g., "2023年")

## Usage Examples

### Basic Import with DateTime

```csharp
using Chsword.Excel2Object;

public class Person
{
    [ExcelTitle("姓名")]
    public string Name { get; set; }
    
    [ExcelTitle("出生日期")]
    public DateTime? Birthday { get; set; }
    
    [ExcelTitle("创建时间")]
    public DateTime? CreateTime { get; set; }
}

var importer = new ExcelImporter();
var persons = importer.ExcelToObject<Person>("data.xlsx");
```

### Custom Date Format

```csharp
public class Employee
{
    [ExcelTitle("姓名")]
    public string Name { get; set; }
    
    [ExcelColumn("入职日期", Format = "yyyy-MM-dd HH:mm:ss")]
    public DateTime HireDate { get; set; }
    
    [ExcelColumn("离职日期", Format = "yyyy-MM-dd HH:mm:ss")]
    public DateTime? TerminationDate { get; set; }
}
```

### Supported Excel Cell Types

The library handles dates stored in Excel in two ways:

1. **Numeric Cell Type**: Excel stores dates as numeric values (days since 1900-01-01). The library automatically converts these using NPOI's `DateCellValue` property.

2. **String Cell Type**: When dates are stored as text in Excel, the library attempts to parse them using the comprehensive format list above.

When a date cell is read into a `string` property or a `Dictionary<string, object>`, it comes out as the date it shows rather than as the serial number Excel stores: `yyyy-MM-dd`, `HH:mm:ss` or `yyyy-MM-dd HH:mm:ss`, depending on whether the cell's format displays a date, a time or both. An elapsed-time format such as `[h]:mm` is a number, and reads as one. A numeric property (`double`, `decimal`, ...) always receives the number the cell stores, whatever its format.

## Exporting Dates

`DateTime` and `DateTime?` columns are written as **real date cells** - a serial number with a date number format - so Excel can sort, filter and calculate with them. A `null` becomes a blank cell.

The `Format` of `[ExcelColumn]` is a .NET date format string; it is translated into the Excel number format that displays the same thing. Without a `Format` the column shows `yyyy-mm-dd hh:mm:ss`.

| `[ExcelColumn(Format = ...)]` | Excel cell format               | Displayed as          |
|-------------------------------|---------------------------------|-----------------------|
| *(none)*                      | `yyyy-mm-dd hh:mm:ss`           | `2026-09-11 14:30:45` |
| `yyyy-MM-dd`                  | `yyyy-mm-dd`                    | `2026-09-11`          |
| `yyyy年MM月dd日`               | `yyyy"年"mm"月"dd"日"`           | `2026年09月11日`       |
| `MM/dd/yyyy hh:mm tt`         | `mm/dd/yyyy hh:mm AM/PM`        | `09/11/2026 02:30 PM` |
| `HH:mm:ss.fff`                | `hh:mm:ss.000`                  | `14:30:45.123`        |
| `d` (standard format)         | `m/d/yy` (Excel builtin 14)     | per Excel locale      |
| `m/d/yy`, `[$-409]d-mmm-yy`, `yyyy/m/d;@` | unchanged - Excel formats pass through | per Excel |

Time zone offsets (`zzz`, `K`) and eras (`g`) have no Excel counterpart and are dropped, and `hh` without `tt` shows the 24-hour clock - Excel has no 12-hour clock without an AM/PM marker. Dates before 1900-01-01 - `default(DateTime)` above all - are outside Excel's calendar (before 1904-01-01 in a workbook on the 1904 date system) and are written as text rendered with the `Format` instead.

`DateTime` values in a `Dictionary<string, object>` or a `DataTable` export become date cells the same way.

To keep the text export of versions before 2.4.0 (every date rendered with its `Format` and written as a string), set `ExcelExporterOptions.DateTimeAsText = true`. A `Format` already in Excel's spelling has no .NET rendering, so such a column is written as `yyyy-MM-dd HH:mm:ss` then.

## Regional Settings

The library attempts to parse dates using both `CultureInfo.InvariantCulture` and `CultureInfo.CurrentCulture` to handle regional differences. This means:

- US format dates (MM/dd/yyyy) will be correctly parsed
- European format dates (dd/MM/yyyy) will be correctly parsed
- ISO 8601 formats will always be parsed correctly regardless of regional settings

## Best Practices

1. **Use Nullable DateTime**: For optional date fields, use `DateTime?` to handle empty cells gracefully.

2. **Prefer Numeric Date Storage**: When possible, store dates in Excel as actual date values (numeric format) rather than text, as this is more reliable.

3. **Specify Format for Export**: When exporting to Excel with custom date formats, use the `Format` parameter in `ExcelColumnAttribute` (see [Exporting Dates](#exporting-dates)):
   ```csharp
   [ExcelColumn("日期", Format = "yyyy-MM-dd HH:mm:ss")]
   public DateTime MyDate { get; set; }
   ```

4. **Use ISO 8601 for Text Dates**: If dates must be stored as text in Excel, use ISO 8601 format (yyyy-MM-dd or yyyy-MM-ddTHH:mm:ss) for maximum compatibility.

5. **Avoid Ambiguous Formats**: Formats like "01/02/2023" can be interpreted as either January 2nd or February 1st depending on regional settings. Use year-first formats or be explicit about the ordering.

## Edge Cases Handled

- **Leap Years**: The library correctly handles February 29 in leap years
- **Time Zones**: ISO 8601 formats with 'Z' suffix are supported
- **Milliseconds**: Formats with milliseconds (.fff) are supported
- **Midnight/Noon**: Times at 00:00:00 and 12:00:00 are handled correctly
- **Time-Only Values**: When only time is provided (no date), it's combined with the current date
- **Empty Cells**: Empty cells are handled as null for nullable DateTime fields

## Performance Considerations

The library tries multiple parsing strategies in order:
1. Standard `DateTime.TryParse` (fastest)
2. Culture-specific exact format parsing with InvariantCulture
3. Culture-specific exact format parsing with CurrentCulture
4. Special handling for time-only formats
5. Fallback strategies for partial dates

This ensures maximum compatibility while maintaining reasonable performance.
