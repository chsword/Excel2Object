# Excel Functions in Formula Columns

A formula column is a C# lambda that is never executed: it is read as an expression tree and translated
to Excel formula text. **334 Excel functions** are available, in ten categories, plus the C# operators and
a set of ordinary .NET members that map onto Excel functions.

### Use formula

``` csharp
var bytes = new ExcelExporter().ObjectToExcelBytes(list, options =>
            {
                options.ExcelType = ExcelType.Xlsx;
                options.FormulaColumns.Add(new FormulaColumn
                {
                    Title = "BirthYear",
                    Formula = c => (int) c["Age"] + DateTime.Now.Year,
                    AfterColumnTitle = "Column1"
                });
            });
            // c => (int) c["Age"] + DateTime.Now.Year will convert to like =A3+YEAR(NOW())
```

### Base

|Function|Syntax |Description|
|---|:---|:----|
|A4 | ```  c => c["One"] ```| Cell in current row|
|A2 | ```  c => c["One",2] ```| Specify any Cell|
|A2:B4 | ```  c => c.Matrix("One",2,"Two",4) ```|
|A:D | ```  c => c.Columns("One","Four") ```| Whole columns|

### Refer to columns by model property

When exporting a typed model, `FormulaColumns.Add<TModel>` passes the model as a second lambda parameter so
columns can be referenced by property instead of by title string. The compiler checks the property names,
and each property is mapped to its column through its `[ExcelTitle]` / `[Display]` attribute. Properties
without one of those attributes are not exported and throw `Excel2ObjectException` when referenced.

``` csharp
public class OrderLine
{
    [ExcelTitle("Product")] public string Product { get; set; }
    [ExcelTitle("Price")] public decimal Price { get; set; }
    [ExcelTitle("Qty")] public int Qty { get; set; }
}

var bytes = new ExcelExporter().ObjectToExcelBytes(lines, options =>
{
    options.FormulaColumns.Add<OrderLine>("Total", (c, m) => m.Price * m.Qty);          // =B2*C2
    options.FormulaColumns.Add<OrderLine>("Label", (c, m) => m.Product + "-" + m.Qty);  // =A2&"-"&C2
    options.FormulaColumns.Add<OrderLine>("Level",
        (c, m) => ExcelFunctions.Condition.If(m.Qty > 5, "bulk", "single"));            // =IF(C2>5,"bulk","single")
});
```

`m.Property` always means the cell of that column on the current row; use the first parameter for anything
else (`c["Price", 2]`, `c.Matrix(...)`, `c.Sheet(...)`), and mix the two freely. String `+` is written as `&`,
`DateTime` members (`m.Ordered.Year`) and `.Value` on nullable properties are supported, and parentheses
follow Excel operator precedence.

### Other sheets

`c.Sheet("title")` refers to another sheet of the same workbook. Its column titles are read from that sheet's
header row, so columns are named the same way as on the current sheet. The sheet must already be in the workbook
when the formula column is written, e.g. write it first and append the formula sheet with `AppendObjectToExcelBytes`
(or `options.SourceExcelBytes`); otherwise an `Excel2ObjectException` is thrown.

|Function|Syntax |Description|
|---|:---|:----|
|'Products'!B4 | ```  c => c.Sheet("Products")["Price"] ```| Cell in current row|
|'Products'!B2 | ```  c => c.Sheet("Products")["Price",2] ```| Specify any Cell|
|'Products'!A2:B20 | ```  c => c.Sheet("Products").Matrix("Name",2,"Price",20) ```|
|'Products'!A:B | ```  c => c.Sheet("Products").Columns("Name","Price") ```| Whole columns|
|VLOOKUP(A4,'Products'!A:B,2,FALSE) | ```  c => ExcelFunctions.Reference.VLookup(c["Product"], c.Sheet("Products").Columns("Name","Price"), 2, false) ```|

Sheet titles are always quoted, so titles with spaces or apostrophes work. If a title cannot be found on the other
sheet, a plain column letter (`"A"`, `"BC"`) is used as-is; any other unknown title throws `Excel2ObjectException`.

### Ranges and single values

Ranges are `ColumnMatrix` values - `c.Matrix("One", 2, "One", 9)`, `c.Columns("One", "Four")`,
`c.Sheet("Products").Matrix(...)` - and a range converts to a single value implicitly, so functions that
accept either take both:

``` csharp
c => ExcelFunctions.Statistics.Sum(c["Qty"], c.Matrix("Qty", 2, "Qty", 9))   // =SUM(C4,C2:C9)
c => ExcelFunctions.Math.SumIf(c.Matrix("Qty", 2, "Qty", 9), ">100",
         c.Matrix("Price", 2, "Price", 9))                                   // =SUMIF(C2:C9,">100",B2:B9)
```

Functions that take criteria pairs (`SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `IFS`, `SWITCH`)
take them as trailing arguments, in the order Excel expects them.

### C# written straight into a formula

Besides the function library, ordinary C# inside the lambda is translated. This works on cells
(`c["Qty"] > 5 && c["Price"] > 1`) as well as on model properties:

|C#|Formula|
|:---|:---|
|`a + b`, `a - b`, `a * b`, `a / b`|`a+b`, `a-b`, `a*b`, `a/b`|
|`"x" + a` (string)|`"x"&a`|
|`a % b`|`MOD(a,b)`|
|`a ^ b` (on a cell)|`a^b` (Excel's power operator, not a bitwise xor)|
|`a & b`, `a \| b`, `a ^ b` (on integers)|`_xlfn.BITAND(a,b)`, `_xlfn.BITOR(a,b)`, `_xlfn.BITXOR(a,b)`|
|`-a`|`-a`|
|`a == b`, `a != b`, `a > b`, `a >= b`, `a < b`, `a <= b`|`a=b`, `a<>b`, `a>b`, `a>=b`, `a<b`, `a<=b`|
|`a && b`, `a \|\| b`, `!a`, `a ^ b` (on bools)|`AND(a,b)`, `OR(a,b)`, `NOT(a)`, `_xlfn.XOR(a,b)`|
|`d.Ticks` below a second|added as its fraction of a day, since `TIME` takes whole seconds|
|`cond ? x : y`|`IF(cond,x,y)`|
|`Math.Abs/Sqrt/Pow/Ceiling/Floor/Truncate/Log/Max/Min/...`|`ABS`, `SQRT`, `POWER`, `_xlfn.CEILING.MATH`, `_xlfn.FLOOR.MATH`, ...|
|`s.ToUpper()`, `s.ToLower()`, `s.Length`|`UPPER(s)`, `LOWER(s)`, `LEN(s)`|
|`s.Substring(i)`, `s.Substring(i, n)`|`MID(s,i+1,LEN(s))`, `MID(s,i+1,n)`|
|`s.Replace(a, b)`, `s.Contains(x)`|`SUBSTITUTE(s,a,b)`, `ISNUMBER(FIND(x,s))`|
|`s.StartsWith(x)`, `s.EndsWith(x)`, `s.IndexOf(x)`|`EXACT(LEFT(s,LEN(x)),x)`, `EXACT(RIGHT(...),x)`, `IFERROR(FIND(x,s)-1,-1)`|
|`DateTime.Now`, `DateTime.Today`|`NOW()`, `TODAY()`|
|`d.Year/Month/Day/Hour/Minute/Second`|`YEAR(d)`, `MONTH(d)`, ...|
|`d.AddMonths(n)`, `d.AddYears(n)`, `d.AddDays(n)`|`EDATE(d,n)`, `EDATE(d,(n)*12)`, `(d+n)`|
|`p.Value`, `p.HasValue` (nullable)|the cell itself, `NOT(ISBLANK(cell))`|
|a captured variable, `new DateTime(2024, 3, 1)`|the value as a literal, `DATE(2024,3,1)`|

`FIND` and `EXACT` are used rather than `SEARCH` and `=` because the .NET methods are case sensitive and
take no wildcards, and `IndexOf` is wrapped in `IFERROR` because Excel errors where .NET returns -1.

Only members Excel can express exactly are translated; a member whose Excel counterpart would compute
something else is left unsupported rather than translated to a near equivalent:

- `Math.Round` rounds halves to even, Excel's `ROUND` rounds them away from zero
- `string.Trim()` strips only the ends, Excel's `TRIM` also collapses runs of spaces inside the text
- `string.IsNullOrWhiteSpace` counts tabs and the Unicode spaces that Excel's `TRIM` does not
- `IndexOf(text, startIndex)` and any overload taking a `StringComparison`, a `CultureInfo` or a
  `MidpointRounding` would have to drop an argument Excel has nowhere to put

Call `ExcelFunctions.Math.Round(...)` or `ExcelFunctions.Text.Trim(...)` when the Excel function is what
you want - then the formula says so explicitly.

Numbers are always written with an invariant decimal point, and text literals are quoted with embedded
quotes doubled, so a formula built under any culture parses in Excel.

## Function reference

¹ Marks a function Excel gained after the 2007 file format was fixed. Those are *stored* under an
`_xlfn.` prefix (`_xlfn.IFS`, `_xlfn.STDEV.P`, and `_xlfn._xlws.FILTER` / `_xlfn._xlws.SORT` for the two
worksheet-only ones) - Excel shows them without it. This library adds the prefix for you; a workbook that
stores the bare name shows `#NAME?` instead of a result. Your Excel version still has to have the function.

### Math and Trigonometry Functions (73)

|Function|C#|
|:---|:---|
|ABS|`ExcelFunctions.Math.Abs(...)`|
|ACOS|`ExcelFunctions.Math.Acos(...)`|
|ACOSH|`ExcelFunctions.Math.Acosh(...)`|
|ACOT ¹|`ExcelFunctions.Math.Acot(...)`|
|ACOTH ¹|`ExcelFunctions.Math.Acoth(...)`|
|AGGREGATE ¹|`ExcelFunctions.Math.Aggregate(...)`|
|ARABIC ¹|`ExcelFunctions.Math.Arabic(...)`|
|ASIN|`ExcelFunctions.Math.Asin(...)`|
|ASINH|`ExcelFunctions.Math.Asinh(...)`|
|ATAN|`ExcelFunctions.Math.Atan(...)`|
|ATAN2|`ExcelFunctions.Math.Atan2(...)`|
|ATANH|`ExcelFunctions.Math.Atanh(...)`|
|BASE ¹|`ExcelFunctions.Math.Base(...)`|
|CEILING|`ExcelFunctions.Math.Ceiling(...)`|
|CEILING.MATH ¹|`ExcelFunctions.Math.CeilingMath(...)`|
|COMBIN|`ExcelFunctions.Math.Combin(...)`|
|COMBINA ¹|`ExcelFunctions.Math.CombinA(...)`|
|COS|`ExcelFunctions.Math.Cos(...)`|
|COSH|`ExcelFunctions.Math.Cosh(...)`|
|COT ¹|`ExcelFunctions.Math.Cot(...)`|
|COTH ¹|`ExcelFunctions.Math.Coth(...)`|
|CSC ¹|`ExcelFunctions.Math.Csc(...)`|
|CSCH ¹|`ExcelFunctions.Math.Csch(...)`|
|DECIMAL ¹|`ExcelFunctions.Math.Decimal(...)`|
|DEGREES|`ExcelFunctions.Math.Degrees(...)`|
|EVEN|`ExcelFunctions.Math.Even(...)`|
|EXP|`ExcelFunctions.Math.Exp(...)`|
|FACT|`ExcelFunctions.Math.Fact(...)`|
|FACTDOUBLE|`ExcelFunctions.Math.FactDouble(...)`|
|FLOOR|`ExcelFunctions.Math.Floor(...)`|
|FLOOR.MATH ¹|`ExcelFunctions.Math.FloorMath(...)`|
|GCD|`ExcelFunctions.Math.Gcd(...)`|
|INT|`ExcelFunctions.Math.Int(...)`|
|LCM|`ExcelFunctions.Math.Lcm(...)`|
|LN|`ExcelFunctions.Math.Ln(...)`|
|LOG|`ExcelFunctions.Math.Log(...)`|
|LOG10|`ExcelFunctions.Math.Log10(...)`|
|MDETERM|`ExcelFunctions.Math.MDeterm(...)`|
|MINVERSE|`ExcelFunctions.Math.MInverse(...)`|
|MMULT|`ExcelFunctions.Math.MMult(...)`|
|MOD|`ExcelFunctions.Math.Mod(...)`|
|MROUND|`ExcelFunctions.Math.MRound(...)`|
|MULTINOMIAL|`ExcelFunctions.Math.MultiNomial(...)`|
|MUNIT ¹|`ExcelFunctions.Math.MUnit(...)`|
|ODD|`ExcelFunctions.Math.Odd(...)`|
|PI|`ExcelFunctions.Math.PI(...)`|
|POWER|`ExcelFunctions.Math.Power(...)`|
|PRODUCT|`ExcelFunctions.Math.Product(...)`|
|QUOTIENT|`ExcelFunctions.Math.Quotient(...)`|
|RADIANS|`ExcelFunctions.Math.Radians(...)`|
|RAND|`ExcelFunctions.Math.Rand(...)`|
|RANDARRAY ¹|`ExcelFunctions.Math.RandArray(...)`|
|RANDBETWEEN|`ExcelFunctions.Math.RandBetween(...)`|
|ROMAN|`ExcelFunctions.Math.Roman(...)`|
|ROUND|`ExcelFunctions.Math.Round(...)`|
|ROUNDDOWN|`ExcelFunctions.Math.RoundDown(...)`|
|ROUNDUP|`ExcelFunctions.Math.RoundUp(...)`|
|SEC ¹|`ExcelFunctions.Math.Sec(...)`|
|SECH ¹|`ExcelFunctions.Math.Sech(...)`|
|SERIESSUM|`ExcelFunctions.Math.SeriesSum(...)`|
|SIGN|`ExcelFunctions.Math.Sign(...)`|
|SIN|`ExcelFunctions.Math.Sin(...)`|
|SINH|`ExcelFunctions.Math.Sinh(...)`|
|SQRT|`ExcelFunctions.Math.Sqrt(...)`|
|SQRTPI|`ExcelFunctions.Math.SqrtPi(...)`|
|SUBTOTAL|`ExcelFunctions.Math.Subtotal(...)`|
|SUMIF|`ExcelFunctions.Math.SumIf(...)`|
|SUMIFS|`ExcelFunctions.Math.SumIfs(...)`|
|SUMPRODUCT|`ExcelFunctions.Math.SumProduct(...)`|
|SUMSQ|`ExcelFunctions.Math.SumSq(...)`|
|TAN|`ExcelFunctions.Math.Tan(...)`|
|TANH|`ExcelFunctions.Math.Tanh(...)`|
|TRUNC|`ExcelFunctions.Math.Trunc(...)`|

### Statistical Functions (80)

|Function|C#|
|:---|:---|
|AVEDEV|`ExcelFunctions.Statistics.AveDev(...)`|
|AVERAGE|`ExcelFunctions.Statistics.Average(...)`|
|AVERAGEA|`ExcelFunctions.Statistics.AverageA(...)`|
|AVERAGEIF|`ExcelFunctions.Statistics.AverageIf(...)`|
|AVERAGEIFS|`ExcelFunctions.Statistics.AverageIfs(...)`|
|BINOM.DIST ¹|`ExcelFunctions.Statistics.BinomDist(...)`|
|CHISQ.TEST ¹|`ExcelFunctions.Statistics.ChiSqTest(...)`|
|CONFIDENCE.NORM ¹|`ExcelFunctions.Statistics.ConfidenceNorm(...)`|
|CONFIDENCE.T ¹|`ExcelFunctions.Statistics.ConfidenceT(...)`|
|CORREL|`ExcelFunctions.Statistics.Correl(...)`|
|COUNT|`ExcelFunctions.Statistics.Count(...)`|
|COUNTA|`ExcelFunctions.Statistics.CountA(...)`|
|COUNTBLANK|`ExcelFunctions.Statistics.CountBlank(...)`|
|COUNTIF|`ExcelFunctions.Statistics.CountIf(...)`|
|COUNTIFS|`ExcelFunctions.Statistics.CountIfs(...)`|
|COVARIANCE.P ¹|`ExcelFunctions.Statistics.CovarianceP(...)`|
|COVARIANCE.S ¹|`ExcelFunctions.Statistics.CovarianceS(...)`|
|DEVSQ|`ExcelFunctions.Statistics.DevSq(...)`|
|EXPON.DIST ¹|`ExcelFunctions.Statistics.ExponDist(...)`|
|FORECAST|`ExcelFunctions.Statistics.Forecast(...)`|
|FORECAST.LINEAR ¹|`ExcelFunctions.Statistics.ForecastLinear(...)`|
|FREQUENCY|`ExcelFunctions.Statistics.Frequency(...)`|
|GEOMEAN|`ExcelFunctions.Statistics.GeoMean(...)`|
|GROWTH|`ExcelFunctions.Statistics.Growth(...)`|
|HARMEAN|`ExcelFunctions.Statistics.HarMean(...)`|
|INTERCEPT|`ExcelFunctions.Statistics.Intercept(...)`|
|KURT|`ExcelFunctions.Statistics.Kurt(...)`|
|LARGE|`ExcelFunctions.Statistics.Large(...)`|
|LINEST|`ExcelFunctions.Statistics.LinEst(...)`|
|LOGEST|`ExcelFunctions.Statistics.LogEst(...)`|
|MAX|`ExcelFunctions.Statistics.Max(...)`|
|MAXA|`ExcelFunctions.Statistics.MaxA(...)`|
|MAXIFS ¹|`ExcelFunctions.Statistics.MaxIfs(...)`|
|MEDIAN|`ExcelFunctions.Statistics.Median(...)`|
|MIN|`ExcelFunctions.Statistics.Min(...)`|
|MINA|`ExcelFunctions.Statistics.MinA(...)`|
|MINIFS ¹|`ExcelFunctions.Statistics.MinIfs(...)`|
|MODE|`ExcelFunctions.Statistics.Mode(...)`|
|MODE.MULT ¹|`ExcelFunctions.Statistics.ModeMult(...)`|
|MODE.SNGL ¹|`ExcelFunctions.Statistics.ModeSngl(...)`|
|NORM.DIST ¹|`ExcelFunctions.Statistics.NormDist(...)`|
|NORM.INV ¹|`ExcelFunctions.Statistics.NormInv(...)`|
|NORM.S.DIST ¹|`ExcelFunctions.Statistics.NormSDist(...)`|
|NORM.S.INV ¹|`ExcelFunctions.Statistics.NormSInv(...)`|
|PERCENTILE|`ExcelFunctions.Statistics.Percentile(...)`|
|PERCENTILE.EXC ¹|`ExcelFunctions.Statistics.PercentileExc(...)`|
|PERCENTILE.INC ¹|`ExcelFunctions.Statistics.PercentileInc(...)`|
|PERCENTRANK|`ExcelFunctions.Statistics.PercentRank(...)`|
|PERMUT|`ExcelFunctions.Statistics.Permut(...)`|
|PERMUTATIONA ¹|`ExcelFunctions.Statistics.PermutationA(...)`|
|POISSON.DIST ¹|`ExcelFunctions.Statistics.PoissonDist(...)`|
|PROB|`ExcelFunctions.Statistics.Prob(...)`|
|QUARTILE|`ExcelFunctions.Statistics.Quartile(...)`|
|QUARTILE.EXC ¹|`ExcelFunctions.Statistics.QuartileExc(...)`|
|QUARTILE.INC ¹|`ExcelFunctions.Statistics.QuartileInc(...)`|
|RANK|`ExcelFunctions.Statistics.Rank(...)`|
|RANK.AVG ¹|`ExcelFunctions.Statistics.RankAvg(...)`|
|RANK.EQ ¹|`ExcelFunctions.Statistics.RankEq(...)`|
|RSQ|`ExcelFunctions.Statistics.Rsq(...)`|
|SKEW|`ExcelFunctions.Statistics.Skew(...)`|
|SLOPE|`ExcelFunctions.Statistics.Slope(...)`|
|SMALL|`ExcelFunctions.Statistics.Small(...)`|
|STANDARDIZE|`ExcelFunctions.Statistics.Standardize(...)`|
|STDEV|`ExcelFunctions.Statistics.StDev(...)`|
|STDEV.P ¹|`ExcelFunctions.Statistics.StDevP(...)`|
|STDEV.S ¹|`ExcelFunctions.Statistics.StDevS(...)`|
|STDEVA|`ExcelFunctions.Statistics.StDevA(...)`|
|STDEVPA|`ExcelFunctions.Statistics.StDevPA(...)`|
|STEYX|`ExcelFunctions.Statistics.SteyX(...)`|
|SUM|`ExcelFunctions.Statistics.Sum(...)`|
|T.DIST ¹|`ExcelFunctions.Statistics.TDist(...)`|
|T.INV ¹|`ExcelFunctions.Statistics.TInv(...)`|
|T.TEST ¹|`ExcelFunctions.Statistics.TTest(...)`|
|TREND|`ExcelFunctions.Statistics.Trend(...)`|
|TRIMMEAN|`ExcelFunctions.Statistics.TrimMean(...)`|
|VAR|`ExcelFunctions.Statistics.Var(...)`|
|VAR.P ¹|`ExcelFunctions.Statistics.VarP(...)`|
|VAR.S ¹|`ExcelFunctions.Statistics.VarS(...)`|
|VARA|`ExcelFunctions.Statistics.VarA(...)`|
|VARPA|`ExcelFunctions.Statistics.VarPA(...)`|

### Logical Functions (11)

|Function|C#|
|:---|:---|
|AND|`ExcelFunctions.Condition.And(...)`|
|FALSE|`ExcelFunctions.Condition.False(...)`|
|IF|`ExcelFunctions.Condition.If(...)`|
|IFERROR|`ExcelFunctions.Condition.IfError(...)`|
|IFNA ¹|`ExcelFunctions.Condition.IfNa(...)`|
|IFS ¹|`ExcelFunctions.Condition.Ifs(...)`|
|NOT|`ExcelFunctions.Condition.Not(...)`|
|OR|`ExcelFunctions.Condition.Or(...)`|
|SWITCH ¹|`ExcelFunctions.Condition.Switch(...)`|
|TRUE|`ExcelFunctions.Condition.True(...)`|
|XOR ¹|`ExcelFunctions.Condition.Xor(...)`|

### Lookup and Reference Functions (33)

|Function|C#|
|:---|:---|
|ADDRESS|`ExcelFunctions.Reference.Address(...)`|
|AREAS|`ExcelFunctions.Reference.Areas(...)`|
|CHOOSE|`ExcelFunctions.Reference.Choose(...)`|
|CHOOSECOLS ¹|`ExcelFunctions.Reference.ChooseCols(...)`|
|CHOOSEROWS ¹|`ExcelFunctions.Reference.ChooseRows(...)`|
|COLUMN|`ExcelFunctions.Reference.Column(...)`|
|COLUMNS|`ExcelFunctions.Reference.Columns(...)`|
|DROP ¹|`ExcelFunctions.Reference.Drop(...)`|
|EXPAND ¹|`ExcelFunctions.Reference.Expand(...)`|
|FILTER ¹|`ExcelFunctions.Reference.Filter(...)`|
|FORMULATEXT ¹|`ExcelFunctions.Reference.FormulaText(...)`|
|HLOOKUP|`ExcelFunctions.Reference.HLookup(...)`|
|HSTACK ¹|`ExcelFunctions.Reference.HStack(...)`|
|HYPERLINK|`ExcelFunctions.Reference.Hyperlink(...)`|
|INDEX|`ExcelFunctions.Reference.Index(...)`|
|INDIRECT|`ExcelFunctions.Reference.Indirect(...)`|
|LOOKUP|`ExcelFunctions.Reference.Lookup(...)`|
|MATCH|`ExcelFunctions.Reference.Match(...)`|
|OFFSET|`ExcelFunctions.Reference.Offset(...)`|
|ROW|`ExcelFunctions.Reference.Row(...)`|
|ROWS|`ExcelFunctions.Reference.Rows(...)`|
|SEQUENCE ¹|`ExcelFunctions.Reference.Sequence(...)`|
|SORT ¹|`ExcelFunctions.Reference.Sort(...)`|
|SORTBY ¹|`ExcelFunctions.Reference.SortBy(...)`|
|TAKE ¹|`ExcelFunctions.Reference.Take(...)`|
|TOCOL ¹|`ExcelFunctions.Reference.ToCol(...)`|
|TOROW ¹|`ExcelFunctions.Reference.ToRow(...)`|
|TRANSPOSE|`ExcelFunctions.Reference.Transpose(...)`|
|UNIQUE ¹|`ExcelFunctions.Reference.Unique(...)`|
|VLOOKUP|`ExcelFunctions.Reference.VLookup(...)`|
|VSTACK ¹|`ExcelFunctions.Reference.VStack(...)`|
|XLOOKUP ¹|`ExcelFunctions.Reference.XLookup(...)`|
|XMATCH ¹|`ExcelFunctions.Reference.XMatch(...)`|

### Date and Time Functions (25)

|Function|C#|
|:---|:---|
|DATE|`ExcelFunctions.DateAndTime.Date(...)`|
|DATEDIF|`ExcelFunctions.DateAndTime.DateDif(...)`|
|DATEVALUE|`ExcelFunctions.DateAndTime.DateValue(...)`|
|DAY|`ExcelFunctions.DateAndTime.Day(...)`|
|DAYS ¹|`ExcelFunctions.DateAndTime.Days(...)`|
|DAYS360|`ExcelFunctions.DateAndTime.Days360(...)`|
|EDATE|`ExcelFunctions.DateAndTime.EDate(...)`|
|EOMONTH|`ExcelFunctions.DateAndTime.EoMonth(...)`|
|HOUR|`ExcelFunctions.DateAndTime.Hour(...)`|
|ISOWEEKNUM ¹|`ExcelFunctions.DateAndTime.IsoWeekNum(...)`|
|MINUTE|`ExcelFunctions.DateAndTime.Minute(...)`|
|MONTH|`ExcelFunctions.DateAndTime.Month(...)`|
|NETWORKDAYS|`ExcelFunctions.DateAndTime.NetworkDays(...)`|
|NETWORKDAYS.INTL ¹|`ExcelFunctions.DateAndTime.NetworkDaysIntl(...)`|
|NOW|`ExcelFunctions.DateAndTime.Now(...)`|
|SECOND|`ExcelFunctions.DateAndTime.Second(...)`|
|TIME|`ExcelFunctions.DateAndTime.Time(...)`|
|TIMEVALUE|`ExcelFunctions.DateAndTime.TimeValue(...)`|
|TODAY|`ExcelFunctions.DateAndTime.Today(...)`|
|WEEKDAY|`ExcelFunctions.DateAndTime.Weekday(...)`|
|WEEKNUM|`ExcelFunctions.DateAndTime.WeekNum(...)`|
|WORKDAY|`ExcelFunctions.DateAndTime.WorkDay(...)`|
|WORKDAY.INTL ¹|`ExcelFunctions.DateAndTime.WorkDayIntl(...)`|
|YEAR|`ExcelFunctions.DateAndTime.Year(...)`|
|YEARFRAC|`ExcelFunctions.DateAndTime.YearFrac(...)`|

### Text Functions (39)

|Function|C#|
|:---|:---|
|ASC|`ExcelFunctions.Text.Asc(...)`|
|CHAR|`ExcelFunctions.Text.Char(...)`|
|CLEAN|`ExcelFunctions.Text.Clean(...)`|
|CODE|`ExcelFunctions.Text.Code(...)`|
|CONCAT ¹|`ExcelFunctions.Text.Concat(...)`|
|CONCATENATE|`ExcelFunctions.Text.Concatenate(...)`|
|DOLLAR|`ExcelFunctions.Text.Dollar(...)`|
|EXACT|`ExcelFunctions.Text.Exact(...)`|
|FIND|`ExcelFunctions.Text.Find(...)`|
|FINDB|`ExcelFunctions.Text.FindB(...)`|
|FIXED|`ExcelFunctions.Text.Fixed(...)`|
|LEFT|`ExcelFunctions.Text.Left(...)`|
|LEFTB|`ExcelFunctions.Text.LeftB(...)`|
|LEN|`ExcelFunctions.Text.Len(...)`|
|LENB|`ExcelFunctions.Text.LenB(...)`|
|LOWER|`ExcelFunctions.Text.Lower(...)`|
|MID|`ExcelFunctions.Text.Mid(...)`|
|MIDB|`ExcelFunctions.Text.MidB(...)`|
|NUMBERVALUE ¹|`ExcelFunctions.Text.NumberValue(...)`|
|PROPER|`ExcelFunctions.Text.Proper(...)`|
|REPLACE|`ExcelFunctions.Text.Replace(...)`|
|REPT|`ExcelFunctions.Text.Rept(...)`|
|RIGHT|`ExcelFunctions.Text.Right(...)`|
|RIGHTB|`ExcelFunctions.Text.RightB(...)`|
|SEARCH|`ExcelFunctions.Text.Search(...)`|
|SEARCHB|`ExcelFunctions.Text.SearchB(...)`|
|SUBSTITUTE|`ExcelFunctions.Text.Substitute(...)`|
|T|`ExcelFunctions.Text.T(...)`|
|TEXT|`ExcelFunctions.Text.Text(...)`|
|TEXTAFTER ¹|`ExcelFunctions.Text.TextAfter(...)`|
|TEXTBEFORE ¹|`ExcelFunctions.Text.TextBefore(...)`|
|TEXTJOIN ¹|`ExcelFunctions.Text.TextJoin(...)`|
|TEXTSPLIT ¹|`ExcelFunctions.Text.TextSplit(...)`|
|TRIM|`ExcelFunctions.Text.Trim(...)`|
|UNICHAR ¹|`ExcelFunctions.Text.UniChar(...)`|
|UNICODE ¹|`ExcelFunctions.Text.Unicode(...)`|
|UPPER|`ExcelFunctions.Text.Upper(...)`|
|VALUE|`ExcelFunctions.Text.Value(...)`|
|WIDECHAR|`ExcelFunctions.Text.WideChar(...)`|

### Information Functions (20)

|Function|C#|
|:---|:---|
|CELL|`ExcelFunctions.Information.Cell(...)`|
|ERROR.TYPE|`ExcelFunctions.Information.ErrorType(...)`|
|INFO|`ExcelFunctions.Information.Info(...)`|
|ISBLANK|`ExcelFunctions.Information.IsBlank(...)`|
|ISERR|`ExcelFunctions.Information.IsErr(...)`|
|ISERROR|`ExcelFunctions.Information.IsError(...)`|
|ISEVEN|`ExcelFunctions.Information.IsEven(...)`|
|ISFORMULA ¹|`ExcelFunctions.Information.IsFormula(...)`|
|ISLOGICAL|`ExcelFunctions.Information.IsLogical(...)`|
|ISNA|`ExcelFunctions.Information.IsNa(...)`|
|ISNONTEXT|`ExcelFunctions.Information.IsNonText(...)`|
|ISNUMBER|`ExcelFunctions.Information.IsNumber(...)`|
|ISODD|`ExcelFunctions.Information.IsOdd(...)`|
|ISREF|`ExcelFunctions.Information.IsRef(...)`|
|ISTEXT|`ExcelFunctions.Information.IsText(...)`|
|N|`ExcelFunctions.Information.N(...)`|
|NA|`ExcelFunctions.Information.Na(...)`|
|SHEET ¹|`ExcelFunctions.Information.Sheet(...)`|
|SHEETS ¹|`ExcelFunctions.Information.Sheets(...)`|
|TYPE|`ExcelFunctions.Information.Type(...)`|

### Financial Functions (21)

|Function|C#|
|:---|:---|
|CUMIPMT|`ExcelFunctions.Financial.CumIPmt(...)`|
|CUMPRINC|`ExcelFunctions.Financial.CumPrinc(...)`|
|DB|`ExcelFunctions.Financial.Db(...)`|
|DDB|`ExcelFunctions.Financial.Ddb(...)`|
|EFFECT|`ExcelFunctions.Financial.Effect(...)`|
|FV|`ExcelFunctions.Financial.Fv(...)`|
|IPMT|`ExcelFunctions.Financial.IPmt(...)`|
|IRR|`ExcelFunctions.Financial.Irr(...)`|
|MIRR|`ExcelFunctions.Financial.MIrr(...)`|
|NOMINAL|`ExcelFunctions.Financial.Nominal(...)`|
|NPER|`ExcelFunctions.Financial.NPer(...)`|
|NPV|`ExcelFunctions.Financial.Npv(...)`|
|PMT|`ExcelFunctions.Financial.Pmt(...)`|
|PPMT|`ExcelFunctions.Financial.PPmt(...)`|
|PV|`ExcelFunctions.Financial.Pv(...)`|
|RATE|`ExcelFunctions.Financial.Rate(...)`|
|SLN|`ExcelFunctions.Financial.Sln(...)`|
|SYD|`ExcelFunctions.Financial.Syd(...)`|
|VDB|`ExcelFunctions.Financial.Vdb(...)`|
|XIRR|`ExcelFunctions.Financial.XIrr(...)`|
|XNPV|`ExcelFunctions.Financial.XNpv(...)`|

### Engineering Functions (20)

|Function|C#|
|:---|:---|
|BIN2DEC|`ExcelFunctions.Engineering.Bin2Dec(...)`|
|BIN2HEX|`ExcelFunctions.Engineering.Bin2Hex(...)`|
|BIN2OCT|`ExcelFunctions.Engineering.Bin2Oct(...)`|
|BITAND ¹|`ExcelFunctions.Engineering.BitAnd(...)`|
|BITLSHIFT ¹|`ExcelFunctions.Engineering.BitLShift(...)`|
|BITOR ¹|`ExcelFunctions.Engineering.BitOr(...)`|
|BITRSHIFT ¹|`ExcelFunctions.Engineering.BitRShift(...)`|
|BITXOR ¹|`ExcelFunctions.Engineering.BitXor(...)`|
|CONVERT|`ExcelFunctions.Engineering.Convert(...)`|
|DEC2BIN|`ExcelFunctions.Engineering.Dec2Bin(...)`|
|DEC2HEX|`ExcelFunctions.Engineering.Dec2Hex(...)`|
|DEC2OCT|`ExcelFunctions.Engineering.Dec2Oct(...)`|
|DELTA|`ExcelFunctions.Engineering.Delta(...)`|
|GESTEP|`ExcelFunctions.Engineering.GeStep(...)`|
|HEX2BIN|`ExcelFunctions.Engineering.Hex2Bin(...)`|
|HEX2DEC|`ExcelFunctions.Engineering.Hex2Dec(...)`|
|HEX2OCT|`ExcelFunctions.Engineering.Hex2Oct(...)`|
|OCT2BIN|`ExcelFunctions.Engineering.Oct2Bin(...)`|
|OCT2DEC|`ExcelFunctions.Engineering.Oct2Dec(...)`|
|OCT2HEX|`ExcelFunctions.Engineering.Oct2Hex(...)`|

### Database Functions (12)

|Function|C#|
|:---|:---|
|DAVERAGE|`ExcelFunctions.Database.DAverage(...)`|
|DCOUNT|`ExcelFunctions.Database.DCount(...)`|
|DCOUNTA|`ExcelFunctions.Database.DCountA(...)`|
|DGET|`ExcelFunctions.Database.DGet(...)`|
|DMAX|`ExcelFunctions.Database.DMax(...)`|
|DMIN|`ExcelFunctions.Database.DMin(...)`|
|DPRODUCT|`ExcelFunctions.Database.DProduct(...)`|
|DSTDEV|`ExcelFunctions.Database.DStDev(...)`|
|DSTDEVP|`ExcelFunctions.Database.DStDevP(...)`|
|DSUM|`ExcelFunctions.Database.DSum(...)`|
|DVAR|`ExcelFunctions.Database.DVar(...)`|
|DVARP|`ExcelFunctions.Database.DVarP(...)`|

### Symbol

|Symbol|Mode|Syntax|Supported|Description|
|---|---|---|:----|:----|
|-|Unary|```c=>-c["One"]```|Yes||
|+|Binary|```c=>c["One"]+c["Two"]```|Yes||
|-|Binary||Yes||
|\*|Binary||Yes||
|/|Binary||Yes||
|%|Binary|```c=>c["One"]%2```|Yes|Written as `MOD(A4,2)`|
|\^|Binary|```c=>c["One"]^2```|Yes|Power|
|=|Binary|```c=>c["One"]==c["Two"]```|Yes||
|\<\> | Binary|```c=>c["One"]!=c["Two"]```|Yes||
|\>|Binary||Yes||
|\<|Binary||Yes||
|\>=|Binary||Yes||
|\<=|Binary||Yes||
|&|Binary|```c=>c["One"]&c["Two"]```|Yes|Join the string|
|:|Binary|```c => c.Matrix("One", 1, "Two", 2)```|Yes|A1:B2|
|,|Binary||Yes|use params Array|
|Space|Binary||No|Range intersection|

#### Reference

- All Functions | https://support.office.com/en-us/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=en-US&rs=en-US&ad=US
