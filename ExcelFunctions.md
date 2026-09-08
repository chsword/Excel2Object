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


### Math Functions

|Function|Syntax |Description|
|:---|:---|:----|
|ABS | ```  c => ExcelFunctions.Math.Abs(c["One"]) ```|
|PI|```  c => ExcelFunctions.Math.PI() ```|

### Statistics Functions

|Function|Syntax |Description|
|:---|:---|:----|
|SUM | ``` c => ExcelFunctions.Statistics.Sum(c.Matix("One", 1, "Two", 2), c.Matrix("Six", 11, "Five", 2)) ```|SUM(A1:B2,F11:E2)|

### Condition Functions
|Function|Syntax |Description|
|:---|:---|:----|
|IF | ``` c => ExcelFunctions.Statistics.Sum(c.Matix("One", 1, "Two", 2), c.Matrix("Six", 11, "Five", 2)) ```|SUM(A1:B2,F11:E2)|

### Reference Functions
|Function|Syntax |Description|
|:---|:---|:----|
|LOOKUP | ``` c => ExcelFunctions.Reference.Lookup(4.19,c.Matrix("One", 2, "One", 6),c.Matrix("Two", 2, "Two", 6)) ```|LOOKUP(4.19, A2:A6, B2:B6)|
|VLOOKUP | ```c => ExcelFunctions.Reference.VLookup(c["One"], c.Matrix("One", 10, "Three", 20), 2, true)```| VLOOKUP(A4,A10:C20,2,TRUE)|
|MATCH|
|CHOOSE|
|INDEX|

### Date and Time Functions
|Function|Syntax |Description  |
|:---|:---|:-----|
|DATE |                        |
|DATEDIF |                     |
|DAYS||

### Text Functions
|Function|Syntax |Description|
|---|:---|:------|
|FIND | ```c => ExcelFunctions.Text.Find("M",c["One",2]) ```|FIND("M",A2)     |
|ASC |                                                                         |

### Symbol

|Symbol|Mode|Synatx|Supported|Description|
|---|---|---|:----|:----|
|-|Unary|```c=>-c["One"]```|Yes                              |
|+|Binary|```c=>c["One"]+c["Two"]```| Yes                   |
|-|Binary|| Yes                                              |
|\*|Binary|| Yes                                             |
|/|Binary|| Yes                                              |
|%|Binary                                                    |
|\^|Binary                                                   |
|=|Binary|```c=>c["One"]==c["Two"]```| Yes                  |
|\<\> | Binary|```c=>c["One"]!=c["Two"]```| Yes             |
|\>|Binary|| Yes                                             |
|\<|Binary|| Yes                                             |
|\>=|Binary|| Yes                                            |
|\<=|Binary|| Yes                                            |
|&|Binary|```c=>c["One"]&c["Two"]```|Yes|Join the string    |
|:|Binary|```c => c.Matrix("One", 1, "Two", 2)```|Yes|A1:B2 |
|,|Binary|||use params Array                                 |
|Space|Binary                                                |

#### Reference

- All Functions | https://support.office.com/en-us/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=en-US&rs=en-US&ad=US
