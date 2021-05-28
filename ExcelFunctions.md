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

Function|Syntax |Description
---|:---|
A4 | ```  c => c["One"] ```| Cell in current row
A2 | ```  c => c["One",2] ```| Specify any Cell
A2:B4 | ```  c => c.Matrix("One",2,"Two",4) ```|


### Math Functions

Function|Syntax |Description
---|:---|
ABS | ```  c => ExcelFunctions.Math.Abs(c["One"]) ```|
PI|```  c => ExcelFunctions.Math.PI() ```

### Statistics Functions

Function|Syntax |Description
---|:---|
SUM | ``` c => ExcelFunctions.Statistics.Sum(c.Matix("One", 1, "Two", 2), c.Matrix("Six", 11, "Five", 2)) ```|SUM(A1:B2,F11:E2)

### Condition Functions
Function|Syntax |Description
---|:---|
IF | ``` c => ExcelFunctions.Statistics.Sum(c.Matix("One", 1, "Two", 2), c.Matrix("Six", 11, "Five", 2)) ```|SUM(A1:B2,F11:E2)

### Reference Functions
Function|Syntax |Description
---|:---|
LOOKUP | ``` c => ExcelFunctions.Reference.Lookup(4.19,c.Matrix("One", 2, "One", 6),c.Matrix("Two", 2, "Two", 6)) ```|LOOKUP(4.19, A2:A6, B2:B6)
VLOOKUP | ```c => ExcelFunctions.Reference.VLookup(c["One"], c.Matrix("One", 10, "Three", 20), 2, true)```| VLOOKUP(A4,A10:C20,2,TRUE)
MATCH|
CHOOSE|
INDEX|

### Date and Time Functions
Function|Syntax |Description
---|:---|
DATE | 
DATEDIF |
DAYS|

### Text Functions
Function|Syntax |Description
---|:---|
FIND | ```c => ExcelFunctions.Text.Find("M",c["One",2]) ```|FIND("M",A2)
ASC | 

### Symbol

Symbol|Mode|Synatx|Supported|Description
---|---|---
-|Unary|```c=>-c["One"]```|Yes
+|Binary|```c=>c["One"]+c["Two"]```| Yes
-|Binary|| Yes
\*|Binary|| Yes
/|Binary|| Yes
%|Binary
\^|Binary
=|Binary|```c=>c["One"]==c["Two"]```| Yes
\<\> | Binary|```c=>c["One"]!=c["Two"]```| Yes
\>|Binary|| Yes
\<|Binary|| Yes
\>=|Binary|| Yes
\<=|Binary|| Yes
&|Binary|```c=>c["One"]&c["Two"]```|Yes|Join the string
:|Binary|```c => c.Matrix("One", 1, "Two", 2)```|Yes|A1:B2
,|Binary|||use params Array
Space|Binary

#### Reference

- All Functions | https://support.office.com/en-us/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb?ui=en-US&rs=en-US&ad=US